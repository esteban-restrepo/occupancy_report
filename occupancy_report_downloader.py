#Importo librerias
from bs4 import BeautifulSoup
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
from rich.console import Console
from rich.style import Style
import io
import pandas as pd
import requests
import time
import yaml
#Arranco rich e imprimo el encabezado
console = Console(log_path=False)
console.print('\nWINKS DOWNLOADER OCCUPANCY REPORT\n', style='blue bold' )
# Leo el archivo de configuración
console.log('Config file reading')
time.sleep(1)
try:
    with open('config.yaml', 'r') as ymlfile:
        cfg = yaml.safe_load(ymlfile)
    console.log('Config file loaded')
except yaml.YAMLError as error:
    console.log('Config file error')
    time.sleep(5)
    quit()
# Leo la pagina del login de winks para saber el token
console.log('Loggin to winks')
time.sleep(1)
s = requests.session()
#abro la pagina de login
req = s.get('https://pms.winks.com.ar/users/sign_in')
#parseo la pagina web del login
html = BeautifulSoup(req.text, 'html.parser')
#busco el token para el loggin
token = html.find('input', {'name': 'authenticity_token'}).attrs['value']
# Genero el requerimiento de logueo
#genero el diccionario con el payload
payload = {
	'utf8': '✓', 
    'authenticity_token': token, 
	'user[login]': cfg['login']['user'], 
	'user[password]': cfg['login']['password'],
    'commit': 'Iniciar Sesión'
}
#intento el login
req = s.post('https://pms.winks.com.ar/users/sign_in', data = payload)
#valido la respuesta que esté bien
if req.status_code == 200:
    #valido que el login fuera correcto
    html = BeautifulSoup(req.text, 'html.parser')
    found_div = html.find('div', {'class': 'error'})
    if found_div != None:
        #como encuentra el div con clase error error se muestra
        found_div = found_div.find('div').contents
        console.log('Winks response (' + found_div[0] + ')')
        time.sleep(5)
        quit()
    else:
        #Ejecuto el login
        console.log('Successfully logged in')
        time.sleep(1)
else:
    console.log('Request failed with status code: ' + req.status_code)
    time.sleep(5)
    quit()
# Defino el payload para la generación del reporte
console.log('Defining filters')
time.sleep(1)
#defino el inicio de las fecha
dte_from = date.today() + relativedelta(months = cfg['data_range']['past_months'], day = 1)
#defno el fin de las fecha
dte_to = date.today() + relativedelta(months = cfg['data_range']['future_months'], day = 31)
#armo el diccionario con el request
payload = {
    'from' : dte_from.strftime('%Y/%m/%d'),
    'to' : dte_to.strftime('%Y/%m/%d'),
    'include_cancelations' : cfg['filters']['include_cancelations'],
    'include_tentatives' : cfg['filters']['include_tentatives'],
    'include_taxes' : cfg['filters']['include_taxes'],
    'group_by' : cfg['filters']['group_by'],
    'revenue_type' : cfg['filters']['revenue_type'],
    'use_original_layout' : cfg['filters']['use_original_layout'],
    'show_revpar_information' : cfg['filters']['show_revpar_information']
}
# Recorro los hoteles y guardo el reporte de excel
#defino el indice del archivo
index = 0
#recorro todo el diccionario para guardar archivos por cada hotel
with console.status("[bold green]Fetching data...") as status:
    for id_hotel in cfg['hotels']:
        console.log(r'Exporting report for \[' + cfg['hotels'][id_hotel]['name'] + ']')
        time.sleep(1)
        #actualizo el hotel del payload
        payload['accommodation_ids[]'] = [id_hotel]
        #actualizo la moneda del payload
        payload['currency_id'] = cfg['hotels'][id_hotel]['currency']
        #genero el request
        req = s.get('https://pms.winks.com.ar/occupancy_reports/summary.csv', params = payload)
        #valido que estuviera correcta la respuesta
        if req.status_code != 200:
            console.log('Request failed with status code: ' + req.status_code)
            time.sleep(5)
            quit()        
        #valido que quiera archivos separados o que esté consolidado
        if cfg['consolidated'] == 'true':
            #como lo quiero consolidado recorro el archivo pero primero lo debo convertir en tipo archivo el streaming
            file = io.BytesIO(req.content)
            #valido que sea el index 1 para crear el dataframe
            if index == 0:
                df = pd.io.excel.read_excel(file, sheet_name=None)
            #como ya no es el primero entonces lo concateno con los anteriores
            else:
                for sheet in df:
                    df[sheet] = pd.concat([df[sheet], pd.io.excel.read_excel(file, sheet_name=sheet)], ignore_index=True)
            time.sleep(1)
        else:
            #abro_archivo
            file = open(cfg['download_path'] + 'report_file_' + cfg['hotels'][id_hotel]['name'] + '_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '.xlsx', mode = 'wb')
            #como no lo quiero consolidado guardo el archivo
            file.write(req.content)
            #cierro el archivo
            file.close()
            #imprimo el log del hotel completado
            time.sleep(1)
        #adelanto el indice
        index += 1
#valido nuevamente que quiera el archivo consolidado para guardarlo
if cfg['consolidated'] == 'true':
    writer = pd.ExcelWriter(cfg['download_path'] + 'report_file_consolidated_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '.xlsx', engine = 'openpyxl')   
    for sheet in df:
        df[sheet].to_excel(writer, sheet_name=sheet, index=False)   
    writer.close()
    console.log('Consolidated file saved')
    time.sleep(1)
#Ipmrimo el log del proceso completado    
console.log('Export completed')
time.sleep(5)