#######PARA EL BOT##############################
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import datetime, time
from datetime import date, timedelta
################################################

#########MANEJO DE FECHA Y HORA OS
import time
import os
import shutil
##################################################

#######PARA EL TRATAMIENTO DE LA INFO EXCEL
import pandas as pd
import numpy as np
#################################################

################PARA EL ENVIO DE CORREO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
#################################################
import sys

##PASO1.######EL BOT HACE SU TRABAJO: ABRE SISCOVID FILTRA DIA ANTERIOR DESCARGA ARCHIVO ############
browser = webdriver.Chrome()
browser.get("https://logincovid19.minsa.gob.pe/accounts/login/") 
time.sleep(10)

#Datos que extraemos de la pagina inspeccionando elemento
browser.find_element_by_name("username").send_keys("xxxxxxx")#usuario
browser.find_element_by_name("password").send_keys("*******")#password

browser.find_element_by_id("botonEnviarLogin").click()
#browser.find_element_by_name("nombreBoton").click() si fuera name agregamos .click()
time.sleep(5)
browser.get("https://siscovid.minsa.gob.pe/reporte-prueba-rapida/") 
time.sleep(5)
#login_attempt.submit()

fecha1 = browser.find_element_by_id('id_fecha')
fecha2 = browser.find_element_by_id('id_fecha_fin')

date = datetime.date.today()
td=timedelta(1)#EL NUMERO ES LA DIFERENCIA DE DIAS 
date = datetime.date.today()-td

year = date.strftime("%Y")
month = date.strftime("%m")
day = date.strftime("%d")

fecha_r=day+'/'+month+'/'+year#restada
fecha1.send_keys(fecha_r)
fecha2.send_keys(fecha_r)

time.sleep(5)
browser.find_element_by_id("btnBuscar").click()
time.sleep(15)
browser.find_element_by_id("btnGroupDrop1").click()
time.sleep(1)
browser.find_element_by_id("idExportarExcel").click()
time.sleep(20)

###MUEVE EL ARCHIVO DESCARGADO A LA CARPETA DONDE SE ENCUENTRA EL SCRIPT#######
filepath = "C:/Users/Alex/Downloads"
filename = max([filepath + "/" + f for f in os.listdir(filepath)], key=os.path.getctime)
current_dir = os.path.dirname(os.path.realpath(__file__))

shutil.move(os.path.join(current_dir,filename),"DATASOURCE/DIARIO.xlsx")
#######################FINALIZA PASO1.##############################
browser.close()

###########################TRATAMIENTO DEL EXCEL#############################
current_dir = os.path.dirname(os.path.realpath(__file__)) 
###########################ABRIR ARCHIVOS 2021 PARA UNIR############
diario = os.path.join(current_dir, "DATASOURCE/DIARIO.xlsx")
#LEER ARCHIVOS 2021
file=pd.read_excel(diario)
#CONVERTIR EN DATAFRAME
df=pd.DataFrame(file)

#CAMBIAR EL NOMBRE DE LA COLUMNA(ÍNDICE 79) Resultado A Resultado.1 
df.columns.values[79]="Resultado.1"

df=df.loc[:,['Nro Documento','nombres','Apellido Paterno','Apellido Materno','comun_sexo_paciente','Edad','Provincia','Distrito','Tipo de Prueba','Fecha Ejecucion Prueba','Resultado.1','Usuario','Fecha Inicio Sintomas de la Ficha Paciente','cod_establecimiento_ejecuta','Establecimiento_Ejecuta','Direccion','Etnia']]
df['NombreCompleto']=df['nombres']+" "+df['Apellido Paterno']+" " + df['Apellido Materno']

df['comun_sexo_paciente']=df['comun_sexo_paciente'].replace(['FEMENINO','MASCULINO'],['F','M'])

df.insert(1,"Nro",0,allow_duplicates=True)
df.insert(9,"ResultadoFinal","",allow_duplicates=True)
df.insert(10,"Fuente","SISCOVID",allow_duplicates=True)

df.insert(8,"Semana","",allow_duplicates=True)
df.insert(9,"AÑO","",allow_duplicates=True)
df.insert(10,"MES","",allow_duplicates=True)
df.insert(11,"DIA","",allow_duplicates=True)
df.insert(12,"Ultimos 07 Dias","0",allow_duplicates=True)
df.insert(13,"Dias Transcurridos","",allow_duplicates=True)
df.insert(14,"Estado Actual","",allow_duplicates=True)
df.insert(15,"Fallecido","",allow_duplicates=True)

df.insert(1,"Dni_Final","",allow_duplicates=True)
df.insert(5,"Etapa de Vida","",allow_duplicates=True)

df["ResultadoFinal"]=["IgG" if str(r)=='IgG Reactivo' else r for r in df['Resultado.1']]
df["ResultadoFinal"]=['POSITIVO' if str(r).find('IgM')!= -1 or str(r)=='Reactivo' else r for r in df['ResultadoFinal']]
df["ResultadoFinal"]=[r if str(r)=='POSITIVO' or str(r)=='IgG' else 'NEGATIVO' for r in df['ResultadoFinal']]

df["Dni_Final"]=["00000000"+ str(r) for r in df['Nro Documento']]
df["Dni_Final"]=[str(r)[-8:] for r in df['Dni_Final']]
df["Tipo de Prueba"]=["Px Antígeno" if str(r).find('Antígeno')!= -1 else "Px Rápida" for r in df['Tipo de Prueba']]

df["Dni_Final"]=df['Dni_Final'].str.replace('nan','00000000')

df["Distrito"]=df['Distrito'].str.upper()
df["Provincia"]=df['Provincia'].str.upper()

df["Distrito"]=df['Distrito'].str.replace('Á','A')
df["Distrito"]=df['Distrito'].str.replace('É','E')
df["Distrito"]=df['Distrito'].str.replace('Í','I')
df["Distrito"]=df['Distrito'].str.replace('Ó','O')

df["Provincia"]=df['Provincia'].str.replace('Ó','O')

###############################
#Create a List of our conditions
conditionsp=[
(df['Distrito']=='BALSAPUERTO'),
(df['Distrito']=='JEBEROS'),
(df['Distrito']=='LAGUNAS'),
(df['Distrito']=='SANTA CRUZ'),
(df['Distrito']=='TENIENTE CESAR LOPEZ ROJAS'),
(df['Distrito']=='YURIMAGUAS'),
(df['Distrito']=='ANDOAS'),
(df['Distrito']=='BARRANCA'),
(df['Distrito']=='CAHUAPANAS'),
(df['Distrito']=='MANSERICHE'),
(df['Distrito']=='MORONA'),
(df['Distrito']=='PASTAZA'),
(df['Distrito']=='NAUTA'),
(df['Distrito']=='PARINARI'),
(df['Distrito']=='TIGRE'),
(df['Distrito']=='TROMPETEROS'),
(df['Distrito']=='URARINAS'),
(df['Distrito']=='ALTO NANAY'),
(df['Distrito']=='BELEN'),
(df['Distrito']=='INDIANA'),
(df['Distrito']=='IQUITOS'),
(df['Distrito']=='FERNANDO LORES'),
(df['Distrito']=='LAS AMAZONAS'),
(df['Distrito']=='MAZAN'),
(df['Distrito']=='NAPO'),
(df['Distrito']=='PUNCHANA'),
(df['Distrito']=='TORRES CAUSANA'),
(df['Distrito']=='SAN JUAN BAUTISTA'),
(df['Distrito']=='PEBAS'),
(df['Distrito']=='RAMON CASTILLA'),
(df['Distrito']=='SAN PABLO'),
(df['Distrito']=='YAVARI'),
(df['Distrito']=='ALTO TAPICHE'),
(df['Distrito']=='CAPELO'),
(df['Distrito']=='EMILIO SAN MARTIN'),
(df['Distrito']=='JENARO HERRERA'),
(df['Distrito']=='MAQUIA'),
(df['Distrito']=='PUINAHUA'),
(df['Distrito']=='REQUENA'),
(df['Distrito']=='SAQUENA'),
(df['Distrito']=='SOPLIN'),
(df['Distrito']=='TAPICHE'),
(df['Distrito']=='YAQUERANA'),
(df['Distrito']=='CONTAMANA'),
(df['Distrito']=='INAHUAYA'),
(df['Distrito']=='PADRE MARQUEZ'),
(df['Distrito']=='PAMPA HERMOSA'),
(df['Distrito']=='SARAYACU'),
(df['Distrito']=='VARGAS GUERRA'),
(df['Distrito']=='PUTUMAYO'),
(df['Distrito']=='ROSA PANDURO'),
(df['Distrito']=='TENIENTE MANUEL CLAVERO'),
(df['Distrito']=='YAGUAS'),
]

valuesp=['ALTO AMAZONAS',
'ALTO AMAZONAS',
'ALTO AMAZONAS',
'ALTO AMAZONAS',
'ALTO AMAZONAS',
'ALTO AMAZONAS',
'DATEM DEL MARAÑON',
'DATEM DEL MARAÑON',
'DATEM DEL MARAÑON',
'DATEM DEL MARAÑON',
'DATEM DEL MARAÑON',
'DATEM DEL MARAÑON',
'LORETO',
'LORETO',
'LORETO',
'LORETO',
'LORETO',
'MAYNAS',
'MAYNAS',
'MAYNAS',
'MAYNAS',
'MAYNAS',
'MAYNAS',
'MAYNAS',
'MAYNAS',
'MAYNAS',
'MAYNAS',
'MAYNAS',
'RAMON CASTILLA',
'RAMON CASTILLA',
'RAMON CASTILLA',
'RAMON CASTILLA',
'REQUENA',
'REQUENA',
'REQUENA',
'REQUENA',
'REQUENA',
'REQUENA',
'REQUENA',
'REQUENA',
'REQUENA',
'REQUENA',
'REQUENA',
'UCAYALI',
'UCAYALI',
'UCAYALI',
'UCAYALI',
'UCAYALI',
'UCAYALI',
'PUTUMAYO',
'PUTUMAYO',
'PUTUMAYO',
'PUTUMAYO'
]
df["Provincia"]=np.select(conditionsp,valuesp)
#############
df["Distrito"]=df['Distrito'].str.replace('BALSAPUERTO','160202 - BALSAPUERTO')
df["Distrito"]=df['Distrito'].str.replace('JEBEROS','160205 - JEBEROS')
df["Distrito"]=df['Distrito'].str.replace('LAGUNAS','160206 - LAGUNAS')
df["Distrito"]=df['Distrito'].str.replace('SANTA CRUZ','160210 - SANTA CRUZ')
df["Distrito"]=df['Distrito'].str.replace('TENIENTE CESAR LOPEZ ROJAS','160211 - TENIENTE CESAR LOPEZ ROJAS')
df["Distrito"]=df['Distrito'].str.replace('YURIMAGUAS','160201 - YURIMAGUAS')
df["Distrito"]=df['Distrito'].str.replace('ANDOAS','160706 - ANDOAS')
df["Distrito"]=df['Distrito'].str.replace('BARRANCA','160701 - BARRANCA')
df["Distrito"]=df['Distrito'].str.replace('CAHUAPANAS','160702 - CAHUAPANAS')
df["Distrito"]=df['Distrito'].str.replace('MANSERICHE','160703 - MANSERICHE')
df["Distrito"]=df['Distrito'].str.replace('MORONA','160704 - MORONA')
df["Distrito"]=df['Distrito'].str.replace('PASTAZA','160705 - PASTAZA')
df["Distrito"]=df['Distrito'].str.replace('NAUTA','160301 - NAUTA')
df["Distrito"]=df['Distrito'].str.replace('PARINARI','160302 - PARINARI')
df["Distrito"]=df['Distrito'].str.replace('TIGRE','160303 - TIGRE')
df["Distrito"]=df['Distrito'].str.replace('TROMPETEROS','160304 - TROMPETEROS')
df["Distrito"]=df['Distrito'].str.replace('URARINAS','160305 - URARINAS')
df["Distrito"]=df['Distrito'].str.replace('ALTO NANAY','160102 - ALTO NANAY')
df["Distrito"]=df['Distrito'].str.replace('BELEN','160112 - BELEN')
df["Distrito"]=df['Distrito'].str.replace('INDIANA','160104 - INDIANA')
df["Distrito"]=df['Distrito'].str.replace('IQUITOS','160101 - IQUITOS')
df["Distrito"]=df['Distrito'].str.replace('FERNANDO LORES','160103 - FERNANDO LORES')
df["Distrito"]=df['Distrito'].str.replace('LAS AMAZONAS','160105 - LAS AMAZONAS')
df["Distrito"]=df['Distrito'].str.replace('MAZAN','160106 - MAZAN')
df["Distrito"]=df['Distrito'].str.replace('NAPO','160107 - NAPO')
df["Distrito"]=df['Distrito'].str.replace('PUNCHANA','160108 - PUNCHANA')
df["Distrito"]=df['Distrito'].str.replace('TORRES CAUSANA','160110 - TORRES CAUSANA')
df["Distrito"]=df['Distrito'].str.replace('SAN JUAN BAUTISTA','160113 - SAN JUAN BAUTISTA')
df["Distrito"]=df['Distrito'].str.replace('PEBAS','160402 - PEBAS')
df["Distrito"]=df['Distrito'].str.replace('RAMON CASTILLA','160401 - RAMON CASTILLA')
df["Distrito"]=df['Distrito'].str.replace('SAN PABLO','160404 - SAN PABLO')
df["Distrito"]=df['Distrito'].str.replace('YAVARI','160403 - YAVARI')
df["Distrito"]=df['Distrito'].str.replace('ALTO TAPICHE','160502 - ALTO TAPICHE')
df["Distrito"]=df['Distrito'].str.replace('CAPELO','160503 - CAPELO')
df["Distrito"]=df['Distrito'].str.replace('EMILIO SAN MARTIN','160504 - EMILIO SAN MARTIN')
df["Distrito"]=df['Distrito'].str.replace('JENARO HERRERA','160510 - JENARO HERRERA')
df["Distrito"]=df['Distrito'].str.replace('MAQUIA','160505 - MAQUIA')
df["Distrito"]=df['Distrito'].str.replace('PUINAHUA','160506 - PUINAHUA')
df["Distrito"]=df['Distrito'].str.replace('REQUENA','160501 - REQUENA')
df["Distrito"]=df['Distrito'].str.replace('SAQUENA','160507 - SAQUENA')
df["Distrito"]=df['Distrito'].str.replace('SOPLIN','160508 - SOPLIN')
df["Distrito"]=df['Distrito'].str.replace('TAPICHE','160509 - TAPICHE')
df["Distrito"]=df['Distrito'].str.replace('YAQUERANA','160511 - YAQUERANA')
df["Distrito"]=df['Distrito'].str.replace('CONTAMANA','160601 - CONTAMANA')
df["Distrito"]=df['Distrito'].str.replace('INAHUAYA','160602 - INAHUAYA')
df["Distrito"]=df['Distrito'].str.replace('PADRE MARQUEZ','160603 - PADRE MARQUEZ')
df["Distrito"]=df['Distrito'].str.replace('PAMPA HERMOSA','160604 - PAMPA HERMOSA')
df["Distrito"]=df['Distrito'].str.replace('SARAYACU','160605 - SARAYACU')
df["Distrito"]=df['Distrito'].str.replace('VARGAS GUERRA','160606 - VARGAS GUERRA')
df["Distrito"]=df['Distrito'].str.replace('PUTUMAYO','160801 - PUTUMAYO')
df["Distrito"]=df['Distrito'].str.replace('ROSA PANDURO','160802 - ROSA PANDURO')
df["Distrito"]=df['Distrito'].str.replace('TENIENTE MANUEL CLAVERO','160803 - TENIENTE MANUEL CLAVERO')
df["Distrito"]=df['Distrito'].str.replace('YAGUAS','160804 - YAGUAS')
#Create a List of our conditions
conditions=[
	(df['Edad']<12),
	(df['Edad']<18),
	(df['Edad']<30),
	(df['Edad']<60),
	(df['Edad']<200),
]
#Create a List values of each conditions
values =['Niño',
'Adolescente',
'Joven',
'Adulto',
'Adulto Mayor']
df["Etapa de Vida"]=np.select(conditions,values)

#df["Fecha Ejecucion Prueba"]=[str(r)[:-9] for r in df['Fecha Ejecucion Prueba']]
df["Fecha Ejecucion Prueba"]=pd.to_datetime(df['Fecha Ejecucion Prueba'])
df["Semana"]=df['Fecha Ejecucion Prueba'].dt.isocalendar().week
df["AÑO"]=df['Fecha Ejecucion Prueba'].dt.year
df["MES"]=df['Fecha Ejecucion Prueba'].dt.month
df["DIA"]=df['Fecha Ejecucion Prueba'].dt.day

td=timedelta(0)#EL NUMERO ES LA DIFERENCIA DE DIAS 

today = pd.datetime.now() - td

df["Dias Transcurridos"]=(pd.to_datetime(today) - df["Fecha Ejecucion Prueba"]).dt.days

df= df.loc[:,['Tipo de Prueba','Dni_Final','NombreCompleto','comun_sexo_paciente','Edad','Etapa de Vida','Distrito','Provincia','Fecha Inicio Sintomas de la Ficha Paciente','Fecha Ejecucion Prueba','Semana','AÑO','MES','DIA','Ultimos 07 Dias','Dias Transcurridos','Estado Actual','Fallecido','Resultado.1','ResultadoFinal','Fuente','Usuario','Etnia','cod_establecimiento_ejecuta','Establecimiento_Ejecuta','Direccion']]

df["Tipo de Prueba"]=[r if str(r).find('Antígeno')!= -1 else 'NO_PRUEBA' for r in df['Tipo de Prueba']]
df_mask1=df['Tipo de Prueba']!='NO_PRUEBA'
df=df[df_mask1]

#date = datetime.date.today()
#year = date.strftime("%Y")
#month = date.strftime("%m")
#day = date.strftime("%d")

#fecha_hoy=day+'-'+month+'-'+year

df=df.reset_index(drop=True)
df.index = df.index+1
###############################################
date = datetime.date.today()
td=timedelta(1)#EL NUMERO ES LA DIFERENCIA DE DIAS 
date = datetime.date.today()-td

year = date.strftime("%Y")
month = date.strftime("%m")
day = date.strftime("%d")

fecha_r=day + '-'+ month + '-' + year#restada
#####################################################

nombre_archivo="SISCOVID_"+ fecha_r + ".xlsx"
df.to_excel("RESULT/"+nombre_archivo)
###################################TERMINA TRATAMIENDO DEL EXCEL##################################
#################################################################################################

#######################PASO 3. ENVIA CORREO ELECTRONICO######################################
# Iniciamos los parámetros del script
server = smtplib.SMTP(host='mail.diresaloreto.gob.pe')
password='rancio#763'
remitente = 'noreply@diresaloreto.gob.pe'
destinatarios = ['tucorreo@gmail.com','otrocorreo@outllok.com']#'cirene1_pdv@hotmail.com','Patyqr133@gmail.com','patyqr1@hotmail.com','fertulum@gmail.com',
asunto = 'PruebasAntígeno+/RefMem-038-2021'
cuerpo = 'Hola, envío el reporte diario de Pruebas, Saludos'
ruta_adjunto = ('RESULT/'+nombre_archivo)
nombre_adjunto = nombre_archivo
# Creamos el objeto mensaje
mensaje = MIMEMultipart()
# Establecemos los atributos del mensaje
mensaje['From'] = remitente
mensaje['To'] = ", ".join(destinatarios)
mensaje['Subject'] = asunto
# Agregamos el cuerpo del mensaje como objeto MIME de tipo texto
mensaje.attach(MIMEText(cuerpo, 'plain'))
 # Abrimos el archivo que vamos a adjuntar
archivo_adjunto = open(ruta_adjunto, 'rb')
# Creamos un objeto MIME base
adjunto_MIME = MIMEBase('application', 'octet-stream')
######################################################################
# Y le cargamos el archivo adjunto
adjunto_MIME.set_payload((archivo_adjunto).read())
# Codificamos el objeto en BASE64
encoders.encode_base64(adjunto_MIME)
# Agregamos una cabecera al objeto
adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)
# Y finalmente lo agregamos al mensaje
mensaje.attach(adjunto_MIME)
######################################################################
#ciframos la conexion
server.starttls()
#Iniciamos sesión en l Servidor
server.login(mensaje['From'], password)


# Convertimos el objeto mensaje a texto
texto = mensaje.as_string()

# Enviamos el mensaje
server.sendmail(remitente, destinatarios, texto)

time.sleep(120)

server.quit()

time.sleep(60)
#########terminamos el script de python 
sys.exit()
