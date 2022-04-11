""" 
IDEA: Crear un ejecutable que corra un query, lo cargue en un excel lo envíe por mail a un ID.

STEP 1= Armar la conexión a STFMVS1 y correr el query
STEP 2= Pasar el output del query a un excel y darle formato
STEP 3= Enviar el excel a un ID 
STEP 4=  Copilarlo en un .exe 

Enhacements a futuro:

1- Modificar el campo del password para que sea invisible. Es posible agregarle el ojito como para definir si queremos que sea o no visible? --> LISTO!
2- Linkear una base de datos al script cosa de que los datos ingresados, como usuario, sender y recepient queden cargados y los usuarios
    puedan controlar los datos que desean tener cargados.
3- 

"""

import pandas as pd
import ibm_db
import ibm_db_dbi as dbi
import ftplib as FTP
import os
import logging
from sender.gateway.email_server import EmailServer
import PySimpleGUI as sg

#__________________________________________________________ Variables Globales  __________________________________________________________#

output_file = "output.xlsx"
userid="xxxxxxx" #input("Enter your STFMVS1 ID username: ")
psswd="xxxxxxx" #input("Enter your STFMVS1 ID password: ")#gp.getpass("Enter your STFMVS1 ID password (hidden): ")


query = (f"""SELECT 'LR',CDATE,SUM(AMTLOC),COUNT(*)
FROM MIRPCC.USDETAIL_CM_D
WHERE CTY='897'
AND STREAM_ID = 'US'
AND TOLI='L'
AND AMTLOC <> 0.00
AND MAJ > '599'
AND ITYP <> 'L'
AND DIV NOT IN ('8P','MW','23','C3','8F')
AND DIV||MAJ||MINOR||SMIN  != 'DG92099990000'
GROUP BY CDATE
ORDER BY 2,3,4,1 """)
#__________________________________________________________ GUI para Credenciales  __________________________________________________________#

new_query = ''

layout = [
        [sg.Text("Please enter your ID, password and query")],
        [sg.Text('User ID:', size=(12,1)),sg.InputText()],
        [sg.Text('Password:', size=(12,1)),sg.InputText(password_char='*')], #'',key='Password',password_char='*')],
        [sg.Text('Query 2:',size=(12,1)), sg.Multiline(default_text=query, size=(40, 3))],
        
        [sg.Button("OK"), sg.Cancel()]
        ]

layout2 = [
        [sg.Text("Mail Info")],
        [sg.Text('Sender:', size=(12,1)),sg.InputText()],
        [sg.Text('Recipient:', size=(12,1)),sg.InputText()],
        [sg.Text('Additional Comments:', size=(12,2)), sg.Multiline(default_text="FYI", size=(40, 3))],
        [sg.Button("OK"), sg.Cancel()]
]

window = sg.Window("STFMVS1 Node - Enter data", layout)

event, values = window.read()

#_____________________________ MAIL INFO ______________________________#

#window2 = sg.Window("STFMVS1 Node - Enter data", layout2)

#event, values2 = window2.read()

window.close()

input_userid = values[0]
input_psswd = values[1]

if values[2] != query:
    query = ((str(values[2])))


window2 = sg.Window("STFMVS1 Node - Enter data", layout2)

event, values2 = window2.read()

window.close()

#print("\n",query,"\n","\n",type(query))
#print(input_userid,input_psswd)

#__________________________________________________________ STEP 1 __________________________________________________________#

#Connection to DB

conn_str = (f"DATABASE= USIBMVRDP1H;HOSTNAME=stfmvs1.pok.ibm.com;PORT=5007;PROTOCOL= TCPIP;UID={input_userid};PWD={input_psswd}")
 
conn = ibm_db.connect(conn_str, '', '')

pconn = dbi.Connection(conn)
 
prep = ibm_db.prepare(conn, query)


#__________________________________________________________ STEP 2 __________________________________________________________#

try:                                 #Acá armamos un dataframe con pandas utilizando la información que extrae el query y guardamos el output en un excel
    data =pd.read_sql(query, pconn)        
    data.to_excel(output_file, index=False)
   
except:                              #En caso de que falle por alguna razón, nos arroja una ventana con el error de SQL
    sg.Window(title="Statement Error",layout=[[sg.Text("SQL statement didn't run, error: {}".format(ibm_db.stmt_error))]], margins=(40,25)).read()
    #print("SQL statement didn't run, error: ", ibm_db.stmt_error)

#__________________________________________________________ STEP 3 __________________________________________________________#


def send_mail(message):                        # En este step definimos una función para enviar un mail a través de notes definiendo variables como
    funcion=EmailServer()                      #el sender, recepient y el mensaje, todos estos datos se completan con el prompt del layout 2 y 
                                               # se linkean con los valores en values2. Finalmente se adjunta el archivo y recibimos el output.
    sender= values2[0]
    recipient= values2[1]
    subject= f'{input_userid} - STFMVS1 Query Output'
    
    funcion.send_normal_mail_message_attach(sender,recipient,subject,message, output_file) 
    sg.Window(title="Mail Confirmation",layout=[[sg.Text(f"Email Successfully sent to: {recipient}")]], margins=(40,25)).read()



send_mail(values2[2])
