import requests
import time
import json
from getpass import getuser
from autenticarse import loguearse

def remitos(entrega,token,factura_papel,remito_papel,modelo,pedido_cliente):
    #---------------------------------------------------Buscar-------------------------------------------------------------#
    entrega=entrega
    #payload=json.dump('filter',{'key':"entrega_id",'value':83691438,'operator':"="})
    headers={'accept': 'application/json','Authorization':token,'PYTHONHTTPSVERIFY':'0'}
    response = requests.get("https://api.nosconecta.com.ar:443/search/353?filter=[{\"key\": \"entrega_id\",     \"value\":" +  str(entrega) + ",     \"operator\": \"=\" }]",headers=headers,stream=True,verify=False)

    id=None
    if response.status_code==200:
        payload=response.json()
        results=payload['message']['results']
        print(results)
        if results!=[]:
            for i in results:
                if i.get('entrega_id')==entrega: #and (i.get('lote')[0]=="A" or i.get('lote')[0]=="R" or i.get('lote')[0]=="F"):
                    id=i.get('id')
                    print(id)
                    break
            
            if id==None:
                for i in results:
                    if i.get('entrega_id')==entrega and (i.get('lote')[0]=="M"):
                        id=i.get('id')
                        print(id)
                        break 
            
    #-------------------------------------------------------Descargar-----------------------------------------------------------#
            headers={'accept': 'application/json','Authorization':token,'PYTHONHTTPSVERIFY':'0'}
            response = requests.get("https://api.nosconecta.com.ar:443/file/353/" + str(id),headers=headers,stream=True,verify=False)
            print(f"Listo para descargar el remito!.")
            with open(modelo,'wb') as file:
                for entrega in response.iter_content():
                    file.write(entrega)


token = loguearse()
user = getuser()
modelo = 'C:/Users/' + user +  '/Desktop/remitos'
entrega = "85419120"
factura_papel = 1
remito_papel = 1
pedido_cliente = 1

remitos(entrega, token, factura_papel, remito_papel, modelo, pedido_cliente)