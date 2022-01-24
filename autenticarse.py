import requests

def loguearse():
    #-----------------------------------------------autenticarse------------------------------------------------#

    usuario={'user':'aalarcon@scienza.com.ar','password':'Adrian2020','PYTHONHTTPSVERIFY':'0'}
    # response = requests.request("GET","https://api.nosconecta.com.ar/auth",headers=usuario,verify=False)

    response = requests.get("https://api.nosconecta.com.ar/auth",headers=usuario,verify=False)

    if response.status_code==200:
        payload=response.json()
        token=payload['message']['token']
        token="Bearer " + token

    return token 
    
