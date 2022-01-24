from autenticarse import loguearse
from descarga_remito import remitos
from getpass import getuser
import time
import sys

def proceso(num_hilo,id_cliente,token,user,alcance,facturas,entregas,factura_papel,pedido_scienza,
            pedido_cliente,remito_papel,clasedearmado,requierepresupuesto,requiereoc,fila):
    
    contador=0   
    for elemento in range(0,alcance):
        if facturas[elemento]==None:
                continue
        else:
            if clasedearmado=="osde":
                carpeta=('C:/Users/' + user +  '/Desktop/legajos' + '/' + str(id_cliente) + '/' + 'Remitos')
                modelo=str(carpeta) + '/' + "0" + str(factura_papel[elemento]) + "_" + str(remito_papel[elemento]) + '.pdf'
            else:
                carpeta=('C:/Users/' + user +  '/Desktop/legajos' + '/' + str(id_cliente) + '/' + str(factura_papel[elemento]) + '/' + 'Remitos')
                modelo=str(carpeta) + '/' + str(remito_papel[elemento]) + ' - ' + str(pedido_cliente[elemento]) + '.pdf'
            try: 
                if entregas[elemento]==None:
                    continue
                else:
                    remitos(str(entregas[elemento]),token,factura_papel[elemento],remito_papel[elemento],modelo,pedido_cliente[elemento])    
            except:
                print(str(sys.exc_info()))
                continue
        contador=contador+1
    print("remitos de " + str(id_cliente) + " " + str(contador))
        
        
