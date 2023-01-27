# Etabs-Connection

import comtypes.client 

#%% Conexión con modelo abierto
def open_model_conection():
    # Establecer conexión primera instacia de Etabs abierta
    etabs = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
    
    # Obtener el modelo asociado
    sap_model = etabs.SapModel
    
    return [etabs, sap_model]

#%% Conexión con modelo cerrado
def close_model_conection(model_path):
    # Establecer conexión con una instacia auxiliar de Etabs
    helper = comtypes.client.CreateObject('ETABSv1.Helper')
    helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
    etabs = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
    
    # Abrir Etabs
    etabs.ApplicationStart()
    
    # Obtener el modelo asociado y abrir el correspondiente a la ruta
    sap_model = etabs.SapModel
    sap_model.File.OpenFile(model_path)
    
    return [etabs, sap_model]
