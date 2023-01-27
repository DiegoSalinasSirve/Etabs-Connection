# Etabs-Connection

#### La conexión con etabs requiere obtener el objeto etabs el cual permite manipular el programa y el objeto sap_model el cual permite manipular el modelo como tal. Esta conexión puede establecerse  con un modelo abierto como con un modelo cerrado.

#### En primer lugar es necesario importar la libreria comtypes.client:

```
import comtypes.client 
```

#### Para la conexión con un modelo abierto se utiliza la siguiente función:
```
def open_model_conection():
    # Establecer conexión primera instacia de Etabs abierta
    etabs = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
    
    # Obtener el modelo asociado
    sap_model = etabs.SapModel
    
    return [etabs, sap_model]
    
```

#### Para la conexión con un modelo cerrado  se utiliza la siguiente función:

```
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
    
```
