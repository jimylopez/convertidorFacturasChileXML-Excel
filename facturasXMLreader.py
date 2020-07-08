import pandas as pd
from datetime import date
import win32api

fechaHoy = date.today()

path = 'C:/TU RUTA DE FACTURAS/XML'

file = input("Ingrese nombre del archivo o 'exit' para salir: ")

contador = open(path+file,'r')
lineas = 0
for line in contador:
    lineas += 1
contador.close()


archivo = open(path+file,'r')
info = archivo.readlines()
archivo.close()

# variables
texto = ""
rut = ''
ruts = []
folio =''
folios = []
fecha = ''
fechas = []
proveedor = ''
proveedores = []
ean = ''
eans = []
codInterno = ''
codInternos = []
item = ''
items = []
unidadMedida = ' '
unidadMedidas = []
cantidad = ''
cantidades = []
precio = ''
precios = []
montoItem = ''
montoItems = []
totSinIva = ''
totSinIvas = []
iva = ''
ivas =[]
excento = ''
excentos = []
totConIva = ''
totConIvas = []

for y in range(lineas-1):
    i = info[y].lstrip()
    if i[0:7] == "<Folio>":
        i = i.replace("<Folio>","")
        i = i.replace("</Folio>","")
        i = i.replace("\n","")
        folio = i
    if i[0:9] == "<FchEmis>":
        i = i.replace("<FchEmis>", "")
        i = i.replace("</FchEmis>", "")
        i = i.replace("\n", "")
        fecha = i
    if i[0:11] == "<RUTEmisor>":
        i = i.replace("<RUTEmisor>","")
        i = i.replace("</RUTEmisor>","")
        i = i.replace("\n","")
        rut = i
    if i[0:8] == "<RznSoc>":
        i = i.replace("<RznSoc>","")
        i = i.replace("</RznSoc>","")
        i = i.replace("\n","")
        proveedor = i
    if i[0:9] == "<MntNeto>":
        i = i.replace("<MntNeto>","")
        i = i.replace("</MntNeto>","")
        i = i.replace("\n","")
        totSinIva=i
    if i[0:8] == "<MntExe>":
        i = i.replace("<MntExe>","")
        i = i.replace("</MntExe>","")
        i = i.replace("\n","")
        excento = i
    if i[0:5] == "<IVA>":
        i = i.replace("<IVA>","")
        i = i.replace("</IVA>","")
        i = i.replace("\n","")
        iva = i
    if i[0:10] == "<MntTotal>":
        i = i.replace("<MntTotal>","")
        i = i.replace("</MntTotal>","")
        i = i.replace("\n","")
        totConIva=i
    
    # detalles:
    if i[0:11] == "<VlrCodigo>":
        i = i.replace("<VlrCodigo>","")
        i = i.replace("</VlrCodigo>","")
        i = i.replace("\n","")
        if len(i) <=8:
            codInterno = i
        else: 
            ean = i
    if i[0:9] == "<NmbItem>":
        i = i.replace("<NmbItem>","")
        i = i.replace("</NmbItem>","")
        i = i.replace("\n","")
        item = i
    if i[0:9] == "<QtyItem>":
        i = i.replace("<QtyItem>","")
        i = i.replace("</QtyItem>","")
        i = i.replace("\n","")
        cantidad = i
    if i[0:10] == "<UnmdItem>":
        i = i.replace("<UnmdItem>","")
        i = i.replace("</UnmdItem>","")
        i = i.replace("\n","")
        unidadMedida = i
    if i[0:9] == "<PrcItem>":
        i = i.replace("<PrcItem>","")
        i = i.replace("</PrcItem>","")
        i = i.replace("\n","")
        precio = i
    if i[0:11] == "<MontoItem>":
        i = i.replace("<MontoItem>","")
        i = i.replace("</MontoItem>","")
        i = i.replace("\n","")
        montoItem = i

    #sube factura
    if i[0:10] == "</Detalle>":
        folios.append(folio)
        fechas.append(fecha)
        ruts.append(rut)
        proveedores.append(proveedor)
        totSinIvas.append(totSinIva)
        excentos.append(excento)
        ivas.append(iva)
        totConIvas.append(totConIva)
        codInternos.append(codInterno)
        eans.append(ean)
        items.append(item)
        cantidades.append(cantidad)
        unidadMedidas.append(unidadMedida)
        precios.append(precio)
        montoItems.append(montoItem)
        
        
    #blanquea next factura
    if i[0:12] == '</Documento>':
        fechaFinal = fecha
        proveedorFinal = proveedor
        folioFinal = folio
        rut = ''
        folio =''
        fecha = ''
        proveedor = ''
        codInterno = ''
        ean = ''
        item = ''
        unidadMedida = ' '
        cantidad = ''
        precio = ''
        montoItem = ''
        totSinIva = ''
        iva = ''
        excento = ''
        totConIva = ''
        
        
data = pd.DataFrame(
    {
        "RUT": ruts,
        "Razón Social":proveedores,
        "Fecha" : fechas,
        "Factura": folios,
        "EAN" : eans,
        "CodigoInterno" : codInternos,
        "Descripción" : items,
        "Cantidad" : cantidades,
        "Unidad Medida":unidadMedidas,
        "Precio Item" : precios,
        "Monto Item": montoItems,
        "MontoFactSinIva": totSinIvas,
        "IVA":ivas,
        "Excento":excentos,
        "MontoFactConIva": totConIvas
        
    }
)
if proveedores[0] != proveedores[-1]:
    data.to_excel(path+'{} Resumen Facturas.xlsx'.format(fechaHoy.isoformat()))
    mensajelog = 'Archivo: {} Resumen Facturas.xlsx creado exitosamente'.format(fechaHoy.isoformat())
    win32api.MessageBox(0, mensajelog, 'Mensaje')
else:
    data.to_excel(path+'{} {} - Factura {}.xlsx'.format(fechaFinal,proveedorFinal,folioFinal))
    mensajelog = 'Archivo: {} {} - Factura {}.xlsx creado exitosamente'.format(fechaFinal,proveedorFinal,folioFinal)
    win32api.MessageBox(0, mensajelog, 'Mensaje')