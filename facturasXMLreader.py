from xml.etree import cElementTree as ET
import pandas as pd

path = 'D:/Google Drive/LA BANQUETA/PEDIDOS PRODUCTOS/FACTURAS/'
file = input("Nombre del archivo: ")
archivo = open(path+file, 'r')
factura = archivo.readlines()
archivo.close()

texto = ""
folio = 0
fecha = ''
proveedor = ''
eans = []
codInterno = []
items = []
cantidad = []
precios = []

for i in factura:
    i = i.lstrip()
    if i[0:7] == "<Folio>":
        i = i.replace("<Folio>", "")
        i = i.replace("</Folio>", "")
        i = i.replace("\n", "")
        folio = i
        i = i + '\n'
    if i[0:8] == "<FchRef>":
        i = i.replace("<FchRef>", "")
        i = i.replace("</FchRef>", "")
        i = i.replace("\n", "")
        fecha = i
        i = i + '\n'
    if i[0:8] == "<RznSoc>":
        i = i.replace("<RznSoc>", "")
        i = i.replace("</RznSoc>", "")
        i = i.replace("\n", "")
        proveedor = i
        i = i + '\n'

    if i[0:11] == "<VlrCodigo>":
        i = i.replace("<VlrCodigo>", "")
        i = i.replace("</VlrCodigo>", "")
        i = i.replace("\n", "")
        if len(i) <= 8:
            codInterno.append(i)
        else:
            eans.append(i)
        i = i + '\n'
    if i[0:9] == "<NmbItem>":
        i = i.replace("<NmbItem>", "")
        i = i.replace("</NmbItem>", "")
        i = i.replace("\n", "")
        items.append(i)
        i = i + '\n'
    if i[0:9] == "<QtyItem>":
        i = i.replace("<QtyItem>", "")
        i = i.replace("</QtyItem>", "")
        i = i.replace("\n", "")
        cantidad.append(int(i))
        i = i + '\n'
    if i[0:9] == "<PrcItem>":
        i = i.replace("<PrcItem>", "")
        i = i.replace("</PrcItem>", "")
        i = i.replace("\n", "")
        precios.append(int(i))
        i = i + '\n'

    texto += i

data = pd.DataFrame(
    {
       # "EAN": eans,
        "CodigoInterno": codInterno,
        "Descripción": items,
        "Cantidad": cantidad,
        "Precio": precios,
        "Venta": 0,
        "Fecha": fecha,
        "Factura": folio,
        "Proveedor": proveedor
    }
)
data["Venta"] = data["Precio"]*data["Cantidad"]
data.head(15)


data.to_excel(path+'{} {} - Factura {}.xlsx'.format(fecha,proveedor,folio))
