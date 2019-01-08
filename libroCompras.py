import openpyxl

maxRow = 0
iva = ['10.50', '21', '27']
tipoComprobante = {'FCC':{'A':'001', 'B': '006', 'C' : '011'},
                'NCC':{'A':'003', 'B': '008', 'C' : '013'},
                'NDC':{'A':'002', 'B': '007', 'C' : '012'},
                'REC':{'A':'004', 'B': '009', 'C' : '015'}
                }

compra = openpyxl.load_workbook ('Compras.xlsx')
model  = openpyxl.load_workbook ('MODELO DE FORMULARIO.xlsx')

sheetCompra = compra.active
sheetModel = model.get_sheet_by_name('Carga de Datos')

def codigoDeComprobante (tipo, varToTrim):
    if (tipo is 'FFC'):
        if (varToTrim [:1] is 'A'):
            return tipoComprobante ['FCC']['A']
        elif (varToTrim [:1] is 'B'):
            return tipoComprobante ['FCC']['B']
        elif (varToTrim [:1] is 'C'):
            return tipoComprobante ['FCC']['C']

    if (tipo is 'NCC'):
        if (varToTrim [:1] is 'A'):
            return tipoComprobante ['NCC']['A']
        elif (varToTrim [:1] is 'B'):
            return tipoComprobante ['NCC']['B']
        elif (varToTrim [:1] is 'C'):
            return tipoComprobante ['NCC']['C']

    if (tipo is 'NDC'):
        if (varToTrim [:1] is 'A'):
            return tipoComprobante ['NDC']['A']
        elif (varToTrim [:1] is 'B'):
            return tipoComprobante ['NDC']['B']
        elif (varToTrim [:1] is 'C'):
            return tipoComprobante ['NDC']['C']

    if (tipo is 'REC'):
        if (varToTrim [:1] is 'A'):
            return tipoComprobante ['REC']['A']
        elif (varToTrim [:1] is 'B'):
            return tipoComprobante ['REC']['B']
        elif (varToTrim [:1] is 'C'):
            return tipoComprobante ['REC']['C']

    return "Nada"

def ptoDeVenta (varToTrim):
    return varToTrim [1:5].split("-")[0]

def nroComprobante (varToTrim):
    return varToTrim [6:14].split("/")[0]

def cuit (idComprobante):
    auxString = idComprobante [:2] + idComprobante [3:11] + idComprobante [12:13]
    return auxString

#Se resta 7 en el for porque la planilla de Compras tiene 7 filas semi-completas que no se usan.
for i in range (sheetCompra.max_row-7):
    #Copia celda a celda. Las primeras 4

    #Copia fecha
    sheetModel.cell (row = 13+i, column = 1).value = sheetCompra.cell (row=2+i, column = 1).value
    aux = str (sheetCompra.cell (row=2+i, column = 2).value)
    aux2 = str (sheetCompra.cell (row=2+i, column = 3).value)
    aux3 = str (sheetCompra.cell (row=2+i, column = 4).value)

    #Copia comprobante, pto venta, nro comprobante, cuit, etc.
    sheetModel.cell (row = 13+i, column = 2).value = codigoDeComprobante (aux,aux2)
    sheetModel.cell (row = 13+i, column = 3).value = ptoDeVenta (aux2)
    sheetModel.cell (row = 13+i, column = 4).value = nroComprobante (aux2)
    sheetModel.cell (row = 13+i, column = 7).value = cuit (aux3)
    sheetModel.cell (row = 13+i, column = 8).value = sheetCompra.cell (row=1+i, column = 6).value

model.save ('MODELO DE FORMULARIO.xlsx')