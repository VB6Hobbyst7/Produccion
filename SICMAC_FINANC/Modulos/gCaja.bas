Attribute VB_Name = "gCaja"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A805B6C0323"
Option Base 0
Option Explicit
'Utilizado para llenar en una cuadricula los billetjes segun la moneda que se
'ingrese
'##ModelId=3A805BA300B6
Public Function CargaBilletajes(psMoneda As String, poFlexGrid As MSHFlexGrid) As Boolean
    On Error GoTo CargaBilletajesErr

    'your code goes here...

    Exit Function
CargaBilletajesErr:
    Call RaiseError(MyUnhandledError, "ClaseCaja:CargaBilletajes Method")
End Function

'Calcula los totales de la(s) columna(s) ingresada(s) en un FlexGrid
'##ModelId=3A805BAB017F
Public Function TotalesGrid(pnColumna As Integer, poFlexGrid As MSHFlexGrid) As Currency
    On Error GoTo TotalesGridErr

    'your code goes here...

    Exit Function
TotalesGridErr:
    Call RaiseError(MyUnhandledError, "ClaseCaja:TotalesGrid Method")
End Function

'Método que me permite levantar cualquier archivo y de cualquier tipo
'##ModelId=3A8D5D2902BB
Public Function CargaArchivo(ByVal psNomArchivo As String, ByVal psRutaArchivo As String) As Boolean
    On Error GoTo CargaArchivoErr

    'your code goes here...

    Exit Function
CargaArchivoErr:
    Call RaiseError(MyUnhandledError, "ClaseCaja:CargaArchivo Method")
End Function
