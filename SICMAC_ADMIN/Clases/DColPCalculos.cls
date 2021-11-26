VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DColPCalculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String

Dim coConex As DConecta

Private Sub Class_Initialize()
    Dim loIni As ClasIni
    
    Set loIni = New ClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing
    
Set coConex = New DConecta
If coConex.AbreConexion(csConexion) = False Then
    Call RaiseError(MyUnhandledError, "DColPCalculos:Initialize. Error en Conexion a Base de datos")
End If

End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

Public Function dObtieneColocParametro(ByVal pnParametro As Integer) As Double
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = "SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & pnParametro

Set lrs = coConex.CargaRecordSet(lsSQL)

If lrs.BOF And lrs.EOF Then
    dObtieneColocParametro = 0
Else
    dObtieneColocParametro = lrs("nParamValor")
End If

Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Parametros Colocaciones <<dObtieneColocParametro>>", Err.Description
    
End Function

Public Function dObtieneListaParametros(ByVal pnProducto As Producto) As Recordset
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = "SELECT nParamVar, cParamDesc, nParamValor, Mod = '' FROM ColocParametro " _
    & " WHERE nParamVar = " & pnProducto

Set lrs = coConex.CargaRecordSet(lsSQL)
    Set dObtieneListaParametros = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Parametros Colocaciones <<dObtieneColocParametro>>", Err.Description
    
End Function
Public Function dObtieneTasaInteres(ByVal psCodLinCred As String, ByVal pnTipoTasa As Integer) As Double
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

coConex.AbreConexion (csConexion)

lsSQL = "SELECT nTasaIni, nTasaFin FROM ColocLineaCreditoTasa " _
      & "WHERE cLineaCred ='" & psCodLinCred & "' AND nColocLinCredTasaTpo =" & pnTipoTasa & " "

Set lrs = coConex.CargaRecordSet(lsSQL)

If lrs.BOF And lrs.EOF Then
    dObtieneTasaInteres = 0
Else
    dObtieneTasaInteres = lrs("nTasaIni")
End If

Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Tasa Interes <<dObtieneColocParametro>>", Err.Description
    
End Function




