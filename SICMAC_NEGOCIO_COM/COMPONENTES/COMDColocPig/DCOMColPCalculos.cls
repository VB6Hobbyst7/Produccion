VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMColPCalculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim coConex As New COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String

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
'madm 20091228 -----------------------------
Public Function dObtieneListaDescripcParametros(ByVal pnParamVar As String) As Recordset
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = "SELECT -1 nConsValor,'--SELECCIONE--' cConsDescripcion UNION SELECT nParamvar as nConsValor,cParamDesc as cConsDescripcion FROM ColocParametro WHERE nParamVar in  (" & pnParamVar & ") "

Set lrs = coConex.CargaRecordSet(lsSQL)
    Set dObtieneListaDescripcParametros = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Parametros Colocaciones <<dObtieneColocParametro>>", Err.Description
    
End Function
'------------------------------------------
Public Function dObtieneTasaInteres(ByVal psCodLinCred As String, ByVal pnTipoTasa As Integer, Optional ByVal pnMaxMin As Integer = 0) As Double
'pnMaxMin : se creo el parametro RECO20140912 ERS032-2014
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

coConex.AbreConexion

lsSQL = "SELECT nTasaIni, nTasaFin FROM ColocLineaCreditoTasa " _
      & "WHERE cLineaCred ='" & psCodLinCred & "' AND nColocLinCredTasaTpo =" & pnTipoTasa & " "

Set lrs = coConex.CargaRecordSet(lsSQL)

If lrs.BOF And lrs.EOF Then
    dObtieneTasaInteres = 0
Else
    dObtieneTasaInteres = lrs("nTasaIni")
    'RECO20140912 ERS032-2014***********************
    If pnMaxMin = 1 Then
        dObtieneTasaInteres = lrs("nTasaIni")
    Else
        dObtieneTasaInteres = lrs("nTasaFin")
    End If
    'FIN RECO***************************************
End If

Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Tasa Interes <<dObtieneColocParametro>>", Err.Description
    
End Function


Public Function dObtieneLineaCredito(ByVal pnCapital As Currency) As Recordset
Dim lsSQL As String

On Error GoTo ErrorLineaCredito

lsSQL = "SELECT cLineaCred, " & _
    "TasaComp = (SELECT nTasaIni FROM colocLineaCreditoTasa WHERE cLineaCred = LC.cLineaCred and nColocLinCredTasaTpo = 1), " & _
    "TasaMora = (SELECT nTasaIni FROM colocLineaCreditoTasa WHERE cLineaCred = LC.cLineaCred and nColocLinCredTasaTpo = 3) " & _
    "FROM ColocLineaCredito LC " & _
    "WHERE Substring(cLineaCred,7,3) = '305' AND (" & pnCapital & " Between nMontoMin AND nMontoMax) AND bEstado = 1"

Set dObtieneLineaCredito = coConex.CargaRecordSet(lsSQL)

Exit Function
    
ErrorLineaCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function dObtienePrecioMaterial(ByVal pnTipoVal, ByVal pnTipoMat As Integer, pnTipoMon As Integer) As Currency
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo ErrorPrecioMaterial

lsSQL = "SELECT TOP 1 nValor FROM ColocPigPrecioMaterial WHERE nTipoValor = " & pnTipoVal & " AND nMaterial = " & pnTipoMat & _
        " AND nMoneda = " & pnTipoMon & " AND DATEDIFF(d, dInicial, '" & Format(Date, "mm/dd/yyyy") & "') > 0 ORDER BY dInicial Desc"
Set lrs = coConex.CargaRecordSet(lsSQL)

    If lrs.EOF And lrs.BOF Then
        dObtienePrecioMaterial = 0
    Else
        dObtienePrecioMaterial = lrs!nValor
    End If
Set lrs = Nothing
    
Exit Function
    
ErrorPrecioMaterial:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function nNumRetasacion(ByVal psCtaCod As String) As Integer
Dim sSql As String
Dim RSTEMP As ADODB.Recordset
 Set RSTEMP = New ADODB.Recordset
 nNumRetasacion = 1

    On Error GoTo ErrorMensaje
    
    sSql = "Select nroret=case when max(nnumretasacion) is null then 1 else max(nnumretasacion)+1  end "
    sSql = sSql & " from HistColocPigJoyaDetRet where cctacod='" & psCtaCod & "'"
    RSTEMP.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTEMP.EOF And RSTEMP.BOF) Then
        nNumRetasacion = RSTEMP!nroret
    End If
    

Set coConex = Nothing
Set RSTEMP = Nothing
    
    Exit Function
    
ErrorMensaje:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Private Sub Class_Initialize()
    Dim loIni As New COMConecta.DCOMClasIni
  
    csConexion = loIni.CadenaConexion
    csNegocio = loIni.BaseNegocio
    csCentralPer = loIni.BasePersonas
    csCentralCom = loIni.BaseComunes
    csCentralImg = loIni.BaseImagenes
    csAdminist = loIni.BaseAdministracion
    If coConex.AbreConexion(csConexion) = False Then
        Call oError.RaiseError(oError.MyUnhandledError, "DColPContrato:Initialize. Error en Conexion a Base de datos")
    End If
End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub
'RECO20140623 ERS081-2014*************************************
Public Function dObtieneParamPignoCostoNotif(ByVal psParamDesc As String, ByVal psAgeCod As String) As Double
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = "stp_sel_DevuelveParamPignoCostoNotif '" & psParamDesc & "','" & psAgeCod & "'"

Set lrs = coConex.CargaRecordSet(lsSQL)

If lrs.BOF And lrs.EOF Then
    dObtieneParamPignoCostoNotif = 0
Else
    dObtieneParamPignoCostoNotif = lrs("nParamValor")
End If

Set lrs = Nothing

Exit Function
    
dError:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'RECO FIN*****************************************************

'*** PEAC 20161104
Public Function dObtieneParamCredRevolNueAnt(ByVal psCtaCod As String) As Boolean
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Set lrs = New Recordset
On Error GoTo dError

lsSQL = "stp_sel_ObtieneParamCredRevolNueAnt '" & psCtaCod & "' "
Set lrs = coConex.CargaRecordSet(lsSQL)

If lrs.BOF And lrs.EOF Then
    dObtieneParamCredRevolNueAnt = 0
Else
    dObtieneParamCredRevolNueAnt = lrs!bAntNueRevol
End If
Set lrs = Nothing

Exit Function
    
dError:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'PEAC 20200728
Public Function nObtieneDsctoInt(ByVal pnMonto As Double, ByVal psNroContrato As String) As Recordset

Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = "exec stp_sel_DsctoIntCampanaPignoRecupContratoAdju " & pnMonto & " ," & psNroContrato

Set lrs = coConex.CargaRecordSet(lsSQL)
    Set nObtieneDsctoInt = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Parametros para dsctos Int Pigno <<nObtieneDsctoInt>>", Err.Description

End Function
