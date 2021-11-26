VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMRCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim OCon As New COMConecta.DCOMConecta

Public Function ObtenerIBM1(ByVal psServerConsol As String, ByVal i As Integer) As ADODB.Recordset
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dError
   lsSQL = "Select count(*) as Total from " & psServerConsol & " IBM1 where cTipDoc='" & i & "'"
   OCon.AbreConexion
   Set rs = OCon.CargaRecordSet(lsSQL)
      Set ObtenerIBM1 = rs
   Set rs = Nothing
   Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function

Public Function ObtenernVen30MN(ByVal psServerConsol As String) As ADODB.Recordset
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dError
   lsSQL = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen30MN>0"
   OCon.AbreConexion
   Set rs = OCon.CargaRecordSet(lsSQL)
      Set ObtenernVen30MN = rs
   Set rs = Nothing
   Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function

Public Function ObtenernVen31MN(ByVal psServerConsol As String) As ADODB.Recordset
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dError
   lsSQL = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen31MN>0"
   OCon.AbreConexion
   Set rs = OCon.CargaRecordSet(lsSQL)
      Set ObtenernVen31MN = rs
   Set rs = Nothing
   Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function


Public Function ObtenerTotalCalifica(ByVal psServerConsol As String, ByVal i As Integer) As ADODB.Recordset
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dError
   lsSQL = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='" & i & "'"
   OCon.AbreConexion
   Set rs = OCon.CargaRecordSet(lsSQL)
      Set ObtenerTotalCalifica = rs
   Set rs = Nothing
   Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function

Public Function ObtenerDeudaVigente(ByVal psServerConsol As String, ByVal pnNorm As String) As ADODB.Recordset
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dError
   lsSQL = "SELECT replace(convert(varchar(20),SUM( " & pnNorm & ")),'.','')  AS TOTAL From " & psServerConsol & "IBM1"
  
   OCon.AbreConexion
   Set rs = OCon.CargaRecordSet(lsSQL)
      Set ObtenerDeudaVigente = rs
   Set rs = Nothing
   Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function

Public Function ObtenerDeudaDirecRefinanciada(ByVal psServerConsol As String, ByVal pnRef As String) As ADODB.Recordset
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dError
   lsSQL = "SELECT replace(convert(varchar(20),SUM(" & pnRef & " )),'.','') AS TOTAL From " & psServerConsol & "IBM1"
   OCon.AbreConexion
   Set rs = OCon.CargaRecordSet(lsSQL)
      Set ObtenerDeudaDirecRefinanciada = rs
   Set rs = Nothing
   Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function

Public Function ObtenerDeudaDirecVencida30(ByVal psServerConsol As String, ByVal pnVen30 As String) As ADODB.Recordset
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dError
   lsSQL = "SELECT replace(convert(varchar(20),SUM(" & pnVen30 & " )),'.','') AS TOTAL From " & psServerConsol & "IBM1"
   OCon.AbreConexion
   Set rs = OCon.CargaRecordSet(lsSQL)
      Set ObtenerDeudaDirecVencida30 = rs
   Set rs = Nothing
   Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function

Public Function ObtenerDeudaDirecVencida31(ByVal psServerConsol As String, ByVal pnVen31 As String) As ADODB.Recordset
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dError
   lsSQL = "SELECT replace(convert(varchar(20),SUM( " & pnVen31 & " )),'.','') AS TOTAL From " & psServerConsol & "IBM1"
   OCon.AbreConexion
   Set rs = OCon.CargaRecordSet(lsSQL)
      Set ObtenerDeudaDirecVencida31 = rs
   Set rs = Nothing
   Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function
Public Function ObtenerDeudaDirecJudicial(ByVal psServerConsol As String, ByVal pnCobJud As String) As ADODB.Recordset
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dError
   lsSQL = "SELECT replace(convert(varchar(20),SUM( " & pnCobJud & " )),'.','') AS TOTAL From " & psServerConsol & "IBM1"
   OCon.AbreConexion
   Set rs = OCon.CargaRecordSet(lsSQL)
      Set ObtenerDeudaDirecJudicial = rs
   Set rs = Nothing
   Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function

Public Function ObtenerLineasCredito(ByVal psServerConsol As String, ByVal nLinCred As String) As ADODB.Recordset
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dError
   lsSQL = "SELECT replace(convert(varchar(20),SUM( " & nLinCred & " )),'.','') AS TOTAL From " & psServerConsol & "IBM1"
   OCon.AbreConexion
   Set rs = OCon.CargaRecordSet(lsSQL)
      Set ObtenerLineasCredito = rs
   Set rs = Nothing
   Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function

Public Function ObtenerCreditosCastigados(ByVal psServerConsol As String, ByVal nCast As String) As ADODB.Recordset
  Dim lsSQL As String
  Dim rs As New ADODB.Recordset
On Error GoTo dError
   lsSQL = "SELECT replace(convert(varchar(20),SUM(" & nCast & " )),'.','') AS TOTAL From " & psServerConsol & "IBM1"
   OCon.AbreConexion
   Set rs = OCon.CargaRecordSet(lsSQL)
      Set ObtenerCreditosCastigados = rs
   Set rs = Nothing
   Exit Function
dError:
    Err.Raise Err.Number, "Error al cargar los Datos", Err.Description
End Function




