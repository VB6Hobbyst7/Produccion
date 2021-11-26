VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NEncajeBCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

'Public Function GetParametroEncaje(Optional pnCodigo As Long = -1) As ADODB.Recordset
'psSql = "SELECT * FROM ParamEncaje " & IIf(pnCodigo = -1, "", " WHERE nCodigo = " & pnCodigo)
'Set GetParametroEncaje = dbConec.CargaRecordSet(psSql)
'End Function

Public Function GetParametroEncaje(pdFecha As Date, Optional pnCodigo As Long = -1) As ADODB.Recordset

Dim sqlAdd As String

If pnCodigo = -1 Then
   sqlAdd = " Where dFecha = (Select Max(dFecha) FROM ParamEncaje  Where dFecha <= '" & Format(pdFecha, gsFormatoFecha) & "') Order By nCodigo "
Else
   sqlAdd = " And dFecha = (Select Max(dFecha) FROM ParamEncaje  Where dFecha <= '" & Format(pdFecha, gsFormatoFecha) & "') Order By nCodigo "
End If

psSql = "SELECT nCodigo, cDescripcion, nValor FROM ParamEncaje " & IIf(pnCodigo = -1, "", " WHERE nCodigo = " & pnCodigo) & sqlAdd

Set GetParametroEncaje = dbConec.CargaRecordSet(psSql)

End Function


Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As New ClasIni
   Set oIni = Nothing
   
   Set dbConec = New DConecta
   dbConec.AbreConexion
End Sub

Private Sub Class_Terminate()
   dbConec.CierraConexion
   Set dbConec = Nothing
End Sub
'ALPA 20080811**********
Public Function ObtenerRepColumnaCol(Optional cCodOpe As String = "") As ADODB.Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_ObtenerRepColumnaCol '" & cCodOpe & "'"
    Set ObtenerRepColumnaCol = dbConec.CargaRecordSet(sqlEnc)
End Function
'***********************
'MIOL 20120828, SEGUN RQ11368 **************************************************
Public Function ObtenerParamEncajeDiario() As ADODB.Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_ParamEncDiario"
    Set ObtenerParamEncajeDiario = dbConec.CargaRecordSet(sqlEnc)
End Function

Public Function ObtenerParamEncajeDiarioxCod(ByVal cCodParamEnc As String, Optional psMoneda As String = "", Optional psFecha As String = "") As ADODB.Recordset 'NAGL Agregó psMoneda 20181001 y psFecha 20181015
    Dim sqlEnc As String
    sqlEnc = "Exec stp_sel_ParamEncDiarioxCod '" & cCodParamEnc & "','" & psMoneda & "','" & psFecha & "'" 'NAGL Agregó psMoneda 20181001 y psFecha 20181015
    Set ObtenerParamEncajeDiarioxCod = dbConec.CargaRecordSet(sqlEnc)
End Function

Public Function UpdateParamEncajeDiario(ByVal nMonto As Double, ByVal cMovParEncDiario As String)
    Dim sqlEnc As String
    sqlEnc = "stp_upd_ParamEncDiario '" & nMonto & "','" & cMovParEncDiario & "'"
    Set UpdateParamEncajeDiario = dbConec.CargaRecordSet(sqlEnc)
End Function

Public Function InsParamEncajeDiario(ByVal cCodParEncDiario As String, ByVal cMovParEncDiario As String, ByVal nMonto As Double, ByVal dFecha As Date)
    Dim sqlEnc As String
    sqlEnc = "stp_ins_MovParamEncajeDiario '" & cCodParEncDiario & "','" & cMovParEncDiario & "','" & nMonto & "','" & Format(dFecha, "yyyymmdd") & "'"
    Set InsParamEncajeDiario = dbConec.CargaRecordSet(sqlEnc)
End Function

Public Function ObtenerSaldoAhoPlaFijCTS(ByVal cFecha As String, ByVal nMoneda As Integer, ByVal cProducto As String) As Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_AhorrosPlazoFijoCTS '" & cFecha & "','" & nMoneda & "','" & cProducto & "'"
    Set ObtenerSaldoAhoPlaFijCTS = dbConec.CargaRecordSet(sqlEnc)
End Function

Public Function ObtenerSaldoChequeAhoPlaFijCTS(ByVal cFecha As String, ByVal nMoneda As Integer, ByVal cProducto As String) As Recordset
    Dim sqlEnc As String
    If cProducto = "232" Then
        sqlEnc = "stp_sel_AhorrosSaldoChequesAnexoDiario '" & cFecha & "','" & nMoneda & "'"
    End If
    If cProducto = "233" Then
        sqlEnc = "stp_sel_PFSaldoChequesAnexoDiario '" & cFecha & "','" & nMoneda & "'"
    End If
    If cProducto = "234" Then
        sqlEnc = "stp_sel_CTSSaldoChequesAnexoDiario '" & cFecha & "','" & nMoneda & "'"
    End If
    Set ObtenerSaldoChequeAhoPlaFijCTS = dbConec.CargaRecordSet(sqlEnc)
End Function

Public Function ObtenerFechaAnioMesDia(ByVal sFecha As String) As Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_FechaAnioMesDia '" & sFecha & "'"
    Set ObtenerFechaAnioMesDia = dbConec.CargaRecordSet(sqlEnc)
End Function

Public Function ObtenerSaldoBCRPEncajeDiario(ByVal cFecha As String, ByVal nMoneda As Integer) As Recordset
    Dim sqlEnc As String
    If nMoneda = 1 Then
        sqlEnc = "stp_SaldoBRCP_MN_AnexoDiario '" & cFecha & "'"
    Else
        sqlEnc = "stp_SaldoBRCP_ME_AnexoDiario '" & cFecha & "'"
    End If
    Set ObtenerSaldoBCRPEncajeDiario = dbConec.CargaRecordSet(sqlEnc)
End Function

Public Function ObtenerSaldoCajasCracsEncajeDiario(ByVal cFecha As String, ByVal nMoneda As Integer, ByVal cProducto As String) As Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_CajasCracsAnexoDiario '" & cFecha & "','" & nMoneda & "','" & cProducto & "'"
    Set ObtenerSaldoCajasCracsEncajeDiario = dbConec.CargaRecordSet(sqlEnc)
End Function

Public Function ObtenerSaldoCajasObligExoneradas(ByVal cFecha As String, ByVal nMoneda As Integer) As Recordset
    Dim sqlEnc As String
    If nMoneda = 1 Then
        sqlEnc = "stp_sel_DescompEfectivoMNAnexoDiario '" & cFecha & "'"
    End If
    If nMoneda = 2 Then
        sqlEnc = "stp_sel_DescompEfectivoMEAnexoDiario '" & cFecha & "'"
    End If
    Set ObtenerSaldoCajasObligExoneradas = dbConec.CargaRecordSet(sqlEnc)
End Function

Public Function ObtenerIntDevEncajeDiario(ByVal cFecha As String, ByVal nMoneda As Integer) As Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_SaldoInteresDevengadoAnexoDiario '" & cFecha & "','" & nMoneda & "'"
    Set ObtenerIntDevEncajeDiario = dbConec.CargaRecordSet(sqlEnc)
End Function

Public Function ObtenerPlazoFijoxRangoEncajeDiario(ByVal cFecha As String, ByVal nMoneda As Integer) As Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_PlazoFijoxRangoAnexoDiario '" & cFecha & "','" & nMoneda & "'"
    Set ObtenerPlazoFijoxRangoEncajeDiario = dbConec.CargaRecordSet(sqlEnc)
End Function

Public Function ObtenerSaldoCoopAhoEncajeDiario(ByVal cFecha As String, ByVal nMoneda As Integer, ByVal cProducto As String) As Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_CoopeAhorros '" & cFecha & "','" & nMoneda & "','" & cProducto & "'"
    Set ObtenerSaldoCoopAhoEncajeDiario = dbConec.CargaRecordSet(sqlEnc)
End Function

Public Function ObtenerSaldoCracsEncajeDiario(ByVal cFecha As String, ByVal nMoneda As Integer, ByVal cProducto As String) As Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_CracsPlaFijAnexoDiario '" & cFecha & "','" & nMoneda & "','" & cProducto & "'"
    Set ObtenerSaldoCracsEncajeDiario = dbConec.CargaRecordSet(sqlEnc)
End Function

Public Function ObtenerSaldoAhoExoEnc() As Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_ValAhoExoEncaje"
    Set ObtenerSaldoAhoExoEnc = dbConec.CargaRecordSet(sqlEnc)
End Function
'END MIOL **********************************************************************
'*** PASI 20140408
Public Function InsertaCalenProyEncaje(ByVal pdCalFecha As Date, ByVal pnCalDepPla As Currency, ByVal pnCalDepAho As Currency, ByVal pnCalDepBcrp As Currency, ByVal pnOblInm As Currency, ByVal pnEfecCaja As Currency, ByVal pnMoneda As Integer)
    Dim sqlEnc As String
    sqlEnc = "stp_ins_InsertarCalendProyEncaje '" & Format(pdCalFecha, "yyyyMMdd") & "'," & pnCalDepPla & "," & pnCalDepAho & "," & pnCalDepBcrp & "," & pnOblInm & "," & pnEfecCaja & "," & pnMoneda
    Set InsertaCalenProyEncaje = dbConec.CargaRecordSet(sqlEnc)
End Function
Public Function ObtenerCalenProyEncaje(ByVal pdCalFechaIni As Date, ByVal dCalFechaFin As Date, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_ObtenerCalendProyEncaje '" & Format(pdCalFechaIni, "yyyyMMdd") & "','" & Format(dCalFechaFin, "yyyyMMdd") & "'," & pnMoneda
    Set ObtenerCalenProyEncaje = dbConec.CargaRecordSet(sqlEnc)
End Function
Public Function ActualizaCalenProyEncaje(ByVal pnCalCod As Integer, ByVal pnnCalDepPla As Currency, ByVal pnnCalDepAho As Currency, ByVal pnCalDepBcrp As Currency, ByVal pnOblInm As Currency, ByVal pnEfecCaja As Currency)
    Dim sqlEnc As String
    sqlEnc = "stp_upd_ActualizaCalendProyEncaje " & pnCalCod & "," & pnnCalDepPla & "," & pnnCalDepAho & "," & pnCalDepBcrp & "," & pnOblInm & "," & pnEfecCaja
    Set ActualizaCalenProyEncaje = dbConec.CargaRecordSet(sqlEnc)
End Function
Public Function ObtenerSaldoEdypimesEncajeDiario(ByVal pdFecha As Date, ByVal pnMoneda As Integer, ByVal pnProducto As Integer) As ADODB.Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_SaldoEdypimesAnexoDiario '" & Format(pdFecha, "yyyyMMdd") & "'," & pnMoneda & "," & pnProducto
    Set ObtenerSaldoEdypimesEncajeDiario = dbConec.CargaRecordSet(sqlEnc)
End Function
Public Function ObtenerSaldoDepPlaCalenProy(ByVal pdFecha As Date, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_SaldoDepPlaCalendProy '" & Format(pdFecha, "yyyyMMdd") & "'," & pnMoneda
    Set ObtenerSaldoDepPlaCalenProy = dbConec.CargaRecordSet(sqlEnc)
End Function
Public Function ObtenerSaldoDepAhoCalenProy(ByVal pdFecha As Date, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_SaldoDepAhoCalendProy '" & Format(pdFecha, "yyyyMMdd") & "'," & pnMoneda
    Set ObtenerSaldoDepAhoCalenProy = dbConec.CargaRecordSet(sqlEnc)
End Function
Public Function ObtenerSaldoDepBcrpCalenProy(ByVal pdFecha As Date, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_SaldoDepBcrpCalendProy '" & Format(pdFecha, "yyyyMMdd") & "'," & pnMoneda
    Set ObtenerSaldoDepBcrpCalenProy = dbConec.CargaRecordSet(sqlEnc)
End Function
Public Function ObtenerSaldoOblInmCalenProy(ByVal pdFecha As Date, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_SaldoOblInmCalendProy '" & Format(pdFecha, "yyyyMMdd") & "'," & pnMoneda
    Set ObtenerSaldoOblInmCalenProy = dbConec.CargaRecordSet(sqlEnc)
End Function
Public Function ObtenerSaldoEfecCajaCalenProy(ByVal pdFecha As Date, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim sqlEnc As String
    sqlEnc = "stp_sel_SaldoEfecCajaCalendProy '" & Format(pdFecha, "yyyyMMdd") & "'," & pnMoneda
    Set ObtenerSaldoEfecCajaCalenProy = dbConec.CargaRecordSet(sqlEnc)
End Function
'***END PASI


