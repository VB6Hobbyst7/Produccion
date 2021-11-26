VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAdeudCal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsConexion As String
Dim vsCentralPers As String
Dim vsCentralCom As String
Dim vsCentralImg As String
Dim sql As String
Dim oConec As DConecta

Public Sub inicio(poConec As ADODB.Connection)
Attribute inicio.VB_UserMemId = 0
oConec.ConexionActiva = poConec
End Sub
    
Public Sub InsertCalendario(psPersCod As String, psIFTpo As String, psCtaIfCod As String, psTpoCuota As String, pnNroCuota As Integer, pdVencimiento As Date, psEstado As String, pnDiasPago As Integer, pnCapital As Currency, pnInteres As Currency, pnInteresPagado As Currency)
On Error GoTo ErrorInsertCal

sql = "INSERT INTO CtaIFCalendario(cPersCod, cIFTpo, cCtaIFCod, cTpoCuota, nNroCuota , dVencimiento, cEstado, nDiasPago, nCapital, nInteres, nInteresPagado ) " _
    & "VALUES ('" & psPersCod & "','" & psIFTpo & "','" & psCtaIfCod & "','" & psTpoCuota & "'," & pnNroCuota & ",'" & Format(pdVencimiento, gsFormatoFecha) & "','" & psEstado & "'," & pnDiasPago & ", " & pnCapital & "," & pnInteres & "," & pnInteresPagado & ")"
oConec.Ejecutar sql
Exit Sub
ErrorInsertCal:
    Err.Raise Err.Number, "DAdeudCal:InsertCalendario", Err.Description
End Sub
Public Sub UpdateCalendario(ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIfCod As String, ByVal pnNroCuota As Integer, ByVal pnInteresPagado As Currency, psMovNro As String, Optional pnEstado As CGTipoEstCuotaAdeud = -1, Optional pbCancela As Boolean = False)
Dim lsActCampo As String
On Error GoTo ErrorUpdateCal
If pnEstado > -1 Then
    lsActCampo = lsActCampo & ", cEstado = '" & pnEstado & "' "
End If
sql = "UPDATE CtaIFCalendario SET nInteresPagado = nInteresPagado + " & pnInteresPagado & ", cMovNro = '" & psMovNro & "' " & lsActCampo _
    & "WHERE  cPersCod='" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' AND cCtaIFCod = '" & psCtaIfCod & "' AND nNroCuota = " & pnNroCuota

oConec.Ejecutar sql
If lbCancela And pnEstado > -1 Then
    sql = "UPDATE CtaIFCalendario SET cMovNro = '" & psMovNro & "' " & lsActCampo _
        & "WHERE cPersCod='" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' AND cCtaIFCod = '" & psCtaIfCod & "' AND nNroCuota > " & pnNroCuota
End If
Exit Sub
ErrorUpdateCal:
    Err.Raise Err.Number, "dAdeudCal: UpdateCalendario", Err.Description
End Sub
Public Sub EliminarCalendario(psPersCod As String, psCtaIfCod As String, psTpoCuota As String, pnNroCuota As Integer)
On Error GoTo ErrorEliminarCal

sql = "DELETE CtaIFCalendario " _
    & "WHERE  cPersCod='" & psPersCod & "' AND cCtaIFCod='" & psCtaIfCod & "' AND cTpoCuota= '" & psTpoCuota & "' AND nNroCuota= " & pnNroCuota & ""

oConec.Ejecutar sql
Exit Sub
ErrorEliminarCal:
    Err.Raise Err.Number, "dAdeudCal", Err.Description
End Sub
Public Sub InsertaCalendDet(psPersCod As String, psIFTpo As String, psCtaIfCod As String, psTpoCuota As String, pnNroCuota As Integer, psAdeudConcepto As String, pnMonto As Currency, pnMontoPago As Currency, psFlag As String)
On Error GoTo ErrorInsertCalDet
sql = "INSERT INTO CtaIFCalenDet(cPersCod, cIFTpo, cCtaIFCod, cTpoCuota, nNroCuota, cAdeudConcepto, nMonto, nMontoPagoMN, nMontoPagoME, cFlag) " _
    & "VALUES('" & psPersCod & "','" & psIFTpo & "', '" & psCtaIfCod & "','" & psTpoCuota & "'," & pnNroCuota & ",'" & psAdeudConcepto & "'," _
    & pnMonto & "," & pnMontoPago & "," & pnMontoPago & ",'" & psFlag & "')"

oConec.Ejecutar sql
Exit Sub
ErrorInsertCalDet:
    Err.Raise Err.Number, "dAdeudCal", Err.Description
End Sub
Public Sub UpdateCalendDet(psPersCod As String, psCtaIfCod As String, psTpoCuota As String, pnNroCuota As Integer, psAdeudConcepto As String, pnMonto As Currency, pnMontoPago As Currency, psFlag As String)
On Error GoTo ErrorUpdateCalDet
sql = "UPDATE CtaIFCalenDet SET cAdeudConcepto='" & psAdeudConcepto & "',nMonto=" & pnMonto & ", nMontoPago=" & pnMontoPago & ", cFlag='" & psFlag & "' " _
    & "WHERE cPersCod='" & psPersCod & "' and cCtaIFCod='" & psCtaIfCod & "' and cTpoCuota='" & psTpoCuota & "' and nNroCuota=" & pnNroCuota

oConec.Ejecutar sql

Exit Sub
ErrorUpdateCalDet:
    Err.Raise Err.Number, "dAdeudCal", Err.Description
End Sub
Public Sub EliminarCalendDet(psPersCod As String, psCtaIfCod As String, psTpoCuota As String, pnNroCuota As Integer)
On Error GoTo ErrorEliminarCalDet
sql = "DELETE CtaIFCalenDet " _
    & "WHERE cPersCod='" & psPersCod & "' and cCtaIFCod='" & psCtaIfCod & "' and cTpoCuota='" & psTpoCuota & "' and nNroCuota=" & pnNroCuota

oConec.Ejecutar sql
Exit Sub
ErrorEliminarCalDet:
    MsgBox "Error N°[" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Public Sub InsertaCtaIFAdeudado(ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIfCod As String, ByVal pnMontoPrestado As Currency, _
                                ByVal pnSaldoCap As Currency, ByVal pnNroCuotas As Integer, ByVal pnPeriodoGracia As Currency, ByVal pnComisionInicial As Currency, _
                                ByVal psPlaza As String, ByVal pnCuotaPagoCap As Integer, ByVal psVencimiento As String, _
                                ByVal pnEjecutado As Integer, ByVal psMonedaPago As String, ByVal psCuotaUltPago As String)
                                
On Error GoTo ErrorInsertCtaIFAdeud
sql = "INSERT INTO CTAIFADEUDADOS(cPersCod, cIFTpo, cCtaIFCod, nMontoPrestado, nCtaIFCuotas, nSaldoCap, nPeriodoGracia, nComisionInicial, cPlaza , nCuotaPagoCap, dVencimiento, bEjecutado, cMonedaPago, cCodLinCred, dCuotaUltPago ) " _
    & "VALUES('" & psPersCod & "','" & psIFTpo & "', '" & psCtaIfCod & "'," & pnMontoPrestado & ", " & pnNroCuotas & ", " & pnSaldoCap & "," & pnPeriodoGracia & "," & pnComisionInicial & ",'" _
    & psPlaza & "'," & pnCuotaPagoCap & ",'" & psVencimiento & "'," & pnEjecutado & ",'" & psMonedaPago & "', '', '" & psCuotaUltPago & "' )"
oConec.Ejecutar sql
Exit Sub
ErrorInsertCtaIFAdeud:
    Err.Raise Err.Number, "dAdeudCal", Err.Description
End Sub

Public Sub UpdateCtaIFAdeudado(ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIfCod As String, _
                                Optional ByVal pnMontoPrestado As Currency = -1, _
                                Optional ByVal pnSaldoCap As Currency = -1, Optional ByVal pnNroCuotas As Integer = -1, Optional ByVal pnPeriodoGracia As Currency = -1, Optional ByVal pnComisionInicial As Currency = -1, _
                                Optional ByVal psPlaza As String = "", Optional ByVal pnCuotaPagoCap As Integer = -1, Optional ByVal psVencimiento As String = "", _
                                Optional ByVal pnEjecutado As Integer = -1, Optional ByVal psMonedaPago As String = "", Optional ByVal psCuotaUltPago As String = "")
Dim lsDatos As String
On Error GoTo ErrorUpdateCtaIFAdeud
lsDatos = ""
If pnMontoPrestado <> -1 Then
    lsDatos = lsDatos & " nMontoPrestado=" & pnMontoPrestado & ","
End If
If pnSaldoCap <> -1 Then
    lsDatos = lsDatos & " nSaldoCap = nSaldoCap - " & pnSaldoCap & ","
End If
If pnPeriodoGracia <> -1 Then
    lsDatos = lsDatos & " nPeriodoGracia=" & pnPeriodoGracia & ","
End If
If pnComisionInicial <> -1 Then
    lsDatos = lsDatos & " nComisionInicial= " & pnComisionInicial & ","
End If
If psPlaza <> "" Then
    lsDatos = lsDatos & " cPlaza= '" & psPlaza & "',"
End If
If pnCuotaPagoCap <> -1 Then
    lsDatos = lsDatos & " nCuotaPagoCap=" & pnCuotaPagoCap & ","
End If
If psVencimiento <> "" Then
    lsDatos = lsDatos & "dVencimiento='" & psVencimiento & "',"
End If
If pnEjecutado <> -1 Then
    lsDatos = lsDatos & "bEjecutado=" & pnEjecutado & ","
End If
If psMonedaPago <> "" Then
    lsDatos = lsDatos & "cMonedaPago= '" & psMonedaPago & "',"
End If
If psCuotaUltPago <> "" Then
    lsDatos = lsDatos & "dCuotaUltPago = '" & psCuotaUltPago & "',"
End If
If lsDatos <> "" Then
    lsDatos = Left(lsDatos, Len(lsDatos) - 1)
    sql = "UPDATE CTAIFADEUDADOS SET " & lsDatos _
        & "WHERE  cPersCod = '" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' and cCtaIfCod = '" & psCtaIfCod & "'"
    oConec.Ejecutar sql
End If
Exit Sub
ErrorUpdateCtaIFAdeud:
   Err.Raise Err.Number, "dAdeudCal", Err.Description
End Sub
Public Sub EliminarCtaIFAdeudado(psPersCod As String, psIFTpo As String, psCtaIfCod As String)
On Error GoTo ErrorEliminarCtaIFAdeud

sql = "DELETE CTAIFADEUDADOS " _
    & "WHERE    cPersCod = '" & psPersCod & "' and cIFTpo = '" & psIFTpo & "'  and cCtaIfCod='" & psCtaIfCod & "'"

oConec.Ejecutar sql
Exit Sub
ErrorEliminarCtaIFAdeud:
    MsgBox "Error N°[" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
End Sub
Public Function VerificaCtaIFAdeudado(psPersCod As String, psIFTpo As String, psCtaIfCod As String) As Boolean
Dim rs As ADODB.Recordset
On Error GoTo ErrorVerCtaIFAdeud

VerificaCtaIFAdeudado = False

sql = "SELECT * FROM CTAIFADEUDADOS " _
    & "WHERE    cPersCod = '" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' and  cCtaIfCod='" & psCtaIfCod & "'"

Set rs = oConec.CargaRecordSet(sql)
VerificaCtaIFAdeudado = Not rs.EOF
RSClose rs
oConec.Ejecutar sql
Exit Function
ErrorVerCtaIFAdeud:
    Err.Raise Err.Number, "dAdeudCal", Err.Description
End Function
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni
vsCentralPers = oIni.BasePersonas
vsCentralCom = oIni.BaseComunes
vsCentralImg = oIni.BaseImagenes

Set oConec = New DConecta
If oConec.AbreConexion() = False Then
    Exit Sub
End If
End Sub
