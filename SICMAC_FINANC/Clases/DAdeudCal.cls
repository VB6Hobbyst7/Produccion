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

Public Sub Inicio(poConec As ADODB.Connection)
oConec.ConexionActiva = poConec
End Sub

'ALPA20130614****************
Public Sub InsertCalendario(psPersCod As String, psIFTpo As CGTipoIF, psCtaIFCod As String, psTpoCuota As String, pnNroCuota As Integer, pdVencimiento As Date, psEstado As String, pnDiasPago As Integer, pnCapital As Currency, pnInteres As Currency, pnComision As Currency, pnInteresPagado As Currency, pnTotalCuota As Currency, psMovNro As String, pnVigente As Integer, psUltimaActualizacion As String, Optional ByVal pnCapitalConce As Currency = 0#, Optional ByVal pnInteresConce As Currency = 0#, Optional ByVal pnComisionConce As Currency = 0#)
On Error GoTo ErrorInsertCal

sql = "INSERT INTO CtaIFCalendario(cPersCod, cIFTpo, cCtaIFCod, cTpoCuota, nNroCuota , dVencimiento, cEstado, nDiasPago, nCapital, nInteres, nComision, nInteresPagado,nTotalCuota, cMovNro,bvigente, cUltimaActualizacion,nCapitalConce,nInteresConce,nComisionConce ) " _
    & "VALUES ('" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psCtaIFCod & "','" & psTpoCuota & "'," & pnNroCuota & ",'" & Format(pdVencimiento, gsFormatoFecha) & "','" & psEstado & "'," & pnDiasPago & ", " & pnCapital & "," & pnInteres & "," & pnComision & "," & pnInteresPagado & "," & pnTotalCuota & ",'" & psMovNro & "'," & pnVigente & ",'" & psUltimaActualizacion & "', " & pnCapitalConce & "," & pnInteresConce & "," & pnComisionConce & ")" 'ALPA20130614
oConec.Ejecutar sql
Exit Sub
ErrorInsertCal:
    Err.Raise Err.Number, "DAdeudCal:InsertCalendario", Err.Description
End Sub

Public Sub UpdateCtaIFAdeudado(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, ByVal psCtaIFCod As String, _
                                Optional ByVal pnMontoPrestado As Currency = -1, _
                                Optional ByVal pnSaldoCap As Currency = -1, Optional ByVal pnNroCuotas As Integer = -1, Optional ByVal pnPeriodoGracia As Currency = -1, Optional ByVal pnComisionInicial As Currency = -1, Optional ByVal pnComisionMonto As Currency = -1, _
                                Optional ByVal psPlaza As String = "", Optional ByVal pnCuotaPagoCap As Integer = -1, Optional ByVal psVencimiento As String = "", _
                                Optional ByVal pnEjecutado As Integer = -1, Optional ByVal psMonedaPago As String = "", Optional ByVal psCuotaUltPago As String = "", Optional ByVal pnTpoCuota As CGAdeudCalTpoCuota = -1, Optional pnTramo As Currency = -1, _
                                Optional pnComisionCuota As Currency = -1, Optional psLinCredCod As String = "", Optional pnFechaFija As Integer = -1, Optional pbAfectaSaldoCap As Boolean = False, _
                                Optional pnSaldoCapLP As Currency = -1, Optional ByVal pnPorcentAfect As Currency = -1, Optional ByVal pnMalPg As Integer = -1, Optional pnMontoEuros As Currency, _
                                Optional pnMontoPresConce As Currency = 0, Optional pnSaldoCapitalConce As Currency = 0, Optional lbGrabarCapitalConcesionado As Boolean = False) 'pnMontoPresConce 20130614*******
                                
Dim lsDatos As String
Dim sql As String
On Error GoTo ErrorUpdateCtaIFAdeud
lsDatos = ""
If pnMontoPrestado <> -1 Then
    lsDatos = lsDatos & " nMontoPrestado=" & pnMontoPrestado & ","
End If

If pnSaldoCap <> -1 Then
    If pbAfectaSaldoCap Then
        lsDatos = lsDatos & " nSaldoCap = nSaldoCap - " & pnSaldoCap & ","
    Else
        lsDatos = lsDatos & " nSaldoCap = " & pnSaldoCap & ","
    End If
End If
If pnPeriodoGracia <> -1 Then
    lsDatos = lsDatos & " nPeriodoGracia=" & pnPeriodoGracia & ","
End If
If pnComisionInicial <> -1 Then
    lsDatos = lsDatos & " nComisionInicial= " & pnComisionInicial & ","
End If
If pnComisionMonto <> -1 Then
    lsDatos = lsDatos & " nComisionMonto= " & pnComisionMonto & ","
End If
If psPlaza <> "" Then
    lsDatos = lsDatos & " cPlaza= '" & psPlaza & "',"
End If
If pnCuotaPagoCap <> -1 Then
    lsDatos = lsDatos & " nCuotaPagoCap=" & pnCuotaPagoCap & ","
End If
If psVencimiento <> "" Then
    lsDatos = lsDatos & " dVencimiento='" & psVencimiento & "',"
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
If pnTpoCuota <> -1 Then
    lsDatos = lsDatos & " nTpoCuota = " & pnTpoCuota & ","
End If
If pnTramo <> -1 Then
    lsDatos = lsDatos & " nTramoConcesion = " & pnTramo & ","
End If
If pnComisionCuota <> -1 Then
    lsDatos = lsDatos & " nComisionCuota = " & pnComisionCuota & ","
End If
If psLinCredCod <> "" Then
    lsDatos = lsDatos & " cCodLinCred = '" & psLinCredCod & "',"
End If
If pnFechaFija <> -1 Then
    lsDatos = lsDatos & " nFechaFija = " & pnFechaFija & ","
End If
If pnNroCuotas <> -1 Then
    lsDatos = lsDatos & " nCtaIFCuotas = " & pnNroCuotas & ","
End If
If pnSaldoCapLP <> -1 Then
    lsDatos = lsDatos & " nSaldoCapLP = " & pnSaldoCapLP & ","
End If
If pnPorcentAfect <> -1 Then
    lsDatos = lsDatos & " nPorcentAfect = " & pnPorcentAfect & ","
End If
If pnMalPg <> -1 Then
    lsDatos = lsDatos & " bMalPg = " & pnMalPg & ","
End If
If pnMontoEuros <> -1 Then
    lsDatos = lsDatos & " nMontoEuros = " & pnMontoEuros & ","
End If
'ALPA20130614***************************************************************
If pnMontoPresConce <> -1 Then
    lsDatos = lsDatos & " nMontoPresConce = " & pnMontoPresConce & ","
End If
If pnSaldoCapitalConce <> -1 And lbGrabarCapitalConcesionado = True Then
    lsDatos = lsDatos & " nSaldoCapConce = " & pnSaldoCapitalConce & ","
End If

'***************************************************************************
If lsDatos <> "" Then
    lsDatos = Left(lsDatos, Len(lsDatos) - 1)
    sql = "UPDATE CTAIFADEUDADOS SET " & lsDatos _
        & " WHERE  cPersCod = '" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' and cCtaIfCod = '" & psCtaIFCod & "'"
    oConec.Ejecutar sql

End If
Exit Sub
ErrorUpdateCtaIFAdeud:
   Err.Raise Err.Number, "dAdeudCal", Err.Description
End Sub


Public Sub EliminarCalendario(psPersCod As String, psCtaIFCod As String, Optional psTpoCuota As String = "", Optional pnNroCuota As Integer = -1)
Dim lsFiltro As String
On Error GoTo ErrorEliminarCal
If Not psTpoCuota = "" Then
    lsFiltro = " and cTpoCuota = '" & psTpoCuota & "' "
End If
If Not pnNroCuota = -1 Then
    lsFiltro = lsFiltro & " and nNroCuota = " & pnNroCuota
End If

sql = "DELETE CtaIFCalendario " _
    & "WHERE  cPersCod='" & psPersCod & "' AND cCtaIFCod='" & psCtaIFCod & "' " & lsFiltro
oConec.Ejecutar sql
Exit Sub
ErrorEliminarCal:
    Err.Raise Err.Number, "dAdeudCal", Err.Description
End Sub

Public Sub InsertaCtaIFAdeudado(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, ByVal psCtaIFCod As String, ByVal pnMontoPrestado As Currency, _
                                ByVal pnSaldoCap As Currency, ByVal pnNroCuotas As Integer, ByVal pnPeriodoGracia As Currency, ByVal pnComisionInicial As Currency, ByVal pnComisionMonto As Currency, _
                                ByVal psPlaza As String, ByVal pnCuotaPagoCap As Integer, ByVal psVencimiento As String, _
                                ByVal pnEjecutado As Integer, ByVal psMonedaPago As String, ByVal psCuotaUltPago As String, pnTpoCuota As CGAdeudCalTpoCuota, pnTramo As Currency, pnComisionCuota As Currency, psLinCredCod As String, pnFechaFija As Integer, ByVal pnPorcentAfect As Currency, ByVal pnMalPg As Integer, ByVal pnMontoEuros As Currency, _
                                Optional pnMontoPresConce As Currency = 0, Optional pnSaldoCapitalConce As Currency = 0) 'ALPA20130614*****
                                
                               
On Error GoTo ErrorInsertCtaIFAdeud
'ALPA 20110703**********************
sql = "INSERT INTO CTAIFADEUDADOS(cPersCod, cIFTpo, cCtaIFCod, nMontoPrestado,nMontoPrestadoReal, nCtaIFCuotas, nSaldoCap, nPeriodoGracia, nComisionInicial, nComisionMonto, cPlaza , nCuotaPagoCap, dVencimiento, bEjecutado, cMonedaPago, cCodLinCred, dCuotaUltPago, nTpoCuota, nTramoConcesion, nComisionCuota, nFechaFija,nPorcentAfect,bMalPg,nMontoEuros,dCuotaUltModSaldos,nMontoPresConce,nSaldoCapConce ) " 'ALPA20130614*****
sql = sql & "VALUES('" & psPersCod & "','" & Format(psIFTpo, "00") & "', '" & psCtaIFCod & "'," & pnMontoPrestado & "," & pnSaldoCap & ", " & pnNroCuotas & ", " & pnSaldoCap & "," & pnPeriodoGracia & "," & pnComisionInicial & "," & pnComisionMonto & ",'"
sql = sql & psPlaza & "'," & pnCuotaPagoCap & ",'" & psVencimiento & "'," & pnEjecutado & ",'" & psMonedaPago & "', '" & psLinCredCod & "', '" & psCuotaUltPago & "', " & pnTpoCuota & "," & pnTramo & "," & pnComisionCuota & "," & pnFechaFija & "," & pnPorcentAfect & "," & pnMalPg & "," & pnMontoEuros & ",'" & Format(psCuotaUltPago, "YYYY/MM/DD") & "'," & pnMontoPresConce & "," & pnSaldoCapitalConce & ")" 'ALPA20130614
'***********************************
'sql = "INSERT INTO CTAIFADEUDADOS(cPersCod, cIFTpo, cCtaIFCod, nMontoPrestado,nMontoPrestadoReal, nCtaIFCuotas, nSaldoCap, nPeriodoGracia, nComisionInicial, nComisionMonto, cPlaza , nCuotaPagoCap, dVencimiento, bEjecutado, cMonedaPago, cCodLinCred, dCuotaUltPago, nTpoCuota, nTramoConcesion, nComisionCuota, nFechaFija,nPorcentAfect,bMalPg,nMontoEuros,dCuotaUltModSaldos ) " _
'    & "VALUES('" & psPersCod & "','" & Format(psIFTpo, "00") & "', '" & psCtaIfCod & "'," & pnMontoPrestado & "," & pnSaldoCap & ", " & pnNroCuotas & ", " & pnSaldoCap & "," & pnPeriodoGracia & "," & pnComisionInicial & "," & pnComisionMonto & ",'" _
'    & psPlaza & "'," & pnCuotaPagoCap & ",'" & psVencimiento & "'," & pnEjecutado & ",'" & psMonedaPago & "', '" & psLinCredCod & "', '" & psCuotaUltPago & "', " & pnTpoCuota & "," & pnTramo & "," & pnComisionCuota & "," & pnFechaFija & "," & pnPorcentAfect & "," & pnMalPg & "," & pnMontoEuros & ",'" & Format(psCuotaUltPago, "MM/dd/YYYY") & "')"
oConec.Ejecutar sql
Exit Sub
ErrorInsertCtaIFAdeud:
    Err.Raise Err.Number, "dAdeudCal", Err.Description
End Sub

Public Sub InsertaMalPagadorAdeudado(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, ByVal psCtaIFCod As String, _
                                 ByVal pnMalPg As Integer, ByVal psMovUltAct As String)
                                
Dim bMalPAnt As Integer
Dim rs As ADODB.Recordset
Dim sql As String
                                
On Error GoTo ErrorInsertCtaIFAdeud

sql = " Select convert(int,bMalPg) bMalPg " & _
      " FROM ctaIFAdeudados " & _
      " where cPErsCod='" & psPersCod & "' and cIFTpo='" & Format(psIFTpo, "00") & "' " & _
      " AND cCtaIFCod= '" & psCtaIFCod & "' "
Set rs = oConec.CargaRecordSet(sql)
If Not (rs.EOF And rs.BOF) Then
    bMalPAnt = rs!bMalPg
End If
      
sql = "INSERT INTO MalPagadorAdeudado(cPersCod,cIFTpo,cCtaIFCod,bMalPgIni,bMalPgFin,cUltimaActualizacion)" & _
      "VALUES('" & psPersCod & "', '" & Format(psIFTpo, "00") & "','" & psCtaIFCod & "',convert(bit,'" & bMalPAnt & " '),Convert(bit,'" & pnMalPg & "'),'" & psMovUltAct & "' )"
oConec.Ejecutar sql

Exit Sub
ErrorInsertCtaIFAdeud:
    Err.Raise Err.Number, "dAdeudCal", Err.Description
End Sub


Public Sub EliminarCtaIFAdeudado(psPersCod As String, psIFTpo As CGTipoIF, psCtaIFCod As String)
On Error GoTo ErrorEliminarCtaIFAdeud

sql = "DELETE CTAIFADEUDADOS " _
    & "WHERE    cPersCod = '" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "'  and cCtaIfCod='" & psCtaIFCod & "'"

oConec.Ejecutar sql
Exit Sub
ErrorEliminarCtaIFAdeud:
    MsgBox "Error N°[" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Public Function VerificaCtaIFAdeudado(psPersCod As String, psIFTpo As CGTipoIF, psCtaIFCod As String) As Boolean
Dim rs As ADODB.Recordset
On Error GoTo ErrorVerCtaIFAdeud

VerificaCtaIFAdeudado = False

sql = "SELECT cPersCod FROM CTAIFADEUDADOS " _
    & "WHERE    cPersCod = '" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' and  cCtaIfCod='" & psCtaIFCod & "'"

Set rs = oConec.CargaRecordSet(sql)
VerificaCtaIFAdeudado = Not rs.EOF
RSClose rs
Exit Function
ErrorVerCtaIFAdeud:
    Err.Raise Err.Number, "dAdeudCal", Err.Description
End Function

Public Sub UpdateCalendario(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, ByVal psCtaIFCod As String, ByVal pnNroCuota As Integer, ByVal pnTpoCuota As CGTipoCuotCalIF, ByVal pnInteresPagado As Currency, ByVal pnInteresPagadoReal As Currency, psMovNro As String, Optional pnEstado As CGTipoEstCuotaAdeud = -1, Optional pbCancela As Boolean = False, Optional bPago As Boolean = False, Optional pnNoCuota6 As String = "")
Dim lsActCampo As String
Dim rs As New ADODB.Recordset
Dim sql As String

On Error GoTo ErrorUpdateCal
If pnEstado > -1 Then
    lsActCampo = lsActCampo & ", cEstado = '" & pnEstado & "' "
End If

If bPago = True Or pbCancela = True Then
    sql = "UPDATE CtaIFCalendario SET cestado=1 , nInteresPagado = nInteresPagado + " & pnInteresPagado & ",nInteresPagado_Real = nInteresPagado_Real + " & pnInteresPagadoReal & ", cMovNro = '" & psMovNro & "' " & lsActCampo _
        & "WHERE  cPersCod='" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' AND cCtaIFCod = '" & psCtaIFCod & "' AND nNroCuota = " & pnNroCuota & " and cTpoCuota = '" & pnTpoCuota & "' "
    oConec.Ejecutar sql
Else
    sql = "UPDATE CtaIFCalendario SET cestado=2 , nInteresPagado = nInteresPagado + " & pnInteresPagado & ",nInteresPagado_Real = nInteresPagado_Real + " & pnInteresPagadoReal & ", cMovNro = '" & psMovNro & "' " & lsActCampo _
        & "WHERE  cPersCod='" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' AND cCtaIFCod = '" & psCtaIFCod & "' AND nNroCuota = " & pnNroCuota & " and cTpoCuota = '" & pnTpoCuota & "' "
    oConec.Ejecutar sql
End If

If pbCancela And pnEstado > -1 Then
    sql = "UPDATE CtaIFCalendario SET cMovNro = '" & psMovNro & "' " & lsActCampo _
        & "WHERE cPersCod='" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' AND cCtaIFCod = '" & psCtaIFCod & "' AND nNroCuota > " & pnNroCuota & "" ' and cTpoCuota = '" & pnTpoCuota & "' "
    oConec.Ejecutar sql
Else
    sql = "UPDATE CtaIFCalendario SET cMovNro = '" & psMovNro & "' " & lsActCampo _
        & "WHERE cPersCod='" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' AND cCtaIFCod = '" & psCtaIFCod & "' AND nNroCuota > " & pnNroCuota & " and cTpoCuota = '2' "
    oConec.Ejecutar sql
    
    If pnNoCuota6 <> "" Then
        sql = "UPDATE CtaIFCalendario SET cMovNro = '" & psMovNro & "' " & lsActCampo _
            & "WHERE cPersCod='" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' AND cCtaIFCod = '" & psCtaIFCod & "' AND nNroCuota > " & pnNroCuota & " and cTpoCuota = '6' "
        oConec.Ejecutar sql
    End If
End If

If bPago = True Then
    If Len(Trim(psMovNro)) > 0 Then
        sql = "SELECT cMovNro From CtaIFCalendario "
        sql = sql & "WHERE  cPersCod='" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' AND cCtaIFCod = '" & psCtaIFCod & "' AND nNroCuota = " & pnNroCuota & "" ' and cTpoCuota = '" & pnTpoCuota & "' "
        sql = sql & " AND len(cMovNro)>0"
        Set rs = oConec.CargaRecordSet(sql)
        If rs.BOF Then
        Else
            If rs!cMovNro <> psMovNro Then
                sql = "INSERT INTO CtaIFAdeudadosRel (cMovNroProv, cMovNroPago)"
                sql = sql & " VALUES('" & rs!cMovNro & "', '" & psMovNro & "')"
                oConec.Ejecutar sql
            End If
        End If
        rs.Close
    End If
End If

Exit Sub
ErrorUpdateCal:
    Err.Raise Err.Number, "dAdeudCal: UpdateCalendario", Err.Description
End Sub


Public Function GetCtaIFAdeudadoMonedaPago(psPersCod As String, psIFTpo As CGTipoIF, psCtaIFCod As String) As String
Dim rs As ADODB.Recordset
On Error GoTo ErrorVerCtaIFAdeud

GetCtaIFAdeudadoMonedaPago = ""
sql = "SELECT cMonedaPago FROM CTAIFADEUDADOS " _
    & "WHERE    cPersCod = '" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' and  cCtaIfCod='" & psCtaIFCod & "'"
Set rs = oConec.CargaRecordSet(sql)
If Not rs.EOF Then
   GetCtaIFAdeudadoMonedaPago = rs!cMonedaPago
End If
RSClose rs
Exit Function
ErrorVerCtaIFAdeud:
    Err.Raise Err.Number, "dAdeudCal", Err.Description
End Function

Public Function VerificaCtaIFAdeudadoCalendario(psPersCod As String, psIFTpo As CGTipoIF, psCtaIFCod As String) As Boolean
Dim rs As ADODB.Recordset
On Error GoTo ErrorVerificaCtaIFAdeudadoCalendario

VerificaCtaIFAdeudadoCalendario = False

sql = "SELECT * FROM CTAIFCalendario " _
    & "WHERE    cPersCod = '" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' and  cCtaIfCod='" & psCtaIFCod & "' And bVigente = 1"

Set rs = oConec.CargaRecordSet(sql)
VerificaCtaIFAdeudadoCalendario = Not rs.EOF
RSClose rs
Exit Function
ErrorVerificaCtaIFAdeudadoCalendario:
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


Public Function ObtenerAlertaVencimientoPago(ByVal pdFecha As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_obtenerAlertaVencimientoPago '" & Format(pdFecha, "YYYY/MM/DD") & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerAlertaVencimientoPago = rs
End Function
Public Function obtenerSitieneElCargo(ByVal psCargo As String) As Boolean
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    obtenerSitieneElCargo = False
    sSql = "exec stp_sel_obtenerSitieneElCargo '" & psCargo & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not (rs.BOF Or rs.EOF) Then
        obtenerSitieneElCargo = True
    End If
    oCon.CierraConexion
End Function
