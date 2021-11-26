VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCtaSaldo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A82E2C1032C"
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sCentralCom As String

Public Function CargaCtaSaldo(Optional psCtaContCod As String = "", Optional psFecha As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
Dim sCond As String
   On Error GoTo CargaCtaSaldoErr
   sCond = ""
   If psCtaContCod <> "" Then
      sCond = " WHERE cCtaContCod = '" & psCtaContCod & "'"
   End If
   If psFecha <> "" Then
      sCond = sCond & IIf(sCond = "", "WHERE", "AND") & " dCtaSaldoFecha = (SELECT MAX(dCtaSaldoFecha) FROM CtaSaldo cs1 WHERE cs1.cCtaContCod = cs.cCtaContCod and dCtaSaldoFecha <= '" & psFecha & "')"
   End If
psSql = "SELECT cCtaContCod, dCtaSaldoFecha, Isnull(nCtaSaldoImporte,0)nCtaSaldoImporte,ISNULL(nCtaSaldoImporteME,0)nCtaSaldoImporteME FROM CtaSaldo cs " & sCond & " ORDER BY cCtaContCod "
   Set CargaCtaSaldo = dbConec.CargaRecordSet(psSql, LockType)
   Exit Function
CargaCtaSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:CargaCtaSaldo Method")
End Function

Public Sub InsertaCtaSaldo(psCtaContCod As String, psFecha As String, pnImporte As Currency, pnImporteME As Currency)
   On Error GoTo InsertaCtaSaldoErr
      psSql = "INSERT CtaSaldo (cCtaContCod, dCtaSaldoFecha, nCtaSaldoImporte, nCtaSaldoImporteME) " _
            & "VALUES ('" & psCtaContCod & "', '" & psFecha & "'," & pnImporte & "," & pnImporteME & ")"
      dbConec.Ejecutar psSql
   Exit Sub
InsertaCtaSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:InsertaCtaSaldo Method")
End Sub

Public Sub ActualizaCtaSaldo(psCtaContCod As String, psFecha As String, pnImporte As Currency, pnImporteME As Currency)
   On Error GoTo ActualizaCtaSaldoErr
      psSql = "UPDATE CtaSaldo SET nCtaSaldoImporte = " & pnImporte & ", nCtaSaldoImporteME = " & pnImporteME _
            & "WHERE cCtaContCod = '" & psCtaContCod & "' and dCtaSaldoFecha = '" & psFecha & "'"
      dbConec.Ejecutar psSql
   Exit Sub
ActualizaCtaSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:ActualizaCtaSaldo Method")
End Sub
Public Sub EliminaCtaSaldo(Optional psCtaContCod As String = "", Optional psFecha As String = "", Optional psSimbolo As String = "=")
Dim sCondCta   As String
Dim sCondFecha As String
   On Error GoTo EliminaCtaSaldoErr
      If psCtaContCod <> "" Then
         sCondCta = "WHERE cCtaContCod = '" & psCtaContCod & "' "
      End If
      If psFecha <> "" Then
         sCondFecha = IIf(sCondCta = "", "WHERE", "AND") & " dCtaSaldoFecha " & psSimbolo & " '" & psFecha & "'"
      End If
      psSql = "DELETE CtaSaldo " & sCondCta & sCondFecha
      dbConec.Ejecutar psSql
   Exit Sub
EliminaCtaSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:EliminaCtaSaldo Method")
End Sub

Public Function CargaCtaObjSaldo(Optional psCtaContCod As String = "", Optional psFecha As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
Dim sCond As String
   On Error GoTo CargaCtaObjSaldoErr
   sCond = ""
   If psCtaContCod <> "" Then
      sCond = " WHERE cCtaContCod = '" & psCtaContCod & "'"
   End If
   If psFecha <> "" Then
      sCond = sCond & IIf(sCond = "", "WHERE", "AND") & " dCtaSaldoFecha = (SELECT MAX(dCtaObjSaldoFecha) FROM CtaObjSaldo cs1 WHERE cs1.cCtaContCod = cs.cCtaContCod and cs1.cObjetoCod = cs.cObjetoCod and dCtaObjSaldoFecha <= '" & psFecha & "')"
   End If
   psSql = "SELECT cCtaContCod, cCtaObjOrden, cObjetoCod, dCtaObjSaldoFecha, nCtaObjSaldoImporte, nCtaObjSaldoImporteME FROM CtaObjSaldo cs " & sCond
   Set CargaCtaObjSaldo = dbConec.CargaRecordSet(psSql, LockType)
   Exit Function
CargaCtaObjSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:CargaCtaObjSaldo Method")
End Function

Public Sub InsertaCtaSaldoEstad(psCtaSaldoFecha, psCodUser As String, psCtaSaldoFecIni As String, psCtaSaldoFecFin As String, pnValido As Integer)
   On Error GoTo InsertaCtaSaldoEstadErr
   psSql = "INSERT CtaSaldoEstad (dCtaSaldoFecha, cCodUser, dCtaSaldoFechaIni, dCtaSaldoFechaFin, bCtaSaldoValido) " _
         & "VALUES ('" & psCtaSaldoFecha & "','" & psCodUser & "', '" & psCtaSaldoFecIni & "','" & psCtaSaldoFecFin & "', '" & pnValido & "')"
   dbConec.Ejecutar psSql
   Exit Sub
InsertaCtaSaldoEstadErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:InsertaCtaSaldoEstad Method")
End Sub

Public Sub ActualizaCtaSaldoEstad(psCtaSaldoFecha, psCodUser As String, pnValido As Integer)
   On Error GoTo ActualizaCtaSaldoEstadErr
   psSql = "UPDATE CtaSaldoEstad set bCtaSaldoValido = " & pnValido & " WHERE dCtaSaldoFecha = '" & psCtaSaldoFecha & "' and cCodUser = '" & psCodUser & "'"
   dbConec.Ejecutar psSql
   Exit Sub
ActualizaCtaSaldoEstadErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:ActualizaCtaSaldoEstad Method")
End Sub

Public Sub InsertaCtaObjSaldo()
   On Error GoTo InsertaCtaObjSaldoErr
   
   Exit Sub
InsertaCtaObjSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:InsertaCtaObjSaldo Method")
End Sub

Public Sub ActualizaCtaObjSaldo()
   On Error GoTo ActualizaCtaObjSaldoErr
   
   Exit Sub
ActualizaCtaObjSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:ActualizaCtaObjSaldo Method")
End Sub

Public Sub EliminaCtaObjSaldo(Optional psCtaContCod As String = "", Optional psFecha As String = "", Optional psSimbolo As String = "=")
Dim sCondCta   As String
Dim sCondFecha As String
   On Error GoTo EliminaCtaObjSaldoErr
      If psCtaContCod <> "" Then
         sCondCta = "WHERE cCtaContCod = '" & psCtaContCod & "' "
      End If
      If psFecha <> "" Then
         sCondFecha = IIf(sCondCta = "", "WHERE", "AND") & " dCtaObjSaldoFecha " & psSimbolo & " '" & psFecha & "'"
      End If
      psSql = "DELETE CtaObjSaldo " & sCondCta & sCondFecha
      dbConec.Ejecutar psSql
   Exit Sub
EliminaCtaObjSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:EliminaCtaObjSaldo Method")
End Sub

Public Sub EliminaCtaObjEfectivoSaldo(Optional psCtaContCod As String = "", Optional psFecha As String = "", Optional psSimbolo As String = "=")
Dim sCondCta   As String
Dim sCondFecha As String
   On Error GoTo EliminaCtaObjEfectivoSaldoErr
      If psCtaContCod <> "" Then
         sCondCta = "WHERE cCtaContCod = '" & psCtaContCod & "' "
      End If
      If psFecha <> "" Then
         sCondFecha = IIf(sCondCta = "", "WHERE", "AND") & " dEfectivoSaldoFecha " & psSimbolo & " '" & psFecha & "'"
      End If
      psSql = "DELETE CtaObjEfectivoSaldo " & sCondCta & sCondFecha
      dbConec.Ejecutar psSql
   Exit Sub
EliminaCtaObjEfectivoSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:EliminaCtaObjEfectivoSaldo Method")
End Sub

Public Sub EliminaCtaIFSaldo(Optional psCtaContCod As String = "", Optional psFecha As String = "", Optional psSimbolo As String = "=")
Dim sCondCta   As String
Dim sCondFecha As String
   On Error GoTo EliminaCtaIFSaldoErr
      If psCtaContCod <> "" Then
         sCondCta = "WHERE cCtaContCod = '" & psCtaContCod & "' "
      End If
      If psFecha <> "" Then
         sCondFecha = IIf(sCondCta = "", "WHERE", "AND") & " dCtaIFSaldo " & psSimbolo & " '" & psFecha & "'"
      End If
      psSql = "DELETE CtaIFSaldo " & sCondCta & sCondFecha
      dbConec.Ejecutar psSql
   Exit Sub
EliminaCtaIFSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:EliminaCtaIFSaldo Method")
End Sub

Public Sub EliminaCtaObjAreaAgenciaSaldo(Optional psCtaContCod As String = "", Optional psFecha As String = "", Optional psSimbolo As String = "=")
Dim sCondCta   As String
Dim sCondFecha As String
   On Error GoTo EliminaCtaObjAreaAgenciaSaldoErr
      If psCtaContCod <> "" Then
         sCondCta = "WHERE cCtaContCod = '" & psCtaContCod & "' "
      End If
      If psFecha <> "" Then
         sCondFecha = IIf(sCondCta = "", "WHERE", "AND") & " dAreaSaldoFecha " & psSimbolo & " '" & psFecha & "'"
      End If
      psSql = "DELETE CtaObjAreaAgenciaSaldo " & sCondCta & sCondFecha
      dbConec.Ejecutar psSql
   Exit Sub
EliminaCtaObjAreaAgenciaSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:EliminaCtaObjAreaAgenciaSaldo Method")
End Sub


Public Function GeneraSaldosContables(pdFecha As Date) As Boolean

Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

   On Error GoTo GeneraSaldosContablesErr
      
      EliminaCtaSaldo "", Format(pdFecha, gsFormatoFecha)
      EliminaCtaObjSaldo "", Format(pdFecha, gsFormatoFecha)
      EliminaCtaObjAreaAgenciaSaldo "", Format(pdFecha, gsFormatoFecha)
      EliminaCtaObjEfectivoSaldo "", Format(pdFecha, gsFormatoFecha)
      EliminaCtaIFSaldo "", Format(pdFecha, gsFormatoFecha)

        Set cmd = New ADODB.Command
        Set prm = New ADODB.Parameter
        cmd.CommandText = "spSaldosContab"
        cmd.CommandType = adCmdStoredProc
        cmd.Name = "spSaldosContab"
        Set prm = cmd.CreateParameter("pdFecha", adDate, adParamInput)
        cmd.Parameters.Append prm

        Set cmd.ActiveConnection = dbConec.ConexionActiva
        cmd.CommandTimeOut = 720
        cmd.Parameters.Refresh

        dbConec.ConexionActiva.spSaldosContab pdFecha
      


Exit Function
GeneraSaldosContablesErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:GeneraSaldosContables Method")
End Function
'Public Function GeneraSaldosContablesNew(pdFecha As Date) As Boolean
Public Function GeneraSaldosContablesNew(pdFecha As String) As Boolean
'Dim cmd As ADODB.Command
'Dim prm As ADODB.Parameter

   On Error GoTo GeneraSaldosContablesErr
      
'        Set cmd = New ADODB.Command
'        Set prm = New ADODB.Parameter
'        cmd.CommandText = "sp_GeneraSaldoContab"
'        cmd.CommandType = adCmdStoredProc
'        cmd.Name = "sp_GeneraSaldoContab"
'        Set prm = cmd.CreateParameter("pdFecha", adDate, adParamInput)
'        cmd.Parameters.Append prm
'
'        Set cmd.ActiveConnection = dbConec.ConexionActiva
'        cmd.CommandTimeOut = 720
'        cmd.Parameters.Refresh
'
'        dbConec.ConexionActiva.spSaldosContab pdFecha
Dim psSql As String
    psSql = "exec sp_GeneraSaldoContab '" & pdFecha & "'"
    dbConec.Ejecutar psSql
    
Exit Function
GeneraSaldosContablesErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:GeneraSaldosContablesNew Method")
End Function

Public Function PermiteActualizarSaldos(psCodUser As String) As Boolean
   On Error GoTo PermiteActualizarSaldosErr
   psSql = "SELECT cs.bCtaSaldoValido, cs.cCodUser, p.cPersNombre " _
         & "FROM CtaSaldoEstad cs JOIN RRHH RH ON RH.cUser = cs.cCodUser " _
         & "     JOIN Persona P ON P.cPersCod = RH.cPersCod " _
         & "WHERE  dCtaSaldoFecha = (SELECT MAX(dCtaSaldoFecha) FROM CtaSaldoEstad)"
   Set prs = dbConec.CargaRecordSet(psSql)
   If Not prs.EOF Then
      If Not prs!bCtaSaldoValido Then
         If prs!cCodUser <> psCodUser Then
            PermiteActualizarSaldos = False
            Err.Raise "50001", "DCtaSaldo", UCase(prs!cpersNombre) & " se encuentra Actualizando Saldos"
         Else
            PermiteActualizarSaldos = True
         End If
      Else
         PermiteActualizarSaldos = True
      End If
   Else
      PermiteActualizarSaldos = True
   End If
Exit Function
PermiteActualizarSaldosErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:PermiteActualizarSaldos Method")
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As New ClasIni
   sCentralCom = oIni.BaseComunes
Set oIni = Nothing
   Set dbConec = New DConecta
   dbConec.AbreConexion
End Sub
'ALPA 20100412*
Public Function ObtenerDatosInteresDevSuspen() As Recordset
Dim lsCtaFiltro As String
    On Error GoTo ObtenerDatosInteresDevSuspenErr
    psSql = "exec stp_sel_InteresesDevengadosyEnSuspenso "
    Set ObtenerDatosInteresDevSuspen = dbConec.CargaRecordSet(psSql)

Exit Function
ObtenerDatosInteresDevSuspenErr:
   Call RaiseError(MyUnhandledError, "NBalanceCont:LeeBalanceHisto Method")
End Function

Public Function ObtenerDatosInteresDevSuspenDetalle() As Recordset
Dim lsCtaFiltro As String
    On Error GoTo ObtenerDatosInteresDevSuspenDetalleErr
    psSql = "exec stp_sel_InteresesDevengadosyEnSuspensoDetalle "
    Set ObtenerDatosInteresDevSuspenDetalle = dbConec.CargaRecordSet(psSql)

Exit Function
ObtenerDatosInteresDevSuspenDetalleErr:
   Call RaiseError(MyUnhandledError, "NBalanceCont:LeeBalanceHisto Method")
End Function
'**************
Private Sub Class_Terminate()
dbConec.CierraConexion
Set dbConec = Nothing
End Sub

'ALPA 20110704
Public Function ObtenerDatosCtaSaldoMaxFecha(ByVal psCtaContCod As String, ByVal pdFecha As Date, ByVal pnTipoCambio As Currency, ByVal psPersCod As String, psIFTpo As String, psCtaIfCod As String, ByVal nNroCuota As Integer) As Currency
Dim lsCtaFiltro As String
Dim nMonto As Currency
Dim oRs As ADODB.Recordset
    On Error GoTo ObtenerDatosCtaSaldoMaxFechaErr
    Set oRs = New ADODB.Recordset
    psSql = "exec stp_sel_CtaIFObtieneSaldoFecha '" & psCtaContCod & "','" & Format(pdFecha, "YYYY/MM/DD") & "', " & pnTipoCambio & ",'" & psPersCod & "','" & psIFTpo & "', '" & psCtaIfCod & "'," & nNroCuota
    Set oRs = dbConec.CargaRecordSet(psSql)
    If oRs.EOF Or oRs.BOF Then
        nMonto = 0
    Else
        nMonto = oRs!nCtaSaldoImporte
    End If
    ObtenerDatosCtaSaldoMaxFecha = nMonto

Exit Function
ObtenerDatosCtaSaldoMaxFechaErr:
   Call RaiseError(MyUnhandledError, "NBalanceCont:LeeBalanceHisto Method")
End Function
'*******************************************
'ALPA 20100412*
Public Sub RegistrarAdeudadoMov(ByVal pnMovNro As Long, ByVal psPersCod As String, psIFTpo As String, psCtaIfCod As String, ByVal nNroCuota As Integer)
Dim lsCtaFiltro As String
    On Error GoTo RegistrarAdeudadoMovErr
    psSql = "exec stp_ins_AdeudadoMov " & pnMovNro & ",'" & psPersCod & "','" & psIFTpo & "', '" & psCtaIfCod & "'," & nNroCuota
    dbConec.CargaRecordSet (psSql)
Exit Sub
RegistrarAdeudadoMovErr:
   Call RaiseError(MyUnhandledError, "NBalanceCont:LeeBalanceHisto Method")
End Sub

'ALPA 20110909*
Public Function ObtenerOperacionesSAF(ByVal psMovNro As String) As ADODB.Recordset
Dim lsCtaFiltro As String
    On Error GoTo ObtenerOperacionesSAFErr
    psSql = "exec SAF_stp_sel_ObtenerOperacionesSAF '" & psMovNro & "'"
    Set ObtenerOperacionesSAF = dbConec.CargaRecordSet(psSql)
Exit Function
ObtenerOperacionesSAFErr:
   Call RaiseError(MyUnhandledError, "NBalanceCont:LeeBalanceHisto Method")
End Function
'EJVG20131122 ***
Public Function ObtenerOperacionesSAF_NEW(ByVal pnMovNro As Long) As ADODB.Recordset
Dim lsCtaFiltro As String
    On Error GoTo ObtenerOperacionesSAFErr
    psSql = "Exec SAF_stp_sel_ObtenerOperacionesSAF_NEW " & pnMovNro
    Set ObtenerOperacionesSAF_NEW = dbConec.CargaRecordSet(psSql)
Exit Function
ObtenerOperacionesSAFErr:
   Call RaiseError(MyUnhandledError, "NBalanceCont:ObtenerOperacionesSAF Method")
End Function
'END EJVG *******
Public Sub InsertarDesembolso_Log(ByVal psCtaSaf As String, ByVal psCtaCod As String, ByVal pnMovNro As Long)
Dim lsCtaFiltro As String
    On Error GoTo ObtenerOperacionesSAFErr
    psSql = "exec SAF_stp_ins_Desembolso_Log '" & psCtaSaf & "','" & psCtaCod & "'," & pnMovNro
    dbConec.CargaRecordSet (psSql)
Exit Sub
ObtenerOperacionesSAFErr:
   Call RaiseError(MyUnhandledError, "NBalanceCont:LeeBalanceHisto Method")
End Sub
'ALPA 20130506*************************************************************************************
Public Function GetSaldoFinalBalance(ByVal psCtaContCod As String, ByVal psAnio As String, ByVal psMes As String, ByVal psMoneda As String) As ADODB.Recordset
Dim lrs As ADODB.Recordset
Dim oConecta As DConecta
Set oConecta = New DConecta

    On Error GoTo GetSaldoFinalBalanceErr
    oConecta.AbreConexion
    psSql = "exec stp_sel_saldoFinCtaContBalanceEstad '" & psAnio & "','" & psMes & "','1','" & psMoneda & "','" & psCtaContCod & "'"
    Set GetSaldoFinalBalance = oConecta.CargaRecordSet(psSql)
    oConecta.CierraConexion
    
Exit Function
GetSaldoFinalBalanceErr:
   Call RaiseError(MyUnhandledError, "NBalanceCont:LeeBalanceHisto Method")
    
End Function

'*** PEAC 20130617
Public Function GetGeneraDatosRep25Nuevo(ByVal psFecha As String, ByRef rsDism As ADODB.Recordset, ByRef rsRecReexME As ADODB.Recordset)
Dim lsAnio As String
Dim lsMes As String

On Error GoTo GetGeneraDatosRep25NuevoErr

    lsAnio = Left(psFecha, 4)
    lsMes = Mid(psFecha, 5, 2)

    psSql = "exec stp_sel_Reporte25Aumento '" & psFecha & "'"
    Set rsDism = dbConec.CargaRecordSet(psSql)
      
    psSql = "exec stp_sel_Reporte25Disminucion '" & psFecha & "'"
    Set rsRecReexME = dbConec.CargaRecordSet(psSql)
    
    Exit Function
GetGeneraDatosRep25NuevoErr:
    Call RaiseError(MyUnhandledError, "DCtaSaldo:GetGeneraDatosRep25Nuevo Method")
End Function
