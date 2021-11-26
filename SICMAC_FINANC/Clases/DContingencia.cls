VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DContingencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim sSQL As String
Dim lsMsgErr As String
Dim rs As ADODB.Recordset



'Comentado TORE28032018
'Public Sub RegistraNuevaContigencia(ByVal psNumRegistro As String, ByVal pnOrigen As Integer, ByVal pnMoneda As Integer, ByVal pnMonto As Double, ByVal psDesc As String, ByVal psMovNro As String, Optional ByVal pnEvPerdida As Integer = 0)
'    On Error GoTo ErrorRegistraNuevaContigencia
'
'    Set dbConec = New DConecta
'    dbConec.AbreConexion
'    sSql = "stp_ins_RegistraNuevaContigencia '" & psNumRegistro & "'," & pnOrigen & "," & pnEvPerdida & "," & pnMoneda & "," & pnMonto & ",'" & psDesc & "','" & psMovNro & "'"
'    dbConec.ConexionActiva.Execute sSql
'    dbConec.CierraConexion
'    Set dbConec = Nothing
'    Exit Sub
'
'ErrorRegistraNuevaContigencia:
'    Err.Raise Err.Number, "Nueva Contigencia", Err.Description
'End Sub

'CROB20170726 -> TORE28032018
Public Function RegistraNuevaContigencia(ByVal psNumRegistro As String, ByVal pnTipoConting As Integer, ByVal pnOrigen As Integer, ByVal pnEvPerdida As Integer _
                                     , ByVal pnReportado As Integer, ByVal psContigDesc As String, ByVal pMontoBruto As Double, ByVal pnMonedaRecup As Integer _
                                     , ByVal pnMontoRecup As Double, ByVal pnMonedaProv As Integer, ByVal pnMontoProvision As Double, ByVal pnPerdidaNeta As Double _
                                     , ByVal psCtaContCod As String, ByVal pdFechaOcurrencia As String, ByVal pdFechaDescubrimiento As String, ByVal dFechaRegCont As String _
                                     , ByVal psCodUser As String, ByVal psCodAge As String, ByVal pnMonedaPerdidaSG As Integer, ByVal pnMontoPerdidaSG As Double)
    On Error GoTo ErrorRegistraNuevaContigencia
    
    Set dbConec = New DConecta
    dbConec.AbreConexion
    sSQL = "stp_ins_RegistraNuevaContigencia '" & psNumRegistro & "'," & pnTipoConting & "," & pnOrigen & "," & pnEvPerdida & "," _
                                                & pnReportado & ",'" & psContigDesc & "'," & pMontoBruto & "," & pnMonedaRecup & "," _
                                                & pnMontoRecup & "," & pnMonedaProv & "," & pnMontoProvision & "," & pnPerdidaNeta & ",'" _
                                                & psCtaContCod & "','" & IIf(Left(Format(pdFechaOcurrencia, "yyyymmdd"), 1) = " ", "", Format(pdFechaOcurrencia, "yyyymmdd")) & "','" _
                                                & IIf(Left(Format(pdFechaDescubrimiento, "yyyymmdd"), 1) = " ", "", Format(pdFechaDescubrimiento, "yyyymmdd")) & "','" _
                                                & IIf(Left(Format(dFechaRegCont, "yyyymmdd"), 1) = " ", "", Format(dFechaRegCont, "yyyymmdd")) & "','" _
                                                & psCodUser & "','" & psCodAge & "'," & pnMonedaPerdidaSG & "," & pnMontoPerdidaSG
    'dbConec.ConexionActiva.Execute sSQL
    Set RegistraNuevaContigencia = dbConec.CargaRecordSet(sSQL)
    dbConec.CierraConexion
    Set dbConec = Nothing
    Exit Function
ErrorRegistraNuevaContigencia:
    Err.Raise Err.Number, "Nueva Contigencia", Err.Description
End Function 'CROB20170726

Public Function BuscaContigenciasxArea(Optional ByVal psCodArea As String = "", Optional ByVal psNumRegistro As String = "", Optional ByVal psCadEstado As String = "") As ADODB.Recordset
On Error GoTo BuscaContigenciasxAreaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_BuscaContigenciasxArea '" & psCodArea & "','" & psNumRegistro & "','" & psCadEstado & "'"
      Set BuscaContigenciasxArea = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
BuscaContigenciasxAreaErr:
   Call RaiseError(MyUnhandledError, "DDocumento:BuscaContigenciasxArea Method")
End Function

Public Function BuscaContigenciaSeleccionada(ByVal psNumRegistro As String) As ADODB.Recordset
On Error GoTo BuscaContigenciaSeleccionadaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_BuscaContigenciaSeleccionada '" & psNumRegistro & "'"
      Set BuscaContigenciaSeleccionada = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
BuscaContigenciaSeleccionadaErr:
   Call RaiseError(MyUnhandledError, "DDocumento:BuscaContigenciaSeleccionada Method")
End Function


Public Sub ExtornaContingencia(ByVal psNumRegistro As String)
    On Error GoTo ErrorExtornaContingencia
    
    Set dbConec = New DConecta
    dbConec.AbreConexion
    sSQL = "stp_upd_ExtornaContingencia '" & psNumRegistro & "'"
    dbConec.ConexionActiva.Execute sSQL
    dbConec.CierraConexion
    Set dbConec = Nothing
    Exit Sub

ErrorExtornaContingencia:
    Err.Raise Err.Number, "Extorno Contigencia", Err.Description
End Sub

Public Function BuscaInfsTecnicosxConting(Optional ByVal psNumRegistro As String = "") As ADODB.Recordset
On Error GoTo BuscaInfsTecnicosxContingErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_BuscaInformesTecxConting '" & psNumRegistro & "'"
      Set BuscaInfsTecnicosxConting = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
BuscaInfsTecnicosxContingErr:
   Call RaiseError(MyUnhandledError, "DDocumento:BuscaInfsTecnicosxConting Method")
End Function

Public Function BuscaContigenciasSeguimiento(Optional ByVal psCodArea As String = "", Optional ByVal psTipoConting As String = "", Optional ByVal psCadEstado As String = "") As ADODB.Recordset
On Error GoTo BuscaContigenciasSeguimientoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_BuscaContigenciasSeguimiento '" & psCodArea & "','" & psTipoConting & "','" & psCadEstado & "'"
      Set BuscaContigenciasSeguimiento = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
BuscaContigenciasSeguimientoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:BuscaContigenciasSeguimiento Method")
End Function

Public Function VerificaContingenciaSiPuedeLiberarse(ByVal psNumRegistro As String) As Boolean
On Error GoTo VerificaContingenciaSiPuedeLiberarseErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_VerificaContingenciaSiPuedeLiberarseoDesestimarse '" & psNumRegistro & "'"
      Set rs = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
      If Not rs.BOF Or Not rs.EOF Then
        If rs!nCalif = 3 Then 'Cierta o Probable
          VerificaContingenciaSiPuedeLiberarse = True
        Else
          VerificaContingenciaSiPuedeLiberarse = False
        End If
      End If
   End If
   Set dbConec = Nothing
   Exit Function
VerificaContingenciaSiPuedeLiberarseErr:
   Call RaiseError(MyUnhandledError, "DDocumento:VerificaContingenciaSiPuedeLiberarse Method")
End Function

Public Function VerificaContingenciaSiPuedeDesestimarse(ByVal psNumRegistro As String) As Boolean
On Error GoTo VerificaContingenciaSiPuedeDesestimarseErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_VerificaContingenciaSiPuedeLiberarseoDesestimarse '" & psNumRegistro & "'"
      Set rs = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
      If Not rs.BOF Or Not rs.EOF Then
        If Left(psNumRegistro, 1) = 1 Then
            If rs!nCalif <> 3 Then 'Diferente de CIERTO
              VerificaContingenciaSiPuedeDesestimarse = True
            Else
              VerificaContingenciaSiPuedeDesestimarse = False
            End If
        Else
            If rs!nCalif = 1 Then 'REMOTO
              VerificaContingenciaSiPuedeDesestimarse = True
            Else
              VerificaContingenciaSiPuedeDesestimarse = False
            End If
        End If
      End If
   End If
   Set dbConec = Nothing
   Exit Function
VerificaContingenciaSiPuedeDesestimarseErr:
   Call RaiseError(MyUnhandledError, "DDocumento:VerificaContingenciaSiPuedeDesestimarse Method")
End Function

Public Function BuscaContingenciaParaLiberar(ByVal psNumRegistro As String) As ADODB.Recordset
On Error GoTo BuscaContingenciaParaLiberarErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_BuscaContingenciaParaLiberar '" & psNumRegistro & "'"
      Set BuscaContingenciaParaLiberar = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
BuscaContingenciaParaLiberarErr:
   Call RaiseError(MyUnhandledError, "DDocumento:BuscaContingenciaParaLiberar Method")
End Function

Public Function BuscaContingenciaParaDesestimar(ByVal psNumRegistro As String) As ADODB.Recordset
On Error GoTo BuscaContingenciaParaDesestimarErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_BuscaContingenciaParaDesestimar '" & psNumRegistro & "'"
      Set BuscaContingenciaParaDesestimar = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
BuscaContingenciaParaDesestimarErr:
   Call RaiseError(MyUnhandledError, "DDocumento:BuscaContingenciaParaDesestimar Method")
End Function

Public Function ObtieneAgenciaUserRegistraConting(ByVal psCodUser As String) As String
On Error GoTo ObtieneAgenciaUserRegistraContingErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_ObtieneAgenciaUserRegistraConting '" & psCodUser & "'"
      Set rs = dbConec.CargaRecordSet(sSQL)
      ObtieneAgenciaUserRegistraConting = rs!cAgenciaActual
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
ObtieneAgenciaUserRegistraContingErr:
   Call RaiseError(MyUnhandledError, "DDocumento:ObtieneAgenciaUserRegistraConting Method")
End Function

Public Function ObtenerUserxCodPersEncargado(ByVal psCodPers As String) As String
On Error GoTo ObtenerUserxCodPersEncargadoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "Select cUser From RRHH where cPersCod='" & psCodPers & "'"
      Set rs = dbConec.CargaRecordSet(sSQL)
      ObtenerUserxCodPersEncargado = rs!cUser
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
ObtenerUserxCodPersEncargadoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:ObtenerUserxCodPersEncargado Method")
End Function

Public Sub LiberarContingencia(ByVal pnItem As Integer, ByVal psNumRegistro As String, ByVal psSubCtaCont As String, ByVal psSubCtaContCta As String, _
                            ByVal pnCalif As Integer, ByVal pnTipo As Integer, ByVal pnMoneda As Integer, ByVal pnProvision As Currency, _
                            ByVal pnDemanda As Currency, ByVal psOpeCod As String, ByVal psMovNro As String, ByRef psMsj As String)

On Error GoTo ErrorLiberarContingencia
    
Dim oDMov As DMov
Dim lnMovNro As String
Dim lbTrans As Boolean
'Dim psMsj As String
Dim ErrMovCta As String
Dim DCtaCont As String
Dim HCtaCont As String
Dim DCtaContDem As String
Dim HCtaContDem As String
Set oDMov = New DMov

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, IIf(Left(psNumRegistro, 1) = "1", "Activo", "Pasivo") & " Contingente Calificacion" & IIf(Left(psNumRegistro, 1) = "1", " Cierta", " Posible"), gMovEstContabMovContable
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnProvision, "0", "0"

If Left(psNumRegistro, 1) = "1" Then
    If pnCalif = 3 Then 'Cierta
        Set rs = BuscaCtaContxInformeTecnico(pnItem, psNumRegistro)
        DCtaCont = "11130" & psSubCtaCont & psSubCtaContCta
        HCtaCont = rs!cCtaContDeudorA '1509070902 & "01"
    End If
ElseIf Left(psNumRegistro, 1) = "2" Then
    If pnCalif = 3 Then 'Probable
        Set rs = BuscaCtaContxInformeTecnico(pnItem, psNumRegistro)
        DCtaCont = rs!cCtaContProvP '2702020101
        HCtaCont = "11130" & psSubCtaCont & psSubCtaContCta
        DCtaContDem = rs!cCtaContAcreedorP '720902090102
        HCtaContDem = rs!cCtaContDeudorP '7109
    End If
End If

oDMov.InsertaMovCta lnMovNro, 1, DCtaCont, pnProvision, , ErrMovCta
If ErrMovCta <> "" Then
    psMsj = ErrMovCta
    oDMov.RollbackTrans
    Exit Sub
End If
oDMov.InsertaMovCta lnMovNro, 2, HCtaCont, pnProvision * -1, , ErrMovCta
If ErrMovCta <> "" Then
    psMsj = ErrMovCta
    oDMov.RollbackTrans
    Exit Sub
End If

If DCtaContDem <> "" And HCtaContDem <> "" Then
    oDMov.InsertaMovCta lnMovNro, 3, DCtaContDem, pnDemanda
    If ErrMovCta <> "" Then
        psMsj = ErrMovCta
        oDMov.RollbackTrans
        Exit Sub
    End If
    oDMov.InsertaMovCta lnMovNro, 4, HCtaContDem, pnDemanda * -1
    If ErrMovCta <> "" Then
        psMsj = ErrMovCta
        oDMov.RollbackTrans
        Exit Sub
    End If
End If

If pnMoneda = gMonedaExtranjera Then
    Dim nTCFijo As Currency
    Dim dFecha As Date
    dFecha = CDate(GetFechaMov(psMovNro, True))
        
    Dim oTC As New nTipoCambio
    nTCFijo = oTC.EmiteTipoCambio(dFecha, TCFijoDia)

    oDMov.GeneraMovME lnMovNro, psMovNro, nTCFijo
End If
    
Set dbConec = New DConecta
dbConec.AbreConexion
sSQL = "stp_upd_LiberarContingencia '" & psNumRegistro & "'," & lnMovNro
dbConec.ConexionActiva.Execute sSQL
dbConec.CierraConexion
Set dbConec = Nothing

oDMov.CommitTrans
Exit Sub

ErrorLiberarContingencia:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, "Liberar Contigencia", Err.Description
End Sub

Public Function BuscaContigenciasLiberadas(ByVal pnTipoConting As Integer, ByVal pdFechaIni As String, ByVal pdFechaFin As String, Optional psAreaCod As String) As ADODB.Recordset
On Error GoTo BuscaContigenciasLiberadasErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_BuscaContigenciasLiberadas " & pnTipoConting & ",'" & Format(pdFechaIni, "yyyymmdd") & "','" & Format(pdFechaFin, "yyyymmdd") & "','" & psAreaCod & "'"
      Set BuscaContigenciasLiberadas = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
BuscaContigenciasLiberadasErr:
   Call RaiseError(MyUnhandledError, "DDocumento:BuscaContigenciasLiberadas Method")
End Function

Public Sub ExtornaLiberacionContingencia(ByVal psNumRegistro As String)
    On Error GoTo ErrorExtornaLiberacionContingencia
    
    Set dbConec = New DConecta
    dbConec.AbreConexion
    sSQL = "stp_upd_ExtornaLiberacionContingencia '" & psNumRegistro & "'"
    dbConec.ConexionActiva.Execute sSQL
    dbConec.CierraConexion
    Set dbConec = Nothing
    Exit Sub

ErrorExtornaLiberacionContingencia:
    Err.Raise Err.Number, "Extorno Liberacion de Contigencia", Err.Description
End Sub

Public Sub DesestimaContingencia(ByVal psNumRegistro As String, ByVal psGlosaDesest As String, ByVal psMovNro As String)
    On Error GoTo ErrorDesestimaContingencia
    
    Set dbConec = New DConecta
    dbConec.AbreConexion
    sSQL = "stp_upd_DesestimaContingencia '" & psNumRegistro & "','" & psGlosaDesest & "','" & psMovNro & "'"
    dbConec.ConexionActiva.Execute sSQL
    dbConec.CierraConexion
    Set dbConec = Nothing
    Exit Sub

ErrorDesestimaContingencia:
    Err.Raise Err.Number, "Desestimar Contigencia", Err.Description
End Sub

Public Function ObtieneCantInformesTecxConting(ByVal psNumRegistro As String) As Integer
On Error GoTo ObtieneCantInformesTecxContingErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_ObtieneCantInformesTecxConting'" & psNumRegistro & "'"
      Set rs = dbConec.CargaRecordSet(sSQL)
      ObtieneCantInformesTecxConting = rs!CantIT
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
ObtieneCantInformesTecxContingErr:
   Call RaiseError(MyUnhandledError, "DDocumento:ObtieneCantInformesTecxConting Method")
End Function

Public Function ObtieneUltimoInformeTecConting(ByVal psNumRegistro As String) As ADODB.Recordset
On Error GoTo ObtieneUltimoInformeTecContingErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_ObtieneUltimoInformeTecConting'" & psNumRegistro & "'"
      Set ObtieneUltimoInformeTecConting = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
ObtieneUltimoInformeTecContingErr:
   Call RaiseError(MyUnhandledError, "DDocumento:ObtieneUltimoInformeTecConting Method")
End Function

Public Function ObtieneValorConstante(ByVal pnConsValor As Integer, ByVal pnConsCod As Integer) As ADODB.Recordset
On Error GoTo ObtieneValorConstanteErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_ObtieneValorConstante " & pnConsValor & "," & pnConsCod
      Set ObtieneValorConstante = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
ObtieneValorConstanteErr:
   Call RaiseError(MyUnhandledError, "DDocumento:ObtieneValorConstante Method")
End Function

Public Sub RegistraInformeTecActivo(ByVal psNumRegistro As String, ByVal psNroInforme As String, ByVal pdNroInforme As String, ByVal psUserEncargado As String, _
                                    ByVal pnCalif As Integer, ByVal pnMoneda As Integer, ByVal pnProvision As Currency, ByVal psEntidad As String, _
                                    ByVal DocOrigen As String, ByVal FecOrigenA As String, ByVal psNombreRef As String, _
                                    ByVal cCtaContDeudor As String, ByVal cCtaContAcreedor As String, ByVal cCtaContAcreedoraAnt As String, _
                                    ByVal psOpeCod As String, ByVal psMovNro As String, ByRef sMensaje As String, ByRef nGeneraAsiento As Boolean, Optional VistaPrevia As Boolean = False, _
                                    Optional ByRef DCtaCont As String = "", Optional ByRef HCtaCont As String = "", Optional ByRef nMovImporte As Currency = 0, Optional ByRef DCtaContDem As String = "", _
                                    Optional ByRef HCtaContDem As String = "", Optional ByRef nMovImporteDem As Currency = 0)
On Error GoTo ErrorRegistraInformeTecActivo
    
Dim oDMov As DMov
Dim lnMovNro As String
Dim lbTrans As Boolean
'Dim psMsj As String
Dim ErrMovCta As String
'Dim cAgeUserReg As String
'Dim nMovImporte As Currency
'Dim nMovImporteDem As Currency
Dim pdFecCierreCont As String
Dim dFecSis As String
Dim rsConting As ADODB.Recordset
'Dim DCtaCont As String
'Dim HCtaCont As String
'Dim DCtaContDem As String
'Dim HCtaContDem As String
Set oDMov = New DMov

DCtaCont = ""
HCtaCont = ""
DCtaContDem = ""
HCtaContDem = ""
nMovImporte = 0
nMovImporteDem = 0

Set rs = ObtieneValorConstante(pnCalif, 5091)
If VistaPrevia = False Then
    oDMov.BeginTrans
    lbTrans = True
    oDMov.InsertaMov psMovNro, psOpeCod, "Activo Contingente Calificacion " & rs!cConsDescripcion, gMovEstContabMovContable
    lnMovNro = oDMov.GetnMovNro(psMovNro)
End If
'oDMov.InsertaMovCont lnMovNro, pnProvision, "0", "0"

'Set rsConting = BuscaContigenciaSeleccionada(psNumRegistro)
'cAgeUserReg = ObtieneAgenciaUserRegistraConting(rsConting!cUserReg)

pdFecCierreCont = "31/12/" & CStr(Year(Date) - 1)
dFecSis = GetFechaMov(psMovNro, True)

If ObtieneCantInformesTecxConting(psNumRegistro) = 0 Then
    If pnCalif = 3 Then 'Cierta
        DCtaCont = cCtaContDeudor '"1509070902" & "01" 'cAgeUserReg
        HCtaCont = cCtaContAcreedor '"520109" & cAgeUserReg
        nMovImporte = pnProvision
    End If
Else
    Set rs = ObtieneUltimoInformeTecConting(psNumRegistro)
    If rs!nCalif = 3 And pnCalif <> 3 Then 'Anterior Cierto y Nuevo No es Cierta
        If Year(dFecSis) = Year(rs!dFechaRegInfTec) Then 'DateDiff("D", pdFecCierreCont, rs!dFechaRegInfTec) > 0 Then 'Ejercicio Actual
            DCtaCont = cCtaContAcreedor '"520109" & cAgeUserReg
            HCtaCont = cCtaContDeudor '"1509070902" & "01" 'cAgeUserReg
        Else
            DCtaCont = cCtaContAcreedoraAnt '"650109" & cAgeUserReg
            HCtaCont = cCtaContDeudor '"1509070902" & "01" 'cAgeUserReg
        End If
        nMovImporte = rs!nProvision
    ElseIf rs!nCalif = 3 And pnCalif = 3 Then 'Anterior Cierta y Nuevo Cierta
        If Year(dFecSis) = Year(rs!dFechaRegInfTec) Then 'DateDiff("D", pdFecCierreCont, rs!dFechaRegInfTec) > 0 Then 'Ejercicio Actual
            If pnProvision > rs!nProvision Then
                DCtaCont = cCtaContDeudor '"1509070902" & "01" 'cAgeUserReg
                HCtaCont = cCtaContAcreedor '"520109" & cAgeUserReg
                nMovImporte = pnProvision - rs!nProvision
            ElseIf pnProvision < rs!nProvision Then
                DCtaCont = cCtaContAcreedor '"520109" & cAgeUserReg
                HCtaCont = cCtaContDeudor '"1509070902" & "01" 'cAgeUserReg
                nMovImporte = rs!nProvision - pnProvision
            ElseIf pnProvision = rs!nProvision Then
                DCtaCont = ""
                HCtaCont = ""
            End If
        Else
            If pnProvision > rs!nProvision Then
                DCtaCont = cCtaContDeudor '"1509070902" & "01" 'cAgeUserReg
                HCtaCont = cCtaContAcreedor '"520109" & cAgeUserReg
                nMovImporte = pnProvision - rs!nProvision
            ElseIf pnProvision < rs!nProvision Then
                DCtaCont = cCtaContAcreedoraAnt '"650109" & cAgeUserReg
                HCtaCont = cCtaContDeudor '"1509070902" & "01" 'cAgeUserReg
                nMovImporte = rs!nProvision - pnProvision
            ElseIf pnProvision = rs!nProvision Then
                DCtaCont = ""
                HCtaCont = ""
            End If
        End If
    ElseIf rs!nCalif <> 3 And pnCalif = 3 Then 'Anterior No es Cierta y Nuevo Cierta
        DCtaCont = cCtaContDeudor '"1509070902" & "01" 'cAgeUserReg
        HCtaCont = cCtaContAcreedor '"520109" & cAgeUserReg
        nMovImporte = pnProvision
    End If
End If

If VistaPrevia = False Then
    oDMov.InsertaMovCont lnMovNro, nMovImporte, "0", "0"
    
    If DCtaCont <> "" And HCtaCont <> "" Then
        oDMov.InsertaMovCta lnMovNro, 1, DCtaCont, nMovImporte, , ErrMovCta
        If ErrMovCta <> "" Then
            sMensaje = ErrMovCta
            oDMov.RollbackTrans
            Exit Sub
        End If
        oDMov.InsertaMovCta lnMovNro, 2, HCtaCont, nMovImporte * -1, , ErrMovCta
        If ErrMovCta <> "" Then
            sMensaje = ErrMovCta
            oDMov.RollbackTrans
            Exit Sub
        End If
        
        If pnMoneda = gMonedaExtranjera Then
            Dim nTCFijo As Currency
            Dim dFecha As Date
            dFecha = CDate(GetFechaMov(psMovNro, True))
            
            Dim oTC As New nTipoCambio
            nTCFijo = oTC.EmiteTipoCambio(dFecha, TCFijoDia)
        
            oDMov.GeneraMovME lnMovNro, psMovNro, nTCFijo
        End If
        nGeneraAsiento = True
    Else
        nGeneraAsiento = False
    End If
    
    Set dbConec = New DConecta
    dbConec.AbreConexion
    sSQL = "stp_upd_ins_RegistraInformeTecActivo '" & psNumRegistro & "','" & psNroInforme & "','" & Format(pdNroInforme, "yyyymmdd") & "','" & psUserEncargado & "'," & pnCalif & "," _
                                                    & pnMoneda & "," & pnProvision & ",'" & psEntidad & "','" & DocOrigen & "','" & Format(FecOrigenA, "yyyymmdd") & "','" & psNombreRef & "','" _
                                                    & cCtaContDeudor & "','" & cCtaContAcreedor & "','" & cCtaContAcreedoraAnt & "'," & lnMovNro
    dbConec.ConexionActiva.Execute sSQL
    dbConec.CierraConexion
    Set dbConec = Nothing
    
    oDMov.CommitTrans
Else
    If DCtaCont <> "" And HCtaCont <> "" Then
        nGeneraAsiento = True
    Else
        nGeneraAsiento = False
    End If
'    Dim oImpr As NContImprimir
'    Dim oPrevio As clsPrevioFinan
'    Set oImpr = New NContImprimir
'    Set oPrevio = New clsPrevioFinan
'
'    oPrevio.Show oImpr.VistaPreviaAsientoContable(psMovNro, 66, 79, , , , , , , psOpeCod, psNumRegistro, pnMoneda, DCtaCont, HCtaCont, nMovImporte, DCtaContDem, HCtaContDem, nMovImporteDem), gsOpeDesc, False, 66, gImpresora
End If
Exit Sub

ErrorRegistraInformeTecActivo:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, "Registro de IT para Contigencia Activa", Err.Description
End Sub

Public Sub RegistraInformeTecPasivo(ByVal psNumRegistro As String, ByVal psNroInforme As String, ByVal pdNroInforme As String, ByVal psUserEncargado As String, _
                                    ByVal pnCalif As Integer, ByVal pnTipo As Integer, ByVal pnMoneda As Integer, ByVal pnProvision As Currency, ByVal psEntidad As String, ByVal pnMonedaDem As Integer, _
                                    ByVal pnMontoDemanda As Currency, ByVal DocOrigen As String, ByVal FecOrigenA As String, ByVal cNombreRef As String, _
                                    ByVal cCtaContDeudor As String, ByVal cCtaContAcreedor As String, ByVal cCtaContGastos As String, ByVal cCtaContProv As String, ByVal cCtaContAcreedorAnt As String, _
                                    ByVal psOpeCod As String, ByVal psMovNro As String, ByRef sMensaje As String, ByRef nGeneraAsiento As Boolean, Optional VistaPrevia As Boolean = False, _
                                    Optional ByRef DCtaCont As String = "", Optional ByRef HCtaCont As String = "", Optional ByRef nMovImporte As Currency = 0, Optional ByRef DCtaContDem As String = "", _
                                    Optional ByRef HCtaContDem As String = "", Optional ByRef nMovImporteDem As Currency = 0)
                                    'JUEZ 20131028 Se agregó pnMonedaDem
                                    
On Error GoTo ErrorRegistraInformeTecPasivo
    
Dim oDMov As DMov
Dim lnMovNro As String
Dim lbTrans As Boolean
'Dim psMsj As String
Dim ErrMovCta As String
Dim cAgeUserReg As String
'Dim nMovImporte As Currency
'Dim nMovImporteDem As Currency
Dim pdFecCierreCont As String
Dim dFecSis As String
Dim rsConting As ADODB.Recordset
'Dim DCtaCont As String
'Dim HCtaCont As String
'Dim DCtaContDem As String
'Dim HCtaContDem As String
Set oDMov = New DMov

DCtaCont = ""
HCtaCont = ""
DCtaContDem = ""
HCtaContDem = ""
nMovImporte = 0
nMovImporteDem = 0

Set rs = ObtieneValorConstante(pnCalif, 5092)
If VistaPrevia = False Then
    oDMov.BeginTrans
    lbTrans = True
    oDMov.InsertaMov psMovNro, psOpeCod, "Pasivo Contingente Calificacion " & rs!cConsDescripcion, gMovEstContabMovContable
    lnMovNro = oDMov.GetnMovNro(psMovNro)
End If
'oDMov.InsertaMovCont lnMovNro, pnProvision, "0", "0"

Set rsConting = BuscaContigenciaSeleccionada(psNumRegistro)
cAgeUserReg = ObtieneAgenciaUserRegistraConting(rsConting!cUserReg)

pdFecCierreCont = "31/12/" & CStr(Year(Date) - 1)
dFecSis = GetFechaMov(psMovNro, True)

If ObtieneCantInformesTecxConting(psNumRegistro) = 0 Then
    If pnCalif = 1 Then 'Remoto
        DCtaCont = ""
        HCtaCont = ""
    ElseIf pnCalif = 2 Then 'Posible
        DCtaCont = cCtaContDeudor '"7109"
        HCtaCont = cCtaContAcreedor '"720902090102"
        nMovImporte = pnMontoDemanda 'rsConting!nMonto
    ElseIf pnCalif = 3 Then 'Probable
'        If pnTipo = 1 Then 'Laboral
        DCtaCont = cCtaContDeudor '"7109"
        HCtaCont = cCtaContAcreedor '"720902090102"
        DCtaContDem = cCtaContGastos '"4305090101" & cAgeUserReg
        HCtaContDem = cCtaContProv '"2702020101"
        nMovImporte = pnMontoDemanda
        nMovImporteDem = pnProvision 'rsConting!nMonto
'        Else 'Otros
'            DCtaCont = "7109"
'            HCtaCont = "720902090102"
'            DCtaContDem = "4305090109" & cAgeUserReg
'            HCtaContDem = "27020909"
'            nMovImporte = pnMontoDemanda
'            nMovImporteDem = pnProvision 'rsConting!nMonto
'        End If
    End If
Else
    Set rs = ObtieneUltimoInformeTecConting(psNumRegistro)
    If pnCalif = 1 Then 'Nuevo Remoto
        If rs!nCalif = 2 Then 'Anterior Posible
            DCtaCont = cCtaContAcreedor '"720902090102"
            HCtaCont = cCtaContDeudor '"7109"
            nMovImporte = rs!nDemandaLab 'Por ver
        ElseIf rs!nCalif = 3 Then 'Anterior Probable
            'If rs!nTipo = 1 Then 'Anterior Laboral
            If Year(dFecSis) = Year(rs!dFechaRegInfTec) Then 'DateDiff("D", pdFecCierreCont, rs!dFechaRegInfTec) > 0 Then 'Ejercicio Actual
                DCtaCont = cCtaContAcreedor '"720902090102"
                HCtaCont = cCtaContDeudor '"7109"
                DCtaContDem = cCtaContProv '"2702020101"
                HCtaContDem = cCtaContGastos '"64010909" & cAgeUserReg
                nMovImporte = rs!nDemandaLab
                nMovImporteDem = rs!nProvision
            Else
                DCtaCont = cCtaContAcreedor '"720902090102"
                HCtaCont = cCtaContDeudor '"7109"
                DCtaContDem = cCtaContProv '"2702020101"
                HCtaContDem = cCtaContAcreedorAnt '"4305090101" & cAgeUserReg
                nMovImporte = rs!nDemandaLab
                nMovImporteDem = rs!nProvision
            End If
'            ElseIf rs!nTipo = 2 Then 'Anterior Otros
'                If DateDiff("D", pdFecCierreCont, rs!dFechaRegInfTec) < 0 Then 'Ejercicio Anterior
'                    DCtaCont = "720902090102"
'                    HCtaCont = "7109"
'                    DCtaContDem = "27020909"
'                    HCtaContDem = "64010909" & cAgeUserReg
'                    nMovImporte = pnMontoDemanda
'                    nMovImporteDem = rs!nProvision
'                Else
'                    DCtaCont = "720902090102"
'                    HCtaCont = "7109"
'                    DCtaContDem = "27020909"
'                    HCtaContDem = "4305090109" & cAgeUserReg
'                    nMovImporte = pnMontoDemanda
'                    nMovImporteDem = rs!nProvision
'                End If
'            End If
        End If
    ElseIf pnCalif = 2 Then 'Nuevo Posible
        If rs!nCalif = 1 Then 'Anterior Remoto
            DCtaCont = cCtaContDeudor '"7109"
            HCtaCont = cCtaContAcreedor '"720902090102"
            nMovImporte = pnMontoDemanda 'rsConting!nMonto
        ElseIf rs!nCalif = 2 Then 'Anterior Posible
            If pnMontoDemanda > rs!nDemandaLab Then '> Anterior Demanda
                DCtaCont = cCtaContAcreedor '"7109"
                HCtaCont = cCtaContDeudor '"720902090102"
                nMovImporte = pnMontoDemanda - rs!nDemandaLab
            ElseIf pnMontoDemanda < rs!nDemandaLab Then '< Anterior Demanda
                DCtaCont = cCtaContDeudor '"720902090102"
                HCtaCont = cCtaContAcreedor '"7109"
                nMovImporte = rs!nDemandaLab - pnMontoDemanda
            ElseIf pnMontoDemanda = rs!nDemandaLab Then '= Anterior Demanda
                DCtaCont = ""
                HCtaCont = ""
            End If
        ElseIf rs!nCalif = 3 Then 'Anterior Probable
            If Year(dFecSis) = Year(rs!dFechaRegInfTec) Then 'DateDiff("D", pdFecCierreCont, rs!dFechaRegInfTec) > 0 Then 'Ejercicio Actual
                DCtaCont = cCtaContProv '"2702020101"
                HCtaCont = cCtaContGastos '"4305090101" & cAgeUserReg
                nMovImporte = rs!nProvision
            Else
                DCtaCont = cCtaContProv '"2702020101"
                HCtaCont = cCtaContAcreedorAnt '"64010909"
                nMovImporte = rs!nProvision
            End If
        End If
'DCA: cCtaContDeudor
'ACA: cCtaContAcreedor
'DGA: cCtaContGastos
'APA: cCtaContProv
'AEO: cCtaContAcreedorAnt
    ElseIf pnCalif = 3 Then 'Nuevo Probable
        'If pnTipo = 1 Then 'Nuevo Laboral
        If rs!nCalif = 1 Then 'Anterior Remoto
            DCtaCont = cCtaContDeudor '"7109"
            HCtaCont = cCtaContAcreedor '"720902090102"
            DCtaContDem = cCtaContGastos '"4305090101" & cAgeUserReg
            HCtaContDem = cCtaContProv '"2702020101"
            nMovImporte = pnMontoDemanda
            nMovImporteDem = pnProvision 'rsConting!nMonto
        ElseIf rs!nCalif = 2 Then 'Anterior Posible
            DCtaCont = cCtaContGastos '"4305090101" & cAgeUserReg
            HCtaCont = cCtaContProv '"2702020101"
            nMovImporte = pnProvision 'rsConting!nMonto
        ElseIf rs!nCalif = 3 Then 'Anterior Probable
            'If rs!nTipo = 1 Then 'Anterior Laboral
            If pnProvision > rs!nProvision Then '> Anterior Provision
                DCtaCont = cCtaContGastos '"4305090101" & cAgeUserReg
                HCtaCont = cCtaContProv '"2702020101"
                nMovImporte = pnProvision - rs!nProvision
            ElseIf pnProvision < rs!nProvision Then '< Anterior Provision
                If Year(dFecSis) = Year(rs!dFechaRegInfTec) Then 'DateDiff("D", pdFecCierreCont, rs!dFechaRegInfTec) > 0 Then 'Ejercicio Actual
                    DCtaCont = cCtaContProv '"2702020101"
                    HCtaCont = cCtaContGastos '"4305090101" & cAgeUserReg
                    nMovImporte = rs!nProvision - pnProvision
                Else
                    DCtaCont = cCtaContProv '"2702020101"
                    HCtaCont = cCtaContAcreedorAnt '"64010909" & cAgeUserReg
                    nMovImporte = rs!nProvision - pnProvision
                End If
            ElseIf pnProvision = rs!nProvision Then '= Anterior Provision
                DCtaCont = ""
                HCtaCont = ""
            End If
            'End If
        End If
'        ElseIf pnTipo = 2 Then 'Nuevo Otros
'            If rs!nCalif = 1 Then 'Anterior Remoto
'                DCtaCont = "7109"
'                HCtaCont = "720902090102"
'                DCtaContDem = "4305090101" & cAgeUserReg
'                HCtaContDem = "27020909"
'                nMovImporte = pnMontoDemanda
'                nMovImporteDem = pnProvision 'rsConting!nMonto
'            ElseIf rs!nCalif = 2 Then 'Anterior Posible
'                DCtaCont = "4305090109" & cAgeUserReg
'                HCtaCont = "27020909"
'                nMovImporte = pnProvision 'rsConting!nMonto
'            ElseIf rs!nCalif = 3 Then 'Anterior Probable
'                If rs!nTipo = 2 Then 'Anterior Otros
'                    If pnProvision > rs!nProvision Then '> Anterior Provision
'                        DCtaCont = "4305090109" & cAgeUserReg
'                        HCtaCont = "27020909"
'                        nMovImporte = pnProvision - rs!nProvision
'                    ElseIf pnProvision < rs!nProvision Then '< Anterior Provision
'                        If DateDiff("D", pdFecCierreCont, rs!dFechaRegInfTec) > 0 Then 'Ejercicio Actual
'                            DCtaCont = "27020909"
'                            HCtaCont = "4305090109" & cAgeUserReg
'                            nMovImporte = rs!nProvision - pnProvision
'                        Else
'                            DCtaCont = "27020909"
'                            HCtaCont = "64010909" & cAgeUserReg
'                            nMovImporte = rs!nProvision - pnProvision
'                        End If
'                    ElseIf pnProvision = rs!nProvision Then '= Anterior Provision
'                        DCtaCont = ""
'                        HCtaCont = ""
'                    End If
'                End If
'            End If
'        End If
    End If
End If

If VistaPrevia = False Then
    oDMov.InsertaMovCont lnMovNro, nMovImporte, "0", "0"
    If DCtaCont <> "" And HCtaCont <> "" Then
        oDMov.InsertaMovCta lnMovNro, 1, DCtaCont, nMovImporte, , ErrMovCta
        If ErrMovCta <> "" Then
            sMensaje = ErrMovCta
            oDMov.RollbackTrans
            Exit Sub
        End If
        oDMov.InsertaMovCta lnMovNro, 2, HCtaCont, nMovImporte * -1, , ErrMovCta
        If ErrMovCta <> "" Then
            sMensaje = ErrMovCta
            oDMov.RollbackTrans
            Exit Sub
        End If
        
        If DCtaContDem <> "" And HCtaContDem <> "" Then
            oDMov.InsertaMovCta lnMovNro, 3, DCtaContDem, nMovImporteDem, , ErrMovCta
            If ErrMovCta <> "" Then
                sMensaje = ErrMovCta
                oDMov.RollbackTrans
                Exit Sub
            End If
            oDMov.InsertaMovCta lnMovNro, 4, HCtaContDem, nMovImporteDem * -1, , ErrMovCta
            If ErrMovCta <> "" Then
                sMensaje = ErrMovCta
                oDMov.RollbackTrans
                Exit Sub
            End If
        End If
        
        If pnMoneda = gMonedaExtranjera Then
            Dim nTCFijo As Currency
            Dim dFecha As Date
            dFecha = CDate(GetFechaMov(psMovNro, True))
            
            Dim oTC As New nTipoCambio
            nTCFijo = oTC.EmiteTipoCambio(dFecha, TCFijoDia)
        
            oDMov.GeneraMovME lnMovNro, psMovNro, nTCFijo
        End If
        nGeneraAsiento = True
    Else
        nGeneraAsiento = False
    End If
    
    Set dbConec = New DConecta
    dbConec.AbreConexion
    sSQL = "stp_upd_ins_RegistraInformeTecPasivo '" & psNumRegistro & "','" & psNroInforme & "','" & Format(pdNroInforme, "yyyymmdd") & "','" & psUserEncargado & "'," & pnCalif & "," & pnTipo & "," & _
                                                    pnMoneda & "," & pnProvision & ",'" & psEntidad & "','" & pnMonedaDem & "'," & pnMontoDemanda & ",'" & DocOrigen & "','" & Format(FecOrigenA, "yyyymmdd") & "','" & cNombreRef & "','" & _
                                                    cCtaContDeudor & "','" & cCtaContAcreedor & "','" & cCtaContGastos & "','" & cCtaContProv & "','" & cCtaContAcreedorAnt & "'," & lnMovNro
                                                    'JUEZ 20131028 Se agregó pnMonedaDem
    dbConec.ConexionActiva.Execute sSQL
    dbConec.CierraConexion
    Set dbConec = Nothing
    
    oDMov.CommitTrans
Else
    If DCtaCont <> "" And HCtaCont <> "" Then
        nGeneraAsiento = True
    Else
        nGeneraAsiento = False
    End If
'    Dim oImpr As NContImprimir
'    Dim oPrevio As clsPrevioFinan
'    Set oImpr = New NContImprimir
'    Set oPrevio = New clsPrevioFinan
'
'    oPrevio.Show oImpr.VistaPreviaAsientoContable(psMovNro, 66, 79, , , , , , , psOpeCod, psNumRegistro, pnMoneda, DCtaCont, HCtaCont, nMovImporte, DCtaContDem, HCtaContDem, nMovImporteDem), gsOpeDesc, False, 66, gImpresora
End If
Exit Sub

ErrorRegistraInformeTecPasivo:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, "Registro de IT para Contigencia Pasiva", Err.Description
End Sub

Public Function ObtieneNroRegistroxnMovNroIT(ByVal psMovNro As String) As String
On Error GoTo ObtieneNroRegistroxnMovNroITErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      ObtieneNroRegistroxnMovNroIT = ""
      sSQL = "exec stp_sel_ObtieneNroRegistroxnMovNroIT '" & psMovNro & "'"
      Set rs = dbConec.CargaRecordSet(sSQL)
      If Not rs.EOF Then 'Juez 20120921
        ObtieneNroRegistroxnMovNroIT = rs!cNumRegistro
      End If
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
ObtieneNroRegistroxnMovNroITErr:
   Call RaiseError(MyUnhandledError, "DDocumento:ObtieneNroRegistroxnMovNroIT Method")
End Function

Public Sub ExtornaITConting(ByVal nItem As Integer, ByVal psNumRegistro As String)
    On Error GoTo ErrorExtornaITConting
    
    Set dbConec = New DConecta
    dbConec.AbreConexion
    sSQL = "stp_upd_ExtornaITConting " & nItem & ",'" & psNumRegistro & "'"
    dbConec.ConexionActiva.Execute sSQL
    dbConec.CierraConexion
    Set dbConec = Nothing
    Exit Sub

ErrorExtornaITConting:
    Err.Raise Err.Number, "Extorno IT de Contigencia", Err.Description
End Sub

Public Function RecuperaInfTecnicoxItem(ByVal pnItem As Integer, ByVal psNumRegistro As String) As ADODB.Recordset
On Error GoTo RecuperaInfTecnicoxItemErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_RecuperaInfTecnicoxItem " & pnItem & ",'" & psNumRegistro & "'"
      Set RecuperaInfTecnicoxItem = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
RecuperaInfTecnicoxItemErr:
   Call RaiseError(MyUnhandledError, "DDocumento:RecuperaInfTecnicoxItem Method")
End Function

Public Sub ActualizaInformeTecnico(ByVal nItem As Integer, ByVal psNumRegistro As String, psNroInforme As String, _
                                   ByVal psFecInforme As String, ByVal psUser As String, ByVal psDocOrigen As String, _
                                   ByVal psFecOrigen As String, ByVal psEntidad As String, ByVal psNombreRef As String)
    On Error GoTo ErrorActualizaInformeTecnico
    
    Set dbConec = New DConecta
    dbConec.AbreConexion
    sSQL = "stp_upd_ActualizaInformeTecnico " & nItem & ",'" & psNumRegistro & "','" & psNroInforme & "','" & Format(psFecInforme, "yyyymmdd") & "','"
    sSQL = sSQL & psUser & "','" & psDocOrigen & "','" & Format(psFecOrigen, "yyyymmdd") & "','" & psEntidad & "','" & psNombreRef & "'"
    dbConec.ConexionActiva.Execute sSQL
    dbConec.CierraConexion
    Set dbConec = Nothing
    Exit Sub

ErrorActualizaInformeTecnico:
    Err.Raise Err.Number, "Actualizar Informe Tecnico", Err.Description
End Sub

Public Function BuscaCtaContxInformeTecnico(ByVal pnItem As Integer, ByVal psNumRegistro As String) As ADODB.Recordset
On Error GoTo BuscaCtaContxInformeTecnicoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_BuscaCtaContxInformeTecnico " & pnItem & ",'" & psNumRegistro & "'"
      Set BuscaCtaContxInformeTecnico = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
BuscaCtaContxInformeTecnicoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:BuscaCtaContxInformeTecnico Method")
End Function

Public Function VerificaSiExisteNroInformeTec(ByVal psNumRegistro As String, ByVal psNroInforme As String) As Boolean
On Error GoTo VerificaSiExisteNroInformeTecErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      VerificaSiExisteNroInformeTec = True
      sSQL = "exec stp_sel_VerificaSiExisteNroInformeTec '" & psNumRegistro & "','" & psNroInforme & "'"
      Set rs = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
      
      If rs.EOF Or rs.BOF Then 'si es vacio
        VerificaSiExisteNroInformeTec = False
      End If
   End If
   Set dbConec = Nothing
   Exit Function
VerificaSiExisteNroInformeTecErr:
   Call RaiseError(MyUnhandledError, "DDocumento:VerificaSiExisteNroInformeTec Method")
End Function

'TORE07032018
Public Function ObtenerNroRegistroContingencia(ByVal nTpoConting As Integer) As ADODB.Recordset
On Error GoTo ObtenerNroRegistroContingenciaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sSQL = "exec stp_sel_ObtenerNroRegistroContingencia '" & nTpoConting & "'"
      Set ObtenerNroRegistroContingencia = dbConec.CargaRecordSet(sSQL)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
ObtenerNroRegistroContingenciaErr:
   Call RaiseError(MyUnhandledError, "DDocumento:ObtenerNroRegistroContingencia Method")
End Function

Public Function ListarCtaContPasivosEventosPerdidaReportados()
    On Error GoTo ErrorListarCtaContPasivosEventosPerdidaReportados
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_sel_ListarCtaContPasivosEventosPerdidaReportados"
        Set ListarCtaContPasivosEventosPerdidaReportados = dbConec.CargaRecordSet(sSQL)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrorListarCtaContPasivosEventosPerdidaReportados:
    Call RaiseError(MyUnhandledError, "Error en proceso")
End Function

Public Function ListarCtaContPasivosEventosPerdidaNOReportados()
    On Error GoTo ErrorListarCtaContPasivosEventosPerdidaNOReportados
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_sel_ListarCtaContPasivosEventosPerdidaNOReportados"
        Set ListarCtaContPasivosEventosPerdidaNOReportados = dbConec.CargaRecordSet(sSQL)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrorListarCtaContPasivosEventosPerdidaNOReportados:
    Call RaiseError(MyUnhandledError, "Error en proceso")
End Function

Public Function ActualizarProvisionContingencia(ByVal psNumRegistro As String, ByVal pnProvision As Double, ByVal psUserRegistra As String)
On Error GoTo ErrorActualizarProvisionContingencia
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_upd_ActualizarProvisionContingencia '" & psNumRegistro & "'," & pnProvision & ",'" & psUserRegistra & "'"
        Set ActualizarProvisionContingencia = dbConec.CargaRecordSet(sSQL)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrorActualizarProvisionContingencia:
Call RaiseError(MyUnhandledError, "Error en proceso")
End Function

Public Function ActualizarCtaContableContingencia(ByVal pnCtaContable As String, ByVal psNumRegistro As String)
On Error GoTo ErrorActualizarCtaContableContingencia
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_upd_ActualizarCtaContableContingencia '" & pnCtaContable & "','" & psNumRegistro & "'"
        Set ActualizarCtaContableContingencia = dbConec.CargaRecordSet(sSQL)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrorActualizarCtaContableContingencia:
Call RaiseError(MyUnhandledError, "Error en proceso")
End Function

Public Function ActualizarMontoRecuperadoContingencia(ByVal psNumRegistro As String, ByVal pnRecuperado As Double, ByVal psUserReg As String)
On Error GoTo ErrorActualizarMontoRecuperadoContingencia
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_upd_ActualizarMontoRecuperadoContingencia '" & psNumRegistro & "'," & pnRecuperado & ",'" & psUserReg & "'"
        Set ActualizarMontoRecuperadoContingencia = dbConec.CargaRecordSet(sSQL)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrorActualizarMontoRecuperadoContingencia:
Call RaiseError(MyUnhandledError, "Error en proceso")
End Function

Public Function CargaContingenciasMontos(ByVal psNumRegistro As String) As Recordset
Dim sSQL As String
Dim oConecta As DConecta
On Error GoTo ErrorCargaContingenciasMontos
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSQL = "Select tmpc.nTipoMontoPCID as cTipoMontoPCID_A, tmpc.cTipoMontoPCDesc as cMonto, tmpc.cTipoMontoPCDesc + space(50) + convert(varchar(5), tmpc.nTipoMontoPCID, 103) AS cTipoMonto, cMoneda = (case when cm.nMoneda=1 then 'SOLES' else 'DOLARES' end), nnMoneda = (case when cm.nMoneda=1 then 'SOLES' + space(50) + convert(varchar(1), cm.nMoneda, 103 ) else 'DOLARES' + space(50) + convert(varchar(1), cm.nMoneda, 103 ) end), cm.nMonto, cm.nTipoCambio AS nTipoCamb, dFechaReg As dFecha "
    sSQL = sSQL & " From ContingenciaMontos cm inner join TipoMontoPasivoContingente tmpc ON tmpc.nTipoMontoPCID=cm.nTipoMontoPCID "
    sSQL = sSQL & " WHERE cm.cNumRegistro = '" & psNumRegistro & "'"
    Set CargaContingenciasMontos = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargaContingenciasMontos:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Instituciones Financieras", Err.Description
End Function

Public Function ListarContingenciaMontos(ByVal psNumRegistro As String)
    On Error GoTo ErrorListarContingenciaMontos
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_sel_ListarContingenciaMontos '" & psNumRegistro & "'"
        Set ListarContingenciaMontos = dbConec.CargaRecordSet(sSQL)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrorListarContingenciaMontos:
    Call RaiseError(MyUnhandledError, "Error en proceso")
End Function

'END TORE





