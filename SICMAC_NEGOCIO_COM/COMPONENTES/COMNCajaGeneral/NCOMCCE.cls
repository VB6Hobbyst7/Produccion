VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************
'** Nombre : NCOMCCE
'** Descripción : Para el Registro de Transacciones de Operaciones con la CCE
'** Creación : PASI, 20160613
'**********************************************************************
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim oConect As COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling
Dim sSql As String
Dim rs As adodb.Recordset
Public Sub RollbackTrans()
    oConect.RollbackTrans
End Sub
Public Sub CommitTrans()
    oConect.CommitTrans
End Sub
Public Sub BeginTrans()
    oConect.BeginTrans
End Sub
Private Sub Class_Initialize()
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Call oError.RaiseError(oError.MyUnhandledError, "DMov:Initialize Method. Error en Conexion a Base de datos")
End Sub
Private Sub Class_Terminate()
    oConect.CierraConexion
    Set oConect = Nothing
End Sub
Public Function CCE_ObtieneIFIS() As adodb.Recordset
    sSql = "stp_sel_CCE_ObtieneIFIS"
    Set CCE_ObtieneIFIS = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ConsultaCuentaCCI(ByVal psCtaCod As String) As adodb.Recordset
    sSql = "Exec stp_sel_CCE_ConsultarCuentaCCI '" & psCtaCod & "'"
    Set CCE_ConsultaCuentaCCI = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ObtieneDatosPersona(ByVal psCtaCod As String) As adodb.Recordset
    sSql = "Exec stp_sel_CCE_ObtieneDatosPersona '" & psCtaCod & "'"
    Set CCE_ObtieneDatosPersona = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ObtieneDocumentoPersona(ByVal psPersCod As String) As adodb.Recordset
    sSql = "Exec stp_sel_CCE_ObtieneDocumentoPersona '" & psPersCod & "'"
    Set CCE_ObtieneDocumentoPersona = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ValidaCuentaCCI(ByVal psCtaCod As String) As Boolean
    Dim rs As adodb.Recordset
    sSql = "stp_sel_CCE_ValidaCuentaCCI '" & psCtaCod & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    CCE_ValidaCuentaCCI = rs!nResp
End Function

Public Function CCE_RegistrarTransferenciaOrdinaria(ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String, ByVal psOpeCod As CaptacOperacion, ByVal psGlosa As String, _
                                                    ByVal psCodBCRBcoDest As String, ByVal psPersCodOrd As String, ByVal psPersNombreOrd As String, ByVal pnTpoDocOrd As Integer, _
                                                    ByVal psNroDocOrd As String, ByVal psDireccionOrd As String, ByVal psTelefOrd As String, ByVal pnMoneda As Integer, ByVal pnMontoIncComis As Integer, ByVal pnMonto As Currency, _
                                                    ByVal psCCIDestCod As String, ByVal pnCargoCta As Integer, ByVal psCtaCod As String, psCCIOrigCod As String, ByVal pnMismoTit As Integer, _
                                                    ByVal psTariCod As String, ByVal psTariCriApl As String, ByVal psComisSigno As String, ByVal pnComisMonto As Currency, ByVal pnCMACMComisMonto As Currency, _
                                                    Optional ByVal pnCMACMMontoITF As Currency = 0) As Long
Dim clsCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lnMovNroOffHost As Long
Dim lnTranNro As Long
Dim lsMovNro As String
Dim lsCodOficDestino, lsCodAplicacion As String 'VAPA20170628
Dim lbSabado, lbEnvioTrama, lbTRTSinEnviar As Boolean 'vapa20170628 agregado
Dim ldFechaPresentacion As Date
Dim lsFechaNull As String
Dim lnId As Long

    On Error GoTo ErrRegistraTransf
    Set clsCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
    clsCapMov.dBeginTrans
    
        lbTrans = True
        lnTranNro = clsCapMov.CCE_RegistraTransferencia(gCCETransfOrdinaria, psCodBCRBcoDest, psPersCodOrd, psPersNombreOrd, pnTpoDocOrd, psNroDocOrd, psDireccionOrd, psTelefOrd, pnMoneda, pnMonto)
        'psCCIDestCod
        lbSabado = CCE_EsSabado(pdFecSis)
        lsCodOficDestino = Mid(psCCIDestCod, 4, 3)
        lsCodAplicacion = CCE_ObtieneAplicacion
        lsCodAplicacion = CCE_ValidaAplicacion(lsCodAplicacion, psCodBCRBcoDest)
        lbEnvioTrama = CCE_SeEnvioTrama(lsCodAplicacion, pdFecSis)
        lbTRTSinEnviar = False
        
        If lbEnvioTrama = True Then
            Select Case lsCodAplicacion
                    Case "TRI"
                    lsCodAplicacion = "TRM"
                    Case "TRM"
                    lsCodAplicacion = "TRT"
                    Case "TRT"
                    lsCodAplicacion = "TRI"
                    lbTRTSinEnviar = True
            End Select
        End If
        
        If lbSabado = True Then
        lsCodAplicacion = "TRI"
        lsCodAplicacion = CCE_ValidaAplicacion(lsCodAplicacion, psCodBCRBcoDest)
        ldFechaPresentacion = CCE_ObtieneFecPre(pdFecSis, lsCodAplicacion, psCodBCRBcoDest, lsCodOficDestino)
        End If
                   
        If lbSabado = False Then
                    If lbTRTSinEnviar = True Then
                    lsCodAplicacion = CCE_ValidaAplicacion(lsCodAplicacion, psCodBCRBcoDest)
                    ldFechaPresentacion = CCE_ObtieneFecPre(pdFecSis, lsCodAplicacion, psCodBCRBcoDest, lsCodOficDestino)
                    End If
        End If
        If lbTRTSinEnviar = True Or lbSabado = True Then
        clsCapMov.CCE_RegistraTransfOrdinariaSabado lnTranNro, psCCIDestCod, pnCargoCta, psCtaCod, psCCIOrigCod, pnMismoTit, pnMontoIncComis, pnMonto, psTariCod, psTariCriApl, psComisSigno, pnComisMonto, lsCodAplicacion, ldFechaPresentacion
        Else
        clsCapMov.CCE_RegistraTransfOrdinariaSabado lnTranNro, psCCIDestCod, pnCargoCta, psCtaCod, psCCIOrigCod, pnMismoTit, pnMontoIncComis, pnMonto, psTariCod, psTariCriApl, psComisSigno, pnComisMonto, lsCodAplicacion
        End If
        lsMovNro = clsCapMov.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
        
        
         'clsCapMov.AgregaMov lsMovNro, IIf(Len(psCtaCod) = 0, psOpeCod, IIf(Mid(psCtaCod, 6, 3) = "232", gCMCCETransfOrdCargoCtaAho, gCMCCETransfOrdCargoCtaCts)), psGlosa, gMovEstContabNoContable, gMovFlagVigente
         clsCapMov.AgregaMov lsMovNro, IIf(Len(psCtaCod) = 0, psOpeCod, IIf(Mid(psCtaCod, 6, 3) = "232", gCCETraOrdCtaAho, gCMCCETransfOrdCargoCtaCts)), psGlosa, gMovEstContabMovContable, gMovFlagVigente
        lnMovNro = clsCapMov.GetnMovNro(lsMovNro)
        clsCapMov.AgregaMovGasto lnMovNro, psPersCodOrd, ""
        'pagoenefectivo
        If Len(psCtaCod) = 0 Then
            clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnMonto, pnMoneda, psOpeCod, gCMCCEConceptoMontoEnvioTrCCE
            clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnComisMonto, pnMoneda, psOpeCod, gCMCCEConceptoComisEnvioTrCCE
            If pnCMACMComisMonto > 0 Then clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnCMACMComisMonto, pnMoneda, psOpeCod, gCMCCEConceptoComisionTrCMACM
            If pnCMACMMontoITF > 0 Then clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnCMACMMontoITF, pnMoneda, psOpeCod, gCMCCEConceptoITFTrCMACM
        Else
            'vapa20170119
            'lnMovNroOffHost = clsCapMov.CCE_CargoCta_OffHost(pdFecSis, psCodAge, psCtaCod, pnMonto + pnComisMonto, pnCMACMMontoITF, pnCMACMComisMonto, pnMoneda, IIf(Mid(psCtaCod, 6, 3) = "232", gCMCCETransfOrdCargoCtaAho, gCMCCETransfOrdCargoCtaCts), IIf(Mid(psCtaCod, 6, 3) = "232", gCMCCEComisionTransferenciaAho, gCMCCEComisionTransferenciaCTS), gCMCCEITFTransferenciaAho, psCodUser)
            'lnMovNroOffHost = clsCapMov.CCE_CargoCta_OffHost(pdFecSis, psCodAge, psCtaCod, pnMonto, pnComisMonto, pnCMACMMontoITF, pnCMACMComisMonto, pnMoneda, IIf(Mid(psCtaCod, 6, 3) = "232", gCMCCETransfOrdCargoCtaAho, gCMCCETransfOrdCargoCtaCts), gCMCCEComisionTransferenciaCCE, IIf(Mid(psCtaCod, 6, 3) = "232", gCMCCEComisionTransferenciaAho, gCMCCEComisionTransferenciaCTS), gCMCCEITFTransferenciaAho, psCodUser)
            lnMovNroOffHost = clsCapMov.CCE_CargoCta(pdFecSis, psCodAge, psCtaCod, pnMonto, pnComisMonto, pnCMACMMontoITF, pnCMACMComisMonto, pnMoneda, IIf(Mid(psCtaCod, 6, 3) = "232", gCCETraOrdCtaAho, gCMCCETransfOrdCargoCtaCts), gCCEComisonTraCCE, IIf(Mid(psCtaCod, 6, 3) = "232", gCMCCEComisionTransferenciaAho, gCMCCEComisionTransferenciaCTS), gCMCCEITFTransferenciaAho, psCodUser, lnMovNro, lsMovNro)
                        
            If Not lnMovNroOffHost > 0 Then
                clsCapMov.dRollbackTrans: Set clsCapMov = Nothing
                lbTrans = False
                CCE_RegistrarTransferenciaOrdinaria = 0
                Exit Function
            End If
        End If
        clsCapMov.CCE_RegistrarMovRef lnMovNro, lnTranNro
       ' If lnMovNroOffHost > 0 Then clsCapMov.CCE_RegistrarMovOffHostRef lnMovNroOffHost, lnTranNro 'VAPA20170613 comentado
    clsCapMov.dCommitTrans
    lbTrans = False
    CCE_RegistrarTransferenciaOrdinaria = lnMovNro
    Set clsCapMov = Nothing
    Exit Function
ErrRegistraTransf:
    CCE_RegistrarTransferenciaOrdinaria = 0
    If lbTrans Then
        clsCapMov.dRollbackTrans
        Set clsCapMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CCE_RegistrarTransferenciaPagoHaberes(ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String, ByVal psOpeCod As CaptacOperacion, ByVal psGlosa As String, _
                                                    ByVal psCodBCRBcoDest As String, ByVal psPersCodOrd As String, ByVal psPersNombreOrd As String, ByVal pnTpoDocOrd As Integer, _
                                                    ByVal psNroDocOrd As String, ByVal psDireccionOrd As String, ByVal psTelefOrd As String, ByVal pnMoneda As Integer, ByVal pnMonto As Currency, _
                                                    ByVal psCCIDestCod As String, ByVal psPersCodBenef As String, ByVal psPersNombreBenef As String, ByVal pnTpoDocBenef As Integer, _
                                                    ByVal psNroDocBenef As String, ByVal psDireccionBenef As String, ByVal psTelefBenef As String, ByVal pnTpoHaberCod As Integer, _
                                                    ByVal psTariCod As String, ByVal psTariCriApl As String, ByVal psComisSigno As String, ByVal pnComisMonto As Currency, ByVal pnCMACMComisMonto As Currency, _
                                                    Optional ByVal pnCMACMMontoITF As Currency = 0) As Long
Dim clsCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lnMovNroRef As Long
Dim lnTranNro As Long
Dim lsMovNro As String
Dim lnId As Long
On Error GoTo ErrRegistraTransf
    Set clsCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
     clsCapMov.dBeginTrans
        lbTrans = True
        
        lnTranNro = clsCapMov.CCE_RegistraTransferencia(gCCETransfPagoHaberes, psCodBCRBcoDest, psPersCodOrd, psPersNombreOrd, pnTpoDocOrd, psNroDocOrd, psDireccionOrd, psTelefOrd, pnMoneda, pnMonto)
        clsCapMov.CCE_RegistraTransfPagoHaberes lnTranNro, psCCIDestCod, psPersCodBenef, psPersNombreBenef, pnTpoDocBenef, psNroDocBenef, psDireccionBenef, psTelefBenef, pnTpoHaberCod, pnMonto, _
                                                psTariCod, psTariCriApl, psComisSigno, pnComisMonto
        lsMovNro = clsCapMov.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
        clsCapMov.AgregaMov lsMovNro, psOpeCod, psGlosa, gMovEstContabNoContable, gMovFlagVigente
        lnMovNro = clsCapMov.GetnMovNro(lsMovNro)
        clsCapMov.AgregaMovGasto lnMovNro, psPersCodOrd, psPersCodBenef
        clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnMonto, pnMoneda, psOpeCod, gCMCCEConceptoMontoEnvioTrCCE
        clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnComisMonto, pnMoneda, psOpeCod, gCMCCEConceptoComisEnvioTrCCE
        If pnCMACMComisMonto > 0 Then clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnCMACMComisMonto, pnMoneda, psOpeCod, gCMCCEConceptoComisionTrCMACM
        If pnCMACMMontoITF > 0 Then clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnCMACMMontoITF, pnMoneda, psOpeCod, gCMCCEConceptoITFTrCMACM
        clsCapMov.CCE_RegistrarMovRef lnMovNro, lnTranNro
        
    clsCapMov.dCommitTrans
    lbTrans = False
    CCE_RegistrarTransferenciaPagoHaberes = lnMovNro
    Set clsCapMov = Nothing
Exit Function
ErrRegistraTransf:
    CCE_RegistrarTransferenciaPagoHaberes = 0
    If lbTrans Then
        clsCapMov.dRollbackTrans
        Set clsCapMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CCE_RegistrarTransferenciaPagoCTS(ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String, ByVal psOpeCod As CaptacOperacion, ByVal psGlosa As String, _
                                                    ByVal psCodBCRBcoDest As String, ByVal psPersCodOrd As String, ByVal psPersNombreOrd As String, ByVal pnTpoDocOrd As Integer, _
                                                    ByVal psNroDocOrd As String, ByVal psDireccionOrd As String, ByVal psTelefOrd As String, ByVal pnMoneda As Integer, ByVal pnMonto As Currency, _
                                                    ByVal psCCIDestCod As String, ByVal psPersCodBenef As String, ByVal psPersNombreBenef As String, ByVal pnTpoDocBenef As Integer, ByVal psNroDocBenef As String, _
                                                    ByVal psDireccionBenef As String, ByVal psTelefBenef As String, _
                                                    ByVal psMesPago As String, ByVal psAnio As String, ByVal pnImporteNSueldosBrutos As Currency, _
                                                    ByVal psTariCod As String, ByVal psTariCriApl As String, ByVal psComisSigno As String, ByVal pnComisMonto As Currency, ByVal pnCMACMComisMonto As Currency, _
                                                    Optional ByVal pnCMACMMontoITF As Currency = 0) As Long
Dim clsCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lnMovNroRef As Long
Dim lnTranNro As Long
Dim lsMovNro As String
Dim lnId As Long
Dim lnMaxNroReg As Long
On Error GoTo ErrRegistraTransf
    Set clsCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
    clsCapMov.dBeginTrans
        lbTrans = True
        lnTranNro = clsCapMov.CCE_RegistraTransferencia(gCCETransfPagoCTS, psCodBCRBcoDest, psPersCodOrd, psPersNombreOrd, pnTpoDocOrd, psNroDocOrd, psDireccionOrd, psTelefOrd, pnMoneda, pnMonto)
        clsCapMov.CCE_RegistraTransfPagoCTS lnTranNro, psCCIDestCod, psPersCodBenef, psPersNombreBenef, pnTpoDocBenef, psNroDocBenef, psDireccionBenef, psTelefBenef, psMesPago, psAnio, pnImporteNSueldosBrutos, pnMonto, _
                                        psTariCod, psTariCriApl, psComisSigno, pnComisMonto
        lsMovNro = clsCapMov.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
        clsCapMov.AgregaMov lsMovNro, psOpeCod, psGlosa, gMovEstContabNoContable, gMovFlagVigente
        lnMovNro = clsCapMov.GetnMovNro(lsMovNro)
        clsCapMov.AgregaMovGasto lnMovNro, psPersCodOrd, psPersCodBenef
        clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnMonto, pnMoneda, psOpeCod, gCMCCEConceptoMontoEnvioTrCCE
        clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnComisMonto, pnMoneda, psOpeCod, gCMCCEConceptoComisEnvioTrCCE
        If pnCMACMComisMonto > 0 Then clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnCMACMComisMonto, pnMoneda, psOpeCod, gCMCCEConceptoComisionTrCMACM
        If pnCMACMMontoITF > 0 Then clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnCMACMMontoITF, pnMoneda, psOpeCod, gCMCCEConceptoITFTrCMACM
        clsCapMov.CCE_RegistrarMovRef lnMovNro, lnTranNro
        
    clsCapMov.dCommitTrans
    lbTrans = False
    CCE_RegistrarTransferenciaPagoCTS = lnMovNro
    Set clsCapMov = Nothing
Exit Function
ErrRegistraTransf:
    CCE_RegistrarTransferenciaPagoCTS = 0
    If lbTrans Then
        clsCapMov.dRollbackTrans
        Set clsCapMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CCE_RegistraTransferenciaPagoTarjCred(ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String, ByVal psOpeCod As CaptacOperacion, ByVal psGlosa As String, _
                                                     ByVal psCodBCRBcoDest As String, ByVal psPersCodOrd As String, ByVal psPersNombreOrd As String, ByVal pnTpoDocOrd As Integer, _
                                                    ByVal psNroDocOrd As String, ByVal psDireccionOrd As String, ByVal psTelefOrd As String, ByVal pnMoneda As Integer, ByVal pnMontoIncComis As Integer, ByVal pnMonto As Currency, _
                                                    ByVal psNumTarj As String, ByVal psPersCodBenef As String, ByVal psPersNombreBenef As String, ByVal pnTpoDocBenef As Integer, ByVal psNroDocBenef As String, ByVal psDireccionBenef As String, ByVal psTelefBenef As String, _
                                                    ByVal psTariCod As String, ByVal psTariCriApl As String, ByVal psComisSigno As String, ByVal pnComisMonto As Currency, ByVal pnCMACMComisMonto As Currency, _
                                                    Optional ByVal pnCMACMMontoITF As Currency = 0) As Long
Dim clsCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lnMovNroRef As Long
Dim lnTranNro As Long
Dim lsMovNro As String
Dim lnId As Long
On Error GoTo ErrRegistraTransf
    Set clsCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
    clsCapMov.dBeginTrans
        lbTrans = True
        lnTranNro = clsCapMov.CCE_RegistraTransferencia(gCCETransfPagoTarjCred, psCodBCRBcoDest, psPersCodOrd, psPersNombreOrd, pnTpoDocOrd, psNroDocOrd, psDireccionOrd, psTelefOrd, pnMoneda, pnMonto)
        clsCapMov.CCE_RegistraTransfPagTarjCred lnTranNro, psNumTarj, psPersCodBenef, psPersNombreBenef, pnTpoDocBenef, psNroDocBenef, psDireccionBenef, psTelefBenef, pnMontoIncComis, pnMonto, _
                                            psTariCod, psTariCriApl, psComisSigno, pnComisMonto
        lsMovNro = clsCapMov.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
        clsCapMov.AgregaMov lsMovNro, psOpeCod, psGlosa, gMovEstContabNoContable, gMovFlagVigente
        lnMovNro = clsCapMov.GetnMovNro(lsMovNro)
        clsCapMov.AgregaMovGasto lnMovNro, psPersCodOrd, psPersCodBenef
        clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnMonto, pnMoneda, psOpeCod, gCMCCEConceptoMontoEnvioTrCCE
        If pnCMACMComisMonto > 0 Then clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnCMACMComisMonto + pnComisMonto, pnMoneda, psOpeCod, gCMCCEConceptoComisionTrCMACM
        If pnCMACMMontoITF > 0 Then clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnCMACMMontoITF, pnMoneda, psOpeCod, gCMCCEConceptoITFTrCMACM
        clsCapMov.CCE_RegistrarMovRef lnMovNro, lnTranNro
        
    clsCapMov.dCommitTrans
    lbTrans = False
    CCE_RegistraTransferenciaPagoTarjCred = lnMovNro
    Set clsCapMov = Nothing
Exit Function
ErrRegistraTransf:
    CCE_RegistraTransferenciaPagoTarjCred = 0
    If lbTrans Then
        clsCapMov.dRollbackTrans
        Set clsCapMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CCE_RegistraTransferenciaPagoProveedor(ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String, ByVal psOpeCod As CaptacOperacion, ByVal psGlosa As String, _
                                                    ByVal psCodBCRBcoDest As String, ByVal psPersCodOrd As String, ByVal psPersNombreOrd As String, ByVal pnTpoDocOrd As Integer, _
                                                    ByVal psNroDocOrd As String, ByVal psDireccionOrd As String, ByVal psTelefOrd As String, ByVal pnMoneda As Integer, ByVal pnMonto As Currency, _
                                                    ByVal psCCIDestCod As String, ByVal psPersCodBenef As String, ByVal psPersNombreBenef As String, ByVal pnTpoDocBenef As Integer, ByVal psNroDocBenef As String, _
                                                    ByVal psDireccionBenef As String, ByVal psTelefBenef As String, _
                                                    ByVal pnTpoDocComp As Integer, ByVal psNroDocComp As String, ByVal pdFecEmisionComp As Date, _
                                                    ByVal psTariCod As String, ByVal psTariCriApl As String, ByVal psComisSigno As String, ByVal pnComisMonto As Currency, ByVal pnCMACMComisMonto As Currency, _
                                                    Optional ByVal pnCMACMMontoITF As Currency = 0) As Long
Dim clsCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lsMovNro As String
Dim lnMovNroRef As Long
Dim lnTranNro As Long
Dim lnId As Long
On Error GoTo ErrRegistraTransf
    Set clsCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
    clsCapMov.dBeginTrans
        lbTrans = True
        
        lnTranNro = clsCapMov.CCE_RegistraTransferencia(gCCETransfPagoProveedor, psCodBCRBcoDest, psPersCodOrd, psPersNombreOrd, pnTpoDocOrd, psNroDocOrd, psDireccionOrd, psTelefOrd, pnMoneda, pnMonto)
        clsCapMov.CCE_RegistraTransfPagoProveedor lnTranNro, psCCIDestCod, psPersCodBenef, psPersNombreBenef, pnTpoDocBenef, psNroDocBenef, psDireccionBenef, psTelefBenef, pnTpoDocComp, psNroDocComp, pdFecEmisionComp, pnMonto, _
                                            psTariCod, psTariCriApl, psComisSigno, pnComisMonto
        lsMovNro = clsCapMov.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
        clsCapMov.AgregaMov lsMovNro, psOpeCod, psGlosa, gMovEstContabNoContable, gMovFlagVigente
        lnMovNro = clsCapMov.GetnMovNro(lsMovNro)
        clsCapMov.AgregaMovGasto lnMovNro, psPersCodOrd, psPersCodBenef
        clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnMonto, pnMoneda, psOpeCod, gCMCCEConceptoMontoEnvioTrCCE
        clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnComisMonto, pnMoneda, psOpeCod, gCMCCEConceptoComisEnvioTrCCE
        If pnCMACMComisMonto > 0 Then clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnCMACMComisMonto, pnMoneda, psOpeCod, gCMCCEConceptoComisionTrCMACM
        If pnCMACMMontoITF > 0 Then clsCapMov.CCE_RegistraMovOpeVarias lnMovNro, pnCMACMMontoITF, pnMoneda, psOpeCod, gCMCCEConceptoITFTrCMACM
        clsCapMov.CCE_RegistrarMovRef lnMovNro, lnTranNro
        
    clsCapMov.dCommitTrans
    lbTrans = False
    CCE_RegistraTransferenciaPagoProveedor = lnMovNro
    Set clsCapMov = Nothing
Exit Function
ErrRegistraTransf:
    CCE_RegistraTransferenciaPagoProveedor = 0
    If lbTrans Then
        clsCapMov.dRollbackTrans
        Set clsCapMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CCE_ObtieneComisionTransf(ByVal psOpeTransf As String, ByVal psAgeCod As String, ByVal psCodBCRDest As String, ByVal psOficCodDest As String, ByVal pnMonto As Currency, ByVal pnMoneda As Integer) As adodb.Recordset
    sSql = "stp_sel_CCE_ObtieneComisionTransf '" & psOpeTransf & "','" & psAgeCod & "','" & psCodBCRDest & "','" & psOficCodDest & "'," & pnMonto & "," & pnMoneda
    Set CCE_ObtieneComisionTransf = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ObtieneSubTpoHaberes() As adodb.Recordset
    sSql = "stp_sel_CCE_ObtieneSubTpoHaberes"
    Set CCE_ObtieneSubTpoHaberes = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ObtieneOpeDocProveedor() As adodb.Recordset
    sSql = "stp_sel_CCE_ObtieneOpeDocProveedor"
    Set CCE_ObtieneOpeDocProveedor = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ObtieneIFIxCodBCR(ByVal psCodBCR As String) As adodb.Recordset
    sSql = "stp_sel_CCE_ObtieneIFIxCodBCR '" & psCodBCR & "'"
    Set CCE_ObtieneIFIxCodBCR = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_GeneraTrama_TRAPRE(ByVal psCodInstrumento As String, ByVal psCodArchivo As String, psCodSesion As String, ByVal psCodAplicacion As String, ByVal pdFecSis As Date, ByVal psCodUser As String) As adodb.Recordset
    sSql = "STP_SEL_CCE_TRA_GENERA_TRAMA_ENVIO '" & psCodInstrumento & "','" & psCodInstrumento & "','" & psCodSesion & "','" & psCodAplicacion & "','" & Format(pdFecSis, "YYYYmmdd") & "','" & psCodUser & "'"
    Set CCE_GeneraTrama_TRAPRE = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ObtieneLimitesTransf() As adodb.Recordset
    sSql = "stp_sel_CCE_ObtieneLimitesTransf"
    Set CCE_ObtieneLimitesTransf = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ListaOpeTransf() As adodb.Recordset
    sSql = "stp_sel_CCE_ListaOpeTransf "
    Set CCE_ListaOpeTransf = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ListaTransfParaExtorno(ByVal pdFecha As Date, ByVal psOpeCod As String, ByVal pnMovNro As Long) As adodb.Recordset
    sSql = "stp_sel_CCE_ListaTransfxExtorno '" & Format(pdFecha, "YYYYMMDD") & "','" & psOpeCod & "'," & pnMovNro
    Set CCE_ListaTransfParaExtorno = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ExtornaTransf(ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String, ByVal psOpeCod As String, psGlosa As String, ByVal pnTranNro As Long)
Dim clsCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lsMovNro As String
On Error GoTo ErrExtornaTransf
    Set clsCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
    clsCapMov.dBeginTrans
        lbTrans = True
        lsMovNro = clsCapMov.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
        clsCapMov.AgregaMov lsMovNro, psOpeCod, psGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
        lnMovNro = clsCapMov.GetnMovNro(lsMovNro)
        clsCapMov.CCE_ExtornaTransf lnMovNro, pnTranNro
   clsCapMov.dCommitTrans
   lbTrans = False
   CCE_ExtornaTransf = lnMovNro 'VAPA20170708
   Set clsCapMov = Nothing
Exit Function
ErrExtornaTransf:
    If lbTrans Then
        clsCapMov.dRollbackTrans
        Set clsCapMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CCE_ObtieneCuentaDesdeCCI(ByVal psCtaCodCci As String) As String
     Dim oConect As COMConecta.DCOMConecta
     Set oConect = New COMConecta.DCOMConecta
    oConect.AbreConexion 'VAPA20170323
    sSql = "select dbo.Fnc_CCE_ObtieneCuentaDesdeCCI('" & psCtaCodCci & "') cCtaCod"
    Set rs = oConect.CargaRecordSet(sSql)
    CCE_ObtieneCuentaDesdeCCI = rs!cCtaCod
    oConect.CierraConexion 'VAPA20170323
    Set oConect = Nothing
End Function
Public Function fgDiferenciaRedondeoITF(ByVal lnITF As Double) As Double
    Dim lnPos As Integer
    Dim DifITF As Double
    lnPos = InStr(1, Trim(Str(lnITF)), ".")
    If lnPos > 0 Then
        DifITF = Round((lnITF * 100 Mod 10) / 100, 2)
        DifITF = IIf(DifITF = 0.05, 0, IIf(DifITF > 0.05, DifITF - 0.05, DifITF))
    Else
        DifITF = 0
    End If
    fgDiferenciaRedondeoITF = Round(DifITF, 2)
End Function
Public Function DameMontoITF(ByVal pnMonto As Double) As Double
    Dim oITF As New COMDConstSistema.FCOMITF
    Dim lnMonto As Double, lnITF As Double, lnRedondeoITF As Double
    oITF.fgITFParametros
    lnMonto = pnMonto
    lnITF = Format(oITF.fgITFCalculaImpuesto(lnMonto), "#,##0.00")
    lnRedondeoITF = fgDiferenciaRedondeoITF(lnITF)
    If lnRedondeoITF > 0 Then
        lnITF = lnITF - lnRedondeoITF
    End If
    DameMontoITF = lnITF
    Set oITF = Nothing
End Function
Public Function CCE_RegistraIngresoPorComision(ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMontoTranComis As Double, ByVal pnMoneda As String, ByVal pnIdRef As Long, ByVal psEfinCodBCR) As Long
    Dim clsCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento

    Dim lbTrans As Boolean
    Dim lnMovNro As Long
    Dim lsMovNro As String
On Error GoTo ErrRegistraIngreso
    Set clsCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
    clsCapMov.dBeginTrans
        lbTrans = True
        lsMovNro = clsCapMov.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
        clsCapMov.AgregaMov lsMovNro, gCMCCEDepComisionTransferencia, "CCE - INGRESO POR COMISIÓN", gMovEstContabNoContable, gMovFlagVigente
        lnMovNro = clsCapMov.GetnMovNro(lsMovNro)
        clsCapMov.AgregaMovOpeVarias lnMovNro, pnMontoTranComis, "", "", pnMoneda, gCMCCEDepComisionTransferencia, gCMCCECptoComisTrCMACMRec
        clsCapMov.CCE_RegistrarRecRef lnMovNro, pnIdRef, psEfinCodBCR
    clsCapMov.dCommitTrans
    lbTrans = False
    CCE_RegistraIngresoPorComision = lnMovNro
    Set clsCapMov = Nothing
Exit Function
ErrRegistraIngreso:
    CCE_RegistraIngresoPorComision = 0
    If lbTrans Then
        clsCapMov.dRollbackTrans
        Set clsCapMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'GITU - PASI 20161103 CCE ************************************
Public Function CCE_DepositoOrd(ByVal pdFecha As Date, ByVal psCtaCodCci As String, ByVal pnMontoTran As Double, ByVal pnMoneda As String, ByVal pnMontoTranComis As Double, ByVal pbMismoTitular As Boolean, ByVal pnIdRef As Long) As Long
    Dim oConect As COMConecta.DCOMConecta 'vapa20170323
    Dim oMov As New COMNContabilidad.NCOMContFunciones
    Dim clsCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim nMonto As Double, nSaldo As Double
    Dim lnMontoITF As Double
    Dim lnRedondeoITF As Double
    Dim sMovNro As String
    Dim lnMovNro As Long
    Dim psCtaCod As String
    Set oConect = New COMConecta.DCOMConecta 'vapa20170323
     Set clsCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
     oConect.AbreConexion 'vapa20170323
    If Not pbMismoTitular Then lnMontoITF = DameMontoITF(pnMontoTran)
    nMonto = pnMontoTran - lnMontoITF
    psCtaCod = CCE_ObtieneCuentaDesdeCCI(psCtaCodCci)
    sMovNro = clsCapMov.GeneraMovNro(pdFecha, Right(Mid(psCtaCod, 4, 2), 2), "UCCE")
    sSql = "stp_ins_DepositoCCE '" & Format(pdFecha, "YYYYmmdd") & "','" & psCtaCod & "'," & pnMontoTran & "," & lnMontoITF & "," & pnMoneda & ",'" & gCMCCEDepTransfOrdinaria & "','" & gCMCCEDepTransfITF & "','UCCE','" & sMovNro & "'," & pnIdRef & ",'" & Left(psCtaCodCci, 3) & "'"
    Set rs = oConect.CargaRecordSet(sSql)
   'Set oConect = Nothing 'VAPA20170323 COMENTADO
    If Not (rs.EOF And rs.BOF) Then
        If Not rs!nResultado = 0 Then
            Sleep 1000
            CCE_RegistraIngresoPorComision pdFecha, Right(Mid(psCtaCod, 4, 2), 2), "UCCE", pnMontoTranComis, pnMoneda, pnIdRef, Left(psCtaCodCci, 3)
        
        End If
        CCE_DepositoOrd = rs!nResultado
    Else
        CCE_DepositoOrd = 0
    End If
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function CCE_PagoHaberes(ByVal pdFecha As Date, ByVal psCtaCodCci As String, ByVal pnMontoTran As Double, ByVal pnMoneda As String, ByVal pnMontoTranComis As Double, ByVal pnSubTpo As Integer, ByVal pnIdRef As Long) As Long
    Dim oMov As New COMNContabilidad.NCOMContFunciones
    Dim nMonto As Double, nSaldo As Double
    Dim lnMontoITF As Double
    Dim lnRedondeoITF As Double
    Dim sMovNro As String
    Dim psCtaCod As String
    If pnSubTpo = 4 Or pnSubTpo = 8 Then lnMontoITF = DameMontoITF(pnMontoTran)
    nMonto = pnMontoTran - lnMontoITF
    psCtaCod = CCE_ObtieneCuentaDesdeCCI(psCtaCodCci)
    sMovNro = oMov.GeneraMovNro(pdFecha, Right(Mid(psCtaCod, 4, 2), 2), "UCCE")
    sSql = "stp_ins_DepositoCCE '" & Format(pdFecha, "YYYYmmdd") & "','" & psCtaCod & "'," & pnMontoTran & "," & lnMontoITF & "," & pnMoneda & ",'" & gCMCCEDepTransfPagoHaberes & "','" & gCMCCEDepTransfITF & "','UCCE','" & sMovNro & "'," & pnIdRef
    Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        CCE_RegistraIngresoPorComision pdFecha, Right(Mid(psCtaCod, 4, 2), 2), "UCCE", pnMontoTranComis, pnMoneda, pnIdRef, Left(psCtaCodCci, 3)
        CCE_PagoHaberes = rs!nResultado
    Else
        CCE_PagoHaberes = 0
    End If
End Function
Public Function CCE_PagoProveedores(ByVal pdFecha As Date, ByVal psCtaCodCci As String, ByVal pnMontoTran As Double, ByVal pnMoneda As String, ByVal pnMontoTranComis As Double, ByVal pnIdRef As Long) As Long
    Dim oMov As New COMNContabilidad.NCOMContFunciones
    Dim nMonto As Double, nSaldo As Double
    Dim lnMontoITF As Double
    Dim lnRedondeoITF As Double
    Dim sMovNro As String
    Dim psCtaCod As String
    nMonto = pnMontoTran - lnMontoITF
    psCtaCod = CCE_ObtieneCuentaDesdeCCI(psCtaCodCci)
    sMovNro = oMov.GeneraMovNro(pdFecha, Right(Mid(psCtaCod, 4, 2), 2), "UCCE")
    sSql = "stp_ins_DepositoCCE '" & Format(pdFecha, "YYYYmmdd") & "','" & psCtaCod & "'," & pnMontoTran & "," & lnMontoITF & "," & pnMoneda & ",'" & gCMCCEDepTransfPagoProveedor & "','" & gCMCCEDepTransfITF & "','UCCE','" & sMovNro & "'," & pnIdRef
    Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        CCE_RegistraIngresoPorComision pdFecha, Right(Mid(psCtaCod, 4, 2), 2), "UCCE", pnMontoTranComis, pnMoneda, pnIdRef, Left(psCtaCodCci, 3)
        CCE_PagoProveedores = rs!nResultado
    Else
        CCE_PagoProveedores = 0
    End If
End Function
Public Function CCE_DepositoCTS(ByVal pdFecha As Date, ByVal psCtaCodCci As String, ByVal pnMontoTran As Double, ByVal pnMoneda As String, ByVal pnMontoTranComis As Double, ByVal pnMontoAcum As Double, ByVal pnIdRef As Long) As Long
    Dim oMov As New COMNContabilidad.NCOMContFunciones
    Dim nMonto As Double, nSaldo As Double
    Dim psCtaCod As String
    psCtaCod = CCE_ObtieneCuentaDesdeCCI(psCtaCodCci)
    sSql = "stp_ins_DepositoCTSCCE '" + psCtaCod + "', " & pnMontoTran & "'," & pnMoneda & ",'" & Right(Mid(psCtaCod, 4, 2), 2) & "','" + gCMCCEDepTransfPagoCTS + "','UCCE'," & pnMontoAcum & "," & pnIdRef
    Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        CCE_RegistraIngresoPorComision pdFecha, Right(Mid(psCtaCod, 4, 2), 2), "UCCE", pnMontoTranComis, pnMoneda, pnIdRef, Left(psCtaCodCci, 3)
        CCE_DepositoCTS = rs!nResultado
    Else
        CCE_DepositoCTS = 0
    End If
End Function
'END GITU - PASI*****************************************************
Public Function CCE_Confirma_CHE(ByVal pdFecha As Date, ByVal psAgeCod As String, ByVal psIFICodDeb As String, _
                                ByVal psCtaCodDeb As String, ByVal psNroCheque As String) As Boolean
Dim oMov As New COMDMov.DCOMMov
Dim oRSDetCheque As New adodb.Recordset
Dim oRSBcoDep As New adodb.Recordset
Dim lbTrans As Boolean
Dim lsMovNro As String
Dim lnMovNro As Long, lnMovNroRef As Long, lnId As Long
Dim lsIFTpo As String, lsPersCod As String, lsCtaIfCod As String
Dim lsCtaContDepositoD As String, lsCtaContDepositoH As String, lsIDs As String
Dim nMonedaDetChe As Integer

    On Error GoTo ErrConfirmaCHE
    oMov.BeginTrans
    lbTrans = True
    
    
    Set oRSDetCheque = CCE_ObtieneChequeDet(psIFICodDeb, psCtaCodDeb, psNroCheque)
    nMonedaDetChe = oRSDetCheque!nMoneda 'agregado por VAPA20170505
    Set oRSBcoDep = CCE_ObtieneBcoDepxCheque(nMonedaDetChe) 'modificado por VAPA20170505
    If Not oRSBcoDep.EOF Then
        lsPersCod = oRSBcoDep!cPersCod
        lsIFTpo = oRSBcoDep!cIFTpo
        lsCtaIfCod = oRSBcoDep!cCtaIfCod
        lsCtaContDepositoD = oRSBcoDep!cCtaContDepositoD
        lsCtaContDepositoH = oRSBcoDep!cCtaContDepositoH
    End If
    If Not oRSDetCheque.EOF Then
        Do While Not oRSDetCheque.EOF
            lnMovNroRef = oRSDetCheque!nMovNro
            Sleep 1000
            lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, "UCCE")
            oMov.InsertaMov lsMovNro, gChqOpeDeposito, "CCE - Depósito Cheque Nro. " & oRSDetCheque!cNroDoc, gMovEstContabMovContable, gMovFlagVigente
            lnMovNro = oMov.GetnMovNro(lsMovNro)
            oMov.InsertaMovCta lnMovNro, 1, Replace(lsCtaContDepositoD, "M", oRSDetCheque!nMoneda), Abs(oRSDetCheque!nMonto)
            oMov.InsertaMovObj lnMovNro, 1, 1, ObjEntidadesFinancieras
            oMov.InsertaMovObjIF lnMovNro, 1, 1, lsPersCod, lsIFTpo, lsCtaIfCod
            oMov.InsertaMovCta lnMovNro, 2, Replace(lsCtaContDepositoH, "M", oRSDetCheque!nMoneda), Abs(oRSDetCheque!nMonto) * -1
            oMov.InsertaMovRef lnMovNro, lnMovNroRef
            oMov.InsertaMovDoc lnMovNro, oRSDetCheque!nTpoDoc, oRSDetCheque!cNroDoc, Format(pdFecha, "yyyymmdd hh:mm:ss")
            oMov.InsertaMovDocRec lnMovNro, oRSDetCheque!nTpoDoc, oRSDetCheque!cNroDoc, oRSDetCheque!cPersCod, oRSDetCheque!cIFTpo, oRSDetCheque!cIFCta
            oMov.ActualizaNroMovDeposito lnMovNroRef, lnMovNro
            If oRSDetCheque!nMoneda = gMonedaExtranjera Then oMov.GeneraMovME lnMovNro, lsMovNro
            oMov.ActualizaSaldoMovimiento lsMovNro, "+"
            oRSDetCheque.MoveNext
        Loop
        oRSDetCheque.MoveFirst
        oMov.RegistraDepositoCheque oRSDetCheque!nid, pdFecha, lsIFTpo, lsPersCod, lsCtaIfCod
        Sleep 1000
        lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, "UCCE")
        oMov.CCE_ValorizaChequeNeg oRSDetCheque!nid, CDate(pdFecha & " " & Format(Time, "hh:mm:ss")), lsMovNro
        lsIDs = lsIDs & CStr(oRSDetCheque!nid) & ","
        If Len(lsIDs) > 0 Then lsIDs = Left(lsIDs, Len(lsIDs) - 1)
        oMov.InsertaMov lsMovNro, Replace("42M004", "M", oRSDetCheque!nMoneda), Left("Valorización de Cheques N° " & lsIDs, 300), gMovEstContabMovContable, gMovFlagVigente
    End If
    oMov.CommitTrans
    lbTrans = False
    CCE_Confirma_CHE = True
    Set oMov = Nothing
    Exit Function
ErrConfirmaCHE:
    CCE_Confirma_CHE = False
    If lbTrans Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CCE_Rechaza_CHE(ByVal pdFecha As Date, ByVal psAgeCod As String, ByVal psIFICodDeb As String, _
                                ByVal psCtaCodDeb As String, ByVal psNroCheque As String) As Boolean
Dim oCons As COMDConstSistema.NCOMConstSistema
Set oCons = New COMDConstSistema.NCOMConstSistema
Dim oMov As New COMDMov.DCOMMov
Dim oRSDetCheque As New adodb.Recordset
Dim lbTrans As Boolean
Dim ldFechaRechaza As Date
Dim lsMovNro As String
Dim lnMovNro As Long, lnMovNroRef As Long, lnId As Long
 Dim lsCtaContD As String, lsCtaContH As String

    On Error GoTo ErrRechazaCHE
    oMov.BeginTrans
    lbTrans = True
    lsCtaContD = oCons.LeeConstSistema(558)
    lsCtaContH = oCons.LeeConstSistema(557)
    ldFechaRechaza = CDate(pdFecha & " " & Format(Time, "hh:mm:ss"))
    Set oRSDetCheque = CCE_ObtieneChequeDet(psIFICodDeb, psCtaCodDeb, psNroCheque)
    If Not oRSDetCheque.EOF Then
        oMov.CCE_RechazaChequeNeg oRSDetCheque!nid, ldFechaRechaza
        Do While Not oRSDetCheque.EOF
            lnMovNroRef = oRSDetCheque!nMovNro
            Sleep 1000
            lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, "UCCE")
            oMov.InsertaMov lsMovNro, Replace("42M005", "M", oRSDetCheque!nMoneda), "CCE - Rechazo de Cheque Nro. " & oRSDetCheque!cNroDoc, gMovEstContabMovContable, gMovFlagVigente
            lnMovNro = oMov.GetnMovNro(lsMovNro)
            oMov.InsertaChequeEstado oRSDetCheque!nTpoDoc, oRSDetCheque!cNroDoc, oRSDetCheque!cPersCod, oRSDetCheque!cIFTpo, ldFechaRechaza, gChqEstRechazado, lsMovNro, oRSDetCheque!cIFCta
            oMov.InsertaMovCta lnMovNro, 1, Replace(lsCtaContD, "M", oRSDetCheque!nMoneda), Abs(oRSDetCheque!nMonto)
            oMov.InsertaMovCta lnMovNro, 2, Replace(lsCtaContH, "M", oRSDetCheque!nMoneda), Abs(oRSDetCheque!nMonto) * -1
            oMov.InsertaMovRef lnMovNro, lnMovNroRef
            oMov.InsertaMovDoc lnMovNro, oRSDetCheque!nTpoDoc, oRSDetCheque!cNroDoc, Format(ldFechaRechaza, "yyyymmdd hh:mm:ss")
            oMov.InsertaMovDocRec lnMovNro, oRSDetCheque!nTpoDoc, oRSDetCheque!cNroDoc, oRSDetCheque!cPersCod, oRSDetCheque!cIFTpo, oRSDetCheque!cIFCta
            oMov.CCE_ActualizaNroMovRechazo lnMovNroRef, lnMovNro
            If oRSDetCheque!nMoneda = gMonedaExtranjera Then oMov.GeneraMovME lnMovNro, lsMovNro
            oMov.ActualizaSaldoMovimiento lsMovNro, "+"
            oRSDetCheque.MoveNext
        Loop
    End If
    oMov.CommitTrans
    lbTrans = False
    CCE_Rechaza_CHE = True
    Set oMov = Nothing
    Exit Function
ErrRechazaCHE:
    CCE_Rechaza_CHE = False
    If lbTrans Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CCE_ObtieneBcoDepxCheque(ByVal pnMoneda As Integer) As adodb.Recordset 'Agrego Moneda VAPA20170505
    sSql = "stp_sel_CCE_BcoDepxCheque " & pnMoneda
    Set CCE_ObtieneBcoDepxCheque = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ObtieneChequeDet(ByVal psIFICodDeb As String, ByVal psCtaCodDeb As String, ByVal psNroCheque As String) As adodb.Recordset
    sSql = "stp_sel_CCE_ChequeDet '" & psIFICodDeb & "','" & psCtaCodDeb & "','" & psNroCheque & "'"
    Set CCE_ObtieneChequeDet = oConect.CargaRecordSet(sSql)
End Function
Public Function DatosVoucherTransferencia(ByVal pnMovNro As Long) As adodb.Recordset
    sSql = "stp_sel_CCE_DatosVoucherTransferencia " & pnMovNro
    Set DatosVoucherTransferencia = oConect.CargaRecordSet(sSql)
End Function
Public Function ObtieneDigitoChequeoCheque(ByVal psNro As String) As String
    sSql = "select dbo.Fnc_CCE_ObtieneDigitoChequeoCCI('" & psNro & "') cDig"
    Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        ObtieneDigitoChequeoCheque = rs!cDig
    Else
        ObtieneDigitoChequeoCheque = ""
    End If
End Function
Public Function CCE_EsChequeEnviado(ByVal pnId As Long) As Boolean
    sSql = "stp_sel_CCE_EsChequeEnviado " & pnId
    Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        CCE_EsChequeEnviado = IIf(rs!nvalor = 0, False, True)
    Else
        CCE_EsChequeEnviado = False
    End If
End Function
Public Function CCE_EsEntidadReceptorTransferencia(ByVal psEfinCodBCR As String, ByVal psOptrCod As String) As Boolean
    sSql = "stp_sel_CCE_EsEntidadReceptorTransferencia '" & psEfinCodBCR & "','" & psOptrCod & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        CCE_EsEntidadReceptorTransferencia = IIf(rs!bReceptor = 0, False, True)
    Else
        CCE_EsEntidadReceptorTransferencia = False
    End If
End Function
Public Function CCE_EsEntidadGiradoraCheque(ByVal psEfinCodBCR As String)
    sSql = "stp_sel_CCE_EsEntidadGiradoraCheque '" & psEfinCodBCR & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        CCE_EsEntidadGiradoraCheque = IIf(rs!bGiradora = 0, False, True)
    Else
        CCE_EsEntidadGiradoraCheque = False
    End If
End Function
Public Function CCE_SaldosAFavorPresentado() As adodb.Recordset
    sSql = "stp_sel_CCE_SaldosAFavorPresentado"
    Set CCE_SaldosAFavorPresentado = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_SaldosAFavorAjuste() As adodb.Recordset
    sSql = "stp_sel_CCE_SaldosAFavorAjuste"
    Set CCE_SaldosAFavorAjuste = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_DocRecSaldoAFavor(ByVal pnId As Integer, ByVal psUser As String, ByVal psTipo As String, ByVal psFechaSal As Date) 'VAPA20170323
    sSql = "stp_ins_CCE_DocRecSaldoAFavor " & pnId & ",'" & psUser & "','" & psTipo & "','" & Format(psFechaSal, "YYYYmmdd") & "'"
    oConect.Ejecutar sSql
End Function
'VAPA 20160816 CCE ******************************
Public Function CCE_GeneraTrama_Confirmacion(ByVal psCodInstrumento As String, ByVal psCodArchivo As String, psCodSesion As Integer, ByVal psCodAplicacion As String, ByVal pdFecSis As Date, ByVal psCodUser As String) As adodb.Recordset
    sSql = "STP_SEL_CCE_TRA_GENERA_TRAMA_CONFIRMACION '" & psCodInstrumento & "','" & psCodInstrumento & "','" & psCodSesion & "','" & psCodAplicacion & "','" & Format(pdFecSis, "YYYYmmdd") & "','" & psCodUser & "'"
    Set CCE_GeneraTrama_Confirmacion = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ValidaProCheque(ByVal psCtaDebitar As String, ByVal psIfiCod As String, psNroCheque As String, ByVal psFecha As Date) As adodb.Recordset
    sSql = "stp_sel_CCE_ValidaProCheque '" & psCtaDebitar & "','" & psIfiCod & "','" & psNroCheque & "','" & Format(psFecha, "YYYYmmdd") & "'"
    Set CCE_ValidaProCheque = oConect.CargaRecordSet(sSql)
End Function

Public Function CCE_InsertaLogProCheque(ByVal pdFecha As Date, ByVal psAgeCod As String, ByVal psIFICodDeb As String, ByVal psCtaCodDeb As String, ByVal psNroCheque As String) As Boolean
   sSql = "stp_ins_CCE_LogProChe '" & Format(pdFecha, "YYYYmmdd") & "','" & psAgeCod & "','" & psIFICodDeb & "','" & psCtaCodDeb & "','" & psNroCheque & "'"
   oConect.Ejecutar sSql
End Function
Public Function CCE_ObtieneRefCheque(ByVal psFecha As String, ByVal psNroCheque As String, ByVal psIfiCod As String, ByVal psCtaDebitar As String) As Integer
sSql = "stp_sel_CCE_IdRef '" & Format(psFecha, "YYYYmmdd") & "','" & psNroCheque & "','" & psIfiCod & "','" & psCtaDebitar & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    CCE_ObtieneRefCheque = rs!IdRef
End Function
Public Function CCE_ObtieneAgenciasCmac() As adodb.Recordset
    sSql = "stp_sel_CCE_AgenciasCmacm"
    Set CCE_ObtieneAgenciasCmac = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ObtieneEntidadFinanciera() As adodb.Recordset
    sSql = "stp_sel_CCE_EntidadesFinancieras"
    Set CCE_ObtieneEntidadFinanciera = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ValidaNulo_Envio_Tra(ByVal psCodSesion As String, ByVal psNul As String, psFecha As Date, ByVal psCodAplicacion As String) As adodb.Recordset
    sSql = "STP_SEL_CCE_TRA_VALIDA_NULO '" & psCodSesion & "','" & psNul & "','" & Format(psFecha, "YYYYmmdd") & "','" & psCodAplicacion & "'"
    Set CCE_ValidaNulo_Envio_Tra = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ValidaNulo_Envio_Che(ByVal psCodAplica As String, psFecha As Date) As adodb.Recordset
    sSql = "STP_SEL_CCE_CHE_VALIDA_NULO '" & psCodAplica & "','" & Format(psFecha, "YYYYmmdd") & "'"
    Set CCE_ValidaNulo_Envio_Che = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_GeneraTrama_Nulo_Che(ByVal psCodSesion As String, psFecha As Date, ByVal psCodAplicacion As String, ByVal psRecibido As String, ByVal psCodUser As String) As adodb.Recordset
    sSql = "STP_SEL_CCE_CHE_GENERA_ARCHIVO_NULO '" & psCodSesion & "','" & Format(psFecha, "YYYYmmdd") & "','" & psCodAplicacion & "','" & psRecibido & "','" & psCodUser & "'"
    Set CCE_GeneraTrama_Nulo_Che = oConect.CargaRecordSet(sSql)
End Function

Public Function CCE_GeneraTrama_Nulo_Tra(ByVal psCodSesion As String, psFecha As Date, ByVal psCodAplicacion As String, ByVal psTra As String, ByVal psRecibido As String, ByVal psCodUser As String) As adodb.Recordset
    sSql = "STP_SEL_CCE_TRA_GENERA_ARCHIVO_NULO '" & psCodSesion & "','" & Format(psFecha, "YYYYmmdd") & "','" & psCodAplicacion & "','" & psTra & "','" & psRecibido & "','" & psCodUser & "'"
    Set CCE_GeneraTrama_Nulo_Tra = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ObtieneSesionChe(ByVal psnIdInstrumento As Integer) As adodb.Recordset
    sSql = "stp_sel_CCE_ObtieneSesionChe'" & psnIdInstrumento & "'"
    Set CCE_ObtieneSesionChe = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_CondicionTra(ByVal psID As Long, ByVal psCondicion As Integer)
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
    oConect.AbreConexion 'VAPA20170323
    sSql = "STP_INS_CCE_TRA_CONDICION " & psID & "," & psCondicion & ""
    oConect.Ejecutar sSql
    oConect.CierraConexion 'VAPA20170323
    Set oConect = Nothing
End Function
Public Function CCE_SeleccionaChequesConfirmados(ByVal psFecha As Date) As adodb.Recordset
    sSql = "stp_sel_CCE_ValidaCHeConf'" & Format(psFecha, "YYYYmmdd") & "'"
    Set CCE_SeleccionaChequesConfirmados = oConect.CargaRecordSet(sSql)
End Function

Public Function CCE_SeleccionaPresentados(ByVal psFecha As Date) As adodb.Recordset
    sSql = "stp_sel_CCE_ChePre'" & Format(psFecha, "YYYYmmdd") & "'"
    Set CCE_SeleccionaPresentados = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_SeleccionaRechazados(ByVal psFecha As Date) As adodb.Recordset
    sSql = "stp_sel_CCE_CheRec'" & Format(psFecha, "YYYYmmdd") & "'"
    Set CCE_SeleccionaRechazados = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_AplicaTra(ByVal psCodAplicacion As String, ByVal psFecha As Date) As adodb.Recordset
    sSql = "stp_sel_CCE_AplicaTra '" & psCodAplicacion & "','" & Format(psFecha, "YYYYmmdd") & "'"
    Set CCE_AplicaTra = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_GeneraTrama_Devolucion(ByVal psCodInstrumento As String, ByVal psCodArchivo As String, psCodSesion As Integer, ByVal psCodAplicacion As String, ByVal pdFecSis As Date, ByVal psCodUser As String) As adodb.Recordset
    sSql = "STP_SEL_CCE_TRA_GENERA_TRAMA_DEVOLUCION '" & psCodInstrumento & "','" & psCodInstrumento & "','" & psCodSesion & "','" & psCodAplicacion & "','" & Format(pdFecSis, "YYYYmmdd") & "','" & psCodUser & "'"
    Set CCE_GeneraTrama_Devolucion = oConect.CargaRecordSet(sSql)
End Function

Public Function CCE_GeneraTrama_CHEPRE(ByVal psCodInstrumento As String, ByVal psCodArchivo As String, psCodSesion As Integer, ByVal psCodAplicacion As String, ByVal pdFecSis As Date, ByVal psCodUser As String) As adodb.Recordset
    sSql = "STP_SEL_CCE_CHE_GENERA_TRAMA_ENVIO '" & psCodInstrumento & "','" & psCodInstrumento & "','" & psCodSesion & "','" & psCodAplicacion & "','" & Format(pdFecSis, "YYYYmmdd") & "','" & psCodUser & "'"
    Set CCE_GeneraTrama_CHEPRE = oConect.CargaRecordSet(sSql)
End Function

Public Function CCE_HorarioCheque() As adodb.Recordset
    sSql = "stp_sel_CCE_HorarioChe"
    Set CCE_HorarioCheque = oConect.CargaRecordSet(sSql)
End Function

Public Function CCE_HorarioTransferencia() As adodb.Recordset
    sSql = "stp_sel_CCE_HorarioTra"
    Set CCE_HorarioTransferencia = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ObtieneInstrumento() As adodb.Recordset
    sSql = "stp_sel_CCE_ObtieneInstrumento"
    Set CCE_ObtieneInstrumento = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ValidaPresentadosRec(ByVal psFecha As String, ByVal psCodAplicacion As String) As adodb.Recordset
    sSql = "STP_INS_CCE_TRA_VALIDACION '" & psFecha & "','" & psCodAplicacion & "'"
    'Set CCE_ObtieneInstrumento = oConect.CargaRecordSet(ssql)
    oConect.Ejecutar sSql
End Function
Public Function CCE_ObtieneSesionxIdInstrumento(ByVal psnIdInstrumento As Integer) As adodb.Recordset
    sSql = "stp_sel_CCE_ObtieneSesionxIdInstrumento'" & psnIdInstrumento & "'"
    Set CCE_ObtieneSesionxIdInstrumento = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ObtieneAplicacionTransferencia(ByVal psSesion As Integer) As adodb.Recordset
    sSql = "stp_sel_CCE_ObtieneAplicacionTransferencia " & psSesion
    Set CCE_ObtieneAplicacionTransferencia = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ObtieneAplicacionSesionCheques(ByVal psnIdInstrumento As Integer) As adodb.Recordset
    sSql = "stp_sel_CCE_ObtieneAplicacionSesionCheques " & psnIdInstrumento
    Set CCE_ObtieneAplicacionSesionCheques = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ConfirmaTRA_Enviada(ByVal pnMovNro As Long, ByVal pdFecha As Date) As Long
    sSql = "stp_ins_CCE_ConfirmaTRA_Enviada " & pnMovNro & ",'" & Format(pdFecha, "YYYYmmdd") & "'"
    Set rs = oConect.CargaRecordSet(sSql)
'    If Not (rs.EOF And rs.BOF) Then
    CCE_ConfirmaTRA_Enviada = rs!nMovNro
'    Else
'    CCE_ConfirmaTRA_Enviada = 0
'    End If
End Function
Public Function CCE_RechazaTRA_Enviada(ByVal pnMovNro As Long, ByVal pdFecha As Date)
    sSql = "stp_ins_CCE_RechazaTRA_Enviada " & pnMovNro & ",'" & Format(pdFecha, "YYYYmmdd") & "'"
    oConect.Ejecutar sSql
End Function
Public Function CCE_ObtieneArchivosBlanco() As adodb.Recordset
    sSql = "spt_sel_CCE_ObtieneArchivosBlanco"
    Set CCE_ObtieneArchivosBlanco = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ConstArchivos() As adodb.Recordset
    sSql = "stp_sel_CCE_ConstArchivos"
    Set CCE_ConstArchivos = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ObtieneArchivosRecibidos() As adodb.Recordset
    sSql = "spt_sel_CCE_ObtieneArchivosRecibidos"
    Set CCE_ObtieneArchivosRecibidos = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ValidaCabeceraNulas(ByVal psCodArchivo As String, ByVal psFecha As String, ByVal psCodMoneda As String, ByVal psCodAplicacion As String, ByVal psEstado As Integer) As String
sSql = "stp_sel_CCE_ValidaCabeceraNulas '" & psCodArchivo & "','" & psFecha & "','" & psCodMoneda & "','" & psCodAplicacion & "','" & psEstado & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    CCE_ValidaCabeceraNulas = rs!cVal
End Function
Public Function CCE_ValidaCabecera(ByVal psDato As String) As String
sSql = "stp_sel_CCE_ValidaCabecera '" & psDato & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    CCE_ValidaCabecera = rs!cVal
End Function
Public Function CCE_InsertaIntercambioTransferenciaDet(ByVal psIdArchivo As Long, ByVal pnIdRef As Long, ByVal psTipoRegistro As String, ByVal pscDato As String, ByVal psTipTra As String, ByVal psnEstado As Integer)
    sSql = "STP_INS_CCE_IntercambioTransferenciaRecDet " & psIdArchivo & "," & pnIdRef & ",'" & psTipoRegistro & "','" & pscDato & "','" & psTipTra & "','" & psnEstado & "'"
    oConect.Ejecutar sSql
End Function
Public Function CCE_InsIntercambioTransferenciaRec(ByVal psCodArchivo As String, ByVal psFechaArchivo As String, ByVal psMoneda As String, ByVal psCodAplicacion As String, ByVal psNumArchivo As Integer, ByVal psNumLote As Integer, ByVal psnRegistros As String, ByVal psRecibido As String, ByVal psUsuario As String, ByVal psnEstado As Integer) As Integer
    sSql = "STP_INS_CCE_IntercambioTransferenciaRec '" & psCodArchivo & " ','" & psFechaArchivo & "','" & psMoneda & "','" & psCodAplicacion & "','" & psNumArchivo & "','" & psNumLote & "','" & psnRegistros & "','" & psRecibido & "', '" & psUsuario & "','" & psnEstado & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    CCE_InsIntercambioTransferenciaRec = rs!nid
End Function
Public Function CCE_EliminaOficinas()
sSql = "stp_del_CCE_DeleteOficinas"
oConect.Ejecutar sSql
End Function
Public Function CCE_InsertaOficinas(ByVal psEfinCodBCR As String, _
                                    ByVal psOficCod As String, _
                                    ByVal psOficCodMismaPlaza As String, _
                                    ByVal psOficUbigPlazaComp As String, _
                                    ByVal psOficUbigPlazaCompTr As String, _
                                    ByVal psOficNombre As String, _
                                    ByVal psOficDomicilio As String, _
                                    ByVal psOficLocalidad As String, _
                                    ByVal psOficCodPostalDist As String, _
                                    ByVal psOficNombrePlaza As String, _
                                    ByVal psOficTipo As String, _
                                    ByVal psOficUbigINEI As String, _
                                    ByVal psOficPlazaExc As String, _
                                    ByVal psOficMotivoActualiza As String, _
                                    ByVal psFechaActualiza As String)
                                    
    sSql = "stp_ins_CCE_Oficinas '" & psEfinCodBCR & "','" _
                                    & psOficCod & "','" _
                                    & psOficCodMismaPlaza & "','" _
                                    & psOficUbigPlazaComp & "','" _
                                    & psOficUbigPlazaCompTr & "','" _
                                    & psOficNombre & "','" _
                                    & psOficDomicilio & "','" _
                                    & psOficLocalidad & "','" _
                                    & psOficCodPostalDist & "','" _
                                    & psOficNombrePlaza & "','" _
                                    & psOficTipo & "','" _
                                    & psOficUbigINEI & "','" _
                                    & psOficPlazaExc & "','" _
                                    & psOficMotivoActualiza & "','" _
                                    & psFechaActualiza & "'"
                                    
    Set rs = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_EsSabado(ByVal pdFecha As Date) As Boolean
    sSql = "stp_sel_CCE_EsSabado '" & Format(pdFecha, "YYYYmmdd") & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    CCE_EsSabado = IIf(rs!cValor = 0, False, True)
End Function
Public Function CCE_ObtieneFecPre(ByVal pdFecha As Date, ByVal pdCodAplicacion As String, ByVal psCodBCR As String, ByVal psCodOfic)
    sSql = "stp_sel_CCE_ObtienePresentacion '" & Format(pdFecha, "YYYYmmdd") & "','" & pdCodAplicacion & "','" & psCodBCR & "','" & psCodOfic & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    CCE_ObtieneFecPre = rs!dFecha
End Function
Public Function CCE_ObtieneAplicacion() As String
    sSql = "stp_sel_CCECodAplicacion"
    Set rs = oConect.CargaRecordSet(sSql)
    CCE_ObtieneAplicacion = rs!cCodAplicacion
End Function
Public Function CCE_ValidaAplicacion(ByVal psCodAplicacion As String, ByVal psCodBCR As String) As String
    Dim rs As adodb.Recordset
    sSql = "stp_sel_CCEValidaCodAplicacion '" & psCodAplicacion & "','" & psCodBCR & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    CCE_ValidaAplicacion = rs!cCodAplicacion
End Function
Public Function CCE_SeEnvioTrama(ByVal psCodAplicacion As String, ByVal pdFecha As Date) As Boolean
    Dim rs As adodb.Recordset
    sSql = "stp_sel_CCESeEnvioTrama '" & psCodAplicacion & "','" & Format(pdFecha, "YYYYmmdd") & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    CCE_SeEnvioTrama = IIf(rs!cVal = 0, False, True)
End Function
Public Function CCE_ObtieneAplicacionReporte() As adodb.Recordset
    sSql = "stp_sel_CCE_ObtieneAplicacionRep"
    Set CCE_ObtieneAplicacionReporte = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ReporteSaldos(ByVal psSesion As String, ByVal pdFecha As Date) As adodb.Recordset
    sSql = "stp_sel_CCE_ReporteSaldos '" & psSesion & "','" & Format(pdFecha, "YYYYmmdd") & "'"
    Set CCE_ReporteSaldos = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ExtornoTrama(ByVal psSesion As String, ByVal pdFecha As Date) As adodb.Recordset
    sSql = "stp_sel_CCE_ExtornoTrama '" & psSesion & "','" & Format(pdFecha, "YYYYmmdd") & "'"
    Set CCE_ExtornoTrama = oConect.CargaRecordSet(sSql)
End Function
Public Function CCE_ExtornaTramaPre(ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String, ByVal psOpeCod As String, psGlosa As String, ByVal pnTranNro As Long)
Dim clsCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lsMovNro As String
On Error GoTo ErrExtornaTransf
    Set clsCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
    clsCapMov.dBeginTrans
        lbTrans = True
        lsMovNro = clsCapMov.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
        clsCapMov.AgregaMov lsMovNro, psOpeCod, psGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
        lnMovNro = clsCapMov.GetnMovNro(lsMovNro)
        clsCapMov.CCE_ExtornaTramaPesentados pnTranNro
   clsCapMov.dCommitTrans
   lbTrans = False
   Set clsCapMov = Nothing
Exit Function
ErrExtornaTransf:
    If lbTrans Then
        clsCapMov.dRollbackTrans
        Set clsCapMov = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function DatosVoucherTransferenciaExtorno(ByVal pnMovNro As Long) As adodb.Recordset
    sSql = "stp_sel_CCE_DatosVoucherTransferenciaExtorno " & pnMovNro
    Set DatosVoucherTransferenciaExtorno = oConect.CargaRecordSet(sSql)
End Function
'end *******************


