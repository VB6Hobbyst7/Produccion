VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMSeguros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public oImp As COMFunciones.FCOMVarImpresion 'JUEZ 20140515
'*** FRHU 20140528 ***************************************************************************************
'REMESA: CERTIFICADOS POR AGENCIAS
Public Function ObtenerSegTarjetaRemesaAgencia() As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ObtenerSegTarjetaRemesaAgencia = oDCOMSeg.ObtenerSegTarjetaRemesaAgencia()
    Set oDCOMSeg = Nothing
End Function
Public Sub InsertaSegTarjetaRemesaAgencia(ByVal psAgeCod As String, ByVal psRangoDel As String, ByVal psRangoAl As String, ByVal pdFecha As Date, ByVal pnEstado As Integer)
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
        oDCOMSeg.InsertaSegTarjetaRemesaAgencia psAgeCod, psRangoDel, psRangoAl, pdFecha, pnEstado
    Set oDCOMSeg = Nothing
End Sub
Public Function ValidarSegTarjetaRemesaAgencia(ByVal pnId As Integer, ByVal psRangoDel As String, ByVal psRangoAl As String, ByVal pnOperacion As Integer) As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ValidarSegTarjetaRemesaAgencia = oDCOMSeg.ValidarSegTarjetaRemesaAgencia(pnId, psRangoDel, psRangoAl, pnOperacion)
    Set oDCOMSeg = Nothing
End Function
Public Sub ActualizarSegTarjetaRemesaAgencia(ByVal pnId As Integer, ByVal psRangoDel As String, ByVal psRangoAl As String, ByVal pdFecha As Date, ByVal pnEstado As Integer, ByVal pnOperacion As Integer)
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
        oDCOMSeg.ActualizarSegTarjetaRemesaAgencia pnId, psRangoDel, psRangoAl, pdFecha, pnEstado, pnOperacion
    Set oDCOMSeg = Nothing
End Sub
'FIN REMESA: CERTIFICADOS POR AGENCIAS
'*** FIN FRHU 20140528 *****************************************************************
'JUEZ 20140515 **********************************************************************
Public Sub IniciaImpresora(Optional ByVal nImpresora As COMFunciones.Impresoras = gEPSON)
Set oImp = New COMFunciones.FCOMVarImpresion
oImp.Inicia nImpresora
End Sub
Private Sub Class_Initialize()
IniciaImpresora
End Sub

Private Sub Class_Terminate()
Set oImp = Nothing
End Sub
Public Function VerificaOperacionSegTarjeta(ByVal psOpeCod As String) As Boolean
Dim oGen As COMDConstSistema.DCOMGeneral
Dim sOpeCod As String
Dim sCod() As String
Dim i As Integer
    
    Set oGen = New COMDConstSistema.DCOMGeneral
    sOpeCod = oGen.LeeConstSistema(470)
    
    sCod = Split(sOpeCod, ",")
    VerificaOperacionSegTarjeta = False
    
    For i = 0 To UBound(sCod)
        If psOpeCod = sCod(i) Then
            VerificaOperacionSegTarjeta = True
            Exit Function
        End If
    Next i
End Function
Public Sub InsertaNroSolicitudAfiliacionMes(ByVal psNumTarj As String, ByVal pdFecSis As String)
Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    oDCOMSeg.InsertaNroSolicitudAfiliacionMes psNumTarj, pdFecSis
Set oDCOMSeg = Nothing
End Sub
Public Sub InsertaSegTarjetaAfiliacion(ByVal psNumCert As String, ByVal psNumTarj As String, ByVal psCtaCod As String, ByVal pdFecSis As String, ByVal psMovNro As String, ByVal pnMovNro As Long, ByVal psPersCod As String)
Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    oDCOMSeg.InsertaSegTarjetaAfiliacion psNumCert, psNumTarj, psCtaCod, pdFecSis, psMovNro, pnMovNro, psPersCod
Set oDCOMSeg = Nothing
End Sub
Public Function ImprimeBoletaAfilicacionSeguroTarjeta(ByVal pnMovNro As Long, ByVal psMovNro As String, ByVal psNomAge As String, ByVal pbImpTMU As Boolean) As String
Dim lscad As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMax As Integer
Dim sFecha As String
Dim sHora As String
Dim sRegistroSBS As String
Dim cNumPoliza As String
Dim sImporte As String
Dim oNSeg As COMDCaptaGenerales.DCOMSeguros
Dim rs As ADODB.Recordset

'Set oImp = New COMFunciones.FCOMVarImpresion
    
    lscad = ""
    lscad = lscad & oImp.gPrnInicializa
    lscad = lscad & oImp.gPrnEspaLineaN   'espaciamiento lineas 1/6 pulg.
    lscad = lscad & oImp.gPrnTamPagina22  'Longitud de página a 22 líneas'
    lscad = lscad & oImp.gPrnTamLetra10CPI   'Tamaño 10 cpi
    lscad = lscad & oImp.gPrnTpoLetraRoman     'Tipo de Letra Sans Serif
    lscad = lscad & oImp.gPrnCondensadaOFF ' cancela condensada
    lscad = lscad & oImp.gPrnNegritaOFF ' desactiva negrita

    sSep = 14
    sIni = 0
    sMax = 39

    sFecha = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)
    sHora = Mid(psMovNro, 9, 2) & ":" + Mid(psMovNro, 11, 2) & ":" + Mid(psMovNro, 13, 2)
    
    Set oNSeg = New COMDCaptaGenerales.DCOMSeguros
    Set rs = oNSeg.RecuperaSegTarjetaAfiliacion(pnMovNro)
    Set oNSeg = Nothing
    'sImporte = "S/. " & CStr(Format(rs!nMontoPrima, "#,##0.00"))
    sImporte = IIf(Mid(rs!cCtaCodDebito, 9, 1) = "1", "S/ ", "US$ ") & CStr(Format(rs!nMontoPrima, "#,##0.00")) 'JUEZ 20150510
    
If Not pbImpTMU Then
    lscad = lscad & oImp.gPrnNegritaON  'Activa Negrita
    lscad = lscad & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "CMAC MAYNAS" & Space(sMax - Len("CMAC MAYNAS") + sSep) & "CMAC MAYNAS" & Space(sMax - Len("CMAC MAYNAS")) & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & psNomAge & Space(31 - Len(psNomAge)) & pnMovNro & Space(sSep) & psNomAge & Space(31 - Len(psNomAge)) & pnMovNro & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(sSep) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImp.gPrnSaltoLinea
    'lscad = lscad & Space(sIni) & "--SEGURO PROTECCION TARJETA DE DEBITO--" & Space(sSep) & "--SEGURO PROTECCION TARJETA DE DEBITO--" & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "--SEGURO PROTECCION DE TARJETA DEBITO--" & Space(sSep) & "--SEGURO PROTECCION DE TARJETA DEBITO--" & oImp.gPrnSaltoLinea 'APRI20171027 ERS028-2017
    lscad = lscad & oImp.gPrnNegritaOFF 'Desactiva Negrita
    lscad = lscad & Space(sIni) & "REGISTRO SBS    " & rs!cRegistroSBS & Space(sMax - Len("REGISTRO SBS    " & rs!cRegistroSBS) + sSep) & "REGISTRO SBS    " & rs!cRegistroSBS & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "POLIZA N*       " & rs!cNumPoliza & Space(sMax - Len("POLIZA Nº       " & rs!cNumPoliza) + sSep) & "POLIZA N*       " & rs!cNumPoliza & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "CERTIFICADO N*  " & rs!cNumCertificado & Space(sMax - Len("CERTIFICADO Nº  " & rs!cNumCertificado) + sSep) & "CERTIFICADO N*  " & rs!cNumCertificado & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "FECHA INICIO    " & sFecha & Space(sMax - Len("FECHA INICIO    " & sFecha) + sSep) & "FECHA INICIO    " & sFecha & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "FORMA DE PAGO:  " & IIf(rs!bPrimaAnual = 0, "MENSUAL", "ANUAL") & Space(sMax - Len("FORMA DE PAGO:  " & IIf(rs!bPrimaAnual = 0, "MENSUAL", "ANUAL")) + sSep) & "FORMA DE PAGO:  " & IIf(rs!bPrimaAnual = 0, "MENSUAL", "ANUAL") & oImp.gPrnSaltoLinea 'APRI20181023 ERS071-2018
    lscad = lscad & Space(sIni) & Left(rs!cPersNombre, 39) & Space(sMax - Len(Left(rs!cPersNombre, 39)) + sSep) & Left(rs!cPersNombre, 39) & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "DNI/CE/RUC      " & rs!cPersIDnro & Space(sMax - Len("DNI/CE/RUC      " & rs!cPersIDnro) + sSep) & "DNI/CE/RUC      " & rs!cPersIDnro & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "TARJ.TITULAR    " & rs!cNumTarjeta & Space(sMax - Len("TARJ.TITULAR    " & rs!cNumTarjeta) + sSep) & "TARJ.TITULAR    " & rs!cNumTarjeta & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "NRO. CUENTA     " & rs!cCtaCodDebito & Space(sMax - Len("NRO. CUENTA     " & rs!cCtaCodDebito) + sSep) & "NRO. CUENTA     " & rs!cCtaCodDebito & oImp.gPrnSaltoLinea
    'CTI2 FERIMORO ERS034-2019   :    31082020
    lscad = lscad & Space(sIni) & "Usuario Venta   " & rs!cUserVendedor & Space(sMax - Len("Usuario Venta   " & rs!cUserVendedor) + sSep) & "Usuario Venta   " & rs!cUserVendedor & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    '-------------------------------------------------------
    lscad = lscad & oImp.gPrnNegritaON  'Activa Negrita
    lscad = lscad & Space(sIni) & String(sMax, "-") & Space(sSep) & String(sMax, "-") & oImp.gPrnSaltoLinea
    'lscad = lscad & Space(sIni) & "Importe Cuota" & Space(sMax - Len("Importe Cuota") - Len(sImporte)) & sImporte & Space(sSep) & "Importe Cuota" & Space(sMax - Len("Importe Cuota") - Len(sImporte)) & sImporte & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "Prima" & Space(sMax - Len("Importe Cuota") - Len(sImporte)) & sImporte & Space(sSep + 8) & "Prima" & Space(sMax - Len("Importe Cuota") - Len(sImporte)) & sImporte & oImp.gPrnSaltoLinea 'APRI20171027 ERS028-2017
    lscad = lscad & Space(sIni) & String(sMax, "-") & Space(sSep) & String(sMax, "-") & oImp.gPrnSaltoLinea
    lscad = lscad & oImp.gPrnNegritaOFF 'Desactiva Negrita
'    lscad = lscad & Space(sIni) & "*** Con el  uso  de  mi  clave  secreta" & Space(sSep) & "*** Con el  uso  de  mi  clave  secreta" & oImp.gPrnSaltoLinea
'    lscad = lscad & Space(sIni) & "autorizo la contratacion del seguro de " & Space(sSep) & "autorizo la contratacion del seguro de " & oImp.gPrnSaltoLinea
'    lscad = lscad & Space(sIni) & "tarjeta de debito ***" & Space(18) & Space(sSep) & "tarjeta de debito ***" & oImp.gPrnSaltoLinea
    'APRI20171027 ERS028-2017
    lscad = lscad & Space(sIni) & "*** Dejo constancia de haber tomado " & Space(sSep + 1) & "*** Dejo constancia de haber tomado " & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "conocimiento de las condiciones generales," & Space(sSep - 3) & "conocimiento de las condiciones generales," & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "particulares y cláusulas adicionales" & Space(sSep + 3) & "particulares y cláusulas adicionales" & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "de la presente póliza***" & Space(sSep + 15) & "de la presente póliza***" & oImp.gPrnSaltoLinea
    lscad = lscad & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "Firma y DOI del Asegurado" & Space(sSep + 14) & "Firma y DOI del Asegurado" & oImp.gPrnSaltoLinea
    lscad = lscad & oImp.gPrnSaltoLinea
    lscad = lscad & oImp.gPrnSaltoLinea
    lscad = lscad & oImp.gPrnSaltoLinea
    'END APRI20171027
    lscad = lscad & Space(sIni) & Right(psMovNro, 4) & Space(sMax - Len(Right(psMovNro, 4)) + sSep) & Right(psMovNro, 4) & oImp.gPrnSaltoLinea
Else
    lscad = lscad & oImp.gPrnNegritaON  'Activa Negrita
    lscad = lscad & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "CMAC MAYNAS" & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & psNomAge & Space(31 - Len(psNomAge)) & pnMovNro & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImp.gPrnSaltoLinea
    'lscad = lscad & Space(sIni) & "--SEGURO PROTECCION TARJETA DE DEBITO--" & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "--SEGURO PROTECCION DE TARJETA DEBITO--" & oImp.gPrnSaltoLinea 'APRI20171027 ERS028-2017
    'lscad = lscad & oImp.gPrnNegritaOFF 'Desactiva Negrita
    lscad = lscad & Space(sIni) & "REGISTRO SBS    " & rs!cRegistroSBS & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "POLIZA N*       " & rs!cNumPoliza & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "CERTIFICADO N*  " & rs!cNumCertificado & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "FECHA INICIO    " & sFecha & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "FORMA DE PAGO:  " & IIf(rs!bPrimaAnual = 0, "MENSUAL", "ANUAL") & oImp.gPrnSaltoLinea  'APRI20181023 ERS071-2018
    lscad = lscad & Space(sIni) & Left(rs!cPersNombre, 39) & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "DNI/CE/RUC      " & rs!cPersIDnro & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "TARJ.TITULAR    " & rs!cNumTarjeta & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "NRO. CUENTA     " & rs!cCtaCodDebito & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    'lscad = lscad & oImp.gPrnNegritaON  'Activa Negrita
    'lscad = lscad & Space(sIni) & "Importe Cuota" & Space(sMax - Len("Importe Cuota") - Len(sImporte)) & sImporte & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "Prima" & Space(sMax - Len("Importe Cuota") - Len(sImporte)) & sImporte & oImp.gPrnSaltoLinea 'APRI20171027 ERS028-2017
    lscad = lscad & Space(sIni) & String(sMax, "-") & oImp.gPrnSaltoLinea
    'lscad = lscad & oImp.gPrnNegritaOFF 'Desactiva Negrita
'    lscad = lscad & Space(sIni) & "*** Con el  uso  de  mi  clave  secreta" & oImp.gPrnSaltoLinea
'    lscad = lscad & Space(sIni) & "autorizo la contratacion del seguro de " & oImp.gPrnSaltoLinea
'    lscad = lscad & Space(sIni) & "tarjeta de debito ***" & Space(18) & oImp.gPrnSaltoLinea
    'APRI20171027 ERS028-2017
    lscad = lscad & Space(sIni) & "*** Dejo constancia de haber tomado " & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "conocimiento de las condiciones" & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "generales, particulares y cláusulas" & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "adicionales de la presente póliza***" & oImp.gPrnSaltoLinea
    lscad = lscad & oImp.gPrnSaltoLinea
    lscad = lscad & Space(sIni) & "Firma y DOI del Asegurado" & oImp.gPrnSaltoLinea
    lscad = lscad & oImp.gPrnSaltoLinea
    lscad = lscad & oImp.gPrnSaltoLinea
    lscad = lscad & oImp.gPrnSaltoLinea
    'END APRI20171027
    lscad = lscad & Space(sIni) & Right(psMovNro, 4) & oImp.gPrnSaltoLinea
    lscad = lscad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
    lscad = lscad & oImp.gPrnNegritaOFF 'Desactiva Negrita
End If

ImprimeBoletaAfilicacionSeguroTarjeta = lscad
End Function
Public Sub ActualizaEstadoSegTarjetaAfiliacion(ByVal pcNumCertificado As String, ByVal pnMovNro As Long, ByVal pnSegEstado As Integer)
'APRI20171027 ADD pcNumCertificado ERS028-2017
Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    oDCOMSeg.ActualizaEstadoSegTarjetaAfiliacion pcNumCertificado, pnMovNro, pnSegEstado
Set oDCOMSeg = Nothing
End Sub
'END JUEZ ***************************************************************************
'******** FRHU 20140610 ERS068-2014 *************************************************
'Configuracion de documentos
Public Function ObtenerSegTarjetaConfigDoc(ByVal pnOperacion As Integer) As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ObtenerSegTarjetaConfigDoc = oDCOMSeg.ObtenerSegTarjetaConfigDoc(pnOperacion)
    Set oDCOMSeg = Nothing
End Function
Public Sub InsertaActualizaSegTarjetaConfigDoc(ByVal pnId As Integer, ByVal psDescripcion As String, ByVal pnEstado As Integer)
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
        oDCOMSeg.InsertaActualizaSegTarjetaConfigDoc pnId, psDescripcion, pnEstado
    Set oDCOMSeg = Nothing
End Sub
Public Function QuitarSegTarjetaConfigDoc(ByVal pnId As Integer) As Integer
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Dim rs As ADODB.Recordset
    Set rs = oDCOMSeg.QuitarSegTarjetaConfigDoc(pnId)
    If Not rs.BOF And Not rs.EOF Then
        QuitarSegTarjetaConfigDoc = rs!valor
    End If
    Set oDCOMSeg = Nothing
End Function
'SOLICITUD: Proceso de Activación de Seguro de Tarjeta
Public Function ValidarSegTarjetaGrabarSolicitud(ByVal psNumSolicitud As String, ByVal psNumActSiniestro As String) As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ValidarSegTarjetaGrabarSolicitud = oDCOMSeg.ValidarSegTarjetaGrabarSolicitud(psNumSolicitud, psNumActSiniestro)
    Set oDCOMSeg = Nothing
End Function
Public Sub InsertaSegTarjetaEstado(ByVal psNumSolicitud As String, ByVal pdPrdEstado As Date, ByVal pnPrdEstado As Integer, ByVal pcDescripcion As String, ByVal pnMonto As Currency, ByVal psUltimaActualizacion As String)
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
        oDCOMSeg.InsertaSegTarjetaEstado psNumSolicitud, pdPrdEstado, pnPrdEstado, pcDescripcion, pnMonto, psUltimaActualizacion
    Set oDCOMSeg = Nothing
End Sub
Public Function ValidarTarjetasAfiliadasSegTarjetaXPersona(ByVal psCodPersona As String) As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ValidarTarjetasAfiliadasSegTarjetaXPersona = oDCOMSeg.ValidarTarjetasAfiliadasSegTarjetaXPersona(psCodPersona)
    Set oDCOMSeg = Nothing
End Function
Public Function ValidarSolicitudesEnProcesoSegTarjeta(ByVal psCodPersona As String) As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ValidarSolicitudesEnProcesoSegTarjeta = oDCOMSeg.ValidarSolicitudesEnProcesoSegTarjeta(psCodPersona)
    Set oDCOMSeg = Nothing
End Function
Public Sub GrabarNuevaSolicitud(ByVal psNumSolicitud As String, ByVal pdFechaSolicitud As Date, ByVal pnMontoSolicitud As Currency, ByVal pnPrdEstado As Integer, _
                                ByVal cPersCodAsegurado As String, ByVal cPersCodTercero As String, ByVal pcRelacionTercero As String, _
                                ByVal psNumActSiniestro As String, ByVal psAgeCodSiniestro As String, ByVal psDescripcionSiniestro As String, ByVal psUbigeoSiniestro As String, _
                                ByVal psDireccionRefSiniestro As String, ByVal pdFechaSiniestro As Date, ByVal psHoraSiniestro As String, ByVal psUltimaActualizacion As String, _
                                ByVal pMatSiniestro As Variant, ByVal pMatDocumento As Variant, pMatTipoDocumento As Variant, ByVal pdFecSis As Date)
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Dim x As Integer, fila As Integer
    On Error GoTo ErrorGrabarNuevaSolicitud
        oDCOMSeg.dBeginTrans
        Call oDCOMSeg.InsertaSegTarjetaSolicitud(psNumSolicitud, pdFechaSolicitud, pnMontoSolicitud, pnPrdEstado, cPersCodAsegurado, cPersCodTercero, pcRelacionTercero, _
                                                 psNumActSiniestro, psAgeCodSiniestro, psDescripcionSiniestro, psUbigeoSiniestro, psDireccionRefSiniestro, pdFechaSiniestro, psHoraSiniestro, psUltimaActualizacion)
        For x = LBound(pMatSiniestro) To UBound(pMatSiniestro)
            If pMatSiniestro(x) <> 0 Then Call oDCOMSeg.InsertaSegTarjetaSolicitudSiniestro(psNumSolicitud, pMatSiniestro(x))
        Next
        For fila = LBound(pMatDocumento) To UBound(pMatDocumento)
            If pMatDocumento(fila) <> 0 Then Call oDCOMSeg.InsertaSegTarjetaSolicitudDocumento(psNumSolicitud, pMatDocumento(fila), pMatTipoDocumento(fila))
        Next
        Call oDCOMSeg.InsertaSegTarjetaEstado(psNumSolicitud, pdFecSis, pnPrdEstado, "Solicitud de Activación de Seguro de Tarjeta", pnMontoSolicitud, psUltimaActualizacion)
        oDCOMSeg.dCommitTrans
        Set oDCOMSeg = Nothing
    Exit Sub
    
ErrorGrabarNuevaSolicitud:
    oDCOMSeg.dRollbackTrans
    Err.Raise Err.Number, "Error En Proceso de Grabar Nueva Solicitud Seg. tarjeta", Err.Description
End Sub
'Proceso Rechazo Solicitud
Public Sub RechazarSolicitud(ByVal psNumSolicitud As String, ByVal psMovNro As String, ByVal pnEstado As Integer, ByVal pdFechaResp As Date, _
                             ByVal pdPrdEstado As Date, ByVal pcDescripcion As String, ByVal pnMonto As Currency)
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    On Error GoTo ErrorRechazarSolicitud
        oDCOMSeg.dBeginTrans
        Call oDCOMSeg.ActualizarSegTarjetaSolicitudARechazado(psNumSolicitud, psMovNro, pnEstado, pdFechaResp)
        Call oDCOMSeg.InsertaSegTarjetaEstado(psNumSolicitud, pdPrdEstado, pnEstado, pcDescripcion, pnMonto, psMovNro)
        oDCOMSeg.dCommitTrans
        Set oDCOMSeg = Nothing
    Exit Sub
    
ErrorRechazarSolicitud:
    oDCOMSeg.dRollbackTrans
    Err.Raise Err.Number, "Error En Proceso de Rechazar Solicitud Seg. tarjeta", Err.Description
End Sub
Public Function ObtenerSegTarjetSolictudXNSolicitud(ByVal psNumSolicitud As String) As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ObtenerSegTarjetSolictudXNSolicitud = oDCOMSeg.ObtenerSegTarjetSolictudXNSolicitud(psNumSolicitud)
    Set oDCOMSeg = Nothing
End Function
'Aceptacion de Solicitud
Public Function ObtenerSegTarjetaMovAprobadas() As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ObtenerSegTarjetaMovAprobadas = oDCOMSeg.ObtenerSegTarjetaMovAprobadas()
    Set oDCOMSeg = Nothing
End Function
Public Function ObtenerSegTarjetaSolicitudesAprobadasXMov(ByVal psMovNro As String) As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ObtenerSegTarjetaSolicitudesAprobadasXMov = oDCOMSeg.ObtenerSegTarjetaSolicitudesAprobadasXMov(psMovNro)
    Set oDCOMSeg = Nothing
End Function
Public Sub ActualizarSegTarjetaSolicitudAceptado(ByVal psNumSolicitud As String, ByVal psUltimaActualizacion As String, ByVal pnPrdEstado As Integer, ByVal pnMonto As Currency)
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
        oDCOMSeg.ActualizarSegTarjetaSolicitudAceptado psNumSolicitud, psUltimaActualizacion, pnPrdEstado, pnMonto
    Set oDCOMSeg = Nothing
End Sub
Public Sub InsertarSegTarjetaSolicitudEstadoAceptado(ByVal psNumSolicitud As String, ByVal pdPrdEstado As Date, ByVal pnPrdEstado As Integer, ByVal psDescripcion As String, ByVal pnMonto As Currency, ByVal psUltimaActualizacion As String)
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
        oDCOMSeg.InsertarSegTarjetaSolicitudEstadoAceptado psNumSolicitud, pdPrdEstado, pnPrdEstado, psDescripcion, pnMonto, psUltimaActualizacion
    Set oDCOMSeg = Nothing
End Sub
'FIN FRHU 20140610 ****************************
'FRHU 20140903 ANEXO-03
Public Function ObtenerSegTarjetaMovPendienteXcMovNro(ByVal psMovNro As String) As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ObtenerSegTarjetaMovPendienteXcMovNro = oDCOMSeg.ObtenerSegTarjetaMovPendienteXcMovNro(psMovNro)
    Set oDCOMSeg = Nothing
End Function
'FIN FRHU 20140903
'JUEZ 20150331 ***********************************************
Public Function GrabarAnulacionTarjetaSinDevolucion(ByVal pnMovNroReg As Long, ByVal psMovNro As String) As String
'Dim oDCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento
'Dim oDSeg As COMDCaptaGenerales.DCOMSeguros
'Dim rs As ADODB.Recordset
'Dim nMovNro As Long
'Dim nMontoCom As Double
'Dim bTrans As Boolean
'
'On Error GoTo ErrGrabar
'    GrabarAnulacionTarjetaSinDevolucion = ""
'    bTrans = False
'
'    Set oDSeg = New COMDCaptaGenerales.DCOMSeguros
'    Set rs = oDSeg.RecuperaSegTarjetaAfiliacion(pnMovNroReg)
'    Set oDSeg = Nothing
'    If Not rs.EOF And Not rs.BOF Then
'        nMontoCom = CDbl(rs!nMontoPrima)
'    Else
'        GrabarAnulacionTarjetaSinDevolucion = "No se pudo obtener datos de la afiliación de la tarjeta"
'        Exit Function
'    End If
'
'    Set oDCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
'    oDCapMov.dbCmact.BeginTrans
'    bTrans = True
'    oDCapMov.AgregaMov psMovNro, "300437", "Anulacion Tarjeta Debito sin Devolucion", gMovEstContabMovContable, gMovFlagVigente
'    nMovNro = oDCapMov.GetnMovNro(psMovNro)
'    oDCapMov.AgregaMovOpeVarias nMovNro, nMontoCom, "", "", gMonedaNacional, "300437"
'    oDCapMov.AgregaMovRef nMovNro, pnMovNroReg
'    oDCapMov.dbCmact.CommitTrans
'    bTrans = False
'
'    Set oDSeg = New COMDCaptaGenerales.DCOMSeguros
'    oDSeg.ActualizaEstadoSegTarjetaAfiliacion pnMovNroReg, 504
'    Set oDSeg = Nothing
'Exit Function
'
'ErrGrabar:
'    If bTrans Then oDCapMov.dbCmact.RollbackTrans
'    Set oDCapMov = Nothing
'    GrabarAnulacionTarjetaSinDevolucion = "No se pudo realizar la anulación"
'    Err.Raise Err.Number, "Error en GrabarAnulacionTarjetaSinDevolucion", Err.Description
End Function
'END JUEZ ****************************************************
'JUEZ 20150510 *********************************************************
Public Function GrabarAnulacionTarjetaDevolucionPorCobrar(ByVal pnMovNroReg As Long, ByVal psMovNro As String, ByVal pdFecSis As Date) As String
Dim oDCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim oDSeg As COMDCaptaGenerales.DCOMSeguros
Dim rs As ADODB.Recordset
Dim nMovNro As Long
Dim nMontoCom As Double
Dim sPersCod As String
Dim sNumCert As String
Dim sCtaCodDeb As String
Dim bTrans As Boolean

On Error GoTo ErrGrabar
    GrabarAnulacionTarjetaDevolucionPorCobrar = ""
    bTrans = False
    
    Set oDSeg = New COMDCaptaGenerales.DCOMSeguros
    Set rs = oDSeg.RecuperaSegTarjetaAfiliacion(pnMovNroReg)
    Set oDSeg = Nothing
    If Not rs.EOF And Not rs.BOF Then
        nMontoCom = CDbl(rs!nMontoPrima)
        sPersCod = rs!cPersCod
        sNumCert = rs!cNumCertificado
        sCtaCodDeb = rs!cCtaCodDebito
    Else
        GrabarAnulacionTarjetaDevolucionPorCobrar = "No se pudo obtener datos de la afiliación de la tarjeta"
        Exit Function
    End If
    
    Set oDCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
    oDCapMov.dbCmact.BeginTrans
    bTrans = True
    oDCapMov.AgregaMov psMovNro, "300437", "Anulacion Tarjeta Debito Devolucion Cargo Cuenta por Cobrar", gMovEstContabMovContable, gMovFlagVigente
    nMovNro = oDCapMov.GetnMovNro(psMovNro)
    oDCapMov.AgregaMovOpeVarias nMovNro, nMontoCom, "", "", gMonedaNacional, "300437"
    oDCapMov.AgregaMovRef nMovNro, pnMovNroReg
    oDCapMov.dbCmact.CommitTrans
    bTrans = False
    Set oDCapMov = Nothing
    
    Set oDSeg = New COMDCaptaGenerales.DCOMSeguros
    oDSeg.dBeginTrans
    bTrans = True
    oDSeg.InsertaSegTarjetaAnulaDevPendiente sNumCert, sPersCod, sCtaCodDeb, pdFecSis, nMontoCom
    oDSeg.ActualizaEstadoSegTarjetaAfiliacion sNumCert, pnMovNroReg, 504
    oDSeg.dCommitTrans
    bTrans = False
    Set oDSeg = Nothing
Exit Function

ErrGrabar:
    If bTrans Then oDCapMov.dbCmact.RollbackTrans
    Set oDCapMov = Nothing
    GrabarAnulacionTarjetaDevolucionPorCobrar = "No se pudo realizar la anulación"
    Err.Raise Err.Number, "Error en GrabarAnulacionTarjetaDevolucionPorCobrar", Err.Description
End Function
'END JUEZ **************************************************************
'RECO20151102 ERS073-2015************************************************
Public Function ObtieneParametrosSegSepelio() As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ObtieneParametrosSegSepelio = oDCOMSeg.ObtieneParametrosSegSepelio
    Set oDCOMSeg = Nothing
End Function
Public Function ObtieneCriteriosSegSepelio() As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ObtieneCriteriosSegSepelio = oDCOMSeg.ObtieneCriteriosSegSepelio
    Set oDCOMSeg = Nothing
End Function
Public Sub ActualizaValoresParamSepelio(ByVal pnParamID As Integer, ByVal pnValor As Double)
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Call oDCOMSeg.ActualizaValoresParamSepelio(pnParamID, pnValor)
End Sub
Public Sub ActualizaCriterioSegSepelio(ByVal pnCriterioID As Integer, ByVal pnValAnio As Integer, ByVal pnValMes As Integer, ByVal pnValDia As Integer)
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Call oDCOMSeg.ActualizaCriterioSegSepelio(pnCriterioID, pnValAnio, pnValMes, pnValDia)
    Set oDCOMSeg = Nothing
End Sub
Public Function ObtieneParamNumCertifSepelio()
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ObtieneParamNumCertifSepelio = oDCOMSeg.ObtieneParamNumCertifSepelio
    Set oDCOMSeg = Nothing
End Function
Public Sub ActualizaParamNumCertifSepelio(ByVal psAgeCod As String, ByVal pnDesde As Integer, ByVal pnHasta As Integer, ByVal pnStock As Integer, ByVal psUser As String)
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Call oDCOMSeg.ActualizaParamNumCertifSepelio(psAgeCod, pnDesde, pnHasta, pnStock, psUser)
    Set oDCOMSeg = Nothing
End Sub
Public Sub RegistraAseguradoSepelio(ByVal psNumCertificado As String, ByVal psAseguradoNombre As String, ByVal pnRelacion As Integer, ByVal pnParticipacion As Integer, ByVal pnEstado As Integer, Optional ByVal pConn As COMConecta.DCOMConecta = Nothing)
    'RIRO 20170505 ADD pConn
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Call oDCOMSeg.RegistraAseguradoSepelio(psNumCertificado, psAseguradoNombre, pnRelacion, pnParticipacion, pnEstado, pConn)
    Set oDCOMSeg = Nothing
End Sub
Public Function ObtieneDatosPersSegSepelio(ByVal psPersCod As String) As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ObtieneDatosPersSegSepelio = oDCOMSeg.ObtieneDatosPersSegSepelio(psPersCod)
    Set oDCOMSeg = Nothing
End Function
Public Function ActualizaEstadoSeguroSepelio(ByVal psNumCertificado As String, ByVal psFecAfiliacion As String, ByVal pnMovNroReg As Long, ByVal pnEstado As Integer) As String
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Dim rs As New ADODB.Recordset
    
    Set rs = oDCOMSeg.ActualizaEstadoSeguroSepelio(psNumCertificado, psFecAfiliacion, pnMovNroReg, pnEstado)
    If Not (rs.BOF And rs.EOF) Then
        ActualizaEstadoSeguroSepelio = rs!cNumCertificado
    End If
    Set oDCOMSeg = Nothing
End Function
Public Function ObtieneClienteAfiliacionSepelio(ByVal psFecha As String, ByVal psPersCod As String) As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set ObtieneClienteAfiliacionSepelio = oDCOMSeg.ObtieneClienteAfiliacionSepelio(psFecha, psPersCod)
    Set oDCOMSeg = Nothing
End Function
Public Function DevulveClienteSeguroActivo(ByVal psPersCod As String) As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set DevulveClienteSeguroActivo = oDCOMSeg.DevulveClienteSeguroActivo(psPersCod)
    Set oDCOMSeg = Nothing
End Function
Public Function VerificaAfiliacionDesactivada(ByVal psPersCod As String) As ADODB.Recordset
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Set VerificaAfiliacionDesactivada = oDCOMSeg.VerificaAfiliacionDesactivada(psPersCod)
    Set oDCOMSeg = Nothing
End Function

Public Function ImprimeBoletaAfilicacionSeguroSepelio(ByVal pnMovNro As Long, ByVal psMovNro As String, ByVal psNomAge As String, ByVal pbImpTMU As Boolean, _
                                                      Optional ByVal pConn As COMConecta.DCOMConecta = Nothing, Optional ByVal nProxFechaCronograma As String = "", Optional ByVal prsPagos As ADODB.Recordset = Nothing) As String
    'RIRO 20170612 ADD pConn
    'CTI2 FERIMORO ERS030-2019 :: 25082020 : nProxFechaCronograma, prsPagos
    'CTI2 FERIMORO : ERS034-2019    26082020
    Dim oNSeg As New COMDCaptaGenerales.DCOMSeguros
    Dim sSep As Integer, sIni As Integer, sMax As Integer
    Dim sFecha As String, sHora As String
    Dim sRegistroSBS As String, cNumPoliza As String, lscad As String
    Dim sImporte As String
    Dim rs As ADODB.Recordset

    lscad = ""
    lscad = lscad & oImp.gPrnInicializa
    lscad = lscad & oImp.gPrnEspaLineaN   'espaciamiento lineas 1/6 pulg.
    lscad = lscad & oImp.gPrnTamPagina22  'Longitud de página a 22 líneas'
    lscad = lscad & oImp.gPrnTamLetra10CPI   'Tamaño 10 cpi
    lscad = lscad & oImp.gPrnTpoLetraRoman     'Tipo de Letra Sans Serif
    lscad = lscad & oImp.gPrnCondensadaOFF ' cancela condensada
    lscad = lscad & oImp.gPrnNegritaOFF ' desactiva negrita

    sSep = 14
    sIni = 0
    sMax = 39

    sFecha = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)
    sHora = Mid(psMovNro, 9, 2) & ":" + Mid(psMovNro, 11, 2) & ":" + Mid(psMovNro, 13, 2)
    
    Set rs = oNSeg.SegSepelioDatosOperacion(pnMovNro, pConn)
    Set oNSeg = Nothing
    
    sImporte = "S/. " & CStr(Format(rs!nMontoPrima, "#,##0.00"))
    
    If Not pbImpTMU Then
        lscad = lscad & oImp.gPrnNegritaON
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "CMAC MAYNAS" & Space(sMax - Len("CMAC MAYNAS") + sSep) & "CMAC MAYNAS" & Space(sMax - Len("CMAC MAYNAS")) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psNomAge & Space(31 - Len(psNomAge)) & pnMovNro & Space(sSep) & psNomAge & Space(31 - Len(psNomAge)) & pnMovNro & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(sSep) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImp.gPrnSaltoLinea
        If rs!cOpeDesc = "" Then
            'lscad = lscad & Space(sIni) & "--AFILIACION AL MICROSEGURO DE SEPELIO--" & Space(sSep) & "--AFILIACION AL MICROSEGURO DE SEPELIO--" & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "-AFILIACION MICROSEGURO VIDA Y SEPELIO-" & Space(sSep) & "-AFILIACION MICROSEGURO VIDA Y SEPELIO-" & oImp.gPrnSaltoLinea 'APRI20171024 ERS0287-2017
        Else
            lscad = lscad & Space(sIni) & "--" & rs!cOpeDesc & "--" & Space(sSep) & "--" & rs!cOpeDesc & "--" & oImp.gPrnSaltoLinea
        End If
        'lscad = lscad & Space(sIni) & rs!CiaSeguro & Space(sSep + 2) & rs!CiaSeguro & oImp.gPrnSaltoLinea
        '*********************APRI20171007 MEJORA********************************
        lscad = lscad & Space(sIni) & Left(rs!CiaSeguro, 39) & Space(sSep + 2) & Left(rs!CiaSeguro, 39) & oImp.gPrnSaltoLinea
        If Len(rs!CiaSeguro) > 39 Then
            lscad = lscad & Space(sIni) & Right(rs!CiaSeguro, Len(rs!CiaSeguro) - 39) & Space(sSep + (39 - (Len(rs!CiaSeguro) - 39))) & Right(rs!CiaSeguro, Len(rs!CiaSeguro) - 39) & oImp.gPrnSaltoLinea
        End If
        '*********************END APRI20171007********************************
        lscad = lscad & oImp.gPrnNegritaOFF
        lscad = lscad & Space(sIni) & "REGISTRO SBS    " & rs!cRegistroSBS & Space(sMax - Len("REGISTRO SBS    " & rs!cRegistroSBS) + sSep) & "REGISTRO SBS    " & rs!cRegistroSBS & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "POLIZA N        " & rs!cNumPoliza & Space(sMax - Len("POLIZA Nº       " & rs!cNumPoliza) + sSep) & "POLIZA N*       " & rs!cNumPoliza & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "CERTIFICADO N   " & rs!cNumCertificado & Space(sMax - Len("CERTIFICADO Nº  " & rs!cNumCertificado) + sSep) & "CERTIFICADO N*  " & rs!cNumCertificado & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "FECHA INICIO    " & Format(rs!dFecAfiliacion, "dd/MM/yyyy") & Space(sMax - Len("FECHA INICIO    " & Format(rs!dFecAfiliacion, "dd/MM/yyyy")) + sSep) & "FECHA INICIO    " & Format(rs!dFecAfiliacion, "dd/MM/yyyy") & oImp.gPrnSaltoLinea '& oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "FORMA DE PAGO:  " & IIf(rs!bPrimaAnual = 0, "MENSUAL", "ANUAL") & Space(sMax - Len("FORMA DE PAGO:  " & IIf(rs!bPrimaAnual = 0, "MENSUAL", "ANUAL")) + sSep) & "FORMA DE PAGO:  " & IIf(rs!bPrimaAnual = 0, "MENSUAL", "ANUAL") & oImp.gPrnSaltoLinea 'APRI20181023 ERS071-2018
        lscad = lscad & Space(sIni) & Left(rs!cPersNombre, 39) & Space(sMax - Len(Left(rs!cPersNombre, 39)) + sSep) & Left(rs!cPersNombre, 39) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "DOI             " & rs!cPersIDnro & Space(sMax - Len("DNI/CE/RUC      " & rs!cPersIDnro) + sSep) & "DOI             " & rs!cPersIDnro & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        If rs!cOpeDesc = "" Then
            lscad = lscad & Space(sIni) & "N BENEFICIARIOS " & rs!nCantBenf & Space(sMax - Len("N BENEFICIARIOS      " & rs!cPersIDnro) + sSep + 12) & "N BENEFICIARIOS      " & rs!nCantBenf & oImp.gPrnSaltoLinea
            'lscad = lscad & Space(sIni) & "MEDIO DE PAGO   " & rs!cMedioPago & Space(sMax - Len("MEDIO DE PAGO      " & rs!cPersIDnro) + 17) & "MEDIO DE PAGO      " & rs!cMedioPago & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "MEDIO DE PAGO   " & rs!cMedioPago & Space(sMax - Len("MEDIO DE PAGO      " & rs!cMedioPago) + 17) & "MEDIO DE PAGO      " & rs!cMedioPago & oImp.gPrnSaltoLinea 'APRI20171024 ERS028-2017
            lscad = lscad & Space(sIni) & "NRO. CUENTA     " & rs!cCtaCodDebito & Space(sMax - Len("NRO. CUENTA     " & rs!cCtaCodDebito) + sSep) & "NRO. CUENTA     " & rs!cCtaCodDebito & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        End If
        lscad = lscad & Space(sIni) & "Usuario Venta       " & rs!cUserVenta & Space(sMax - Len("Usuario Venta       " & rs!cUserVenta) + sSep) & "Usuario Venta       " & rs!cUserVenta & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        'lscad = lscad & Space(sIni) & "Usuario Cobra       " & rs!cUserSistema & Space(sMax - Len("Usuario Cobra       " & rs!cUserSistema) + sSep) & "Usuario Cobra       " & rs!cUserSistema & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea 'CTI2 FERIMORO : ERS034-2019 (03112020)
        lscad = lscad & Space(sIni) & String(sMax, "-") & Space(sSep) & String(sMax, "-") & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnNegritaON
        If rs!cOpeDesc = "" Then
            'lscad = lscad & Space(sIni) & "PRIMA" & Space(sMax - Len("PRIMA") - Len(sImporte)) & sImporte & Space(sSep) & "PRIMA" & Space(sMax - Len("PRIMA") - Len(sImporte)) & sImporte & oImp.gPrnSaltoLinea
            'CTI2 FERIMORO : ERS034-2019    26082020
            'lscad = lscad & Space(sIni) & "COBRO PRIMA" & Space(sMax - Len("COBRO PRIMA") - Len(sImporte)) & sImporte & Space(sSep) & "COBRO PRIMA" & Space(sMax - Len("PRIMA") - Len(sImporte)) & sImporte & oImp.gPrnSaltoLinea 'APRI20171024 ERS028-2017
            Dim ntotal As Double
            Dim bResulTotal As String
            Dim nConcepto As String
            Dim nConcMonto As String
 
            If prsPagos Is Nothing Then  'CORREGIDO CTI2 FERIMORO : ERROR EN EL PASE 12012021
                ntotal = CStr(Format(rs!nMontoPrima, "#,##0.00"))
            Else
                Do While Not prsPagos.EOF
                    nConcepto = "COBRO PRIMA " & prsPagos!mes
                    nConcMonto = prsPagos!Monto
                    lscad = lscad & Space(sIni) & nConcepto & Space(sMax - Len(nConcepto) - Len(nConcMonto)) & nConcMonto & Space(sSep) & nConcepto & Space(sMax - Len(nConcepto) - Len(nConcMonto)) & nConcMonto & oImp.gPrnSaltoLinea
                    ntotal = ntotal + CDbl(prsPagos!nPrima)
                    prsPagos.MoveNext
                Loop
            End If
            
            bResulTotal = "S/ " & CStr(Format(ntotal, "#,##0.00"))
            lscad = lscad & Space(sIni) & String(sMax, "-") & Space(sSep) & String(sMax, "-") & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "TOTAL PRIMA" & Space(sMax - Len("TOTAL PRIMA") - Len(bResulTotal)) & bResulTotal & Space(sSep) & "TOTAL PRIMA" & Space(sMax - Len("TOTAL PRIMA") - Len(bResulTotal)) & bResulTotal & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & String(sMax, "-") & Space(sSep) & String(sMax, "-") & oImp.gPrnSaltoLinea
            
            'lscad = lscad & Space(sIni) & "Prox. Pago" & Space(sMax - Len("Prox. Pago") - Len(rs!dFecProxPag)) & Format(rs!dFecProxPag, "dd/MM/yyyy") & Space(sSep) & "Prox. Pago" & Space(sMax - Len("Prox. Pago") - Len(rs!dFecProxPag)) & Format(rs!dFecProxPag, "dd/MM/yyyy") & oImp.gPrnSaltoLinea
            If (nProxFechaCronograma <> "") Then
              lscad = lscad & Space(sIni) & "Prox. Pago" & Space(sMax - Len("Prox. Pago") - Len(nProxFechaCronograma)) & Format(nProxFechaCronograma, "dd/MM/yyyy") & Space(sSep) & "Prox. Pago" & Space(sMax - Len("Prox. Pago") - Len(nProxFechaCronograma)) & Format(nProxFechaCronograma, "dd/MM/yyyy") & oImp.gPrnSaltoLinea
            Else
              lscad = lscad & Space(sIni) & "Prox. Pago" & Space(sMax - Len("Prox. Pago") - Len(rs!dFecProxPag)) & Format(rs!dFecProxPag, "dd/MM/yyyy") & Space(sSep) & "Prox. Pago" & Space(sMax - Len("Prox. Pago") - Len(rs!dFecProxPag)) & Format(rs!dFecProxPag, "dd/MM/yyyy") & oImp.gPrnSaltoLinea
            End If
            
            
        Else
            lscad = lscad & Space(sIni) & "MONTO DEVOLVER  :   " & Space(sMax - Len("MONTO DEVOLVER  :   ") - Len(sImporte) - 10) & sImporte & Space(sSep - 1) & "MONTO DEVOLVER  :   " & Space(sMax - Len("MONTO DEVOLVER  :   ") - Len(sImporte)) & sImporte & oImp.gPrnSaltoLinea
        End If
        lscad = lscad & Space(sIni) & String(sMax, "-") & Space(sSep) & String(sMax, "-") & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnNegritaOFF
        'lscad = lscad & Space(sIni) & "*** Con el  uso  de  mi  clave  secreta " & Space(sSep) & "*** Con el  uso  de  mi  clave  secreta " & oImp.gPrnSaltoLinea
        If rs!cOpeDesc = "" Then
            'APRI20171024 ERS028-2017
            lscad = lscad & Space(sIni) & "*** Dejo constancia de haber tomado " & Space(sSep + 3) & "*** Dejo constancia de haber tomado " & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "conocimiento de las condiciones generales," & Space(sSep - 3) & "conocimiento de las condiciones generales," & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "particulares y cláusulas adicionales" & Space(sSep + 3) & "particulares y cláusulas adicionales" & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "de la presente póliza***" & Space(sSep + 15) & "de la presente póliza***" & oImp.gPrnSaltoLinea
            'END APRI20171024
        Else
            lscad = lscad & Space(sIni) & "*** Con el  uso  de  mi  clave  secreta " & Space(sSep) & "*** Con el  uso  de  mi  clave  secreta " & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "autorizo la desafiliación del Microseguro" & Space(sSep) & "autorizo la desafiliación del Microseguro" & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "de Sepelio ***" & Space(18) & Space(sSep) & Space(8) & "de Sepelio ***" & oImp.gPrnSaltoLinea
        End If
        'lscad = lscad & Space(sIni) & "de Sepelio ***" & Space(18) & Space(sSep) & Space(8) & "de Sepelio ***" & oImp.gPrnSaltoLinea
        'APRI20171024 ERS028-2017
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "Firma y DOI del Asegurado" & Space(sSep + 14) & "Firma y DOI del Asegurado" & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        'END APRI
        lscad = lscad & Space(sIni) & Right(psMovNro, 4) & Space(sMax - Len(Right(psMovNro, 4)) + sSep) & Right(psMovNro, 4) & oImp.gPrnSaltoLinea
    Else
        lscad = lscad & oImp.gPrnNegritaON
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "CMAC MAYNAS" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psNomAge & Space(31 - Len(psNomAge)) & pnMovNro & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImp.gPrnSaltoLinea
        If rs!cOpeDesc = "" Then
            'lscad = lscad & Space(sIni) & "--AFILIACION AL MICROSEGURO DE SEPELIO--" & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "--AFILIACION MICROSEGURO VIDA Y SEPELIO--" & oImp.gPrnSaltoLinea 'APRI20171024 ERS028-2017
        Else
            lscad = lscad & Space(sIni) & "--" & rs!cOpeDesc & "--" & oImp.gPrnSaltoLinea
        End If
        
        'lscad = lscad & Space(sIni) & rs!CiaSeguro & oImp.gPrnSaltoLinea
        '*********************APRI20171007 MEJORA********************************
        lscad = lscad & Space(sIni) & Left(rs!CiaSeguro, 39) & oImp.gPrnSaltoLinea
        If Len(rs!CiaSeguro) > 39 Then
            lscad = lscad & Space(sIni) & Right(rs!CiaSeguro, Len(rs!CiaSeguro) - 39) & oImp.gPrnSaltoLinea
        End If
        '*********************END APRI20171007********************************
        lscad = lscad & Space(sIni) & "REGISTRO SBS  :      " & rs!cRegistroSBS & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "POLIZA N      :      " & rs!cNumPoliza & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "CERTIFICADO N :      " & rs!cNumCertificado & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "FECHA INICIO  :      " & Format(rs!dFecAfiliacion, "dd/MM/yyyy") & oImp.gPrnSaltoLinea '& oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "FORMA DE PAGO :      " & IIf(rs!bPrimaAnual = 0, "MENSUAL", "ANUAL") & oImp.gPrnSaltoLinea 'APRI20181023 ERS071-2018
        'CTI2 FERIMORO : CORRIGE PROCESO DE IMPRESION
        'lscad = lscad & Space(sIni) & "CLIENTE: " & rs!cpersnombre & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "" & rs!cPersNombre & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "DOI    : " & rs!cPersIDnro & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        
'CTI2 FERIMORO : CORRIGE IMMPRESION DE VOUCHER
'        If rs!cOpeDesc = "" Then
'            lscad = lscad & Space(sIni) & "N BENEFICIARIOS :   " & rs!nCantBenf & oImp.gPrnSaltoLinea
'            lscad = lscad & Space(sIni) & "MEDIO DE PAGO   :   " & rs!cMedioPago & oImp.gPrnSaltoLinea
'            lscad = lscad & Space(sIni) & "CUENTA VINCULADA:   " & rs!cCtaCodDebito & oImp.gPrnSaltoLinea
'            lscad = lscad & Space(sIni) & "PRIMA           :   " & sImporte & oImp.gPrnSaltoLinea
'            lscad = lscad & Space(sIni) & "Prox. Pago      :   " & Format(rs!dFecProxPag, "dd/MM/yyyy") & oImp.gPrnSaltoLinea
'        Else
'            lscad = lscad & Space(sIni) & "MONTO DEVOLVER  :   " & sImporte & oImp.gPrnSaltoLinea
'        End If
        If rs!cOpeDesc = "" Then
            lscad = lscad & Space(sIni) & "N BENEFICIARIOS " & rs!nCantBenf & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "MEDIO DE PAGO   " & rs!cMedioPago & oImp.gPrnSaltoLinea 'APRI20171024 ERS028-2017
            lscad = lscad & Space(sIni) & "NRO. CUENTA     " & rs!cCtaCodDebito & oImp.gPrnSaltoLinea
        End If
        lscad = lscad & Space(sIni) & "Usuario Venta       " & rs!cUserVenta & oImp.gPrnSaltoLinea
        
        lscad = lscad & Space(sIni) & String(sMax, "-") & oImp.gPrnSaltoLinea
        
        If rs!cOpeDesc = "" Then
            'lscad = lscad & Space(sIni) & "PRIMA" & oImp.gPrnSaltoLinea
            'CTI2 FERIMORO : ERS034-2019    26082020
            'lscad = lscad & Space(sIni) & "COBRO PRIMA" & Space(sMax - Len("COBRO PRIMA") - Len(sImporte)) & sImporte & Space(sSep) & "COBRO PRIMA" & Space(sMax - Len("PRIMA") - Len(sImporte)) & sImporte & oImp.gPrnSaltoLinea 'APRI20171024 ERS028-2017

 
            If prsPagos Is Nothing Then  'CORREGIDO CTI2 FERIMORO : ERROR EN EL PASE 12012021
                ntotal = CStr(Format(rs!nMontoPrima, "#,##0.00"))
            Else
                Do While Not prsPagos.EOF
                    nConcepto = "COBRO PRIMA " & prsPagos!mes
                    nConcMonto = prsPagos!Monto
                    lscad = lscad & Space(sIni) & nConcepto & oImp.gPrnSaltoLinea
                    ntotal = ntotal + CDbl(prsPagos!nPrima)
                    prsPagos.MoveNext
                Loop
            End If
            
            bResulTotal = "S/ " & CStr(Format(ntotal, "#,##0.00"))
            lscad = lscad & Space(sIni) & String(sMax, "-") & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "TOTAL PRIMA" & bResulTotal & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & String(sMax, "-") & oImp.gPrnSaltoLinea
            
            'lscad = lscad & Space(sIni) & "Prox. Pago" & Space(sMax - Len("Prox. Pago") - Len(rs!dFecProxPag)) & Format(rs!dFecProxPag, "dd/MM/yyyy") & Space(sSep) & "Prox. Pago" & Space(sMax - Len("Prox. Pago") - Len(rs!dFecProxPag)) & Format(rs!dFecProxPag, "dd/MM/yyyy") & oImp.gPrnSaltoLinea
            If (nProxFechaCronograma <> "") Then
              lscad = lscad & Space(sIni) & "Prox. Pago" & Space(sMax - Len("Prox. Pago") - Len(nProxFechaCronograma)) & Format(nProxFechaCronograma, "dd/MM/yyyy") & oImp.gPrnSaltoLinea
            Else
              lscad = lscad & Space(sIni) & "Prox. Pago" & Space(sMax - Len("Prox. Pago") - Len(rs!dFecProxPag)) & Format(rs!dFecProxPag, "dd/MM/yyyy") & oImp.gPrnSaltoLinea
            End If
            
            
        Else
            lscad = lscad & Space(sIni) & "MONTO DEVOLVER" & Space(sMax - Len("MONTO") - Len(sImporte) - 10) & sImporte & oImp.gPrnSaltoLinea
        End If
        lscad = lscad & Space(sIni) & String(sMax, "-") & oImp.gPrnSaltoLinea
        
        'lscad = lscad & Space(sIni) & "***Con el uso de mi clave secreta " & oImp.gPrnSaltoLinea
        If rs!cOpeDesc = "" Then
             'APRI20171024 ERS028-2017
            lscad = lscad & Space(sIni) & "*** Dejo constancia de haber tomado " & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "conocimiento de las condiciones" & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "generales, particulares y cláusulas" & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "adicionales de la presente póliza***" & oImp.gPrnSaltoLinea
            'END APRI20171024
        Else
            lscad = lscad & Space(sIni) & "***Con el uso de mi clave secreta " & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "autorizo la desafiliación del Microseguro" & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & "de Sepelio***" & oImp.gPrnSaltoLinea
        End If
        'lscad = lscad & Space(sIni) & "de Sepelio***" & oImp.gPrnSaltoLinea
        'APRI20171024 ERS028-2017
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "Firma y DOI del Asegurado" & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & Right(psMovNro, 4) & oImp.gPrnSaltoLinea
        'END APRI
        lscad = lscad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
        lscad = lscad & oImp.gPrnNegritaOFF
    End If
    ImprimeBoletaAfilicacionSeguroSepelio = lscad
End Function

Public Function ImprimeBoletaDesafilicacionSegSepelio(ByVal pnMovNro As Long, ByVal psMovNro As String, ByVal psNomAge As String, ByVal pbImpTMU As Boolean) As String
    Dim oNSeg As New COMDCaptaGenerales.DCOMSeguros
    Dim sSep As Integer, sIni As Integer, sMax As Integer
    Dim sFecha As String, sHora As String
    Dim sRegistroSBS As String, cNumPoliza As String, lscad As String
    Dim sImporte As String
    Dim rs As ADODB.Recordset

    lscad = ""
    lscad = lscad & oImp.gPrnInicializa
    lscad = lscad & oImp.gPrnEspaLineaN   'espaciamiento lineas 1/6 pulg.
    lscad = lscad & oImp.gPrnTamPagina22  'Longitud de página a 22 líneas'
    lscad = lscad & oImp.gPrnTamLetra10CPI   'Tamaño 10 cpi
    lscad = lscad & oImp.gPrnTpoLetraRoman     'Tipo de Letra Sans Serif
    lscad = lscad & oImp.gPrnCondensadaOFF ' cancela condensada
    lscad = lscad & oImp.gPrnNegritaOFF ' desactiva negrita

    sSep = 14
    sIni = 0
    sMax = 39

    sFecha = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)
    sHora = Mid(psMovNro, 9, 2) & ":" + Mid(psMovNro, 11, 2) & ":" + Mid(psMovNro, 13, 2)
    
    Set rs = oNSeg.SegSepelioDatosOperacion(pnMovNro)
    Set oNSeg = Nothing
    
    sImporte = 0
    
    If Not pbImpTMU Then
        lscad = lscad & oImp.gPrnNegritaON
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "CMAC MAYNAS" & Space(sMax - Len("CMAC MAYNAS") + sSep) & "CMAC MAYNAS" & Space(sMax - Len("CMAC MAYNAS")) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psNomAge & Space(31 - Len(psNomAge)) & pnMovNro & Space(sSep) & psNomAge & Space(31 - Len(psNomAge)) & pnMovNro & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(sSep) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "--DESAFILIACION MICROSEGURO DE SEPELIO--" & Space(sSep) & "--DESAFILIACION MICROSEGURO DE SEPELIO--" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & rs!CiaSeguro & Space(sSep + 2) & rs!CiaSeguro & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnNegritaOFF
        lscad = lscad & Space(sIni) & "REGISTRO SBS    " & rs!cRegistroSBS & Space(sMax - Len("REGISTRO SBS    " & rs!cRegistroSBS) + sSep) & "REGISTRO SBS    " & rs!cRegistroSBS & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "POLIZA N        " & rs!cNumPoliza & Space(sMax - Len("POLIZA Nº       " & rs!cNumPoliza) + sSep) & "POLIZA N*       " & rs!cNumPoliza & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "CERTIFICADO N   " & rs!cNumCertificado & Space(sMax - Len("CERTIFICADO Nº  " & rs!cNumCertificado) + sSep) & "CERTIFICADO N*  " & rs!cNumCertificado & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "FECHA INICIO    " & Format(rs!dFecAfiliacion, "dd/MM/yyyy") & Space(sMax - Len("FECHA INICIO    " & Format(rs!dFecAfiliacion, "dd/MM/yyyy")) + sSep) & "FECHA INICIO    " & Format(rs!dFecAfiliacion, "dd/MM/yyyy") & oImp.gPrnSaltoLinea '& oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "FORMA DE PAGO:  " & IIf(rs!bPrimaAnual = 0, "MENSUAL", "ANUAL") & Space(sMax - Len("FORMA DE PAGO:  " & IIf(rs!bPrimaAnual = 0, "MENSUAL", "ANUAL")) + sSep) & "FORMA DE PAGO:  " & IIf(rs!bPrimaAnual = 0, "MENSUAL", "ANUAL") & oImp.gPrnSaltoLinea 'APRI20181023 ERS071-2018
        lscad = lscad & Space(sIni) & Left(rs!cPersNombre, 39) & Space(sMax - Len(Left(rs!cPersNombre, 39)) + sSep) & Left(rs!cPersNombre, 39) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "DOI             " & rs!cPersIDnro & Space(sMax - Len("DNI/CE/RUC      " & rs!cPersIDnro) + sSep) & "DOI             " & rs!cPersIDnro & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnNegritaON
        lscad = lscad & Space(sIni) & String(sMax, "-") & Space(sSep) & String(sMax, "-") & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnNegritaOFF
        lscad = lscad & Space(sIni) & "*** Con el  uso  de  mi  clave  secreta " & Space(sSep) & "*** Con el  uso  de  mi  clave  secreta " & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "autorizo la desafiliación del Microseguro" & Space(sSep) & "autorizo la desafiliación del Microseguro" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "de Sepelio ***" & Space(18) & Space(sSep) & Space(8) & "de Sepelio ***" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & Right(psMovNro, 4) & Space(sMax - Len(Right(psMovNro, 4)) + sSep) & Right(psMovNro, 4) & oImp.gPrnSaltoLinea
    Else
        lscad = lscad & oImp.gPrnNegritaON
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "CMAC MAYNAS" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psNomAge & Space(31 - Len(psNomAge)) & pnMovNro & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "--DESAFILIACION MICROSEGURO DE SEPELIO--" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & rs!CiaSeguro & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "REGISTRO SBS  :      " & rs!cRegistroSBS & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "POLIZA N      :      " & rs!cNumPoliza & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "CERTIFICADO N :      " & rs!cNumCertificado & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "FECHA INICIO  :      " & Format(rs!dFecAfiliacion, "dd/MM/yyyy") & oImp.gPrnSaltoLinea '& oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "FORMA DE PAGO :      " & IIf(rs!bPrimaAnual = 0, "MENSUAL", "ANUAL") & oImp.gPrnSaltoLinea 'APRI20181023 ERS071-2018
        lscad = lscad & Space(sIni) & "CLIENTE: " & rs!cPersNombre & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "DOI    : " & rs!cPersIDnro & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "***Con el uso de mi clave secreta " & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "autorizo la contratacion del Microseguro" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "de Sepelio***" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & Right(psMovNro, 4) & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
        lscad = lscad & oImp.gPrnNegritaOFF
    End If
    ImprimeBoletaDesafilicacionSegSepelio = lscad
End Function

Public Function SepelioBeneficiarios(ByVal psPersCod As String) As ADODB.Recordset
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Set SepelioBeneficiarios = obj.SepelioBeneficiarios(psPersCod)
    Set obj = Nothing
End Function
Public Function SepelioCoberturas() As ADODB.Recordset
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Set SepelioCoberturas = obj.SepelioCoberturas
    Set obj = Nothing
End Function
Public Sub RegistraSepelioCoberturas(ByVal pnIdSolicitud As Integer, ByVal pnIdCobertura As Integer)
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    obj.RegistraSepelioCoberturas pnIdSolicitud, pnIdCobertura
End Sub
Public Sub SepelioSolicCtaBenef(ByVal pnAseguradoID As Integer, ByVal psCtaCod As String)
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    obj.SepelioSolicCtaBenef pnAseguradoID, psCtaCod
End Sub
Public Function SepelioDevolverMovimientoExtornar(ByVal psDatoBusqueda As String, ByVal psFecha As String, ByVal psOpeCod As String, ByVal psAgeCod As String, ByVal pnTipoBusqueda As Integer) As ADODB.Recordset
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Set SepelioDevolverMovimientoExtornar = obj.SepelioDevolverMovimientoExtornar(psDatoBusqueda, psFecha, psOpeCod, psAgeCod, pnTipoBusqueda)
    Set obj = Nothing
End Function
Public Function SepelioObtieneCriterios() As ADODB.Recordset
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Set SepelioObtieneCriterios = obj.SepelioObtieneCriterios
    Set obj = Nothing
End Function
Public Function SepelioObtieneDatosExtorno(ByVal pnMovNro As Long) As ADODB.Recordset
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Set SepelioObtieneDatosExtorno = obj.SepelioObtieneDatosExtorno(pnMovNro)
    Set obj = Nothing
End Function
Public Function SepelioObtieneDatos(ByVal psPersCod As String) As ADODB.Recordset
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Set SepelioObtieneDatos = obj.SepelioObtieneDatos(psPersCod)
    Set obj = Nothing
End Function
Public Sub SepelioActualizaDatos(ByVal psNumCertifNue As String, ByVal psNumCertifAnt As String, ByVal pnTpoPago As Integer, ByVal psCtaCod As String, ByVal pnMovNro As Long, ByVal pnEstado As Integer)
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Call obj.SepelioActualizaDatos(psNumCertifNue, psNumCertifAnt, pnTpoPago, psCtaCod, pnMovNro, pnEstado)
    Set obj = Nothing
End Sub
Public Sub SepelioLimpiaBeneficiario(ByVal psNumCertif As String)
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Call obj.SepelioLimpiaBeneficiario(psNumCertif)
    Set obj = Nothing
End Sub
Public Function SepelioVerificaPagoMes(ByVal psFecSis As String, ByVal psPersCod As String) As Boolean
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Dim oRS As New ADODB.Recordset
    Set oRS = obj.SepelioVerificaPagoMes(psFecSis, psPersCod)
    
    If Not (oRS.EOF And oRS.BOF) Then
        SepelioVerificaPagoMes = True
    Else
        SepelioVerificaPagoMes = False
    End If
    
    Set obj = Nothing
End Function
'RECO FIN ***************************************************************
'APRI20171004 ERS028-2017
Public Function SepelioVerificaTpoPrograma(ByVal nTpoSeguro As Integer, ByVal nProd As String, ByVal nTpoPrograma As Integer) As Boolean
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Dim oRS As New ADODB.Recordset
    Set oRS = obj.SepelioVerificaTpoPrograma(nTpoSeguro, nProd, nTpoPrograma)
    
    If Not (oRS.EOF And oRS.BOF) Then
        SepelioVerificaTpoPrograma = oRS!nValue
    Else
        SepelioVerificaTpoPrograma = False
    End If
    
    Set obj = Nothing
End Function
Public Function SepelioVerificaPagoEfectivo(ByVal psPersCod As String) As Boolean
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Dim oRS As New ADODB.Recordset
    Set oRS = obj.SepelioVerificaPagoEfectivo(psPersCod)
    
    If Not (oRS.EOF And oRS.BOF) Then
        SepelioVerificaPagoEfectivo = oRS!nVal
    Else
        SepelioVerificaPagoEfectivo = False
    End If
    
    Set obj = Nothing
End Function
Public Function SepelioVerificaProducto(ByVal pcPersCod As String) As Boolean
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Dim oRS As New ADODB.Recordset
    Set oRS = obj.SepelioVerificaProducto(pcPersCod)
    
    If Not (oRS.EOF And oRS.BOF) Then
        SepelioVerificaProducto = True
    Else
        SepelioVerificaProducto = False
    End If
    
    Set obj = Nothing
End Function
Public Function SegSepelioParametros(ByVal psParamID As Integer) As Double
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Dim oRS As New ADODB.Recordset
    Set oRS = obj.SepelioParametros(psParamID)
    
    If Not (oRS.EOF And oRS.BOF) Then
        SegSepelioParametros = oRS!nValor
    Else
        SegSepelioParametros = 0
    End If
    
    Set obj = Nothing
End Function
Public Function ObservacionesCredSegDesg(ByVal psCtaCod As String) As Integer
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Dim oRS As New ADODB.Recordset
    Set oRS = obj.ObtenerObservacionesCredSegDesg(psCtaCod)
    
    If Not (oRS.EOF And oRS.BOF) Then
            ObservacionesCredSegDesg = oRS!ntotal
    Else
        ObservacionesCredSegDesg = 0
    End If
    
    Set obj = Nothing
    
End Function
'END APRI

'CTI2 FERIMORO : ERS034-2019   25082020
Public Function ActualizaSegurosCronograma(ByVal tpoOpcion As Integer, ByVal psFecAfiliacion As String, ByVal tpoSeguro As Integer, _
                                           ByVal pnMovNroReg As Long, ByVal psNumCertificado As String, ByVal nExiSegCrono As Integer)
       
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Call oDCOMSeg.ActuSeguroCronograma(tpoOpcion, psFecAfiliacion, tpoSeguro, pnMovNroReg, psNumCertificado, nExiSegCrono)
    Set oDCOMSeg = Nothing

End Function
'CTI2 FERIMORO : ERS034-2019    29082020
Public Function BuscaFormularioExiste(ByVal cNroFormulario As String) As Integer
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Dim oRS As New ADODB.Recordset
    Set oRS = obj.BuscaFormulariosExisteD(cNroFormulario)
    
    If Not (oRS.EOF And oRS.BOF) Then
            BuscaFormularioExiste = oRS!existe
    Else
        BuscaFormularioExiste = 0
    End If
    
    Set obj = Nothing
    
End Function
'CTI2 FERIMORO : ERS034-2019       29082020
Public Function ImprimeBoletaAfiliacionSegSoat(ByVal pnMovNro As Long, ByVal psMovNro As String, ByVal psNomAge As String, ByVal pbImpTMU As Boolean, _
                                               Optional ByVal pConn As COMConecta.DCOMConecta = Nothing)
 Dim lscad As String
 Dim sSep As Integer, sIni As Integer, sMax As Integer
 Dim oNSeg As New COMDCaptaGenerales.DCOMSeguros
 Dim rs As ADODB.Recordset
 Dim sFecha As String
 Dim sHora As String
 
    lscad = ""
    
    lscad = lscad & oImp.gPrnInicializa
    lscad = lscad & oImp.gPrnEspaLineaN   'espaciamiento lineas 1/6 pulg.
    lscad = lscad & oImp.gPrnTamPagina22  'Longitud de página a 22 líneas'
    lscad = lscad & oImp.gPrnTamLetra10CPI   'Tamaño 10 cpi
    lscad = lscad & oImp.gPrnTpoLetraRoman     'Tipo de Letra Sans Serif
    lscad = lscad & oImp.gPrnCondensadaOFF ' cancela condensada
    lscad = lscad & oImp.gPrnNegritaOFF ' desactiva negrita
    
    sSep = 14
    sIni = 0
    sMax = 39
    
    sFecha = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4)
    sHora = Mid(psMovNro, 9, 2) & ":" + Mid(psMovNro, 11, 2) & ":" + Mid(psMovNro, 13, 2)
    
    Set rs = oNSeg.ObtieneDatosSegSoat(pnMovNro, pConn)
    Set oNSeg = Nothing
    
    If Not pbImpTMU Then
        
        lscad = lscad & oImp.gPrnNegritaON
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "CMAC MAYNAS" & Space(sMax - Len("CMAC MAYNAS") + sSep) & "CMAC MAYNAS" & Space(sMax - Len("CMAC MAYNAS")) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psNomAge & Space(31 - Len(psNomAge)) & pnMovNro & Space(sSep) & psNomAge & Space(31 - Len(psNomAge)) & pnMovNro & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(sSep) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImp.gPrnSaltoLinea
            
        lscad = lscad & Space(sIni) & "------ AFILIACION SEGURO SOAT ------" & Space(sSep + 3) & "------ AFILIACION SEGURO SOAT ------" & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnNegritaOFF
        lscad = lscad & Space(sIni) & "FORMULARIO N    " & rs!cNumCertificado & Space(sMax - Len("FORMULARIO N    " & rs!cNumCertificado) + sSep) & "FORMULARIO N    " & rs!cNumCertificado & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "USO             " & rs!cUso & Space(sMax - Len("USO             " & rs!cUso) + sSep) & "USO             " & rs!cUso & oImp.gPrnSaltoLinea
        'CTI2 FERIMORO: CORREO GELU : SE MEJORA MODULO
        'lscad = lscad & Space(sIni) & "USUARIO VENTA   " & "BRSO" & Space(sMax - Len("USUARIO VENTA   " & "BRSO") + sSep) & "USUARIO VENTA   " & "BRSO" & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "USUARIO VENTA   " & rs!cUsuario & Space(sMax - Len("USUARIO VENTA   " & rs!cUsuario) + sSep) & "USUARIO VENTA   " & rs!cUsuario & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    
        lscad = lscad & Space(sIni) & rs!cPersNombre & Space(sMax - Len(rs!cPersNombre) + sSep) & rs!cPersNombre & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "DOI             " & rs!cDOI & Space(sMax - Len("DOI             " & rs!cDOI) + sSep) & "DOI             " & rs!cDOI & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & String(sMax, "-") & Space(sSep) & String(sMax, "-") & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "PRIMA " & Space(sMax - Len("PRIMA ") - Len("S/ " & CStr(Format(rs!nPrima, "#,##0.00")))) & "S/ " & CStr(Format(rs!nPrima, "#,##0.00")) & Space(sSep) & "PRIMA " & Space(sMax - Len("PRIMA ") - Len("S/ " & CStr(Format(rs!nPrima, "#,##0.00")))) & "S/ " & CStr(Format(rs!nPrima, "#,##0.00")) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & String(sMax, "-") & Space(sSep) & String(sMax, "-") & oImp.gPrnSaltoLinea
        'lscad = lscad & Space(sIni) & "Prox. Pago" & Space(sMax - Len("Prox. Pago") - Len(rs!fFecTermino)) & Format(rs!fFecTermino, "dd/MM/yyyy") & Space(sSep) & "Prox. Pago" & Space(sMax - Len("Prox. Pago") - Len(rs!fFecTermino)) & Format(rs!fFecTermino, "dd/MM/yyyy") & oImp.gPrnSaltoLinea
        'lscad = lscad & Space(sIni) & String(sMax, "-") & Space(sSep) & String(sMax, "-") & oImp.gPrnSaltoLinea
    
        lscad = lscad & Space(sIni) & "*** Dejo constancia de haber tomado " & Space(sSep + 3) & "*** Dejo constancia de haber tomado " & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "conocimiento de las condiciones generales," & Space(sSep - 3) & "conocimiento de las condiciones generales," & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "particulares y cláusulas adicionales" & Space(sSep + 3) & "particulares y cláusulas adicionales" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "de la presente póliza***" & Space(sSep + 15) & "de la presente póliza***" & oImp.gPrnSaltoLinea
    
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "Firma y DOI del Asegurado" & Space(sSep + 14) & "Firma y DOI del Asegurado" & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & Right(psMovNro, 4) & Space(sMax - Len(Right(psMovNro, 4)) + sSep) & Right(psMovNro, 4) & oImp.gPrnSaltoLinea

        lscad = lscad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
        lscad = lscad & oImp.gPrnNegritaOFF
    'CTI2 FERIMORO : 11/02/2021 ANEXO2 -- SE CORRIGE POR MOTIVOS DE IMPRESION INADECUADA : CORREO: GELU
    Else
    
        lscad = lscad & oImp.gPrnNegritaON
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "CMAC MAYNAS" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & psNomAge & Space(31 - Len(psNomAge)) & pnMovNro & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImp.gPrnSaltoLinea
            
        lscad = lscad & Space(sIni) & "------ AFILIACION SEGURO SOAT ------" & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnNegritaOFF
        lscad = lscad & Space(sIni) & "FORMULARIO N    " & rs!cNumCertificado & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "USO             " & rs!cUso & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "USUARIO VENTA   " & rs!cUsuario & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "CLIENTE   " & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & rs!cPersNombre & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "DOI             " & rs!cDOI & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & String(sMax, "-") & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "PRIMA " & Space(sMax - Len("PRIMA ") - Len("S/ " & CStr(Format(rs!nPrima, "#,##0.00")))) & "S/ " & CStr(Format(rs!nPrima, "#,##0.00")) & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & String(sMax, "-") & oImp.gPrnSaltoLinea
    
        lscad = lscad & Space(sIni) & "*** Dejo constancia de haber tomado " & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "conocimiento de las condiciones generales, particulares y cláusulas adicionales" & oImp.gPrnSaltoLinea
        'lscad = lscad & Space(sIni) & "particulares y cláusulas adicionales" & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "de la presente póliza***" & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & "Firma y DOI del Asegurado" & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & Space(sIni) & Right(psMovNro, 4) & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnNegritaOFF
    '********************************************************
    End If
    ImprimeBoletaAfiliacionSegSoat = lscad
 
End Function
'CTI2 FERIMORO : ERS034-2019    29082020
Public Function BuscaTpoDoc(ByVal cPersCod As String) As String
    Dim obj As New COMDCaptaGenerales.DCOMSeguros
    Dim oRS As New ADODB.Recordset
    Set oRS = obj.BuscaTpoDoc(cPersCod)
    
    If Not (oRS.EOF And oRS.BOF) Then
            BuscaTpoDoc = oRS!cConsdescripcion
    Else
        BuscaTpoDoc = ""
    End If
    
    Set obj = Nothing
    
End Function
'ADD BY PTI1 20210927
Public Function ValidaPerSegSepelioA(ByVal sCuenta As String, ByVal fsPersCod As String) As Boolean
    
    Dim paResultado As Boolean
    
    Dim oCont As COMDCaptaGenerales.DCOMSeguros
    Set oCont = New COMDCaptaGenerales.DCOMSeguros

    paResultado = oCont.ValidaPerSegSepelioA(sCuenta, fsPersCod)
    ValidaPerSegSepelioA = paResultado

End Function
