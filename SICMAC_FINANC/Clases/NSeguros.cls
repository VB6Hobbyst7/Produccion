VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NSeguros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'PASIERS1362014
Public Type TSegAgencia
    sAgencia As String
    nMonto As Currency
End Type
Public Type TRetPagoSegDet
    sCtaCod As String
    nCuota As Integer
    nNroCalen As Integer
    nMonto As Currency
    sAgeCod As String
    dFechaVig As Date
End Type
'END PASI
Public Function ObtenerSegTarjetaSolicitudesPendientes(ByVal pnTipo As Integer, Optional ByVal psNumSolicitud As String = "") As ADODB.Recordset
    Dim oDSeg As New DSeguros
    Set ObtenerSegTarjetaSolicitudesPendientes = oDSeg.ObtenerSegTarjetaSolicitudesPendientes(pnTipo, psNumSolicitud)
    Set oDSeg = Nothing
End Function
Public Function RecuperarCtaAhorro(ByVal psNumSolicitud As String) As ADODB.Recordset
    Dim oDSeg As New DSeguros
    Set RecuperarCtaAhorro = oDSeg.RecuperarCtaAhorro(psNumSolicitud)
    Set oDSeg = Nothing
End Function
'Public Function GrabarDepositoXActivacionSegTarjeta(ByVal psPersCodIf As String, ByVal pdFechaDeposito As Date, ByVal pnMontoDeposito As Currency, ByVal pMatdepositos As Variant, ByVal pdFecsis As Date, ByVal pnMatTotalOpe As Integer, ByVal psCodAge As String, ByVal psCodUsu As String, _
'                                               ByVal psOpeCod As String, ByVal psCtaDebe As String, ByVal psObjetoIF As String, ByVal psCtaHaber As String, ByVal pMatPendiente As Variant, ByVal pbPendiente As Boolean) As String
'    Dim oDMov As New DMov
'    Dim bTrans As Boolean
'    Dim x As Integer, y As Integer
'    Dim lsMovNro As String
'    Dim lnMovNro As Long
'    Dim lnMovItem As Integer, lnMovOrden As Integer
'
'    On Error GoTo ErrGraba
'        oDMov.BeginTrans
'        bTrans = True
'
'        lsMovNro = oDMov.GeneraMovNro(pdFecsis, psCodAge, psCodUsu)
'        oDMov.InsertaMov lsMovNro, psOpeCod, "DEPOSITO PARA ACTIVACION DE SEGURO", gMovEstContabMovContable, gMovFlagVigente
'        lnMovNro = oDMov.GetnMovNro(lsMovNro)
'        lnMovItem = lnMovItem + 1
'        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, pnMontoDeposito 'pnTotalImporte
'        lnMovOrden = lnMovOrden + 1
'        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
'        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psObjetoIF, 4, 13), Mid(psObjetoIF, 1, 2), Mid(psObjetoIF, 18, 10)
'
'        For x = 1 To pnMatTotalOpe
'            Call oDMov.InsertaSegTarjetaEstado(pMatdepositos(x, 1), pdFecsis, 1001, pMatdepositos(x, 4), pMatdepositos(x, 3), lsMovNro)
'            'pMatdepositos(x, 1) = Solicitud, pMatdepositos(x, 4) = Glosa, pMatdepositos(x, 3) = Monto,pMatdepositos(x, 2) = CuentaAhorro
'            Call oDMov.ActualizarSegTarjetaSolicitudAproAcep(pMatdepositos(x, 1), lsMovNro, 1001, pMatdepositos(x, 2), CCur(pMatdepositos(x, 3)))
''            lsFiltro = ""
''            lsFiltro = oCont.GetFiltroObjetos(ObjProductosCMACT, psCtaHaber, rs!cObjetoCod, False)
''            lsFiltro = lsFiltro & oCont.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaHaber, rs!AreaCod & rs!AgeCod, False)
'        Next
'
'        lnMovItem = lnMovItem + 1
'        lnMovOrden = 0
'        'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, rs!Importe * -1
'        'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, CCur(pMatPendiente(3)) * -1
'        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnMontoDeposito * -1
'        lnMovOrden = lnMovOrden + 1
'        'oDSeg.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, rs!cObjetoCod
'        'oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, "6002232"
'        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, pMatdepositos(1, 7)
'        lnMovOrden = lnMovOrden + 1
'        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
'        'oDSeg.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, rs!AgeCod, rs!AreaCod
'        'oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, "01", "026"
'        oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, pMatdepositos(1, 6), pMatdepositos(1, 5)
'
'        If pdFechaDeposito <> pdFecsis Then
'            'oDMov.InsertaMovRef lnMovNro, rs!nMovNro
'            oDMov.InsertaMovRef lnMovNro, pMatPendiente(0)
'        End If
'
'        If Mid(psOpeCod, 3, 1) = "2" Then
'            oDMov.GeneraMovME lnMovNro, lsMovNro
'        End If
'        oDMov.ActualizaSaldoMovimiento lsMovNro, "+"
'
'        oDMov.CommitTrans
'        bTrans = False
'        GrabarDepositoXActivacionSegTarjeta = lsMovNro
'        Exit Function
'ErrGraba:
'    If bTrans Then oDMov.RollbackTrans
'    Set oDMov = Nothing
'    Err.Raise Err.Number, "", Err.Description
'End Function
Public Function RecuperarDepositosPendientes(ByVal psPersCod As String, ByVal psIFTpo As String, psCtaIFCod As String, ByVal pnImporte As Currency, ByVal pdFechaDeposito As Date) As ADODB.Recordset
    Dim oDSeg As New DSeguros
    Set RecuperarDepositosPendientes = oDSeg.RecuperarDepositosPendientes(psPersCod, psIFTpo, psCtaIFCod, pnImporte, pdFechaDeposito)
    Set oDSeg = Nothing
End Function
'JUEZ 20170711 ************************************************************
Public Sub GrabarRetiroPagoSeguroTarjeta(ByVal pdFecsis As String, ByVal psCodAge As String, ByVal psCodUsu As String, ByVal psOpecod As String, ByVal psPeriodo As String, _
                                         ByVal pnTipoDoc As Integer, ByVal psNroDoc As String, ByVal psFechaPago As String, ByVal psMovDesc As String, ByVal pnMontoRetiro As Currency, _
                                         ByVal psCtaHaber As String, ByVal psPersCodIf As String, ByVal psTipoIF As String, ByVal psCtaIFCod As String, ByRef psMovNro As String)

Dim oDCon As New DConecta
Dim oDMov As New DMov
Dim oDSeg As New DSeguros
Dim rs As ADODB.Recordset
Dim bTrans As Boolean
Dim x As Integer, y As Integer
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lnMovItem As Integer, lnMovOrden As Integer
Dim cCtaDebe As String
Dim sTime As String
    
    oDCon.AbreConexion
    sTime = oDCon.GetHoraServer
    oDCon.CierraConexion
    
    On Error GoTo ErrGraba
        oDMov.BeginTrans
        bTrans = True
        Set rs = oDMov.Ejecutar("select cCtaDebe=replace(cCtaContCod,'M','1') from OpeCtaNeg where cOpeCod='200374' and nPersoneria=1 and cOpeCtaDH='H'")
        cCtaDebe = rs!cCtaDebe
        
        lsMovNro = oDMov.GeneraMovNro(pdFecsis, psCodAge, psCodUsu)
        oDMov.InsertaMov lsMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
        lnMovNro = oDMov.GetnMovNro(lsMovNro)
        
        lnMovItem = lnMovItem + 1
        lnMovOrden = 0
        oDMov.InsertaMovCta lnMovNro, lnMovItem, cCtaDebe, pnMontoRetiro
        
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnMontoRetiro * -1
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIf, psTipoIF, psCtaIFCod
               
        oDMov.InsertaMovDoc lnMovNro, pnTipoDoc, psNroDoc, Format(pdFecsis, "yyyymmdd")
               
        If Mid(psOpecod, 3, 1) = "2" Then
            oDMov.GeneraMovME lnMovNro, lsMovNro
        End If
        oDMov.ActualizaSaldoMovimiento lsMovNro, "+"
        
        oDMov.InsertaSegTarjetaOpeBancos lnMovNro, Left(psPeriodo, 4), Right(psPeriodo, 2), psFechaPago, sTime, pnMontoRetiro, 1
        
        oDMov.CommitTrans
        bTrans = False
        
        'JUEZ 20150510 ******************************************
        Set oDSeg = New DSeguros
            oDSeg.ActualizaSegTarjetaAnulaDevPendiente "", "", True, True, lnMovNro
        Set oDSeg = Nothing
        'END JUEZ ***********************************************
        
        psMovNro = lsMovNro
        Exit Sub
ErrGraba:
    If bTrans Then oDMov.RollbackTrans
    psMovNro = ""
    Set oDMov = Nothing
    Err.Raise Err.Number, "", Err.Description
End Sub
Public Sub GrabarDepositoComisionSeguroTarjeta(ByVal pdFecsis As String, ByVal psCodAge As String, ByVal psCodUser As String, ByVal psOpecod As String, ByVal psPeriodo As String, _
                                               ByVal psFechaPago As String, ByVal psMovDesc As String, ByVal pnMontoDeposito As Double, ByVal psCtaDebe As String, ByVal psPersCodIf As String, _
                                               ByVal psTpoIF As String, ByVal psCtaIFCod As String, ByRef psMovNro As String)

Dim oDCon As New DConecta
Dim oDMov As New DMov
Dim oDSeg As New DSeguros
Dim bTrans As Boolean
Dim x As Integer, y As Integer
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lnMovItem As Integer, lnMovOrden As Integer
Dim cCtaHaber As String
Dim sTime As String
    
    oDCon.AbreConexion
    sTime = oDCon.GetHoraServer
    oDCon.CierraConexion
    
    On Error GoTo ErrGraba
        oDMov.BeginTrans
        bTrans = True
        cCtaHaber = "15170909"
        
        lsMovNro = oDMov.GeneraMovNro(pdFecsis, psCodAge, psCodUser)
        oDMov.InsertaMov lsMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
        lnMovNro = oDMov.GetnMovNro(lsMovNro)
        
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, pnMontoDeposito
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIf, psTpoIF, psCtaIFCod
        
        lnMovItem = lnMovItem + 1
        lnMovOrden = 0
        oDMov.InsertaMovCta lnMovNro, lnMovItem, cCtaHaber, pnMontoDeposito * -1
               
        If Mid(psOpecod, 3, 1) = "2" Then
            oDMov.GeneraMovME lnMovNro, lsMovNro
        End If
        oDMov.ActualizaSaldoMovimiento lsMovNro, "+"
        
        oDMov.InsertaSegTarjetaOpeBancos lnMovNro, Left(psPeriodo, 4), Right(psPeriodo, 2), psFechaPago, sTime, pnMontoDeposito, 2
        
        oDMov.CommitTrans
        bTrans = False
        psMovNro = lsMovNro
        Exit Sub
ErrGraba:
    If bTrans Then oDMov.RollbackTrans
    psMovNro = ""
    Set oDMov = Nothing
    Err.Raise Err.Number, "", Err.Description
End Sub
                                               
'END JUEZ *****************************************************************
'APRI20180130 ERS028-2017
Public Sub GrabarTransferenciaPrimaSeguros(ByVal pdFecProceso As String, ByVal psCodAge As String, ByVal psCodUsu As String, ByVal psOpecod As String, ByVal psPeriodo As String, _
                                         ByVal pnTipoDoc As Integer, ByVal psNroDoc As String, ByVal psMovDesc As String, ByVal pnMontoRetiro As Currency, _
                                         ByVal psCtaHaber As String, ByVal psPersCodIf As String, ByVal psTipoIF As String, ByVal psCtaIFCod As String, ByRef psMovNro As String, ByVal pnTipoSeg As Integer)
'Public Sub GrabarTransferenciaPrimaSeguros(ByVal pdFecsis As String, ByVal psCodAge As String, ByVal psCodUsu As String, ByVal psOpecod As String, ByVal psPeriodo As String, _
'                                         ByVal pnTipoDoc As Integer, ByVal psNroDoc As String,ByVal psFechaPago As String, ByVal psMovDesc As String, ByVal pnMontoRetiro As Currency, _
'                                         ByVal psCtaHaber As String, ByVal psPersCodIf As String, ByVal psTipoIF As String, ByVal psCtaIFCod As String, ByRef psMovNro As String, ByVal pnTipoSeg As Integer)

Dim oDCon As New DConecta
Dim oDMov As New DMov
Dim oDSeg As New DSeguros
Dim rs As ADODB.Recordset
Dim bTrans As Boolean
Dim x As Integer, y As Integer
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lnMovItem As Integer, lnMovOrden As Integer
Dim cCtaDebe As String
Dim sTime As String

    
    oDCon.AbreConexion
    sTime = oDCon.GetHoraServer
    oDCon.CierraConexion
    
    On Error GoTo ErrGraba
        oDMov.BeginTrans
        bTrans = True

        If pnTipoSeg = 1 Or pnTipoSeg = 2 Then '1:Tarjeta y 2:Sepelio
            Set rs = oDMov.Ejecutar("select cCtaDebe=replace(cCtaContCod,'M','1') from OpeCtaNeg where cOpeCod='" & IIf(pnTipoSeg = 1, "200374", "200380") & "' and nPersoneria=1 and cOpeCtaDH='H'")
        Else '3: Incendio y 4:Desgravamen , 5 : MYPE
            'Set rs = oDMov.Ejecutar("select cCtaDebe=replace(cCtaContCod,'M','" & Mid(psOpecod, 3, 1) & "') from OpeCtaNeg where cOpeCod='100202' and nPersoneria=1 and cOpeCtaDH='H' and nconcepto= " & IIf(pnTipoSeg = 3, 1231, 1217))
            'APRI20180919 ERS061-2018
            Set rs = oDMov.Ejecutar("select cCtaDebe=replace(cCtaContCod,'M','" & Mid(psOpecod, 3, 1) & "') from OpeCtaNeg where cOpeCod='100202' and nPersoneria=1 and cOpeCtaDH='H' and nconcepto= " & IIf(pnTipoSeg = 3, 1231, IIf(pnTipoSeg = 4, 1217, 1272)))
            'END APRI
        End If
        cCtaDebe = rs!cCtaDebe
        
        'lsMovNro = oDMov.GeneraMovNro(pdFecsis, psCodAge, psCodUsu)
        lsMovNro = oDMov.GeneraMovNro(pdFecProceso, psCodAge, psCodUsu)
        oDMov.InsertaMov lsMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
        lnMovNro = oDMov.GetnMovNro(lsMovNro)
        
        lnMovItem = lnMovItem + 1
        lnMovOrden = 0
        oDMov.InsertaMovCta lnMovNro, lnMovItem, cCtaDebe, pnMontoRetiro
        
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnMontoRetiro * -1
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIf, psTipoIF, psCtaIFCod
               
        'oDMov.InsertaMovDoc lnMovNro, pnTipoDoc, psNroDoc, Format(pdFecsis, "yyyymmdd")
        oDMov.InsertaMovDoc lnMovNro, pnTipoDoc, psNroDoc, Format(pdFecProceso, "yyyymmdd")
        
        If Mid(psOpecod, 3, 1) = "2" Then
            oDMov.GeneraMovME lnMovNro, lsMovNro
        End If
        oDMov.ActualizaSaldoMovimiento lsMovNro, "+"
        
        'oDMov.InsertaSegurosOpeBancos lnMovNro, Left(psPeriodo, 4), Right(psPeriodo, 2), psFechaPago, sTime, pnMontoRetiro, pnTipoSeg
        oDMov.InsertaSegurosOpeBancos lnMovNro, Left(psPeriodo, 4), Right(psPeriodo, 2), pdFecProceso, sTime, pnMontoRetiro, CInt(Mid(psOpecod, 3, 1)), pnTipoSeg
        
        oDMov.CommitTrans
        bTrans = False
        
        psMovNro = lsMovNro
        Exit Sub
ErrGraba:
    If bTrans Then oDMov.RollbackTrans
    psMovNro = ""
    Set oDMov = Nothing
    Err.Raise Err.Number, "", Err.Description
End Sub
'END APRI
'FRHU 20140812
Public Function GrabarDepositoXActivacionSegTarjeta(ByVal psPersCodIf As String, ByVal pdFechaDeposito As Date, ByVal pnMontoDeposito As Currency, ByVal pMatdepositos As Variant, ByVal pdFecsis As Date, ByVal pnMatTotalOpe As Integer, ByVal psCodAge As String, ByVal psCodUsu As String, _
                                               ByVal psOpecod As String, ByVal psCtaDebe As String, ByVal psObjetoIF As String, ByVal psCtaHaber As String, ByVal pMatPendiente As Variant, ByVal pbPendiente As Boolean) As String
    Dim oDMov As New DMov
    Dim bTrans As Boolean
    Dim x As Integer
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lnMovItem As Integer, lnMovOrden As Integer
    'psObjetoIF = Cuenta de la Insitución Financiera
    On Error GoTo ErrGraba
        oDMov.BeginTrans
        bTrans = True
       
        lsMovNro = oDMov.GeneraMovNro(pdFecsis, psCodAge, psCodUsu)
        oDMov.InsertaMov lsMovNro, psOpecod, "DEPOSITO PARA ACTIVACION DE SEGURO", gMovEstContabMovContable, gMovFlagVigente
        lnMovNro = oDMov.GetnMovNro(lsMovNro)
        oDMov.InsertaMovCont lnMovNro, pnMontoDeposito, "0", "0"
        'Guardamos la cuenta en el Debe
        lnMovItem = 0: lnMovOrden = 0
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, pnMontoDeposito 'pnTotalImporte
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psObjetoIF, 4, 13), Mid(psObjetoIF, 1, 2), Mid(psObjetoIF, 18, 10)
        'Guardamos en las Tablas de Seguro de Tarjeta
        For x = 1 To pnMatTotalOpe
            Call oDMov.InsertaSegTarjetaEstado(pMatdepositos(x, 1), pdFecsis, 1001, pMatdepositos(x, 4), pMatdepositos(x, 3), lsMovNro)
            'pMatdepositos(x, 1) = Solicitud, pMatdepositos(x, 4) = Glosa, pMatdepositos(x, 3) = Monto,pMatdepositos(x, 2) = CuentaAhorro
            Call oDMov.ActualizarSegTarjetaSolicitudAproAcep(pMatdepositos(x, 1), lsMovNro, 1001, pMatdepositos(x, 2), CCur(pMatdepositos(x, 3)))
'            lsFiltro = ""
'            lsFiltro = oCont.GetFiltroObjetos(ObjProductosCMACT, psCtaHaber, rs!cObjetoCod, False)
'            lsFiltro = lsFiltro & oCont.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaHaber, rs!AreaCod & rs!AgeCod, False)
        Next
        'Guardamos la cuenta en el Haber
        lnMovItem = lnMovItem + 1
        lnMovOrden = 0
        'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, rs!Importe * -1
        'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, CCur(pMatPendiente(3)) * -1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnMontoDeposito * -1
        'Registra la pendiente en la tabla MovPendientesRend, como siempre va a ser pendiente se quita if onde validad si la cuenta es una pendiente(Caja General)
        'Dim lsCtaCaracterOut As String
        'If oDMov.CuentaEsPendiente(psCtaHaber, lsCtaCaracterOut, IIf(pbDeposito, "A", "D")) Then
            'If (lsCtaCaracterOut = "A" And pbDeposito) Or (lsCtaCaracterOut = "D" And Not pbDeposito) Then
                oDMov.InsertaMovPendientesRend lnMovNro, psCtaHaber, pnMontoDeposito
            'End If
        'End If
        
        If Mid(psOpecod, 3, 1) = "2" Then
            oDMov.GeneraMovME lnMovNro, lsMovNro
        End If
        oDMov.ActualizaSaldoMovimiento lsMovNro, "+"
        
        oDMov.CommitTrans
        bTrans = False
        GrabarDepositoXActivacionSegTarjeta = lsMovNro
        Exit Function
ErrGraba:
    If bTrans Then oDMov.RollbackTrans
    Set oDMov = Nothing
    Err.Raise Err.Number, "", Err.Description
End Function
Public Function GrabarDepositoXActivacionSegTarjetaFechaAnterior(ByVal pMatdepositos As Variant, ByVal pdFecsis As Date, ByVal pnMatTotalOpe As Integer, ByVal psCodAge As String, ByVal psCodUsu As String, ByVal psOpecod As String, ByVal pMatPendiente As Variant) As Boolean
    Dim oDMov As New DMov
    Dim bTrans As Boolean
    Dim x As Integer
    Dim lsMovNro As String
    Dim lnMovNro As Long
    
    GrabarDepositoXActivacionSegTarjetaFechaAnterior = False
    On Error GoTo ErrGraba
        oDMov.BeginTrans
        bTrans = True
        
        lsMovNro = oDMov.GeneraMovNro(pdFecsis, psCodAge, psCodUsu)
        oDMov.InsertaMov lsMovNro, psOpecod, "DEPOSITO PARA ACTIVACION DE SEGURO: FECHA ANTERIOR - REFERENCIA(SIN ASIENTO)", gMovEstContabMovContable, gMovFlagVigente
        lnMovNro = oDMov.GetnMovNro(lsMovNro)
        oDMov.InsertaMovRef lnMovNro, pMatPendiente(0)
        
        For x = 1 To pnMatTotalOpe
            Call oDMov.InsertaSegTarjetaEstado(pMatdepositos(x, 1), pdFecsis, 1001, pMatdepositos(x, 4), pMatdepositos(x, 3), lsMovNro)
            'pMatdepositos(x, 1) = Solicitud, pMatdepositos(x, 4) = Glosa, pMatdepositos(x, 3) = Monto,pMatdepositos(x, 2) = CuentaAhorro
            Call oDMov.ActualizarSegTarjetaSolicitudAproAcep(pMatdepositos(x, 1), lsMovNro, 1001, pMatdepositos(x, 2), CCur(pMatdepositos(x, 3)))
'            lsFiltro = ""
'            lsFiltro = oCont.GetFiltroObjetos(ObjProductosCMACT, psCtaHaber, rs!cObjetoCod, False)
'            lsFiltro = lsFiltro & oCont.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaHaber, rs!AreaCod & rs!AgeCod, False)
        Next
        oDMov.CommitTrans
        bTrans = False
        GrabarDepositoXActivacionSegTarjetaFechaAnterior = True
        Exit Function
ErrGraba:
    If bTrans Then oDMov.RollbackTrans
    Set oDMov = Nothing
    Err.Raise Err.Number, "", Err.Description
End Function
'FIN FRHU 20140812
'PASIERS1362014
Public Sub GrabarRetiroPagoSeguroDesgravIncen(ByVal pdFecsis As String, ByVal psCodAge As String, ByVal psCodUsu As String, ByVal psOpecod As String, ByVal psPeriodo As String, _
                                                                        ByVal pnTipoDoc As Integer, ByVal psNroDoc As String, ByVal psFechaPago As String, ByVal psMovDesc As String, ByRef TSegAge() As TSegAgencia, _
                                                                        ByRef TRetPagSegDet() As TRetPagoSegDet, ByVal pnMontoRetiro As Currency, ByVal psCtaHaber As String, ByVal psPersCodIf As String, ByVal psTipoIF As String, ByVal psCtaIFCod As String, _
                                                                        ByRef psMovNro As String)
Dim oDCon As New DConecta
Dim oDMov As New DMov
Dim oDSeg As New DSeguros
Dim rs As ADODB.Recordset
Dim bTrans As Boolean
Dim lsMovNro As String
Dim lnMovNro As Long
Dim I, lnMovItem, lnMovOrden As Integer

    oDCon.AbreConexion
    oDCon.CierraConexion
    lnMovItem = 0
    lnMovOrden = 0
    On Error GoTo ErrGraba
        oDMov.BeginTrans
        bTrans = True
        Set rs = oDMov.Ejecutar("select cCtaDebe=replace(cCtaContCod,'M','" & Mid(psOpecod, 3, 1) & "') from OpeCtaNeg where cOpeCod='100202' and nPersoneria=1 and cOpeCtaDH='H' and nconcepto= " & IIf(Mid(psOpecod, 4, 3) = Mid(OpeCGOtrosOpeRetPagSeguroDesgravamenMN, 4, 3), 1217, 1210))
        cCtaDebe = rs!cCtaDebe
        
        lsMovNro = oDMov.GeneraMovNro(pdFecsis, psCodAge, psCodUsu)
        oDMov.InsertaMov lsMovNro, psOpecod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
        lnMovNro = oDMov.GetnMovNro(lsMovNro)
        For I = 1 To UBound(TSegAge)
            If TSegAge(I).nMonto > 0 Then
                lnMovItem = lnMovItem + 1
                oDMov.InsertaMovCta lnMovNro, lnMovItem, Replace(rs!cCtaDebe, "AG", TSegAge(I).sAgencia), TSegAge(I).nMonto
            End If
        Next
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnMontoRetiro * -1
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIf, psTipoIF, psCtaIFCod
        oDMov.InsertaMovDoc lnMovNro, pnTipoDoc, psNroDoc, Format(pdFecsis, "yyyymmdd")
        If Mid(psOpecod, 3, 1) = "2" Then
                oDMov.GeneraMovME lnMovNro, lsMovNro
        End If
        oDMov.InsertaMovRetPagoSeguro lnMovNro, Left(psPeriodo, 4), Right(psPeriodo, 2), psFechaPago, pnMontoRetiro, CInt(Mid(psOpecod, 3, 1))
        For I = 1 To UBound(TRetPagSegDet)
            If TRetPagSegDet(I).nMonto > 0 Then
                oDMov.InsertaMovRetPagoSeguroDet lnMovNro, Left(psPeriodo, 4), Right(psPeriodo, 2), TRetPagSegDet(I).sCtaCod, TRetPagSegDet(I).nCuota, TRetPagSegDet(I).nNroCalen, TRetPagSegDet(I).nMonto, TRetPagSegDet(I).sAgeCod
            End If
        Next
        oDMov.ActualizaSaldoMovimiento lsMovNro, "+"
        oDMov.CommitTrans
        bTrans = False
        psMovNro = lsMovNro
        Exit Sub
ErrGraba:
    If bTrans Then oDMov.RollbackTrans
    psMovNro = ""
    Set oDMov = Nothing
    Err.Raise Err.Number, "", Err.Description
End Sub
Public Function GetMovRetPagoxExtorno(ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal psOpecod As String) As ADODB.Recordset
    Dim oDSeg As New DSeguros
    Set GetMovRetPagoxExtorno = oDSeg.GetMovRetPagoSeguroxExtorno(pdFecIni, pdFecFin, psOpecod)
    Set oDSeg = Nothing
End Function
'end PASI

