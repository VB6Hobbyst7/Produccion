VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOpeReimprVoucher 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reimpresión de Voucher"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   Icon            =   "frmOpeReimprVoucher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9945
      Begin MSDataGridLib.DataGrid DBGrdVoucher 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   4
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "nMovNro"
            Caption         =   "#"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cCliente"
            Caption         =   "Cliente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cOpeCod"
            Caption         =   "cOpeCod"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cTpoOpe"
            Caption         =   "Operación"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "cMoneda"
            Caption         =   "Mon."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "nMonto"
            Caption         =   "Monto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "dHora"
            Caption         =   "Hora"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "hh:mm:ss AMPM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   4
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "cCheque"
            Caption         =   "Cheque"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "nNroReu"
            Caption         =   "Reu"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            SizeMode        =   1
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            Size            =   800
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   3495.118
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtGlosa 
         Height          =   540
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   3750
         Width           =   7575
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8640
         TabIndex        =   3
         Top             =   3960
         Width           =   1170
      End
      Begin VB.CommandButton cmdReimprimir 
         Caption         =   "Reimprimir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8640
         TabIndex        =   2
         Top             =   3480
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Glosa"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3480
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmOpeReimprVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmOpeReimprVoucher
'** Descripción : Lista de operaciones disponibles para reimprimir su boleta creado segun RFC079-2012
'** Creación : JUEZ, 20120723 04:00:00 PM
'********************************************************************

Option Explicit

Private Sub cmdReimprimir_Click()

    Dim oImprDesembPagoCred As COMNCredito.NCOMCredDoc
    Dim oDatosVoucher As COMDCredito.DCOMCredDoc
    Dim oImprPagJud As COMNColocRec.NCOMColRecImpre
    Dim oImprCF As COMNCartaFianza.NCOMCartaFianzaImpre
    Dim rsDatos As ADODB.Recordset
    Dim ImprDesemb As Variant
    Dim oDatosVoucherCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim oImprCaptac As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oImprPig As COMNColoCPig.NCOMColPImpre
    Dim oImpSeg As COMNCaptaGenerales.NCOMSeguros 'JUEZ 20140711
    Dim sImpreBoleta As String
    Dim oImp As COMFunciones.FCOMVarImpresion 'PASI20151022 ERS0692015
    Dim oImpOpEsp As COMNCaptaGenerales.NCOMCaptaImpresion 'PASI20151103 ERS0692015
    Dim oImpCapServ As COMNCaptaServicios.NCOMCaptaServicios 'PASI20151103 ERS0692015
    Dim bExisteBoleta As Boolean 'PASI20151022 ERS0692015
    
    Set oImprDesembPagoCred = New COMNCredito.NCOMCredDoc
    Set oDatosVoucher = New COMDCredito.DCOMCredDoc
    Set oImprPagJud = New COMNColocRec.NCOMColRecImpre
    Set oImprCF = New COMNCartaFianza.NCOMCartaFianzaImpre
    Set oDatosVoucherCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set oImprCaptac = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oImprPig = New COMNColoCPig.NCOMColPImpre
    Set oImpSeg = New COMNCaptaGenerales.NCOMSeguros 'JUEZ 20140711
    Set oImp = New COMFunciones.FCOMVarImpresion  'PASI20151022 ERS0692015
    Set oImpCapServ = New COMNCaptaServicios.NCOMCaptaServicios  'PASI20151022 ERS0692015
    Set oImpOpEsp = New COMNCaptaGenerales.NCOMCaptaImpresion 'PASI20151103 ERS0692015
    
    Dim bResultadoVisto As Boolean
    Dim oVisto As frmVistoElectronico
    
    Dim pnMovNro As Long
    Dim psOpeCod As String
    pnMovNro = DBGrdVoucher.Columns(0)
    psOpeCod = DBGrdVoucher.Columns(2)
    
    '***JGPA20190607
    Dim pnNroREU As Long
    Dim pnResponseREU As Integer
    Dim pnImporte As Currency
    Dim rsDatosIntReu As ADODB.Recordset
    Dim oDatosIntReu As COMDCredito.DCOMCredDoc
    Dim MatInterviniestesReu As Variant
    Dim MatRealiza As Variant
    Dim MatOrdena As Variant
    Dim MatBeneficia As Variant
    Dim frmML As frmMovLavDinero
    Dim sImpreBoletaReu As String
    Dim ir As Integer
    Dim ListaPersonasRealizan() As PersonaLavado
    Dim ListaPersonasOrdenan() As PersonaLavado
    Dim ListaPersonasBenefician() As PersonaLavado
    
    Set oDatosIntReu = New COMDCredito.DCOMCredDoc
    Set frmML = New frmMovLavDinero
    If Me.DBGrdVoucher.Columns(8) <> "" Then
        pnNroREU = CLng(Me.DBGrdVoucher.Columns(8))
        pnImporte = CCur(Me.DBGrdVoucher.Columns(5))
    End If
    '***JGPA20190607
    
    '******* Variables Apertura Captaciones *******
    Dim bReImp As Boolean, bProd As Boolean
    Dim sTipDep As String
    Dim sModDep As String, sTipApe As String
    Dim sNomTit As String, sNroDoc As String
    Dim nTipoPag As Integer, nDocumento As Integer
    Dim nSaldoCnt As Double, nSaldoDisp As Double
    '**********************************************
    '******* Variables Deposito Captaciones *******
    Dim sMsgOpe As String
    Dim nDiasTranscurridos As Integer, nIntGanado As Double
    '**********************************************
    'PASI20151028 ERS0692015
    Dim dUltRetInt As Date
    Dim dFechaRenovacionPFM As Date
    Dim nDias As Integer
    Dim sMenProx As String
    Dim dFechaProx As Date
    Dim lsMenProx2 As String
    'end PASI
    '**********************************************
    
    '************PASI20151106********************
    '*****Variables para Otras Operaciones*******
    Dim sImpDJ As String
    '********************************************
    
    If Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Es necesario que escriba la glosa", vbInformation, "Aviso"
        txtGlosa.SetFocus
        Exit Sub
    End If
    
    Select Case psOpeCod
    
    'DESEMBOLSOS
    Case gCredDesembEfec, gCredDesembCheque, gCredDesembCtaNueva, gCredDesembCtaNuevaDOA, gCredDesembCtaExist, gCredDesembCtaExistDOA
    
        Set rsDatos = oDatosVoucher.ObtieneDatosReimprDesemb(pnMovNro)
        
        Select Case psOpeCod
        
        Case gCredDesembEfec, gCredDesembCheque
            ImprDesemb = oImprDesembPagoCred.ImprimeDocumentosDesembolso(rsDatos!cCtaCod, pnMovNro, gsNomAge, _
                    gsCodUser, gdFecSis, , , , , , , , , , , , , , , , gsNomCmac, gsInstCmac, gsCodCMAC, rsDatos!nMontoITF, rsDatos!nMontoPoliza)
        
        Case gCredDesembCtaNueva, gCredDesembCtaNuevaDOA
            ImprDesemb = oImprDesembPagoCred.ImprimeDocumentosDesembolso(rsDatos!cCtaCod, pnMovNro, gsNomAge, _
                    gsCodUser, gdFecSis, 1, rsDatos!sCtaAho, rsDatos!nMontoCol, 0, 0, 0, rsDatos!nMontoCol, _
                    0, 0, rsDatos!nSaldoDisponible, rsDatos!nSaldoContable, 0, 0, 0, 0, gsNomCmac, gsInstCmac, _
                    gsCodCMAC, rsDatos!nMontoITF, rsDatos!nMontoPoliza)
    
        Case gCredDesembCtaExist, gCredDesembCtaExistDOA
            ImprDesemb = oImprDesembPagoCred.ImprimeDocumentosDesembolso(rsDatos!cCtaCod, pnMovNro, gsNomAge, _
                    gsCodUser, gdFecSis, 2, rsDatos!sCtaAho, 0, 0, 0, 0, rsDatos!nMontoCap, 0, 0, _
                    rsDatos!nSaldoDisponible - rsDatos!nMontoITF, rsDatos!nSaldoContable - rsDatos!nMontoITF, 0, _
                    0, 0, 0, gsNomCmac, gsInstCmac, gsCodCMAC, rsDatos!nMontoITF, rsDatos!nMontoPoliza)
        
        End Select
        
    Dim i As Integer
    For i = 0 To UBound(ImprDesemb) - 1
        'If ImprDesemb(i, 0) = "1"  Then 'Para Boleta 'comento 'JOEP20190412
        If ImprDesemb(i, 0) = "1" And ImprDesemb(i, 1) <> "" Then 'JOEP20190412 'Para Boleta
            sImpreBoleta = sImpreBoleta & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
            sImpreBoleta = sImpreBoleta & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
            sImpreBoleta = sImpreBoleta & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
            sImpreBoleta = sImpreBoleta & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
            sImpreBoleta = sImpreBoleta & Chr$(27) + Chr$(18)  ' cancela condensada
            sImpreBoleta = sImpreBoleta & Chr$(27) + Chr$(72)  ' desactiva negrita
            sImpreBoleta = sImpreBoleta & ImprDesemb(i, 1)
    'JOEP20190412
        ElseIf ImprDesemb(i, 0) = "3" And ImprDesemb(i, 1) <> "" Then 'Para Boleta
            sImpreBoleta = sImpreBoleta & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
            sImpreBoleta = sImpreBoleta & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
            sImpreBoleta = sImpreBoleta & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
            sImpreBoleta = sImpreBoleta & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
            sImpreBoleta = sImpreBoleta & Chr$(27) + Chr$(18)  ' cancela condensada
            sImpreBoleta = sImpreBoleta & Chr$(27) + Chr$(72)  ' desactiva negrita
            sImpreBoleta = sImpreBoleta & ImprDesemb(i, 1)
        End If
    'JOEP20190412
    Next i
    
    Case "120201", "120202", "120204" 'PASI20150929 ERS0692015
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpDesembPigno(pnMovNro)
        sImpreBoleta = oImprPig.nPrintReciboDesembolso(rsDatos!dVenc, rsDatos!cCtaCod, rsDatos!nSaldoCap, rsDatos!FechaHoraPrend, _
                        rsDatos!nMontoEntregar, rsDatos!nInteresComp, gsNomAge, gsCodUser, rsDatos!nITF, gImpresora, IIf(psOpeCod = "120201", False, True), IIf(psOpeCod = "120201", "", IIf(psOpeCod = "120202", rsDatos!cCtaAho, "")), IIf(psOpeCod = "120201", "", IIf(psOpeCod = "120202", "", rsDatos!cCtaAho)))
    
    'PAGOS CREDITOS: PAGO NORMAL
    Case gCredPagNorNorEfec, gCredPagNorNorCC, gCredPagNorNorCh, gCredPagNorVenEfec 'PASI20151218 agrego:gCredPagNorVenEfec
        Set rsDatos = oDatosVoucher.ObtieneDatosReimprPagoCred(pnMovNro)
        
        sImpreBoleta = oImprDesembPagoCred.ImprimeBoleta(rsDatos!cCtaCod, rsDatos!cPersNombre, gsNomAge, rsDatos!cmoneda, _
                        rsDatos!nCuotasPag, gdFecSis, rsDatos!dHora, rsDatos!nTransacc + 1, "", _
                        rsDatos!nCapitalPag, rsDatos!nIntCompPag, rsDatos!nIntCompVencPag, rsDatos!nIntMorPag, rsDatos!nGastoPag, _
                        rsDatos!nIntGraciaPag, rsDatos!nIntSuspPag + rsDatos!nIntReprogPag, rsDatos!nSaldoCap, rsDatos!dNroProxCuota, _
                        gsCodUser, sLpt, gsInstCmac, IIf(rsDatos!bCheque = 1, True, False), rsDatos!cNumCheque, gsCodCMAC, rsDatos!nITF, _
                        rsDatos!nIntDesagioPag, False, , , gbImpTMU)
    
    '100212 PASI20150926 ERS0692015
    Case gCredPagLeasingCU
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpPagoLeasing(pnMovNro)
        sImpreBoleta = oImprDesembPagoCred.ImprimeBoleta(rsDatos!cCtaCod, rsDatos!cPersNombre, gsNomAge, rsDatos!cmoneda, _
                        rsDatos!nCuotasPag, gdFecSis, rsDatos!dHora, rsDatos!nTransacc + 1, "", _
                        rsDatos!nCapitalPag, rsDatos!nIntCompPag, _
                        0, _
                        rsDatos!nIntMorPag, rsDatos!nGastoPag, _
                        0, _
                        0, _
                        rsDatos!nSaldoCap, rsDatos!dNroProxCuota, _
                        gsCodUser, sLpt, gsInstCmac, IIf(rsDatos!bCheque = 1, True, False), rsDatos!cNumCheque, gsCodCMAC, rsDatos!nITF, _
                        rsDatos!nIntDesagioPag, False, , , gbImpTMU)
    
    '100218 PASI20150928 ERS0692015
    Case gCredPagHonramiento
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpPagoxHonramiento(pnMovNro)
        sImpreBoleta = oImprDesembPagoCred.ImprimePagoHonramiento(rsDatos!cCtaCod, rsDatos!cPersNombre, gsNomAge, gdFecSis, rsDatos!nMonto, gsInstCmac, rsDatos!nITF, gsCodCMAC, gsCodUser, gbImpTMU)
                             
    'PAGO JUDICIAL
    Case gColRecOpePagJudCDEfe, gColRecOpePagJudCDEfe, gColRecOpePagCastEfe, gColRecOpePagJudCDChq, gColRecOpePagJudCDChq, gColRecOpePagCastChq
        
        Set rsDatos = oDatosVoucher.ObtieneDatosReimprPagoJud(pnMovNro)
        
        sImpreBoleta = oImprPagJud.nPrintReciboPagoCredRecup(gsNomAge, fgFechaHoraGrab(rsDatos!cMovNro), rsDatos!cCtaCod, _
            rsDatos!cPersNombre, rsDatos!nMonto, gsCodUser, " ", rsDatos!nITF, gImpresora, gbImpTMU, psOpeCod)
            'WIOR 20150615 AGREGO psOpeCod
            
    'CARTA FIANZA
    Case gColCFOpeComisEfe, "322001" 'PASI20150930 ERS0692015 agrego 322001
    
        Set rsDatos = oDatosVoucher.ObtieneDatosReimprCFComision(pnMovNro)
            
        sImpreBoleta = oImprCF.nPrintReciboCFComision(gsNomAge, fgFechaHoraGrab(rsDatos!cMovNro), rsDatos!cCtaCod, rsDatos!cNomPersona, _
                        rsDatos!cNomAcreedor, rsDatos!nMontoApr, rsDatos!dVencApr, rsDatos!nMonto, gsCodUser, "", gsCodCMAC, rsDatos!nITF, gImpresora, gbImpTMU)

    'AHORROS / PLAZO FIJO / CTS
    '   APERTURAS
    Case gAhoApeEfec, gAhoApeChq, gAhoApeTransf, _
         gPFApeEfec, gPFApeChq, gPFApeTransf, _
         gCTSApeEfec, gCTSApeChq, gCTSApeTransf
         
        Set rsDatos = oDatosVoucherCap.ObtieneDatosReimprAperturaCaptac(pnMovNro)
            
        sTipDep = rsDatos!cmoneda
        nSaldoCnt = rsDatos!nSaldoContable
        nSaldoDisp = rsDatos!nSaldoDisponible
            
        If rsDatos!bCheque = 1 Then
            sModDep = "Depósito Cheque"
            sNroDoc = rsDatos!cNumCheque
            nDocumento = TpoDocCheque
        Else
            If psOpeCod = gAhoApeTransf Or psOpeCod = gPFApeTransf Or psOpeCod = gCTSApeTransf Then
                sModDep = "Depósito Transferencia"
            Else
                sModDep = "Depósito Efectivo"
            End If
        End If
        Select Case rsDatos!nProducto
            Case gCapAhorros
                bProd = gITF.gbITFAsumidoAho
                If rsDatos!bOrdPag = 1 Then
                    sTipApe = "APERTURA AHORROS CON OP"
                Else
                    sTipApe = "APERTURA AHORROS"
                End If
            Case gCapPlazoFijo
                bProd = gITF.gbITFAsumidoPF
                sTipApe = "APERTURA PLAZO FIJO"
            Case gCapCTS
                bProd = True
                sTipApe = "APERTURA CTS"
        End Select
        bReImp = False
        If gbITFAplica And rsDatos!nProducto <> gCapCTS And bProd = False Then
            If Trim(Left(rsDatos!nValoresITF, 6)) = "990101" Then
                nSaldoCnt = nSaldoCnt - CDbl(Trim(Right(rsDatos!nValoresITF, 6)))
                nSaldoDisp = nSaldoDisp - CDbl(Trim(Right(rsDatos!nValoresITF, 6)))
            End If
        End If

        sNomTit = ImpreCarEsp(clsMant.GetNombreTitulares(rsDatos!cCtaCod))
        
        Set clsMant = Nothing
        
        bReImp = False
        '****** validacion CTS *****
        If rsDatos!nProducto = gCapCTS Then
            'If nSaldRetiro = 0 Then
                nSaldoDisp = (nSaldoDisp * rsDatos!nPorc) '/ 100
            'End If
        End If
        '***************************
        
        'Obtener el Tipo de pago de ITF
        If Trim(Left(rsDatos!nValoresITF, 6)) = "990102" Then
            nTipoPag = 1
        Else
            nTipoPag = 2
        End If
        
            If rsDatos!bCheque = 1 Then
                If nDocumento = TpoDocCheque Then
                    sImpreBoleta = oImprCaptac.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, psOpeCod, rsDatos!nMonto, sNomTit, rsDatos!cCtaCod, Format$(rsDatos!dFecValorizacion, "dd/mm/yyyy"), nSaldoDisp, 0, "Fecha Valor", 1, nSaldoCnt, , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, False, gsCodCMAC, , , , , , , , , , , , True, nTipoPag, CDbl(Trim(Right(rsDatos!nValoresITF, 6))), True, , , gbImpTMU)
                ElseIf nDocumento = TpoDocNotaAbono Then
                    sImpreBoleta = oImprCaptac.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, psOpeCod, rsDatos!nMonto, sNomTit, rsDatos!cCtaCod, "", nSaldoDisp, 0, "", 1, nSaldoCnt, , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, , , , , , , , , , , , True, nTipoPag, CDbl(Trim(Right(rsDatos!nValoresITF, 6))), True, , , gbImpTMU)
                End If
            Else
                sImpreBoleta = oImprCaptac.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep), psOpeCod, rsDatos!nMonto, sNomTit, rsDatos!cCtaCod, "", nSaldoDisp, 0, "", 1, nSaldoCnt, , , , , , , , , , gdFecSis, Trim(gsNomAge), gsCodUser, sLpt, False, gsCodCMAC, , , , , , , , , , , , True, nTipoPag, CDbl(Trim(Right(rsDatos!nValoresITF, 6))), True, , , gbImpTMU)
            End If
            
    '200104 PASI20151022 ERS0692015
    Case gAhoApeLoteEfec
        Dim sImpreBoletaCab As String
        Dim sImpreBoletaDet As String
        Dim rsDatosCab As ADODB.Recordset
        Dim rsDatosDet As ADODB.Recordset

        Set rsDatosCab = oDatosVoucher.ObtieneDatosReimpApeLoteEfectivoCab(pnMovNro)
        sImpreBoletaCab = oImprCaptac.ImprimeBoletaRes(rsDatosCab!cTipApe, ImpreCarEsp(rsDatosCab!cModDep), Trim(rsDatosCab!nOperacion), _
        rsDatosCab!nMonto, rsDatosCab!cNomTit, CStr(rsDatosCab!NroCtas), "", _
        rsDatosCab!nSaldoDisp, 0, "", 1, _
        rsDatosCab!nSaldoCont, , , , , , , , , , _
        gdFecSis, Trim(gsNomAge), gsCodUser, sLpt, False, , , , True, , , , , , , , , True, rsDatosCab!nTpoPago, 0, _
        True, , , , , , , , , , , , , , rsDatosCab!nMovNro) & oImp.gPrnSaltoLinea

        Set rsDatosDet = oDatosVoucher.ObtieneDatosReimpApeLoteEfectivodet(pnMovNro)
        Do While Not rsDatosDet.EOF
            sImpreBoletaDet = sImpreBoletaDet & oImprCaptac.ImprimeBoleta(rsDatosDet!cTipApe, ImpreCarEsp(rsDatosDet!cModDep), Trim(rsDatosDet!nOperacion), _
            rsDatosDet!nMonto, rsDatosDet!cNomTit, rsDatosDet!cCtaCod, "", _
            rsDatosDet!nSaldoDisp, 0, "", 1, _
            rsDatosDet!nSaldoCont, , , , , , , , , , _
            gdFecSis, Trim(gsNomAge), gsCodUser, sLpt, False, , , , True, , , , , , , , , True, rsDatosDet!ntpopag, rsDatosDet!nITF, _
            True, , , , , , , , , , , , , , rsDatosDet!nMovNro) & oImp.gPrnSaltoLinea
            rsDatosDet.MoveNext
        Loop
                    
    '220104 PASI20151030 ERS0692015
    Case gCTSApeLoteEfec
        Dim sImpreBoletaCTSCab As String
        Dim sImpreBoletaCTSDet As String
        Dim rsDatosCTSCab As ADODB.Recordset
        Dim rsDatosCTSDet As ADODB.Recordset
        
        Set rsDatosCTSCab = oDatosVoucher.ObtieneDatosReimpApeCTSLoteEfectivoCab(pnMovNro)
        sImpreBoletaCTSCab = oImprCaptac.ImprimeBoletaRes(rsDatosCTSCab!cTipApe, ImpreCarEsp(rsDatosCTSCab!cModDep), Trim(rsDatosCTSCab!nOperacion), _
        rsDatosCTSCab!nMonto, rsDatosCTSCab!cNomTit, CStr(rsDatosCTSCab!NroCtas), "", _
        rsDatosCTSCab!nSaldoDisp, 0, "", 1, _
        rsDatosCTSCab!nSaldoCont, , , , , , , , , , _
        gdFecSis, Trim(gsNomAge), gsCodUser, sLpt, False, , , , True, , , , , , , , , True, rsDatosCTSCab!nTpoPago, 0, _
        True, , , , , , , , , , , , , , rsDatosCTSCab!nMovNro) & oImp.gPrnSaltoLinea

        Set rsDatosCTSDet = oDatosVoucher.ObtieneDatosReimpApeCTSLoteEfectivoDet(pnMovNro)
        Do While Not rsDatosCTSDet.EOF
            sImpreBoletaCTSDet = sImpreBoletaCTSDet & oImprCaptac.ImprimeBoleta(rsDatosCTSDet!cTipApe, ImpreCarEsp(rsDatosCTSDet!cModDep), Trim(rsDatosCTSDet!nOperacion), _
            rsDatosCTSDet!nMonto, rsDatosCTSDet!cNomTit, rsDatosCTSDet!cCtaCod, "", _
            rsDatosCTSDet!nSaldoDisp, 0, "", 1, _
            rsDatosCTSDet!nSaldoCont, , , , , , , , , , _
            gdFecSis, Trim(gsNomAge), gsCodUser, sLpt, False, , , , True, , , , , , , , , True, rsDatosCTSDet!ntpopag, rsDatosCTSDet!nITF, _
            True, , , , , , , , , , , , , , rsDatosCTSDet!nMovNro) & oImp.gPrnSaltoLinea
            rsDatosCTSDet.MoveNext
        Loop
                
    '200203,200209,200245  PASI20151023 ERS0692015
    Case gAhoDepTransf, gAhoDepOtrosIngRRHH, "200245"
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpDepositoCaptacII(pnMovNro)
        If psOpeCod = gAhoDepTransf Then
            sMsgOpe = "Depósito Transferencia"
        ElseIf psOpeCod = gAhoDepOtrosIngRRHH Then
            sMsgOpe = "Otros Conceptos RRHH"
        ElseIf psOpeCod = "200245" Then
            sMsgOpe = " Abono Cta. Haberes Transf."
        End If
        nDiasTranscurridos = DateDiff("d", rsDatos!dUltCierre, gdFecSis) - 1
        If nDiasTranscurridos < 0 Then
            nDiasTranscurridos = 0
        End If
        nIntGanado = oDatosVoucherCap.GetInteres(rsDatos!nsaldoant, rsDatos!nTasaInteres, nDiasTranscurridos, TpoCalcIntSimple)
        oDatosVoucherCap.IniciaImpresora gImpresora
        sImpreBoleta = oDatosVoucherCap.ImprimeBoleta(rsDatos!cTipApe, ImpreCarEsp(sMsgOpe), rsDatos!cOpeCod, CStr(Trim(rsDatos!nMonto)), ImpreCarEsp(clsMant.GetNombreTitulares(rsDatos!cCtaCod)), rsDatos!cCtaCod, "", rsDatos!nSaldoDisp, nIntGanado, "", rsDatos!nExtracto, rsDatos!nSaldoCont, True, , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, rsDatos!cAgeCod, , False, CStr(rsDatos!cOpeCod), , , , , , , , False, rsDatos!ntpopag, CDbl(rsDatos!nITF), True, rsDatos!ComiDepOtraAge, , gbImpTMU, rsDatos!cNumTarjeta, False, , , , , , , , , , , rsDatos!cPersCodEcoTaxi, rsDatos!nLogEcotaxi)
            
    '302002 PASI0151111 ERS0692015
    Case "302002"
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpNotaAbono(pnMovNro)
        nDiasTranscurridos = DateDiff("d", rsDatos!dUltCierre, gdFecSis) - 1
        If nDiasTranscurridos < 0 Then
            nDiasTranscurridos = 0
        End If
        nIntGanado = oDatosVoucherCap.GetInteres(rsDatos!nSaldoDisp - rsDatos!nMonto, rsDatos!nTasaInteres, nDiasTranscurridos, TpoCalcIntSimple)
        sImpreBoleta = oDatosVoucherCap.ImprimeBoleta(rsDatos!cTipApe, ImpreCarEsp(rsDatos!cMsg), rsDatos!cOpeCod, CStr(Trim(rsDatos!nMonto)), ImpreCarEsp(clsMant.GetNombreTitulares(rsDatos!cCtaCod)), rsDatos!cCtaCod, "", rsDatos!nSaldoDisp, nIntGanado, "", rsDatos!nExtracto, rsDatos!nSaldoCont, True, , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, rsDatos!cAgeCod, , False, CStr(rsDatos!cOpeCod), , , , , , , , False, rsDatos!ntpopag, CDbl(rsDatos!nITF), True, rsDatos!ComiDepOtraAge, , gbImpTMU, rsDatos!cNumTarjeta, False, , , , , , , , , , , rsDatos!cPersCodEcoTaxi, rsDatos!nLogEcotaxi)
                
    '310101,310301 PASI0151111 ERS0692015
    Case "310101", "310301"
        Dim sDest() As String
        ReDim sDest(4)
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpGiro(pnMovNro)
        sDest(0) = rsDatos!cDestinatario
        sDest(1) = ""
        sDest(2) = ""
        sDest(3) = ""
        
        If psOpeCod = "310101" Then
            'sImpreBoleta = oImpCapServ.ImprimeBoletaGiros(rsDatos!cCtaCod, rsDatos!cRemitente, sDest(), rsDatos!nMonto, rsDatos!nComision, gdFecSis, Format$(Time, "hh:mm:ss"), gsNomAge, gsCodUser, sLpt, , 0, rsDatos!cagedest, gbImpTMU) 'Comentado by NAGL 20181030
             gbImpTMU = False 'NAGL 20181030
             sImpreBoleta = oImpCapServ.ImprimeBoletaGirosNew(rsDatos!cCtaCod, rsDatos!cRemitente, sDest(), rsDatos!cDNIDestinatario, rsDatos!nMonto, rsDatos!nComision, gdFecSis, Format$(Time, "hh:mm:ss"), gsNomAge, gsCodUser, rsDatos!nITF, rsDatos!cagedest, gbImpTMU, psOpeCod, rsDatos!cDNIRemitente, IIf(rsDatos!pbITFEfect = "02", True, False)) 'NAGL Según RFC1807260001
        ElseIf psOpeCod = "310301" Then
             gbImpTMU = False 'NAGL 20181030
             sImpreBoleta = oImpCapServ.ImprimeBoletaCancelacionGiros(rsDatos!cCtaCod, rsDatos!cRemitente, sDest(), rsDatos!nMonto, gdFecSis, Format$(Time, "hh:mm:ss"), gsNomAge, gsCodUser, sLpt, , gbImpTMU, psOpeCod, rsDatos!nITF)
             'NAGL Agregó los parámetros psOpeCod, rsDatos!nITF
        End If
    
    '310401  PASI0151111 ERS0692015
    Case "310401"
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpGiroCambDest(pnMovNro)
        sImpreBoleta = oImpCapServ.ImprimeBoletaGirosCambDestinatario(rsDatos!cCtaCod, rsDatos!cRemitente, rsDatos!cDestinatario, rsDatos!nMonto, rsDatos!nComision, gdFecSis, Format$(Time, "hh:mm:ss"), rsDatos!cagedest, gsCodUser, "", , 0, "", 0)
    
    '310402  PASI0151111 ERS0692015
    Case "310402"
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpGiroAnulacion(pnMovNro)
        sImpreBoleta = oImpCapServ.ImprimeBoletaGirosAnulacion(rsDatos!cCtaCod, rsDatos!cRemitente, rsDatos!cDestinatario, rsDatos!nMonto, rsDatos!nComision, gdFecSis, Format$(Time, "hh:mm:ss"), rsDatos!cagedest, gsCodUser, "", , 0, "", 0)
         
    'DEPOSITOS AHORROS
    Case gAhoDepEfec, gAhoDepChq
        
        Set rsDatos = oDatosVoucherCap.ObtieneDatosReimprDepositoCaptac(pnMovNro)
        
        sTipApe = "DEPOSITO AHORROS"
        
        If psOpeCod = gAhoDepEfec Then
            sMsgOpe = "Depósito Efectivo"
        ElseIf psOpeCod = gAhoDepChq Then
            sMsgOpe = "Dep.Chq"
        End If
        
        nDiasTranscurridos = DateDiff("d", rsDatos!dUltCierre, gdFecSis) - 1
        If nDiasTranscurridos < 0 Then
            nDiasTranscurridos = 0
        End If
        nIntGanado = oImprCaptac.GetInteres(rsDatos!nSaldoDisponible - rsDatos!nMonto, rsDatos!nTasaInteres, nDiasTranscurridos, TpoCalcIntSimple)
        
        sNomTit = ImpreCarEsp(clsMant.GetNombreTitulares(rsDatos!cCtaCod))
        
        'Obtener el Tipo de pago de ITF
        If Trim(Left(rsDatos!nValoresITF, 6)) = "990102" Then
            nTipoPag = 1
        Else
            nTipoPag = 2
        End If
        
        If rsDatos!bCheque = 1 Then
            'If sCodCmac <> "" Then
            '    sImpreBoleta = sImpreBoleta & oImprCaptac.ImprimeBoleta(sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, psOpeCod, rsDatos!nMonto, sNomTit, rsDatos!cCtaCod, Format$(rsDatos!dFecValorizacion, "dd/mm/yyyy"), rsDatos!nSaldoDisponible, nIntGanado, "Fecha Valor", rsDatos!nTransacc, nSaldoCnt, True, , , , , , , gsNomCmac, , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, Mid(rsDatos!cMovNro, 18, 2), , False, CStr(psOpeCod), , , , , , , , False, nTipoPag, CDbl(Trim(Right(rsDatos!nValoresITF, 6))), True, , , gbImpTMU)
            'Else
                sImpreBoleta = sImpreBoleta & oImprCaptac.ImprimeBoleta(sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & rsDatos!cNumCheque, psOpeCod, rsDatos!nMonto, sNomTit, rsDatos!cCtaCod, Format$(rsDatos!dFecValorizacion, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", rsDatos!nTransacc, nSaldoCnt, True, , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, Mid(rsDatos!cMovNro, 18, 2), , False, CStr(psOpeCod), , , , , , , , False, nTipoPag, CDbl(Trim(Right(rsDatos!nValoresITF, 6))), True, , , gbImpTMU)
            'End If
        Else
            'If sCodCMAC <> "" Then
            '    sTipApe = "DEPOSITO CMAC AHORROS"
            '    sImpreBoleta = sImpreBoleta & oImprCaptac.ImprimeBoleta(sTipApe, ImpreCarEsp(sMsgOpe), psOpeCod, rsDatos!nMonto, sNomTit, rsDatos!cCtaCod, "", rsDatos!nSaldoDisponible, nIntGanado, "", rsDatos!nTransacc, rsDatos!nSaldoContable, True, , , , , , , gsNomCmac, , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, Mid(rsDatos!cMovNro, 18, 2), , False, CStr(psOpeCod), , , , , , , , False, nTipoPag, CDbl(Trim(Right(rsDatos!nValoresITF, 6))), True, rsDatos!nITFOtraAge, , gbImpTMU)
            'Else
                sImpreBoleta = sImpreBoleta & oImprCaptac.ImprimeBoleta(sTipApe, ImpreCarEsp(sMsgOpe), psOpeCod, rsDatos!nMonto, sNomTit, rsDatos!cCtaCod, "", rsDatos!nSaldoDisponible, nIntGanado, "", rsDatos!nTransacc, rsDatos!nSaldoContable, True, , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, Mid(rsDatos!cMovNro, 18, 2), , False, CStr(psOpeCod), , , , , , , , False, nTipoPag, CDbl(Trim(Right(rsDatos!nValoresITF, 6))), True, rsDatos!nITFOtraAge, , gbImpTMU, rsDatos!cNumTarjeta)
            'End If
        End If
        
    'DEPOSITOS CTS
    Case gCTSDepEfec, gCTSDepChq, gCTSDepTransf 'PASI20151102 incluyo gCTSDepTransf

        Set rsDatos = oDatosVoucherCap.ObtieneDatosReimprDepositoCaptac(pnMovNro)
    
        If psOpeCod = gCTSDepEfec Then
            sMsgOpe = "Depósito Efectivo"
        ElseIf psOpeCod = gCTSDepChq Then
            sMsgOpe = "Depósito Cheque"
        ElseIf psOpeCod = gCTSDepTransf Then 'PASI20151102
            sMsgOpe = "Depósito Transferencia"
        End If
        
        sTipApe = "DEPOSITO CTS"
        
        sNomTit = ImpreCarEsp(clsMant.GetNombreTitulares(rsDatos!cCtaCod))
        
        nDiasTranscurridos = DateDiff("d", rsDatos!dUltCierre, gdFecSis) - 1
        If nDiasTranscurridos < 0 Then
            nDiasTranscurridos = 0
        End If
        nIntGanado = oImprCaptac.GetInteres(rsDatos!nSaldoDisponible - rsDatos!nMonto, rsDatos!nTasaInteres, nDiasTranscurridos, TpoCalcIntSimple)
        
        If rsDatos!bCheque = 1 Then
            'Select Case nTipoDoc
            '    Case TpoDocCheque
                    sImpreBoleta = sImpreBoleta & oImprCaptac.ImprimeBoleta(sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & rsDatos!cNumCheque, psOpeCod, rsDatos!nMonto, sNomTit, rsDatos!cCtaCod, Format$(rsDatos!dFecValorizacion, "dd/mm/yyyy"), rsDatos!nSaldoDisponible, nIntGanado, "Fecha Valor", rsDatos!nTransacc, rsDatos!nSaldoContable, True, , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, , , , , , , , , , , , False, , , , , , gbImpTMU)
            'End Select
        Else
            'If sCodCmac <> "" Then
            '    sTipApe = "DEPOSITO CMAC CTS"
            '    psImpBoletas = psImpBoletas & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, , , , , , , , , , , , False, , , , , , pbImpTMU)
            'Else
                sImpreBoleta = sImpreBoleta & oImprCaptac.ImprimeBoleta(sTipApe, ImpreCarEsp(sMsgOpe), psOpeCod, rsDatos!nMonto, sNomTit, rsDatos!cCtaCod, "", rsDatos!nSaldoDisponible, nIntGanado, "", 1, rsDatos!nSaldoContable, True, , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, , , , , , , , , , , , False, , , , , , gbImpTMU)
            'End If
        End If
        
    'RETIROS AHORROS
    Case gAhoRetEfec, gAhoRetOP, gAhoRetEmiChq
        
        Set rsDatos = oDatosVoucherCap.ObtieneDatosReimprRetiroCaptac(pnMovNro)
        
        sTipApe = "RETIRO AHORROS"
        
        If psOpeCod = gAhoRetOP Then
            sMsgOpe = "Retiro OP"
        ElseIf psOpeCod = gAhoRetEfec Then
            sMsgOpe = "Retiro Efectivo"
        ElseIf psOpeCod = gAhoRetEmiChq Then
            sMsgOpe = "Retiro Emisión Cheque"
        End If
        
        sNomTit = ImpreCarEsp(clsMant.GetNombreTitulares(rsDatos!cCtaCod))
        
        nDiasTranscurridos = DateDiff("d", rsDatos!dUltCierre, gdFecSis) - 1
        If nDiasTranscurridos < 0 Then
            nDiasTranscurridos = 0
        End If
        nIntGanado = oImprCaptac.GetInteres(rsDatos!nSaldoDisponible - rsDatos!nMonto, rsDatos!nTasaInteres, nDiasTranscurridos, TpoCalcIntSimple)
                
        'Obtener el Tipo de pago de ITF
        If Trim(Left(rsDatos!nValoresITF, 6)) = "990102" Then
            nTipoPag = 1
        Else
            nTipoPag = 2
        End If
                
        If rsDatos!bCheque = 1 Then
        Select Case rsDatos!nTipoDoc
            Case TpoDocOrdenPago
                'If sCodCmac <> "" Then
                '    sTipApe = "RETIRO AHORROS CMACMAYNAS"
                '    sImpreBoleta = sImpreBoleta & oImprCaptac.ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, False, False, , , , True, , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, , , bImpreSaldos, CStr(nOperacion), , , , , , , , False, pnTipoITF, CDbl(pnITFValor), True, pnComiRetOtraAge, pnComiRetxMaxOpe, pbImpTMU)
                'Else
                    sImpreBoleta = sImpreBoleta & oImprCaptac.ImprimeBoleta(sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & rsDatos!cNumCheque, psOpeCod, rsDatos!nMonto, sNomTit, rsDatos!cCtaCod, "", rsDatos!nSaldoDisponible - CDbl(Trim(Right(rsDatos!nValoresITF, 6))), nIntGanado, "", rsDatos!nTransacc, rsDatos!nSaldoContable, False, False, , , , True, , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, , , False, CStr(psOpeCod), , , , , , , , False, nTipoPag, CDbl(Trim(Right(rsDatos!nValoresITF, 6))), True, rsDatos!nComRetOtraAge, rsDatos!nComRetOtraAge, gbImpTMU)
                'End If
            End Select
        Else
            'If sCodCmac <> "" Then
            '    sTipApe = "RETIRO AHORROS CMACMAYNAS"
            '    sImpreBoleta = sImpreBoleta & oImprCaptac.ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, False, False, , , , True, , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, , , bImpreSaldos, CStr(nOperacion), , , , , , , , True, pnTipoITF, CDbl(pnITFValor), True, pnComiRetOtraAge, pnComiRetxMaxOpe, pbImpTMU)
            'Else
                sImpreBoleta = sImpreBoleta & oImprCaptac.ImprimeBoleta(sTipApe, ImpreCarEsp(sMsgOpe), psOpeCod, rsDatos!nMonto, sNomTit, rsDatos!cCtaCod, "", rsDatos!nSaldoDisponible - CDbl(Trim(Right(rsDatos!nValoresITF, 6))), nIntGanado, "", rsDatos!nTransacc, rsDatos!nSaldoContable, False, False, , , , True, , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, , , False, CStr(psOpeCod), , , , , , , , True, nTipoPag, CDbl(Trim(Right(rsDatos!nValoresITF, 6))), True, rsDatos!nComRetOtraAge, rsDatos!nComRetOtraAge, gbImpTMU, rsDatos!cNumTarjeta, , , , , 0, "")
            'End If
        End If
        
    '200304 PASI20151026 ERS0692015
    Case gAhoRetOPCanje
        Dim sImpreBoletaOPCanje As String
        Dim sImpreBoletaOPCanjeComi As String
        Dim rsDatosOPCanje As ADODB.Recordset
        Dim rsDatosOPCanjeComi As ADODB.Recordset
        
        Set rsDatosOPCanje = oDatosVoucher.ObtieneDatosReimpRetiroCaptacOrdPagCanje(pnMovNro)
        Set rsDatosOPCanjeComi = oDatosVoucher.ObtienDatosReimpRetiroCaptacOrdPagCanjeComision(rsDatosOPCanje!nMovNro, rsDatosOPCanje!cUser, rsDatosOPCanje!cCtaCod)
        nDiasTranscurridos = DateDiff("d", rsDatosOPCanje!dUltCierre, gdFecSis) - 1
        If nDiasTranscurridos < 0 Then
            nDiasTranscurridos = 0
        End If
        nIntGanado = oDatosVoucherCap.GetInteres(rsDatosOPCanje!nsaldoant, rsDatosOPCanje!nTasaInteres, nDiasTranscurridos, TpoCalcIntSimple)
        oDatosVoucherCap.IniciaImpresora gImpresora
        sImpreBoletaOPCanje = oDatosVoucherCap.ImprimeBoleta(rsDatosOPCanje!cTipApe, ImpreCarEsp(rsDatosOPCanje!cModDep), rsDatosOPCanje!cOpeCod, Trim(rsDatosOPCanje!nMonto), ImpreCarEsp(clsMant.GetNombreTitulares(rsDatosOPCanje!cCtaCod)), rsDatosOPCanje!cCtaCod, "", rsDatosOPCanje!nSaldoDisp, nIntGanado, "", rsDatosOPCanje!nTransacc, rsDatosOPCanje!nSaldoCont, False, False, , , , True, , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, , , False, CStr(rsDatosOPCanje!cOpeCod), , , , , , , , False, rsDatosOPCanje!ntpopag, CDbl(rsDatosOPCanje!nITF), True, rsDatosOPCanje!ComiRetOtraAge, rsDatosOPCanje!ComiRetxMaxOpe, gbImpTMU)
        nDiasTranscurridos = DateDiff("d", rsDatosOPCanjeComi!dUltCierre, gdFecSis) - 1
        If nDiasTranscurridos < 0 Then
            nDiasTranscurridos = 0
        End If
        nIntGanado = oDatosVoucherCap.GetInteres(rsDatosOPCanjeComi!nsaldoant, rsDatosOPCanjeComi!nTasaInteres, nDiasTranscurridos, TpoCalcIntSimple)
        sImpreBoletaOPCanjeComi = oDatosVoucherCap.ImprimeBoleta(rsDatosOPCanjeComi!cTipApe, ImpreCarEsp(rsDatosOPCanjeComi!cModDep), rsDatosOPCanjeComi!cOpeCod, Trim(rsDatosOPCanjeComi!nMonto), ImpreCarEsp(clsMant.GetNombreTitulares(rsDatosOPCanjeComi!cCtaCod)), rsDatosOPCanjeComi!cCtaCod, "", rsDatosOPCanjeComi!nSaldoDisp, nIntGanado, "", rsDatosOPCanjeComi!nTransacc, rsDatosOPCanjeComi!nSaldoCont, False, False, , , , True, , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, , , False, CStr(rsDatosOPCanjeComi!cOpeCod), , , , , , , , True, rsDatosOPCanjeComi!ntpopag, CDbl(rsDatosOPCanjeComi!nITF), True, rsDatosOPCanjeComi!ComiRetOtraAge, rsDatosOPCanjeComi!ComiRetxMaxOpe, gbImpTMU, rsDatosOPCanjeComi!cNumTarjeta, , , , , 0, "", , , "")
        
    '200323 PASI20151027 ERS0692015
    Case gAhoRetComTransferencia
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpRetiroCaptacComisTransf(pnMovNro)
        nDiasTranscurridos = DateDiff("d", rsDatos!dUltCierre, gdFecSis) - 1
        If nDiasTranscurridos < 0 Then
            nDiasTranscurridos = 0
        End If
        nIntGanado = oDatosVoucherCap.GetInteres(rsDatos!nsaldoant, rsDatos!nTasaInteres, nDiasTranscurridos, TpoCalcIntSimple)
        sImpreBoleta = oDatosVoucherCap.ImprimeBoleta(rsDatos!cTipRet, ImpreCarEsp(rsDatos!cModRet), rsDatos!cOpeCod, Trim(rsDatos!nMonto), ImpreCarEsp(clsMant.GetNombreTitulares(rsDatos!cCtaCod)), rsDatos!cCtaCod, "", rsDatos!nSaldoDisp, nIntGanado, "", rsDatos!nTransacc, rsDatos!nSaldoCont, False, False, , , , True, , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, , , False, CStr(rsDatos!cOpeCod), , , , , , , , True, rsDatos!ntpopag, CDbl(rsDatos!nITF), True, rsDatos!ComiRetOtraAge, rsDatos!ComiRetxMaxOpe, gbImpTMU, rsDatos!cNumTarjeta, , , , , 0, "", , , "")
    
    '210201 PASI20151028 ERS0692015
    Case gPFRetInt
        
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpRetiroPFIntEfec(pnMovNro)
        oDatosVoucherCap.IniciaImpresora gImpresora
        dUltRetInt = oDatosVoucherCap.GetFechaUltimoRetiroIntPF(rsDatos!cCtaCod)
        nDias = Format$(DateDiff("d", dUltRetInt, gdFecSis), "#0")
        
        Select Case rsDatos!nFormaRetiro
            Case 1
                dFechaRenovacionPFM = DateAdd("d", rsDatos!nPlazo, rsDatos!dRenovacion)
                dFechaProx = DateAdd("d", 30, gdFecSis)
                
                If dFechaRenovacionPFM > dFechaProx Then
                   sMenProx = "RET. MENSUAL PROX RETIRO:"
                   sMenProx = Trim(sMenProx) & " " & Format(dFechaProx, "dd/MM/yyyy")
                Else
                    If DateDiff("d", dFechaRenovacionPFM, dFechaProx) < 30 Then
                        dFechaProx = DateAdd("d", 30, dFechaRenovacionPFM)
                        lsMenProx2 = "FECHA RENOVACIÓN:"
                        lsMenProx2 = Trim(lsMenProx2) & " " & Format(dFechaRenovacionPFM, "dd/MM/yyyy")
                        sMenProx = "RET. MENSUAL PROX RETIRO:"
                        sMenProx = Trim(sMenProx) & " " & Format(dFechaProx, "dd/MM/yyyy")
                    End If
                    
                End If
            Case 2
                dFechaProx = DateAdd("d", rsDatos!nPlazo, rsDatos!dRenovacion)
                sMenProx = "RET. FIN PLAZO PROX RENOV:"
                sMenProx = Trim(sMenProx) & " " & Format(dFechaProx, "dd/MM/yyyy")
        End Select
        
        sImpreBoleta = oDatosVoucherCap.ImprimeBoleta(rsDatos!cTipApe, ImpreCarEsp(rsDatos!cModDep), rsDatos!cOpeCod, rsDatos!nMonto, ImpreCarEsp(clsMant.GetNombreTitulares(rsDatos!cCtaCod)), rsDatos!cCtaCod, "", rsDatos!nSaldoDisp, rsDatos!nMonto, "", 1, rsDatos!nSaldoCont, , , Trim(nDias), , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, , , , , , , sMenProx, , , , , True, rsDatos!ntpopag, rsDatos!nITF, True, , , gbImpTMU, , , , , , , , , , "", lsMenProx2)
    
    '210202 PASI20151028 ERS0692015
    Case gPFRetIntAboAho
        Dim sImpreBoletaPF As String
        Dim sImpreBoletaAho As String
        Dim rsDatosPF As ADODB.Recordset
        Dim rsDatosAho As ADODB.Recordset
        
        Set rsDatosPF = oDatosVoucher.ObtieneDatosReimpRetiroPFIntAboCtaBoletaPF(pnMovNro)
        Set rsDatosAho = oDatosVoucher.ObtieneDatosReimpRetiroPFIntAboCtaBoletaAho(pnMovNro)
        oDatosVoucherCap.IniciaImpresora gImpresora
        dUltRetInt = oDatosVoucherCap.GetFechaUltimoRetiroIntPF(rsDatosPF!cCtaCod)
        nDias = Format$(DateDiff("d", dUltRetInt, gdFecSis), "#0")
        
        Select Case rsDatosPF!nFormaRetiro
            Case 1
                dFechaRenovacionPFM = DateAdd("d", rsDatosPF!nPlazo, rsDatosPF!dRenovacion)
                dFechaProx = DateAdd("d", 30, gdFecSis)
                
                If dFechaRenovacionPFM > dFechaProx Then
                   sMenProx = "RET. MENSUAL PROX RETIRO:"
                   sMenProx = Trim(sMenProx) & " " & Format(dFechaProx, "dd/MM/yyyy")
                Else
                    If DateDiff("d", dFechaRenovacionPFM, dFechaProx) < 30 Then
                        dFechaProx = DateAdd("d", 30, dFechaRenovacionPFM)
                        lsMenProx2 = "FECHA RENOVACIÓN:"
                        lsMenProx2 = Trim(lsMenProx2) & " " & Format(dFechaRenovacionPFM, "dd/MM/yyyy")
                        sMenProx = "RET. MENSUAL PROX RETIRO:"
                        sMenProx = Trim(sMenProx) & " " & Format(dFechaProx, "dd/MM/yyyy")
                    End If
                    
                End If
            Case 2
                dFechaProx = DateAdd("d", rsDatosPF!nPlazo, rsDatosPF!dRenovacion)
                sMenProx = "RET. FIN PLAZO PROX RENOV:"
                sMenProx = Trim(sMenProx) & " " & Format(dFechaProx, "dd/MM/yyyy")
        End Select
        sImpreBoletaPF = oDatosVoucherCap.ImprimeBoleta(rsDatosPF!cTipApe, ImpreCarEsp(rsDatosPF!cModDep), rsDatosPF!cOpeCod, Trim(rsDatosPF!nMonto - rsDatosPF!nITF), ImpreCarEsp(clsMant.GetNombreTitulares(rsDatosPF!cCtaCod)), rsDatosPF!cCtaCod, "", rsDatosPF!nSaldoDisp, rsDatosPF!nMonto, "", 1, rsDatosPF!nSaldoCont, , , Trim(nDias), , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, , , , , , , sMenProx, , , , , True, rsDatosPF!ntpopag, rsDatosPF!nITF, True, , , gbImpTMU, , , , , , , , , , "", lsMenProx2)
        
        nDiasTranscurridos = DateDiff("d", rsDatosAho!dUltCierre, gdFecSis) - 1
        If nDiasTranscurridos < 0 Then
            nDiasTranscurridos = 0
        End If
        nIntGanado = oImprCaptac.GetInteres(rsDatosAho!nsaldoant, rsDatosAho!nTasaInteres, nDiasTranscurridos, TpoCalcIntSimple)
        sImpreBoletaAho = oDatosVoucherCap.ImprimeBoleta(rsDatosAho!cTipApe, ImpreCarEsp(rsDatosAho!cModDep), rsDatosAho!cOpeCod, rsDatosAho!nMonto, ImpreCarEsp(clsMant.GetNombreTitulares(rsDatosAho!cCtaCod)), rsDatosAho!cCtaCod, "", rsDatosAho!nSaldoDisp, nIntGanado, "", 1, rsDatosAho!nSaldoCont, , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, , , , , , , , , , , , True, rsDatosAho!ntpopag, rsDatosAho!nITF, True, , , gbImpTMU, , , , , , , , , , 0)
    
    '210206 PASI20151028 ERS0692015
    Case gPFRetIntAdelantado
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpRetiroPFIntCash(pnMovNro)
        dUltRetInt = oDatosVoucherCap.GetFechaUltimoRetiroIntPF(rsDatos!cCtaCod)
        nDias = Format$(DateDiff("d", dUltRetInt, gdFecSis), "#0")
        nIntGanado = oImprCaptac.GetInteres(rsDatos!nSaldoDisp, rsDatos!nTasaInteres, rsDatos!nPlazo, TpoCalcIntAdelantado)
        
        sImpreBoleta = oDatosVoucherCap.ImprimeBoleta(rsDatos!cTipApe, ImpreCarEsp(rsDatos!cModDep), rsDatos!cOpeCod, Trim(rsDatos!nMonto - rsDatos!nITF), ImpreCarEsp(clsMant.GetNombreTitulares(rsDatos!cCtaCod)), rsDatos!cCtaCod, "", rsDatos!nSaldoDisp, nIntGanado, "", 1, rsDatos!nSaldoCont, , , Trim(nDias), , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, , , , , , , "", , , , , True, rsDatos!ntpopag, rsDatos!nITF, True, , , gbImpTMU, , , , , , , , , , "", "")
    
    'RETIROS CTS
        
    '************************************* Agregado por riro el 20130314
    
    Case gDepositoRecaudo
    
        Dim oConvenioRecaudo As COMNCaptaGenerales.NCOMCaptaMovimiento
        Set oConvenioRecaudo = New COMNCaptaGenerales.NCOMCaptaMovimiento
        
        Set rsDatos = New ADODB.Recordset
        Set rsDatos = oConvenioRecaudo.getDatosVaucherRecaudo(pnMovNro)
        
        sImpreBoleta = sImpreBoleta & oConvenioRecaudo.ImprimeVaucherRecaudo(rsDatos, gbImpTMU, gsCodUser)
        
    '******************************************* fin riro
    
    
    Case gCTSRetEfec
    
        Set rsDatos = oDatosVoucherCap.ObtieneDatosReimprRetiroCaptac(pnMovNro)
        
        sTipApe = "RETIRO CTS"
        
        If psOpeCod = gCTSRetEfec Then
            sMsgOpe = "Retiro Efectivo"
        End If
        
        sNomTit = ImpreCarEsp(clsMant.GetNombreTitulares(rsDatos!cCtaCod))
    
        nDiasTranscurridos = DateDiff("d", rsDatos!dUltCierre, gdFecSis) - 1
        If nDiasTranscurridos < 0 Then
            nDiasTranscurridos = 0
        End If
        nIntGanado = oImprCaptac.GetInteres(rsDatos!nSaldoDisponible - rsDatos!nMonto, rsDatos!nTasaInteres, nDiasTranscurridos, TpoCalcIntSimple)
              
        If rsDatos!bCheque = 1 Then
            sImpreBoleta = sImpreBoleta & oImprCaptac.ImprimeBoleta(sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & rsDatos!cNumCheque, psOpeCod, rsDatos!nMonto, sNomTit, rsDatos!cCtaCod, "", rsDatos!nSaldoDisponible, nIntGanado, "", rsDatos!nTransacc, rsDatos!nSaldoContable, True, False, , , , True, , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , , , , , , , , , , , , , True, , , , , , gbImpTMU)
        Else
            'If sCodCmac <> "" Then
            '    sTipApe = "RETIRO CMAC CTS"
            '    sImpreBoleta = sImpreBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, False, , , , , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , , , , , , , , , , , , , True, , , , , , pbImpTMU)
            'Else
                sImpreBoleta = sImpreBoleta & oImprCaptac.ImprimeBoleta(sTipApe, ImpreCarEsp(sMsgOpe), psOpeCod, rsDatos!nMonto, sNomTit, rsDatos!cCtaCod, "", rsDatos!nSaldoDisponible, nIntGanado, "", 1, rsDatos!nSaldoContable, True, False, , , , True, , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , , , , , , , , , , , , , True, , , , , , gbImpTMU, , , , , , 0, "")
            'End If
        End If
        
    'CREDITO PIGNORATICIO RENOVACION
    Case gColPOpeRenovacEFE, gColPOpeRenovNorEFE, gColPOpeRenovMorEFE, gColPOpeRenovacCHQ, gColPOpeRenovNorCHQ, gColPOpeRenovMorCHQ
        
        Set rsDatos = oDatosVoucher.ObtieneDatosReimprPigno(pnMovNro)
        
        sImpreBoleta = oImprPig.nPrintReciboRenovacion(gsNomAge, fgFechaHoraGrab(rsDatos!cMovNro), rsDatos!cCtaCod, rsDatos!cPersNombre, _
                        Format(rsDatos!dVigencia, "mm/dd/yyyy"), DateDiff("d", rsDatos!dVenc, gdFecSis), rsDatos!nsaldoant, rsDatos!nCapPag, _
                        rsDatos!nIntCompPag, rsDatos!nIntMoratPag, rsDatos!nImpuestoPag, rsDatos!nCustodiaVencPag + rsDatos!nCustodiaPag, _
                        rsDatos!nCostRematePag, rsDatos!nMontoRenov, rsDatos!nSaldo, rsDatos!nTasaInteres, _
                        rsDatos!nNroRenov, Format(rsDatos!dVenc, "dd/mm/yyyy"), gsCodUser, 30, "", rsDatos!nCostNotifPag, " ", rsDatos!nITF, gImpresora, rsDatos!nIntVencPag, gbImpTMU)
    
    '121600 PASI20150929 ERS0692015
    Case gColPOpeImpDuplicado
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpDupliContPigno(pnMovNro)
        sImpreBoleta = oImprPig.nPrintReciboDuplicadoContrato(gsNomAge, rsDatos!FechaHoraGrab, rsDatos!cCtaCod, rsDatos!cPersNombre, _
                        rsDatos!nMonto, rsDatos!nNroDuplic, rsDatos!nTasaInteres, gsCodUser, "", rsDatos!nITF, gImpresora)
    
    '121900 PASI20150930 ERS0692015
    Case gColPOpeCobCusDiferida
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpCustoDifPigno(pnMovNro)
        sImpreBoleta = oImprPig.nPrintReciboCobroCustodia(gsNomAge, rsDatos!FechaHoraGrab, rsDatos!cCtaCod, rsDatos!cPersNombre, _
                        rsDatos!nMonto, 0, 0, gsCodUser, "", rsDatos!nITF, gImpresora)
    
    '122900 PASI20150930 ERS0692015
    Case "122900"
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpRecupContPigno(pnMovNro)
        sImpreBoleta = oImprPig.ImpRecupSub(rsDatos!cCtaCod, PstaNombre(rsDatos!cPersNombre, False), rsDatos!nmontopreventabruta, rsDatos!nMontoTotal, rsDatos!nITF, gsNomAge, gdFecSis, gsCodUser, gImpresora)
    
    'COMISIONES 'JUEZ 20130902
    'Case gComisionDiversasAhoGasto, gComisionDiversasAhoCom, gAhoCargoCobroComDiversasAho, gCTSCargoCobroComDiversasAho
    Case gAhoCargoCobroComDiversasAho, gCTSCargoCobroComDiversasAho  'PASI20160215
        
        Set rsDatos = oDatosVoucher.ObtieneDatosReimprComisiones(pnMovNro)
        
        sImpreBoleta = oImprDesembPagoCred.ImprimeBoletaComision("COMISIONES VARIAS", rsDatos!cCodCom & " - " & rsDatos!cComDesc, "", Str(CDbl(rsDatos!nMovImporte)), rsDatos!cPersNombre, rsDatos!cPersIDnro, IIf(rsDatos!cCtaCod <> "", rsDatos!cCtaCod, "________" & rsDatos!nmoneda), False, "", "", , gdFecSis, gsNomAge, gsCodUser, sLpt, , gbImpTMU, True, rsDatos!cTipoPago)
    
    'SEGUROS 'JUEZ 20140711
    Case gAhoCargoAfilSegTarjeta, gCTSCargoAfilSegTarjeta
        
        Dim oDSeg As COMDCaptaGenerales.DCOMSeguros
        Set oDSeg = New COMDCaptaGenerales.DCOMSeguros
        Set rsDatos = oDSeg.RecuperaSegTarjetaAfiliacion(pnMovNro)
        Set oDSeg = Nothing
        sImpreBoleta = oImpSeg.ImprimeBoletaAfilicacionSeguroTarjeta(pnMovNro, rsDatos!cMovNroReg, gsNomAge, gbImpTMU)
            
    '200401,200403  PASI20151027 ERS0692015
    Case gAhoCancAct, gAhoCancTransfAct
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpCancelActivCaptac(pnMovNro)
        If psOpeCod = gAhoCancAct Then
            sMsgOpe = "Cancelación Efectivo"
        ElseIf psOpeCod = gAhoCancTransfAct Then
            sMsgOpe = "Cancelación Transferencia"
        End If
        oDatosVoucherCap.IniciaImpresora gImpresora
        sImpreBoleta = oDatosVoucherCap.ImprimeBoletaInteres(rsDatos!cTitCan, ImpreCarEsp(sMsgOpe), rsDatos!cOpeCod, Trim(rsDatos!nMonto), ImpreCarEsp(clsMant.GetNombreTitulares(rsDatos!cCtaCod)), rsDatos!cCtaCod, "", 0, Trim(0), "", rsDatos!nTransacc, "Interes Retirado", Trim(rsDatos!nIntAcum), , , , gsNomAge, gdFecSis, gsCodUser, sLpt, gsCodCMAC, True, rsDatos!ntpopag, rsDatos!nITF, True, gbImpTMU, rsDatos!nMontoComVB, 0)
        
    '210301 PASI20151102 ERS0692015
    Case gPFCancEfec
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpCancelaPFEfectivo(pnMovNro)
        oDatosVoucherCap.IniciaImpresora gImpresora
        sImpreBoleta = oDatosVoucherCap.ImprimeBoletaInteres(rsDatos!cTipApe, rsDatos!cModDep, rsDatos!cOpeCod, Trim(rsDatos!nMonto), ImpreCarEsp(clsMant.GetNombreTitulares(rsDatos!cCtaCod)), rsDatos!cCtaCod, "", 0, Trim(0), "", rsDatos!nTransacc, "Interes Retirado", Trim(rsDatos!nintret), , , , gsNomAge, gdFecSis, gsCodUser, sLpt, gsCodCMAC, True, rsDatos!ntpopag, rsDatos!nITF, True, gbImpTMU, , 0)
               
    '220401 PASI20151102 ERS0692015
    Case gCTSCancEfec
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpCancelaCTSEfectivo(pnMovNro)
        oDatosVoucherCap.IniciaImpresora gImpresora
        sImpreBoleta = oDatosVoucherCap.ImprimeBoletaInteres(rsDatos!cTipApe, ImpreCarEsp(rsDatos!cModDep), rsDatos!cOpeCod, Trim(rsDatos!nMonto), ImpreCarEsp(clsMant.GetNombreTitulares(rsDatos!cCtaCod)), rsDatos!cCtaCod, "", 0, Trim(0), "", rsDatos!nTransacc, "Interes Retirado", Trim(rsDatos!nIntGan), , , , gsNomAge, gdFecSis, gsCodUser, sLpt, gsCodCMAC, , , , , gbImpTMU, rsDatos!nComisionVB, 0)
               
    '210801,210807 PASI20151030 ERS0692015
    Case gPFAumCapEfec, gPFAumCapTrans
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpPFAumentCapEfec(pnMovNro)
        If psOpeCod = gAhoCancAct Then
            sMsgOpe = "AUMENTO CAP PF EFECT"
        ElseIf psOpeCod = gPFAumCapTrans Then
            sMsgOpe = "AUM CAP PF TRANF"
        End If
        oDatosVoucherCap.IniciaImpresora gImpresora
        sImpreBoleta = oDatosVoucherCap.ImprimeBoletaInteres(sMsgOpe, rsDatos!cModDep, rsDatos!cOpeCod, Trim(rsDatos!nMonto), ImpreCarEsp(clsMant.GetNombreTitulares(rsDatos!cCtaCod)), rsDatos!cCtaCod, "", rsDatos!nSaldoDisp, Trim(0), "", rsDatos!nTransacc, "Interes Ganado", Trim(rsDatos!nIntGan), True, Trim(rsDatos!nSaldoCont), , gsNomAge, gdFecSis, gsCodUser, sLpt, gsCodCMAC, False, rsDatos!ntpopag, rsDatos!nITF, True, gbImpTMU)
    
    '300416,300421,300470,300503,300520,300521 PASI20151102 ERS0692015
    Case "300416", "300421", "300470", "300503", "300520", "300521"
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpOtrasOperaciones(pnMovNro)
        If psOpeCod = "300416" Then
            sMsgOpe = Left("OTROS INGRESOS PAGO DE CASILLERO EFECTIVO", 36)
        ElseIf psOpeCod = "300421" Then
            sMsgOpe = "DEVOLUCIÓN DE VIATICOS Y E/R"
        ElseIf psOpeCod = "300470" Then
            sMsgOpe = "COBRO COMISIÓN REPOSICIÓN TARJETA"
        ElseIf psOpeCod = "300503" Then
            sMsgOpe = "DEVOLUCIÓN CRÉDITOS PERSONALES"
        ElseIf psOpeCod = "300520" Then
            sMsgOpe = "EGRESO POR COMISIÓN DE CONVENIO"
        ElseIf psOpeCod = "300521" Then
            sMsgOpe = "EGRESO POR DESCUENTO POR PLANILLA"
        End If
        sImpreBoleta = oImpOpEsp.ImprimeBoleta(rsDatos!cTipApe, sMsgOpe, "", Str(rsDatos!nMonto), rsDatos!cPersNombre, "________" & rsDatos!nmoneda, rsDatos!cDocNro, 0, "0", IIf(Len(rsDatos!cDocNro) = 0, "", "Nro Documento"), 0, 0, False, False, , , , False, , "Nro Ope. : " & Str(rsDatos!nMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False, rsDatos!nITF)
            
      '300493,300494,300523,300595,300596 PASI20151102 ERS0692015
    Case "300493", "300494", "300523", "300595", "300596"
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpOtraOpeDepCtaBanco(pnMovNro)
        sImpreBoleta = oImpOpEsp.ImprimeBoletaDepCtaBanco(gsNomCmac, gsNomAge, gdFecSis, rsDatos!cPersNombre, _
                        rsDatos!cCtaIFDesc, rsDatos!cDocNro, rsDatos!nMovImporte, sLpt, rsDatos!nmoneda, gsCodUser, "DEP. A CTA. BANCO", True)
           
     '300524,300525 PASI20151106 ERS0692015
    Case "300524", "300525"
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpDesemOtroGasto(pnMovNro)
        If psOpeCod = "300524" Then
            sMsgOpe = "DESEMBOLSO POR A RENDIR CUENTA"
        ElseIf psOpeCod = "300525" Then
            sMsgOpe = "DESEMBOLSO PARA VIATICOS"
    
        End If
        oDatosVoucherCap.IniciaImpresora gImpresora
        sImpreBoleta = oDatosVoucherCap.ImprimeBoleta(sMsgOpe, ImpreCarEsp("Monto Desembolsado " & IIf(rsDatos!nmoneda = gMonedaNacional, "S/.", "$.")), CStr(Trim(rsDatos!cOpeCod)), Trim(rsDatos!nMovImporte), rsDatos!cPersNombre, "", "", 0, "", "", rsDatos!nMovNro, 0, False, False, , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , , , CInt(rsDatos!nmoneda), , , , , , , , , , False, , , , , , gbImpTMU)
        sImpDJ = oDatosVoucherCap.imprimirDJ(rsDatos!nMovNro, sLpt, gbImpTMU)
        sImpreBoleta = sImpreBoleta & sImpDJ
        
    '300526 PASI20151106 ERS0692015
    Case "300526"
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpDesemOtroGastoCCHICA(pnMovNro)
        sMsgOpe = "DESEM. PARA CAJA CHICA"
        oDatosVoucherCap.IniciaImpresora gImpresora
        sImpreBoleta = oDatosVoucherCap.ImprimeBoleta(sMsgOpe & rsDatos!nProcNro, ImpreCarEsp("Monto Desembolsado " & IIf(rsDatos!nmoneda = gMonedaNacional, "S/.", "$.")), CStr(Trim(rsDatos!cOpeCod)), Trim(rsDatos!nMovImporte), rsDatos!cPersNombre, "", "", 0, "", "", rsDatos!nMovNro, 0, False, False, , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , , , CInt(rsDatos!nmoneda), , , , , , , , , , False, , , , , , gbImpTMU)
               
    '300528 PASI20151110 ERS0692015
    Case "300528"
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpDevSobOtrasOpeCheque(pnMovNro)
        sImpreBoleta = oImpOpEsp.ImprimeBoleta(rsDatos!cCab, Left(rsDatos!cMsg, 15), rsDatos!cOpeCod, Str(rsDatos!nMonto), ImpreCarEsp(rsDatos!cTitular), "________" & rsDatos!nmoneda, rsDatos!cDocNro, 0, "0", IIf(Len(rsDatos!cDocNro) = 0, "", "Nro Documento"), 0, 0, False, False, , , , False, , "Nro Ope. : " & Str(rsDatos!nMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False, (rsDatos!nITF * -1), , , rsDatos!cDNI)
               
    Case "300901", "300904"
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpComisionOtros(pnMovNro)
        sImpreBoleta = oImprDesembPagoCred.ImprimeBoletaComision(rsDatos!cTit, Left(rsDatos!cMsg, 36), "", Str(CDbl(rsDatos!nMovImporte)), rsDatos!cCliente, rsDatos!cDOI, "________" & rsDatos!nmoneda, False, "", "", , gdFecSis, gsNomAge, gsCodUser, sLpt, , gbImpTMU)
    
    '301000 PASI20151110 ERS0692015
    Case "301000"
        Set rsDatos = oDatosVoucher.ObtieneDatosReimpDebitoSP(pnMovNro)
        sImpreBoleta = oImpOpEsp.imprimirBoletaServicioPago(gsNomCmac, gsNomAge, rsDatos!nmoneda, rsDatos!nMovNro, _
                                                            rsDatos!Fecha, rsDatos!Hora, rsDatos!cEmpresa, rsDatos!cDOI, _
                                                            rsDatos!cBeneficiario, rsDatos!nMonto, gsCodUser, sLpt, gbImpTMU, , , rsDatos!cConvenio)

    'CTI2 FERIMORO : MEJORA EL PROCESO PARA REIMPRIMIR VOUCHER DE COBRANZA DE SOAT
        Case "300155"

            Dim oSeg As New COMNCaptaGenerales.NCOMSeguros
        
        Set rsDatos = oDatosVoucher.ObtieneDatosSoat(pnMovNro)
        
        sImpreBoleta = oSeg.ImprimeBoletaAfiliacionSegSoat(pnMovNro, rsDatos!cMovNro, gsNomAge, gbImpTMU)

            '**************
    End Select

    '***JGPA20190607
    If pnNroREU > 0 Then
        Set rsDatos = oDatosVoucher.ObtieneDatosOperacionREU(pnMovNro, pnNroREU, psOpeCod)
        Set rsDatosIntReu = oDatosIntReu.ObtieneIntervinientesOperacionREU(pnMovNro, pnNroREU)
        If Not (rsDatosIntReu.BOF And rsDatosIntReu.EOF) Then
            ReDim MatInterviniestesReu(rsDatosIntReu.RecordCount, 3)
            For ir = 1 To rsDatosIntReu.RecordCount
                MatInterviniestesReu(ir, 1) = CLng(rsDatosIntReu!nMovNro)
                MatInterviniestesReu(ir, 2) = CStr(rsDatosIntReu!cPersCod)
                MatInterviniestesReu(ir, 3) = CInt(rsDatosIntReu!nMLDRelac)
                rsDatosIntReu.MoveNext
            Next ir
            For ir = 1 To UBound(MatInterviniestesReu)
                Select Case MatInterviniestesReu(ir, 3)
                    Case RealizaTransaccion
                        MatRealiza = frmML.ObtenerDatosReimpresionReu(MatInterviniestesReu(ir, 2))
                    Case OrdenaTransaccion
                        MatOrdena = frmML.ObtenerDatosReimpresionReu(MatInterviniestesReu(ir, 2))
                    Case BeneficiaTransaccion
                        MatBeneficia = frmML.ObtenerDatosReimpresionReu(MatInterviniestesReu(ir, 2))
                End Select
            Next ir
            
            If IsArray(MatRealiza) Then
                If UBound(MatRealiza) Then
                    ReDim ListaPersonasRealizan(0)
                    ListaPersonasRealizan(0).PersCod = MatRealiza(1, 1)
                    ListaPersonasRealizan(0).Nombre = MatRealiza(1, 2)
                    ListaPersonasRealizan(0).DocumentoId = MatRealiza(1, 3)
                    ListaPersonasRealizan(0).Direccion = MatRealiza(1, 4)
                    ListaPersonasRealizan(0).Ocupacion = MatRealiza(1, 5)
                    ListaPersonasRealizan(0).Nacionalidad = MatRealiza(1, 6)
                    ListaPersonasRealizan(0).Residente = MatRealiza(1, 7)
                    ListaPersonasRealizan(0).Peps = MatRealiza(1, 8)
                End If
            End If
            If IsArray(MatOrdena) Then
                If UBound(MatOrdena) Then
                    ReDim ListaPersonasOrdenan(0)
                    ListaPersonasOrdenan(0).PersCod = MatOrdena(1, 1)
                    ListaPersonasOrdenan(0).Nombre = MatOrdena(1, 2)
                    ListaPersonasOrdenan(0).DocumentoId = MatOrdena(1, 3)
                    ListaPersonasOrdenan(0).Direccion = MatOrdena(1, 4)
                    ListaPersonasOrdenan(0).Ocupacion = MatOrdena(1, 5)
                    ListaPersonasOrdenan(0).Nacionalidad = MatOrdena(1, 6)
                    ListaPersonasOrdenan(0).Residente = MatOrdena(1, 7)
                    ListaPersonasOrdenan(0).Peps = MatOrdena(1, 8)
                End If
            End If
            If IsArray(MatBeneficia) Then
                If UBound(MatBeneficia) Then
                    ReDim ListaPersonasBenefician(0)
                    ListaPersonasBenefician(0).PersCod = MatBeneficia(1, 1)
                    ListaPersonasBenefician(0).Nombre = MatBeneficia(1, 2)
                    ListaPersonasBenefician(0).DocumentoId = MatBeneficia(1, 3)
                    ListaPersonasBenefician(0).Direccion = MatBeneficia(1, 4)
                    ListaPersonasBenefician(0).Ocupacion = MatBeneficia(1, 5)
                    ListaPersonasBenefician(0).Nacionalidad = MatBeneficia(1, 6)
                    ListaPersonasBenefician(0).Residente = MatBeneficia(1, 7)
                    ListaPersonasBenefician(0).Peps = MatBeneficia(1, 8)
                End If
            End If
            
        End If
        If Not (rsDatos.BOF And rsDatos.EOF) Then
            sImpreBoletaReu = oImpOpEsp.ImprimeBoletaLavadoDinero(gsNomCmac, gsNomAge, gdFecSis, rsDatos!cCuenta, rsDatos!cTitular, rsDatos!cDOI, rsDatos!cDir, rsDatos!cGiro, _
                                                               ListaPersonasRealizan(0).Nombre, ListaPersonasOrdenan(0).DocumentoId, ListaPersonasRealizan(0).Direccion, ListaPersonasRealizan(0).Ocupacion, _
                                                               ListaPersonasBenefician(0).Nombre, ListaPersonasBenefician(0).DocumentoId, ListaPersonasBenefician(0).Direccion, ListaPersonasBenefician(0).Ocupacion, _
                                                               rsDatos!cOperacion, pnImporte, sLpt, _
                                                               ListaPersonasOrdenan(0).Nombre, ListaPersonasBenefician(0).DocumentoId, ListaPersonasBenefician(0).Direccion, ListaPersonasBenefician(0).Ocupacion, , True, , , gbImpTMU, gsCodAge, Trim(CStr(rsDatos!cOrigenEfectivo)), , , , , , , , , , , , , _
                                                               CInt(rsDatos!nTipoREU), CStr(pnNroREU), ListaPersonasRealizan, ListaPersonasOrdenan, ListaPersonasBenefician)
        End If
    End If
    '***End JGPA
    
    Set oImprDesembPagoCred = Nothing
    Set oDatosVoucher = Nothing
    Set oImprPagJud = Nothing
    Set oImprCF = Nothing
    Set oDatosVoucherCap = Nothing
    Set clsMant = Nothing
    Set oImprCaptac = Nothing
    Set oImprPig = Nothing
    Set oImpSeg = Nothing 'JUEZ 20140711
    
    bExisteBoleta = True
    If psOpeCod = gAhoApeLoteEfec Then
        If sImpreBoletaCab = "" Then bExisteBoleta = False
    ElseIf psOpeCod = gAhoRetOPCanje Then
        If sImpreBoletaOPCanje = "" Then bExisteBoleta = False
    ElseIf psOpeCod = gPFRetIntAboAho Then
        If sImpreBoletaPF = "" Then bExisteBoleta = False
    ElseIf psOpeCod = gCTSApeLoteEfec Then
        If sImpreBoletaCTSCab = "" Then bExisteBoleta = False
    Else
        If sImpreBoleta = "" Then bExisteBoleta = False
    End If
    
    If Not bExisteBoleta Then
        MsgBox "Por el momento la reimpresión de voucher de esta operación no está disponible, Disculpe las molestias", vbInformation, "Aviso"
        Exit Sub
    End If
    'PASI end
    
    
    Set oVisto = New frmVistoElectronico
    bResultadoVisto = oVisto.Inicio(3) ', "909909")
    If Not bResultadoVisto Then
        Exit Sub
    End If
    
    '***JGPA20190607-------------------------------
    If pnNroREU > 0 Then
        pnResponseREU = MsgBox("La operación tiene Registro de Operación (RO)." & Chr(13) & Chr(13) & _
               "¿Desea reimprimir el RO en lugar del voucher?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
    Else
        If MsgBox("Se va a reimprimir el voucher, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        End If
    End If
    'If MsgBox("Se va a reimprimir el voucher, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub 'Comentado
    '***JGPA20190607-------------------------------
    
    Dim oImpr As COMDCredito.DCOMCredActBD
    Set oImpr = New COMDCredito.DCOMCredActBD
    
    Dim loContFunct As COMNContabilidad.NCOMContFunciones
    Dim psMovNro As String
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
    
    psMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    Call oImpr.InsertaDatosImprVoucher(pnMovNro, Me.txtGlosa.Text, psMovNro)
    
    'MARG ERS052-2017----------------------------
    oVisto.RegistraVistoElectronico pnMovNro, , gsCodUser, pnMovNro
    'END MARG --------------------------------------
    
    '***JGPA20190607 Quitar antes de subir al TFS
    Dim clsprevio As clsprevio
    Set clsprevio = New clsprevio
    If pnResponseREU = vbYes Then
        Do
            clsprevio.PrintSpool sLpt, sImpreBoletaReu
        Loop While MsgBox("Desea Reimprimir Registro de Operación?", vbInformation + vbYesNo, "Aviso") = vbYes
        txtGlosa.Text = ""
        Exit Sub
    End If
    '***JGPA20190607
    
    'PASI20151022 ERS0692015
    If psOpeCod = gAhoApeLoteEfec Then
        Do
            clsprevio.PrintSpool sLpt, sImpreBoletaCab
        Loop While MsgBox("Desea Reimprimir Boleta Resumen?", vbInformation + vbYesNo, "Aviso") = vbYes
        Do
            clsprevio.PrintSpool sLpt, sImpreBoletaDet
        Loop While MsgBox("Desea Reimprimir Boleta Detalle?", vbInformation + vbYesNo, "Aviso") = vbYes
        txtGlosa.Text = ""
        Exit Sub
    ElseIf psOpeCod = gCTSApeLoteEfec Then
        Do
            clsprevio.PrintSpool sLpt, sImpreBoletaCTSCab
        Loop While MsgBox("Desea Reimprimir Boleta Resumen?", vbInformation + vbYesNo, "Aviso") = vbYes
        Do
            clsprevio.PrintSpool sLpt, sImpreBoletaCTSDet
        Loop While MsgBox("Desea Reimprimir Boleta Detalle?", vbInformation + vbYesNo, "Aviso") = vbYes
        txtGlosa.Text = ""
        Exit Sub
    ElseIf psOpeCod = gAhoRetOPCanje Then
        Do
            clsprevio.PrintSpool sLpt, sImpreBoletaOPCanje
        Loop While MsgBox("Desea Reimprimir Boleta?", vbInformation + vbYesNo, "Aviso") = vbYes
        Do
            clsprevio.PrintSpool sLpt, sImpreBoletaOPCanjeComi
        Loop While MsgBox("Desea Reimprimir Boleta de Comisión?", vbInformation + vbYesNo, "Aviso") = vbYes
        txtGlosa.Text = ""
        Exit Sub
    ElseIf psOpeCod = gPFRetIntAboAho Then
        sImpreBoletaPF = sImpreBoletaAho & sImpreBoletaPF
        Do
            clsprevio.PrintSpool sLpt, sImpreBoletaPF
        Loop While MsgBox("Desea Reimprimir el voucher?", vbInformation + vbYesNo, "Aviso") = vbYes
    End If
    'end if
    
    Do
        clsprevio.PrintSpool sLpt, sImpreBoleta
    Loop While MsgBox("Desea Reimprimir el voucher?", vbInformation + vbYesNo, "Aviso") = vbYes
    
    txtGlosa.Text = ""
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oMov As DCOMMov
    Dim rsMov As ADODB.Recordset

    Set oMov = New DCOMMov
    
    Set rsMov = oMov.ObtieneMovimientosOpeUsuario(gsCodUser, gdFecSis)
    Set oMov = Nothing
    
    If rsMov.RecordCount = 0 Then
        MsgBox "No se Encontraron Movimientos", vbInformation, "Aviso"
        'Call CmdSalir_Click
        cmdReimprimir.Visible = False
    Else
        Set DBGrdVoucher.DataSource = rsMov
        DBGrdVoucher.Refresh
        cmdReimprimir.Visible = True
    End If
End Sub

Private Sub txtGlosa_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txtGlosa.Text <> "" Then
            If Len(txtGlosa.Text) = 2 Then
                txtGlosa.Text = Mid(txtGlosa.Text, 1, IIf(Len(txtGlosa.Text) = 0, 2, Len(txtGlosa.Text)) - 2)
            End If
        End If
    End If
End Sub
