VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOperaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Opciones"
   ClientHeight    =   5970
   ClientLeft      =   1920
   ClientTop       =   1740
   ClientWidth     =   7125
   HelpContextID   =   210
   Icon            =   "frmOperaciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView tvOpe 
      Height          =   5625
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   9922
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imglstFiguras"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.Animation Logo 
      Height          =   645
      Left            =   480
      TabIndex        =   6
      Top             =   210
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1138
      _Version        =   393216
      FullWidth       =   45
      FullHeight      =   43
   End
   Begin VB.Frame frmMoneda 
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   150
      TabIndex        =   3
      Top             =   1020
      Width           =   1275
      Begin VB.OptionButton optMoneda 
         Caption         =   "M. &E."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   540
         Width           =   795
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "M. &N."
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   150
      TabIndex        =   2
      Top             =   4950
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   150
      TabIndex        =   1
      Top             =   5370
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox rtxt 
      Height          =   315
      Left            =   300
      TabIndex        =   7
      Top             =   4350
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmOperaciones.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstFiguras 
      Left            =   420
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperaciones.frx":038A
            Key             =   "Padre"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperaciones.frx":06DC
            Key             =   "Hijo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperaciones.frx":0A2E
            Key             =   "Hijito"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperaciones.frx":0D80
            Key             =   "Bebe"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lExpand  As Boolean
Dim lExpandO As Boolean
Dim sArea  As String
Dim dFecha As Date, dFecha2 As Date

Public Sub inicio(sObj As String, Optional plExpandO As Boolean = False)
    sArea = sObj
    lExpandO = plExpandO
    Me.Show 0, frmMdiMain
End Sub
Private Sub cmdAceptar_Click()
On Error GoTo ErrorAceptar
    If tvOpe.Nodes.Count = 0 Then
        MsgBox "Lista de operaciones se encuentra vacia", vbInformation, "Aviso"
        Exit Sub
    End If
    If tvOpe.SelectedItem.Tag = "1" Then
        MsgBox "Operación seleccionada no valida...!", vbInformation, "Aviso"
        tvOpe.SetFocus
        Exit Sub
    End If
    gsSimbolo = ""
    gsOpeCod = Mid(tvOpe.SelectedItem.Text, 1, 6)
    If Mid(gsOpeCod, 3, 1) = "1" Then
        gsSimbolo = gcMN
    End If
    If Mid(gsOpeCod, 3, 1) = "2" Then
        gsSimbolo = gcME
    End If
    
    If Left(tvOpe.SelectedItem.Key, 1) <> "P" Then
        gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60)) & ": " & Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescHijo = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescPadre = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60))
    Else
      gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
    End If
    
    'gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60)) & ": " & Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
    Select Case Mid(gsOpeCod, 1, 6)
        Case gCGArendirCtaSolMN, gCGArendirCtaSolME
                frmARendirSolicitud.inicio gArendirTipoCajaGeneral, False
        Case gCGArendirCtaRechMN, gCGArendirCtaRechME
                frmArendirAtencion.inicio gArendirTipoCajaGeneral, ArendirRechazo
        '***Agregado por ELRO el 20120430, según OYP-RFC016-2012
        Case gCGArendirCtaAprobMN, gCGArendirCtaAprobME
                frmARendirAutorizacion.iniciarAprobacionARendir
        Case gCGArendirCtatSust2MN, gCGArendirCtatSust2ME
                FrmARendirSustentacion.iniciarSustentacion gArendirTipoCajaGeneral, ArendirSustentacion
        Case gCGArendirCtaRend2MN, gCGArendirCtaRend2ME
                frmARendirLista.inicio gArendirTipoCajaGeneral, ArendirRendicion, False
        '***Fin Agregado por ELRO*******************************
        Case gCGArendirCtaAtencMN, gCGArendirCtaAtencME
                frmArendirAtencion.inicio gArendirTipoCajaGeneral, ArendirAtencion
        Case gCGArendirCtaSustMN, gCGArendirCtaSustME
             frmARendirLista.inicio gArendirTipoCajaGeneral, ArendirSustentacion, False
        Case gCGArendirCtaRendMN, gCGArendirCtaRendME
                frmARendirLista.inicio gArendirTipoCajaGeneral, ArendirRendicion, False
        Case gCGArendirCtaExtAtencMN, gCGArendirCtaExtAtencME
                frmARendirExtorno.inicio gArendirTipoCajaGeneral, ArendirExtornoAtencion, False
        '***Agregado por ELRO el 20120430, según OYP-RFC016-2012
        Case gCGArendirCtaExtAprobMN, gCGArendirCtaExtAprobME
                frmARendirExtorno.inicio gArendirTipoCajaGeneral, 8, False
        '***Fin Agregado por ELRO*******************************
        Case gCGArendirCtaExtRendMN, gCGArendirCtaExtRendME
                frmARendirExtorno.inicio gArendirTipoCajaGeneral, ArendirExtornoRendicion, False
'        Case gCGArendirViatRend2MN, gCGArendirViatRend2ME
'                frmARendirLista.Inicio gArendirTipoCajaGeneral, ArendirExtornoRendicion, False
        Case gCGArendirCtaRecxHonorarioProvBusca, gCGArendirViatRecxHonorarioProvBusca, gCHRendContabRecxHonorarioProvBusca 'EJVG20140721
            frmProveedorBuscaxAcumulado.Show 1
        'A rendir Agencias
        Case gCGArendirCtaSolMNAge, gCGArendirCtaSolMEAge
                frmARendirSolicitud.inicio gArendirTipoAgencias, False
        Case gCGArendirCtaRechMNAge, gCGArendirCtaRechMEAge
                frmARendirLista.inicio gArendirTipoAgencias, ArendirRechazo
        Case gCGArendirCtaAtencMNAge, gCGArendirCtaAtencMEAge
                frmArendirAtencion.inicio gArendirTipoAgencias, ArendirAtencion
        Case gCGArendirCtaSustMNAge, gCGArendirCtaSustMEAge
                frmARendirLista.inicio gArendirTipoAgencias, ArendirSustentacion, False
        Case gCGArendirCtaRendMNAge, gCGArendirCtaRendMEAge
                frmARendirLista.inicio gArendirTipoAgencias, ArendirRendicion, False
        Case gCGArendirCtaExtAtencMNAge, gCGArendirCtaExtAtencMEAge
                frmARendirExtorno.inicio gArendirTipoAgencias, ArendirExtornoAtencion, False
        Case gCGArendirCtaExtRendMNAge, gCGArendirCtaExtRendMEAge
                frmARendirExtorno.inicio gArendirTipoAgencias, ArendirExtornoRendicion, False

        'A rendir Viaticos
        Case gCGArendirViatSolMN, gCGArendirViatSolME
                frmViaticosSol.inicio True
        Case gCGArendirViatRechMN, gCGArendirViatRechME
                frmArendirAtencion.inicio gArendirTipoViaticos, ArendirRechazo
        '***Agregado por ELRO el 20120321, según OYP-RFC005-2012
        Case gCGArendirViatAprobMN, gCGArendirViatAprobME
                frmARendirAutorizacion.iniciarAprobacionViaticos
        Case gCGArendirViatSust2MN, gCGArendirViatSust2ME
                FrmARendirSustentacion.iniciarSustentacion gArendirTipoViaticos, ArendirSustentacion
        Case gCGArendirViatRend2MN, gCGArendirViatRend2ME
                frmARendirLista.inicio gArendirTipoViaticos, ArendirRendicion, False
        '***Fin Agregado por ELRO*******************************
        Case gCGArendirViatAtencMN, gCGArendirViatAtencME
                frmArendirAtencion.inicio gArendirTipoViaticos, ArendirAtencion
        Case gCGArendirViatSustMN, gCGArendirViatSustME
                frmARendirLista.inicio gArendirTipoViaticos, ArendirSustentacion, False
        Case gCGArendirViatRendMN, gCGArendirViatRendME
                frmARendirLista.inicio gArendirTipoViaticos, ArendirRendicion, False
        Case gCGArendirViatExtAtencMN, gCGArendirViatExtAtencME
                frmARendirExtorno.inicio gArendirTipoViaticos, ArendirExtornoAtencion, False
        '***Agregado por ELRO el 20120321, según OYP-RFC005-2012
        Case gCGArendirViatExtAprobMN, gCGArendirViatExtAprobME
                frmARendirExtorno.inicio gArendirTipoViaticos, 7, False
        '***Fin Agregado por ELRO*******************************
        Case gCGArendirViatExtRendMN, gCGArendirViatExtRendME
                frmARendirExtorno.inicio gArendirTipoViaticos, ArendirExtornoRendicion, False
        Case gCGArendirViatAmpMN, gCGArendirViatAmpME
                frmViaticosSol.inicio False
                 
        '***Agregado por VAPI segun ERS1792014
        Case gCGArendirViatReimpMN, gCGArendirViatReimpME
                frmViaticoSolReimpr.Show 1
        
        '***Fin VAPI***
                 
        'Caja Chica
        Case gCHHabilitaNuevaMN, gCHHabilitaNuevaME
            frmCajaChicaHabilitacion.inicio True
        Case gCHMantenimientoMN, gCHMantenimientoME
             frmCajaChicaHabilitacion.inicio False
             
        'MIOL 20130422, SEGUN RQ13152 **************************
        Case "401306"
            frmCajaChicaCambioEncargado.Show 1
        Case "401307"
            frmCajaChicaConfEncargatura.Show 1
        'END MIOL **********************************************
        Case gCHAutorizaDesembMN, gCHAutorizaDesembME
             frmCajaChicaRendicion.inicio gCHTipoProcHabilitacion
             
        Case gCHDesembEfectivoMN, gCHDesembEfectivoME, gCHDesembOrdenPagoMN, gCHDesembOrdenPagoME, gCHDesembAbonoCtaMN, gCHDesembAbonoCtaME, gCHDesembCaChiCtaPendienteMN, gCHDesembCaChiCtaPendienteME
             frmCajaChicaEgreDirec.inicio ArendirAtencion, gCHTipoProcDesembolso
             
        Case gCHExtDesembEfectivoMN, gCHExtDesembEfectivoME, gCHExtDesembOrdenPagoMN, gCHExtDesembOrdenPagoME, gCHExtDesembCtaPendienteMN, gCHExtDesembCtaPendienteME
             frmCajaChicaEgreDirec.inicio ArendirExtornoAtencion, gCHTipoProcDesembolso
             
        '***Agregado por ELRO el 20120604, según OYP-RFC047-2012
        Case gCHExtApropacionApeMN, gCHExtApropacionApeME
             frmCajaChicaEgreDirec.inicio ArendirExtornoAtencion, gCHTipoProcDesembolso
        '***Fin Agregado por ELRO*******************************
        
        'Case gCHDesembCaChiCtaPendienteMN, gCHDesembCaChiCtaPendienteME
        '     frmCajaChicaEgreDirec.Inicio ArendirAtencion, gCHTipoProcDesembolso
        
        
        Case gCHArendirCtaSolMN, gCHArendirCtaSolME
             frmARendirSolicitud.inicio gArendirTipoCajaChica, False
        Case gCHArendirCtaRechMN, gCHArendirCtaRechME
             frmCajaChicaLista.inicio ArendirRechazo
        Case gCHArendirCtaAtencMN, gCHArendirCtaAtencME
             frmCajaChicaLista.inicio ArendirAtencion
        Case gCHArendirCtaSustMN, gCHArendirCtaSustME
             '***Modificado por ELRO el 20120604, según OYP-RFC047-2012
             'frmCajaChicaLista.Inicio ArendirSustentacion
             frmCajaChicaSustentacion.inicio ArendirSustentacion
             '***Fin Modificado por ELRO*******************************

        Case gCHArendirCtaRendMN, gCHArendirCtaRendME 'cambiado me
             frmCajaChicaLista.inicio ArendirRendicion
        Case gCHArendirCtaRendExtAtencMN, gCHArendirCtaRendExtAtencME
             frmCajaChicaLista.inicio ArendirExtornoAtencion
        '***Agregado por ELRO el 20120604, según OYP-RFC047-2012
        Case gCHArendirCtaRendExtMN2, gCHArendirCtaRendExtME2
             frmCajaChicaLista.inicio ArendirExtornoRendicion
        '***Fin Agregado por ELRO*******************************
        Case gCHArendirCtaRendExtExactMN, gCHArendirCtaRendExtExactME, _
             gCHArendirCtaRendExtIngMN, gCHArendirCtaRendExtIngME, gCHArendirCtaRendExtEgrMN, _
             gCHArendirCtaRendExtEgrME
            frmCajaChicaLista.inicio ArendirExtornoRendicion
            
        Case gCHEgreDirectoSolMN, gCHEgreDirectoSolME
            Set frmOpeDocChica = Nothing
            frmOpeDocChica.InicioEgresoDirecto
        Case gCHEgreDirectoRechMN, gCHEgreDirectoRechME
            frmCajaChicaEgreDirec.inicio ArendirRechazo
        Case gCHEgreDirectoAtencMN, gCHEgreDirectoAtencME
            frmCajaChicaEgreDirec.inicio ArendirAtencion
        Case gCHEgreDirectoExtAtencMN, gCHEgreDirectoExtAtencME
            frmCajaChicaEgreDirec.inicio ArendirExtornoAtencion
        Case gCHRendContabMN, gCHRendContabME
            frmCajaChicaRendicion.inicio gCHTipoProcRendicion
        Case gCHArqueoContabMN, gCHArqueoContabME
            frmCajaChicaArqueo.Show 1
'        Case 401396
        Case gCHRendContabExtMN, gCHRendContabExtME
            frmCajaChicaRendicion.inicio gCHTipoProcCancelada
            
        
        'boveda CajaGeneral
        Case gOpeBoveCGHabAgeMN, gOpeBoveCGHabAgeME
            frmCajaGenHabilitacion.Show 1
        Case gOpeBoveCGConfHabAgeBoveMN, gOpeBoveCGConfHabAgeBoveME
            If gbBitCentral Then
                frmCajaGenLista.Show 1
            Else
                frmCajaGenHabilitacion.Show 1
            End If
        Case gOpeBoveCGExtHabAgeMN, gOpeBoveCGExtHabAgeME
            frmCajaGenLista.Show 1
        Case gOpeBoveCGExtConfHabAgeBovMN, gOpeBoveCGExtConfHabAgeBovME
            frmCajaGenLista.Show 1
            
        'Boveda Agencia
        Case gOpeBoveAgeConfHabCGMN, gOpeBoveAgeConfHabCGME
            frmCajaGenLista.Show 1
        Case gOpeBoveAgeHabAgeACGMN, gOpeBoveAgeHabAgeACGME
            frmCajaGenHabilitacion.Show 1
        Case gOpeBoveAgeHabEntreAgeMN, gOpeBoveAgeHabEntreAgeME
            frmCajaGenHabilitacion.Show 1
            
        'Remesa con Cheques
        Case gOpeCGRemChequesMN, gOpeCGRemChequesME
            frmCajaGenRemCheques.Show 1
        
        Case gOpeCGRemChequesMantMN, gOpeCGRemChequesMantME
            FrmCajaGenRemChequeMant.Show 1
            
        'Regularizaciones de Ventanilla
        Case gOpeCGRVentanaIngresoMN, gOpeCGRVentanaIngresoME
            frmOpeRegVentanilla.Show 1
        Case gOpeCGRVentanaIngresoMNExt, gOpeCGRVentanaEgresoMNExt, _
             gOpeCGRVentanaIngresoMEExt, gOpeCGRVentanaEgresoMEExt
            frmCajaGenExtornos.Show 1
            

        'TRANSFERENCIAS
        'Se Agrego LA operacion Transferencia entre agencias GITU 01/07/2008
        Case gOpeCGTransfBancosMN, gOpeCGTransfBancosCMACSMN, gOpeCGTransfCMACSBancosMN, gOpeCGTransfMismoBancoMN, "401425", _
             gOpeCGTransfBancosME, gOpeCGTransfBancosCMACSME, gOpeCGTransfCMACSBancosME, gOpeCGTransfMismoBancoME, "402425"
            frmCajaGenTransf.Show 1
        
        Case gOpeCGTransfExtBancosMN, gOpeCGTransfExtCMACSBancosMN, gOpeCGTransfExtBancosCMACSMN, gOpeCGTransfExtMismoBancoMN, "401435", _
             gOpeCGTransfExtBancosME, gOpeCGTransfExtCMACSBancosME, gOpeCGTransfExtBancosCMACSME, gOpeCGTransfExtMismoBancoME, "402435"
            frmCajaGenExtornos.Show 1
        
        'OPERACIONES CON BANCOS
        Case gOpeCGOpeBancosDepEfecMN, gOpeCGOpeBancosDepEfecME
            frmCajaGenMovEfectivo.Show 1
            Set frmCajaGenMovEfectivo = Nothing
        Case gOpeCGOpeBancosRetEfecMN, gOpeCGOpeBancosRetEfecME
            frmCajaGenMovEfectivo.Show 1
        Case gOpeCGOpeBancosConfRetEfecMN, gOpeCGOpeBancosConfRetEfecME
            frmCajaGenExtornos.Show 1
        Case gOpeCGOpeBancosRegChequesMN, gOpeCGOpeBancosRegChequesME
            'frmIngCheques.inicio False, gsOpeCod, False, 0, Mid(gsOpeCod, 3, 1)
        Case gOpeCGOpeBancosDepChequesMN, gOpeCGOpeBancosDepChequesME
            frmCajaGenDepCheques.Show 1
        Case gOpeCGOpeBancosDepDivBancosMN, gOpeCGOpeBancosDepDivBancosME
            frmCajaGenOpeDivBancos.inicio True
        Case gOpeCGOpeBancosRetDivBancosMN, gOpeCGOpeBancosRetDivBancosME
            frmCajaGenOpeDivBancos.inicio False
        Case gOpeCGOpeBancosRetCtasBancosDetracMN, gOpeCGOpeBancosRetCtasBancosDetracME, _
             gOpeCGOpeBancosRetCtasBancosEmbargoMN, gOpeCGOpeBancosRetCtasBancosEmbargoME
            frmOpePagProvDetrac.Show 1
        'EJVG20140808 ***
        Case gOpeCGOpeBancosRemesaAgenciaMN, gOpeCGOpeBancosRemesaAgenciaMe
            frmRemesaIFiToAgencia.Inicia gsOpeCod, gsOpeDescHijo
        Case gOpeCGOpeBancosRemesaVBHabilitacion, gOpeCGOpeBancosRemesaVBHabilitacionme
            frmRemesaAprobHab.Inicia gsOpeCod, gsOpeDescHijo
        Case gOpeCGOpeBancosExtRemesaAgencia, gOpeCGOpeBancosExtRemesaAgenciaME
            frmRemesaIFiToAgenciaExt.inicio gsOpeCod, gsOpeDescHijo
        'END EJVG *******
        '*** PEAC 20091112
        Case "701450"
            frmRegVentaEmite.Show 1
        Case "701451"
            frmEmiteCheques.Show 1
        '*** FIN PEAC
        
            'JEOM
        Case gOpeCGOpeBancosOtrosDepositosMN, gOpeCGOpeBancosOtrosDepositosME
             frmCajaGenNegocioBancos.inicio True
            
        Case gOpeCGOpeBancosOtrosRetirosMN, gOpeCGOpeBancosOtrosRetirosME
             frmCajaGenNegocioBancos.inicio False
             
        Case gOpeCGOpeBancosRetFielCumplimientoMN, gOpeCGOpeBancosRetFielCumplimientoME
             frmOpePagProvDetrac.Show 1
        'FRHU 20140522 ERS068-2014 RQ14283
        Case gOpeCGOpeBancosDepositoXActivacionSegTarjetaMN
             frmSegTarjetaDepositoPorActivacion.Show 1
        'FIN FRHU 20140522
        Case gOpeCGOpeBancosRetPagoSegTarjMN, gOpeCGOpeBancosRetPagoSegTarjMNExt 'JUEZ 20140711
            frmSegTarjRetPagoBanco.inicio Mid(gsOpeCod, 1, 6)
        Case gOpeCGOpeBancosDepComSegTarjMN, gOpeCGOpeBancosDepComSegTarjMNExt 'JUEZ 20140711
            frmSegTarjDepPagoBanco.inicio Mid(gsOpeCod, 1, 6)
        Case gOpeCGOpeAperCorrienteMN, gOpeCGOpeAperAhorroMN, gOpeCGOpeAperPlazoMN, _
             gOpeCGOpeAperCorrienteME, gOpeCGOpeAperAhorroME, gOpeCGOpeAperPlazoME
            frmCajaGenAperCtas.Show 1
            Set frmCajaGenAperCtas = Nothing
        Case gOpeCGOpeConfApertMN, gOpeCGOpeConfApertME
            frmCajaGenExtornos.Show 1
        Case gOpeCGOpeIntDevPFMN, gOpeCGOpeIntDevPFME
            frmCajaIntDeveng.Show 1
        Case gOpeCGOpeGastComBancosMN, gOpeCGOpeGastComBancosME
            frmCajaGenOpeDivBancos.inicio False, False
        Case gOpeCGOpeCapIntPFMN, gOpeCGOpeCapIntPFME
            frmCajaIntDeveng.Show 1
        Case gOpeCGOpeCancCtaCteMN, gOpeCGOpeCancCtaCteME
            frmCajaIntDeveng.Show 1
        Case gOpeCGOpeCancCtaAhoMN, gOpeCGOpeCancCtaAhoME
            frmCajaIntDeveng.Show 1
        Case gOpeCGOpeCancCtaPFMN, gOpeCGOpeCancCtaPFME
            frmCajaIntDeveng.Show 1
        Case gOpeCGOpeInteresAhoCtaCteMN, gOpeCGOpeInteresAhoCtaCteME
            frmCajaGenOpeDivBancos.inicio True, False
        Case gOpeCGOpeMantCtaBancosMN, gOpeCGOpeMantCtaBancosME, gOpeCGOpeMantCtasBancosMN, gOpeCGOpeMantCtasBancosME, _
             gOpeCGOpeMantCtasBancosHaMN, gOpeCGOpeMantCtasBancosHaME, _
             gOpeCGOpeMantCtasBancosReMN, gOpeCGOpeMantCtasBancosReME
            frmCajaGenMantCtas.Show 1
            Set frmCajaGenMantCtas = Nothing
        'Extornos de Bancos
        Case gOpeCGExtBcoDepEfectivo, gOpeCGExtBcoRetEfectivo, _
             gOpeCGExtBcoConfRetEfectivo, gOpeCGExtBcoRegCheques, _
             gOpeCGExtBcoDepCheques, gOpeCGExtBcoDepDiv, _
             gOpeCGExtBcoRetDiv, gOpeCGExtBcoRecepChqRegAgencias, _
             gOpeCGExtBcoApertCta, gOpeCGExtBcoConfApert, _
             gOpeCGExtBcoIntDevengPF, gOpeCGExtBcoGastComision, _
             gOpeCGExtBcoCapitalizaIntDPF, gOpeCGExtBcoCancelaCtas, _
             gOpeCGExtBcoIntCtasAho, _
             gOpeCGExtBcoDepEfectivoME, gOpeCGExtBcoRetEfectivoME, _
             gOpeCGExtBcoConfRetEfectivoME, gOpeCGExtBcoRegChequesME, _
             gOpeCGExtBcoDepChequesME, gOpeCGExtBcoDepDivME, _
             gOpeCGExtBcoRetDivME, gOpeCGExtBcoRecepChqRegAgenciasME, _
             gOpeCGExtBcoApertCtaME, gOpeCGExtBcoConfApertME, _
             gOpeCGExtBcoIntDevengPFME, gOpeCGExtBcoGastComisionME, _
             gOpeCGExtBcoCapitalizaIntDPFME, gOpeCGExtBcoCancelaCtasME, _
             gOpeCGExtBcoIntCtasAhoME, _
             gOpeCGOpeBancosOtrosDepositosMNExt, gOpeCGOpeBancosOtrosRetirosMNExt, _
             gOpeCGOpeBancosOtrosDepositosMEExt, gOpeCGOpeBancosOtrosRetirosMEExt
            frmCajaGenExtornos.Show 1
             
        'CAJA GENERAL OPERACIONES CMACS
        Case gOpeCGOpeCMACDepDivMN, gOpeCGOpeCMACDepDivME
            frmCajaGenOpeDivBancos.inicio True
        Case gOpeCGOpeCMACRetDivMN, gOpeCGOpeCMACRetDivME
            frmCajaGenOpeDivBancos.inicio False
        Case gOpeCGOpeCMACRegularizMN, gOpeCGOpeCMACRegularizME
             frmCajaGenOpeDivBancos.inicio True, False
'        '***Modificado por ELRO el 20110923, según Acta 263-2011/TI-D
        Case gOpeCGOpeCMACDepEfeMN, gOpeCGOpeCMACDepEfeME
             frmCajaGenMovEfectivo.Show 1
        Case gOpeCGOpeCMACRetEfeMN, gOpeCGOpeCMACRetEfeME
             frmCajaGenMovEfectivo.Show 1
        Case gOpeCGOpeCMACConRetEfeMN, gOpeCGOpeCMACConRetEfeME
             frmCajaGenExtornos.Show 1
        '***fin Modificado por ELRO**********************************
        '***Modificado por ELRO el 20110930, según Acta 269-2011/TI-D
        Case gOpeCGOpeCMACExtDepEfeMN, gOpeCGOpeCMACExtDepEfeME, _
             gOpeCGOpeCMACExtRetEfeMN, gOpeCGOpeCMACExtRetEfeME, _
             gOpeCGOpeCMACExtConRetEfeMN, gOpeCGOpeCMACExtConRetEfeME
             frmCajaGenExtornos.Show 1
        '***fin Modificado por ELRO**********************************
        Case gOpeCGOpeCMACAperAhorrosMN, gOpeCGOpeCMACAperAhorrosME
            frmCajaGenAperCtas.Show 1
        Case gOpeCGOpeCMACAperPFMN, gOpeCGOpeCMACAperPFME
            frmCajaGenAperCtas.Show 1
        Case gOpeCGOpeCMACConfAperMN, gOpeCGOpeCMACConfAperME
            frmCajaGenExtornos.Show 1
        Case gOpeCGOpeCMACIntDevPFMN, gOpeCGOpeCMACIntDevPFME
            frmCajaIntDeveng.Show 1
        Case gOpeCGOpeCMACGastosComMN, gOpeCGOpeCMACGastosComME
            frmCajaGenOpeDivBancos.inicio False, False
        Case gOpeCGOpeCMACCapIntDevPFMN, gOpeCGOpeCMACCapIntDevPFME
            frmCajaIntDeveng.Show 1
        Case gOpeCGOpeCMACCancAhorrosMN, gOpeCGOpeCMACCancAhorrosME
            frmCajaIntDeveng.Show 1
        Case gOpeCGOpeCMACCancPFMN, gOpeCGOpeCMACCancPFME
            frmCajaIntDeveng.Show 1
        Case gOpeCGOpeCMACMantCtas1MN, gOpeCGOpeCMACMantCtas1ME
            frmCajaGenMantCtas.Show 1
            Set frmCajaGenMantCtas = Nothing
        Case gOpeCGOpeCMACInteresAhoMN, gOpeCGOpeCMACInteresAhoME
            frmCajaGenOpeDivBancos.inicio True, False
        Case gOpeCGOpeCMACDepDivMNExt, gOpeCGOpeCMACRetDivMNExt, _
             gOpeCGOpeCMACRegularizMNExt, gOpeCGOpeCMACAperCtasMNExt, _
             gOpeCGOpeCMACConfAperMNExt, gOpeCGOpeCMACIntDevPFMNExt, _
             gOpeCGOpeCMACGastosComMNExt, gOpeCGOpeCMACCapIntDevPFMNExt, _
             gOpeCGOpeCMACCancelaMNExt, gOpeCGOpeCMACInteresAhoMNExt, _
             gOpeCGOpeCMACDepDivMEExt, gOpeCGOpeCMACRetDivMEExt, _
             gOpeCGOpeCMACRegularizMEExt, gOpeCGOpeCMACAperCtasMEExt, _
             gOpeCGOpeCMACConfAperMEExt, gOpeCGOpeCMACIntDevPFMEExt, _
             gOpeCGOpeCMACGastosComMEExt, gOpeCGOpeCMACCapIntDevPFMEExt, _
             gOpeCGOpeCMACCancelaMEExt, gOpeCGOpeCMACInteresAhoMEExt, _
             gOpeCGOpeBancosRetCtasBancosDetracMNExt, gOpeCGOpeBancosRetCtasBancosDetracMEExt, _
             gOpeCGOpeBancosRetCtasBancosEmbargoMNExt, gOpeCGOpeBancosRetCtasBancosEmbargoMEExt, _
             gOpeCGOpeBancosRetFielCumplimientoMNExt, gOpeCGOpeBancosRetFielCumplimientoMEExt
            frmCajaGenExtornos.Show 1
        'FRHU 20140526 ERS068-2014 RQ14284
        Case gOpeCGOpeBancosDepositoXActivacionSegTarjetaMNExt
            frmSegTarjetaExtornoDepositoXActivacion.Show 1
        'FRHU 20140526 ERS068-2014
        'ADEUDADOS
        Case gOpeCGAdeudaCalendarioMN, gOpeCGAdeudaCalendarioME
            frmAdeudCal.inicio False, "", "", 0, gdFecSis
        Case gOpeCGAdeudaRegPagareMN, gOpeCGAdeudaRegPagareME
            frmCajaGenRegPAdeud.Show 1
        Case gOpeCGAdeudaRegPagareConfMN, gOpeCGAdeudaRegPagareConfMe
            frmCajaGenExtornos.Show 1
        Case gOpeCGAdeudaProvisionMN, gOpeCGAdeudaProvisionME
            frmAdeudProv1.Show 0, Me
        Case gOpeCGAdeudaPagoCuotaMN, gOpeCGAdeudaPagoCuotaME
            frmAdeudOperaciones1.inicio "D", 1
        Case gOpeCGAdeudaMntPagaresMN, gOpeCGAdeudaMntPagaresME
            frmCajaGenMantCtas.Show 1
            Set frmCajaGenMantCtas = Nothing
        Case gOpeCGAdeudaMntAjusteEurosMN, gOpeCGAdeudaMntAjusteEurosME
            frmCajaGenMantCtas.Ini True
            Set frmCajaGenMantCtas = Nothing
        Case gOpeCGAdeudaMntPagaresVinculadosMN, gOpeCGAdeudaMntPagaresVinculadosME
            frmCajaGenMantCtasVinculados.Show 1
            Set frmCajaGenMantCtas = Nothing
        Case gOpeCGAdeudaPagoCuotaLoteMN, gOpeCGAdeudaPagoCuotaLoteME
            frmAdeudPagLote.Show 0, Me
            
        'Extornos Adeudados
        Case gOpeCGAdeudaExtRegistroMN, gOpeCGAdeudaExtRegistroME, _
             gOpeCGAdeudaExtConfRegiMN, gOpeCGAdeudaExtConfRegiME, _
             gOpeCGAdeudaExtProvisiónMN, gOpeCGAdeudaExtProvisiónME, _
             gOpeCGAdeudaExtPagoCuotaMN, gOpeCGAdeudaExtPagoCuotaME, _
             gOpeCGAdeudaExtReprogramaMN, gOpeCGAdeudaExtReprogramaME
            frmCajaGenExtornos.Show 1
            
        'Caja general Moneda Extranjera
        Case gOpeMECompraAInst, gOpeMEVentaAInst
            frmCajaGenTransf.Show 1
        Case gOpeMECompraEfect, gOpeMEVentaEfec
            frmCajaGenCompraMEEfect.Show 1
        Case gOpeMEExtCompraAInst, gOpeMEExtVentaAInst, gOpeMEExtCompraEfect, gOpeMEExtVentaEfec
            frmCajaGenExtornos.Show 1
        
        'Cajero Moneda Extranjera
        'Case gOpeCajeroMETipoCambio
        '    frmMantTipoCambio.Show 1
        Case gOpeCajeroMECompra
            frmCajeroCompraVenta.Show 1
        Case gOpeCajeroMEVenta
            frmCajeroCompraVenta.Show 1
        Case gOpeCajeroMEExtCompra, gOpeCajeroMEExtVenta
            frmCajeroExtornos.Show 1
        

        '***************** PAGO DE IMPUESTOS **************************
        Case OpeCGPagoSunatBoletaPag, OpeCGPagoSunatIGVRenta, OpeCGPagoSunatSegSocial
            frmOpePagSunat.Show 1
        Case OpeCGPagoSunatSegSocialExt, OpeCGPagoSunatBoletaPagExt, OpeCGPagoSunatIGVRentaExt
            frmCajaGenExtornos.Show 1

            
        '***************** CARTAS FIANZA ********************************
        Case OpeCGCartaFianzaIng, OpeCGCartaFianzaIngME
            frmCartaFianzaIngreso.Show 1
        Case OpeCGCartaFianzaSal, OpeCGCartaFianzaSalME
            frmCartaFianzaSalida.Show 1
            
        Case OpeCGCartaFianzaIngExt, OpeCGCartaFianzaSalExt, _
             OpeCGCartaFianzaIngMEExt, OpeCGCartaFianzaSalMEExt
            frmCajaGenExtornos.Show 1
            
        '***************** OTRAS OPERACIONES DE CAJA GENERAL ********************
                    
        Case OpeCGOtrosOpeEfecIngr, OpeCGOtrosOpeEfecIngrME
        Case OpeCGOtrosOpeEfecEgre, OpeCGOtrosOpeEfecEgreME
        Case OpeCGOtrosOpeEfecCamb, OpeCGOtrosOpeEfecCambme
            frmCajaGenCompraMEEfect.Show 1
        'EJVG20140125 ***
        Case OpeCGOtrosOpeChequeValorizacion, OpeCGOtrosOpeChequeValorizacionME
            frmChequeNegocioValorizacion.inicio gsOpeCod, gsOpeDescHijo
        Case OpeCGOtrosOpeChequeExtValorizacion, OpeCGOtrosOpeChequeExtValorizacionME, OpeCGOtrosOpeChequeExtRechazo, OpeCGOtrosOpeChequeExtRechazoME
            frmChequeNegocioValorizacionExtorno.inicio gsOpeCod, gsOpeDescHijo
        'END EJVG *******
        Case OpeCGOtrosOpeEfecOtro, OpeCGOtrosOpeEfecOtroME
            frmAsientoRegistro.inicio "", -1, False, True
        
        Case OpeCGOtrosOpeEfecIngrExt, OpeCGOtrosOpeEfecEgreExt, _
             OpeCGOtrosOpeEfecCambExt, OpeCGOtrosOpeEfecOtroExt, _
             OpeCGOtrosOpeEfecIngrMEExt, OpeCGOtrosOpeEfecEgreMEExt, _
             OpeCGOtrosOpeEfecCambMEExt, OpeCGOtrosOpeEfecOtroMEExt
            frmCajaGenExtornos.Show 1
        
        'JEOM
        Case OpeCGOpeProvPago, OpeCGOpeProvPagoME
            'frmOpePagProv.Show , Me
            'frmOpePagProv.Ini False, "PAGO A PROVEEDORES"
            frmOpePagoProv_NEW.Show 1 'EJVG20131209
            
        Case OpeCGOpeProvRechazo, OpeCGOpeProvRechazoME
             frmOpePagProv.Show , Me
             
        'Embargo
        Case OpeCGOpeProvPagoListSUNAT
            frmOpePagProv.Ini True, "LISTADO CONSULTA SUNAT"
        Case OpeCGOpeProvPagoRegSUNAT
            frmACGMntProveedorValidacion.Show 1
        Case OpeCGOpeProvPagoPagUNAT
            frmACGMntProveedorPagoSunat.Show 1
            
        'FIN
            

        Case OpeCGOpeProvEntrOP, OpeCGOpeProvEntrOPME, OpeCGOpeProvEntrCh, OpeCGOpeProvEntrChME
            frmOpePagProvEntrega.Show , Me
            
        Case gOpeCGExtProvPago, gOpeCGExtProvPagoEfectivo, gOpeCGExtProvPagoTransfer, _
             gOpeCGExtProvPagoOPago, gOpeCGExtProvPagoCheque, _
             gOpeCGExtProvPagoAbono, gOpeCGExtProvPagoRechazo, _
             gOpeCGExtProvEntrOPago, gOpeCGExtProvEntrCheques, _
             gOpeCGExtProvPagoME, gOpeCGExtProvPagoEfectivoME, gOpeCGExtProvPagoTransferME, _
             gOpeCGExtProvPagoOPagoME, gOpeCGExtProvPagoChequeME, _
             gOpeCGExtProvPagoAbonoME, gOpeCGExtProvPagoRechazoME, _
             gOpeCGExtProvEntrOPagoME, gOpeCGExtProvEntrChequesME, gOpeCGExtProvPagoSUNAT, _
             gOpeCGExtProvDevSUNAT, _
            OpeCGOtrosOpeRetPagSeguroDesgravamenMNExt, OpeCGOtrosOpeRetPagSeguroIncendioMNExt, _
             OpeCGOtrosOpeRetPagSeguroDesgravamenMEExt, OpeCGOtrosOpeRetPagSeguroIncendioMEExt
             'PASIERS1242014 agregó:gOpeCGExtProvDevSUNAT
             'PASIERS1362014 agregó:OpeCGOtrosOpeRetPagSeguroDesgravamenMNExt,OpeCGOtrosOpeRetPagSeguroIncendioMNExt,OpeCGOtrosOpeRetPagSeguroDesgravamenMEExt,OpeCGOtrosOpeRetPagSeguroIncendioMEExt
             frmCajaGenExtornos.Show 1
        '*******************INVERSIONES********************************
        
'        'JACA 20110913**************************
        Case "421301", "422301"
             frmInversiones.Show 1
        Case "421302", "422302"
             frmInversionesConfirmacion.Show 1
        Case "421303", "422303"
             frmInversionesCancelacion.Show 1
        Case "421304", "422304"
             frmInversionesProvision.Show 1
        Case "421306", "422306", "421307", "422307", _
             "421308", "422308", "421309", "422309"
             frmInversionesExtorno.Show 1
        Case "421310", "422310"
             frmInversionesMantenimiento.Show 1
        'JACA END*******************************
        
        'MIOL 20130304, ERS025 *****************
        Case "421401", "422401"
             frmRegLiquidezPotencial.Show 1
        Case "421402", "422402"
             frmMantLiquidezPotencial.Show 1
        'END MIOL ******************************
             
        '**************** PRESUPUESTO ADEUDADOS *************
        Case gOpeCGAdeudaPresuProyecMN, gOpeCGAdeudaPresuProyecME
            frmPresGenPlanPend.Show , Me
        
        '********************* REPORTES DE CAJA GENERAL **************************
        
        Case OpeCGRepRepBancosFlujoMensMN, OpeCGRepRepBancosFlujoMensME
            frmCajaGenRepFlujos.Show , Me
        Case OpeCGRepRepBancosFlujoPFMN, OpeCGRepRepBancosFlujoPFME
            frmCajaGenRepFlujos.Show , Me
        
        Case OpeCGRepRepCMACSFlujoMensMN, OpeCGRepRepCMACSFlujoMensME
            frmCajaGenRepFlujos.Show , Me
        Case OpeCGRepRepCMACSFlujoPFMN, OpeCGRepRepCMACSFlujoPFME
            frmCajaGenRepFlujos.Show , Me
        
        Case OpeCGRepRepOPGirMN, OpeCGRepRepOPGirME
        
        Case OpeCGRepRepChqRecDetMN, OpeCGRepRepChqRecDetME
        Case OpeCGRepRepChqRecResMN, OpeCGRepRepChqRecResME
        
        Case OpeCGRepRepChqValDetMN, OpeCGRepRepChqValDetME
        Case OpeCGRepRepChqValResMN, OpeCGRepRepChqValResME
        
        Case OpeCGRepRepChqValorizadosDetMN, OpeCGRepRepChqValorizadosDetME
        Case OpeCGRepRepChqValorizadosResMN, OpeCGRepRepChqValorizadosResME
        
        Case OpeCGRepRepChqAnulDetMN, OpeCGRepRepChqAnulDetME
        Case OpeCGRepRepChqAnulResMN, OpeCGRepRepChqAnulResME

        Case OpeCGRepRepChqObsDetMN, OpeCGRepRepChqObsDetME
        Case OpeCGRepRepChqObsResMN, OpeCGRepRepChqObsResME
        
      'Operaciones con Pendientes
      'Alquileres
      Case gAnaRegGarantía, gAnaRegGarantíaME
        frmLogProvisionPago.inicio False, False, "", True, True, , True
      Case gAnaDevGarantía, gAnaDevGarantíaME
        frmAnalisisRegulaPend.inicio "H", True, False, True, True, True, , , , , True
      Case gAnaDevGarantíaGasto, gAnaDevGarantíaGastoME
        frmAnalisisRegulaPend.inicio "H", False, False, True, True, True, , , 1

      'Demandas
      Case gAnaDemandaIngr, gAnaDemandaIngrME
        frmAsientoRegistro.inicio "", 0, , True, True, True
      Case gAnaDemandaDescargo, gAnaDemandaDescargoME
        frmAnalisisRegulaPend.inicio "H", True, False, True, True, True, , , , , True
      Case gAnaDemandaDescargoGasto, gAnaDemandaDescargoGastoME
        frmAnalisisRegulaPend.inicio "H", False, False, True, True, True, , , 1
      
      'SubSidio Pre y Post Natal
      Case gAnaSubsidioPrePostAbonoSeguro, gAnaSubsidioPrePostAbonoSeguroME
        frmAnalisisRegulaPend.inicio "H", True, False, False, True, True, , False

      'SubSidio Enfermedad
      Case gAnaSubsidioEnfermAbonoSeguro, gAnaSubsidioEnfermAbonoSeguroME
        frmAnalisisRegulaPend.inicio "H", True, False, False, True, True, , False

      'Siniestros
      Case gAnaIndemnizaSiniestroRegistro, gAnaIndemnizaSiniestroRegistroME
        frmAsientoRegistro.inicio "", 0, , True, True, True
      Case gAnaIndemnizaPago, gAnaIndemnizaPagoME
        frmAnalisisRegulaPend.inicio "H", True, False, False, True, True, , , , , True
      Case gAnaIndemnizaPagoGasto, gAnaIndemnizaPagoGastoME
        frmAnalisisRegulaPend.inicio "H", False, False, True, True, True, , , 1
    
      'Otras Cuentas por Cobrar
      Case gAnaOtrasCtaCobrarRegistroAsiento, gAnaOtrasCtaCobrarRegistroAsientoME
        frmAsientoRegistro.inicio "", 0, , True, True, True
      Case gAnaOtrasCtaCobrarProvisPago, gAnaOtrasCtaCobrarProvisPagoME
        frmLogProvisionPago.inicio False, False, "", True, True
      Case gAnaOtrasCtaCobrarCancelaAsiento, gAnaOtrasCtaCobrarCancelaAsientoME
        frmAnalisisRegulaPend.inicio "H", False, False, False, False, False, False, False, 2
      Case gAnaOtrasCtaCobrarCancelaDeuda, gAnaOtrasCtaCobrarCancelaDeudaME
        frmAnalisisRegulaPend.inicio "H", True, False, True, True, True, False, , , , True

      'Otras Operaciones de Caja General
      Case gAnaOtraOpeCGRegistroAsiento, gAnaOtraOpeCGRegistroAsientoME
        frmAsientoRegistro.inicio "", 0, , True, True, True
      Case gAnaOtraOpeCGRegulaDocSustentatorio, gAnaOtraOpeCGRegulaDocSustentatorioME
        frmAnalisisRegulaPend.inicio "H", False, False, False, False, False, False, True, 1
      Case gAnaOtraOpeCGRegulaAsiento, gAnaOtraOpeCGRegulaAsientoME
        frmAnalisisRegulaPend.inicio "H", False, False, False, False, False, False, True, 2
      Case gAnaOtraOpeCGRegulaCajaGeneral, gAnaOtraOpeCGRegulaCajaGeneralME
        frmAnalisisRegulaPend.inicio "H", True, False, False, True, True, True, True, 0, , True
    
      'Otras Provisiones
      Case gAnaOtraProvisRegistProvision, gAnaOtraProvisRegistProvisionME
        frmAsientoRegistro.inicio "", 0, , True, True, True
      Case gAnaOtraProvisRegulaProvision, gAnaOtraProvisRegulaProvisionME
        frmAnalisisRegulaPend.inicio "D", False, False, False, False, False, False, False, 1
      Case gAnaOtraProvisRegulaProvAsiento, gAnaOtraProvisRegulaProvAsientoME
        frmAnalisisRegulaPend.inicio "D", False, False, False, False, False, False, False, 2
         
      'Canje de Cheque
      Case gAnaCanjeOPRegulariza, gAnaCanjeOPRegularizaME
        frmAnalisisRegulaPend.inicio "D", False, False, False, False, False, False, False, 2
      
      'Transferencias
      Case gAnaTranferRegulariza, gAnaTranferRegularizaME
        frmAnalisisRegulaPend.inicio "D", False, False, False, False, False, False, False, 2

      'Operaciones en Tramite de Caja General
      Case gAnaOtrasOpeLiqRegistAsiento, gAnaOtrasOpeLiqRegistAsientoME
        frmAsientoRegistro.inicio "", 0, , True, True, True
      
      Case gAnaOtrasOpeLiqRegulaAsiento, gAnaOtrasOpeLiqRegulaAsientoME
        frmAnalisisRegulaPend.inicio "D", False, False, False, False, False, False, True, 2
      Case gAnaOtrasOpeLiqRegulaCajaGen, gAnaOtrasOpeLiqRegulaCajaGenME
        frmAnalisisRegulaPend.inicio "D", True, False, True, True, True, True, True, 0, , False
        
      'Recursos Humanos
      Case gAnaRecursosHumRegulaAsiento, gAnaRecursosHumRegulaAsientoME
        frmAnalisisRegulaPend.inicio "D", False, False, False, False, False, False, False, 2
        
        
        '************************* CONTABILIDAD **********************************
        'Provision de Pago a Proveedores
        Case gContProvOrdenCompraMN, gContProvOrdenCompraME
            frmLogProvisionSeleccion.inicio True
        Case gContProvOrdenServicMN, gContProvOrdenServicME
            frmLogProvisionSeleccion.inicio False
        Case gContProvDirectaMN, gContProvDirectaME
        'ALPA 20090303 ***********************************************************
            'frmLogProvisionPago.Inicio False, False, ""
            frmLogProvisionPago.inicio False, False, "", , , , , True
        '************************************************************************
        Case gContProvDirectaOCMN, gContProvDirectaOCME
            frmLogProvisionPago.inicio False, False, ""
        Case gContProvDirectaRRHHMN, gContProvDirectaRRHHME
            frmLogProvisionPago.inicio False, False, ""
        'EJVG20131113 ***
        Case gContProvLogComprobanteMN, gContProvLogComprobanteME
            frmLogProvisionPago.inicio False, False, "", , , , , True
        'END EJVG *******
        
        'Provision de Cartera de Creditos
        Case gContProvCarteraCredMN, gContProvCarteraCredME
        
        Case gContLibroDiario
               frmContabDiario.Show , Me
        Case gContLibroMayor
               frmContabMayor.Show , Me
        Case gContLibroMayCta
               frmContabMayorDet.Show , Me
        Case gContRegCompraGastos
               frmRegCompraGastos.Show , Me
        Case gContRegVentas
               frmRegVenta.Show , Me
        Case gContRepBaseFormula
            frmRepBaseFormula.Show , Me
        Case gContRepCompraVenta
            frmRepResCVenta.Show , Me
        Case gHistoContLibroMayCta 'ALPA 20111229
            frmHistoContabMayorDet.Show , Me

        'Otros Ajustes
        Case gContAjReclasiCartera, gContAjReclasiCarteraME
            frmAjusteReCartera.Show 0, Me
        Case gContAjReclasiGaranti, gContAjReclasiGarantiME
            frmAjusteGarantias.Show 0, Me
        Case gContAjInteresDevenga, gContAjInteresDevengaME
            frmAjusteIntDevengado.inicio 1
        Case gContAjInteresSuspens, gContAjInteresSuspensME
            frmAjusteIntDevengado.inicio 2
        Case gContCapCreditoCastig, gContCapCreditoCastigME
            frmAjusteIntDevengado.inicio 5
        Case gContIntCreditoCastig, gContIntCreditoCastigME
            frmAjusteIntDevengado.inicio 6
        Case gContAsProvisionCarte, gContAsProvisionCarteME
            frmAjusteIntDevengado.inicio 3
        Case gContAsProvisionCFianza, gContAsProvisionCFianzaME 'ASIENTOS POR CALIFICACION DE CREDITO
            'frmAjusteIntDevengado.Inicio 8
            frmAjusteIntDevengado.inicio 12
        Case gContRevInteresDeveng, gContRevInteresDevengME
            frmAjusteIntDevengado.inicio 7
        Case gContAsCalificaCredConting, gContAsCalificaCredContingME 'ASIENTOS POR RIESGO PONDERADO
            'frmAjusteIntDevengado.Inicio 9
            frmAjusteIntDevengado.inicio 13
        Case gContAsProvisionCConting, gContAsProvisionCContingME
            'frmAjusteIntDevengado.Inicio 10 'JOMARK
            frmAjusteIntDevengado.inicio 14
        Case "701237", "702237"
            'frmAjusteIntDevengado.Inicio 10 'JOMARK
            frmAjusteIntDevengado.inicio 15
       Case "701238", "702238"
            frmAjusteIntDevengado.inicio 16
'        Case gContAsRevIntCalifDudosoPerdida, gContAsRevIntCalifDudosoPerdidaME
'            frmAjusteIntDevengado.Inicio 11
'            'Set frmAjusteIntDevengado = Nothing
'        Case gContAsRevIntCredRefinan, gContAsRevIntCredRefinanME
'            frmAjusteIntDevengado.Inicio 12
'            'Set frmAjusteIntDevengado = Nothing
        'ALPA 20090529******************************
        Case "701239", "702239"
            frmAjusteIntDevengado.inicio 17
        
        '*******************************************
        'ALPA 20090529******************************
        Case "701240", "702240"
            frmAjusteIntDevengado.inicio 18
        Case "701241", "702241"
            frmAjusteIntDevengado.inicio 19
        Case "701243", "702243"
            frmAjusteIntDevengado.inicio 21
        '*******************************************
        'JUEZ 20130116 *****************************
        Case gContGastoCreditoCastig, gContGastoCreditoCastigME
            frmAjusteIntDevengado.inicio 22
        'END JUEZ **********************************
        Case gContAjComisionCF, gContAjComisionCFME 'EJVG20130322
            frmAjusteIntDevengado.inicio 23
        Case "701246", "702246" 'ALPA20131129
            frmAjusteIntDevengado.inicio 24
        'ALPA 20131202****************************
        Case "701247", "702247" 'ALPA20140925
            frmAjusteIntDevengado.inicio 25
        Case 701248, 702248
            frmAjusteIntDevengado.inicio 26
        '*****************************************
        'PASIERS0142015****************************************
        Case 701249, 702249
            frmAjusteDepreAdjudicado.inicio "701249"
        'END PASI**************************************************
        'PASI20160216 ERS0762015
        Case 701250
            frmAjusteGramosDeOroEnCustodia.inicio "701250"
        'END PASI
        Case "701245"
        
        'PASIERS0152017 20170104****************************
        Case 701251, 702251
            frmAjusteIntDevengado.inicio 28
        'PASI END ******************************************
        
        'NAGL 202007 Según ACTA N°049-2020******************
        Case 701253, 702253
            frmAjusteIntDevengado.inicio 29
        'NAGL END ******************************************
        
        'NAGL 202007 Según ACTA N°049-2020******************
        Case 701254, 702254
            frmAjusteIntDevengado.inicio 30
        'NAGL END ******************************************
        
        'NAGL 202007 Según ACTA N°049-2020******************
        Case 701255, 702255
            frmAjusteIntDevengado.inicio 31
        'NAGL END ******************************************
        
        'NAGL 202008 Según ACTA N°063-2020******************
        Case 701256, 702256
            frmAjusteIntDevengado.inicio 32
        'NAGL END ******************************************
        
        'NAGL 202008 Según ACTA N°063-2020******************
        Case 701257, 702257
            frmAjusteIntDevengado.inicio 33
        'NAGL END ******************************************
        
        'NAGL 202102 Según ACTA N°017-2021******************
        Case 701258, 702258
            frmAjusteIntDevengado.inicio 34
        'NAGL END ******************************************
        
        'FONCODES
        Case gContFoncodesNAbono, gContFoncodesNAbonoME
            
        Case gContFoncodesNCargo, gContFoncodesNCargoME

        'ANEXOS
        Case gContAnx07
            frmAnexo7RiesgoInteres.inicio True
        
        Case gContRegistroAsiento, gContRegistroAsientoME
             frmAsientoRegistro.inicio "", False
        Case gContRegAsientoMoneda, gContRegAsientoMonedaME
             frmAsientoRegistro.inicio "", False, , True
        Case gOpeBancoyCajasMN, gOpeBancoyCajasME
            Call frmBanyOtrasInstSisFinan.inicio(gsOpeCod)
        Case OpeCGOpeProvDevSUNAT 'PASIERS1242014
            frmMntProveedorDevolucionSunat.Show 1
        Case OpeCGOtrosOpeRetPagSeguroDesgravamenMN, OpeCGOtrosOpeRetPagSeguroIncendioMN, _
                OpeCGOtrosOpeRetPagSeguroDesgravamenME, OpeCGOtrosOpeRetPagSeguroIncendioME 'PASIERS1362014
            Call frmSegDesgravamenRetPago.inicio(gsOpeCod)
        'APRI2018 ERS028-2017
        Case "421008", "422008", "421018", "422018"
            frmSegTransferenciaPrimas.inicio Mid(gsOpeCod, 1, 6)
        'END APRI
    End Select
    
    If Mid(gsOpeCod, 1, 6) = "701228" Or Mid(gsOpeCod, 1, 6) = "702228" Then
        frmAjusteIntDevengado.inicio 11
        
    End If
    
    'MADM 20110805
    If Mid(gsOpeCod, 1, 6) = "701242" Or Mid(gsOpeCod, 1, 6) = "702242" Then
        frmAjusteIntDevengado.inicio 20
    End If
    'END MADM
    
    
    Exit Sub
ErrorAceptar:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso Error"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
tvOpe.SetFocus
Err.Clear
End Sub

Private Sub Form_Load()
    Dim sCod As String
    On Error GoTo ERROR
    CentraForm Me
    frmMdiMain.Enabled = False
    If Dir(App.path & "\videos\LogoA.avi") <> "" Then
        Logo.AutoPlay = True
        Logo.Open App.path & "\videos\LogoA.avi"
    End If
    If Not lExpandO Then
       Dim oConst As New NConstSistemas
       sCod = oConst.LeeConstSistema(gConstSistContraerListaOpe)
       If sCod <> "" Then
         lExpand = IIf(UCase(Trim(sCod)) = "FALSE", False, True)
       End If
       Set oConst = Nothing
    Else
       lExpand = lExpandO
    End If
    LoadOpeUsu "2"
    Exit Sub
ERROR:
    MsgBox TextErr(Err.Description), vbExclamation, Me.Caption
End Sub

Sub LoadOpeUsu(psMoneda As String)
Dim clsGen As DGeneral
Dim rsUsu As Recordset
Dim sOperacion As String, sOpeCod As String
Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String
Dim nodOpe As Node

Set clsGen = New DGeneral
Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, sArea, MatOperac, NroRegOpe, psMoneda)

Set clsGen = Nothing
tvOpe.Nodes.Clear
Do While Not rsUsu.EOF
    sOpeCod = rsUsu("cOpeCod")
    sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
    Select Case rsUsu("nOpeNiv")
        Case "1"
            sOpePadre = "P" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
            nodOpe.Tag = sOpeCod
        Case "2"
            sOpeHijo = "H" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
            nodOpe.Tag = sOpeCod
        Case "3"
            sOpeHijito = "J" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
            nodOpe.Tag = sOpeCod
        Case "4"
            Set nodOpe = tvOpe.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
            nodOpe.Tag = sOpeCod
    End Select
    nodOpe.Expanded = lExpand
    rsUsu.MoveNext
Loop
RSClose rsUsu
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMdiMain.Enabled = True
End Sub

Private Sub OptMoneda_Click(index As Integer)
    Dim sDig As String
    Dim sCod As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    On Error GoTo ERROR
    If optMoneda(0) Then
        sDig = "2"
        gsSimbolo = gcMN
    Else
        sDig = "1"
        gsSimbolo = gcME
    End If
    oConec.AbreConexion
    LoadOpeUsu sDig
    oConec.CierraConexion
    tvOpe.SetFocus
    Exit Sub
ERROR:
    MsgBox TextErr(Err.Description), vbExclamation, Me.Caption
End Sub

Private Sub tvOpe_Collapse(ByVal Node As MSComctlLib.Node)
    Node.ForeColor = "&H80000008"
End Sub

Private Sub tvOpe_DblClick()
    If tvOpe.Nodes.Count > 0 Then
       cmdAceptar_Click
    End If
End Sub

Private Sub tvOpe_Expand(ByVal Node As MSComctlLib.Node)
    Node.ForeColor = "&H8000000D"
End Sub

Private Sub tvOpe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdAceptar_Click
    End If
End Sub
