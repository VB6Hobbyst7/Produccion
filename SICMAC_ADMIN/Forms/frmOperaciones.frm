VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOperaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proceso de Contabilidad: Selección de Operaciones"
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
   StartUpPosition =   2  'CenterScreen
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
      Enabled         =   -1  'True
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

Public Sub Inicio(sObj As String, Optional plExpandO As Boolean = False)
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
    gsOpeCod = Mid(tvOpe.SelectedItem.Text, 1, 6)
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
                frmARendirSolicitud.Inicio gArendirTipoCajaGeneral, False
        Case gCGArendirCtaRechMN, gCGArendirCtaRechME
                frmArendirAtencion.Inicio gArendirTipoCajaGeneral, ArendirRechazo
        Case gCGArendirCtaAtencMN, gCGArendirCtaAtencME
                frmArendirAtencion.Inicio gArendirTipoCajaGeneral, ArendirAtencion
        Case gCGArendirCtaSustMN, gCGArendirCtaSustME
                frmARendirLista.Inicio gArendirTipoCajaGeneral, ArendirSustentacion, False
        Case gCGArendirCtaRendMN, gCGArendirCtaRendME
                frmARendirLista.Inicio gArendirTipoCajaGeneral, ArendirRendicion, False
        Case gCGArendirCtaExtAtencMN, gCGArendirCtaExtAtencME
                frmARendirExtorno.Inicio gArendirTipoCajaGeneral, ArendirExtornoAtencion, False
        Case gCGArendirCtaExtRendMN, gCGArendirCtaExtRendME
                frmARendirExtorno.Inicio gArendirTipoCajaGeneral, ArendirExtornoRendicion, False
        'A rendir Viaticos
        Case gCGArendirViatSolMN, gCGArendirViatSolME
                frmViaticosSol.Inicio True
        Case gCGArendirViatRechMN, gCGArendirViatRechME
                frmArendirAtencion.Inicio gArendirTipoViaticos, ArendirRechazo
        Case gCGArendirViatAtencMN, gCGArendirViatAtencME
                frmArendirAtencion.Inicio gArendirTipoViaticos, ArendirAtencion
        Case gCGArendirViatSustMN, gCGArendirViatSustME
                frmARendirLista.Inicio gArendirTipoViaticos, ArendirSustentacion, False
        Case gCGArendirViatRendMN, gCGArendirViatRendME
                frmARendirLista.Inicio gArendirTipoViaticos, ArendirRendicion, False
        Case gCGArendirViatExtAtencMN, gCGArendirViatExtAtencME
                frmARendirExtorno.Inicio gArendirTipoViaticos, ArendirExtornoAtencion, False
        Case gCGArendirViatExtRendMN, gCGArendirViatExtRendME
                frmARendirExtorno.Inicio gArendirTipoViaticos, ArendirExtornoRendicion, False
        Case gCGArendirViatAmpMN, gCGArendirViatAmpME
                frmViaticosSol.Inicio False
        'Caja Chica
        
        Case gCHHabilitaNuevaMN, gCHHabilitaNuevaME
            frmCajaChicaHabilitacion.Inicio True
        Case gCHMantenimientoMN, gCHMantenimientoME
            frmCajaChicaHabilitacion.Inicio False
        
        Case gCHAutorizaDesembMN, gCHAutorizaDesembME
            frmCajaChicaRendicion.Inicio gCHTipoProcHabilitacion
        Case gCHDesembEfectivoMN, gCHDesembEfectivoME, gCHDesembOrdenPagoMN, gCHDesembOrdenPagoME
            frmCajaChicaEgreDirec.Inicio ArendirAtencion, gCHTipoProcDesembolso
        Case gCHExtDesembEfectivoMN, gCHExtDesembEfectivoMN, gCHExtDesembOrdenPagoMN, gCHExtDesembOrdenPagoME
            frmCajaChicaEgreDirec.Inicio ArendirExtornoAtencion, gCHTipoProcDesembolso
        
        Case gCHArendirCtaSolMN, gCHArendirCtaSolME
                frmARendirSolicitud.Inicio gArendirTipoCajaChica, False
        Case gCHArendirCtaRechMN, gCHArendirCtaRechME
                frmCajaChicaLista.Inicio ArendirRechazo
        Case gCHArendirCtaAtencMN, gCHArendirCtaAtencME
                frmCajaChicaLista.Inicio ArendirAtencion
        Case gCHArendirCtaSustMN, gCHArendirCtaSustME
                frmCajaChicaLista.Inicio ArendirSustentacion
        Case gCHArendirCtaRendMN, gCHArendirCtaRendMN
            frmCajaChicaLista.Inicio ArendirRendicion
        Case gCHArendirCtaRendExtAtencMN, gCHArendirCtaRendExtAtencME
            frmCajaChicaLista.Inicio ArendirExtornoAtencion
            
        Case gCHArendirCtaRendExtExactMN, gCHArendirCtaRendExtExactME, _
             gCHArendirCtaRendExtIngMN, gCHArendirCtaRendExtIngME, gCHArendirCtaRendExtEgrMN, _
             gCHArendirCtaRendExtEgrME
            
            frmCajaChicaLista.Inicio ArendirExtornoRendicion
        Case gCHEgreDirectoSolMN, gCHEgreDirectoSolME
            Set frmOpeDocChica = Nothing
            frmOpeDocChica.InicioEgresoDirecto
        Case gCHEgreDirectoRechMN, gCHEgreDirectoRechME
            frmCajaChicaEgreDirec.Inicio ArendirRechazo
        Case gCHEgreDirectoAtencMN, gCHEgreDirectoAtencME
            frmCajaChicaEgreDirec.Inicio ArendirAtencion
        Case gCHEgreDirectoExtAtencMN, gCHEgreDirectoExtAtencME
            frmCajaChicaEgreDirec.Inicio ArendirExtornoAtencion
        Case gCHRendContabMN, gCHRendContabME
            frmCajaChicaRendicion.Inicio gCHTipoProcRendicion
        Case gCHArqueoContabMN, gCHArqueoContabME
            frmCajaChicaArqueo.Show 1
        
        'boveda CajaGeneral
        Case gOpeBoveCGHabAgeMN, gOpeBoveCGHabAgeMN
            frmCajaGenHabilitacion.Show 1
        Case gOpeBoveCGConfHabAgeBoveMN, gOpeBoveCGConfHabAgeBoveME
            frmCajaGenLista.Show 1
        Case gOpeBoveCGExtHabAgeMN, gOpeBoveCGExtHabAgeMN
            frmCajaGenLista.Show 1
        Case gOpeBoveCGExtConfHabAgeBovMN, gOpeBoveCGExtConfHabAgeBovME
            frmCajaGenLista.Show 1
        'Boveda Agencia
        Case gOpeBoveAgeConfHabCGMN, gOpeBoveAgeConfHabCGMN
            frmCajaGenLista.Show 1
        Case gOpeBoveAgeHabAgeACGMN, gOpeBoveAgeHabAgeACGME
            frmCajaGenHabilitacion.Show 1
        Case gOpeBoveAgeHabEntreAgeMN, gOpeBoveAgeHabEntreAgeME
            frmCajaGenHabilitacion.Show 1
        Case gOpeBoveAgeHabCajeroMN, gOpeBoveAgeHabCajeroME
            frmCajeroHab.Show 1
        Case gOpeBoveAgeExtConfHabCGMN, gOpeBoveAgeExtConfHabCGME, _
            gOpeBoveAgeExtHabAgeACGMN, gOpeBoveAgeExtHabAgeACGME, _
            gOpeBoveAgeExtHabEntreAgeMN, gOpeBoveAgeExtHabEntreAgeMe
            frmCajaGenLista.Show 1
        Case gOpeBoveAgeExtHabCajeroMN, gOpeBoveAgeExtHabCajeroME
            frmCajeroExtornos.Show 1
    
        'TRANSFERENCIAS
        Case gOpeCGTransfBancosMN, gOpeCGTransfBancosCMACSMN, gOpeCGTransfCMACSBancosMN, gOpeCGTransfMismoBancoMN, _
             gOpeCGTransfBancosME, gOpeCGTransfBancosCMACSME, gOpeCGTransfCMACSBancosME, gOpeCGTransfMismoBancoME
            frmCajaGenTransf.Show 1
        
        Case gOpeCGTransfExtBancosMN, gOpeCGTransfExtCMACSBancosMN, gOpeCGTransfExtBancosCMACSMN, gOpeCGTransfExtMismoBancoMN, _
            gOpeCGTransfExtBancosME, gOpeCGTransfExtCMACSBancosME, gOpeCGTransfExtBancosCMACSME, gOpeCGTransfExtMismoBancoME
            frmCajaGenExtornos.Show 1
        
        'OPERACIONES CON BANCOS
        Case gOpeCGOpeBancosDepEfecMN, gOpeCGOpeBancosDepEfecME
            frmCajaGenMovEfectivo.Show 1
        Case gOpeCGOpeBancosRetEfecMN, gOpeCGOpeBancosRetEfecME
            frmCajaGenMovEfectivo.Show 1
        Case gOpeCGOpeBancosConfRetEfecMN, gOpeCGOpeBancosConfRetEfecME
            frmCajaGenExtornos.Show 1
        Case gOpeCGOpeBancosRegChequesMN, gOpeCGOpeBancosRegChequesME
            frmIngCheques.Inicio False, gsOpeCod, False, 0, Mid(gsOpeCod, 3, 1)
        Case gOpeCGOpeBancosDepChequesMN, gOpeCGOpeBancosDepChequesME
            frmCajaGenDepCheques.Show 1
        Case gOpeCGOpeBancosDepDivBancosMN, gOpeCGOpeBancosDepDivBancosME
            frmCajaGenOpeDivBancos.Inicio True
        Case gOpeCGOpeBancosRetDivBancosMN, gOpeCGOpeBancosRetDivBancosME
            frmCajaGenOpeDivBancos.Inicio False
        Case gOpeCGOpeAperCorrienteMN, gOpeCGOpeAperAhorroMN, gOpeCGOpeAperPlazoMN, _
            gOpeCGOpeAperCorrienteME, gOpeCGOpeAperAhorroME, gOpeCGOpeAperPlazoME
            frmCajaGenAperCtas.Show 1
        Case gOpeCGOpeConfApertMN, gOpeCGOpeConfApertME
            frmCajaGenExtornos.Show 1
        Case gOpeCGOpeIntDevPFMN, gOpeCGOpeIntDevPFME
            frmCajaIntDeveng.Show 1
        Case gOpeCGOpeGastComBancosMN, gOpeCGOpeGastComBancosME
            frmCajaGenOpeDivBancos.Inicio False
        Case gOpeCGOpeCapIntPFMN, gOpeCGOpeCapIntPFME
            frmCajaIntDeveng.Show 1
        Case gOpeCGOpeCancCtaCteMN, gOpeCGOpeCancCtaCteME
            frmCajaIntDeveng.Show 1
        Case gOpeCGOpeCancCtaAhoMN, gOpeCGOpeCancCtaAhoME
            frmCajaIntDeveng.Show 1
        Case gOpeCGOpeCancCtaPFMN, gOpeCGOpeCancCtaPFME
            frmCajaIntDeveng.Show 1
        Case gOpeCGOpeMantCtaBancosMN, gOpeCGOpeMantCtaBancosME
            frmCajaGenmantCtas.Show 1
        
        'CAJA GENERAL OPERACIONES CMACS
        Case OpeCGOpeCMACDepDivMN, OpeCGOpeCMACDepDivME
            frmCajaGenOpeDivBancos.Inicio True
        Case OpeCGOpeCMACRetDivMN, OpeCGOpeCMACRetDivME
            frmCajaGenOpeDivBancos.Inicio False
        Case OpeCGOpeCMACRegularizMN, OpeCGOpeCMACRegularizME
            frmCajaGenOpeDivBancos.Inicio
        Case OpeCGOpeCMACAperAhorrosMN, OpeCGOpeCMACAperAhorrosME
            frmCajaGenAperCtas.Show 1
        Case OpeCGOpeCMACAperPFMN, OpeCGOpeCMACAperPFME
            frmCajaGenAperCtas.Show 1
        Case OpeCGOpeCMACConfAperMN, OpeCGOpeCMACConfAperME
            frmCajaGenExtornos.Show 1
        Case OpeCGOpeCMACIntDevPFMN, OpeCGOpeCMACIntDevPFME
            frmCajaIntDeveng.Show 1
        Case OpeCGOpeCMACGastosComMN, OpeCGOpeCMACGastosComME
            frmCajaGenOpeDivBancos.Inicio False
        Case OpeCGOpeCMACCapIntDevPFMN, OpeCGOpeCMACCapIntDevPFME
            frmCajaIntDeveng.Show 1
        Case OpeCGOpeCMACCancAhorrosMN, OpeCGOpeCMACCancAhorrosME
            frmCajaIntDeveng.Show 1
        Case OpeCGOpeCMACCancPFMN, OpeCGOpeCMACCancPFME
            frmCajaIntDeveng.Show 1
        Case OpeCGOpeCMACMantCtasMN, OpeCGOpeCMACMantCtasME
            frmCajaGenmantCtas.Show 1
        'Caja general Moneda Extranjera
        Case gOpeMECompraAInst, gOpeMEVentaAInst
            frmCajaGenTransf.Show 1
        
        Case gOpeMECompraEfect, gOpeMEVentaEfec
            frmCajaGenCompraMEEfect.Show 1
        
        Case gOpeMEExtCompraAInst, gOpeMEExtVentaAInst, gOpeMEExtCompraEfect, gOpeMEExtVentaEfec
            frmCajaGenExtornos.Show 1
        
        'Cajero Moneda Extranjera
        Case gOpeCajeroMETipoCambio
            frmMantTipoCambio.Show 1
        Case gOpeCajeroMECompra
            frmCajeroCompraVenta.Show 1
        Case gOpeCajeroMEVenta
            frmCajeroCompraVenta.Show 1
        Case gOpeCajeroMEExtCompra, gOpeCajeroMEExtVenta
            frmCajeroExtornos.Show 1
        
        
        'Operaciones Cajero
        Case gOpeHabCajRegEfectMN, gOpeHabCajRegEfectME
            frmCajaGenEfectivo.RegistroEfectivo
        Case gOpeHabCajDevABoveMN, gOpeHabCajDevABoveME
            frmCajeroHab.Show 1
        Case gOpeHabCajTransfEfectCajerosMN, gOpeHabCajTransfEfectCajerosME
            frmCajeroHab.Show 1
        Case gOpeHabCajConfHabBovAgeMN, gOpeHabCajConfHabBovAgeME
            frmCajeroExtornos.Show 1
        Case gOpeHabCajRegSobFaltMN, gOpeHabCajRegSobFaltME
            frmCajeroIngEgre.Show 1
        
        Case gOpeHabCajIngEfectRegulaFaltMN, gOpeHabCajIngEfectRegulaFaltME
            frmCajeroRegFaltSob.Show 1
            
        Case gOpeHabCajExtTransfEfectCajerosMN, gOpeHabCajExtTransfEfectCajerosME
            frmCajeroExtornos.Show 1
        Case gOpeHabCajExtConfHabBovAgeMN, gOpeHabCajExtConfHabBovAgeME
            frmCajeroExtornos.Show 1
        Case gOpeHabCajExtIngEfectRegulaFaltMN, gOpeHabCajExtIngEfectRegulaFaltME
            frmCajeroExtornos.Show 1
            
        Case gOpeHabCajExtDevABoveMN, gOpeHabCajExtDevABoveME
            frmCajeroExtornos.Show 1
        
        '********************* REPORTES DE CAJA GENERAL **************************
        
        Case OpeCGRepRepBancosFlujoMensMN, OpeCGRepRepBancosFlujoMensME
        
        Case OpeCGRepRepBancosFlujoPFMN, OpeCGRepRepBancosFlujoPFME
        Case OpeCGRepRepBancosSaldosCtasMN, OpeCGRepRepBancosSaldosCtasME
        Case OpeCGRepRepCMACSFlujoMensMN, OpeCGRepRepCMACSFlujoMensME
        Case OpeCGRepRepCMACSFlujoPFMN, OpeCGRepRepCMACSFlujoPFME
        Case OpeCGRepRepCMACSSaldosCtasMN, OpeCGRepRepCMACSSaldosCtasME
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
        
        
        '************************* CONTABILIDAD **********************************
        Case gContRegistroAsiento
               frmAsientoRegistro.Inicio "", 0
        Case gContLibroDiario
               frmContabDiario.Show 0, Me
        Case gContLibroMayor
               frmContabMayor.Show 0, Me
        Case gContLibroMayCta
               frmContabMayorDet.Show 0, Me
        Case gContRegCompraGastos
               frmRegCompraGastos.Show 0, Me
        Case gContRegVentas
               frmRegVenta.Show 0, Me
        Case gContRepBaseFormula
            frmRepBaseFormula.Show 0, Me
        Case gContRepCompraVenta
            frmRepResCVenta.Show 0, Me
        Case gContRep_FSD
            frmFondoSeguroDep.Inicio IIf(optMoneda(0).Value, "1", "2")

        'Otros Ajustes
        Case gContAjReclasiCartera
            frmAjusteReCartera.Show 0, Me
        Case gContAjReclasiGaranti
            frmAjusteGarantias.Show 0, Me
        Case gContAjInteresDevenga
            frmAjusteIntDevengado.Inicio True
        Case gContAjInteresSuspens
            frmAjusteIntDevengado.Inicio False
            
        'ANEXOS
        Case gContAnx07
            frmAnexo7RiesgoInteres.Inicio True
    End Select
    Exit Sub
ErrorAceptar:
    MsgBox Err.Description, vbInformation, "Aviso Error"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
tvOpe.SetFocus
End Sub

Private Sub Form_Load()
    Dim sCod As String
    On Error GoTo ERROR
    frmMdiMain.Enabled = False
    
    Logo.AutoPlay = True
    Logo.Open App.Path & "\videos\LogoA.avi"

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
    MsgBox Err.Description, vbExclamation, Me.Caption
End Sub

Sub LoadOpeUsu(psMoneda As String)
Dim clsGen As DGeneral
Dim rsUsu As Recordset
Dim sOperacion As String, sOpeCod As String
Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String
Dim nodOpe As Node

Set clsGen = New DGeneral
Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, sArea, psMoneda)
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

Private Sub optMoneda_Click(Index As Integer)
    Dim sDig As String
    Dim sCod As String
    On Error GoTo ERROR
    If optMoneda(0) Then
        sDig = "2"
    Else
        sDig = "1"
    End If
    AbreConexion
    LoadOpeUsu sDig
    CierraConexion
    tvOpe.SetFocus
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, Me.Caption
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
