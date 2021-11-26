VERSION 5.00
Begin VB.Form frmCredReprogSimulador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulador de Reprogramación"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10890
   Icon            =   "frmCredReprogSimulador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frbotonesSimulador 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   33
      Top             =   6240
      Width           =   10815
      Begin VB.CommandButton cmdReprogramarSimulador 
         Caption         =   "Generar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8160
         TabIndex        =   37
         Top             =   240
         Width           =   1410
      End
      Begin VB.CommandButton cmdSalirSimulador 
         Cancel          =   -1  'True
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
         Height          =   405
         Left            =   9600
         TabIndex        =   36
         Top             =   240
         Width           =   1050
      End
      Begin VB.ComboBox CmbReprogNatEspecialesSimulador 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox cmbCovidReprogSimulador 
         Height          =   315
         ItemData        =   "frmCredReprogSimulador.frx":030A
         Left            =   3120
         List            =   "frmCredReprogSimulador.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Frame frSimulador 
      Height          =   3015
      Left            =   0
      TabIndex        =   32
      Top             =   3240
      Width           =   10815
      Begin SICMACT.FlexEdit FECalendSimulador 
         Height          =   2625
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   4630
         Cols0           =   13
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Fecha-Nro-Monto-Capital-Int. Comp-Int. Mor-Int. Reprog-Int Gracia-Gasto-Saldo-Estado-nCapPag"
         EncabezadosAnchos=   "400-1000-400-1000-1000-1000-1000-1000-1000-1000-1200-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-2-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   65535
         BackColorControl=   65535
         BackColorControl=   65535
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483635
      End
   End
   Begin VB.Frame frMontCuotaCovidSimulador 
      Caption         =   "Monto Cuota"
      Height          =   615
      Left            =   9120
      TabIndex        =   31
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
      Begin SICMACT.EditMoney txtMontoCuotaSimulador 
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame FraDatosSimulador 
      Height          =   3225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.CommandButton CmdBuscarSimulador 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         TabIndex        =   7
         Top             =   260
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.ListBox LstCtasSimulador 
         Height          =   645
         Left            =   8325
         TabIndex        =   6
         Top             =   210
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.TextBox txtTEASimulador 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   5
         Text            =   "0%"
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtTCEADesSimulador 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Text            =   "0%"
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtTCEAReprogSimulador 
         Alignment       =   1  'Right Justify
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5520
         TabIndex        =   3
         Text            =   "0%"
         Top             =   2880
         Width           =   855
      End
      Begin VB.Frame fr_TasaEspecialSimulador 
         Caption         =   "Tasa Especial"
         Height          =   615
         Left            =   7800
         TabIndex        =   1
         Top             =   2520
         Width           =   1215
         Begin VB.Label lbl_TasaEspecialSimulador 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   735
         End
      End
      Begin SICMACT.ActXCodCta ActxCtaSimulador 
         Height          =   480
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   847
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
         CMAC            =   "109"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Interes : "
         Height          =   285
         Left            =   4440
         TabIndex        =   30
         Top             =   2565
         Width           =   1020
      End
      Begin VB.Label LblTasaSimulador 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5520
         TabIndex        =   29
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   28
         Top             =   960
         Width           =   525
      End
      Begin VB.Label lblTitularSimulador 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   27
         Top             =   915
         Width           =   5985
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Analista :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   26
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label LblAnalistaSimulador 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         Top             =   2115
         Width           =   5985
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Préstamo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   24
         Top             =   2565
         Width           =   765
      End
      Begin VB.Label LblPrestamoSimulador 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   23
         Top             =   2520
         Width           =   1140
      End
      Begin VB.Label Saldo 
         AutoSize        =   -1  'True
         Caption         =   "Saldo :"
         Height          =   285
         Left            =   2640
         TabIndex        =   22
         Top             =   2565
         Width           =   495
      End
      Begin VB.Label LblSaldoSimulador 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3240
         TabIndex        =   21
         Top             =   2520
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Capital Reprog.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7440
         TabIndex        =   20
         Top             =   1755
         Width           =   1575
      End
      Begin VB.Label lblSaldoRepSimulador 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   9360
         TabIndex        =   19
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ultima Cuota:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7440
         TabIndex        =   18
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label lblfecUltCuotaSimulador 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9360
         TabIndex        =   17
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo Producto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   165
         TabIndex        =   16
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo Crédito:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   1755
         Width           =   1095
      End
      Begin VB.Label lblTipoCreditoSimulador 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   1725
         Width           =   5985
      End
      Begin VB.Label lblTipoProductoSimulador 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   1320
         Width           =   5985
      End
      Begin VB.Label lblDiasReprogSimulador 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9360
         TabIndex        =   12
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Dias a Reprogramar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7440
         TabIndex        =   11
         Top             =   1365
         Width           =   1485
      End
      Begin VB.Label Label10 
         Caption         =   "TCEA Des.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   2930
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "TCEA Reprg. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         Top             =   2930
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "TEA. :"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   2930
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCredReprogSimulador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'REDISEÑADO X JOEP 20201009

Option Explicit
 
Private nFilaEditar As Integer
Private dFecTemp As Date
Private nMontoApr As Double
Private fnTasaInteres As Double
Private MatCalend As Variant
Private nTipoReprogCred As Integer
Dim ldVigencia As Date
Dim dFecUltCuota As Date
Dim nCuoPag As Integer
Dim nCuoNoPag As Integer
Dim lnCapital As Double
Dim lnIntComp As Double
Dim lnIntGra As Double
Dim loVistoElectronico As frmVistoElectronico
Dim lbVistoVal As Boolean
Dim lsPersCod As String

'Para el manejo Parametrizado de la Reprogramacion
Dim nReprogUltimaCuotaFija As Integer

'**DAOR 20070410**********************************
Private MatGastos As Variant
Private nNumGastos As Integer
Private bCalendGen As Boolean
Private bGastosGen As Boolean
Private nTipoPeriodo As Integer
Private nPlazo As Integer
'*************************************************
Dim objPista As COMManejador.Pista  '' *** PEAC 20090126
'ALPA 20100907 ***********************************
Dim lnPerFechaFijaAct As Integer
Dim lnDiaFijoColocEstado As Integer
'*************************************************

Dim fnTipoComision As Integer 'JUEZ 20130412
Dim fnPersoneria As Integer 'JUEZ 20130412
Dim fbReprogDiasAtraso As Boolean 'JUEZ 20131104
Dim lnValor As Double

Dim fnIntGraciaPend As Double 'JOEP
Dim NewTCEA As Double 'JOEP
Dim nMontoCuota As Double 'angc variable global

'->***** LUCV20180601, según ERS022-2018
Private MatCalendReprogramado As Variant 'Obtiene los registros pendientes a ser pag.
Dim MatCalendSegDes As Variant
Dim rsCalend As ADODB.Recordset
Dim fdFechaCuotaPend As Date
Dim fnTasaSegDes As Double
'<-***** Fin LUCV20180601

Dim dFecVencUltimoPago As Date
Dim dFechaCuotaPend As Date
Dim dDesembolso As Date
Dim nPrimaPerGracia As Double

Dim nMod As Integer
Dim nCuotaReprog As Currency
Dim nCalf As Integer
Dim bRespSalir As Integer

'Add JOEP20210306 garantia covid
Dim gnMenuOpcion As Integer
Dim gvArrayDatos As Variant
'Add JOEP20210306 garantia covid

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

'Public Sub Inicio(ByVal pcCtaCod As String, ByVal pnModalidad As Integer, ByRef pnCuota As Currency, ByRef pCalf As Integer)'Comento JOEP20210306 garantia covid
Public Sub Inicio(ByVal pcCtaCod As String, ByVal pnModalidad As Integer, ByRef pnCuota As Currency, ByRef pCalf As Integer, Optional ByRef pnMatrizCalend As Variant = Nothing, Optional ByVal pnMenuOpcion As Integer = -1, Optional ByRef pvArrayDatos As Variant = Nothing) 'Add JOEP20210306 garantia covid
    bRespSalir = 0
    gnMenuOpcion = pnMenuOpcion 'Add JOEP20210306 garantia covid
    ActxCtaSimulador.NroCuenta = pcCtaCod
    nMod = pnModalidad
    Call ActxCtaSimulador_KeyPress(13)
    Me.Show 1
        
    If bRespSalir = 0 Then
        pnCuota = 0
        pCalf = 0
    'Add JOEP20210306 garantia covid
        Set MatCalend = Nothing
        Set gvArrayDatos = Nothing
    'Add JOEP20210306 garantia covid
    Else
        pnCuota = nCuotaReprog
        pCalf = nCalf
    'Add JOEP20210306 garantia covid
        pnMatrizCalend = MatCalend
        pvArrayDatos = gvArrayDatos
    'Add JOEP20210306 garantia covid
    End If
    
End Sub

Private Sub HabilitaControlesReprog(ByVal pbHabilita As Boolean)
    FECalendSimulador.lbEditarFlex = pbHabilita
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
    Dim oCredito As COMDCredito.DCOMCredito
    Dim rsCred As ADODB.Recordset
    Dim rsCal As ADODB.Recordset
    Dim rsReprogApr As ADODB.Recordset
    Dim rsDatosReprog As ADODB.Recordset
    Dim bAutorizado As Boolean
    Dim nPrdEstado As Integer
    Dim lnSaldoNew As Double
    Dim lnSegDes As Double 'JOEP
    
    Dim nMontIntGraciaTotal As Double 'JOEP
    Dim nMontIntGracCuotaPag As Double 'JOEP

    CargaDatos = False
    
    On Error GoTo ErrorCargaDatos
    LimpiaFlex FECalendSimulador
    MatCalend = ""
    
    Call CargaCombo 'JOEP20200706 Mantener Cuota
    
    'Obtiene los datos del crédito a reprogramar
    Set oCredito = New COMDCredito.DCOMCredito
    'Call oCredito.SimuladoReprogramacionCargarDatos(psCtaCod, rsCred, rsCal)'Comento JOEP20210306 garantia covid
    Call oCredito.SimuladoReprogramacionCargarDatos(psCtaCod, rsCred, rsCal, gnMenuOpcion) 'Add JOEP20210306 garantia covid
    Set oCredito = Nothing
            
    'Asignación de valores según tablas de memoria
    nPrdEstado = rsCred!nPrdEstado
    fnTasaInteres = CDbl(Format(rsCred!nTasaInteres, "#,##0.0000"))
    lblTitularSimulador.Caption = PstaNombre(rsCred!cTitular)
    LblAnalistaSimulador.Caption = PstaNombre(rsCred!cAnalista)
    LblSaldoSimulador.Caption = Format(rsCred!nSaldo, "#,##0.00")
    LblPrestamoSimulador.Caption = Format(rsCred!nMontoCol, "#,##0.00")
    LblTasaSimulador.Caption = Format(fnTasaInteres, "#0.0000")
    ldVigencia = Format(rsCred!dVigencia, "dd/mm/yyyy")
    fnTasaSegDes = Format(rsCred!nTasaSegDesg, "#0.0000")
    lsPersCod = rsCred!cPersCod
    'ALPA 20100606***************
    lblTipoCreditoSimulador.Caption = rsCred!cTpoCredDes
    lblTipoProductoSimulador.Caption = rsCred!cTpoProdDes
    '****************************
    fdFechaCuotaPend = Format(rsCred!dVenc_Cuota, "dd/mm/yyyy") 'LUCV20180601, Según ERS022-2018
    lblDiasReprogSimulador.Caption = rsCred!nDiasReprogramar 'DateDiff("d", CDate(rsCred!dFecCuotaVenc), CDate(rsCred!dFecNuevaCuotaVenc))
    'txtDiasporReprog.Text = DateDiff("d", CDate(rsDatosReprog!dFecCuotaVenc), CDate(rsDatosReprog!dFecNuevaCuotaVenc))
    
    dFecVencUltimoPago = rsCred!dFecVencUltimoPago
    dFechaCuotaPend = rsCred!dFechaCuotaPend
    dDesembolso = rsCred!dFecVencReprog
    nPrimaPerGracia = rsCred!nPrimaPerGracia
                
    'JOEP20200425 cuota igual covid (en Reprogramados)
    txtTEASimulador = rsCred!TEA
    txtTCEADesSimulador = rsCred!TCEA
    'JOEP20200425 cuota igual covid (en Reprogramados)
    
    rsCred.Close
    Set rsCred = Nothing

    lnSaldoNew = 0
    'Add by Gitu
    nCuoPag = 0
    nCuoNoPag = 0
    lnCapital = 0
    lnIntComp = 0
    lnIntGra = 0
    lnSegDes = 0 'JOEP
    nMontIntGraciaTotal = 0 'JOEP
    nMontIntGracCuotaPag = 0 'JOEP
    fnIntGraciaPend = 0 'JOEP
    'End Gitu
    
    nMontoApr = rsCal!nSaldoPactado
    'Recorrido del calendario Actual del crédito
    Do While Not rsCal.EOF
        'Add by Gitu
        nCuoPag = nCuoPag + 1
        nCuoNoPag = nCuoNoPag + 1
        If rsCal!nColocCalendEstado = gColocCalendEstadoPagado Then
            lnCapital = rsCal!nCapital
            lnIntComp = rsCal!nIntComp
            lnIntGra = rsCal!nIntGracia
            
        Else
            lnCapital = rsCal!nCapital - rsCal!nCapitalPag
            lnIntComp = rsCal!nIntComp - rsCal!nIntCompPag
            lnIntGra = rsCal!nIntGracia - rsCal!nIntGraciaPag
            lnSegDes = rsCal!nGasto - rsCal!nGastoPag 'JOEP
        End If
        
        FECalendSimulador.AdicionaFila
        FECalendSimulador.TextMatrix(rsCal.Bookmark, 1) = Format(rsCal!dVenc, "dd/mm/yyyy")
        FECalendSimulador.TextMatrix(rsCal.Bookmark, 2) = Trim(str(rsCal!nCuota))
        FECalendSimulador.TextMatrix(rsCal.Bookmark, 3) = Format(IIf(IsNull(lnCapital), 0, lnCapital) + _
                                        IIf(IsNull(lnIntComp), 0, lnIntComp) + _
                                        IIf(IsNull(lnIntGra), 0, lnIntGra) + _
                                        IIf(IsNull(rsCal!nIntMor), 0, rsCal!nIntMor) + _
                                        IIf(IsNull(rsCal!nIntReprog), 0, rsCal!nIntReprog) + _
                                        IIf(IsNull(rsCal!nGasto), 0, rsCal!nGasto), "#0.00")
        FECalendSimulador.TextMatrix(rsCal.Bookmark, 4) = Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")
        FECalendSimulador.TextMatrix(rsCal.Bookmark, 5) = Format(IIf(IsNull(lnIntComp), 0, lnIntComp), "#0.00")
        FECalendSimulador.TextMatrix(rsCal.Bookmark, 6) = Format(IIf(IsNull(rsCal!nIntMor), 0, rsCal!nIntMor), "#0.00")
        FECalendSimulador.TextMatrix(rsCal.Bookmark, 7) = Format(IIf(IsNull(rsCal!nIntReprog), 0, rsCal!nIntReprog), "#0.00")
        FECalendSimulador.TextMatrix(rsCal.Bookmark, 8) = Format(IIf(IsNull(lnIntGra), 0, lnIntGra), "#0.00")
        FECalendSimulador.TextMatrix(rsCal.Bookmark, 9) = Format(IIf(IsNull(rsCal!nGasto), 0, rsCal!nGasto), "#0.00")
        nMontoApr = nMontoApr - IIf(IsNull(rsCal!nCapital), 0, rsCal!nCapital)
        nMontoApr = CDbl(Format(nMontoApr, "#0.00"))
        FECalendSimulador.TextMatrix(rsCal.Bookmark, 10) = Format(nMontoApr, "#0.00")
        FECalendSimulador.TextMatrix(rsCal.Bookmark, 11) = Trim(str(rsCal!nColocCalendEstado))
        FECalendSimulador.TextMatrix(rsCal.Bookmark, 12) = Format(IIf(IsNull(rsCal!nCapitalPag), 0, rsCal!nCapitalPag), "#0.00")
        
        lnSaldoNew = lnSaldoNew + IIf(IsNull(rsCal!nCapital), 0, rsCal!nCapital) - IIf(IsNull(rsCal!nCapitalPag), 0, rsCal!nCapitalPag)
        'End Gitu
        
        If rsCal!nColocCalendEstado = gColocCalendEstadoPagado Then
            FECalendSimulador.row = rsCal.Bookmark
            Call FECalendSimulador.ForeColorRow(vbRed)
            nCuoNoPag = nCuoNoPag - 1
        End If
        If rsCal.RecordCount = rsCal.Bookmark Then
            lblfecUltCuotaSimulador = Format(rsCal!dVenc, "dd/mm/yyyy")
        End If
        
        'JOEP
        nMontIntGraciaTotal = nMontIntGraciaTotal + rsCal!nIntGracia
        'If rsCal!nColocCalendEstado = 1 Then'JOEP20200321 Comento Mejora Reprogramacion
        nMontIntGracCuotaPag = nMontIntGracCuotaPag + rsCal!nIntGraciaPag
        'End If'JOEP20200321 Comento Mejora Reprogramacion
        'JOEP
        
        rsCal.MoveNext
    Loop
    
    Set rsCalend = New ADODB.Recordset 'LUCV20180601, Agregó, Según ERS022-2018
    Set rsCalend = rsCal.Clone 'LUCV20180601, Agregó, Según ERS022-2018
    
    rsCal.Close
    Set rsCal = Nothing
    lblSaldoRepSimulador = Format(lnSaldoNew, "#,##0.00")
    
    'If bAutorizado Then
        CmbReprogNatEspecialesSimulador.Visible = True 'JOEP20200428 Covid Cuota Iguales
    'End If
    
    fnIntGraciaPend = nMontIntGraciaTotal - nMontIntGracCuotaPag 'JOEP
    
    CargaDatos = True
    
    Exit Function
ErrorCargaDatos:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Private Sub ActxCtaSimulador_KeyPress(KeyAscii As Integer)
    Dim oDCredito As New COMDCredito.DCOMCredito 'JUEZ 20131104
    Dim bCredito As Boolean
    Dim bTieneDiasAtraso As Boolean  'JUEZ 20131104
    If KeyAscii = 13 Then
       
        'END JUEZ *********************************************
        If Len(ActxCtaSimulador.NroCuenta) = 18 Then 'JUEZ 20130412
            If CargaDatos(ActxCtaSimulador.NroCuenta) Then
                'JUEZ 20130412 ******************************************************************
                'Dim odCredito As COMDCredito.DCOMCredito 'JUEZ 20131104
                Set oDCredito = New COMDCredito.DCOMCredito
                If oDCredito.ExisteComisionVigente(ActxCtaSimulador.NroCuenta, gComisionReprogCredito) Then
                    fnTipoComision = 1 'Pagado
                Else
                    Dim R As ADODB.Recordset
                    Dim lsPrdConceptoCod As Integer
                    Set oDCredito = New COMDCredito.DCOMCredito
                    Set R = oDCredito.RecuperaDatosComision(ActxCtaSimulador.NroCuenta, 1)
                    fnPersoneria = R!nPersoneria
                    lsPrdConceptoCod = IIf(R!nPersoneria = 1, gColocConceptoCodGastoComisionReprogNat, gColocConceptoCodGastoComisionReprogJur)
                    
                    Set R = oDCredito.RecuperaProductoConcepto(lsPrdConceptoCod)
                    'JUEZ 20151229 ************************************************
                    Dim lnTCVenta As Double
                    Dim oDGeneral As COMDConstSistema.NCOMTipoCambio
                    Set oDGeneral = New COMDConstSistema.NCOMTipoCambio
                        lnTCVenta = oDGeneral.EmiteTipoCambio(gdFecSis, TCVenta)
                    Set oDGeneral = Nothing
                    lnValor = CDbl(R!nValor) / IIf(Mid(ActxCtaSimulador.NroCuenta, 9, 1) = "1", 1, lnTCVenta)
                    'END JUEZ *****************************************************
                    
                    MsgBox "Se aplicará el pago de " & IIf(Mid(ActxCtaSimulador.NroCuenta, 9, 1) = "1", "S/ ", "$ ") & Format(lnValor, "#,##0.00") & " por concepto de Comisión de Reprogramación de Créditos, en la siguiente cuota.", vbInformation, "Aviso"
                    fnTipoComision = 2 'Cobrar en proxima cuota
                End If
                'END JUEZ ***********************************************************************
                FraDatosSimulador.Enabled = False
                
                'cmdEditar.Enabled = True
                cmdReprogramarSimulador.Enabled = True
                
                'CmdGastos.Enabled = True 'DAOR 20070410
                HabilitarReprogramar True
                
                'joep20201002 Tasa especial y reduccion de monto
                Dim obOCM As COMDCredito.DCOMCredito
                Dim rsRepgOCM As ADODB.Recordset
                Set obOCM = New COMDCredito.DCOMCredito
                'Set rsRepgOCM = obOCM.SimuladoReprogramacion(ActxCtaSimulador.NroCuenta)'Comento JOEP20210306 garantia covid
                Set rsRepgOCM = obOCM.SimuladoReprogramacion(ActxCtaSimulador.NroCuenta, gnMenuOpcion) 'Add JOEP20210306 garantia covid
                If Not (rsRepgOCM.BOF And rsRepgOCM.EOF) Then
                    Call OCMControl(nMod)
                End If
                Set obOCM = Nothing
                RSClose rsRepgOCM
                'joep20201002 Tasa especial y reduccion de monto
            Else
                cmdCancelarSimulador_Click
            End If
        End If
    End If
End Sub

Private Sub cmdCancelarSimulador_Click()
    Call cmdNuevo_Click
    Call VisibleBotones(0, False)
End Sub

Private Sub cmdNuevo_Click()
    nTipoReprogCred = -1
    FraDatosSimulador.Enabled = True
    LimpiaFlex FECalendSimulador
    nFilaEditar = -1
    ActxCtaSimulador.NroCuenta = ""
    ActxCtaSimulador.CMAC = gsCodCMAC
    ActxCtaSimulador.Age = gsCodAge
    cmdSalirSimulador.Enabled = True
    HabilitaControlesReprog False
    LblAnalistaSimulador.Caption = ""
    LblPrestamoSimulador.Caption = "0.00"
    LblSaldoSimulador.Caption = "0.00"
    lblTitularSimulador.Caption = ""
    LblTasaSimulador.Caption = "0.0000"
    MatCalend = ""
    lblSaldoRepSimulador = "0.00"
    LstCtasSimulador.Clear
    lblTipoCreditoSimulador.Caption = ""
    lblTipoProductoSimulador.Caption = ""
    
    CmbReprogNatEspecialesSimulador.ListIndex = -1
    cmbCovidReprogSimulador.Visible = False
    CmbReprogNatEspecialesSimulador.Visible = False
    fbReprogDiasAtraso = False
    lblfecUltCuotaSimulador.Caption = ""
    Me.lblDiasReprogSimulador.Caption = ""
    cmdReprogramarSimulador.Enabled = False
    HabilitarReprogramar True
    lnValor = 0
    txtMontoCuotaSimulador.Text = "0.00"
    txtTCEAReprogSimulador.Text = "0%"
    txtTEASimulador.Text = "0%"
    txtTCEADesSimulador.Text = "0%"
    nMontoCuota = 0
End Sub

Private Sub HabilitarReprogramar(ByVal pbHabilita As Boolean)
    cmdReprogramarSimulador.Enabled = pbHabilita
    CmbReprogNatEspecialesSimulador.Enabled = pbHabilita
    cmbCovidReprogSimulador.Enabled = pbHabilita
    txtMontoCuotaSimulador.Enabled = pbHabilita
End Sub

Private Sub cmdReprogramarSimulador_Click()

If ValidaDatos(0) = True Then
    Exit Sub
End If

If CmbReprogNatEspecialesSimulador.Visible = True And (CmbReprogNatEspecialesSimulador.Text) = "" Then
    MsgBox "Seleccione Opción de reprogramación", vbInformation, "Aviso"
    CmbReprogNatEspecialesSimulador.SetFocus
    Exit Sub
End If
If cmbCovidReprogSimulador.Visible = True And (cmbCovidReprogSimulador.Text) = "" Then
    MsgBox "Seleccione Opción de reprogramación", vbInformation, "Aviso"
    cmbCovidReprogSimulador.SetFocus
    Exit Sub
End If

If CmbReprogNatEspecialesSimulador.Visible = True And Right(CmbReprogNatEspecialesSimulador.Text, 1) = 1 Then
    'Normal=1,Mnatener cuota=2,Tasa especial(Solo a la primera cuota Reprogramda)=3,Reduccion de monto=4,Reduccion de Tasa (aplica a toda las cuotas)=5
    If cmbCovidReprogSimulador.Visible = True And _
    (Right(cmbCovidReprogSimulador.Text, 1) = 1 Or Right(cmbCovidReprogSimulador.Text, 1) = 2 Or _
     Right(cmbCovidReprogSimulador.Text, 1) = 3 Or Right(cmbCovidReprogSimulador.Text, 1) = 4 Or _
     Right(cmbCovidReprogSimulador.Text, 1) = 5) Then
        cmdReprog_Click
    'ElseIf cmbCovidReprogSimulador.Visible = True And (Right(cmbCovidReprogSimulador.Text, 1) = 2 Or Right(cmbCovidReprogSimulador.Text, 1) = 4) Then
        'cmdReprog_Click
    End If
Else
    cmdReprog_Click
End If

EnableBotones 0, False

'Add JOEP20200428 covid cuotas iguales

HabilitarReprogramar False

End Sub

Private Sub cmdSalirSimulador_Click()
Dim i As Integer
nCuotaReprog = 0
If cmdReprogramarSimulador.Enabled = False Then
'Obtenemos la cuota de la reprogramacion
    For i = 1 To FECalendSimulador.rows - 1
        If FECalendSimulador.TextMatrix(i, 11) = 0 Then
            nCuotaReprog = FECalendSimulador.TextMatrix(i, 3)
            Exit For
        End If
    Next i
    
'Obtenemos el valor si calificac la reprogramacion
    If ValidaDatos(1) = True Or ValidaDatos(0) = True Then
        nCalf = 2
    Else
        nCalf = 1
    End If
    bRespSalir = 1
    Unload Me
Else
    If MsgBox("Esta seguro de salir, sin generar la simulacion de reprogramación, Para continuar presione [SI] ", vbYesNo, "Aviso") = vbYes Then
        bRespSalir = 0
        Unload Me
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then 'F12
        Dim bRetSinTarjeta As Boolean
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColComercEmp, bRetSinTarjeta)
        If sCuenta <> "" Then
            ActxCtaSimulador.NroCuenta = sCuenta
            ActxCtaSimulador.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    DisableCloseButton Me
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    CentraForm Me
    ActxCtaSimulador.CMAC = gsCodCMAC
    ActxCtaSimulador.Age = gsCodAge
    nFilaEditar = -1
    ValidarFechaActual
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredReprogramarCredito
    fnPersoneria = 0
    fbReprogDiasAtraso = False
    HabilitarReprogramar True
    fr_TasaEspecialSimulador.Visible = False
    bRespSalir = 0
End Sub

Private Sub ValidarFechaActual()
Dim lsFechaValidador As String

    lsFechaValidador = validarFechaSistema
    If lsFechaValidador <> "" Then
        If gdFecSis <> CDate(lsFechaValidador) Then
            MsgBox "La Fecha de tu sesión en el Negocio no coincide con la fecha del Sistema", vbCritical, "Aviso"
            Unload Me
            End
        End If
    End If
End Sub


'**DAOR 20070410
Sub EstablecerGastos(pMatGastos As Variant, pbGastosGen As Boolean, pnNumGastos As Integer, pnTipoPeriodo As Integer, pnPlazo As Integer)
    MatGastos = pMatGastos
    bGastosGen = pbGastosGen
    nNumGastos = pnNumGastos
    nTipoPeriodo = pnTipoPeriodo
    nPlazo = pnPlazo
End Sub
Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
    'permite que solo sean ingresados los numeros, el ENTER y el RETROCESO
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        SoloNumeros = 0
    Else
        SoloNumeros = KeyAscii
    End If
    ' teclas especiales permitidas
    If KeyAscii = 8 Then SoloNumeros = KeyAscii ' borrado atras
    If KeyAscii = 13 Then SoloNumeros = KeyAscii 'Enter
End Function

'->***** LUCV20180601, Modificó evento, según ERS022-2018
Private Sub cmdReprog_Click()
    Dim oDCOMCredito As COMDCredito.DCOMCredito 'LUCV20180601
    Dim oNCOMCredito As COMNCredito.NCOMCredito
    Dim oDCOMConecta As COMConecta.DCOMConecta
    
    Dim lnSaldoNew As Double
    Dim dFechaReprog As Date
    Dim nSaldoCapital As Double
    Dim i As Integer
    Dim j As Integer
    Dim nGastos As Currency
    
    'JOEP INICIO
    Dim nUltPago As Date
    Dim nUltPagoNoComp As Date
    Dim fnCantAfiliadosSegDes As Integer
    Dim nTasaSegDes As Double
    Dim nTotalCapital As Double
    Dim nAjuste As Double
    Dim nValCuoataAju As Double
    Dim nCuotaAjustada As Double
    Dim rsTipPeriodo As ADODB.Recordset
    Dim rsTasaEspecialCovid As ADODB.Recordset 'Joep20200910 Tasa Especial covid
    Dim nTasaEspCovid As Double 'Joep20200910 Tasa Especial covid
    Dim nOpCovid As Integer 'Joep20200910 Tasa Especial covid
    'JOEP FIN
    
    '->***** LUCV20180601
    Dim oNCOMCalendario As COMNCredito.NCOMCalendario
    Dim rsDatosAprob As ADODB.Recordset
    Dim nDiasPorReprogramacion As Integer
    
    Dim nGastoSegDesg As Double
    Dim nGastoIncendio As Double
    Dim nGastoIncendioGracia As Double
    
    'RIRO 20200825 Mejora en liquidación
    Dim nIntCompCalculado As Double
    Dim nDiasCalculo As Integer
    Dim nIntCompDiferenciaCapitalizado As Double
    Dim nIntGraciaGenerado As Double
    Dim nIntGraciaCapitalizado As Double
    Dim nIntGraciaAsignado As Double
    'RIRO 20200825 Mejora en liquidación
    
    'Dim nIntGraciaPendCap As Double
    'Dim nIntGraciaPendTotal As Double
    'Para Gastos
    Dim nMontoPoliza As Double
    Dim nTasaSegInc As Double
    Dim oNGasto As COMNCredito.NCOMGasto
    Set oNGasto = New COMNCredito.NCOMGasto
    '<-***** Fin LUCV20180601
    
    Dim rsLiquidacionConceptos As ADODB.Recordset 'Add JOEP20200414 Covid
    
    'RIRO 20210212 ********************
    Dim nPolizaMen As Double        ' Póliza mensual sin el prorrateo, concepto 1231
    Dim nPolizaCuotReprog As Double ' Póliza de la primera cuota, concepto 1231
    Dim nPolizaProrrateo As Double  ' Póliza prorrateada en cada cuota, concepto 1279
    Dim dFechaCorte As Date         ' Fecha de corte hasta donde se aplicarán los intereses
    'END RIRO *************************
    
    'Seteo de variables
    HabilitaControlesReprog True
    LimpiaFlex FECalendSimulador
    MatCalend = ""
    lnSaldoNew = 0: i = 0: j = 0: lnCapital = 0: lnIntComp = 0: lnIntGra = 0: nMontoPoliza = 0: nDiasPorReprogramacion = 0
    nMontoApr = rsCalend!nSaldoPactado
    nSaldoCapital = val(Replace(LblSaldoSimulador.Caption, ",", ""))
    nDiasPorReprogramacion = CInt(lblDiasReprogSimulador)
    dFechaReprog = fdFechaCuotaPend + nDiasPorReprogramacion
     nTasaEspCovid = 0 'Joep20200910 Tasa Especial covid
    nOpCovid = 0 'Joep20200910 Tasa Especial covid
    
    'Joep20200910 Tasa Especial covid
    If cmbCovidReprogSimulador.Visible = True Then
        nOpCovid = Right(cmbCovidReprogSimulador.Text, 1)
    End If
    'Joep20200910 Tasa Especial covid
    
    'Estados datos del crédito con estado Aprobado
    Set oDCOMCredito = New COMDCredito.DCOMCredito
    Set rsDatosAprob = oDCOMCredito.RecuperaColocacEstado(ActxCtaSimulador.NroCuenta, gColocEstAprob)
    Set oDCOMCredito = Nothing
    
     'Joep20200910 Tasa Especial covid
    Set oDCOMCredito = New COMDCredito.DCOMCredito
    'Set rsTasaEspecialCovid = oDCOMCredito.ReprogramacionObtTasaEspecial(ActxCtaSimulador.NroCuenta, nSaldoCapital)'Comento JOEP20210306 garantia covid
    Set rsTasaEspecialCovid = oDCOMCredito.ReprogramacionObtTasaEspecial(ActxCtaSimulador.NroCuenta, nSaldoCapital, gnMenuOpcion, nOpCovid) 'Add JOEP20210306 garantia covid
        If Not (rsTasaEspecialCovid.BOF And rsTasaEspecialCovid.EOF) Then
            nTasaEspCovid = rsTasaEspecialCovid!nTasaInteres
        'Add JOEP20210306 garantia covid
            If gnMenuOpcion = 2 And nOpCovid = 5 Then
                fnTasaInteres = nTasaEspCovid
            End If
        'Add JOEP20210306 garantia covid
        End If
    Set oDCOMCredito = Nothing
    'Joep20200910 Tasa Especial covid
    
    'Seguro Desgravamen
    Set oNCOMCalendario = New COMNCredito.NCOMCalendario
    If fnTasaSegDes <> 0 Then
        nTasaSegDes = fnTasaSegDes 'Tasa SegDes. Desembolso
    Else
        Set oNCOMCredito = New COMNCredito.NCOMCredito 'JOEP20200317 Mejora
        nTasaSegDes = oNCOMCredito.ObtenerTasaSeguroDesg(ActxCtaSimulador.NroCuenta, gdFecSis, fnCantAfiliadosSegDes) 'Tasa SegDes. Actual
    End If
    
    'Seguro Incendio del desembolso
    'Comento JOEP20200414 Covid
'    nMontoPoliza = oNGasto.RecuperaMontoPoliza(ActxCta.NroCuenta, _
'                                                nCuoNoPag, _
'                                                gColocConceptoCodGastoPolizaIncendioHipoteca, _
'                                                nTasaSegInc)
    'Comento JOEP20200414 Covid
    
'RIRO 20210215 COMENTADO ***************************************************************************
'    'Add JOEP20200414 Covid
'    nMontoPoliza = oNGasto.RecuperaMontoPoliza(ActxCtaSimulador.NroCuenta, _
'                                                nCuoNoPag, _
'                                                gColocConceptoCodGastoPolizaIncendioHipoteca, _
'                                                nTasaSegInc, , , , 1)
'    'Add JOEP20200414 Covid
'RIRO 20210215 COMENTADO ***************************************************************************

    Dim nSegIncAnt As Double
'RIRO 20210211 Se realiza de esta manera por mergencia y evitar el pase *****
    Dim oConPoliza As COMConecta.DCOMConecta
    Dim rsPoliza As ADODB.Recordset
    Dim ssql As String
    Set rsPoliza = New ADODB.Recordset
    
    ssql = "exec stp_sel_ObtieneSeguroIncendio '" & ActxCtaSimulador.NroCuenta & "'," & nCuoNoPag & "," & nCuoNoPag & ", " & _
            nOpCovid & ", " & CCur(txtMontoCuotaSimulador) & "," & nDiasPorReprogramacion
            
    Set oConPoliza = New COMConecta.DCOMConecta
    oConPoliza.AbreConexion
    Set rsPoliza = oConPoliza.CargaRecordSet(ssql)
    oConPoliza.CierraConexion
    Set oConPoliza = Nothing
    
    nPolizaMen = 0
    nPolizaCuotReprog = 0
    nPolizaProrrateo = 0
    nMontoPoliza = 0
    dFechaCorte = "01/01/1900"
    
    If Not rsPoliza Is Nothing Then
        If rsPoliza.State = 1 Then
            If Not rsPoliza.EOF And Not rsPoliza.BOF Then
                If rsPoliza.RecordCount > 0 Then
                    nPolizaMen = Round(rsPoliza!nPolizaMen, 2)
                    nPolizaCuotReprog = Round(rsPoliza!nPolizaCuotReprog, 2)
                    nPolizaProrrateo = Round(rsPoliza!nPolizaProrrateo, 2)
                    nMontoPoliza = Round(rsPoliza!nPolizaMen + rsPoliza!nPolizaProrrateo, 2)
                    dFechaCorte = rsPoliza!dVencCuotReprog
                    nSegIncAnt = Round(rsPoliza!nPolizaPend, 2)
                End If
            End If
        End If
    End If
    If dFechaCorte = "01/01/1900" Then
        MsgBox "Se han presentado inconvenientes al validar la póliza contra incendios, favor de comunicarse con T.I.", vbInformation, "Validación Póliza"
        Exit Sub
    End If
'END RIRO *******************************************************************
    
    'Liquidación de la deuda:
    Dim MatCalendIni As Variant          'Matriz del Calendario Pend. a pagar
    Dim vArrayDatos As Variant           'Array de parametros de la liquidación de la deuda
    Dim nCapital As Double               'Saldo Capital
    Dim nInteresCompAFecha As Double     'Interés Compensatorio (Hasta la Fecha Reprogramación)
    Dim nInteresGraciaAFecha As Double   'Interés Gracia pendiente
    Dim nInteresCompVencAFecha As Double 'Interés Compensatorio Vencido
    Dim nInteresMoratorio As Double      'Interés Moratorio (de todas las cuotas)
    Dim nSegDesgAnt As Double
    'Dim nSegIncAnt As Double
    Dim nSegIncGraciaAnt As Double
    
    'Calendario de pagos pendiente
    Set oNCOMCredito = New COMNCredito.NCOMCredito
    MatCalendIni = oNCOMCredito.RecuperaMatrizCalendarioPendiente(ActxCtaSimulador.NroCuenta)
    
    'Capital
    nCapital = oNCOMCredito.MatrizCapitalAFecha(ActxCtaSimulador.NroCuenta, MatCalendIni)
    
    'Liq. Interes Compensatorio.
    'nInteresCompAFecha = oNCOMCredito.MatrizInteresCompAFecha(ActxCtaSimulador.NroCuenta, MatCalendIni, gdFecSis) 'Comento JOEP20200414 Covid 'Cumple cuando la cuota no tiene días de atraso'DesComento JOEP_RIRO_20200914
    nInteresCompAFecha = oNCOMCredito.MatrizInteresCompAFecha(ActxCtaSimulador.NroCuenta, MatCalendIni, dFechaCorte) 'Comento JOEP20200414 Covid 'Cumple cuando la cuota no tiene días de atraso'DesComento JOEP_RIRO_20200914
    
    'Liq. Interés de Gracia.
    'nInteresGraciaAFecha = fnIntGraciaPend
    nInteresGraciaAFecha = oNCOMCredito.MatrizInteresGraciaFecha(ActxCtaSimulador.NroCuenta, MatCalendIni, dFechaCorte) 'ADD RIRO 20210214
    'nInteresGraciaAFecha = oNCOMCredito.MatrizInteresGraciaFecha(ActxCta.NroCuenta, MatCalendIni, gdFecSis)
    
    'Liq. Interés Moratorio
    If gnMenuOpcion = 2 And nOpCovid = 5 Then
        nInteresMoratorio = 0
    Else
        nInteresMoratorio = oNCOMCredito.MatrizIntMoratorioCalendario(MatCalendIni)
    End If
    
    'Liq. Interés Compensatorio Vencido. (Este proceso esta en proceso de implementación)
    nInteresCompVencAFecha = oNCOMCredito.MatrizInteresCompVencidoFecha(ActxCtaSimulador.NroCuenta, MatCalendIni)
    
    'Liq. de Gastos
    'nSegDesgAnt = oNCOMCredito.TotalGastosAFecha(ActxCtaSimulador.NroCuenta, gdFecSis, gColocConceptoCodGastoSeguro7) 'Comento JOEP20200414 Covid 'Descomento JOEP_RIRO_20200914
    nSegDesgAnt = oNCOMCredito.TotalGastosAFecha(ActxCtaSimulador.NroCuenta, dFechaCorte, gColocConceptoCodGastoSeguro7, gnMenuOpcion, nOpCovid) 'Add JOEP20210306 garantia covid
    
    'Comento JOEP20200414 Covid
    'nSegIncAnt = oNCOMCredito.TotalGastosAFecha(ActxCtaSimulador.NroCuenta, Format(gdFecSis, "mm/dd/yyyy"), gColocConceptoCodGastoPolizaIncendioHipoteca) 'DesComento JOEP_RIRO_20200914
    'Comento JOEP20200414 Covid
    'nSegIncAnt = 0 'Add JOEP20200414 Covid 'Comento JOEP_RIRO_20200914
    'nSegIncGraciaAnt = oNCOMCredito.TotalGastosAFecha(ActxCtaSimulador.NroCuenta, Format(gdFecSis, "mm/dd/yyyy"), gColocConceptoCodGastoPolizaIncendioHipotecaGracia)
    
    'Descomentar para pruebas.
'    MsgBox "Capital=" & nCapital & Chr(13) & _
'    " InteresComp=" & nInteresCompAFecha & Chr(13) & _
'    " InteresGracia=" & nInteresGraciaAFecha & Chr(13) & _
'    " Moratorio=" & nInteresMoratorio & Chr(13) & _
'    " InteresCompVencAFecha=" & nInteresCompVencAFecha & Chr(13) & _
'    " SegDes=" & nSegDesgAnt & Chr(13) & _
'    "SegInc=" & nSegIncAnt & _
'    " SegIncGracia=" & nPolizaProrrateo & " TasaSegDes=" & nTasaSegDes, vbInformation, "Liquidacion"
'Descomentar para pruebas.

    ReDim gvArrayDatos(8)
    gvArrayDatos(0) = nCapital
    gvArrayDatos(1) = nInteresCompAFecha
    gvArrayDatos(2) = nInteresGraciaAFecha
    gvArrayDatos(3) = nInteresMoratorio
    gvArrayDatos(4) = nInteresCompVencAFecha
    gvArrayDatos(5) = nSegDesgAnt
    gvArrayDatos(6) = nSegIncAnt
    gvArrayDatos(7) = nPolizaProrrateo
    
    'Agrupación de importes liquidados.
    ReDim vArrayDatos(15)
    vArrayDatos(0) = nInteresCompAFecha
    vArrayDatos(1) = nInteresGraciaAFecha
    vArrayDatos(2) = nInteresMoratorio
    vArrayDatos(3) = nInteresCompVencAFecha
    vArrayDatos(4) = nSegDesgAnt
    vArrayDatos(5) = nPolizaMen
    vArrayDatos(6) = nPolizaCuotReprog
    vArrayDatos(7) = nTasaEspCovid 'Joep20200910 Tasa Especial covid
    vArrayDatos(8) = nOpCovid 'Joep20200910 Tasa Especial covid
    vArrayDatos(9) = nPolizaProrrateo
    vArrayDatos(10) = dFecVencUltimoPago
    vArrayDatos(11) = dFechaCuotaPend
    vArrayDatos(12) = dDesembolso
    vArrayDatos(13) = nPrimaPerGracia
    vArrayDatos(14) = IIf(CCur(txtMontoCuotaSimulador) = 0, 0, CCur(txtMontoCuotaSimulador)) - nMontoPoliza 'nMontoPoliza 'IIf(CCur(txtMontoCuotaSimulador) = 0, 0, (CCur(txtMontoCuotaSimulador) - nPolizaProrrateo))
    vArrayDatos(15) = 0 'Cuotas incrementadas
    'Fin Liquidación
        
    'Generacion del calendario de pagos de las cuotas no pagadas
    ReDim MatCalendReprogramado(nCuoNoPag)
                    
    MatCalendReprogramado = oNCOMCalendario.SimuladorReprogGeneraCalendario(CDbl(LblSaldoSimulador), _
                                                            fnTasaInteres, _
                                                            nCuoNoPag, _
                                                            IIf(IsNull(rsDatosAprob!nPlazo), 0, rsDatosAprob!nPlazo), _
                                                            gdFecSis, _
                                                            Fija, _
                                                            IIf(rsDatosAprob!nPeriodoFechaFija > 0, 2, 1), _
                                                            PrimeraCuota, _
                                                            nDiasPorReprogramacion, _
                                                            Day(dFechaReprog), _
                                                            IIf(IsNull(rsDatosAprob!nProxMes), 0, rsDatosAprob!nProxMes) _
                                                            , , , , , , , , , , , , , , _
                                                            ActxCtaSimulador.NroCuenta, , , _
                                                            nInteresGraciaAFecha, _
                                                            , , , , nTasaSegDes, _
                                                            MatCalendSegDes, , _
                                                            nMontoPoliza, _
                                                            nTasaSegInc, _
                                                            vArrayDatos)
      
If (Right(cmbCovidReprogSimulador.Text, 1) = 2 Or Right(cmbCovidReprogSimulador.Text, 1) = 4) Then
    Dim nCuotasPag As Integer
    Dim X As Integer
    nCuotasPag = nCuoPag - nCuoNoPag
    nCuotasPag = nCuotasPag + vArrayDatos(14)
        
    'Generacion del calendario Reprogramado
    'ReDim MatCalend(nCuoPag, 17) 'LUCV20180601, Modificó 11 por 17
    'ReDim MatCalend(nCuotasPag, 23) 'LUCV20180601, Modificó 11 por 17/ RIRO Se modificó a 23
    ReDim MatCalend(nCuotasPag, 27) 'Add JOEP20210306 garantia covid
    'JOEP-Lucv coivd
  
    'Do While Not rsCalend.EOF'JOEP-Lucv coivd
    For X = 1 To nCuotasPag '- 1 'JOEP-Lucv coivd
        FECalendSimulador.AdicionaFila
        
        If (rsCalend.EOF) Then
            nGastos = 0
            nGastoSegDesg = 0
            nGastoIncendio = 0
            nGastoIncendioGracia = 0
            
            'RIRO 20200829 Liquidación
            nIntCompCalculado = 0
            nDiasCalculo = 0
            nIntCompDiferenciaCapitalizado = 0
            nIntGraciaGenerado = 0
            nIntGraciaCapitalizado = 0
            nIntGraciaAsignado = 0
            'RIRO 20200829 Liquidación
            
            'Cuotas Pendientes
            FECalendSimulador.TextMatrix(X, 1) = MatCalendReprogramado(j, 0)  'FechaCuota (Fila, Colum)
            MatCalend(i, 0) = MatCalendReprogramado(j, 0) 'FechaVenc.
            lnCapital = MatCalendReprogramado(j, 3) 'Capital
            lnIntComp = MatCalendReprogramado(j, 4) 'IntComp
            lnIntGra = MatCalendReprogramado(j, 5) 'IntGrac
                    
            nGastoSegDesg = MatCalendReprogramado(j, 8)
            nGastoIncendio = CDbl(MatCalendReprogramado(j, 15))
            nGastoIncendioGracia = CDbl(MatCalendReprogramado(j, 16)) 'SegInc por días de gracia
                                
            'nGastos = CDbl(CDbl(MatCalendReprogramado(j, 8)) + nMontoPoliza + CDbl(MatCalendReprogramado(j, 16))) 'Add JOEP20200414 Covid
            nGastos = CDbl(CDbl(MatCalendReprogramado(j, 8)) + nMontoPoliza) 'Add JOEP20210306 garantia covid
            j = j + 1
                        
            'Asignación de valores en el Flex del calendario
            FECalendSimulador.TextMatrix(X, 2) = X                                      'Nro. Cuota
            FECalendSimulador.TextMatrix(X, 3) = Format(IIf(IsNull(lnCapital), 0, lnCapital) + _
                                                        IIf(IsNull(lnIntComp), 0, lnIntComp) + _
                                                        IIf(IsNull(lnIntGra), 0, lnIntGra) + _
                                                        0 + _
                                                        0 + _
                                                        nGastos, "#0.00") 'Importe Cuota
            FECalendSimulador.TextMatrix(X, 4) = Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")           'Capital
            FECalendSimulador.TextMatrix(X, 5) = Format(IIf(IsNull(lnIntComp), 0, lnIntComp), "#0.00")           'Interés Compensatorio
            'LUCV20180601, Agregó según ERS022-2018
            FECalendSimulador.TextMatrix(X, 6) = Format(0, "#0.00") 'Interés Moratorio
            'FIN LUCV20180601.
            
            FECalendSimulador.TextMatrix(X, 7) = Format(0, "#0.00")
            FECalendSimulador.TextMatrix(X, 8) = Format(IIf(IsNull(lnIntGra), 0, lnIntGra), "#0.00")                 'Interés de Gracia
            FECalendSimulador.TextMatrix(X, 9) = Format(nGastos, "#0.00")
            nMontoApr = nMontoApr - IIf(IsNull(lnCapital), 0, lnCapital)
            nMontoApr = CDbl(Format(nMontoApr, "#0.0000"))
            FECalendSimulador.TextMatrix(X, 10) = Format(nMontoApr, "#0.00")
            FECalendSimulador.TextMatrix(X, 11) = 0                            'Estado Cuota
            FECalendSimulador.TextMatrix(X, 12) = Format(0, "#0.00")
            lnSaldoNew = lnSaldoNew + IIf(IsNull(lnCapital), 0, lnCapital) - 0
    
            'Asignación de valores a la Matriz del calendario de pagos
            MatCalend(i, 1) = X
            MatCalend(i, 2) = Format(IIf(IsNull(lnCapital), 0, lnCapital) + _
                                    IIf(IsNull(lnIntComp), 0, lnIntComp) + _
                                    IIf(IsNull(lnIntGra), 0, lnIntGra) + _
                                    0 + _
                                    0 + _
                                    nGastos, "#0.00")
            MatCalend(i, 3) = Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")
            MatCalend(i, 4) = Format(IIf(IsNull(lnIntComp), 0, lnIntComp), "#0.00")
            MatCalend(i, 5) = Format(IIf(IsNull(lnIntGra), 0, lnIntGra), "#0.00")
            MatCalend(i, 6) = Format(nGastos, "#0.00")
            MatCalend(i, 7) = Format(lnSaldoNew, "#0.00")
            MatCalend(i, 8) = Format(nGastoSegDesg, "#0.00")
            MatCalend(i, 9) = Format(lnIntComp, "#0.00") 'nInteres1 + nInterespro 'Add JOEP20200415 covid
            
            MatCalend(i, 15) = Format(nGastoIncendio, "#0.00")
            MatCalend(i, 16) = Format(nGastoIncendioGracia, "#0.00")
            MatCalend(i, 17) = 0 'LUCV20180601. Agregó
            i = i + 1
        Else
            nGastos = IIf(IsNull(rsCalend!nGasto), 0, rsCalend!nGasto)
            nGastoSegDesg = 0
            nGastoIncendio = 0
            nGastoIncendioGracia = 0
            
            'RIRO 20200829 Liquidación
            nIntCompCalculado = 0
            nDiasCalculo = 0
            nIntCompDiferenciaCapitalizado = 0
            nIntGraciaGenerado = 0
            nIntGraciaCapitalizado = 0
            nIntGraciaAsignado = 0
            'RIRO 20200829 Liquidación
                  
            'Cuotas Pagadas
            If rsCalend!nColocCalendEstado = gColocCalendEstadoPagado Then
                FECalendSimulador.TextMatrix(rsCalend.Bookmark, 1) = Format(rsCalend!dVenc, "dd/mm/yyyy")
                MatCalend(i, 0) = Format(rsCalend!dVenc, "dd/mm/yyyy")
                lnCapital = rsCalend!nCapital
                lnIntComp = rsCalend!nIntComp
                lnIntGra = rsCalend!nIntGracia
            'Cuotas Pendientes
            Else
                FECalendSimulador.TextMatrix(rsCalend.Bookmark, 1) = MatCalendReprogramado(j, 0)  'FechaCuota (Fila, Colum)
                MatCalend(i, 0) = MatCalendReprogramado(j, 0) 'FechaVenc.
                lnCapital = MatCalendReprogramado(j, 3) 'Capital
                lnIntComp = MatCalendReprogramado(j, 4) 'IntComp
                lnIntGra = MatCalendReprogramado(j, 5) 'IntGrac
                    
                nGastoSegDesg = MatCalendReprogramado(j, 8)
                nGastoIncendio = CDbl(MatCalendReprogramado(j, 15))
                nGastoIncendioGracia = CDbl(MatCalendReprogramado(j, 16)) 'SegInc por días de gracia
                                
                'RIRO 20200829 Liquidación *******************************
                nIntCompCalculado = CDbl(MatCalendReprogramado(j, 17))
                nDiasCalculo = CInt(MatCalendReprogramado(j, 18))
                nIntCompDiferenciaCapitalizado = CDbl(MatCalendReprogramado(j, 19))
                nIntGraciaGenerado = CDbl(MatCalendReprogramado(j, 20))
                nIntGraciaCapitalizado = CDbl(MatCalendReprogramado(j, 21))
                nIntGraciaAsignado = CDbl(MatCalendReprogramado(j, 22))
                'RIRO 20200829 Liquidación *******************************
                                
                'nGastos = CDbl(CDbl(MatCalendReprogramado(j, 8)) + nMontoPoliza + CDbl(MatCalendReprogramado(j, 16))) 'Add JOEP20200414 Covid
                nGastos = CDbl(CDbl(MatCalendReprogramado(j, 8)) + nMontoPoliza) 'Add JOEP20210306 garantia covid
                j = j + 1
            End If
            
            'Asignación de valores en el Flex del calendario
            FECalendSimulador.TextMatrix(rsCalend.Bookmark, 2) = Trim(str(rsCalend!nCuota))                                      'Nro. Cuota
            FECalendSimulador.TextMatrix(rsCalend.Bookmark, 3) = Format(IIf(IsNull(lnCapital), 0, lnCapital) + _
                                                        IIf(IsNull(lnIntComp), 0, lnIntComp) + _
                                                        IIf(IsNull(lnIntGra), 0, lnIntGra) + _
                                                        IIf(rsCalend!nColocCalendEstado = gColocCalendEstadoPagado, IIf(IsNull(rsCalend!nIntMor), 0, rsCalend!nIntMor), 0) + _
                                                        IIf(IsNull(rsCalend!nIntReprog), 0, rsCalend!nIntReprog) + _
                                                        nGastos, "#0.00")                                               'Importe Cuota
            FECalendSimulador.TextMatrix(rsCalend.Bookmark, 4) = Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")           'Capital
            FECalendSimulador.TextMatrix(rsCalend.Bookmark, 5) = Format(IIf(IsNull(lnIntComp), 0, lnIntComp), "#0.00")           'Interés Compensatorio
            'LUCV20180601, Agregó según ERS022-2018
            If rsCalend!nColocCalendEstado = gColocCalendEstadoPagado Then
                FECalendSimulador.TextMatrix(rsCalend.Bookmark, 6) = Format(IIf(IsNull(rsCalend!nIntMor), 0, rsCalend!nIntMor), "#0.00") 'Interés Moratorio
            Else
                FECalendSimulador.TextMatrix(rsCalend.Bookmark, 6) = Format(0, "#0.00") 'Interés Moratorio
            End If
            'FIN LUCV20180601.
            
            FECalendSimulador.TextMatrix(rsCalend.Bookmark, 7) = Format(IIf(IsNull(rsCalend!nIntReprog), 0, rsCalend!nIntReprog), "#0.00")
            FECalendSimulador.TextMatrix(rsCalend.Bookmark, 8) = Format(IIf(IsNull(lnIntGra), 0, lnIntGra), "#0.00")                 'Interés de Gracia
            FECalendSimulador.TextMatrix(rsCalend.Bookmark, 9) = Format(nGastos, "#0.00")
            nMontoApr = nMontoApr - IIf(IsNull(lnCapital), 0, lnCapital)
            nMontoApr = CDbl(Format(nMontoApr, "#0.0000"))
            FECalendSimulador.TextMatrix(rsCalend.Bookmark, 10) = Format(nMontoApr, "#0.00")
            FECalendSimulador.TextMatrix(rsCalend.Bookmark, 11) = Trim(str(rsCalend!nColocCalendEstado))                             'Estado Cuota
            FECalendSimulador.TextMatrix(rsCalend.Bookmark, 12) = Format(IIf(IsNull(rsCalend!nCapitalPag), 0, rsCalend!nCapitalPag), "#0.00")
            lnSaldoNew = lnSaldoNew + IIf(IsNull(lnCapital), 0, lnCapital) - IIf(IsNull(rsCalend!nCapitalPag), 0, rsCalend!nCapitalPag)
    
            'Asignación de valores a la Matriz del calendario de pagos
            MatCalend(i, 1) = Trim(str(rsCalend!nCuota))
            MatCalend(i, 2) = Format(IIf(IsNull(lnCapital), 0, lnCapital) + _
                                    IIf(IsNull(lnIntComp), 0, lnIntComp) + _
                                    IIf(IsNull(lnIntGra), 0, lnIntGra) + _
                                    IIf(IsNull(rsCalend!nIntMor), 0, rsCalend!nIntMor) + _
                                    IIf(IsNull(rsCalend!nIntReprog), 0, rsCalend!nIntReprog) + _
                                    nGastos, "#0.00")
            MatCalend(i, 3) = Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")
            MatCalend(i, 4) = Format(IIf(IsNull(lnIntComp), 0, lnIntComp), "#0.00")
            MatCalend(i, 5) = Format(IIf(IsNull(lnIntGra), 0, lnIntGra), "#0.00")
            MatCalend(i, 6) = Format(nGastos, "#0.00")
            MatCalend(i, 7) = Format(lnSaldoNew, "#0.00")
            MatCalend(i, 8) = Format(nGastoSegDesg, "#0.00")
            MatCalend(i, 9) = Format(lnIntComp, "#0.00") 'nInteres1 + nInterespro 'Add JOEP20200415 covid
            
            MatCalend(i, 15) = Format(nGastoIncendio, "#0.00")
            MatCalend(i, 16) = Format(nGastoIncendioGracia, "#0.00")
            MatCalend(i, 17) = rsCalend!nColocCalendEstado 'LUCV20180601. Agregó
            
            'RIRO 20200825 Corrección Liquidación
            MatCalend(i, 18) = nIntCompCalculado
            MatCalend(i, 19) = nDiasCalculo
            MatCalend(i, 20) = nIntCompDiferenciaCapitalizado
            MatCalend(i, 21) = nIntGraciaGenerado
            MatCalend(i, 22) = nIntGraciaCapitalizado
            MatCalend(i, 23) = nIntGraciaAsignado
            'RIRO 20200825 Corrección Liquidación
        'Add JOEP20210306 garantia covid
            MatCalend(i, 24) = rsCalend!nNroCalen
            MatCalend(i, 25) = rsCalend!dPago
            MatCalend(i, 26) = nOpCovid
        'Add JOEP20210306 garantia covid
            If rsCalend!nColocCalendEstado = gColocCalendEstadoPagado Then
                FECalendSimulador.row = rsCalend.Bookmark
                Call FECalendSimulador.ForeColorRow(vbRed)
            End If
            If rsCalend.RecordCount = rsCalend.Bookmark Then
                lblfecUltCuotaSimulador = Format(rsCalend!dVenc, "dd/mm/yyyy")
            End If
    
            i = i + 1
            rsCalend.MoveNext
        End If
        
            nTotalCapital = nTotalCapital + Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")
    Next X
        
    Set oDCOMCredito = New COMDCredito.DCOMCredito
    Set rsTipPeriodo = oDCOMCredito.IdentificarTipoPeriodo(ActxCtaSimulador.NroCuenta)
    If Not (rsTipPeriodo.EOF And rsTipPeriodo.BOF) Then
        nTipoPeriodo = rsTipPeriodo!nTpPeriodo
    End If
    
    NewTCEA = oNCOMCredito.GeneraTasaCostoEfectivoAnual(CDate(Format(ldVigencia, "dd/mm/yyyy")), CDbl(LblPrestamoSimulador), MatCalend, CDbl(fnTasaInteres), ActxCtaSimulador.NroCuenta, nTipoPeriodo)  'Para calcular la TCEA
    txtTCEAReprogSimulador = NewTCEA & " %"
    rsCalend.Close
    Set rsCalend = Nothing
    lblSaldoRepSimulador = Format(lnSaldoNew, "#,##0.00")
    
Else
    'Generacion del calendario Reprogramado
    'ReDim MatCalend(nCuoPag, 23) 'LUCV20180601, Modificó 11 por 17 / RIRO 20200829 de 17 a 23
    ReDim MatCalend(nCuoPag, 27) 'Add JOEP20210306 garantia covid
    Do While Not rsCalend.EOF
        FECalendSimulador.AdicionaFila
        nGastos = IIf(IsNull(rsCalend!nGasto), 0, rsCalend!nGasto)
        nGastoSegDesg = 0
        nGastoIncendio = 0
        nGastoIncendioGracia = 0
        
        'RIRO 20200829 Liquidación
        nIntCompCalculado = 0
        nDiasCalculo = 0
        nIntCompDiferenciaCapitalizado = 0
        nIntGraciaGenerado = 0
        nIntGraciaCapitalizado = 0
        nIntGraciaAsignado = 0
        'RIRO 20200829 Liquidación
        
        'Cuotas Pagadas
        If rsCalend!nColocCalendEstado = gColocCalendEstadoPagado Then
            FECalendSimulador.TextMatrix(rsCalend.Bookmark, 1) = Format(rsCalend!dVenc, "dd/mm/yyyy")
            MatCalend(i, 0) = Format(rsCalend!dVenc, "dd/mm/yyyy")
            lnCapital = rsCalend!nCapital
            lnIntComp = rsCalend!nIntComp
            lnIntGra = rsCalend!nIntGracia
        'Add JOEP20210306 garantia covid
            nGastoIncendio = rsCalend!nGastoPolizaIncendio
            nGastoIncendioGracia = rsCalend!nGastoPolizaIncendioGracia
        'Add JOEP20210306 garantia covid
        'Cuotas Pendientes
        Else
            FECalendSimulador.TextMatrix(rsCalend.Bookmark, 1) = MatCalendReprogramado(j, 0)  'FechaCuota (Fila, Colum)
            MatCalend(i, 0) = MatCalendReprogramado(j, 0) 'FechaVenc.
            
            lnCapital = MatCalendReprogramado(j, 3) 'Capital
            lnIntComp = MatCalendReprogramado(j, 4) 'IntComp
            lnIntGra = MatCalendReprogramado(j, 5) 'IntGrac
            
            nGastoSegDesg = MatCalendReprogramado(j, 8)
            nGastoIncendio = CDbl(MatCalendReprogramado(j, 15))
            nGastoIncendioGracia = CDbl(MatCalendReprogramado(j, 16)) 'SegInc por días de gracia
            
            nIntCompCalculado = CDbl(MatCalendReprogramado(j, 17)) 'RIRO 20200825 Interés Compensatorio Calculado
            nDiasCalculo = CInt(MatCalendReprogramado(j, 18)) 'RIRO 20200825 Interés Compensatorio Calculado
            nIntCompDiferenciaCapitalizado = CDbl(MatCalendReprogramado(j, 19)) 'RIRO 20200825 Interés Compensatorio Calculado
            nIntGraciaGenerado = CDbl(MatCalendReprogramado(j, 20)) 'RIRO 20200825 Interés Compensatorio Calculado
            nIntGraciaCapitalizado = CDbl(MatCalendReprogramado(j, 21)) 'RIRO 20200825 Interés Compensatorio Calculado
            nIntGraciaAsignado = CDbl(MatCalendReprogramado(j, 22)) 'RIRO 20200825 Interés Compensatorio Calculado
                                                  
            'nGastos = CDbl(CDbl(MatCalendReprogramado(j, 8)) + CDbl(MatCalendReprogramado(j, 15)) + CDbl(MatCalendReprogramado(j, 16)))'Comento JOEP20200414 Covid
            'nGastos = CDbl(CDbl(MatCalendReprogramado(j, 8)) + nMontoPoliza + CDbl(MatCalendReprogramado(j, 16))) 'Add JOEP20200414 Covid
            nGastos = CDbl(CDbl(MatCalendReprogramado(j, 8)) + nMontoPoliza) 'Add JOEP20210306 garantia covid
            j = j + 1
        End If

        'Asignación de valores en el Flex del calendario
        FECalendSimulador.TextMatrix(rsCalend.Bookmark, 2) = Trim(str(rsCalend!nCuota))                                      'Nro. Cuota
        FECalendSimulador.TextMatrix(rsCalend.Bookmark, 3) = Format(IIf(IsNull(lnCapital), 0, lnCapital) + _
                                                    IIf(IsNull(lnIntComp), 0, lnIntComp) + _
                                                    IIf(IsNull(lnIntGra), 0, lnIntGra) + _
                                                    IIf(rsCalend!nColocCalendEstado = gColocCalendEstadoPagado, IIf(IsNull(rsCalend!nIntMor), 0, rsCalend!nIntMor), 0) + _
                                                    IIf(IsNull(rsCalend!nIntReprog), 0, rsCalend!nIntReprog) + _
                                                    nGastos, "#0.00")                                               'Importe Cuota
        FECalendSimulador.TextMatrix(rsCalend.Bookmark, 4) = Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")           'Capital
        FECalendSimulador.TextMatrix(rsCalend.Bookmark, 5) = Format(IIf(IsNull(lnIntComp), 0, lnIntComp), "#0.00")           'Interés Compensatorio
        'LUCV20180601, Agregó según ERS022-2018
        If rsCalend!nColocCalendEstado = gColocCalendEstadoPagado Then
            FECalendSimulador.TextMatrix(rsCalend.Bookmark, 6) = Format(IIf(IsNull(rsCalend!nIntMor), 0, rsCalend!nIntMor), "#0.00") 'Interés Moratorio
        Else
            FECalendSimulador.TextMatrix(rsCalend.Bookmark, 6) = Format(0, "#0.00") 'Interés Moratorio
        End If
        'FIN LUCV20180601.
        
        FECalendSimulador.TextMatrix(rsCalend.Bookmark, 7) = Format(IIf(IsNull(rsCalend!nIntReprog), 0, rsCalend!nIntReprog), "#0.00")
        FECalendSimulador.TextMatrix(rsCalend.Bookmark, 8) = Format(IIf(IsNull(lnIntGra), 0, lnIntGra), "#0.00")                 'Interés de Gracia
        FECalendSimulador.TextMatrix(rsCalend.Bookmark, 9) = Format(nGastos, "#0.00")
        nMontoApr = nMontoApr - IIf(IsNull(lnCapital), 0, lnCapital)
        nMontoApr = CDbl(Format(nMontoApr, "#0.0000"))
        FECalendSimulador.TextMatrix(rsCalend.Bookmark, 10) = Format(nMontoApr, "#0.00")
        FECalendSimulador.TextMatrix(rsCalend.Bookmark, 11) = Trim(str(rsCalend!nColocCalendEstado))                             'Estado Cuota
        FECalendSimulador.TextMatrix(rsCalend.Bookmark, 12) = Format(IIf(IsNull(rsCalend!nCapitalPag), 0, rsCalend!nCapitalPag), "#0.00")
        lnSaldoNew = lnSaldoNew + IIf(IsNull(lnCapital), 0, lnCapital) - IIf(IsNull(rsCalend!nCapitalPag), 0, rsCalend!nCapitalPag)

        'Asignación de valores a la Matriz del calendario de pagos
        MatCalend(i, 1) = Trim(str(rsCalend!nCuota))
        MatCalend(i, 2) = Format(IIf(IsNull(lnCapital), 0, lnCapital) + _
                                IIf(IsNull(lnIntComp), 0, lnIntComp) + _
                                IIf(IsNull(lnIntGra), 0, lnIntGra) + _
                                IIf(IsNull(rsCalend!nIntMor), 0, rsCalend!nIntMor) + _
                                IIf(IsNull(rsCalend!nIntReprog), 0, rsCalend!nIntReprog) + _
                                nGastos, "#0.00")
        MatCalend(i, 3) = Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")
        MatCalend(i, 4) = Format(IIf(IsNull(lnIntComp), 0, lnIntComp), "#0.00")
        MatCalend(i, 5) = Format(IIf(IsNull(lnIntGra), 0, lnIntGra), "#0.00")
        MatCalend(i, 6) = Format(nGastos, "#0.00")
        MatCalend(i, 7) = Format(lnSaldoNew, "#0.00")
        'MatCalend(i, 8) = rsCalend!nColocCalendEstado 'LUCV20180601, Comentó
        MatCalend(i, 8) = Format(nGastoSegDesg, "#0.00")
        'MatCalend(i, 9) = rsCalend!nIntCompPag + lnIntComp 'nInteres1 + nInterespro 'Comento JOEP20200415 covid
        MatCalend(i, 9) = Format(lnIntComp, "#0.00") 'nInteres1 + nInterespro 'Add JOEP20200415 covid
        
        MatCalend(i, 15) = Format(nGastoIncendio, "#0.00")
        MatCalend(i, 16) = Format(nGastoIncendioGracia, "#0.00")
        MatCalend(i, 17) = rsCalend!nColocCalendEstado 'LUCV20180601. Agregó
        
        'RIRO 20200825 Corrección Liquidación
        MatCalend(i, 18) = nIntCompCalculado
        MatCalend(i, 19) = nDiasCalculo
        MatCalend(i, 20) = nIntCompDiferenciaCapitalizado
        MatCalend(i, 21) = nIntGraciaGenerado
        MatCalend(i, 22) = nIntGraciaCapitalizado
        MatCalend(i, 23) = nIntGraciaAsignado
        'RIRO 20200825 Corrección Liquidación
    'Add JOEP20210306 garantia covid
        MatCalend(i, 24) = rsCalend!nNroCalen
        MatCalend(i, 25) = rsCalend!dPago
        MatCalend(i, 26) = nOpCovid
    'Add JOEP20210306 garantia covid
    
        If rsCalend!nColocCalendEstado = gColocCalendEstadoPagado Then
            FECalendSimulador.row = rsCalend.Bookmark
            Call FECalendSimulador.ForeColorRow(vbRed)
        End If
        If rsCalend.RecordCount = rsCalend.Bookmark Then
            lblfecUltCuotaSimulador = Format(rsCalend!dVenc, "dd/mm/yyyy")
        End If

        i = i + 1
        rsCalend.MoveNext

        nTotalCapital = nTotalCapital + Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")
    Loop

    'JOEP Identificar Tipo de Periodo(Para calcular TCEA)
    Set oDCOMCredito = New COMDCredito.DCOMCredito
    Set rsTipPeriodo = oDCOMCredito.IdentificarTipoPeriodo(ActxCtaSimulador.NroCuenta)
    If Not (rsTipPeriodo.EOF And rsTipPeriodo.BOF) Then
        nTipoPeriodo = rsTipPeriodo!nTpPeriodo
    End If
    
    NewTCEA = oNCOMCredito.GeneraTasaCostoEfectivoAnual(CDate(Format(ldVigencia, "dd/mm/yyyy")), CDbl(LblPrestamoSimulador), MatCalend, CDbl(fnTasaInteres), ActxCtaSimulador.NroCuenta, nTipoPeriodo)  'Para calcular la TCEA
    txtTCEAReprogSimulador = NewTCEA & " %" 'Add JOEP20200425 Cuota Igual
    rsCalend.Close
    Set rsCalend = Nothing
    lblSaldoRepSimulador = Format(lnSaldoNew, "#,##0.00")
End If

End Sub
'<-*****Fin LUCV20180601

Private Sub CmbReprogNatEspecialesSimulador_Click()
Dim rsCovidOpciones As ADODB.Recordset
Dim oDCOMCred As COMDConstantes.DCOMConstantes
Set oDCOMCred = New COMDConstantes.DCOMConstantes

Dim rsMsgbox As ADODB.Recordset
Dim objMsg As COMDCredito.DCOMCredito

    If Right(CmbReprogNatEspecialesSimulador.Text, 2) = 1 Then
        Set rsCovidOpciones = oDCOMCred.RecuperaConstantes(2090)
        If Not (rsCovidOpciones.BOF And rsCovidOpciones.EOF) Then
            Call Llenar_Combo_con_Recordset(rsCovidOpciones, cmbCovidReprogSimulador)
            Call CambiaTamañoCombo(cmbCovidReprogSimulador, 100)
        End If
        cmbCovidReprogSimulador.Visible = True
    Else
        cmbCovidReprogSimulador.ListIndex = -1
        cmbCovidReprogSimulador.Visible = False
    End If
Set oDCOMCred = Nothing
RSClose rsCovidOpciones
End Sub

Private Sub EnableBotones(ByVal Opcion As Integer, Optional ByVal bValor As Boolean)
    If Opcion = 0 Then
        If CmbReprogNatEspecialesSimulador.Visible = True And Right(cmbCovidReprogSimulador.Text, 1) = "2" Then
            CmbReprogNatEspecialesSimulador.Enabled = bValor
            cmbCovidReprogSimulador.Enabled = bValor
        End If
    End If
End Sub

Private Sub VisibleBotones(ByVal Opcion As Integer, Optional ByVal bValor As Boolean)
    If Opcion = 0 Then
        If CmbReprogNatEspecialesSimulador.Enabled = True And Right(cmbCovidReprogSimulador.Text, 1) = "2" Then
            frMontCuotaCovidSimulador.Visible = bValor
            txtMontoCuotaSimulador.Enabled = False
            fr_TasaEspecialSimulador.Visible = False
        ElseIf CmbReprogNatEspecialesSimulador.Enabled = True And Right(cmbCovidReprogSimulador.Text, 1) = "4" Then
            frMontCuotaCovidSimulador.Visible = bValor
            txtMontoCuotaSimulador.Enabled = bValor
            frMontCuotaCovidSimulador.Enabled = bValor
            txtMontoCuotaSimulador.Text = Format(0#, "#,#0.00")
            fr_TasaEspecialSimulador.Visible = False
        ElseIf CmbReprogNatEspecialesSimulador.Enabled = True And Right(cmbCovidReprogSimulador.Text, 1) = "3" Then
            fr_TasaEspecialSimulador.Visible = bValor
            frMontCuotaCovidSimulador.Visible = False
    'Add JOEP20210306 garantia covid
        ElseIf CmbReprogNatEspecialesSimulador.Enabled = True And Right(cmbCovidReprogSimulador.Text, 1) = "5" Then
            fr_TasaEspecialSimulador.Visible = bValor
            frMontCuotaCovidSimulador.Visible = False
    'Add JOEP20210306 garantia covid
        Else
            frMontCuotaCovidSimulador.Visible = False
            txtMontoCuotaSimulador.Text = Format(0#, "#,#0.00")
            fr_TasaEspecialSimulador.Visible = False
        End If
    End If
End Sub

Private Sub cmbCovidReprogSimulador_Click()
Dim rsMC As ADODB.Recordset
Dim rsTasaEspecialCovid As ADODB.Recordset
Dim obMc As COMDCredito.DCOMCredito
Set obMc = New COMDCredito.DCOMCredito
nMontoCuota = 0

If CmbReprogNatEspecialesSimulador.Enabled = True And cmbCovidReprogSimulador.Enabled = True Then
    If Right(cmbCovidReprogSimulador.Text, 1) = "2" Or Right(cmbCovidReprogSimulador.Text, 1) = "4" Then
        'Set rsMC = obMc.ReprogramacionObtMantenerCuota(ActxCtaSimulador.NroCuenta)
        Set rsMC = obMc.ReprogramacionObtMantenerCuota(ActxCtaSimulador.NroCuenta, gnMenuOpcion, Right(cmbCovidReprogSimulador.Text, 1)) 'Add JOEP20210310 identificar facilidad de reprogramacion
            If Not (rsMC.BOF And rsMC.EOF) Then
                txtMontoCuotaSimulador.Text = Format(rsMC!nMontoCuota, "#,#0.00")
                nMontoCuota = Format(txtMontoCuotaSimulador, "#,#0.00")
                Call VisibleBotones(0, True)
            End If
    ElseIf Right(cmbCovidReprogSimulador.Text, 1) = "3" Then
        'Set rsTasaEspecialCovid = obMc.ReprogramacionObtTasaEspecial(ActxCtaSimulador.NroCuenta, LblSaldoSimulador)
        'Set rsTasaEspecialCovid = obMc.ReprogramacionObtTasaEspecial(ActxCtaSimulador.NroCuenta, LblSaldoSimulador, gnMenuOpcion) 'Add JOEP20210306 garantia covid
        Set rsTasaEspecialCovid = obMc.ReprogramacionObtTasaEspecial(ActxCtaSimulador.NroCuenta, LblSaldoSimulador, gnMenuOpcion, Right(cmbCovidReprogSimulador.Text, 1)) 'Add JOEP20210310 identificar facilidad de reprogramacion
            If Not (rsTasaEspecialCovid.BOF And rsTasaEspecialCovid.EOF) Then
                lbl_TasaEspecialSimulador = Format(rsTasaEspecialCovid!nTasaInteres, "0.00")
                Call VisibleBotones(0, True)
            End If
    ElseIf Right(cmbCovidReprogSimulador.Text, 1) = "5" Then
        'Set rsTasaEspecialCovid = obMc.ReprogramacionObtTasaEspecial(ActxCtaSimulador.NroCuenta, LblSaldoSimulador, gnMenuOpcion) 'Add JOEP20210211 garantia covid
        Set rsTasaEspecialCovid = obMc.ReprogramacionObtTasaEspecial(ActxCtaSimulador.NroCuenta, LblSaldoSimulador, gnMenuOpcion, Right(cmbCovidReprogSimulador.Text, 1)) 'Add JOEP20210310 identificar facilidad de reprogramacion
            If Not (rsTasaEspecialCovid.BOF And rsTasaEspecialCovid.EOF) Then
                lbl_TasaEspecialSimulador = Format(rsTasaEspecialCovid!nTasaInteres, "0.00")
                Call VisibleBotones(0, True)
            End If
    Else
        Call VisibleBotones(0, True)
    End If
End If

Set obMc = Nothing
RSClose rsMC
End Sub

Private Sub CargaCombo()
    Dim rsReprogEmergencia As ADODB.Recordset
    Dim rsRepgOCM As ADODB.Recordset
    Dim obOCM As COMDCredito.DCOMCredito
    Dim oDCOMCred As COMDConstantes.DCOMConstantes
    Set oDCOMCred = New COMDConstantes.DCOMConstantes
    Set obOCM = New COMDCredito.DCOMCredito
    Dim bOCM As Integer
    bOCM = 0 'JOEP20200928 Reprogramacion OCM

    'Set rsRepgOCM = obOCM.SimuladoReprogramacion(ActxCtaSimulador.NroCuenta)
    Set rsRepgOCM = obOCM.SimuladoReprogramacion(ActxCtaSimulador.NroCuenta, gnMenuOpcion) 'Add JOEP20210306 garantia covid
    If Not (rsRepgOCM.BOF And rsRepgOCM.EOF) Then
        bOCM = rsRepgOCM!Datos
    End If
    
    'Set rsReprogEmergencia = oDCOMCred.RecuperaConstanteReprogaramacion(2080, bOCM)
    Set rsReprogEmergencia = oDCOMCred.RecuperaConstanteReprogaramacion(2080, bOCM, gnMenuOpcion) 'Add JOEP20210306 garantia covid
    If Not (rsReprogEmergencia.BOF And rsReprogEmergencia.EOF) Then
        Call Llenar_Combo_con_Recordset(rsReprogEmergencia, CmbReprogNatEspecialesSimulador)
        Call CambiaTamañoCombo(CmbReprogNatEspecialesSimulador, 200)
    End If
    
    Set oDCOMCred = Nothing
    Set obOCM = Nothing
    RSClose rsRepgOCM
    RSClose rsReprogEmergencia
                
End Sub

Private Function ValidaDatos(ByVal pnBoton As Integer) As Boolean
Dim rsVD As ADODB.Recordset
Dim obVD As COMDCredito.DCOMCredito
Set obVD = New COMDCredito.DCOMCredito
ValidaDatos = False

Dim nCapital As Currency
Dim dFechaVencFinReprg As String
Dim nMontoNegativo As Integer
nCapital = 0
dFechaVencFinReprg = ""
nMontoNegativo = 0
Dim i As Integer
    For i = 1 To (FECalendSimulador.rows - 1)
        If FECalendSimulador.TextMatrix(i, 11) = 0 Then
            nCapital = nCapital + FECalendSimulador.TextMatrix(i, 4)
            dFechaVencFinReprg = FECalendSimulador.TextMatrix(i, 1)
             If (FECalendSimulador.TextMatrix(i, 4) < 0 Or FECalendSimulador.TextMatrix(i, 5) < 0 Or FECalendSimulador.TextMatrix(i, 9) < 0) Then
                nMontoNegativo = 1
            End If
        End If
    Next i

'Set rsVD = obVD.ReprogramacionValidDatos(ActxCtaSimulador.NroCuenta, "Glosa", CCur(txtMontoCuotaSimulador.Text), nCapital, IIf(CmbReprogNatEspecialesSimulador.Visible, 1, 0), IIf(cmbCovidReprogSimulador.Visible, 1, 0), IIf(cmbCovidReprogSimulador.Visible = True, Right(cmbCovidReprogSimulador.Text, 1), -1), nMontoCuota, CCur(lbl_TasaEspecialSimulador), CCur(LblTasaSimulador), dFechaVencFinReprg, pnBoton)
Set rsVD = obVD.ReprogramacionValidDatos(ActxCtaSimulador.NroCuenta, "ModSimulacion", CCur(txtMontoCuotaSimulador.Text), nCapital, IIf(CmbReprogNatEspecialesSimulador.Visible, 1, 0), IIf(cmbCovidReprogSimulador.Visible, 1, 0), IIf(cmbCovidReprogSimulador.Visible = True, Right(cmbCovidReprogSimulador.Text, 1), -1), nMontoCuota, CCur(lbl_TasaEspecialSimulador), CCur(LblTasaSimulador), dFechaVencFinReprg, pnBoton, nMontoNegativo) 'Add JOEP20210306 garantia covid

If Not (rsVD.BOF And rsVD.EOF) Then
    If rsVD!MsgBox <> "" Then
        MsgBox rsVD!MsgBox, vbInformation, "Aviso"
        ValidaDatos = True
    End If
End If
Set obVD = Nothing
RSClose rsVD
End Function

Private Sub OCMControl(ByVal pnModalidad As Integer)
    CmbReprogNatEspecialesSimulador.Visible = True
    CmbReprogNatEspecialesSimulador.ListIndex = 0
    cmbCovidReprogSimulador.ListIndex = pnModalidad - 1
    CmbReprogNatEspecialesSimulador.Enabled = False
    cmbCovidReprogSimulador.Enabled = False
End Sub

Public Function DisableCloseButton(frm As Form) As Boolean
'PURPOSE: Removes X button from a form
'EXAMPLE: DisableCloseButton Me
'RETURNS: True if successful, false otherwise
'NOTES:   Also removes Exit Item from
'         Control Box Menu
    Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
    
    lHndSysMenu = GetSystemMenu(frm.hwnd, 0)

    'remove close button
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)
   'Remove seperator bar
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
    'Return True if both calls were successful
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)
End Function
