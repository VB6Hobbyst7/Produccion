VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCapDepositoLote 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   Icon            =   "frmCapDepositoLote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7980
      Left            =   45
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   45
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   14076
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Deposito en Lote"
      TabPicture(0)   =   "frmCapDepositoLote.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dlgArchivo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraMonto"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraGlosa"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSalir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdGrabar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraCuenta"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraTipo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraTranferecia"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCancelar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "pbProgres"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin ComctlLib.ProgressBar pbProgres 
         Height          =   195
         Left            =   2295
         TabIndex        =   40
         Top             =   7650
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancela&r"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         TabIndex        =   9
         Top             =   7560
         Width           =   1095
      End
      Begin VB.Frame fraTranferecia 
         Caption         =   "Transferencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   2115
         Left            =   135
         TabIndex        =   25
         Top             =   5400
         Visible         =   0   'False
         Width           =   4410
         Begin VB.ComboBox cboTransferMoneda 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   195
            Width           =   1575
         End
         Begin VB.CommandButton cmdTranfer 
            Height          =   315
            Left            =   2520
            Picture         =   "frmCapDepositoLote.frx":0326
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   555
            Width           =   475
         End
         Begin VB.TextBox txtTransferGlosa 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   840
            MaxLength       =   255
            TabIndex        =   26
            Top             =   1290
            Width           =   3465
         End
         Begin VB.Label lbltransferBco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   39
            Top             =   930
            Width           =   3465
         End
         Begin VB.Label lbltransferN 
            AutoSize        =   -1  'True
            Caption         =   "Nro Doc :"
            Height          =   195
            Left            =   60
            TabIndex        =   38
            Top             =   570
            Width           =   690
         End
         Begin VB.Label lbltransferBcol 
            AutoSize        =   -1  'True
            Caption         =   "Banco :"
            Height          =   195
            Left            =   60
            TabIndex        =   37
            Top             =   930
            Width           =   555
         End
         Begin VB.Label lblTrasferND 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   36
            Top             =   555
            Width           =   1575
         End
         Begin VB.Label lblTransferMoneda 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   60
            TabIndex        =   35
            Top             =   225
            Width           =   585
         End
         Begin VB.Label lblTransferGlosa 
            AutoSize        =   -1  'True
            Caption         =   "Glosa :"
            Height          =   195
            Left            =   60
            TabIndex        =   34
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label lblTTCC 
            Caption         =   "TCC"
            Height          =   285
            Left            =   3120
            TabIndex        =   33
            Top             =   180
            Width           =   390
         End
         Begin VB.Label Label11 
            Caption         =   "TCV"
            Height          =   285
            Left            =   3120
            TabIndex        =   32
            Top             =   480
            Width           =   390
         End
         Begin VB.Label lblTTCCD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3570
            TabIndex        =   31
            Top             =   165
            Width           =   735
         End
         Begin VB.Label lblTTCVD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3570
            TabIndex        =   30
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblMonTra 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2925
            TabIndex        =   29
            Top             =   1680
            Width           =   1365
         End
         Begin VB.Label lblSimTra 
            AutoSize        =   -1  'True
            Caption         =   "S/."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   2370
            TabIndex        =   28
            Top             =   1680
            Width           =   300
         End
         Begin VB.Label lblEtiMonTra 
            AutoSize        =   -1  'True
            Caption         =   "Monto Transacción"
            Height          =   195
            Left            =   870
            TabIndex        =   27
            Top             =   1710
            Width           =   1380
         End
      End
      Begin VB.Frame fraTipo 
         Caption         =   "Datos Generales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1080
         Left            =   135
         TabIndex        =   21
         Top             =   405
         Width           =   9570
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   270
            Width           =   1455
         End
         Begin VB.TextBox lblInst 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   3690
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   630
            Width           =   5730
         End
         Begin SICMACT.TxtBuscar txtInstitucion 
            Height          =   315
            Left            =   3690
            TabIndex        =   1
            Top             =   270
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TipoBusqueda    =   3
            sTitulo         =   ""
            TipoBusPers     =   1
            EnabledText     =   0   'False
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   315
            Width           =   810
         End
         Begin VB.Label lblInstEtq 
            AutoSize        =   -1  'True
            Caption         =   "Institución :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2610
            TabIndex        =   23
            Top             =   315
            Width           =   1020
         End
      End
      Begin VB.Frame fraCuenta 
         Caption         =   "Datos Cuenta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3840
         Left            =   135
         TabIndex        =   19
         Top             =   1530
         Width           =   9570
         Begin VB.CommandButton cmdFormato 
            Caption         =   "&Formato"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   135
            TabIndex        =   3
            Top             =   3330
            Width           =   915
         End
         Begin VB.TextBox txtArchivo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   3330
            Width           =   4245
         End
         Begin VB.CommandButton cmdCargar 
            Caption         =   "&Cargar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6795
            TabIndex        =   6
            Top             =   3330
            Width           =   840
         End
         Begin VB.CommandButton cmdBusca 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6210
            TabIndex        =   5
            Top             =   3330
            Width           =   495
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8595
            TabIndex        =   8
            Top             =   3330
            Width           =   855
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7695
            TabIndex        =   7
            Top             =   3330
            Width           =   855
         End
         Begin SICMACT.FlexEdit grdCuenta 
            Height          =   3015
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   5318
            Cols0           =   7
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Cod Cliente-Cuenta-Nombre-DOI-Monto-campo1"
            EncabezadosAnchos=   "500-1500-1800-4000-1200-1200-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-X-X-X-5-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-1-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-L-C-C"
            FormatosEdit    =   "0-0-0-0-0-2-2"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   495
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Archivo :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1170
            TabIndex        =   20
            Top             =   3375
            Width           =   780
         End
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7650
         TabIndex        =   10
         Top             =   7560
         Width           =   960
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8730
         TabIndex        =   11
         Top             =   7560
         Width           =   960
      End
      Begin VB.Frame fraGlosa 
         Caption         =   "Glosa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2115
         Left            =   4590
         TabIndex        =   17
         Top             =   5400
         Width           =   2430
         Begin RichTextLib.RichTextBox txtGlosa 
            Height          =   1770
            Left            =   90
            TabIndex        =   18
            Top             =   225
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   3122
            _Version        =   393217
            Appearance      =   0
            TextRTF         =   $"frmCapDepositoLote.frx":0768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame fraMonto 
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2115
         Left            =   7065
         TabIndex        =   15
         Top             =   5400
         Width           =   2655
         Begin SICMACT.EditMoney txtMonto 
            Height          =   315
            Left            =   675
            TabIndex        =   41
            Top             =   360
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            BackColor       =   12648447
            Text            =   "0.00"
            Enabled         =   -1  'True
         End
         Begin VB.Label lblMon 
            AutoSize        =   -1  'True
            Caption         =   "S/."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   2160
            TabIndex        =   42
            Top             =   360
            Width           =   300
         End
         Begin VB.Label Label1 
            Caption         =   "Monto:"
            Height          =   240
            Left            =   90
            TabIndex        =   16
            Top             =   405
            Width           =   555
         End
      End
      Begin MSComDlg.CommonDialog dlgArchivo 
         Left            =   1575
         Top             =   7380
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
   End
End
Attribute VB_Name = "frmCapDepositoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************************************************************************
'* NOMBRE         : "frmCapDepositoLote"
'* DESCRIPCION    : Formulario creado para el pago en lote.
'* CREACION       : RIRO, 20140430 10:00 AM
'*********************************************************************************************************************************************************
Option Explicit

Private fnProducto As Producto
Private fnOpeCod As CaptacOperacion
Private fsDescOperacion As String
Private fnMoneda As COMDConstantes.Moneda
Private bCargaLote As Boolean
Private nNroDeposito As Integer
Private fnMovNroRVD As Long
Private lnMovNroTransfer As Long
Private nMoneda As COMDConstantes.Moneda

Private Sub cboMoneda_Click()
    nMoneda = CLng(Right(cboMoneda.Text, 1))
    If nMoneda = gMonedaNacional Then
        txtMonto.BackColor = &HC0FFFF
        lblMon.Caption = "S/."
    ElseIf nMoneda = gMonedaExtranjera Then
        txtMonto.BackColor = &HC0FFC0
        lblMon.Caption = "US$"
    End If
    If fnOpeCod = gAhoDepositoHaberesEnLoteTransf Then
        cboTransferMoneda.ListIndex = IndiceListaCombo(cboTransferMoneda, Trim(Right(cboMoneda.Text, 5)))
        SetDatosTransferencia "", "", "", 0, -1, ""
        nNroDeposito = 0
        fnMovNroRVD = 0
        lnMovNroTransfer = 0
    End If
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtInstitucion.SetFocus
    End If
End Sub

Private Sub cboTransferMoneda_Click()
    If nMoneda = gMonedaNacional Then
        lblMonTra.BackColor = &HC0FFFF
        lblSimTra.Caption = "S/."
    ElseIf nMoneda = gMonedaExtranjera Then
        lblMonTra.BackColor = &HC0FFC0
        lblSimTra.Caption = "US$"
    End If
End Sub

Private Sub cmdAgregar_Click()
    If bCargaLote Then
        If MsgBox("Al usar esta opción, se limpiará el Grid, ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            LimpiarGrdCuenta
            txtArchivo.Text = ""
            bCargaLote = False
            txtMonto.Text = "0.00"
            grdCuenta.ColumnasAEditar = "X-1-X-X-X-5"
        Else
            Exit Sub
        End If
    End If
    grdCuenta.lbEditarFlex = True
    grdCuenta.Col = 1
    grdCuenta_RowColChange
    grdCuenta.AdicionaFila
    grdCuenta.SetFocus
    SendKeys "{Enter}"
End Sub

Private Sub cmdBusca_Click()
   On Error GoTo error_handler
    
    txtArchivo.Text = Empty
    
    dlgArchivo.InitDir = "C:\"
    dlgArchivo.Filter = "Archivos de Texto (*.txt)|*.txt|Archivos de Excel (*.xls)|*.xls| Archivos de Excel (*.xlsx)|*.xlsx"
    dlgArchivo.ShowOpen
    If dlgArchivo.Filename <> Empty Then
        txtArchivo.Text = dlgArchivo.Filename
        If fnProducto = gCapCTS Then
            txtArchivo.Locked = True
            cmdAgregar.Enabled = False
        End If
    Else
        txtArchivo.Text = "NO SE ABRIO NINGUN ARCHIVO"
        Exit Sub
    End If
    cmdCargar.Enabled = True
    
     Exit Sub
error_handler:
    
    If err.Number = 32755 Then
    ElseIf err.Number = 1004 Then
        MsgBox "Archivo en uso. Ciérrelo y luego proceda a reemplazar.", vbExclamation, "Aviso"
    Else
        MsgBox "Error al momento de seleccionar el archivo", vbCritical, "Aviso"
    End If
End Sub

Private Sub cmdCancelar_Click()
Limpiar
End Sub

Private Sub cmdCargar_Click()

    Dim lsNroDoc As String
    Dim lsPersCod As String
    Dim lsNombre As String
    Dim lsMonto As String
    Dim lnPersoneria As Integer
    Dim lnOP As Integer
    Dim lnTasaCli As Double
    Dim lnMonApeCli As Double
    Dim objExcel As Excel.Application
    Dim xLibro As Excel.Workbook
    Dim Col As Integer, Fila As Integer
    Dim psArchivoAGrabar As String
    Dim sCad As String
    Dim nFila As Long
    Dim lsNomArch As String
    Dim lsDire As String
    Dim lsTipDOI As String
    Dim X As Integer
    Dim Y As Integer, z As Integer
    Dim bMayorEdad As Boolean, bFormato As Boolean
    Dim psArchivoAGrabarMenores As String
    Dim psArchivoAGrabarPersJurid As String
    
    Dim oBookMenores As Object
    Dim oSheetMenores As Object
    
    Dim oBookPersJurid As Object
    Dim oSheetPersJurid As Object
    
    Dim sClientes As String
    Dim sMontos As String
    
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
        
    Dim oPer As New COMDPersona.DCOMPersonas
    Dim rsPers As New ADODB.Recordset
    Dim rsPersTmp As New ADODB.Recordset
        
    If txtArchivo.Text = "" Then
        MsgBox "No selecciono ningun archivo", vbExclamation, "Aviso"
        Set oPer = Nothing
        Set rsPers = Nothing
        Set rsPersTmp = Nothing
        Exit Sub
    End If
    If MsgBox("¿Esta operación puede tardar minutos, esta seguro de continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    If grdCuenta.Rows >= 2 And Len(Trim(grdCuenta.TextMatrix(1, 1))) > 0 Then
        If MsgBox("Al cargar la trama, se limpiaran los registros del Grid, ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        Else
            grdCuenta.Clear
            grdCuenta.Rows = 2
            grdCuenta.FormaCabecera
        End If
    End If
    grdCuenta.lbEditarFlex = False
    grdCuenta.ColumnasAEditar = "X-X-X-X-X-X"
    pbProgres.Max = 10
    pbProgres.Min = 1
    pbProgres.value = 1
    pbProgres.Visible = True
    DoEvents
    Set objExcel = New Excel.Application
    Set xLibro = objExcel.Workbooks.Open(txtArchivo.Text)
    psArchivoAGrabar = App.path & "\SPOOLER\NoCumpleValidacion_" & Format(gdFecSis, "yyyymmdd") & ".xls"
    
    grdCuenta.SetFocus
                
    cmdEliminar.Enabled = True
    X = 1
    Y = 1: z = 1
    
    If Dir(psArchivoAGrabar) <> "" Then
        Kill psArchivoAGrabar
    End If
    pbProgres.value = 2
    DoEvents
       Set oExcel = CreateObject("Excel.Application")
       Set oBook = oExcel.Workbooks.Add
       Set oSheet = oBook.Worksheets(1)
    pbProgres.value = 3
    DoEvents
        bFormato = True
        bCargaLote = True
        
        If UCase(Trim(xLibro.Sheets(1).Cells(1, 1))) <> "ITEM" Then bFormato = False
        If UCase(Trim(xLibro.Sheets(1).Cells(1, 2))) <> "TIPO DOC (1=DNI 2=RUC)" Then bFormato = False
        If UCase(Trim(xLibro.Sheets(1).Cells(1, 3))) <> "NRO DOC" Then bFormato = False
        If UCase(Trim(xLibro.Sheets(1).Cells(1, 4))) <> "MONTO" Then bFormato = False
        If bFormato = False Then
            MsgBox "El archivo seleccionado no tiene el formato adecuado para la carga en lote, verifíquelo e inténtelo de nuevo", vbInformation, "Aviso"
            If Not objExcel Is Nothing Then
                objExcel.Workbooks.Close
                Set objExcel = Nothing
            End If
            If Not oExcel Is Nothing Then
                oExcel.Workbooks.Close
                Set oExcel = Nothing
            End If
            pbProgres.Visible = False
            Exit Sub
        End If
        Fila = 2
        pbProgres.value = 4
        DoEvents
        With xLibro
            With .Sheets(1)
            Do While Len(Trim(.Cells(Fila, 2))) > 0
                lsNroDoc = Trim(.Cells(Fila, 3))
                lsMonto = Trim(.Cells(Fila, 4))
                sClientes = sClientes & Trim(lsNroDoc) & ","
                sMontos = sMontos & Trim(lsMonto) & ","
                Fila = Fila + 1
            Loop
            End With
        End With
        pbProgres.value = 7
        DoEvents
        If Len(sClientes) > 2 Then
            sClientes = Mid(sClientes, 1, Len(sClientes) - 1)
        End If
        If Len(sMontos) > 2 Then
            sMontos = Mid(sMontos, 1, Len(sMontos) - 1)
        End If
        If Len(sClientes) = 0 Or Len(sMontos) = 0 Then
            MsgBox "La trama seleccionada no contiene datos para la carga", vbInformation, "Aviso"
            If Not objExcel Is Nothing Then
                objExcel.Workbooks.Close
                Set objExcel = Nothing
            End If
            If Not oExcel Is Nothing Then
                oExcel.Workbooks.Close
                Set oExcel = Nothing
            End If
            pbProgres.Visible = False
            Exit Sub
        End If
        Set rsPers = oPer.ValidaTramaDeposito(sClientes, txtInstitucion.Text, Trim(Right(cboMoneda.Text, 5)))
        pbProgres.value = 8
        DoEvents
        If (MostrarErrores(rsPers)) Then
            grdCuenta.Clear
            grdCuenta.Rows = 2
            grdCuenta.FormaCabecera
        Else
            Set rsPers = Nothing
            Set rsPers = oPer.ListarTramaDeposito(sClientes, sMontos, txtInstitucion.Text, Trim(Right(cboMoneda.Text, 5)))
            If Not rsPers Is Nothing Then
                grdCuenta.rsFlex = rsPers
            End If
        End If
        pbProgres.value = 9
        DoEvents
        txtMonto.Text = Format(grdCuenta.SumaRow(5), "#,##0.00")
        pbProgres.value = 10
        DoEvents
        pbProgres.Visible = False
        objExcel.Quit
        Set objExcel = Nothing
        Set xLibro = Nothing
        Set oBook = Nothing
        oExcel.Quit
        Set oExcel = Nothing
End Sub

Private Sub cmdEliminar_Click()
    Dim nFila As Long
    nFila = grdCuenta.row
    If bCargaLote Then Exit Sub
    If MsgBox("¿Desea eliminar la fila seleccionada?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        grdCuenta.EliminaFila nFila
    End If
End Sub

Private Sub cmdFormato_Click()
    
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim nFila, i As Double
        
On Error GoTo error_handler
        
    dlgArchivo.Filename = Empty
    dlgArchivo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx| Archivos de Excel (*.xls)|*.xls"
    dlgArchivo.Filename = "DepositoHaberesLote" & Format(Now, "yyyyMMddhhnnss") & ".xlsx"
    dlgArchivo.ShowSave
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    If fs.FileExists(dlgArchivo.Filename) Then
        MsgBox "El archivo '" & dlgArchivo.FileTitle & "' ya existe, debe asignarle un nombre diferente", vbExclamation, ""
        Exit Sub
    End If
    If fnProducto = gCapAhorros Then
        lsArchivo = App.path & "\FormatoCarta\FormatoDepositoLoteAhorro.xlsx"
        lsNomHoja = "DepositoLoteAhorro"
    End If
    lsArchivo1 = dlgArchivo.Filename
    If fs.FileExists(lsArchivo) Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(lsArchivo)
    Else
        MsgBox "No Existe Plantilla en la Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If
    For Each xlHoja1 In xlsLibro.Worksheets
        If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            lbExisteHoja = True
            Exit For
        End If
    Next
    xlHoja1.SaveAs lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    MsgBox "Se exportó el formato para la carga de archivos", vbInformation, "Aviso"
               
Exit Sub
    
error_handler:
    If err.Number = 32755 Then
    ElseIf err.Number = 1004 Then
        MsgBox "Archivo en uso. Ciérrelo y luego proceda a reemplazar.", vbExclamation, "Aviso"
    Else
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Error al momento de generar el archivo", vbCritical, "Aviso"
    End If
End Sub

Private Function ValidacionDeposito() As String

    Dim oPer As New COMDPersona.DCOMPersonas
    Dim sClientes As String, sMensaje As String
    Dim rsPers As ADODB.Recordset
    Dim i As Integer
    
    ' registros en la grilla
    If grdCuenta.Rows = 2 And Len(Trim(grdCuenta.TextMatrix(1, 1))) = 0 Then
        sMensaje = "Debe al menos ingresar un registro en la grilla" & vbNewLine
    End If
    
    ' seleccion de institucion
    If Len(Trim(txtInstitucion.Text)) = 0 Then
        sMensaje = sMensaje & "Debe seleccionar una institucion" & vbNewLine
    End If
    
    ' texto en la glosa
    If Len(Trim(Replace(txtGlosa.Text, vbNewLine, ""))) = 0 Then
        sMensaje = sMensaje & "Debe ingresar un valor en la glosa" & vbNewLine
    End If
    
    ' registros del voucher
    If fnOpeCod = gAhoDepositoHaberesEnLoteTransf Then
        If nNroDeposito <> grdCuenta.Rows - 1 Then
            sMensaje = sMensaje & "El número de registros ingresados no coincide con el numero de depósitos del voucher" & vbNewLine
        End If
    End If
    ' Monto Depositado
    If CCur(lblMonTra.Caption) <> CCur(txtMonto.Text) And fnOpeCod = gAhoDepositoHaberesEnLoteTransf Then
        sMensaje = sMensaje & "El monto de depósito es diferente al monto del Vouvher" & vbNewLine
    End If
    
    ' Validando las cuentas en funcion a los DOI de los clientes
    For i = 1 To grdCuenta.Rows - 1
        sClientes = sClientes & Trim(grdCuenta.TextMatrix(i, 4)) & ","
    Next
    If Len(sClientes) > 2 Then
        sClientes = Mid(sClientes, 1, Len(sClientes) - 1)
    End If
    Set rsPers = oPer.ValidaTramaDeposito(sClientes, txtInstitucion.Text, Trim(Right(cboMoneda.Text, 5)))
    If Not rsPers Is Nothing Then
        If Not rsPers.EOF And Not rsPers.BOF Then
            sMensaje = sMensaje & "Revisar la titularidad del cliente, la moneda selcionada y su vinculacion con la institución" & vbNewLine
        End If
    End If
    ValidacionDeposito = sMensaje
    Exit Function
End Function


Private Sub cmdGrabar_Click()

Dim nMontoCargo As Double
Dim sCuenta As String, sGlosa As String
Dim lsmensaje As String
Dim lsBoleta As String
Dim lsBoletaITF As String
Dim nFicSal As Integer
Dim Autid As Long
Dim bResult As Boolean
Dim sCuentas() As String

On Error GoTo ErrGraba

lsmensaje = Trim(ValidacionDeposito)
If Len(lsmensaje) > 0 Then
    MsgBox "Se presentaron observaciones: " & vbNewLine & lsmensaje, vbInformation, "Aviso"
    Exit Sub
End If
If MsgBox("¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sMovNro As String, sMovNroV As String
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim rsCtaAbo As ADODB.Recordset
    
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Sleep (1000)
    sMovNroV = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set rsCtaAbo = grdCuenta.GetRsNew()
    sGlosa = Replace(Trim(txtGlosa.Text), vbNewLine, " ")
        
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios, sPersLavDinero As String
    Dim nMontoLavDinero As Double, nTC As Double, sReaPersLavDinero As String, sBenPersLavDinero As String
    
    Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
    Set clsExo = Nothing
    Set clsLav = Nothing
    
    nMontoCargo = CDbl(txtMonto.Text)
    
    If fnOpeCod = gAhoDepositoHaberesEnLoteEfec Then
        bResult = clsCap.CapAbonoLoteCtaSueldo(sCuenta, nMontoCargo, sMovNro, rsCtaAbo, sGlosa, gsNomAge, sLpt, sPersLavDinero, CDbl(gnTipCambioC), CDbl(gnTipCambioV), gbITFAplica, 0, gbITFAsumidoAho, 0, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, , , , , , , gnMovNro, fnOpeCod)
    ElseIf fnOpeCod = gAhoDepositoHaberesEnLoteTransf Then
        bResult = clsCap.CapAbonoLoteCtaSueldo(sCuenta, nMontoCargo, sMovNro, rsCtaAbo, sGlosa, gsNomAge, sLpt, sPersLavDinero, CDbl(gnTipCambioC), CDbl(gnTipCambioV), gbITFAplica, 0, gbITFAsumidoAho, 0, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, , , , , , , gnMovNro, fnOpeCod, , , , , , lnMovNroTransfer, fnMovNroRVD, sMovNroV, nNroDeposito)
    ElseIf fnOpeCod = gAhoDepositoEnLoteCheq Then
    End If
    If bResult Then
     If gnMovNro > 0 Then
         Call frmMovLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, sBenPersLavDinero, , , , , , gnTipoREU, gnMontoAcumulado, gsOrigen) 'JACA 20110224
     End If
     
      If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
      End If
      
      Do
        If Trim(lsBoleta) <> "" Then
           nFicSal = FreeFile
           Open sLpt For Output As nFicSal
              Print #nFicSal, lsBoleta & Chr$(12)
              Print #nFicSal, ""
           Close #nFicSal
        End If
      Loop Until MsgBox("¿Desea reimprimir Boleta de Depósito en lote? ", vbQuestion + vbYesNo, Me.Caption) = vbNo
      
      cmdCancelar_Click
      MsgBox "La operación se realizó correctamente", vbInformation, "Aviso"
    Else
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If
End If
 Set clsCap = Nothing
 gVarPublicas.LimpiaVarLavDinero
 
Exit Sub
ErrGraba:
    MsgBox err.Description, vbExclamation, "Error"
    Exit Sub

End Sub

Private Sub cmdsalir_Click()
    If MsgBox("¿Deseas salir de la formulario?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdTranfer_Click()
    Dim lsGlosa As String
    Dim lsDoc As String
    Dim lsInstit As String
    Dim oform As frmCapRegVouDepBus
    Dim lnTipMot As Integer
    Dim lsDetalle As String
    Dim lnTransferSaldo As Currency
    Dim fsPersCodTransfer As String
    If cboTransferMoneda.Text = "" Then
        MsgBox "Debe escoger la moneda de la transferencia.", vbInformation, "Aviso"
        If cboTransferMoneda.Visible And cboTransferMoneda.Enabled Then cboTransferMoneda.SetFocus
        Exit Sub
    End If
    If gsOpeCod = gAhoDepositoHaberesEnLoteTransf Then
        lnTipMot = 9
    End If
    fnMovNroRVD = 0
    Set oform = New frmCapRegVouDepBus
    SetDatosTransferencia "", "", "", 0, -1, "" 'Limpiamos datos y variables globales
    oform.iniciarFormulario Trim(Right(cboTransferMoneda, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer, lsDetalle
    If fnMovNroRVD = 0 Then
        Exit Sub
    End If
    If IsNumeric(Trim(lsDetalle)) Then
        nNroDeposito = CInt(lsDetalle)
    Else
        nNroDeposito = 0
    End If
    SetDatosTransferencia lsGlosa, lsInstit, lsDoc, lnTransferSaldo, lnMovNroTransfer, lsDetalle
    Me.grdCuenta.row = 1
    Set oform = Nothing
    Exit Sub
End Sub
Private Sub SetDatosTransferencia(ByVal psGlosa As String, ByVal psInstit As String, ByVal psDoc As String, ByVal pnTransferSaldo As Currency, ByVal pnMovNroTransfer As Long, ByVal psDetalle As String)
    
    Dim oPersona As New DCOMPersona
    Dim rsPersona As New ADODB.Recordset
    Dim row As Integer
    
    txtTransferGlosa.Text = psGlosa
    lbltransferBco.Caption = psInstit
    lblTrasferND.Caption = psDoc
    If psDetalle <> "" Then

    End If
    
    If pnMovNroTransfer <> -1 Then
        txtTransferGlosa.SetFocus
    End If
    
    txtTransferGlosa.Locked = True
    txtMonto.Enabled = False
    lblMonTra = Format(pnTransferSaldo, "#,##0.00")
    
    Set rsPersona = Nothing
    Set oPersona = Nothing
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
End Sub

Private Sub Form_Load()
Me.Caption = fsDescOperacion
Me.txtMonto.Enabled = False
bCargaLote = False
lblTTCCD.Caption = Format(gnTipCambioC, "#,#0.0000")
lblTTCVD.Caption = Format(gnTipCambioV, "#,#0.0000")
End Sub

Private Sub grdCuenta_OnCellChange(pnRow As Long, pnCol As Long)
    txtMonto.Text = Format(grdCuenta.SumaRow(5), "#,##0.00")
End Sub

Private Sub grdCuenta_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)

Dim sCta As String
Dim sRelac As String
Dim sEstado As String
Dim i As Long
Dim bDuplicadoCuenta As Boolean

If psDataCod = "" Then
    grdCuenta.EliminaFila pnRow
    Exit Sub
End If

Dim ClsPersona As New COMDPersona.DCOMPersonas
Dim clsCap As New COMNCaptaGenerales.NCOMCaptaGenerales
Dim clsPers As New COMDPersona.UCOMPersona
Dim rsCuenta As New ADODB.Recordset
Dim rsPersona As New ADODB.Recordset
Dim clsCuenta As New UCapCuenta

Set rsCuenta = clsCap.GetCuentasPersona(psDataCod, gCapAhorros, True, , Trim(Right(cboMoneda.Text, 5)), , , "6", True)
Set rsPersona = ClsPersona.BuscaCliente(psDataCod, BusquedaCodigo)

If Not rsCuenta Is Nothing Then
    If Not (rsCuenta.EOF And rsCuenta.EOF) Then
        Do While Not rsCuenta.EOF
            sCta = rsCuenta("cCtaCod")
            sRelac = rsCuenta("cRelacion")
            sEstado = Trim(rsCuenta("cEstado"))
            frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
            rsCuenta.MoveNext
        Loop
    Else
        grdCuenta.EliminaFila pnRow
        MsgBox "Persona no posee cuenta sueldo", vbInformation, "Aviso"
        rsCuenta.Close
        Set rsCuenta = Nothing
        Set clsPers = Nothing
        Set ClsPersona = Nothing
        Set clsCap = Nothing
        Exit Sub
    End If
End If
rsCuenta.Close
grdCuenta.TextMatrix(pnRow, 2) = ""
Set rsCuenta = Nothing
Set clsCuenta = New UCapCuenta
Set clsCuenta = frmCapMantenimientoCtas.inicia
If Not clsCuenta Is Nothing Then
    If clsCuenta.sCtaCod <> "" Then
        If pbEsDuplicado Then
            For i = 1 To grdCuenta.Rows - 1
                If clsCuenta.sCtaCod = grdCuenta.TextMatrix(i, 2) Then
                    MsgBox "El registro seleccionado es duplicado.", vbInformation, "Aviso"
                    grdCuenta.EliminaFila pnRow
                    Exit Sub
                End If
            Next
        End If
        grdCuenta.TextMatrix(grdCuenta.row, 1) = rsPersona!cPersCod
        grdCuenta.TextMatrix(grdCuenta.row, 2) = clsCuenta.sCtaCod
        grdCuenta.TextMatrix(grdCuenta.row, 3) = rsPersona!cPersNombre
        grdCuenta.TextMatrix(grdCuenta.row, 4) = rsPersona!cPersIDnroDNI
        grdCuenta.Col = 4
        SendKeys "{F2}"
    Else
        grdCuenta.EliminaFila pnRow
    End If
Else
    grdCuenta.EliminaFila pnRow
End If
Set clsCuenta = Nothing
            
End Sub

Private Sub grdCuenta_OnRowDelete()
    txtMonto.Text = Format$(grdCuenta.SumaRow(5), "#,##0.00")
End Sub

Private Sub grdCuenta_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(grdCuenta.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub

Private Sub grdCuenta_RowColChange()
    Dim nRow As Long
    Dim nCol As Long
    
    nRow = grdCuenta.row
    nCol = grdCuenta.Col
    If bCargaLote Then
        If nCol = 1 Then
            grdCuenta.lbEditarFlex = False
            Me.KeyPreview = True
        Else
            grdCuenta.lbEditarFlex = True
            Me.KeyPreview = True
        End If
    Else
        If nCol = 1 Then
            grdCuenta.lbEditarFlex = True
            Me.KeyPreview = False
        Else
            grdCuenta.lbEditarFlex = True
            Me.KeyPreview = True
        End If
    End If
    If Not IsNumeric(Trim(grdCuenta.TextMatrix(nRow, 1))) Then
        grdCuenta.TextMatrix(nRow, 1) = ""
    End If
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        If cmdGrabar.Enabled Then cmdGrabar.SetFocus
    End If
End Sub
Private Sub txtInstitucion_EmiteDatos()
    If txtInstitucion.Text <> "" Then
        If txtInstitucion.PersPersoneria < 2 Then
            lblInst.Text = ""
            txtInstitucion.Text = ""
            MsgBox "La institución seleccionada debe tener personería Jurídica", vbInformation, "Aviso"
        Else
            lblInst.Text = txtInstitucion.psDescripcion
            If cmdAgregar.Enabled Then cmdAgregar.SetFocus
        End If
    Else
        lblInst.Text = ""
    End If
End Sub
Public Sub iniciarFormulario(ByVal pnProducto As Producto, ByVal pnOpeCod As CaptacOperacion, _
                             Optional psDescOperacion As String = "")
fnProducto = pnProducto
fnOpeCod = pnOpeCod
fsDescOperacion = psDescOperacion

Dim clsGen As New COMDConstSistema.DCOMGeneral
Dim rsConst As New ADODB.Recordset
Set rsConst = clsGen.GetConstante(gMoneda)

Select Case fnProducto

    Case gCapAhorros
    
        If fnOpeCod = gAhoDepositoHaberesEnLoteEfec Then
            fraTranferecia.Visible = False
        ElseIf fnOpeCod = gAhoDepositoHaberesEnLoteTransf Then
            fraTranferecia.Visible = True
        End If

    Case gCapCTS
            
End Select

If gsOpeCod = gAhoDepositoHaberesEnLoteTransf Then
    Set rsConst = clsGen.GetConstante(gMoneda)
    CargaCombo cboTransferMoneda, rsConst
End If

Set rsConst = clsGen.GetConstante(gMoneda)
CargaCombo cboMoneda, rsConst
If cboMoneda.ListCount > 0 Then
    cboMoneda.ListIndex = 0
End If

Me.Show 1
End Sub

Private Function MostrarErrores(ByVal rsErrores As ADODB.Recordset) As Boolean

    Dim oBook As Object
    Dim oSheet As Object
    Dim sDireccion As String
    Dim oExcel As Object
    Dim bResult As Boolean
    Dim i As Long
    
    On Error GoTo error_handler
    
    bResult = False
    If Not rsErrores Is Nothing Then
        Set oExcel = CreateObject("Excel.Application")
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Worksheets(1)
        sDireccion = App.path & "\SPOOLER\Observaciones_" & Format(CDate(gdFecSis), "yyyyMMdd") & ".xls"
        If Dir(sDireccion) <> "" Then
            Kill sDireccion
        End If
        oSheet.Range("A1:F1").Font.Bold = True
        oSheet.Columns("B:B").NumberFormat = "@"
        oSheet.Columns("C:C").NumberFormat = "@"
        oSheet.Range("A1").value = "#"
        oSheet.Columns("A:A").ColumnWidth = 7
        oSheet.Columns("B:B").ColumnWidth = 21
        oSheet.Columns("C:C").ColumnWidth = 15
        oSheet.Columns("D:D").ColumnWidth = 51
        oSheet.Columns("F:F").ColumnWidth = 80
        oSheet.Range("B1").value = "NRO CUENTA"
        oSheet.Range("C1").value = "COD CLIENTE"
        oSheet.Range("D1").value = "NOMBRE"
        oSheet.Range("E1").value = "DOI"
        oSheet.Range("F1").value = "OBSERVACIONES"
        i = 2
        Do While Not rsErrores.EOF And Not rsErrores.BOF
            oSheet.Range("A" & i).value = i - 1
            oSheet.Range("B" & i).value = rsErrores!cCtaCod
            oSheet.Range("C" & i).value = rsErrores!cPersCod
            oSheet.Range("D" & i).value = rsErrores!cPersNombre
            oSheet.Range("E" & i).value = rsErrores!cPersDoi
            If rsErrores!nRegistrado = 0 Then
                oSheet.Range("F" & i).value = "Persona no registrada"
            ElseIf rsErrores!nTitular = 0 Then
                oSheet.Range("F" & i).value = "Cliente no es titular de la cuenta sueldo"
            ElseIf rsErrores!nMonedaSelec = 0 Then
                oSheet.Range("F" & i).value = "La cuenta sueldo no es de la moneda seleccionada"
            ElseIf rsErrores!nVinculacion = 0 Then
                oSheet.Range("F" & i).value = "La cuenta sueldo no está vinculada a la empresa seleccionada"
            ElseIf rsErrores!nCantCuentas > 1 Then
                oSheet.Range("F" & i).value = "Cliente tiene mas de una cuenta sueldo con la institucion y moneda seleccionados"
            Else
                oSheet.Range("F" & i).value = "La persona presenta observaciones"
            End If
            i = i + 1
            rsErrores.MoveNext
            bResult = True
        Loop
        oBook.SaveAs sDireccion
        oExcel.Quit
        Set oExcel = Nothing
        Set oBook = Nothing
        If i > 2 Then
            If MsgBox("¿Algunos registros de la trama no cumplen con las validaciones respectivas, deseas exportar el detalle a Excel?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                MostrarErrores = bResult
                Exit Function
            End If
            Dim m_Excel As New Excel.Application
            m_Excel.Workbooks.Open (sDireccion)
            m_Excel.Visible = True
        End If
    End If
    
    MostrarErrores = bResult
    Exit Function
    
error_handler:
        oExcel.Quit
        Set oExcel = Nothing
        Set oBook = Nothing
        MsgBox "Error al momento de generar el archivo", vbCritical, "Aviso"
        MostrarErrores = False
        
End Function

Private Sub LimpiarGrdCuenta()
    grdCuenta.Clear
    grdCuenta.Rows = 2
    grdCuenta.FormaCabecera
End Sub

Private Sub Limpiar()
LimpiarGrdCuenta
LimpiarTransferencia
LimpiarControles
End Sub

Private Sub LimpiarTransferencia()
    SetDatosTransferencia "", "", "", 0, -1, "" 'Limpiamos datos y variables globales
End Sub

Private Sub LimpiarControles()
    If cboMoneda.ListCount > 0 Then
        cboMoneda.ListIndex = 0
    End If
    txtArchivo.Text = ""
    txtMonto.Text = "0.00"
    txtInstitucion.Text = ""
    lblInst.Text = ""
    txtGlosa.Text = ""
    grdCuenta.ColumnasAEditar = "X-1-X-X-X-5"
    bCargaLote = False
    cboMoneda.SetFocus
End Sub
