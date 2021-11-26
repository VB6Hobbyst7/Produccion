VERSION 5.00
Begin VB.Form frmCapAbonos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9120
   ClientLeft      =   2655
   ClientTop       =   2175
   ClientWidth     =   8115
   Icon            =   "frmCapAbonos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDevolCtaRecaudo 
      Caption         =   "Devolución a Cuenta Recaudo"
      Height          =   375
      Left            =   3000
      TabIndex        =   68
      Top             =   8520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CheckBox chkIniciarEcotaxi 
      Caption         =   "Inicial Ecotaxi"
      Height          =   375
      Left            =   1320
      TabIndex        =   61
      Top             =   8520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame FrClienteWesterUnion 
      Caption         =   "Cliente Wester Union"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   75
      TabIndex        =   55
      Top             =   4875
      Width           =   7950
      Begin SICMACT.TxtBuscar txtClienteWesterUnion 
         Height          =   375
         Left            =   240
         TabIndex        =   56
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
         Enabled         =   0   'False
         Enabled         =   0   'False
         TipoBusqueda    =   3
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin VB.Label lblClienteWesterUnion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2520
         TabIndex        =   57
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   8400
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   8400
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   8520
      Width           =   1000
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3315
      Left            =   75
      TabIndex        =   15
      Top             =   1425
      Width           =   7950
      Begin VB.Frame fraDevolAfecGarantEcotaxi 
         Height          =   1215
         Left            =   3840
         TabIndex        =   62
         Top             =   2040
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CommandButton cmdBuscaCredGarant 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   320
            Left            =   3650
            TabIndex        =   67
            Top             =   260
            Width           =   255
         End
         Begin VB.TextBox txtClienteGarant 
            Enabled         =   0   'False
            Height          =   285
            Left            =   840
            TabIndex        =   66
            Top             =   690
            Width           =   2895
         End
         Begin SICMACT.ActXCodCta txtCuentaGarant 
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   661
            Texto           =   "Nº Cred."
         End
         Begin VB.CheckBox chkDevAfecGarantEcotaxi 
            Caption         =   "Devolución de Afectación garantía - Ecotaxi"
            DataMember      =   "chkGarant"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   0
            Width           =   3615
         End
         Begin VB.Label Label5 
            Caption         =   "Cliente :"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   735
         End
      End
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1755
         Left            =   105
         TabIndex        =   1
         Top             =   225
         Width           =   7680
         _ExtentX        =   13811
         _ExtentY        =   3096
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Nombre-Relacion-Direccion-ID"
         EncabezadosAnchos=   "250-1700-3800-1500-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   255
         RowHeight0      =   300
         TipoBusPersona  =   1
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Exoneración :"
         Height          =   195
         Left            =   150
         TabIndex        =   54
         Top             =   2805
         Width           =   1335
      End
      Begin VB.Label lblExoneracion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EXONERADA POR  ..."
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
         Height          =   315
         Left            =   1530
         TabIndex        =   53
         Top             =   2745
         Visible         =   0   'False
         Width           =   6030
      End
      Begin VB.Label lblFirmas 
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   3360
         TabIndex        =   25
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "# Firmas :"
         Height          =   195
         Left            =   2640
         TabIndex        =   24
         Top             =   2145
         Width           =   690
      End
      Begin VB.Label lblTipoCuenta 
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   1155
         TabIndex        =   23
         Top             =   2085
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   2138
         Width           =   960
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1350
      Left            =   75
      TabIndex        =   14
      Top             =   75
      Width           =   7920
      Begin VB.Frame fraDatos 
         Height          =   585
         Left            =   105
         TabIndex        =   16
         Top             =   660
         Width           =   7680
         Begin VB.Label lblUltContacto 
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   4215
            TabIndex        =   21
            Top             =   195
            Width           =   1995
         End
         Begin VB.Label lblEtqUltCnt 
            AutoSize        =   -1  'True
            Caption         =   "Ultimo Contacto :"
            Height          =   195
            Left            =   2910
            TabIndex        =   20
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label lblApertura 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   900
            TabIndex        =   18
            Top             =   195
            Width           =   1965
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   135
            TabIndex        =   17
            Top             =   255
            Width           =   690
         End
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   3840
         TabIndex        =   19
         Top             =   150
         Width           =   3960
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
      ForeColor       =   &H00000080&
      Height          =   2625
      Left            =   4635
      TabIndex        =   9
      Top             =   5640
      Width           =   3375
      Begin VB.CheckBox chkITFEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   870
         TabIndex        =   52
         Top             =   1860
         Width           =   705
      End
      Begin VB.Frame fraPeriodoCTS 
         Height          =   1080
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   3120
         Begin VB.ComboBox cboPeriodo 
            Height          =   315
            Left            =   885
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   210
            Width           =   2130
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Período :"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lblCTS 
            AutoSize        =   -1  'True
            Caption         =   "Dispon.del Excedente (%) :"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   675
            Width           =   1905
         End
         Begin VB.Label lblDispCTS 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H8000000D&
            Height          =   300
            Left            =   2160
            TabIndex        =   27
            Top             =   645
            Width           =   795
         End
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   375
         Left            =   870
         TabIndex        =   5
         Top             =   1380
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         ForeColor       =   12582912
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "ITF :"
         Height          =   195
         Left            =   225
         TabIndex        =   51
         Top             =   1860
         Width           =   330
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         Height          =   195
         Left            =   225
         TabIndex        =   50
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label lblITF 
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1680
         TabIndex        =   49
         Top             =   1815
         Width           =   1095
      End
      Begin VB.Label lblTotal 
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
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   870
         TabIndex        =   48
         Top             =   2190
         Width           =   1905
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   165
         TabIndex        =   11
         Top             =   1470
         Width           =   540
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "S/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2925
         TabIndex        =   10
         Top             =   1440
         Width           =   255
      End
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
      ForeColor       =   &H00000080&
      Height          =   2610
      Left            =   120
      TabIndex        =   34
      Top             =   5640
      Width           =   4500
      Begin VB.ComboBox cboTransferMoneda 
         Height          =   315
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   255
         Width           =   1575
      End
      Begin VB.CommandButton cmdTranfer 
         Height          =   350
         Left            =   2520
         Picture         =   "frmCapAbonos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   645
         Width           =   475
      End
      Begin VB.TextBox txtTransferGlosa 
         Appearance      =   0  'Flat
         Height          =   720
         Left            =   855
         MaxLength       =   255
         TabIndex        =   35
         Top             =   1410
         Width           =   3465
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
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   2760
         TabIndex        =   60
         Top             =   2160
         Visible         =   0   'False
         Width           =   1665
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
         Left            =   2400
         TabIndex        =   59
         Top             =   2190
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblEtiMonTra 
         AutoSize        =   -1  'True
         Caption         =   "Monto Transacción"
         Height          =   195
         Left            =   960
         TabIndex        =   58
         Top             =   2220
         Visible         =   0   'False
         Width           =   1380
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
         TabIndex        =   38
         Top             =   615
         Width           =   750
      End
      Begin VB.Label lblTTCCD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3585
         TabIndex        =   39
         Top             =   255
         Width           =   750
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
         Height          =   345
         Left            =   855
         TabIndex        =   47
         Top             =   1020
         Width           =   3465
      End
      Begin VB.Label lbltransferN 
         AutoSize        =   -1  'True
         Caption         =   "Nro Doc :"
         Height          =   195
         Left            =   45
         TabIndex        =   46
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lbltransferBcol 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   90
         TabIndex        =   45
         Top             =   1110
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
         Height          =   345
         Left            =   855
         TabIndex        =   44
         Top             =   645
         Width           =   1575
      End
      Begin VB.Label lblTransferMoneda 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   45
         TabIndex        =   43
         Top             =   315
         Width           =   585
      End
      Begin VB.Label lblTransferGlosa 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   1410
         Width           =   495
      End
      Begin VB.Label lblTTCC 
         Caption         =   "TCC"
         Height          =   285
         Left            =   3210
         TabIndex        =   41
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label4 
         Caption         =   "TCV"
         Height          =   285
         Left            =   3195
         TabIndex        =   40
         Top             =   630
         Width           =   390
      End
   End
   Begin VB.Frame fraDocumento 
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2610
      Left            =   105
      TabIndex        =   12
      Top             =   5640
      Width           =   4500
      Begin VB.TextBox txtglosa 
         Height          =   1395
         Left            =   705
         TabIndex        =   3
         Top             =   1095
         Width           =   3705
      End
      Begin VB.CommandButton cmdDocumento 
         Height          =   350
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   300
         Width           =   475
      End
      Begin VB.Label lblNroDoc 
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
         Height          =   345
         Left            =   915
         TabIndex        =   33
         Top             =   303
         Width           =   1575
      End
      Begin VB.Label lblNombreIF 
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
         Height          =   345
         Left            =   915
         TabIndex        =   32
         Top             =   690
         Width           =   3375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   765
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nro Doc :"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   378
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCapAbonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nProducto As COMDConstantes.Producto
Dim nTipoCuenta As COMDConstantes.ProductoCuentaTipo
Dim nMoneda As COMDConstantes.Moneda
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim bDocumento As Boolean
Public dFechaValorizacion As Date
Public lnDValoriza As Integer
Dim sPersCodCMAC As String, sTipoCuenta As String
Dim sNombreCMAC As String
Public sCodIF As String
Dim nPersoneria As COMDConstantes.PersPersoneria
Dim pbOrdPag As Boolean
Dim sOperacion As String

'Transferencia
Dim lnMovNroTransfer As Long
Dim lnTransferSaldo As Currency
Dim fsPersCodTransfer As String '***Agregado por ELRO el 20120706, según OYP-RFC074-2012
Dim fsOpeCod As String '***Agregado por ELRO el 20120706, según OYP-RFC074-2012
Dim fnMovNroRVD As Long '***Agregado por ELRO el 20120706, según OYP-RFC074-2012
Dim fsPersNombreCVME As String '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
Dim fsPersDireccionCVME As String '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
Dim fsdocumentoCVME As String '***Agregado por ELRO el 20120717, según OYP-RFC024-2012



'***************Variabres Agregadas********************
Dim Gtitular As String
Dim GAutNivel As String
Dim GAutMontoFinSol As Double
Dim GAutMontoFinDol As Double
Dim GMontoAprobado As Double, GNroID As Long, GPersCod As String

'Variables para la impresion de la boleta de Lavado de Dinero


Dim sPersCod As String, sDocId As String, sDireccion As String
Dim sPersCodRea As String, sNombreRea As String, sDocIdRea As String, sDireccionRea As String
Dim sNombre As String


'Variables para el ITF
Dim lbITFCtaExonerada As Boolean

'Variables para Validar Deposito Pandero/Panderito -- AVMM -- 20/02/2007
Dim lnMonDepositoP As Double
Dim lnTpoPrograma As Integer
Dim pnValorChq As Double
'ALPA 20081006**************************************************************
Dim lsCuentaWestern As String
Dim bCuentaWestern As Boolean
'***************************************************************************
Dim nRedondeoITF As Double 'BRGO 20110914

Dim sNumTarj As String
Dim fnDepositoPersRealiza As Boolean 'WIOR 20121114
Dim fnCondicion As Integer 'WIOR 20121114
Dim sPersCodEcotaxi As String 'ALPA 20130107
Dim sPersNomEcotaxi As String 'ALPA 20130107
Dim sPersDNIEcotaxi As String 'ALPA 20130107
Dim sPerEcotaxi() As String
Dim bValEfeTrans As Boolean 'RECO 20131104 ERS-141
Dim sCtaCodAbono As String 'RECO 20131104 ERS-141
Dim bDevolRecaudoEcotaxi As Boolean 'RECO 20131104 ERS-141
Dim oDocRec As UDocRec 'EJVG20140408
Dim bInstFinanc As Boolean 'JUEZ 20140414
'JUEZ 20141014 Nuevos Parámetros *****
Dim bValidaCantDep As Boolean
Dim nParCantDepLib As Integer
Dim nParMontoMinDepSol As Double
Dim nParMontoMinDepDol As Double
Dim nParCantDepAnio As Integer
Dim nParDiasVerifRegSueldo As Integer
Dim nParUltRemunBrutas As Integer
Dim loVistoElectronico As frmVistoElectronico
Dim lbVistoVal As Boolean
'END JUEZ ****************************

'RIRO20170414 ***
Dim sIFTipo As String
'END RIRO *******

Public Function ObtieneDatosTarjeta(ByVal psCodTarj As String) As Boolean
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsTarj As New ADODB.Recordset
Dim sTarjeta As String, sPersona As String
Dim nEstado As COMDConstantes.CaptacTarjetaEstado

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsTarj = clsMant.GetTarjetaCuentas(psCodTarj)
If rsTarj.EOF And rsTarj.BOF Then
    MsgBox "Tarjeta no posee ninguna relación con cuentas activas o Tarjeta no activa.", vbInformation, "Aviso"
    ObtieneDatosTarjeta = False
Else
    ObtieneDatosTarjeta = True
End If
Set rsTarj = Nothing
Set clsMant = Nothing
End Function
Sub Finaliza_Verifone5000()
        If Not GmyPSerial Is Nothing Then
            GmyPSerial.Disconnect
            Set GmyPSerial = Nothing
        End If
End Sub

'Funcion de Impresion de Boletas
Private Sub ImprimeBoleta(ByVal sBoleta As String, Optional ByVal sMensaje As String = "Boleta Operación")
Dim nFicSal As Integer
Do
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    If nProducto = gCapCTS Then
        sBoleta = sBoleta & oImpresora.gPrnSaltoLinea
    End If
    Print #nFicSal, sBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    Print #nFicSal, ""
    Print #nFicSal, ""
    Close #nFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
Dim clsGen As COMDConstSistema.DCOMGeneral 'DGeneral
Dim rsCta As New ADODB.Recordset, rsRel As New ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim nRow As Long
Dim sMsg As String, sMoneda As String, sPersona As String

'******RECO 20131029 ERS141**************
Dim oCapGen As New COMNCaptaGenerales.NCOMCaptaGenerales
Set oCapGen = New COMNCaptaGenerales.NCOMCaptaGenerales
'**************END RECO******************
'----- MADM
Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset
'----- MADM
'JUEZ 20141014 ***************************************
Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim rsPar As ADODB.Recordset
Dim nCantOpeCta As Integer

lbVistoVal = False
'END JUEZ ********************************************

Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
sMsg = clsCap.ValidaCuentaOperacion(sCuenta, True)
Set clsCap = Nothing
grdCliente.lbEditarFlex = False
If sMsg = "" Then
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsCta = New Recordset
        Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    Set clsMant = Nothing
    If Not (rsCta.EOF And rsCta.BOF) Then
    
        Dim dLSCAP As COMDCaptaGenerales.DCOMCaptaGenerales 'DCapMantenimiento
        Set dLSCAP = New COMDCaptaGenerales.DCOMCaptaGenerales

        If dLSCAP.EsCtaConvenio(sCuenta) Then
            MsgBox "Esta es una cuenta de Convenio." & vbCrLf & "Usar operación Abonos de Ctas de Convenio. ", vbOKOnly + vbInformation, "AVISO"
            Set dLSCAP = Nothing
            Exit Sub
        End If
    

        nEstado = rsCta("nPrdEstado")
        nPersoneria = rsCta("nPersoneria")
        
        'If nProducto = gCapAhorros Then
        If nProducto = gCapAhorros Or nProducto = gCapCTS Then 'JUEZ 20140320
            lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
        End If
        'JUEZ 20141014 ******************************************************
        Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        'Set rsPar = clsDef.GetCapParametroNew(nProducto, lnTpoPrograma)
        Set rsPar = clsDef.GetCapParametroNew(nProducto, lnTpoPrograma, sCuenta) 'APRI20190109 ERS077-2018
        
        Dim nIdTarifario As Integer
        nIdTarifario = rsPar!nIdTarifario
             
        'If nIdTarifario = 1 And EsHaberes(sCuenta) Then
         If nIdTarifario = 1 And EsHaberes(sCuenta) And (Trim(nOperacion) <> "200243" And Trim(nOperacion) <> "200244" And Trim(nOperacion) <> "200245") Then 'APRI20200128 MEJORA
                MsgBox "No puede utilizar esta Operación para una Cuenta de Haberes", vbOKOnly + vbExclamation, App.Title
                Exit Sub
        End If
                
        If nProducto = gCapAhorros Then
            nParCantDepLib = rsPar!nCantOpeVentDep
            nParMontoMinDepSol = rsPar!nMontoMinDepSol
            nParMontoMinDepDol = rsPar!nMontoMinDepDol
        Else
            nParCantDepAnio = rsPar!nCantOpeDepAnio
            nParDiasVerifRegSueldo = rsPar!nDiasVerifUltRegSueldo
            nParUltRemunBrutas = rsPar!nUltRemunBrutas
        End If
        Set rsPar = Nothing

        If bValidaCantDep Then
            Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                nCantOpeCta = clsMant.ObtenerCantidadOperaciones(sCuenta, gCapMovDeposito, gdFecSis)
            Set clsMant = Nothing
'COMENTADO POR APRI20190109 ERS077-2018
'            If nCantOpeCta >= IIf(nProducto = gCapAhorros, nParCantDepLib, nParCantDepAnio) Then
'                If MsgBox("Se ha realizado el número máximo de depósitos, se requiere de VB del supervisor para grabar la operación. Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'                Set loVistoElectronico = New frmVistoElectronico
'
'                lbVistoVal = loVistoElectronico.Inicio(3, nOperacion)
'
'                If Not lbVistoVal Then Exit Sub
'            End If
        End If
        'END JUEZ ***********************************************************
        lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy hh:mm:ss")
        'ITF INICIO
        lbITFCtaExonerada = fgITFVerificaExoneracion(sCuenta)
        fgITFParamAsume Mid(sCuenta, 4, 2), Mid(sCuenta, 6, 3)
        'APRI20190109 ERS077-2018
       'If Trim(nOperacion) <> "200243" Or Trim(nOperacion) <> "200244" Or Trim(nOperacion) <> "200245" Then
        If nIdTarifario <> 1 And lnTpoPrograma = 6 And (Trim(nOperacion) <> "200243" And Trim(nOperacion) <> "200244" And Trim(nOperacion) <> "200245") Then
            lbITFCtaExonerada = False
        End If
        'END APRI
        If sPersCodCMAC = "" Then
            If gbITFAsumidoAho Then
                Me.chkITFEfectivo.Visible = False
                Me.chkITFEfectivo.value = 0
            Else
                If nOperacion = gAhoDepChq Then
                    Me.chkITFEfectivo.value = 0
                Else
                    Me.chkITFEfectivo.value = 0
                End If
                Me.chkITFEfectivo.Visible = True
            End If
        Else
            If gbITFAsumidoAho Then
                Me.chkITFEfectivo.Visible = False
                Me.chkITFEfectivo.value = 0
            Else
                Me.chkITFEfectivo.value = 0
                Me.chkITFEfectivo.Visible = True
            End If
            chkITFEfectivo.Enabled = False
        End If
        'ITF FIN
        
        nMoneda = CLng(Mid(sCuenta, 9, 1))
        
        If nMoneda = gMonedaNacional Then
            sMoneda = "MONEDA NACIONAL"
            txtMonto.BackColor = &HC0FFFF
            'lblMon.Caption = "S/."
            lblMon.Caption = gcPEN_SIMBOLO 'APRI20191022 SUGERENCIA CALIDAD
            lblITF.BackColor = &HC0FFFF
            lblTotal.BackColor = &HC0FFFF
        Else
            sMoneda = "MONEDA EXTRANJERA"
            txtMonto.BackColor = &HC0FFC0
            lblITF.BackColor = &HC0FFC0
            lblTotal.BackColor = &HC0FFC0
            lblMon.Caption = "$"
        End If
        
        
        
        If Mid(sCuenta, 4, 2) <> gsCodAge Then
            lblMensaje = "OPERACION REMOTA " & Trim(UCase(rsCta("cAgeDescripcion")))
        Else
            lblMensaje = "OPERACION LOCAL"
        End If
        
    
        Select Case nProducto
             Case gCapAhorros
             
             '
                If rsCta("bOrdPag") Then
                    lblMensaje = lblMensaje & Chr$(13) & "AHORROS CON ORDEN DE PAGO" & Chr$(13) & sMoneda
                    pbOrdPag = True
                Else
                    'AVMM 10-04-2007
                    If lnTpoPrograma = 1 Then
                        lblMensaje = lblMensaje & Chr$(13) & "AHORRO ÑAÑITO" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 2 Then
                        lblMensaje = lblMensaje & Chr$(13) & "AHORROS PANDERITO" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 3 Then
                        '*** PEAC 20090722
                        'lblMensaje = lblMensaje & Chr$(13) & "AHORROS PANDERO" & Chr$(13) & sMoneda
                        lblMensaje = lblMensaje & Chr$(13) & "AHORROS POCO A POCO AHORRO" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 4 Then
                        lblMensaje = lblMensaje & Chr$(13) & "AHORROS DESTINO" & Chr$(13) & sMoneda
                    Else
                        lblMensaje = lblMensaje & Chr$(13) & "AHORROS SIN ORDEN DE PAGO" & Chr$(13) & sMoneda
                    End If
                    pbOrdPag = False
                End If
                lblUltContacto = Format$(rsCta("dUltContacto"), "dd mmm yyyy hh:mm:ss")
                
                If fgITFTipoExoneracion(sCuenta) = gnITFTpoExoPlanilla And (nOperacion <> gAhoDepPlanRRHHAdelantoSueldos And nOperacion <> gAhoDepPlanRRHH And nOperacion <> gAhoDepOtrosIngRRHH And nOperacion <> "200243" And nOperacion <> "200244" And nOperacion <> "200245") Then
                    MsgBox "No puede hacer este tipo de abonos a cuentas exoneradas de ITF por Referencia a Remuneraciones.", vbInformation, "Aviso"
                    LimpiaControles
                    Exit Sub
                End If
                
                
                If lbITFCtaExonerada Then
                    Dim nTipoExo As Integer, sDescripcion As String
                    nTipoExo = fgITFTipoExoneracion(sCuenta, sDescripcion)
                    lblExoneracion.Visible = True
                    lblExoneracion.Caption = sDescripcion
                End If
                Set clsGen = Nothing
            Case gCapCTS
                'JUEZ 20140320 ********************************
                If lnTpoPrograma = 2 Then
                    MsgBox "Esta cuenta es un CTS No Activo, por lo cual no se puede realizar ningún depósito", vbInformation, "Aviso"
                    cmdCancelar_Click
                    Exit Sub
                End If
                'END JUEZ *************************************
                lblUltContacto = rsCta("cInstitucion")
                lblMensaje = lblMensaje & Chr$(13) & "CTS" & Chr$(13) & sMoneda
        End Select
        
        lblTipoCuenta = UCase(rsCta("cTipoCuenta"))
        sTipoCuenta = lblTipoCuenta
        nTipoCuenta = rsCta("nPrdCtaTpo")
        lblFirmas = Format$(rsCta("nFirmas"), "#0")
        Set rsRel = New ADODB.Recordset
        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
            Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        Set clsMant = Nothing
        sPersona = ""
        
        Do While Not rsRel.EOF
            If rsRel("cPersCod") = gsCodPersUser Then
                MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                Unload Me
                Exit Sub
            End If
            If sPersona <> rsRel("cPersCod") Then
                grdCliente.AdicionaFila
                nRow = grdCliente.Rows - 1
                grdCliente.TextMatrix(nRow, 1) = rsRel("cPersCod")
                grdCliente.TextMatrix(nRow, 2) = UCase(PstaNombre(rsRel("Nombre")))
                grdCliente.TextMatrix(nRow, 3) = UCase(rsRel("Relacion")) & space(50) & Trim(rsRel("nPrdPersRelac"))
                grdCliente.TextMatrix(nRow, 4) = rsRel("Direccion") & ""
                grdCliente.TextMatrix(nRow, 5) = rsRel("ID N°")
                sPersona = rsRel("cPersCod")
            End If
            rsRel.MoveNext
        Loop
        
        'JUEZ 20140414 ****************************************
        If nOperacion = gAhoDepEfec Or nOperacion = gAhoDepChq Then
            Dim i As Integer
            For i = 1 To grdCliente.Rows - 1
                If Trim(Left(grdCliente.TextMatrix(i, 3), 10)) = "TITULAR" Then
                    Dim oDInstFinan As COMDPersona.DCOMInstFinac
                    Set oDInstFinan = New COMDPersona.DCOMInstFinac
                    bInstFinanc = oDInstFinan.VerificaEsInstFinanc(Trim(grdCliente.TextMatrix(i, 1)))
                    Set oDInstFinan = Nothing
                    txtMonto_Change
                End If
            Next
        End If
        'END JUEZ *********************************************
        
        '********* firma madm
         Set lafirma = New frmPersonaFirma
         Set ClsPersona = New COMDPersona.DCOMPersonas
        
         Set Rf = ClsPersona.BuscaCliente(grdCliente.TextMatrix(nRow, 1), BusquedaCodigo)
         
         If Not Rf.BOF And Not Rf.EOF Then
            If Rf!nPersPersoneria = 1 Then
            Call frmPersonaFirma.Inicio(Trim(grdCliente.TextMatrix(nRow, 1)), Mid(grdCliente.TextMatrix(nRow, 1), 4, 2), False, True)
            End If
         End If
         Set Rf = Nothing
        '************ firma madm
        
        rsRel.Close
        Set rsRel = Nothing
        fraCliente.Enabled = True
        fraDocumento.Enabled = True
        fraMonto.Enabled = True
        FraCuenta.Enabled = False
        If fraDocumento.Visible Then
            If cmdDocumento.Visible Then
                cmdDocumento.Enabled = True
            Else
                txtGlosa.SetFocus
            End If
        ElseIf fraTranferecia.Visible Then
            If cboTransferMoneda.Visible Then
                fraTranferecia.Enabled = True
                '***Modificado por ELRO el 20121015, según OYP-RFC024-2012
                'cboTransferMoneda.Enabled = True
                If nMoneda = gMonedaNacional Then
                    cboTransferMoneda.ListIndex = 0
                Else
                    cboTransferMoneda.ListIndex = 1
                End If
                cboTransferMoneda.Enabled = False
                'cboTransferMoneda.SetFocus
                '***Fin Modificado por ELRO el 20121015*******************
               
            End If
        End If
        '***********RECO 20131024 ERS141*************
        'If txtCuenta.Prod = "517" Then
        
            If oCapGen.RecuperaSubTipoProducto(txtCuenta.NroCuenta)!nTpoPrograma = 7 Then
                chkDevolCtaRecaudo.Visible = True
                chkIniciarEcotaxi.Visible = True
            Else
                chkDevolCtaRecaudo.Visible = False
                chkIniciarEcotaxi.Visible = False
            End If
        If bValEfeTrans = True Then
            If oCapGen.ValidaFondoGarantEcotaxi(sCuenta) = True Then
                fraDevolAfecGarantEcotaxi.Visible = True
                chkDevolCtaRecaudo.Visible = False
            End If
        End If
        '*****************END RECO*******************
        cmdgrabar.Enabled = True
        cmdcancelar.Enabled = True
    End If
Else
    MsgBox sMsg, vbInformation, "Operacion"
    txtCuenta.SetFocus
End If
End Sub

Private Sub LimpiaControles()
grdCliente.Clear
lbITFCtaExonerada = False
grdCliente.Rows = 2
grdCliente.FormaCabecera
txtGlosa = ""
txtMonto.value = 0
FraCuenta.Enabled = True
Select Case nProducto
    Case gCapCTS
        cboPeriodo.ListIndex = 0
End Select
lblNroDoc = ""
lblNombreIF = ""
cmdgrabar.Enabled = False
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Age = ""
txtCuenta.Cuenta = ""
cmdgrabar.Enabled = False
cmdcancelar.Enabled = False
lblApertura = ""
lblUltContacto = ""
lblFirmas = ""
lblTipoCuenta = ""
fraCliente.Enabled = False
fraDatos.Enabled = False
fraDocumento = False
fraMonto.Enabled = False
fraTranferecia.Enabled = False
txtCuenta.SetFocus
'***Modificado por ELRO el 20121015, según OYP-RFC024-2012
'Me.cboTransferMoneda.Enabled = True
Me.cboTransferMoneda.Enabled = False
'***Fin Modificado por ELRO el 20121015*******************
Me.txtTransferGlosa.Text = ""
Me.lbltransferBco.Caption = ""
Me.lblTrasferND.Caption = ""
lnTransferSaldo = 0
lnMovNroTransfer = -1
Me.lblExoneracion.Visible = False
nRedondeoITF = 0
chkIniciarEcotaxi.value = 0 'ALPA 20130110
chkITFEfectivo.value = 0 'EJVG20130914
fraDevolAfecGarantEcotaxi.Visible = False '***RECO 20131024 ERS141***
txtCuentaGarant.NroCuenta = "" '***RECO 20131024 ERS141***
txtClienteGarant.Text = "" '***RECO 20131024 ERS141***
txtClienteGarant.Enabled = False '***RECO 20131024 ERS141***
txtCuentaGarant.Enabled = False '***RECO 20131024 ERS141***
chkIniciarEcotaxi.Visible = False '***RECO 20131024 ERS141***
sCtaCodAbono = "" '***RECO 20131024 ERS141***
chkDevAfecGarantEcotaxi.value = 0 '***RECO 20131024 ERS141***
bDevolRecaudoEcotaxi = False '***RECO 20131024 ERS141***
chkDevolCtaRecaudo.value = 0 '***RECO 20131024 ERS141***
chkDevolCtaRecaudo.Visible = False '***RECO 20131024 ERS141***
bInstFinanc = False 'JUEZ 20140414
sIFTipo = "" 'RIRO20170714
End Sub

Private Sub IniciaComboCTSPeriodo()
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsConst As New ADODB.Recordset
Dim sCodigo As String * 2
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsConst = clsMant.GetCTSPeriodo()
Set clsMant = Nothing
Do While Not rsConst.EOF
    sCodigo = rsConst("nItem")
    cboPeriodo.AddItem sCodigo & space(2) & UCase(rsConst("cDescripcion")) & space(100) & rsConst("nPorcentaje")
    rsConst.MoveNext
Loop
cboPeriodo.ListIndex = 4
End Sub

Public Sub Inicia(ByVal nProd As Producto, ByVal nOpe As CaptacOperacion, _
        Optional sCodCmac As String = "", Optional sNomCmac As String = "", _
        Optional sDescOperacion As String = "")
nProducto = nProd
nOperacion = nOpe
sPersCodCMAC = sCodCmac
sNombreCMAC = sNomCmac
sOperacion = sDescOperacion


'/*Verificar cantidad de operaciones disponibles ANDE 20171218*/
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim bProsigue As Boolean
    Dim cMsgValid As String
    bProsigue = oCaptaLN.OperacionPermitida(gsCodUser, gdFecSis, nOperacion, cMsgValid)
    If bProsigue = False Then
        MsgBox cMsgValid, vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
'/*end ande*/


Select Case nProd
    Case gCapAhorros
        
        fraPeriodoCTS.Visible = False
        lblEtqUltCnt = "Ult. Contacto :"
        lblUltContacto.Width = 2000
        txtCuenta.Prod = Trim(Str(gCapAhorros))
        If nOpe = "200243" Or nOpe = "200244" Or nOpe = "200245" Then
            Me.chkITFEfectivo.value = vbUnchecked
            Me.chkITFEfectivo.Enabled = False
            
        End If
            
        If sPersCodCMAC = "" Then
            Me.Caption = "Captaciones - Ahorros - " & sDescOperacion
        Else
            Me.Caption = "Captaciones - Ahorros - " & sDescOperacion & " - " & sNombreCMAC
        End If
        '******************RECO 20131104 ERS-141******************************************
        If nOperacion = gAhoDepEfec Or nOperacion = gAhoDepTransf Then
            bValEfeTrans = True
        Else
            bValEfeTrans = False
        End If
        '*******************************END RECO******************************************
    Case gCapCTS
        'By capi 05032009 Acta 025-2009
        FrClienteWesterUnion.Visible = False
        '
        fraPeriodoCTS.Visible = True
        IniciaComboCTSPeriodo
        lblEtqUltCnt = "Institución :"
        lblUltContacto.Width = 4500
        lblUltContacto.Left = lblUltContacto.Left - 450
        txtCuenta.Prod = Trim(Str(gCapCTS))
        If sPersCodCMAC = "" Then
            Me.Caption = "Captaciones - CTS - " & sDescOperacion
        Else
            Me.Caption = "Captaciones - CTS - " & sDescOperacion & " - " & sNombreCMAC
        End If
        Me.chkITFEfectivo.Visible = False
        Me.chkITFEfectivo.value = 0
End Select
'Verifica si la operacion necesita algun documento
If nOperacion = gAhoDepChq Or nOperacion = gCTSDepChq Or nOperacion = "200244" Or nOperacion = "200252" Then
    lblNroDoc.Visible = True
    lblNombreIF.Visible = True
    cmdDocumento.Visible = True
    fraDocumento.Caption = "Cheque"
    bDocumento = True
    txtMonto.Enabled = False
    Label12.Visible = True
    Label13.Visible = True
    Me.fraTranferecia.Visible = False
    chkITFEfectivo.Enabled = True
    Me.chkITFEfectivo.value = vbUnchecked
    'chkITFEfectivo.Enabled = False
    'Me.chkITFEfectivo.value = 1
ElseIf nOperacion = gAhoDepTransf Or nOperacion = gCTSDepTransf Or nOperacion = "200245" Then
    lblNroDoc.Visible = False
    lblNombreIF.Visible = False
    cmdDocumento.Visible = False
    bDocumento = False
    Label12.Visible = False
    Label13.Visible = False
    
    IniciaCombo cboTransferMoneda, gMoneda
    Me.fraDocumento.Visible = False
    Me.fraTranferecia.Visible = True

    chkITFEfectivo.Visible = False
    Me.chkITFEfectivo.value = 1
    
    '***Agregado por ELRO el 20120706, según OYP-RFC024-2012
    If nOperacion = gAhoDepTransf Or nOperacion = gCTSDepTransf Or nOperacion = "200245" Then
        lblEtiMonTra.Visible = True
        lblSimTra.Visible = True
        lblMonTra.Visible = True
        chkITFEfectivo.Visible = True
        'chkITFEfectivo.Enabled = False
        chkITFEfectivo.Enabled = True 'EJVG20130914
        chkITFEfectivo.value = 0
        cboTransferMoneda.Enabled = False
    End If
    '***Fin Agregado por ELRO*******************************
    
Else
    Label3.Visible = False
    lblNroDoc.Visible = False
    lblNombreIF.Visible = False
    cmdDocumento.Visible = False
    bDocumento = False
    Label12.Visible = False
    Label13.Visible = False
    Me.fraTranferecia.Visible = False
    lbltransferN.Visible = False
End If

'JUEZ 20141014 Verificar si operación valida cantidad de depositos en mes ****
Dim oCapDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
bValidaCantDep = oCapDef.ValidaCantOperaciones(nOperacion, nProducto, gCapMovDeposito)
Set oCapDef = Nothing
'END JUEZ ********************************************************************

txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledProd = False
txtCuenta.EnabledCMAC = False
cmdgrabar.Enabled = False
cmdcancelar.Enabled = False
fraCliente.Enabled = False
fraDocumento.Enabled = False
fraMonto.Enabled = False
fraTranferecia.Enabled = False
bInstFinanc = False 'JUEZ 20140414

Me.Show 1
End Sub

Private Sub cboPeriodo_Click()
    lblDispCTS.Caption = Format$(CDbl(Trim(Right(cboPeriodo.Text, 5))) * 100, "#,##0.00")
End Sub

Private Sub cboPeriodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtMonto.Enabled Then
        txtMonto.SetFocus
    Else
        cmdgrabar.SetFocus
    End If
End If
End Sub

'***Agregado por ELRO el 20120823, según OYP-RFC024-2012
Private Sub cboTransferMoneda_Click()
    If Right(cboTransferMoneda, 3) = Moneda.gMonedaNacional Then
        'lblSimTra.Caption = "S/."
        lblSimTra.Caption = gcPEN_SIMBOLO 'APRI20191022 SUGERENCIA CALIDAD
        lblMonTra.BackColor = &HC0FFFF
    Else
        lblSimTra.Caption = "$"
        lblMonTra.BackColor = &HC0FFC0
    End If
End Sub
'***Fin Agregado por ELRO el 20120823*******************

Private Sub cboTransferMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdTranfer.SetFocus
    End If
End Sub



'*****************RECO20131024 ERS2141*********************
Private Sub chkDevAfecGarantEcotaxi_Click()
    If chkDevAfecGarantEcotaxi.value = 1 Then
        ActivarControlesDevGarantEcotaxi True
        
    Else
        ActivarControlesDevGarantEcotaxi False
    End If
End Sub
'**********************END RECO***************************

Private Sub chkDevolCtaRecaudo_Click()
    If chkDevolCtaRecaudo.value = 1 Then
        bDevolRecaudoEcotaxi = True
    Else
        bDevolRecaudoEcotaxi = False
    End If
    
End Sub

Private Sub chkITFEfectivo_Click()
    If chkITFEfectivo.value = 1 Then
        'Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00")
        Me.lblTotal.Caption = Format(Me.txtMonto.value + CCur(Me.lblITF.Caption), "#,##0.00")
    Else
        If nProducto = gCapAhorros And gbITFAsumidoAho Then
                    Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
        ElseIf nProducto = gCapPlazoFijo And gbITFAsumidoPF Then
                    Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
        Else
               '     Me.lblTotal.Caption = Format(txtMonto.value - CCur(Me.lblITF.Caption), "#,##0.00")
                    Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
        End If
    
        'Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00")
    End If
End Sub
'*************************RECO20131024 ERS141*********************************
Private Sub cmdBuscaCredGarant_Click()
  Dim loPers As COMDPersona.UCOMPersona 'UPersona
    Dim lsPersCod As String, lsPersNombre As String
    Dim lsEstados As String
    'Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
    Dim loPersCredito  As COMDCredito.DCOMCreditos
    
    Dim lrCreditos As New ADODB.Recordset
    Dim loCuentas As COMDPersona.UCOMProdPersona
    
        
On Error GoTo ControlError

    Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
    Set loPers = Nothing
    
    

    If Trim(lsPersCod) <> "" Then
        Set loPersCredito = New COMDCredito.DCOMCreditos
            Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod)
        Set loPersCredito = Nothing
    End If

    Set loCuentas = New COMDPersona.UCOMProdPersona
        Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
        If loCuentas.sCtaCod <> "" Then
            If Mid(loCuentas.sCtaCod, 6, 3) = "517" Then
                txtClienteGarant.Text = lsPersNombre
            Else
                MsgBox "Sólo es posible seleccionar número de créditos ecotaxi.", vbCritical, "Aviso"
                txtCuentaGarant.NroCuenta = ""
                txtClienteGarant.Text = ""
                Exit Sub
            End If
            txtCuentaGarant.Enabled = True
            txtCuentaGarant.NroCuenta = loCuentas.sCtaCod
            'AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
            sCtaCodAbono = txtCuentaGarant.NroCuenta
            txtCuentaGarant.SetFocusCuenta
        End If
    Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
     
End Sub
'***********END RECO******************************

Private Sub cmdCancelar_Click()
LimpiaControles
gVarPublicas.LimpiaVarLavDinero
End Sub

Private Sub cmdDocumento_Click()
'EJVG20140408 ***
'frmCapAperturaListaChq.Inicia frmCapAbonos, nOperacion, nmoneda, nProducto
'pnValorChq = txtMonto.value
    Dim oform As New frmChequeBusqueda
    Dim lnOperacion As TipoOperacionCheque

    On Error GoTo ErrCargaDocumento
    If nOperacion = gAhoDepChq Or nOperacion = 200244 Then
        lnOperacion = AHO_Deposito
    ElseIf nOperacion = gPFAumCapchq Then
        lnOperacion = DPF_AumentoCapital
    ElseIf nOperacion = gCTSDepChq Then
        lnOperacion = CTS_Deposito
    Else
        lnOperacion = Ninguno
    End If

    Set oDocRec = oform.iniciarBusqueda(nMoneda, lnOperacion, txtCuenta.NroCuenta)
    Set oform = Nothing
    
    txtGlosa.Text = oDocRec.fsGlosa
    lblNombreIF.Caption = oDocRec.fsPersNombre
    lblNroDoc.Caption = oDocRec.fsNroDoc
    sCodIF = oDocRec.fsPersCod
    
    txtMonto.Text = Format(oDocRec.fnMonto, gsFormatoNumeroView)
    txtMonto_Change

    txtGlosa.Locked = True
    txtMonto.Enabled = False
    Exit Sub
ErrCargaDocumento:
    MsgBox "Ha sucedido un error al cargar los datos del Documento", vbCritical, "Aviso"
'END EJVG *******
End Sub
'EJVG20140408 ***
Private Function ValidaSeleccionCheque() As Boolean
    ValidaSeleccionCheque = True
    If oDocRec Is Nothing Then
        ValidaSeleccionCheque = False
        Exit Function
    End If
    If Len(Trim(oDocRec.fsNroDoc)) = 0 Then
        ValidaSeleccionCheque = False
        Exit Function
    End If
End Function
'END EJVG *******
Private Sub cmdGrabar_Click()
'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, nOperacion, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII

'WIOR 20130301 **************************
Dim fbPersonaReaAhorros As Boolean
Dim fnCondicion As Integer
Dim nI As Integer
nI = 0
'WIOR FIN *******************************
Dim sNroDoc As String
Dim nMonto As Double
Dim sCuenta As String
Dim lsmensaje As String
Dim nComixDep As Double ' BRGO 20110127

Dim loLavDinero As frmMovLavDinero
Dim objPersona As COMDPersona.DCOMPersonas 'JACA 20110512
Set objPersona = New COMDPersona.DCOMPersonas 'JACA 20110512
 
Set loLavDinero = New frmMovLavDinero

Dim loMov As COMDMov.DCOMMov
Set loMov = New COMDMov.DCOMMov
Dim lnLogEcotaxi As Integer
Dim oNCOMContImprimir As COMNContabilidad.NCOMContImprimir '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
Set oNCOMContImprimir = New COMNContabilidad.NCOMContImprimir '***Agregado por ELRO el 20120717, según OYP-RFC024-2012

'PASI20140530
Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
'end PASI
'APRI20190109 ERS077-2018
Dim nCantOpeCta As Integer
Dim nComision As Currency
Dim nComisionLog As Double 'Added by TORE: Correccion comision para Cajeros Corresponsales
'END APRI
nMonto = txtMonto.value
If nMonto = 0 Then
    MsgBox "Monto debe ser mayor a cero", vbInformation, "Aviso"
    If txtMonto.Enabled Then txtMonto.SetFocus
    Exit Sub
End If
sNroDoc = Trim(lblNroDoc.Caption)
If bDocumento Then
    If sNroDoc = "" Then
        MsgBox "Debe seleccionar un cheque válido para la operacion.", vbInformation, "Aviso"
        cmdDocumento.SetFocus
        Exit Sub
    End If
End If

'ANDE 20180419 ERS021-2018 camapaña mundialito
    Dim cperscod As String
    Dim nTitularCod As Integer, nTipoPersona As Integer, bParticipaCamp As Boolean, cTextoDatos As String
    Dim ix As Integer
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    For ix = 1 To grdCliente.Rows - 1
        nTitularCod = Val(Right(Trim(grdCliente.TextMatrix(ix, 3)), 2))
        If nTitularCod = 10 Then
            'nTipoPersona = grdCliente.TextMatrix(ix, 1)
            cperscod = grdCliente.TextMatrix(ix, 1)
            nTipoPersona = oCaptaLN.getVerificarPersonaNatJur(cperscod)
        End If
    Next ix
    'end ande


If nProducto = gCapAhorros Then
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim clsCapMnt As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim nMontoMinDep As Double
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    'txtCuenta.NroCuenta
    'JUEZ 20141014 Nuevos parámetros ***********************************
    'nMontoMinDep = clsDef.GetMontoMinimoDepPersoneria(gCapAhorros, Mid(txtCuenta.NroCuenta, 9, 1), nPersoneria, pbOrdPag)
    'If nMontoMinDep > txtMonto.value Then
    '   MsgBox "El Monto del Abono es menor al mínimo permitido de " & IIf(Mid(txtCuenta.NroCuenta, 9, 1) = 1, "S/. ", "US$. ") & CStr(nMontoMinDep), vbOKOnly + vbInformation, "Aviso"
    '   Exit Sub
    'End If
    nMontoMinDep = IIf(Mid(txtCuenta.NroCuenta, 9, 1) = gMonedaNacional, nParMontoMinDepSol, nParMontoMinDepDol)
    If txtMonto.value < nMontoMinDep Then
        'MsgBox "El Monto de Abono no debe ser menor de " & IIf(Mid(txtCuenta.NroCuenta, 9, 1) = gMonedaNacional, "S/.", "$") & " " & Format(nMontoMinDep, "#,##0.00"), vbInformation, "Aviso"
        MsgBox "El Monto de Abono no debe ser menor de " & IIf(Mid(txtCuenta.NroCuenta, 9, 1) = gMonedaNacional, gcPEN_SIMBOLO, "$") & " " & Format(nMontoMinDep, "#,##0.00"), vbInformation, "Aviso"  'APRI20191022 SUGERENCIA CALIDAD
        txtMonto.SetFocus 'APRI20191022 SUGERENCIA CALIDAD
        txtMonto.value = nMontoMinDep 'APRI20191022 SUGERENCIA CALIDAD
        Exit Sub
    End If
    'END JUEZ **********************************************************
    Set clsDef = Nothing
    Dim clsCC As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set clsCC = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim bEsCtaCC As Boolean
    bEsCtaCC = clsCC.ValidaCuentaCC(txtCuenta.NroCuenta, nOperacion, 0, "", 0, 2)
    Set clsCC = Nothing
    'APRI20190109 ERS077-2018
    If bValidaCantDep Then
        Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        nCantOpeCta = clsMant.ObtenerCantidadOperaciones(txtCuenta.NroCuenta, gCapMovDeposito, gdFecSis)
        nCantOpeCta = nCantOpeCta + 1 'Por la operación actual
        nComision = clsMant.ObtenerCapValorComision(txtCuenta.NroCuenta, nCantOpeCta, 2, 1)
        nComisionLog = CDbl(nComision) 'Added by TORE: Correccion comision para Cajeros Corresponsales
        
        Set clsMant = Nothing
        'If nComision > 0 Then
         'If MsgBox("La operación solicitada genera un cargo de " & IIf(Mid(txtCuenta.NroCuenta, 9, 1) = 1, gcPEN_SIMBOLO & " ", IIf(Mid(txtCuenta.NroCuenta, 9, 1) = 2, "$. ", "Eu.")) & Format$(nComision, "#,##0.00") & ", por exceso de operaciones de depósitos. Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
                'End If
                'ADD JHCU
                If nComision > 0 Then
         If Not bEsCtaCC Then
            If MsgBox("La operación solicitada genera un cargo de " & IIf(Mid(txtCuenta.NroCuenta, 9, 1) = 1, gcPEN_SIMBOLO & " ", IIf(Mid(txtCuenta.NroCuenta, 9, 1) = 2, "$. ", "Eu.")) & Format$(nComision, "#,##0.00") & ", por exceso de operaciones de depósitos. Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
            End If
         End If
         'END JHCU
        End If
    'END APRI
    '***BRGO 20110127 ********************************
    
        'Modificado by JACA 20111021************************************
            If nOperacion = gAhoDepEfec Or nOperacion = gAhoDepChq Or nOperacion = gAhoDepTransf Then
                nComixDep = Round(CalcularComisionDepOtraAge(), 2)
                'ADD JHCU
                If bEsCtaCC Then
                    nComision = 0 'Added by TORE: Correccion comision para Cajeros Corresponsales
                    nComixDep = 0
                End If
            Else
                nComixDep = 0
            End If
            
           
        'JACA END********************************************************
        
        'Comentado by JACA 20111021***************************************************
            'If (lnTpoPrograma = 5 Or lnTpoPrograma = 6) Or nOperacion <> gAhoDepEfec Then
                'nComixDep = 0
            'End If
        'JACA END***********************************************************************
        
    '***END BRGO *************************************
    'APRI20190109 ERS077-2018
    Dim ArrDistCom As Variant
    ReDim ArrDistCom(2)
    ArrDistCom(0) = CCur(nComixDep)
    ArrDistCom(1) = CCur(nComision)
    'END APRI
    If lnTpoPrograma = 2 Or lnTpoPrograma = 4 Then
       
        Dim lsMenAbono As String
        Set clsCapMnt = New COMDCaptaGenerales.DCOMCaptaGenerales
        clsCapMnt.GetAbonoPactadoSubProductoAho txtCuenta.NroCuenta, txtMonto.value, gdFecSis, lnTpoPrograma, lsMenAbono
        If Trim(lsMenAbono) <> "" Then
            MsgBox lsMenAbono, vbInformation, "Aviso"
            Set clsCapMnt = Nothing
            Exit Sub
        Else
            Set clsCapMnt = Nothing
        End If
    End If
    If nOperacion = gAhoDepChq Then
        If lnTpoPrograma = 4 Then
            'Validar Monto de Cheque para Deposito de Aho-Destino... AVMM  16-03-2007
            Dim nValorCh As Double
            Dim nDifValorCh As Double
            Dim nDifTotalCh As Double
            Dim nPagadoTotal As Double
            Set clsCapMnt = New COMDCaptaGenerales.DCOMCaptaGenerales
            'nValorCh = clsCapMnt.GetObtenerMontoCheque(lblNroDoc.Caption)
            'nDifValorCh = Format((CDbl(pnValorChq) - CDbl(nValorCh)), "0.00")
            nPagadoTotal = CDbl(txtMonto.value)
            nDifValorCh = Format((CDbl(pnValorChq) - CDbl(nPagadoTotal)), "0.00")
            'nDifTotalCh = (CDbl(nDifValorCh) - CDbl(nPagadoTotal))
            
            If nDifValorCh < 0 Then
                MsgBox "No se puede realizar el Deposito con Cheque solo dispone de: " & pnValorChq, vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    End If
End If
'ALPA 20130107*************************************
lnLogEcotaxi = 0
ReDim Preserve sPerEcotaxi(1 To 3)
    sPerEcotaxi(1) = ""
    sPerEcotaxi(2) = ""
    sPerEcotaxi(3) = ""
If chkIniciarEcotaxi.value = 1 Then
    sPersCodEcotaxi = ""
    Call frmCapAbonoIniciarEcotaxi.Inicio(sPersCodEcotaxi, sPersNomEcotaxi, sPersDNIEcotaxi)
    If sPersCodEcotaxi = "" Then
        Exit Sub
    End If
    'ReDim Preserve sPerEcotaxi(1 To 3)
    sPerEcotaxi(1) = sPersCodEcotaxi
    sPerEcotaxi(2) = sPersNomEcotaxi
    sPerEcotaxi(3) = sPersDNIEcotaxi
    lnLogEcotaxi = 1
End If
'**************************************************

If bCuentaWestern = True Then
  If txtClienteWesterUnion.Text = "" Then
        MsgBox "Debe ingresar el codigo del cliente Wester Union.", vbInformation, "Aviso"
        Me.txtClienteWesterUnion.SetFocus
        Exit Sub
  End If
End If
                
If nOperacion = gAhoDepTransf Or nOperacion = gCTSDepTransf Then
    If lblTrasferND.Caption = "" Then
        MsgBox "Debe ingresar un numero de transacción.", vbInformation, "Aviso"
        Me.cmdTranfer.SetFocus
        Exit Sub
    End If
End If
'****** Comentado porque para creditos no debe de Existir Validar Cuentas *******

Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
'If chkITFEfectivo.value = 0 Then
'    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
'        If clsCap.ValidaSaldoCuenta(txtCuenta.NroCuenta, CDbl(lblITF.Caption)) = False Then
'        MsgBox "No existe saldo Suficiente para realizar esta Operación", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    Set clsCap = Nothing
'End If

'*********************************************************************************
'** BRGO 20110425*** Valida si el titular de la cuenta tiene actualizado su registro de sueldos
If nProducto = gCapCTS Then
    Dim clsCapt As New DCOMCaptaMovimiento
    Dim clsGen As New COMDConstSistema.DCOMGeneral 'JUEZ 20130815
    Dim R As ADODB.Recordset
    Dim dFecha As Date
    'Dim nDiasVerifica As Integer 'JUEZ 20130815 'JUEZ 20141014 Comentar para nuevos parámetros
    'nDiasVerifica = clsGen.GetParametro(gPrdParamCaptac, 1006) 'JUEZ 20130815 'JUEZ 20141014 Comentar para nuevos parámetros
    Set R = clsCapt.ObtenerFecUltimaActSueldosCTS(txtCuenta.NroCuenta)
    If R.BOF Or R.EOF Then
        MsgBox "No se encontraron registros de sueldos del titular de la cuenta. Debe registrar el total de los " & nParUltRemunBrutas & " últimos sueldos para proceder"
        Exit Sub
    Else
        dFecha = R!FechaAct
        'If DateDiff("d", dFecha, gdFecSis) > 30 Then
        'If DateDiff("d", dFecha, gdFecSis) > nDiasVerifica Then 'JUEZ 20130815
        If DateDiff("d", dFecha, gdFecSis) > nParDiasVerifRegSueldo Then 'JUEZ 20141014
            MsgBox "La última actualización ha caducado. Favor actualice su registro de Sueldos", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    R.Close
    Set R = Nothing
    Set clsCapt = Nothing
End If
'*************************************************************************
'JACA 20110512 *****VERIFICA SI LAS PERSONAS CUENTAN CON OCUPACION E INGRESO PROMEDIO
        Dim rsPersVerifica As Recordset
        Dim i As Integer
        Set rsPersVerifica = New Recordset
        For i = 1 To grdCliente.Rows - 1
            Call VerSiClienteActualizoAutorizoSusDatos(grdCliente.TextMatrix(i, 1), nOperacion) 'FRHU ERS077-2015 20151204
            Set rsPersVerifica = objPersona.ObtenerDatosPersona(Me.grdCliente.TextMatrix(i, 1))
            If rsPersVerifica!nPersIngresoProm = 0 Or rsPersVerifica!cActiGiro1 = "" Then
                If MsgBox("Necesita Registrar la Ocupacion e Ingreso Promedio de: " + Me.grdCliente.TextMatrix(i, 2), vbYesNo) = vbYes Then
                    'frmPersona.Inicio Me.grdCliente.TextMatrix(i, 1), PersonaActualiza
                    frmPersOcupIngreProm.Inicio Me.grdCliente.TextMatrix(i, 1), Me.grdCliente.TextMatrix(i, 2), rsPersVerifica!cActiGiro1, rsPersVerifica!nPersIngresoProm
                End If
            End If
        Next i
    'JACA END***************************************************************************
    'WIOR 20121009 Clientes Observados *************************************
    If nOperacion = gAhoDepEfec Or nOperacion = gAhoDepChq Or nOperacion = gCTSDepEfec Or nOperacion = gCTSDepChq Then
        Dim oDPersona As COMDPersona.DCOMPersona
        Dim rsPersona As ADODB.Recordset
        Dim sCodPersona As String
        Dim Cont As Integer
        
        Set oDPersona = New COMDPersona.DCOMPersona
        
        For Cont = 0 To grdCliente.Rows - 2
            If Trim(Right(grdCliente.TextMatrix(Cont + 1, 3), 5)) = gCapRelPersTitular Then
                sCodPersona = Trim(grdCliente.TextMatrix(Cont + 1, 1))
                Set rsPersona = oDPersona.ObtenerUltimaVisita(sCodPersona)
                If rsPersona.RecordCount > 0 Then
                    If Not (rsPersona.EOF And rsPersona.BOF) Then
                        If Trim(rsPersona!sUsual) = "3" Then
                            MsgBox Trim(grdCliente.TextMatrix(Cont + 1, 2)) & "." & Chr(10) & "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                            Call frmPersona.Inicio(sCodPersona, PersonaActualiza)
                        End If
                    End If
                End If
                Set rsPersona = Nothing
            End If
        Next Cont
    End If
    'WIOR FIN ***************************************************************
    
    'AMDO 20130702 TI-ERS063-2013 ****************************************************
    If nOperacion = gAhoDepEfec Or nOperacion = gCTSDepEfec Then
        Dim oDPersonaAct As COMDPersona.DCOMPersona
        Dim conta As Integer
        Dim sPersCod As String
        
        Set oDPersonaAct = New COMDPersona.DCOMPersona
        For conta = 0 To grdCliente.Rows - 2
        sPersCod = Trim(grdCliente.TextMatrix(conta + 1, 1))
                        If oDPersonaAct.VerificaExisteSolicitudDatos(sPersCod) Then
                            MsgBox Trim("SE SOLICITA DATOS DEL CLIENTE: " & grdCliente.TextMatrix(conta + 1, 2)) & "." & Chr(10), vbInformation, "Aviso"
                            Call frmActInfContacto.Inicio(sPersCod)
                        End If
        Next conta
    End If
    'AMDO FIN ********************************************************************************
    'EJVG20140408 ***
    If nOperacion = gAhoDepChq Or nOperacion = 200244 Or nOperacion = gPFAumCapchq Or nOperacion = gCTSDepChq Then
        If Not ValidaSeleccionCheque Then
            MsgBox "Ud. debe seleccionar un Cheque para continuar", vbInformation, "Aviso"
            If cmdDocumento.Visible And cmdDocumento.Enabled Then cmdDocumento.SetFocus
            Exit Sub
        End If
    End If
    'END EJVG *******
    
    '********************* ANDE 20170623 mejoras en validación de CIUU
    Dim rsPersOcu As Recordset
    Set rsPersOcu = objPersona.ObtenerDatosPersona(Me.grdCliente.TextMatrix(1, 1))
    If Not (rsPersOcu.BOF And rsPersOcu.EOF) Then
        If IsNull(rsPersOcu!cPersCIIU) Then
            MsgBox "Cliente no cuenta con código CIIU, por favor actualice los datos para continuar.", vbInformation + vbOKOnly, "Aviso"
            Exit Sub
        End If
    Else
      MsgBox "No se pudo recuperar los datos del cliente"
    End If
    '********************* end validación de ciuu
    
If MsgBox("¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
   
    Dim sMovNro As String, sPersLavDinero As String
    Dim ClsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
    Dim nSaldo As Double, nPorcDisp As Double
    Dim nMontoLavDinero As Double, nTC As Double
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios  'NCapServicios
    Dim previo As New previo.clsprevio
    Dim lsBoletaImp As String
    Dim lsBoletaImpITF As String
    Dim nFicSal As Integer
    Dim lsBoletaCVME As String '***Agregado por ELRO el 20120726, según OYP-RFC024-2012
    
     'RIRO20170714 ****
    Dim sValida As String
    If nOperacion = gAhoDepTransf Or nOperacion = gCTSDepTransf Then
        sValida = validaVoucher
        If Len(Trim(sValida)) > 0 Then
            MsgBox sValida, vbInformation, "Verifica validación de voucher"
            Exit Sub
        End If
    End If
    'END RIRO ********
    
    'Realiza la Validación para el Lavado de Dinero
    sCuenta = txtCuenta.NroCuenta
    Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
    'If clsLav.EsOperacionEfectivo(Trim(nOperacion)) Then
        Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
        If Not clsExo.EsCuentaExoneradaLavadoDinero(sCuenta) Then
            Set clsExo = Nothing
            
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            If nMoneda = gMonedaNacional Then
                Dim clsTC As COMDConstSistema.NCOMTipoCambio
                Set clsTC = New COMDConstSistema.NCOMTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
            Else
                nTC = 1
            End If
            If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                'bY cAPI 18022008
                Call IniciaLavDinero(loLavDinero, txtClienteWesterUnion.Text, txtClienteWesterUnion.psDescripcion, bCuentaWestern)
                sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 27), True, sTipoCuenta, , , , , CInt(Mid(sCuenta, 9, 1)), , gnTipoREU, gnMontoAcumulado, gsOrigen)
                If loLavDinero.OrdPersLavDinero = "" Then
                    Exit Sub
                End If
            
            End If
        Else
            Set clsExo = Nothing
        End If
    '''''Else *
    '    Set clsLav = Nothing
   ' End If
    'WIOR 20130301 Personas Sujetas a Procedimiento Reforzado *************************************
    fbPersonaReaAhorros = False
    If loLavDinero.OrdPersLavDinero = "Exit" _
            And (nOperacion = gAhoDepEfec Or nOperacion = gAhoDepChq Or nOperacion = gAhoDepTransf _
            Or nOperacion = gCTSDepEfec Or nOperacion = gCTSDepChq Or nOperacion = gCTSDepTransf Or nOperacion = gCTSDepAboOtrosConceptos) Then
            
            Dim oPersonaSPR As UPersona_Cli
            Dim oPersonaU As COMDPersona.UCOMPersona
            Dim nTipoConBN As Integer
            Dim sConPersona As String
            Dim pbClienteReforzado As Boolean
            Dim rsAgeParam As Recordset
            Dim objCap As COMNCaptaGenerales.NCOMCaptaMovimiento
            Dim lnMonto As Double, lnTC As Double
            Dim ObjTc As COMDConstSistema.NCOMTipoCambio
            
            
            Set oPersonaU = New COMDPersona.UCOMPersona
            Set oPersonaSPR = New UPersona_Cli
            
            fbPersonaReaAhorros = False
            pbClienteReforzado = False
            fnCondicion = 0
            
            For nI = 0 To grdCliente.Rows - 2
                oPersonaSPR.RecuperaPersona Trim(grdCliente.TextMatrix(nI + 1, 1))
                                    
                If oPersonaSPR.Personeria = 1 Then
                    If oPersonaSPR.Nacionalidad <> "04028" Then
                        sConPersona = "Extranjera"
                        fnCondicion = 1
                        pbClienteReforzado = True
                        Exit For
                    ElseIf oPersonaSPR.Residencia <> 1 Then
                        sConPersona = "No Residente"
                        fnCondicion = 2
                        pbClienteReforzado = True
                        Exit For
                    ElseIf oPersonaSPR.RPeps = 1 Then
                        sConPersona = "PEPS"
                        fnCondicion = 4
                        pbClienteReforzado = True
                        Exit For
                    ElseIf oPersonaU.ValidaEnListaNegativaCondicion(IIf(Trim(oPersonaSPR.ObtenerDNI) = "", oPersonaSPR.ObtenerNumeroDoc(0), oPersonaSPR.ObtenerDNI), oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                        If nTipoConBN = 1 Or nTipoConBN = 3 Then
                            sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                            fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                            pbClienteReforzado = True
                            Exit For
                        End If
                    End If
                Else
                    If oPersonaU.ValidaEnListaNegativaCondicion(oPersonaSPR.ObtenerDNI, oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                        If nTipoConBN = 1 Or nTipoConBN = 3 Then
                            sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                            fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                            pbClienteReforzado = True
                            Exit For
                        End If
                    End If
                End If
            Next nI
            
            If pbClienteReforzado Then
                MsgBox "El Cliente: " & Trim(grdCliente.TextMatrix(nI + 1, 2)) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
                frmPersRealizaOpeGeneral.Inicia Me.Caption & " (Persona " & sConPersona & ")", nOperacion, 1
                fbPersonaReaAhorros = frmPersRealizaOpeGeneral.PersRegistrar
                
                If Not fbPersonaReaAhorros Then
                    MsgBox "Se va a proceder a Anular el Abono de la Cuenta", vbInformation, "Aviso"
                    cmdgrabar.Enabled = True
                    Exit Sub
                End If
            Else
                fnCondicion = 0
                lnMonto = nMonto
                pbClienteReforzado = False
                
                Set ObjTc = New COMDConstSistema.NCOMTipoCambio
                lnTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set ObjTc = Nothing
            
            
                Set objCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                Set rsAgeParam = objCap.getCapAbonoAgeParam(gsCodAge)
                Set objCap = Nothing
                
                If Mid(Trim(txtCuenta.NroCuenta), 9, 1) = 1 Then
                    lnMonto = Round(lnMonto / lnTC, 2)
                End If
            
                If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                    If lnMonto >= rsAgeParam!nMontoMin And lnMonto <= rsAgeParam!nMontoMax Then
                        frmPersRealizaOpeGeneral.Inicia Me.Caption, nOperacion
                        fbPersonaReaAhorros = frmPersRealizaOpeGeneral.PersRegistrar
                        If Not fbPersonaReaAhorros Then
                            MsgBox "Se va a proceder a Anular el Abono de la Cuenta", vbInformation, "Aviso"
                            cmdgrabar.Enabled = True
                            Exit Sub
                        End If
                    End If
                End If
                
            End If
    End If
    'WIOR FIN ***************************************************************
    If oDocRec Is Nothing Then Set oDocRec = New UDocRec 'EJVG20140408
    Set ClsMov = New COMNContabilidad.NCOMContFunciones 'NContFunciones
    sMovNro = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set ClsMov = Nothing
    On Error GoTo ErrGraba
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
    clsCap.IniciaImpresora gImpresora
    Select Case nProducto
        Case gCapAhorros
                'WIOR 20130301 comento to este codigo -INICIO
                ''JACA 20110317 PARA REGISTRAR LA PERSONA Q REALIZA EL DEPOSITO MENOR CUANTIA************
                'Dim regPersonaReaDep As Boolean
                'Dim rsAgeParam As Recordset
                'Dim lnMonto As Double
                'regPersonaReaDep = False
                'If (lnTpoPrograma = 0 Or lnTpoPrograma = 1 Or lnTpoPrograma = 2 Or lnTpoPrograma = 3 Or lnTpoPrograma = 4 Or lnTpoPrograma = 5) And sPersLavDinero = "" Then 'JACA 20110325
                '
                '        Set clsTC = New COMDConstSistema.NCOMTipoCambio
                '        nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                '        Set clsTC = Nothing
                '
                '        Set rsAgeParam = clsCap.getCapAbonoAgeParam(gsCodAge)
                '        lnMonto = nMonto
                '        If Mid(sCuenta, 9, 1) = 1 Then
                '            lnMonto = Round(lnMonto / nTC, 2)
                '        End If
                '        If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then 'JACA 20110325
                '            If lnMonto >= rsAgeParam!nMontoMin And lnMonto <= rsAgeParam!nMontoMax Then
                '                'frmCapAbonosPersRealiza.Inicia 'WIOR 20130301 COMENTO
                '                'regPersonaReaDep = frmCapAbonosPersRealiza.PersRegistrar 'WIOR 20130301 COMENTO
                '                'WIOR 20130301 **************************************************
                '                frmPersRealizaOpeGeneral.Inicia sOperacion, nOperacion
                '                regPersonaReaDep = frmPersRealizaOpeGeneral.PersRegistrar
                '                'WIOR FIN ***************************************************
                '                If Not regPersonaReaDep Then
                '                    MsgBox "Se va a proceder a Cancelar la Operacion"
                '                    Exit Sub
                '                End If
                '            End If
                '        End If
                'End If
                ''END JACA*********************************************************
                ''WIOR 20121114 Personas Sujetas a Procedimiento Reforzado *************************************
                'If loLavDinero.OrdPersLavDinero = "Exit" And regPersonaReaDep = False And (nOperacion = gAhoDepEfec Or nOperacion = gAhoDepChq Or nOperacion = gAhoDepTransf) Then
                '    Dim oPersonaSPR As UPersona_Cli
                '    Dim rsPersonaSPR As ADODB.Recordset
                '    Dim sCodPersonaSPR As String
                '    Dim ContSPR As Integer
                '    Dim oPersonaU As COMDPersona.UCOMPersona
                '    Dim nTipoConBN As Integer
                '    Dim sNombrePersona As String
                '    Dim sConPersona As String
                '    Dim nOperacionRea As Integer
                '
                '    Select Case nOperacion
                '        Case gAhoDepEfec: nOperacionRea = 5
                '        Case gAhoDepChq: nOperacionRea = 6
                '        Case gAhoDepTransf: nOperacionRea = 7
                '    End Select
                '
                '    Dim pbClienteReforzado As Boolean
                '    Set oPersonaU = New COMDPersona.UCOMPersona
                '    Set oPersonaSPR = New UPersona_Cli
                '    pbClienteReforzado = False
                '    fnCondicion = 0
                '    For ContSPR = 0 To grdCliente.Rows - 2
                '        If Trim(Right(grdCliente.TextMatrix(ContSPR + 1, 3), 5)) = gCapRelPersTitular Then
                '            sCodPersonaSPR = Trim(grdCliente.TextMatrix(ContSPR + 1, 1))
                '            sNombrePersona = Trim(grdCliente.TextMatrix(ContSPR + 1, 2))
                '            oPersonaSPR.RecuperaPersona sCodPersonaSPR
                '            If oPersonaSPR.Personeria = 1 Then
                '                If oPersonaSPR.Nacionalidad <> "04028" Then
                '                    sConPersona = "Extranjera"
                '                    fnCondicion = 1
                '                    pbClienteReforzado = True
                '                    Exit For
                '                ElseIf oPersonaSPR.Residencia <> 1 Then
                '                    sConPersona = "No Residente"
                '                    fnCondicion = 2
                '                   pbClienteReforzado = True
                '                    Exit For
                '                ElseIf oPersonaSPR.RPeps = 1 Then
                '                    sConPersona = "PEPS"
                '                    fnCondicion = 4
                '                    pbClienteReforzado = True
                '                    Exit For
                '                ElseIf oPersonaU.ValidaEnListaNegativaCondicion(IIf(Trim(oPersonaSPR.ObtenerDNI) = "", oPersonaSPR.ObtenerNumeroDoc(0), oPersonaSPR.ObtenerDNI), oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                '                    If nTipoConBN = 1 Or nTipoConBN = 3 Then
                '                        sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                '                        fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                '                        pbClienteReforzado = True
                '                        Exit For
                '                    End If
                '                End If
                '            Else
                '                If oPersonaU.ValidaEnListaNegativaCondicion(oPersonaSPR.ObtenerDNI, oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                '                    If nTipoConBN = 1 Or nTipoConBN = 3 Then
                '                        sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                '                        fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                '                        pbClienteReforzado = True
                '                        Exit For
                '                    End If
                '                End If
                '            End If
                '        End If
                '    Next ContSPR
                '
                '    If pbClienteReforzado Then
                '        MsgBox "El Cliente: " & sNombrePersona & " es un cliente de procedimiento reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
                '        frmPersRealizaOperacion.Inicia "Deposito (Persona " & sConPersona & ")", nOperacionRea, 1
                '        fnDepositoPersRealiza = frmPersRealizaOperacion.PersRegistrar
                '
                '        If Not fnDepositoPersRealiza Then
                '            MsgBox "Se va a proceder a Anular el Deposito de la Cuenta.", vbInformation, "Aviso"
                '            Exit Sub
                '        End If
                '    End If
                'End If
                'WIOR FIN ***************************************************************
                'ANDE 20180419 ERS021-2018 participanción en campaña mundialto
                Dim nCondicion As Integer, nPuntosRef As Integer, nPTotalAcumulado As Integer
                If _
                    nOperacion = gAhoDepEfec Or nOperacion = gAhoDepChq Or nOperacion = gAhoDepTransf _
                Then
                    Dim nOpeTipo As Integer
                    nOpeTipo = 2 '1:DEPOSITO
                    nMoneda = Mid(sCuenta, 9, 1)
                    If nMoneda = gMonedaNacional Then
                        'Dim oCampanhaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
                        Call oCaptaLN.VerificarParticipacionCampMundial(cperscod, sCuenta, nOperacion, nOpeTipo, nMoneda, nMonto, nTipoPersona, bParticipaCamp, sMovNro, , gdFecSis, lnTpoPrograma, txtCuenta.Age, nPuntosRef, nCondicion, , nPTotalAcumulado)
                        If bParticipaCamp Then
                            cTextoDatos = "#" & IIf(bParticipaCamp, "1", "0") & "." & CStr(nPuntosRef) & "$" & CStr(nCondicion) & "_" & CStr(nPTotalAcumulado) & "&"
                            lsmensaje = cTextoDatos
                        End If
                    End If
                End If
            
                
                'WIOR 20130301 comento to este codigo -FIN **************************************************************
                If sPersCodCMAC = "" Then
                    If bDocumento Then
                          'nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, TpoDocCheque, sNroDoc, sCodIF, dFechaValorizacion, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.ReaPersLavDinero, , , , , gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, , loLavDinero.BenPersLavDinero, lsMensaje, lsBoletaImp, lsBoletaImpITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComixDep, , , , , , , , , , , sPerEcotaxi, lnLogEcotaxi, sCtaCodAbono, bDevolRecaudoEcotaxi)
                          dFechaValorizacion = oNCapMov.ObtenerFechaValorizaCheque(oDocRec.fsNroDoc, oDocRec.fsPersCod, oDocRec.fsIFTpo, oDocRec.fsIFCta) 'PASI20140530
                          nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, oDocRec.fnTpoDoc, oDocRec.fsNroDoc, oDocRec.fsPersCod, dFechaValorizacion, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.ReaPersLavDinero, , , , , gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, , loLavDinero.BenPersLavDinero, lsmensaje, lsBoletaImp, lsBoletaImpITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, ArrDistCom, , , , , , , , , , , sPerEcotaxi, lnLogEcotaxi, sCtaCodAbono, bDevolRecaudoEcotaxi, , oDocRec.fsIFTpo, oDocRec.fsIFCta)  'EJVG20140408
                          'APRI20190109 ERS077-2018 CHANGE nComixDep TO ArrDistCom
                    Else
                        If cboTransferMoneda.Text <> "" Then
                            '***Modificado por ELRO el 20120726, según OYP-RFC024-2012
                            'nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, IIf(nOperacion = gAhoDepTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), , , , , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.TitPersLavDinero, , , lnMovNroTransfer, Right(Me.cboTransferMoneda.Text, 3), gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, , loLavDinero.BenPersLavDinero, lsmensaje, lsBoletaImp, lsBoletaImpITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComixDep)
                            nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, IIf(nOperacion = gAhoDepTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), , , , , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.TitPersLavDinero, , , lnMovNroTransfer, Right(Me.cboTransferMoneda.Text, 3), gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, , loLavDinero.BenPersLavDinero, lsmensaje, lsBoletaImp, lsBoletaImpITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, ArrDistCom, , , , , , , , , _
                                                              fnMovNroRVD, CCur(lblMonTra), sPerEcotaxi, lnLogEcotaxi, sCtaCodAbono, bDevolRecaudoEcotaxi)
                            'APRI20190109 ERS077-2018 CHANGE nComixDep TO ArrDistCom
                            '***Modificado por ELRO el 20120726***********************
                        Else
                            nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, IIf(nOperacion = gAhoDepTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), , , , , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.TitPersLavDinero, , , lnMovNroTransfer, , gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, , loLavDinero.BenPersLavDinero, lsmensaje, lsBoletaImp, lsBoletaImpITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, ArrDistCom, , , , , , , , sNumTarj, , , sPerEcotaxi, lnLogEcotaxi, sCtaCodAbono, bDevolRecaudoEcotaxi)
                            'APRI20190109 ERS077-2018 CHANGE nComixDep TO ArrDistCom
                        End If
                    End If
                Else
                    If bDocumento Then
                        'nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, TpoDocCheque, sNroDoc, sCodIF, dFechaValorizacion, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.TitPersLavDinero, , , , , gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCMACCargo, gITFCobroCMAC), , , loLavDinero.BenPersLavDinero, lsMensaje, lsBoletaImp, lsBoletaImpITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComixDep, , , , , , , , , , , sPerEcotaxi, lnLogEcotaxi, sCtaCodAbono, bDevolRecaudoEcotaxi)
                        dFechaValorizacion = oNCapMov.ObtenerFechaValorizaCheque(oDocRec.fsNroDoc, oDocRec.fsPersCod, oDocRec.fsIFTpo, oDocRec.fsIFCta) 'PASI20140530
                        nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, oDocRec.fnTpoDoc, oDocRec.fsNroDoc, oDocRec.fsPersCod, dFechaValorizacion, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.TitPersLavDinero, , , , , gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCMACCargo, gITFCobroCMAC), , , loLavDinero.BenPersLavDinero, lsmensaje, lsBoletaImp, lsBoletaImpITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, ArrDistCom, , , , , , , , , , , sPerEcotaxi, lnLogEcotaxi, sCtaCodAbono, bDevolRecaudoEcotaxi, , oDocRec.fsIFTpo, oDocRec.fsIFCta) 'EJVG20140408
                        'APRI20190109 ERS077-2018 CHANGE nComixDep TO ArrDistCom
                    Else
                        If cboTransferMoneda.Text <> "" Then
                            '***Modificado por ELRO el 20120726, según OYP-RFC024-2012
                            'nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, IIf(nOperacion = gAhoDepTransf, Trim(txtTransferGlosa.Text), Trim(txtglosa.Text)), , , , , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.TitPersLavDinero, , , lnMovNroTransfer, Right(Me.cboTransferMoneda.Text, 3), gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCMACCargo, gITFCobroCMAC), , , loLavDinero.BenPersLavDinero, lsmensaje, lsBoletaImp, lsBoletaImpITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComixDep)
                            nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, IIf(nOperacion = gAhoDepTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), , , , , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.TitPersLavDinero, , , lnMovNroTransfer, Right(Me.cboTransferMoneda.Text, 3), gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCMACCargo, gITFCobroCMAC), , , loLavDinero.BenPersLavDinero, lsmensaje, lsBoletaImp, lsBoletaImpITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, ArrDistCom, , , , , , , , , _
                                                              fnMovNroRVD, CCur(lblMonTra), sPerEcotaxi, lnLogEcotaxi, sCtaCodAbono, bDevolRecaudoEcotaxi)
                            'APRI20190109 ERS077-2018 CHANGE nComixDep TO ArrDistCom
                            '***Modificado por ELRO el 20120726***********************
                        Else
                            nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, IIf(nOperacion = gAhoDepTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), , , , , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.TitPersLavDinero, , , lnMovNroTransfer, , gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCMACCargo, gITFCobroCMAC), , , loLavDinero.BenPersLavDinero, lsmensaje, lsBoletaImp, lsBoletaImpITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, ArrDistCom, , , , , , , , , , , sPerEcotaxi, lnLogEcotaxi, sCtaCodAbono, bDevolRecaudoEcotaxi)
                            'APRI20190109 ERS077-2018 CHANGE nComixDep TO ArrDistCom
                        End If
                    End If
                End If
                'WIOR 20130301 comento to este codigo -INICIO ********************************************************************
                ''JACA 20110317
                'If regPersonaReaDep Then
                '    frmCapAbonosPersRealiza.insertarPersonaDeposita gnMovNro, frmCapAbonosPersRealiza.PersCod, frmCapAbonosPersRealiza.PersDNI, frmCapAbonosPersRealiza.PersNombre
                'End If
                ''END JACA
                ''WIOR 20121114 ************************************
                'If fnDepositoPersRealiza Then
                '    frmPersRealizaOperacion.InsertaPersonaRealizaOperacion gnMovNro, sCuenta, frmPersRealizaOperacion.PersTipoCliente, _
                '    frmPersRealizaOperacion.PersCod, frmPersRealizaOperacion.PersTipoDOI, frmPersRealizaOperacion.PersDOI, frmPersRealizaOperacion.PersNombre, _
                '    frmPersRealizaOperacion.TipoOperacion, frmPersRealizaOperacion.Origen, fnCondicion
                '
                '    fnDepositoPersRealiza = False
                '    fnCondicion = 0
                'End If
                'WIOR FIN *****************************************
                'WIOR 20130301 comento to este codigo -FIN ************************************************************************
        Case gCapCTS
            nPorcDisp = CDbl(lblDispCTS)
            If bDocumento Then
                'nSaldo = clsCap.CapAbonoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nPorcDisp, TpoDocCheque, sNroDoc, sCodIF, dFechaValorizacion, , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.TitPersLavDinero, , , gsCodCMAC, loLavDinero.BenPersLavDinero, lsMensaje, lsBoletaImp, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
                dFechaValorizacion = oNCapMov.ObtenerFechaValorizaCheque(oDocRec.fsNroDoc, oDocRec.fsPersCod, oDocRec.fsIFTpo, oDocRec.fsIFCta) 'PASI20140530
                nSaldo = clsCap.CapAbonoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nPorcDisp, oDocRec.fnTpoDoc, oDocRec.fsNroDoc, oDocRec.fsPersCod, dFechaValorizacion, , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.TitPersLavDinero, , , gsCodCMAC, loLavDinero.BenPersLavDinero, lsmensaje, lsBoletaImp, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, , , oDocRec.fsIFTpo, oDocRec.fsIFCta) 'EJVG20140408
            Else
                If cboTransferMoneda.Text <> "" Then
                    '***Modificado por ELRO el 20120726, según OYP-RFC024-2012
                    'nSaldo = clsCap.CapAbonoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, IIf(nOperacion = gCTSDepTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), nPorcDisp, , , , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.TitPersLavDinero, lnMovNroTransfer, Right(Me.cboTransferMoneda.Text, 3), gsCodCMAC, loLavDinero.BenPersLavDinero, lsmensaje, lsBoletaImp, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
                    nSaldo = clsCap.CapAbonoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, IIf(nOperacion = gCTSDepTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), nPorcDisp, , , , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.TitPersLavDinero, lnMovNroTransfer, Right(Me.cboTransferMoneda.Text, 3), gsCodCMAC, loLavDinero.BenPersLavDinero, lsmensaje, lsBoletaImp, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, _
                                                      fnMovNroRVD, CCur(lblMonTra))
                    '***Modificado por ELRO el 20120726***********************
                Else
                    nSaldo = clsCap.CapAbonoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, IIf(nOperacion = gCTSDepTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), nPorcDisp, , , , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, loLavDinero.TitPersLavDinero, lnMovNroTransfer, , gsCodCMAC, loLavDinero.BenPersLavDinero, lsmensaje, lsBoletaImp, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
                End If
            End If
    End Select
    
    '*****BRGO 20110914 *****************************************************
    If gbITFAplica = True And CCur(lblITF.Caption) > 0 Then
       Call loMov.InsertaMovRedondeoITF(sMovNro, 1, CCur(lblITF.Caption) + nRedondeoITF, CCur(lblITF.Caption)) 'BRGO 20110914
    End If
    Set loMov = Nothing
    '*** End BRGO *****************
    
    If lbVistoVal Then loVistoElectronico.RegistraVistoElectronico (gnMovNro) 'JUEZ 20141014
    
    'ALPA 20081010
    If gnMovNro > 0 Then
        ''Added by TORE: Correccion comision para Cajeros Corresponsales
        Dim clsCCLog As COMNCaptaGenerales.NCOMCaptaMovimiento
        Set clsCCLog = New COMNCaptaGenerales.NCOMCaptaMovimiento
        Dim bInsertLog As Boolean
        bInsertLog = clsCCLog.ValidaCuentaCC(txtCuenta.NroCuenta, nOperacion, nMonto, sMovNro, nComisionLog, 2)
        Set clsCCLog = Nothing
        'End Added
    
        'Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
        Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
        Call clsCap.InsertarClienteWesterDeposito(, , gnMovNro, txtClienteWesterUnion.Text, bCuentaWestern) 'ALPA 20110316
    End If
    'JACA 20110510***********************************************************
    'WIOR 20130301 ************************************************************
    If fbPersonaReaAhorros And gnMovNro > 0 Then
        frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(sCuenta), fnCondicion
        fbPersonaReaAhorros = False
    End If
    'WIOR FIN *****************************************************************
   
 '***************** ANDE 20170623 Modificación para mejora en calidad de validación de CIIU
    'Dim objPersona As COMDPersona.DCOMPersonas
    Dim nAcumulado As Currency
    Dim nMontoPersOcupacion As Currency
                 
    'Set objPersona = New COMDPersona.DCOMPersonas
                
    Set clsTC = New COMDConstSistema.NCOMTipoCambio
    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
    Set clsTC = Nothing
                            
    nAcumulado = objPersona.ObtenerPersAcumuladoMontoOpe(nTC, Mid(Format(gdFecSis, "yyyymmdd"), 1, 6), rsPersOcu!cperscod)
    nMontoPersOcupacion = objPersona.ObtenerParamPersAgeOcupacionMonto(Mid(rsPersOcu!cperscod, 4, 2), CInt(Mid(rsPersOcu!cPersCIIU, 2, 2)))
                
    If nAcumulado >= nMontoPersOcupacion Then
        If Not objPersona.ObtenerPersonaAgeOcupDatos_Verificar(rsPersOcu!cperscod, gdFecSis) Then
            objPersona.insertarPersonaAgeOcupacionDatos gnMovNro, rsPersOcu!cperscod, IIf(nMoneda = gMonedaNacional, lblTotal, lblTotal * nTC), nAcumulado, gdFecSis, sMovNro
        End If
    End If
    '********************* end ande
        
    'JACA END*****************************************************************
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation
     End If
    
    If Trim(lsBoletaImp) <> "" Then ImprimeBoleta lsBoletaImp
    If Trim(lsBoletaImpITF) <> "" Then ImprimeBoleta lsBoletaImpITF, "Boleta ITF"
    
    '***Agregado por ELRO el 20120718, según OYP-RFC024-2012
    'If nOperacion = gAhoDepTransf Or nOperacion = gCTSDepTransf Then
    If nOperacion = gAhoDepTransf Or nOperacion = gCTSDepTransf Or nOperacion = "200245" Then 'EJVG20130923
        If Mid(sCuenta, 9, 1) <> Trim(Right(cboTransferMoneda, 3)) Then
          MsgBox "Coloque papel para la Boleta de Compra/Venta Moneda Extranjera.", vbInformation, "Aviso"
          lsBoletaCVME = oNCOMContImprimir.ImprimeBoletaCompraVentaME("Compra/Venta Moneda Extranjera", "", _
                                                                      fsPersNombreCVME, _
                                                                      fsPersDireccionCVME, _
                                                                      fsdocumentoCVME, _
                                                                      IIf(Trim(Right(cboTransferMoneda, 3)) = Moneda.gMonedaExtranjera, CCur(lblTTCCD), CCur(lblTTCVD)), _
                                                                      IIf(Trim(Right(cboTransferMoneda, 3)) = Moneda.gMonedaExtranjera, gOpeCajeroMECompra, gOpeCajeroMEVenta), _
                                                                      CCur(lblMonTra), _
                                                                      CCur(txtMonto), _
                                                                      gsNomAge, _
                                                                      sMovNro, _
                                                                      sLpt, _
                                                                      gsCodCMAC, _
                                                                      gsNomCmac, _
                                                                      gbImpTMU)
          Do
           If Trim(lsBoletaCVME) <> "" Then
              nFicSal = FreeFile
              Open sLpt For Output As nFicSal
                 Print #nFicSal, lsBoletaCVME
                 Print #nFicSal, ""
              Close #nFicSal
            End If
            
        Loop Until MsgBox("¿Desea reimprimir Boleta de Compra/Venta Moneda Extranjera? ", vbQuestion + vbYesNo, Me.Caption) = vbNo
    End If
    fsPersNombreCVME = ""
    fsPersDireccionCVME = ""
    fsdocumentoCVME = ""
    lsBoletaCVME = ""
  End If
  '***Fin Agregado por ELRO el 20120718*******************
    
    sNumTarj = ""
    
    Set clsLav = Nothing
    Set clsCap = Nothing
    Set rsAgeParam = Nothing
    cmdCancelar_Click
    
    gVarPublicas.LimpiaVarLavDinero
    
    '***Agregado por ELRO el 20120718, según OYP-RFC024-2012
    fnMovNroRVD = 0
    lblMonTra = "0.00"
    '***Fin Agregado por ELRO el 20120718*******************
End If
  
Set loLavDinero = Nothing
Set oNCOMContImprimir = Nothing '***Agregado por ELRO el 20120717, según OYP-RFC024-2012


  '************ Registrar actividad de opertaciones especiales - ANDE 2017-12-18
    'Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim RVerOpe As ADODB.Recordset
    Dim nEstadoActividad As Integer
    nEstadoActividad = oCaptaLN.RegistrarActividad(nOperacion, gsCodUser, gdFecSis)
    
    If nEstadoActividad = 1 Then
        MsgBox "He detectado un problema; su operación no fue afectada, pero por favor comunciar a TI-Desarrollo.", vbError, "Error"
    ElseIf nEstadoActividad = 2 Then
        MsgBox "Ha usado el total de operaciones permitidas para el día de hoy. Si desea realizar más operaciones, comuníquese con el área de Operaciones.", vbInformation + vbOKOnly, "Aviso"
        Unload Me
    End If
    ' END ANDE ******************************************************************

    'INICIO JHCU ENCUESTA 16-10-2019
    Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
    'FIN
Exit Sub
ErrGraba:
    MsgBox err.Description, vbExclamation, "Error"
    Exit Sub
End Sub

'RIRO20170714 ***
Private Function validaVoucher() As String
        
    Dim rsVoucher As ADODB.Recordset
    Dim bVerifica As Boolean
    Dim obj As COMNCaptaGenerales.NCOMCaptaGenerales
    
    On Error GoTo ErrValida
    
    validaVoucher = ""
    bVerifica = False
    Set obj = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsVoucher = obj.obtenerVoucherClienteNroIFSinOperacion(Right(gsCodAge, 2), _
                                                               Trim(Right(cboTransferMoneda.Text, 3)), 2, _
                                                               Left(sIFTipo, 2), Mid(sIFTipo, 4, 13), _
                                                               Trim(Right(lblTrasferND.Caption, 50)), _
                                                               Trim(Left(lblTrasferND.Caption, 50)))
    If (Not rsVoucher Is Nothing) Then
        If rsVoucher.RecordCount > 0 Then
            If (Not rsVoucher.EOF And Not rsVoucher.BOF) Then
                bVerifica = True
            End If
        End If
    End If
    
    If Not bVerifica Then
        validaVoucher = "El voucher seleccionado ya no se encuentra disponible"
    End If
    
    Exit Function
ErrValida:
    validaVoucher = "Se presentó un inconveniente durante la validacion del estado del voucher"
    Exit Function
End Function
'END RIRO *******

Private Sub cmdSalir_Click()
Unload Me
End Sub
'EJVG20130914 ***
'Private Sub cmdTranfer_Click()
'    Dim lsGlosa As String
'    Dim lsDoc As String
'    Dim lsInstit As String
'    Dim oform As New frmCapRegVouDepBus '***Agregado por ELRO el 20120706, según OYP-RFC024-2012
'    Dim lnTipMot As Integer '***Agregado por ELRO el 20120706, según OYP-RFC024-2012
'    Dim i As Integer '***Agregado por ELRO el 20120706, según OYP-RFC024-2012
'
'    If Me.cboTransferMoneda.Text = "" Then
'        MsgBox "Debe escoger la moneda de la transferencia.", vbInformation, "Aviso"
'        cboTransferMoneda.SetFocus
'        Exit Sub
'    End If
'
'    '***Agregado por ELRO el 20120726, según OYP-RFC024-2012
'    If gsOpeCod = gAhoDepTransf Then
'        lnTipMot = 2
'    ElseIf gsOpeCod = gCTSDepTransf Then
'        lnTipMot = 6
'    End If
'    '***Fin Agregado por ELRO el 20120726*******************
'
'    '***Modificado por ELRO el 20120718, según OYP-RFC024-2012
'    'lnMovNroTransfer = frmTransfpendientes.Ini(Right(Me.cboTransferMoneda.Text, 2), lnTransferSaldo, lsGlosa, lsInstit, lsDoc)
'    oform.iniciarFormularioDeposito Trim(Right(cboTransferMoneda, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer, fsPersNombreCVME, fsPersDireccionCVME, fsdocumentoCVME
'    '***Fin Modificado por ELRO el 20120718*******************
'
'    '***Comentado por ELRO el 20120718, según OYP-RFC024-2012
'    'If lnMovNroTransfer = -1 Then
'    '    Me.cboTransferMoneda.Enabled = True
'    '    lnTransferSaldo = 0
'    'Else
'    '    Me.cboTransferMoneda.Enabled = False
'    'End If
'    '***Fin Comentado por ELRO el 20120718******************
'
'    Me.txtTransferGlosa.Text = lsGlosa
'    Me.lbltransferBco.Caption = lsInstit
'    Me.lblTrasferND.Caption = lsDoc
'
'    Me.txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")
'
'    If Mid(txtCuenta.NroCuenta, 9, 1) = Moneda.gMonedaNacional Then
'        If Right(cboTransferMoneda, 3) = Moneda.gMonedaNacional Then
'            Me.txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")
'        Else
'            Me.txtMonto.Text = Format(lnTransferSaldo * CCur(Me.lblTTCCD.Caption), "#,##0.00")
'        End If
'    Else
'        If Right(cboTransferMoneda, 3) = Moneda.gMonedaNacional Then
'            Me.txtMonto.Text = Format(lnTransferSaldo / CCur(Me.lblTTCVD.Caption), "#,##0.00")
'        Else
'            Me.txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")
'        End If
'    End If
'
'    txtMonto_Change '***Agregado por ELRO el 20120726, según OYP-RFC024-2012
'
'    '***Modificado por ELRO el 20120706, según OYP-RFC024-2012
'    'If lnMovNroTransfer <> -1 Then
'    '    Me.txtTransferGlosa.SetFocus
'    'End If
'    If lnTransferSaldo > 0# Then
'        cboTransferMoneda.Enabled = False
'    Else
'        cboTransferMoneda.Enabled = True
'    End If
'    txtTransferGlosa.Locked = True
'    txtMonto.Enabled = False
'    lblMonTra = Format(lnTransferSaldo, "#,##0.00")
'    '***Fin Modificado por ELRO el 20120706*******************
'End Sub
Private Sub cmdTranfer_Click()
    Dim lsGlosa As String
    Dim lsDoc As String
    Dim lsInstit As String
    Dim oform As frmCapRegVouDepBus
    Dim lnTipMot As Integer
    Dim i As Integer
    
    If Len(txtCuenta.NroCuenta) <> 18 Then
        MsgBox "Ud. debe especificar el Nro. de Cuenta", vbInformation, "Aviso"
        If txtCuenta.Visible And txtCuenta.Enabled Then txtCuenta.SetFocusCuenta
        Exit Sub
    End If
    If cboTransferMoneda.Text = "" Then
        MsgBox "Ud. debe seleccionar la moneda de la transferencia.", vbInformation, "Aviso"
        If cboTransferMoneda.Visible And cboTransferMoneda.Enabled Then cboTransferMoneda.SetFocus
        Exit Sub
    End If

    If gsOpeCod = gAhoDepTransf Or gsOpeCod = "200245" Then
        lnTipMot = 2
    ElseIf gsOpeCod = gCTSDepTransf Then
        lnTipMot = 6
    End If

    Set oform = New frmCapRegVouDepBus
    oform.iniciarFormularioDeposito Trim(Right(cboTransferMoneda.Text, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer, fsPersNombreCVME, fsPersDireccionCVME, fsdocumentoCVME, txtCuenta.NroCuenta, sIFTipo 'RIRO20171007
    
    txtTransferGlosa.Text = lsGlosa
    lbltransferBco.Caption = lsInstit
    lblTrasferND.Caption = lsDoc
    
    txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")
    
    If Mid(txtCuenta.NroCuenta, 9, 1) = Moneda.gMonedaNacional Then
        If Right(cboTransferMoneda.Text, 3) = Moneda.gMonedaNacional Then
            txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")
        Else
            txtMonto.Text = Format(lnTransferSaldo * CCur(lblTTCCD.Caption), "#,##0.00")
        End If
    Else
        If Right(cboTransferMoneda.Text, 3) = Moneda.gMonedaNacional Then
            txtMonto.Text = Format(lnTransferSaldo / CCur(lblTTCVD.Caption), "#,##0.00")
        Else
            txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")
        End If
    End If
    
    txtMonto_Change

    txtTransferGlosa.Locked = True
    txtMonto.Enabled = False
    lblMonTra = Format(lnTransferSaldo, "#,##0.00")
    Set oform = Nothing
End Sub
'END EJVG *******
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sNumTar As String
    Dim sClaveTar As String
    Dim nErr As Integer
    Dim sCaption As String
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim nEstado  As COMDConstantes.CaptacTarjetaEstado
    Dim sMaquina As String
    Dim nCOM As Integer
    sMaquina = GetComputerName

    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Set clsGen = New COMDConstSistema.DCOMGeneral
        Dim bRetSinTarjeta As Boolean
        Dim sCuenta As String
        bRetSinTarjeta = clsGen.GetPermisoEspecialUsuario(gCapPermEspRetSinTarj, gsCodUser, gsDominio)
        sCuenta = frmValTarCodAnt.Inicia(nProducto, bRetSinTarjeta, True)
        If Val(Mid(sCuenta, 6, 3)) <> nProducto Then
            MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
    
    If KeyCode = vbKeyF11 And txtCuenta.Enabled = True Then 'F11
        sCaption = Me.Caption
        Me.Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
        sNumTar = ""
        sNumTar = GetNumTarjeta_ACS
        If Len(sNumTar) <> 16 Then
            MsgBox "Error en la Lectura de Tarjeta.", vbInformation, "Aviso"
            Me.Caption = sCaption
            Exit Sub
        End If
        
        Me.Caption = "Ingrese la Clave de la Tarjeta."
        MsgBox "Ingrese la Clave de la Tarjeta.", vbInformation, "AVISO"
        Select Case GetClaveTarjeta_ACS(sNumTar, 1)
            Case gClaveValida
                    Dim rsTarj As ADODB.Recordset
                    Set rsTarj = New ADODB.Recordset
                    Dim ObjTarj As COMNCaptaServicios.NCOMCaptaTarjeta
                    Set ObjTarj = New COMNCaptaServicios.NCOMCaptaTarjeta
                    Set rsTarj = ObjTarj.Get_Datos_Tarj(sNumTar)
                    
                    If rsTarj.EOF And rsTarj.BOF Then
                        MsgBox "Tarjeta no posee ninguna relación con cuentas activas o Tarjeta no activa.", vbInformation, "Aviso"
                        Set ObjTarj = Nothing
                        Me.Caption = sCaption
                        Exit Sub
                    Else
                        nEstado = rsTarj("nEstado")
                        If nEstado = gCapTarjEstBloqueada Or nEstado = gCapTarjEstCancelada Then
                            If nEstado = gCapTarjEstBloqueada Then
                                MsgBox "Número de Tarjeta Bloqueada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
                            ElseIf nEstado = gCapTarjEstCancelada Then
                                MsgBox "Número de Tarjeta Cancelada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
                            End If
                            Me.Caption = sCaption
                            Set ObjTarj = Nothing
                            Exit Sub
                        End If
                        
                        Dim rsPers As New ADODB.Recordset
                        Dim sCta As String, sProducto As String, sMoneda As String
                        Dim clsCuenta As UCapCuenta
                                                
                        Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
                        'Dim rsTarj As New ADODB.Recordset
    
                        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                        'Set rsTarj = clsMant.GetTarjetaCuentas(sNumTar, nProducto)
                        
                        Set rsPers = clsMant.GetCuentasPersona(rsTarj("cPersCod"), nProducto)
                        'Set rsPers = clsMant.GetTarjetaCuentas(sNumTar, nProducto)
                        
                        Set clsMant = Nothing
                        If Not (rsPers.EOF And rsPers.EOF) Then
                            Do While Not rsPers.EOF
                                sCta = rsPers("cCtaCod")
                                sProducto = rsPers("cDescripcion")
                                sMoneda = Trim(rsPers("cMoneda"))
                                frmCapMantenimientoCtas.lstCuentas.AddItem sCta & space(2) & sProducto & space(2) & sMoneda
                                rsPers.MoveNext
                            Loop
                            
                            Set clsCuenta = New UCapCuenta
                            Set clsCuenta = frmCapMantenimientoCtas.Inicia
                            If clsCuenta.sCtaCod <> "" Then
                                txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
                                txtCuenta.Prod = Mid(clsCuenta.sCtaCod, 6, 3)
                                txtCuenta.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
                                txtCuenta.SetFocusCuenta
                                Call txtCuenta_KeyPress(13)
                                'SendKeys "{Enter}"
                            End If
                            Set clsCuenta = Nothing
                        Else
                            MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
                        End If
                        rsPers.Close
                        Set rsPers = Nothing
                    End If
                    Set rsTarj = Nothing
                    Set clsMant = Nothing
                                    
            Case gClaveNOValida
                MsgBox "Clave Incorrecta", vbInformation, "Aviso"
            Case Else
                MsgBox "Tarjeta no Registrada", vbInformation, "Aviso"
        End Select
        Set clsGen = Nothing
        Me.Caption = "Captaciones - Abono - Ahorros " & sOperacion
    End If
    
    '**DAOR 20081125, Para tarjetas ***********************
    If KeyCode = vbKeyF10 And txtCuenta.Enabled Then
        sCuenta = frmATMCargaCuentas.RecuperaCuenta(gsOpeCod, sNumTarj, nProducto)
        If Val(Mid(sCuenta, 6, 3)) <> nProducto And sCuenta <> "" Then
            MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
    '*******************************************************
    
End Sub

Private Sub Form_Load()
    GetTipCambio gdFecSis
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    '***Modificado por ELRO el 20120828, según OYP-RFC024-2012
    'Me.lblTTCCD.Caption = Format(gnTipCambioC, "#.00")
    Me.lblTTCCD.Caption = Format(gnTipCambioC, "#,#0.0000")
    'Me.lblTTCVD.Caption = Format(gnTipCambioV, "#.00")
    Me.lblTTCVD.Caption = Format(gnTipCambioV, "#,#0.0000")
    '***Fin Modificado por ELRO el 20120828*******************
    lnMovNroTransfer = -1
End Sub



Private Sub grdCliente_DblClick()
    Dim R As ADODB.Recordset
    Dim ssql As String
    Dim clsFirma As COMDCaptaGenerales.DCOMCaptaMovimiento 'DCapMovimientos
    Set clsFirma = New COMDCaptaGenerales.DCOMCaptaMovimiento
     
      If Me.grdCliente.TextMatrix(grdCliente.row, 1) = "" Then Exit Sub
    
   Set R = New ADODB.Recordset
    Set R = clsFirma.GetFirma(Me.grdCliente.TextMatrix(grdCliente.row, 1), gsCodAge)
   If R.BOF Or R.EOF Then
       Set R = Nothing
       MsgBox "La visualización del DNI no esta Disponible", vbOKOnly + vbInformation, "AVISO"
       Exit Sub
   End If
            
    If R.RecordCount > 0 Then
       If IsNull(R!iPersFirma) = True Then
           MsgBox "El cliente no posse Firmas ", vbInformation, "Aviso"
           Exit Sub
       End If
       ' Call frmMuestraFirma.IDBFirma.CargarFirma(R)
       frmMuestraFirma.psCodCli = Me.grdCliente.TextMatrix(grdCliente.row, 1)
       Set frmMuestraFirma.rs = R
    End If
    Set clsFirma = Nothing
    frmMuestraFirma.Show 1
End Sub




Private Sub txtClienteWesterUnion_EmiteDatos()
    Me.lblClienteWesterUnion.Caption = Trim(txtClienteWesterUnion.psDescripcion)
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCta As String
    sCta = txtCuenta.NroCuenta
      
    
    If Trim(nOperacion) = "200243" Then
        If EsHaberes(sCta) Then
                ObtieneDatosCuenta sCta
        Else
            MsgBox "Esta no es una cuenta de haberes", vbOKOnly + vbExclamation, App.Title
        End If
    ElseIf Trim(nOperacion) = "200244" Then
        If EsHaberes(sCta) Then
                ObtieneDatosCuenta sCta
        Else
            MsgBox "Esta no es una cuenta de haberes", vbOKOnly + vbExclamation, App.Title
        End If
    ElseIf Trim(nOperacion) = "200245" Then
            If EsHaberes(sCta) Then
                ObtieneDatosCuenta sCta
            Else
                MsgBox "Esta no es una cuenta de haberes", vbOKOnly + vbExclamation, App.Title
            End If
    Else
'APRI20190109 ERS077-2018 LIBERACIÓN CAJA SUELDO
'        If EsHaberes(sCta) Then
'                MsgBox "No puede utilizar esta Operación para una Cuenta de Haberes", vbOKOnly + vbExclamation, App.Title
'                Exit Sub
'        End If
        
            ObtieneDatosCuenta sCta
     End If
     'frmSegSepelioAfiliacion.Inicio sCta 'RECO20151226 ERS074-2015 'COMENTADO POR APRI20180201 ERS028-2017
End If
        'ALPA 20081006**************************************************************************
            Dim loConstSis As COMDConstSistema.NCOMConstSistema
            Set loConstSis = New COMDConstSistema.NCOMConstSistema
            bCuentaWestern = False
            If Mid(sCta, 9, 1) = "1" Then
                lsCuentaWestern = loConstSis.LeeConstSistema(334)
                bCuentaWestern = True
            Else
                lsCuentaWestern = loConstSis.LeeConstSistema(333)
                bCuentaWestern = True
            End If
            If sCta = lsCuentaWestern Then
                txtClienteWesterUnion.Enabled = True
                txtClienteWesterUnion.EnabledText = True
                lblClienteWesterUnion.Enabled = True
                bCuentaWestern = True
            Else
                txtClienteWesterUnion.Enabled = False
                txtClienteWesterUnion.EnabledText = False
                lblClienteWesterUnion.Enabled = False
                bCuentaWestern = False
            End If
        '***************************************************************************************
        
        '***********APRI 20161031
        'COMENTADO APRI 20170623-PENDIENTE PASE A PRODUCCION TI-ERS048-2016
'        If nOperacion = "200201" Or nOperacion = "200202" Or nOperacion = "200203" Then
'            Dim CTA As String
'            CTA = txtCuenta.NroCuenta
'            BusquedaProducto "", CTA, nOperacion
'        End If
        '***********END APRI

End Sub
Private Function EsHaberes(ByVal sCta As String) As Boolean
Dim ssql As String
Dim cCap As COMDCaptaGenerales.COMDCaptAutorizacion
Set cCap = New COMDCaptaGenerales.COMDCaptAutorizacion
    EsHaberes = cCap.EsHaberes(sCta)
Set cCap = Nothing
End Function

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    If fraPeriodoCTS.Visible Then
        cboPeriodo.SetFocus
    Else
    'EJVG20140206 ***
    '    txtMonto.Enabled = True
    '    txtMonto.SetFocus
    If txtMonto.Visible And txtMonto.Enabled Then txtMonto.SetFocus
    'END EJVG *******
    End If
End If
End Sub

Private Sub txtMonto_Change()
If nOperacion <> "200243" And nOperacion <> "200244" And nOperacion <> "200245" Then
    If gbITFAplica And nProducto <> gCapCTS Then       'Filtra para CTS
        If txtMonto.value > gnITFMontoMin Then
            If Not lbITFCtaExonerada Then
                'If nOperacion = gAhoDepTransf Or nOperacion = gAhoDepPlanRRHH Or nOperacion = gAhoDepPlanRRHHAdelantoSueldos Or nOperacion = gAhoDepOtrosIngRRHH Or nOperacion = gAhoDepGratRRHH Then
                If nOperacion = gAhoDepPlanRRHH Or nOperacion = gAhoDepPlanRRHHAdelantoSueldos Or nOperacion = gAhoDepGratRRHH Then
                    Me.lblITF.Caption = Format(0, "#,##0.00")
                ElseIf nProducto = gCapAhorros And (nOperacion <> gAhoDepChq Or nOperacion <> gCMACOAAhoDepChq Or nOperacion <> gCMACOTAhoDepChq) Then
                    Me.lblITF.Caption = Format(fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                Else
                    Me.lblITF.Caption = Format(fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                End If
                nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
                If nRedondeoITF > 0 Then
                    Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
                End If
                '*** END BRGO
                If bInstFinanc Then lblITF.Caption = "0.00" 'JUEZ 20140414
            Else
                Me.lblITF.Caption = "0.00"
            End If
            If (nProducto = gCapAhorros And gbITFAsumidoAho = False) Or (nProducto = gCapPlazoFijo And gbITFAsumidoPF = False) Then
                If nOperacion = gAhoRetOPCanje Or nOperacion = gAhoRetOPCertCanje Or nOperacion = gAhoRetFondoFijoCanje Then
                    Me.lblTotal.Caption = Format(0, "#,##0.00")
                ElseIf nOperacion = gAhoDepChq Then
                     If chkITFEfectivo.value = vbChecked Then
                        Me.lblTotal.Caption = Format(CCur(txtMonto.Text) + CCur(Me.lblITF.Caption), "#,##0.00")
                        Exit Sub
                    Else
                        Me.lblTotal.Caption = Format(CCur(txtMonto.Text), "#,##0.00") ' + CCur(lblITF.Caption)
                        Exit Sub
                    End If
                End If
            Else
                If nProducto = gCapAhorros And gbITFAsumidoAho Then
                    Me.lblITF.Caption = Format(fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                    Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
                ElseIf nProducto = gCapPlazoFijo And gbITFAsumidoPF Then
                    Me.lblITF.Caption = Format(fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                    Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
                Else
                    Me.lblTotal.Caption = Format(txtMonto.value - CCur(Me.lblITF.Caption), "#,##0.00")
                End If
            
            End If
            If bInstFinanc Then lblITF.Caption = "0.00" 'JUEZ 20140414
        End If
    Else
        Me.lblITF.Caption = Format(0, "#,##0.00")
        
        If nProducto = gCTSDepChq Then
            Me.lblTotal.Caption = Format(0, "#,##0.00")
        Else
            Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00")
        End If
    End If
Else
    Me.lblITF.Caption = Format(0, "#,##0.00")
    Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00")
End If
    
    If txtMonto.value = 0 Then
        Me.lblITF.Caption = "0.00"
        Me.lblTotal.Caption = "0.00"
    End If
    chkITFEfectivo_Click
End Sub

Private Sub txtMonto_GotFocus()
txtMonto.MarcaTexto
txtMonto.SelStart = 0
txtMonto.SelLength = Len(txtMonto.Text)

End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdgrabar.SetFocus
End If
End Sub

Private Sub IniciaCombo(ByRef cboConst As ComboBox, nCapConst As ConstanteCabecera)
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsConst As New ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(nCapConst)
    Set clsGen = Nothing
    Do While Not rsConst.EOF
        cboConst.AddItem rsConst("cDescripcion") & space(100) & rsConst("nConsValor")
        rsConst.MoveNext
    Loop
    cboConst.ListIndex = 0
End Sub

Private Sub txtTransferGlosa_GotFocus()
    txtTransferGlosa.SelStart = 0
    txtTransferGlosa.SelLength = 500
End Sub

Private Sub txtTransferGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    If fraPeriodoCTS.Visible Then
        cboPeriodo.SetFocus
    'EJVG20130914 ***
    'Else
    '    txtMonto.Enabled = True
    '    txtMonto.SetFocus
    'END EJVG *******
    End If
End If
End Sub

Private Function Cargousu(ByVal NomUser As String) As String
 Dim rs As New ADODB.Recordset
 Dim oCons As COMDConstSistema.DCOMUAcceso
 Set oCons = New COMDConstSistema.DCOMUAcceso
 
 Set rs = oCons.Cargousu(NomUser)
  If Not (rs.EOF And rs.BOF) Then
    Cargousu = rs(0)
  End If
 Set rs = Nothing
 'rs.Close
 Set oCons = Nothing
End Function

Private Sub IniciaLavDinero(ByRef poLavDin As frmMovLavDinero, Optional lsCuentaWestern As String = "", Optional lsNombreWestern As String = "", Optional bCuentaWestern As Boolean = False)
Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim nMonto As Double
Dim oPersona As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsPers As New ADODB.Recordset
If bCuentaWestern = False Then
For i = 1 To grdCliente.Rows - 1
    nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
    If nPersoneria = gPersonaNat Then
        If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            poLavDin.TitPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDin.TitPersLavDineroNom = grdCliente.TextMatrix(i, 2)
            Exit For
        End If
    Else
        If nRelacion = gCapRelPersTitular Then
            poLavDin.TitPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDin.TitPersLavDineroNom = grdCliente.TextMatrix(i, 2)
        End If
        If nRelacion = gCapRelPersRepTitular Then
            poLavDin.ReaPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDin.ReaPersLavDineroNom = grdCliente.TextMatrix(i, 2)
            If poLavDin.TitPersLavDinero <> "" Then Exit For
        End If
    End If
Next i
Else
     poLavDin.TitPersLavDinero = lsCuentaWestern
     poLavDin.TitPersLavDineroNom = lsNombreWestern
End If
nMonto = txtMonto.value
sTipoCuenta = lblTipoCuenta.Caption

End Sub
'***** BRGO 20110127 ********* COBRO DE COMISIÓN POR DEPÓSITO OTRA AGENCIA ************
Public Function CalcularComisionDepOtraAge() As Currency
    
    'Comentado by JACA 20111021*******************************
        'Dim nComixDep As Double
        'Dim nPorComixDep As Double
        'Dim nMontoComixDepAge As Double
        Dim oCons As COMDConstantes.DCOMAgencias
        'Dim clsGen As COMNCaptaGenerales.NCOMCaptaDefinicion
        'Dim clsTC As COMDConstSistema.NCOMTipoCambio 'nTipoCambio
        
        Set oCons = New COMDConstantes.DCOMAgencias
     'JACA EN**************************************************
   
        '*** Verificar Ubicacion de la Cuenta ***
        If oCons.VerficaZonaAgencia(gsCodAge, Mid(txtCuenta.NroCuenta, 4, 2)) Then
                'Comentado by JACA 20111021*******************************
            
                    'Set clsGen = New COMNCaptaGenerales.NCOMCaptaDefinicion
                
                    'If nmoneda = gMonedaNacional Then
                    '    nMontoComixDepAge = clsGen.GetCapParametro(2046) 'Monto comisión en soles
                    'Else
                    '    nMontoComixDepAge = clsGen.GetCapParametro(2047) 'Monto comisión en dólares
                    'End If
                    
                    'Valida si en la operación no interviene las Agencias Requena o Aguaytía
                    'If (Mid(txtCuenta.NroCuenta, 4, 2) <> 12 And Mid(txtCuenta.NroCuenta, 4, 2) <> 13) And (gsCodAge <> 12 And gsCodAge <> 13) Then
                    '    nComixDep = nMontoComixDepAge
                    'Else
                    '    If (gsCodAge = 12 Or gsCodAge = 13) And (Mid(txtCuenta.NroCuenta, 4, 2) <> 12 And Mid(txtCuenta.NroCuenta, 4, 2) <> 13) Then
                    '        nComixDep = nMontoComixDepAge
                    '    Else
                    '        nPorComixDep = clsGen.GetCapParametro(gComisionRetOtraAge)
                    '        nComixDep = CDbl(txtMonto.Text) * (val(nPorComixDep) / 100)
                    '        If nComixDep < nMontoComixDepAge Then
                    '            nComixDep = nMontoComixDepAge
                    '        End If
                    '    End If
                    'End If
                    
                    'Set clsGen = Nothing
                'JACA END****************************************************
                
                'JACA 20111021*******************************************************
                
                Dim objCap As New COMNCaptaGenerales.NCOMCaptaGenerales
                Dim rs As New Recordset
                
                'Obtenemos los parametros de la comision
                'Set rs = objCap.obtenerValTarifaOpeEnOtrasAge(gsCodAge, lnTpoPrograma, CInt(Mid(Me.txtCuenta.NroCuenta, 9, 1)))
                Set rs = objCap.obtenerValTarifaOpeEnOtrasAge(Me.txtCuenta.NroCuenta) 'APRI20190109 ERS077-2018
                If Not (rs.EOF And rs.BOF) Then
                    
                    If rs!nComision > 0 Then
                        'si la comision es menor que el monto minimo
                        If CDbl(txtMonto.Text) * (rs!nComision / 100) < rs!nMontoMin Then
                            CalcularComisionDepOtraAge = rs!nMontoMin
                        'si la comision es mayor que el monto maximo
                        ElseIf CDbl(txtMonto.Text) * (rs!nComision / 100) > rs!nMontoMax Then
                            CalcularComisionDepOtraAge = rs!nMontoMax
                        'si la comision esta entre en monto min y max
                        Else
                            CalcularComisionDepOtraAge = CDbl(txtMonto.Text) * (rs!nComision / 100)
                        End If
                    Else
                        CalcularComisionDepOtraAge = 0
                    End If
                    
                Else
                     CalcularComisionDepOtraAge = 0
                End If
                'JACA END************************************************************
            
        Else
            'nComixDep = 0 Comentado by JACA 20111021
            CalcularComisionDepOtraAge = 0
        End If
    Set oCons = Nothing
    'CalcularComisionDepOtraAge = nComixDep Comentado by JACA 20111021
End Function
'***************RECO 20131024 ERS141************
Public Sub ActivarControlesDevGarantEcotaxi(ByVal pbValor As Boolean)
    If pbValor = True Then
        txtCuentaGarant.Enabled = True
        txtClienteGarant.Enabled = True
        cmdBuscaCredGarant.Enabled = True
        txtCuentaGarant.EnabledAge = True
        txtCuentaGarant.EnabledCMAC = True
        txtCuentaGarant.EnabledCta = True
        txtCuentaGarant.EnabledProd = True
    Else
        txtCuentaGarant.Enabled = False
        txtClienteGarant.Enabled = False
        cmdBuscaCredGarant.Enabled = False
        txtCuentaGarant.EnabledAge = False
        txtCuentaGarant.EnabledCMAC = False
        txtCuentaGarant.EnabledCta = False
        txtCuentaGarant.EnabledProd = False
    End If
End Sub

Private Sub txtCuentaGarant_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCta As String
        
        Dim loCred As New COMDCredito.DCOMCreditos
        Set loCred = New COMDCredito.DCOMCreditos
        
        Dim loDR As ADODB.Recordset
        Set loDR = New ADODB.Recordset
        
        'Set loDR = loCred.ObtieneDatosCredEcotaxi(txtCuenta.NroCuenta)
        'sCta = txtCuenta.NroCuenta
        
        
        
        
        
    End If
End Sub
'******************END RECO*********************


