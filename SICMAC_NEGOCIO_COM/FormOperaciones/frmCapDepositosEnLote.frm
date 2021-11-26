VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCapDepositosEnLote 
   Caption         =   "Depositos en Lote"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13200
   Icon            =   "frmCapDepositosEnLote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   11160
      TabIndex        =   44
      Top             =   8160
      Width           =   915
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   12120
      TabIndex        =   43
      Top             =   8160
      Width           =   915
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   8160
      Width           =   915
   End
   Begin VB.Frame fraMonto 
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
      Height          =   2040
      Left            =   3960
      TabIndex        =   19
      Top             =   6000
      Width           =   9075
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL (S/.)"
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
         Height          =   195
         Left            =   1440
         TabIndex        =   35
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL ($.)"
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
         Height          =   195
         Left            =   2640
         TabIndex        =   34
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ITF Cargo Cta"
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
         Height          =   195
         Left            =   165
         TabIndex        =   33
         Top             =   870
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ITF Efectivo "
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
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label lblITFES 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1470
         TabIndex        =   31
         Top             =   495
         Width           =   960
      End
      Begin VB.Label lblITFCS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1470
         TabIndex        =   30
         Top             =   825
         Width           =   960
      End
      Begin VB.Label lblITFED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2685
         TabIndex        =   29
         Top             =   495
         Width           =   960
      End
      Begin VB.Label lblITFCD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2685
         TabIndex        =   28
         Top             =   825
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL NETO ( $ )"
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
         Height          =   195
         Left            =   4455
         TabIndex        =   27
         Top             =   855
         Width           =   1590
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL NETO (S/.)"
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
         Height          =   195
         Left            =   4455
         TabIndex        =   26
         Top             =   420
         Width           =   1635
      End
      Begin VB.Label lblITFAD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2685
         TabIndex        =   25
         Top             =   1170
         Width           =   960
      End
      Begin VB.Label lblITFAS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1470
         TabIndex        =   24
         Top             =   1170
         Width           =   960
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ITF Asumido"
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
         Height          =   195
         Left            =   165
         TabIndex        =   23
         Top             =   1170
         Width           =   1065
      End
      Begin VB.Label lblTNS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6210
         TabIndex        =   22
         Top             =   360
         Width           =   960
      End
      Begin VB.Label lblTND 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6210
         TabIndex        =   21
         Top             =   780
         Width           =   960
      End
   End
   Begin VB.Frame fraDocumento 
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2040
      Left            =   120
      TabIndex        =   13
      Top             =   6000
      Width           =   3615
      Begin VB.TextBox txtMonto 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   960
         TabIndex        =   52
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdDocumento 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         Picture         =   "frmCapDepositosEnLote.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   200
         TabIndex        =   53
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro Doc :"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   375
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   840
         Width           =   555
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
         Left            =   960
         TabIndex        =   16
         Top             =   765
         Width           =   2535
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
         Left            =   960
         TabIndex        =   15
         Top             =   360
         Width           =   1575
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
      ForeColor       =   &H00000080&
      Height          =   4065
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   12915
      Begin SICMACT.FlexEdit grdCuentaAbono 
         Height          =   3015
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   12660
         _ExtentX        =   22331
         _ExtentY        =   5318
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Cuenta-Titular-Monto S/.-Monto $-ITF S/.-ITF $-FItf-TipoItf-Total-bExonerada-FCTACARGO"
         EncabezadosAnchos=   "250-1800-3800-1400-1400-1200-1200-0-0-1400-0-0"
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
         ColumnasAEditar =   "X-X-X-3-4-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-R-R-C-C-L-C-C-C-C"
         FormatosEdit    =   "0-0-0-2-2-2-2-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblTotalMN 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5940
         TabIndex        =   40
         Top             =   3255
         Width           =   1395
      End
      Begin VB.Label lblTotalME 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   7320
         TabIndex        =   39
         Top             =   3255
         Width           =   1395
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2880
         TabIndex        =   38
         Top             =   3255
         Width           =   3075
      End
      Begin VB.Label LblITFME 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   9930
         TabIndex        =   37
         Top             =   3240
         Width           =   1245
      End
      Begin VB.Label LblITFMN 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   8685
         TabIndex        =   36
         Top             =   3240
         Width           =   1260
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
      ForeColor       =   &H00000080&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12900
      Begin VB.Frame fraGlosa 
         Caption         =   "Glosa"
         Height          =   795
         Left            =   9600
         TabIndex        =   50
         Top             =   1080
         Width           =   3210
         Begin VB.TextBox txtGlosa 
            Height          =   435
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            Top             =   240
            Width           =   3030
         End
      End
      Begin VB.Frame fraTipoCambio 
         Caption         =   "Tipo Cambio"
         Height          =   915
         Left            =   10200
         TabIndex        =   45
         Top             =   120
         Width           =   1575
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "TCC:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   49
            Top             =   270
            Width           =   360
         End
         Begin VB.Label lblTCC 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   660
            TabIndex        =   48
            Top             =   210
            Width           =   795
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "TCV:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   47
            Top             =   570
            Width           =   360
         End
         Begin VB.Label lblTCV 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   660
            TabIndex        =   46
            Top             =   510
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar"
         Height          =   375
         Left            =   7560
         TabIndex        =   41
         Top             =   1080
         Width           =   1035
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   1995
      End
      Begin VB.ComboBox cboTipoTasa 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   660
         Width           =   1995
      End
      Begin VB.ComboBox cboPrograma 
         Height          =   315
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   2835
      End
      Begin SICMACT.TxtBuscar txtInstitucion 
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Tasa :"
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   720
         Width           =   810
      End
      Begin VB.Label lblInst 
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
         Left            =   4380
         TabIndex        =   9
         Top             =   690
         Width           =   5655
      End
      Begin VB.Label lblInstEtq 
         AutoSize        =   -1  'True
         Caption         =   "Institución :"
         Height          =   195
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Sub Producto:"
         Height          =   195
         Left            =   3300
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lblTasa 
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
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1200
         TabIndex        =   6
         Top             =   1080
         Width           =   1905
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tasa EA (%) :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   960
      End
   End
   Begin MSComDlg.CommonDialog CdlgFile 
      Left            =   1050
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCapDepositosEnLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nProducto As COMDConstantes.Producto
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim bDocumento As Boolean
Dim nmoneda As COMDConstantes.Moneda
Dim nTipoTasa As COMDConstantes.CaptacTipoTasa
Public sCodIF As String
Public dFechaValorizacion As Date
Public lnDValoriza As Integer
Dim nTasaNominal As Double
Dim sError As String, sCodPers As String
Dim vCantITF As Double
Dim oDocRec As UDocRec 'EJVG20140408

Private Sub cmdDocumento_Click()
    'EJVG20140408 ***
    'frmCapAperturaListaChq.Inicia frmCapDepositosEnLote, nOperacion, nmoneda, nProducto
    Dim oform As New frmChequeBusqueda
    Dim lnOperacion As TipoOperacionCheque

    On Error GoTo ErrCargaDocumento
    If nOperacion = gAhoDepositoEnLoteCheq Then
        lnOperacion = AHO_DepositoLote
    Else
        lnOperacion = Ninguno
    End If

    Set oDocRec = oform.Iniciar(nmoneda, lnOperacion)
    Set oform = Nothing
    
    txtGlosa.Text = oDocRec.fsGlosa
    lblNombreIF.Caption = oDocRec.fsPersNombre
    lblNroDoc.Caption = oDocRec.fsNroDoc
    sCodIF = oDocRec.fsPersCod
    txtMonto.Text = Format(oDocRec.fnMonto, gsFormatoNumeroView)
    txtMonto.Locked = True
    
    If nmoneda = gMonedaNacional Then
        lblTNS.Caption = txtMonto.Text
    Else
        lblTND.Caption = txtMonto.Text
    End If
    Exit Sub
ErrCargaDocumento:
    MsgBox "Ha sucedido un error al cargar los datos del Documento", vbCritical, "Aviso"
    'END EJVG *******
End Sub

Private Sub CmdGrabar_Click()
Dim nMontoCargo As Double
Dim sCuenta As String, sGlosa As String

Dim lsMensaje As String
Dim lsBoleta As String
Dim lsBoletaITF As String

Dim nFicSal As Integer

Dim Autid As Long

'nMontoCargo = txtMontoCargo.value
'sCuenta = txtCuenta.NroCuenta

If lblTotalMN = "" Or lblTotalME = "" Then
    MsgBox "Debe ingresar cuenta(s) para el abono", vbInformation, "Aviso"
    'cmdAgregar.SetFocus
    Exit Sub
End If

'If nMontoCargo = 0 Then
'    MsgBox "Monto de Cargo debe ser mayor a cero", vbInformation, "Aviso"
'    txtMontoCargo.SetFocus
'    Exit Sub
'End If
'If nmoneda = gMonedaNacional Then
'    If nMontoCargo <> CDbl(lblTotalMN) Then
'        MsgBox "Suma total no coincide como monto de cargo", vbInformation, "Aviso"
'        cmdAgregar.SetFocus
'        Exit Sub
'    End If
'Else
'    If nMontoCargo <> CDbl(lblTotalME) Then
'        MsgBox "Suma total no coincide como monto de cargo", vbInformation, "Aviso"
'        cmdAgregar.SetFocus
'        Exit Sub
'    End If
'End If

If MsgBox("¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sMovNro As String
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim rsCtaAbo As ADODB.Recordset
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    On Error GoTo ErrGraba
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set rsCtaAbo = grdCuentaAbono.GetRsNew()
    sGlosa = Trim(txtGlosa.Text)
    
    
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios, sPersLavDinero As String
    Dim nMontoLavDinero As Double, nTC As Double, sReaPersLavDinero As String, sBenPersLavDinero As String
    

    'Realiza la Validación para el Lavado de Dinero
    'sCuenta = txtCuenta.NroCuenta
    Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
    'If clsLav.EsOperacionEfectivo(Trim(nOperacion)) Then
        Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
'        If Not clsExo.EsCuentaExoneradaLavadoDinero(sCuenta) Then
'            Set clsExo = Nothing
'            sPersLavDinero = ""
'            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
'            Set clsLav = Nothing
'            If nmoneda = gMonedaNacional Then
'                Dim clsTC As COMDConstSistema.NCOMTipoCambio
'                Set clsTC = New COMDConstSistema.NCOMTipoCambio
'                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
'                Set clsTC = Nothing
'            Else
'                nTC = 1
'            End If
'            If nMontoCargo >= Round(nMontoLavDinero * nTC, 2) Then
'                sPersLavDinero = IniciaLavDinero()
'                sPersLavDinero = frmMovLavDinero.OrdPersLavDinero
'                sReaPersLavDinero = gVarPublicas.gReaPersLavDinero
'                sBenPersLavDinero = gVarPublicas.gBenPersLavDinero
'
'                If sPersLavDinero = "" Then Exit Sub
'
'            End If
'        Else
'            Set clsExo = Nothing
'        End If
        Set clsExo = Nothing
        Set clsLav = Nothing
        If oDocRec Is Nothing Then Set oDocRec = New UDocRec 'EJVG20140408
    'If clsCap.CapAbonoLote(sCuenta, nMontoCargo, sMovNro, rsCtaAbo, sGlosa, gsNomAge, sLpt, sPersLavDinero, CDbl(Me.lblTCC.Caption), CDbl(Me.lblTCV.Caption), gbITFAplica, 0, gbITFAsumidoAho, 0, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, , , , , , , gnMovNro, nOperacion, lblNroDoc.Caption, sCodIF) Then
    If clsCap.CapAbonoLote(sCuenta, nMontoCargo, sMovNro, rsCtaAbo, sGlosa, gsNomAge, sLpt, sPersLavDinero, CDbl(Me.lblTCC.Caption), CDbl(Me.lblTCV.Caption), gbITFAplica, 0, gbITFAsumidoAho, 0, sBenPersLavDinero, lsMensaje, lsBoleta, lsBoletaITF, , , , , , , gnMovNro, nOperacion, oDocRec.fsNroDoc, oDocRec.fsPersCod, oDocRec.fnTpoDoc, oDocRec.fsIFTpo, oDocRec.fsIFCta) Then 'EJVG20140408
     If gnMovNro > 0 Then
        'Call frmMovLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, sBenPersLavDinero, , , , , , gnTipoREU, gnMontoAcumulado, gsOrigen)
         Call frmMovLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, sBenPersLavDinero, , , , , , gnTipoREU, gnMontoAcumulado, gsOrigen) 'JACA 20110224
     End If
     
      If Trim(lsMensaje) <> "" Then
        MsgBox lsMensaje, vbInformation, "Aviso"
      End If
      
      If Trim(lsBoleta) <> "" Then
         nFicSal = FreeFile
         Open sLpt For Output As nFicSal
            Print #nFicSal, lsBoleta & Chr$(12)
            Print #nFicSal, ""
         Close #nFicSal
      End If
      
      If Trim(lsBoletaITF) <> "" Then
         nFicSal = FreeFile
         Open sLpt For Output As nFicSal
            Print #nFicSal, lsBoletaITF & Chr$(12)
            Print #nFicSal, ""
         Close #nFicSal
      End If
      cmdCancelar_Click
    Else
        MsgBox lsMensaje, vbInformation, "Aviso"
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
Private Sub cmdCancelar_Click()
LimpiaControles
End Sub
Private Sub LimpiaControles()

vCantITF = 0

'grdCliente.Clear
'grdCliente.Rows = 2
'grdCliente.FormaCabecera
grdCuentaAbono.Clear
grdCuentaAbono.Rows = 2
grdCuentaAbono.FormaCabecera
'txtMontoCargo.value = 0
cmdGrabar.Enabled = False
'txtCuenta.Age = ""
'txtCuenta.cuenta = ""
'txtCuentaAbo.Age = ""
'txtCuentaAbo.cuenta = ""
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
'fraCuentaAbono.Enabled = False
txtGlosa = ""
fraGlosa.Enabled = False
'txtCuenta.Enabled = True
'txtCuenta.SetFocus
'lblMensaje = ""
lblTotalMN = ""
lblTotalME = ""
'Me.lblITF = "0.00"
'Me.lblTotal = "0.00"
Me.lblITFAD = "0.00"
Me.lblITFAS = "0.00"
Me.lblITFED = "0.00"
Me.lblITFES = "0.00"
Me.lblITFCD = "0.00"
Me.lblITFCS = "0.00"

Me.lblTND.Caption = "0.00"
Me.lblTNS.Caption = "0.00"

Me.LblITFME.Caption = "0.00"
Me.LblITFMN.Caption = "0.00"
Me.lblNroDoc.Caption = ""
Me.lblNombreIF.Caption = ""
Me.txtMonto.Text = ""
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim clsGen As COMDConstSistema.NCOMTipoCambio
    Dim rsTC As ADODB.Recordset
    Set clsGen = New COMDConstSistema.NCOMTipoCambio
    lblTCC = Format$(clsGen.EmiteTipoCambio(gdFecSis, TCCompra), "#0.0000")
    lblTCV = Format$(clsGen.EmiteTipoCambio(gdFecSis, TCVenta), "#0.0000")
    Set clsGen = Nothing
End Sub

Private Sub grdCuentaAbono_OnCellChange(pnRow As Long, pnCol As Long)
Dim nMonCta As COMDConstantes.Moneda
Dim nMonto As Double
Dim nSumMontoSoles As Currency
Dim nSumMontoDolares As Currency
'Dim bExonerada As Boolean

nMonCta = CLng(Mid(grdCuentaAbono.TextMatrix(pnRow, 1), 9, 1))
nMonto = CDbl(grdCuentaAbono.TextMatrix(pnRow, pnCol))

'bExonerada = fgITFVerificaExoneracion(grdCuentaAbono.TextMatrix(pnRow, 1))

If pnCol = 3 Or pnCol = 4 Then
    If nmoneda = gMonedaNacional Then
        If nMonCta = nmoneda Then
            grdCuentaAbono.TextMatrix(pnRow, 4) = "0.00"
        Else
            If pnCol = 4 Then
                grdCuentaAbono.TextMatrix(pnRow, 3) = Format$(nMonto * CDbl(lblTCV), "#0.00")
            Else
                grdCuentaAbono.TextMatrix(pnRow, 4) = Format$(nMonto / CDbl(lblTCV), "#0.00")
            End If
        End If
    ElseIf nmoneda = gMonedaExtranjera Then
        If nMonCta = nmoneda Then
            grdCuentaAbono.TextMatrix(pnRow, 3) = "0.00"
        Else
            If pnCol = 4 Then
                grdCuentaAbono.TextMatrix(pnRow, 3) = Format$(nMonto * CDbl(lblTCC), "#0.00")
            Else
                grdCuentaAbono.TextMatrix(pnRow, 4) = Format$(nMonto / CDbl(lblTCC), "#0.00")
            End If
        End If
    End If
    If grdCuentaAbono.TextMatrix(pnRow, 7) = "N" Then
       If pnCol = 3 Then
            grdCuentaAbono.TextMatrix(pnRow, 5) = "0.00"
            grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 3)), "#,##0.00")
       Else
            grdCuentaAbono.TextMatrix(pnRow, 6) = "0.00"
            grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 4)), "#,##0.00")
       End If
       
       GoTo lblCalculos
       
    Else
    
'    If txtMontoCargo.value > gnITFMontoMin Then
'                        If Not lbITFCtaExonerada Then
'                            Me.LblItf.Caption = Format(fgITFCalculaImpuesto(txtMontoCargo.value), "#,##0.00")
'
'                        Else
'                            Me.LblItf.Caption = "0.00"
'                        End If
'
'                        If gbITFAsumidoAho Then
'                            Me.LblTotal.Caption = Format(txtMontoCargo.value, "#,##0.00")
'                            Exit Sub
'                        ElseIf chkItfEfectivo.value = vbChecked Then
'                            Me.LblTotal.Caption = Format(CCur(txtMontoCargo.Text) + CCur(Me.LblItf.Caption), "#,##0.00")
'                            Exit Sub
'                        Else
'                            Me.LblTotal.Caption = Format(CCur(txtMontoCargo.Text) - CCur(LblItf.Caption), "#,##0.00")
'                            Exit Sub
'                        End If
'    End If
    
    
    
    
       If pnCol = 3 Then
            grdCuentaAbono.TextMatrix(pnRow, 5) = Format(fgITFCalculaImpuesto(grdCuentaAbono.TextMatrix(pnRow, 3)), "#,##0.00")
          
          
            If grdCuentaAbono.TextMatrix(pnRow, 8) = "C" Then
          
                If grdCuentaAbono.TextMatrix(pnRow, 4) <> "" Or val(grdCuentaAbono.TextMatrix(pnRow, 4)) > 0 Then
                    grdCuentaAbono.TextMatrix(pnRow, 6) = Format(fgITFCalculaImpuesto(grdCuentaAbono.TextMatrix(pnRow, 4)), "#,##0.00")
                End If
                If nMonCta = gMonedaNacional Then
                      grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 3)) - grdCuentaAbono.TextMatrix(pnRow, 5), "#,##0.00")
                Else
                      grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 4)) - grdCuentaAbono.TextMatrix(pnRow, 6), "#,##0.00")
                End If
                
            ElseIf grdCuentaAbono.TextMatrix(pnRow, 8) = "E" Then
                
                If grdCuentaAbono.TextMatrix(pnRow, 4) <> "" Or val(grdCuentaAbono.TextMatrix(pnRow, 4)) > 0 Then
                    grdCuentaAbono.TextMatrix(pnRow, 6) = Format(fgITFCalculaImpuesto(grdCuentaAbono.TextMatrix(pnRow, 4)), "#,##0.00")
                End If
                
                If nMonCta = gMonedaNacional Then
                        grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 3)) + grdCuentaAbono.TextMatrix(pnRow, 5), "#,##0.00")
                Else
                        grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 4)) + grdCuentaAbono.TextMatrix(pnRow, 6), "#,##0.00")
                End If
            End If
            
            GoTo lblCalculos
            
       
       ElseIf pnCol = 4 Then
       
            grdCuentaAbono.TextMatrix(pnRow, 6) = Format(fgITFCalculaImpuesto(grdCuentaAbono.TextMatrix(pnRow, 4)), "#,##0.00")
       
            If grdCuentaAbono.TextMatrix(pnRow, 8) = "C" Then
               ' grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 4)) - grdCuentaAbono.TextMatrix(pnRow, 6), "#,##0.00")
                If grdCuentaAbono.TextMatrix(pnRow, 3) <> "" Or val(grdCuentaAbono.TextMatrix(pnRow, 4)) > 0 Then
                   grdCuentaAbono.TextMatrix(pnRow, 5) = Format(fgITFCalculaImpuesto(grdCuentaAbono.TextMatrix(pnRow, 3)), "#,##0.00")
                End If
                
                If nMonCta = gMonedaNacional Then
                      grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 3)) - grdCuentaAbono.TextMatrix(pnRow, 5), "#,##0.00")
                Else
                      grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 4)) - grdCuentaAbono.TextMatrix(pnRow, 6), "#,##0.00")
                End If
                
            ElseIf grdCuentaAbono.TextMatrix(pnRow, 8) = "E" Then
                'grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 4)) + grdCuentaAbono.TextMatrix(pnRow, 6), "#,##0.00")
                If grdCuentaAbono.TextMatrix(pnRow, 3) <> "" Or val(grdCuentaAbono.TextMatrix(pnRow, 4)) > 0 Then
                    grdCuentaAbono.TextMatrix(pnRow, 5) = Format(fgITFCalculaImpuesto(grdCuentaAbono.TextMatrix(pnRow, 3)), "#,##0.00")
                End If
                
                If nMonCta = gMonedaNacional Then
                        grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 3)) + grdCuentaAbono.TextMatrix(pnRow, 5), "#,##0.00")
                Else
                        grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 4)) + grdCuentaAbono.TextMatrix(pnRow, 6), "#,##0.00")
                End If
                
            End If
                GoTo lblCalculos
       End If
    End If
End If
lblCalculos:
CalculaTotales
End Sub
Private Sub CalculaTotales()
Dim i As Long, nFila As Long, nCol As Long
Dim nAcumMN As Double, nAcumME As Double, nMonto As Double
Dim nAcumIEMN As Double, nAcumIEME As Double
Dim nAcumICMN As Double, nAcumICME As Double
Dim nAcumIAMN As Double, nAcumIAME As Double
Dim nAcumTMN As Double, nAcumTME As Double

Dim sCuenta As String

nAcumIEMN = 0: nAcumIEME = 0
nAcumICMN = 0: nAcumICME = 0
nAcumTMN = 0: nAcumTME = 0
nAcumIAMN = 0: nAcumIAME = 0


Dim bValida As Boolean
nFila = grdCuentaAbono.row
nCol = grdCuentaAbono.Col
nAcumMN = 0
nAcumME = 0
For i = 1 To grdCuentaAbono.Rows - 1
    
        sCuenta = Trim(grdCuentaAbono.TextMatrix(i, 1))
    
    '********TOTALES 1
    If grdCuentaAbono.TextMatrix(i, 3) <> "" Then
        nAcumMN = nAcumMN + CDbl(grdCuentaAbono.TextMatrix(i, 3))
    End If
    If grdCuentaAbono.TextMatrix(i, 4) <> "" Then
        nAcumME = nAcumME + CDbl(grdCuentaAbono.TextMatrix(i, 4))
    End If
    '**********TOTALES 1
    
    If Mid(sCuenta, 9, 1) = "1" And grdCuentaAbono.TextMatrix(i, 5) <> "" And grdCuentaAbono.TextMatrix(i, 7) = "S" And grdCuentaAbono.TextMatrix(i, 8) = "E" Then
        nAcumIEMN = nAcumIEMN + CDbl(grdCuentaAbono.TextMatrix(i, 5))
    End If
    If Mid(sCuenta, 9, 1) = "2" And grdCuentaAbono.TextMatrix(i, 6) <> "" And grdCuentaAbono.TextMatrix(i, 7) = "S" And grdCuentaAbono.TextMatrix(i, 8) = "E" Then
        nAcumIEME = nAcumIEME + CDbl(grdCuentaAbono.TextMatrix(i, 6))
    End If
    
    If Mid(sCuenta, 9, 1) = "1" And grdCuentaAbono.TextMatrix(i, 5) <> "" And grdCuentaAbono.TextMatrix(i, 8) = "A" Then
        nAcumIAMN = nAcumIAMN + CDbl(grdCuentaAbono.TextMatrix(i, 5))
    End If
    If Mid(sCuenta, 9, 1) = "2" And grdCuentaAbono.TextMatrix(i, 6) <> "" And grdCuentaAbono.TextMatrix(i, 8) = "A" Then
        nAcumIAME = nAcumIAME + CDbl(grdCuentaAbono.TextMatrix(i, 6))
    End If
          
    If Mid(sCuenta, 9, 1) = "1" And grdCuentaAbono.TextMatrix(i, 5) <> "" And grdCuentaAbono.TextMatrix(i, 7) = "S" And grdCuentaAbono.TextMatrix(i, 8) = "C" Then
        nAcumICMN = nAcumICMN + CDbl(grdCuentaAbono.TextMatrix(i, 5))
    End If
    If Mid(sCuenta, 9, 1) = "2" And grdCuentaAbono.TextMatrix(i, 6) <> "" And grdCuentaAbono.TextMatrix(i, 7) = "S" And grdCuentaAbono.TextMatrix(i, 8) = "C" Then
        nAcumICME = nAcumICME + CDbl(grdCuentaAbono.TextMatrix(i, 6))
    End If
        
    If grdCuentaAbono.TextMatrix(i, 9) <> "" And Mid(sCuenta, 9, 1) = 1 Then
        nAcumTMN = nAcumTMN + CDbl(grdCuentaAbono.TextMatrix(i, 9))
    End If
    If grdCuentaAbono.TextMatrix(i, 9) <> "" And Mid(sCuenta, 9, 1) = 2 Then
        nAcumTME = nAcumTME + CDbl(grdCuentaAbono.TextMatrix(i, 9))
    End If
            
Next i

'nMonto = txtMontoCargo.value

bValida = True

If nmoneda = gMonedaNacional Then
    If nOperacion = gAhoDepositoEnLoteCheq Then
       If txtMonto.Text = "" Then
            MsgBox "Elegir cheque para la operación.", vbInformation, "Aviso"
           bValida = False
       End If
       If bValida = True Then
           If Round(CDbl(txtMonto.Text), 2) < Round(nAcumMN, 2) Then
               MsgBox "SUMA TOTAL supera al monto establecido para cargar.", vbInformation, "Aviso"
               bValida = False
           ElseIf nMonto = nAcumMN Then
        '       cmdAgregar.Enabled = False
           Else
         '      cmdAgregar.Enabled = True
           End If
       End If
    End If
Else
    If txtMonto.Text = "" Then
        MsgBox "Elegir cheque para la operación.", vbInformation, "Aviso"
        bValida = False
    End If
    If bValida = True Then
        If Round(CDbl(txtMonto.Text), 2) < Round(nAcumME, 2) Then
            MsgBox "SUMA TOTAL supera al monto establecido para cargar.", vbInformation, "Aviso"
            bValida = False
        ElseIf nMonto = nAcumME Then
    '        cmdAgregar.Enabled = False
        Else
    '        cmdAgregar.Enabled = True
        End If
    End If
End If
grdCuentaAbono.row = nFila
grdCuentaAbono.Col = nCol


CalcITFPorcentaje

Me.LblITFME.Caption = nAcumIEME + nAcumICME
Me.LblITFMN.Caption = nAcumIEMN + nAcumICMN
Me.lblITFAD.Caption = nAcumIAME
Me.lblITFAS.Caption = nAcumIAMN
Me.lblITFED.Caption = nAcumIEME
Me.lblITFES.Caption = nAcumIEMN
Me.lblITFCD.Caption = nAcumICME
Me.lblITFCS.Caption = nAcumICMN
Me.lblTNS.Caption = nAcumTMN
Me.lblTND.Caption = nAcumTME


If bValida Then
    lblTotalMN = Format$(nAcumMN, "#,##0.00")
    lblTotalME = Format$(nAcumME, "#,##0.00")
    
Else
    grdCuentaAbono.TextMatrix(nFila, 3) = "0.00"
    grdCuentaAbono.TextMatrix(nFila, 4) = "0.00"
End If

End Sub
Private Sub CalcITFPorcentaje()
Dim vCargoCalc As Double, i As Integer

vCantITF = 0
 If Not (vCantITF = grdCuentaAbono.Rows - 1) Then
        'fraCuentaAbono.Enabled = True
'        fraGlosa.Enabled = True
'        cmdEliminar.Enabled = False
'        txtGlosa.SetFocus
        

    For i = 1 To grdCuentaAbono.Rows - 1
        If grdCuentaAbono.TextMatrix(i, 11) = "S" Then
                If nmoneda = gMonedaNacional Then
                    vCantITF = vCantITF + grdCuentaAbono.TextMatrix(i, 3)
                Else
                    vCantITF = vCantITF + grdCuentaAbono.TextMatrix(i, 4)
                End If
        End If
   Next i
                vCargoCalc = vCantITF
        
        
        If gbITFAplica Then       'Filtra para CTS
'                  If txtMontoCargo.value > gnITFMontoMin Then
'                        'ALPA 20091125***********************
'                        'If Not lbITFCtaExonerada Then
'                        If lnITFCtaExonerada = 0 Then
'                        '************************************
'                            Me.lblITF.Caption = Format(fgITFCalculaImpuesto(vCargoCalc), "#,##0.00")
'
'                        Else
'                            Me.lblITF.Caption = "0.00"
'                        End If
'
'                        If gbITFAsumidoAho Then
'                            Me.lblTotal.Caption = Format(vCargoCalc, "#,##0.00")
'                            Exit Sub
'                        ElseIf chkITFEfectivo.value = vbChecked Then
'                            Me.lblTotal.Caption = Format(CCur(txtMontoCargo.Text) + CCur(Me.lblITF.Caption), "#,##0.00")
'                            Exit Sub
'                        Else
'                            Me.lblTotal.Caption = Format(CCur(txtMontoCargo.Text), "#,##0.00") '+ CCur(LblItf.Caption), "#,##0.00")
'                            Exit Sub
'                        End If
'                 End If
    
         End If
        
    End If


End Sub
Private Sub cboTipoTasa_Click()
    nTipoTasa = CLng(Right(cboTipoTasa.Text, 4))
'DefineTasaGrid
End Sub
Private Sub cboMoneda_Click()
nmoneda = CLng(Right(cboMoneda.Text, 1))
If nmoneda = gMonedaNacional Then
    'txtMonto.BackColor = &HC0FFFF
    'lblMon.Caption = "S/."
    'grdCuentaAbono
ElseIf nmoneda = gMonedaExtranjera Then
    'txtMonto.BackColor = &HC0FFC0
    'lblMon.Caption = "US$"
End If
'Me.lblTasa.BackColor = txtMonto.BackColor
'Me.lblITF.BackColor = txtMonto.BackColor
'Me.lblTotal.BackColor = txtMonto.BackColor

If nOperacion = gAhoApeLoteChq Or nOperacion = gPFApeLoteChq Or nOperacion = gCTSApeLoteChq Then
    'Me.txtMonto.value = 0
    Me.lblNroDoc.Caption = ""
    Me.lblNombreIF.Caption = ""
End If

End Sub
Private Sub cboPrograma_Click()
Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim bOrdPag As Boolean
Dim nMonto As Double
Dim nPlazo As Long
Dim nTpoPrograma As Integer

Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
'bOrdPag = IIf(chkOrdenPago.value = 1, True, False)
'nMonto = txtMonto.value
nTpoPrograma = 6

If cboPrograma.Visible Then
    nTpoPrograma = CInt(Right(Trim(cboPrograma.Text), 2))
End If

If nTpoPrograma = 4 Then
    lblInst.Visible = True
    txtInstitucion.Visible = True
    lblInstEtq.Visible = True
Else
    lblInst.Visible = False
    txtInstitucion.Visible = False
    lblInstEtq.Visible = False
End If

If nProducto = gCapPlazoFijo Then
   
ElseIf nProducto = gCapAhorros Then
    nmoneda = 1
    nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, bOrdPag, nTpoPrograma)
    lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
    
Else
    nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma)
    lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
End If
    If cboPrograma.ListIndex = 0 Then
        Me.txtInstitucion.Visible = True
        Me.lblInst.Visible = True
        Me.lblInstEtq.Visible = True
    Else
        Me.txtInstitucion.Visible = False
        Me.lblInst.Visible = False
        Me.lblInstEtq.Visible = False
    End If
Set clsDef = Nothing

'DefineTasaGrid
End Sub
Public Sub Inicia(ByVal nProd As Producto, ByVal nOpe As CaptacOperacion, ByVal sDescOperacion As String)
nProducto = nProd
nOperacion = nOpe

fgITFParamAsume gsCodAge, CStr(nProd)

Me.cboPrograma.Visible = False
Label20.Visible = False
Me.lblTasa.Visible = False
Select Case nProd
    Case gCapAhorros
        Me.Caption = "Captaciones - Ahorros - " & sDescOperacion
        'lblPeriodo.Visible = False
        'cboPeriodoCTS.Visible = False
        'grdCuenta.ColWidth(4) = 0
        'grdCuenta.ColWidth(6) = 0
        'lblDispCTS.Visible = False
        'lblCTS.Visible = False
        lblInst.Visible = False
        lblInstEtq.Visible = False
        txtInstitucion.Visible = False
        'Me.fraITF.Visible = True
        'Me.chkITFEfectivo.value = 0
        'Me.chkITFEfectivo.Enabled = True
        
        
        If gbITFAsumidoAho Then
            'chkITFEfectivo.value = 1
            'chkITFEfectivo.Visible = False
        Else
            'chkITFEfectivo.value = 0
            'chkITFEfectivo.Visible = True
        End If
        IniciaCombo cboPrograma, 2030, 1
        Me.cboPrograma.Visible = True
        Label20.Visible = True
        Me.lblTasa.Visible = True
        Call cboPrograma_Click
        'FraITFAsume.Visible = False
         
End Select

Select Case nOperacion
    Case gAhoDepositoEnLoteEfec, gAhoDepositoEnLoteCarg
        fraDocumento.Visible = False
    Case gAhoDepositoEnLoteCheq
        fraDocumento.Visible = True
End Select

If nProducto = gCapAhorros And gbITFAsumidoAho Then
   ' Me.chkITFEfectivo.value = 0
ElseIf nProducto = gCapPlazoFijo And gbITFAsumidoPF Then
   ' Me.chkITFEfectivo.value = 0
End If

'chkExoITF_Click

IniciaCombo cboMoneda, gMoneda
IniciaCombo cboTipoTasa, gCaptacTipoTasa
'cmdAgregar.Enabled = True
'cmdEliminar.Enabled = False
'txtMonto.Enabled = False

Me.Show 1
End Sub
Private Sub IniciaCombo(ByRef cboConst As ComboBox, ByVal nCapConst As ConstanteCabecera, Optional nTipo As Integer = 0)
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim rsConst As New ADODB.Recordset
Set clsGen = New COMDConstSistema.DCOMGeneral
Set rsConst = clsGen.GetConstante(nCapConst)
Set clsGen = Nothing
If nTipo = 1 Then
    Do While Not rsConst.EOF
        If rsConst("nConsValor") = 6 Then
            cboConst.AddItem rsConst("cDescripcion") & Space(100) & rsConst("nConsValor")
        End If
        rsConst.MoveNext
    Loop
    cboConst.ListIndex = 0
Else
    Do While Not rsConst.EOF
        cboConst.AddItem rsConst("cDescripcion") & Space(100) & rsConst("nConsValor")
        rsConst.MoveNext
    Loop
    cboConst.ListIndex = 0
End If
End Sub
Private Sub CargaDatos()
Dim oDCaptacion As COMDCaptaGenerales.DCOMCaptaGenerales
'Dim oDCaptacion As COMNCredito.NCOMCredito
Dim R As ADODB.Recordset
Dim MatCalend As Variant

    On Error GoTo ErrorCargaDatos
    Set oDCaptacion = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set R = oDCaptacion.ObtenerCtasParaDepositoEnLote(txtInstitucion.Text, Trim(Right(cboMoneda.Text, 2)))
    If R.BOF And R.EOF Then
        MsgBox "No se Encontraron Registros", vbInformation, "Aviso"
        R.Close
        Set R = Nothing
        Exit Sub
    End If
    Do While Not R.EOF
        Call IngresarCuenta(R!cCtaCod)
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Exit Sub
ErrorCargaDatos:
    MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub IngresarCuenta(sCtaCod As String)
Dim bCObraITF As Boolean, nExonerada As Integer, bExiste As Boolean

Dim bSuCuenta As Boolean

bCObraITF = True

'If KeyAscii = 13 Then
    Dim sCta As String, sCtaCargo As String
    sCta = sCtaCod
   ' sCtaCargo = txtCuenta.NroCuenta
'     If sCta = sCtaCargo Then
'        MsgBox "La Cuenta de Abono no puede ser la misma cuenta de Cargo.", vbInformation, "Aviso"
'        txtCuentaAbo.SetFocusCuenta
'        Exit Sub
'    End If
    
    If Not CuentaExisteEnLista(sCta) Then
        bExiste = True
        ObtieneDatosCuentaAbono sCta, , , bCObraITF, nExonerada, bExiste, bSuCuenta

        If bSuCuenta Then
            Unload Me
            Exit Sub
        End If
        
        If bExiste = False Then Exit Sub
        If nExonerada = 3 Then
'            MsgBox "Cuenta de ahorro es una cuenta de haberes. Digitar otra cuenta"
'            grdCuentaAbono.EliminaFila grdCuentaAbono.Row
'            Exit Sub
        Else
            If nExonerada > 0 Then
                 grdCuentaAbono.TextMatrix(grdCuentaAbono.row, 7) = "N"
                 grdCuentaAbono.TextMatrix(grdCuentaAbono.row, 8) = ""
                 grdCuentaAbono.TextMatrix(grdCuentaAbono.row, 10) = "SI"
                 MsgBox "CUENTA EXONERADA DE ITF", vbOKOnly + vbInformation, "AVISO"
            Else
               
               grdCuentaAbono.TextMatrix(grdCuentaAbono.row, 10) = "NO"
               If gbITFAsumidoAho Then
                   grdCuentaAbono.TextMatrix(grdCuentaAbono.row, 8) = "A"
                   MsgBox "ITF asumido", vbOKOnly + vbInformation, "AVISO"
               Else
                    If bCObraITF Then
                        If MsgBox("¿EL cobró de ITF será en efectivo?", vbYesNo + vbDefaultButton2 + vbQuestion, "AVISO") = vbYes Then
                            grdCuentaAbono.TextMatrix(grdCuentaAbono.row, 8) = "E"
                        Else
                            grdCuentaAbono.TextMatrix(grdCuentaAbono.row, 8) = "C"
                        End If
                       
                    End If
                End If
                
            End If
                    
            cmdGrabar.Enabled = True
            cmdCancelar.Enabled = True
            'txtMontoCargo.Enabled = True
        End If
    Else
        MsgBox "Cuenta ya se encuentra en la lista.", vbInformation, "Aviso"
       ' txtCuentaAbo.SetFocusCuenta
    End If
'End If
End Sub
Private Function CuentaExisteEnLista(ByVal sCuenta As String) As Boolean
Dim bExito As Boolean
Dim i As Long
Dim sCuentaLista As String
bExito = False
For i = 1 To grdCuentaAbono.Rows - 1
    sCuentaLista = grdCuentaAbono.TextMatrix(i, 1)
    If sCuenta = sCuentaLista Then
        bExito = True
        Exit For
    End If
Next i
CuentaExisteEnLista = bExito
End Function
Private Sub cmdMostrar_Click()
   ' If Trim(Right(cboPrograma.Text, 2)) = "6" Then
        If txtInstitucion.Text = "" Then
            MsgBox "Ingresar Institución", vbCritical
            Exit Sub
      '  End If
    End If
    'EJVG20140408 ***
    If nOperacion = gAhoDepositoEnLoteCheq Then
        If Len(Trim(lblNroDoc.Caption)) = 0 Then
            MsgBox "Ud. debe especificar primero el Nro. de Documento", vbInformation, "Aviso"
            If cmdDocumento.Visible And cmdDocumento.Enabled Then cmdDocumento.SetFocus
            Exit Sub
        End If
    End If
    'END EJVG *******
    Call CargaDatos
End Sub

Private Sub txtInstitucion_EmiteDatos()
If txtInstitucion.Text <> "" Then
    lblInst.Caption = txtInstitucion.psDescripcion
    'cmdGrabar.SetFocus
End If
End Sub

Private Function ObtieneDatosCuentaAbono(ByVal sCuenta As String, Optional bArchivo As Boolean = False, _
        Optional nMonto As Double = 0, Optional ByRef bCObraITF As Boolean = True, Optional ByRef nExonerada As Integer = 0, Optional ByRef bExiste As Boolean = True, Optional ByRef bSuCuenta As Boolean = False) As Boolean

Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim rsCta As ADODB.Recordset, rsRel As New ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim nFila As Long
Dim sMsg As String, sMoneda As String, sPersona As String
Dim nMonedaAbono As Moneda, sCObraITF As String, i As Integer

Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
sMsg = clsCap.ValidaCuentaOperacion(sCuenta, True)
Set clsCap = Nothing
If sMsg <> "" Then bExiste = False

If sMsg = "" Then

sCObraITF = "S"

nExonerada = fgITFVerificaExoneracionInteger(sCuenta)

If nExonerada Then sCObraITF = "N"
If gbITFAsumidoAho Then sCObraITF = "N"

i = 1

    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        grdCuentaAbono.AdicionaFila
        nFila = grdCuentaAbono.Rows - 1
        grdCuentaAbono.TextMatrix(nFila, 1) = sCuenta
        'EJVG20140408 ***
        grdCuentaAbono.TextMatrix(nFila, 3) = "0.00"
        grdCuentaAbono.TextMatrix(nFila, 4) = "0.00"
        grdCuentaAbono.TextMatrix(nFila, 5) = "0.00"
        grdCuentaAbono.TextMatrix(nFila, 6) = "0.00"
        grdCuentaAbono.TextMatrix(nFila, 9) = "0.00"
        'END EJVG *******
        nMonedaAbono = CLng(Mid(sCuenta, 9, 1))
        
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        sPersona = ""
        
        Dim dlsMant As COMDCaptaGenerales.DCOMCaptaGenerales
        Set dlsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
        
        Do While Not rsRel.EOF
            
            If dlsMant.GetNroOPeradoras(gsCodAge) > 1 Then
                If rsRel("cPersCod") = gsCodPersUser Then
                            MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                            Set dlsMant = Nothing
                            bSuCuenta = True
                            ObtieneDatosCuentaAbono = False
                            Exit Function
                End If
            End If
                    
            If sPersona <> rsRel("cPersCod") And rsRel("nPrdPersRelac") = gCapRelPersTitular Then
                grdCuentaAbono.TextMatrix(nFila, 2) = UCase(PstaNombre(rsRel("Nombre")))
                Exit Do
            End If
            rsRel.MoveNext
        Loop
        rsRel.MoveFirst
        
        Do While Not rsRel.EOF
            
'            For i = 1 To grdCliente.Rows - 1
'                If grdCliente.TextMatrix(i, 3) = rsRel("cPersCod") Then
'                        sCObraITF = "N"
'                        lblITF.Caption = "0.00"
'                        lblTotal.Caption = txtMontoCargo.Text
'                        grdCuentaAbono.TextMatrix(nFila, 11) = "N"
'                        GoTo sContinuar
'                End If
'            Next i
            
            rsRel.MoveNext
        Loop
sContinuar:

        rsRel.Close
        Set rsRel = Nothing
        
            grdCuentaAbono.TextMatrix(nFila, 7) = sCObraITF
            If sCObraITF = "N" Then
                bCObraITF = False
            Else
                grdCuentaAbono.TextMatrix(nFila, 11) = "S"
            End If
                    
        If nMonedaAbono = gMonedaNacional Then
            grdCuentaAbono.BackColorRow vbWhite
            grdCuentaAbono.BackColorControl = vbWhite
            grdCuentaAbono.TextMatrix(nFila, 3) = nMonto
            
        Else
            grdCuentaAbono.BackColorRow &HC0FFC0
            grdCuentaAbono.BackColorControl = &HC0FFC0
            grdCuentaAbono.TextMatrix(nFila, 4) = nMonto
            
        End If
        
        If Not bArchivo Then
            grdCuentaAbono.lbEditarFlex = True
            grdCuentaAbono.SetFocus
            'cmdEliminar.Enabled = True
            cmdGrabar.Enabled = True
        End If
        ObtieneDatosCuentaAbono = True
    End If
Else
    If bArchivo Then
        'sError = sError & sMsg & gPrnSaltoLinea
    Else
        MsgBox sMsg, vbInformation, "Operacion"
        'cmdAgregar.SetFocus
    End If
    ObtieneDatosCuentaAbono = False
End If

If Not bArchivo Then
    'txtCuentaAbo.Visible = False
End If
Set clsMant = Nothing
End Function
Private Function IniciaLavDinero() As String
Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
Dim nPersoneria As COMDConstantes.PersPersoneria, sOperacion As String, sTipoCuenta As String
Dim nMonto As Double
Dim sCuenta As String
Dim oDatos As COMDPersona.DCOMPersonas
Dim rsPersona As ADODB.Recordset

Set oDatos = New COMDPersona.DCOMPersonas



For i = 1 To grdCuentaAbono.Rows - 1
 ' nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
'    If npersoneria = gPersonaNat Then
'        If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
'            sPersCod = grdCliente.TextMatrix(i, 3)
'            sNombre = grdCliente.TextMatrix(i, 1)
'            sDireccion = ""
'            sDocId = ""
'            Exit For
'        End If
'    Else
'        If nRelacion = gCapRelPersTitular Then
'            sPersCod = grdCliente.TextMatrix(i, 3)
'            sNombre = grdCliente.TextMatrix(i, 1)
'            sDireccion = ""
'            sDocId = ""
'            Exit For
'        End If
'    End If
Next i
'nMonto = txtMontoCargo.value
'sCuenta = txtCuenta.NroCuenta
'If sPersCodCMAC <> "" Then
'    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, sCuenta, sOperacion, , sTipoCuenta)
'Else
Set rsPersona = oDatos.dDatosPersonas(sPersCod)
  sDireccion = rsPersona("cPersDireccDomicilio")
  sDocId = rsPersona("cPersIdNro")
  'ALPA 20081009************************************************************************************************
  'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, True, nMonto, sCuenta, "TRANSFERENCIA AHORROS", , sTipoCuenta)
  IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, True, nMonto, sCuenta, "TRANSFERENCIA AHORROS", , sTipoCuenta, , , , , , , gnTipoREU, gnMontoAcumulado, gsOrigen)
  '*************************************************************************************************************
Set oDatos = Nothing
'End If
End Function
