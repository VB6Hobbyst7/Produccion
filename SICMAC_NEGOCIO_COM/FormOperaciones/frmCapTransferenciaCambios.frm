VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCapTransferenciaCambios 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7905
   ClientLeft      =   180
   ClientTop       =   1410
   ClientWidth     =   13140
   Icon            =   "frmCapTransferenciaCambios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   13140
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CdlgFile 
      Left            =   1155
      Top             =   7410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   105
      TabIndex        =   11
      Top             =   7485
      Width           =   915
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   12210
      TabIndex        =   10
      Top             =   7485
      Width           =   915
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   11190
      TabIndex        =   9
      Top             =   7485
      Width           =   915
   End
   Begin VB.Frame fraCuentaAbono 
      Caption         =   "Cuenta Abono"
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
      Height          =   4710
      Left            =   105
      TabIndex        =   20
      Top             =   2640
      Width           =   12990
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   1035
         TabIndex        =   7
         Top             =   3660
         Width           =   915
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   75
         TabIndex        =   6
         Top             =   3660
         Width           =   915
      End
      Begin SICMACT.ActXCodCta txtCuentaAbo 
         Height          =   375
         Left            =   2010
         TabIndex        =   8
         Top             =   3660
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Cuenta N°"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin SICMACT.FlexEdit grdCuentaAbono 
         Height          =   2655
         Left            =   180
         TabIndex        =   5
         Top             =   240
         Width           =   12660
         _ExtentX        =   22331
         _ExtentY        =   4683
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
         Left            =   11760
         TabIndex        =   49
         Top             =   3810
         Width           =   960
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
         Left            =   11760
         TabIndex        =   48
         Top             =   3390
         Width           =   960
      End
      Begin VB.Label Label21 
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
         Left            =   2550
         TabIndex        =   47
         Top             =   1875
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
         Left            =   5715
         TabIndex        =   46
         Top             =   4200
         Width           =   1065
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
         Left            =   7020
         TabIndex        =   45
         Top             =   4200
         Width           =   960
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
         Left            =   8235
         TabIndex        =   44
         Top             =   4200
         Width           =   960
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
         Left            =   10005
         TabIndex        =   43
         Top             =   3450
         Width           =   1635
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
         Left            =   10005
         TabIndex        =   42
         Top             =   3885
         Width           =   1590
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
         Left            =   8190
         TabIndex        =   41
         Top             =   3300
         Width           =   960
      End
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
         Left            =   6975
         TabIndex        =   40
         Top             =   3300
         Width           =   1065
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
         Left            =   8235
         TabIndex        =   39
         Top             =   3855
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
         Left            =   8235
         TabIndex        =   38
         Top             =   3525
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
         Left            =   7020
         TabIndex        =   37
         Top             =   3855
         Width           =   960
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
         Left            =   7020
         TabIndex        =   36
         Top             =   3525
         Width           =   960
      End
      Begin VB.Label Label5 
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
         Left            =   5670
         TabIndex        =   35
         Top             =   3615
         Width           =   1185
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
         Left            =   8865
         TabIndex        =   34
         Top             =   2880
         Width           =   1260
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
         Left            =   10110
         TabIndex        =   33
         Top             =   2880
         Width           =   1245
      End
      Begin VB.Label Label4 
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
         Left            =   5715
         TabIndex        =   32
         Top             =   3900
         Width           =   1200
      End
      Begin VB.Label Label2 
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
         Left            =   3060
         TabIndex        =   24
         Top             =   2895
         Width           =   3075
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
         Left            =   7500
         TabIndex        =   23
         Top             =   2895
         Width           =   1395
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
         Left            =   6120
         TabIndex        =   22
         Top             =   2895
         Width           =   1395
      End
   End
   Begin VB.Frame fraCuentaCargo 
      Caption         =   "Cuenta Cargo"
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
      Height          =   2595
      Left            =   105
      TabIndex        =   12
      Top             =   60
      Width           =   12960
      Begin VB.TextBox txtIdAut 
         Height          =   330
         Left            =   1560
         TabIndex        =   26
         Top             =   2670
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.CommandButton cmdObtDatos 
         Caption         =   "&Obtener Datos"
         Height          =   375
         Left            =   3825
         TabIndex        =   1
         Top             =   300
         Width           =   1230
      End
      Begin VB.Frame fraGlosa 
         Caption         =   "Glosa"
         Height          =   795
         Left            =   9375
         TabIndex        =   25
         Top             =   1665
         Width           =   3450
         Begin VB.TextBox txtGlosa 
            Height          =   435
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   240
            Width           =   3030
         End
      End
      Begin VB.Frame fraTipoCambio 
         Caption         =   "Tipo Cambio"
         Height          =   915
         Left            =   9375
         TabIndex        =   15
         Top             =   735
         Width           =   1575
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
            TabIndex        =   19
            Top             =   510
            Width           =   795
         End
         Begin VB.Label Label3 
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
            TabIndex        =   18
            Top             =   570
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
            TabIndex        =   17
            Top             =   210
            Width           =   795
         End
         Begin VB.Label Label1 
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
            TabIndex        =   16
            Top             =   270
            Width           =   360
         End
      End
      Begin VB.Frame fraMontoCargo 
         Caption         =   "Monto Total Cargo"
         Height          =   1740
         Left            =   5460
         TabIndex        =   13
         Top             =   660
         Width           =   3720
         Begin VB.CheckBox chkItfEfectivo 
            Caption         =   "ITF Efect"
            Height          =   345
            Left            =   120
            TabIndex        =   53
            Top             =   600
            Width           =   1035
         End
         Begin VB.CheckBox chkVBEfectivo 
            Caption         =   "Comision Efect"
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   930
            Visible         =   0   'False
            Width           =   1095
         End
         Begin SICMACT.EditMoney txtMontoCargo 
            Height          =   315
            Left            =   1230
            TabIndex        =   3
            Top             =   255
            Width           =   1890
            _ExtentX        =   3334
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
            ForeColor       =   255
            Text            =   "0"
         End
         Begin VB.Label Label11 
            Caption         =   "S/."
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
            Height          =   255
            Left            =   3150
            TabIndex        =   51
            Top             =   990
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblMonComision 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1230
            TabIndex        =   50
            Top             =   960
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.Label LblTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1230
            TabIndex        =   31
            Top             =   1305
            Width           =   1890
         End
         Begin VB.Label LblItf 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1230
            TabIndex        =   30
            Top             =   615
            Width           =   1890
         End
         Begin VB.Label Label7 
            Caption         =   "S/."
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
            Height          =   255
            Left            =   3150
            TabIndex        =   29
            Top             =   1335
            Width           =   315
         End
         Begin VB.Label Label6 
            Caption         =   "S/."
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
            Height          =   255
            Left            =   3150
            TabIndex        =   28
            Top             =   645
            Width           =   315
         End
         Begin VB.Label lblMon 
            Caption         =   "S/."
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
            Height          =   255
            Left            =   3150
            TabIndex        =   21
            Top             =   300
            Width           =   315
         End
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   180
         TabIndex        =   0
         Top             =   285
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Cuenta N°"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1635
         Left            =   180
         TabIndex        =   2
         Top             =   810
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2884
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Nombre-RE-cperscod-CCodRelacion"
         EncabezadosAnchos=   "250-4000-600-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Id Autorización"
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
         Left            =   240
         TabIndex        =   27
         Top             =   2730
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   4065
         TabIndex        =   14
         Top             =   240
         Width           =   5235
      End
   End
End
Attribute VB_Name = "frmCapTransferenciaCambios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nProducto As COMDConstantes.Producto
Dim nmoneda As COMDConstantes.Moneda
Dim nOperacion As COMDConstantes.CaptacOperacion
Private Const nLongPrimerRegistro = 66
Private Const nLongSegundoRegistro = 33
Dim sError As String, sCodPers As String
Dim vCantITF As Double

Private nPersoneria As COMDConstantes.PersPersoneria 'WIOR 20131009
'***************Variabres Agregadas********************
Dim Gtitular As String
Dim GAutNivel As String
Dim GAutMontoFinSol As Double
Dim GAutMontoFinDol As Double
Dim GMontoAprobado As Double, GNroID As Long, GPersCod As String

'Dim lbITFCtaExonerada As Boolean
Dim lnITFCtaExonerada As Boolean
'********************************************************

Dim sCuenta As String
Dim sNumTarj As String
Dim cGetValorOpe As String

Dim nRedondeoITF As Double 'RIRO20140724

Private Sub ObtieneDatosCuentasAbonar(ByVal sArchivo As String)
Dim sCad As String
Dim bPrimeraLinea As Boolean
Dim nMontoTotal As Double, nSumaTotal As Double
Dim nNumReg As Long, nItem As Long
Dim dFechaAbono As Date, dFechaProceso As Date
Dim sCuentaAbono As String, sCuentaCargo As String
Dim nMontoAbono As Double
On Error GoTo ErrFileOpen
Open sArchivo For Input As #1
bPrimeraLinea = True
nItem = 0
nSumaTotal = 0
sError = ""
Do While Not EOF(1)
    Line Input #1, sCad
    If sCad <> "" Then
        If bPrimeraLinea Then
            If Len(sCad) = nLongPrimerRegistro Then
                sCodPers = Left(sCad, 13)
                sCad = Mid(sCad, 14, Len(sCad) - 13)
                sCuentaCargo = Left(sCad, 18)
                sCad = Mid(sCad, 19, Len(sCad) - 18)
                If ObtieneDatosCuenta(sCuentaCargo, True) Then
                    nNumReg = CLng(Trim(Left(sCad, 8)))
                    sCad = Mid(sCad, 9, Len(sCad) - 8)
                    nMontoTotal = CDbl(Trim(Mid(sCad, 1, 9)) & "." & Trim(Mid(sCad, 10, 2)))
                    sCad = Mid(sCad, 12, Len(sCad) - 11)
                    dFechaAbono = CDate(Mid(sCad, 7, 2) & "/" & Mid(sCad, 5, 2) & "/" & Mid(sCad, 1, 4))
                    sCad = Mid(sCad, 9, Len(sCad) - 8)
                    If DateDiff("d", gdFecSis, dFechaAbono) >= 0 Then
                        dFechaProceso = CDate(Mid(sCad, 7, 2) & "/" & Mid(sCad, 5, 2) & "/" & Mid(sCad, 1, 4))
                        If DateDiff("d", gdFecSis, dFechaProceso) < 0 Then
                            sError = sError & "Fecha de Proceso es mayor que la fecha actual" & gPrnSaltoLinea
                        End If
                    Else
                        sError = sError & "Fecha de Abono es menor que la fecha actual" & gPrnSaltoLinea
                    End If
                    txtCuenta.Age = sCuentaCargo
                End If
                bPrimeraLinea = False
            Else
                sError = sError & "Longitud del primer registro no coincide con formato establecido" & gPrnSaltoLinea
            End If
        Else
            sCad = Mid(sCad, 5, Len(sCad) - 4)
            sCuentaAbono = Left(sCad, 18)
            sCad = Mid(sCad, 19, Len(sCad) - 18)
            nMontoAbono = CDbl(Trim(Mid(sCad, 1, 9)) & "." & Trim(Mid(sCad, 10, 2)))
            If Not CuentaExisteEnLista(sCuentaAbono) Then
                ObtieneDatosCuentaAbono sCuentaAbono, True, nMontoAbono
            Else
                sError = sError & "Cuenta N° " & sCuentaAbono & "Duplicada en la relación" & gPrnSaltoLinea
            End If
            nSumaTotal = nSumaTotal + nMontoAbono
            nItem = nItem + 1
        End If
    End If
Loop
Close #1
If nItem <> nNumReg Then
    sError = sError & "Número de Cuentas NO coincide con el total de registros enviados. " & nNumReg & " - " & nItem & gPrnSaltoLinea
End If
If Round(nMontoTotal, 2) - Round(nSumaTotal, 2) <> 0 Then
    sError = sError & "Monto Total NO coincide con la SUMA TOTAL de MONTOS A ABONAR. " & nMontoTotal & " - " & nSumaTotal & gPrnSaltoLinea
End If
If sError <> "" Then
    Dim oPrevio As previo.clsprevio
    Set oPrevio = New previo.clsprevio
        oPrevio.Show sError, "Errores Cargo Abono en Lote", True, , gImpresora
    Set oPrevio = Nothing
    cmdCancelar_Click
    Exit Sub
End If
txtMontoCargo.value = nMontoTotal
CalculaTotales
cmdGrabar.Enabled = True
cmdCancelar.Enabled = True
fraCuentaAbono.Enabled = True
fraMontoCargo.Enabled = True
txtCuenta.Enabled = False
cmdObtDatos.Enabled = False
grdCuentaAbono.lbEditarFlex = False
Exit Sub
ErrFileOpen:
    Close #1
    cmdCancelar_Click
    MsgBox err.Description, vbExclamation, "Error"
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

nMonto = txtMontoCargo.value

bValida = True
If nmoneda = gMonedaNacional Then
    If Round(nMonto, 2) < Round(nAcumMN, 2) Then
        MsgBox "SUMA TOTAL supera al monto establecido para cargar.", vbInformation, "Aviso"
        bValida = False
    ElseIf nMonto = nAcumMN Then
        cmdAgregar.Enabled = False
    Else
        cmdAgregar.Enabled = True
    End If
Else
    If Round(nMonto, 2) < Round(nAcumME, 2) Then
        MsgBox "SUMA TOTAL supera al monto establecido para cargar.", vbInformation, "Aviso"
        bValida = False
    ElseIf nMonto = nAcumME Then
        cmdAgregar.Enabled = False
    Else
        cmdAgregar.Enabled = True
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

Private Function ObtieneDatosCuenta(ByVal sCuenta As String, Optional bArchivo As Boolean = False) As Boolean
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento  'NCapMovimientos
Dim rsCta As ADODB.Recordset, rsRel As New ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim nRow As Long
Dim sMsg As String, sMoneda As String, sPersona As String
Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    sMsg = clsCap.ValidaCuentaOperacion(sCuenta)
Set clsCap = Nothing


vCantITF = 0


If sMsg = "" Then
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    
     'ITF INICIO
        lnITFCtaExonerada = fgITFVerificaExoneracionInteger(sCuenta)
        fgITFParamAsume Mid(sCuenta, 4, 2), Mid(sCuenta, 6, 3)
            
            Me.chkItfEfectivo.value = 0
            
            If gbITFAsumidoAho Then
                Me.chkItfEfectivo.Visible = False
            Else
                Me.chkItfEfectivo.Visible = True
            End If
        
     'ITF FIN
    
    
    If Not (rsCta.EOF And rsCta.BOF) Then
        nmoneda = CLng(Mid(sCuenta, 9, 1))
        If nmoneda = gMonedaNacional Then
            sMoneda = "MONEDA NACIONAL"
            txtMontoCargo.BackColor = &HC0FFFF
            LblItf.BackColor = &HC0FFFF
            LblTotal.BackColor = &HC0FFFF
            lblMon.Caption = "S/."
        Else
            sMoneda = "MONEDA EXTRANJERA"
            txtMontoCargo.BackColor = &HC0FFC0
            LblItf.BackColor = &HC0FFC0
            LblTotal.BackColor = &HC0FFC0
            lblMon.Caption = "$"
        End If
        
        If rsCta("bOrdPag") Then
            lblMensaje = "AHORROS CON ORDEN DE PAGO" & Chr$(13) & sMoneda
        Else
            lblMensaje = "AHORROS SIN ORDEN DE PAGO" & Chr$(13) & sMoneda
        End If
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        sPersona = ""
        
        Dim dlsMant As COMDCaptaGenerales.DCOMCaptaGenerales 'DCapMantenimiento
        Set dlsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
        nPersoneria = rsCta("nPersoneria") 'WIOR 20131009
        Do While Not rsRel.EOF
        
            If dlsMant.GetNroOPeradoras(gsCodAge) > 1 Then
                   If rsRel("cPersCod") = gsCodPersUser Then
                        MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                        Unload Me
                        Exit Function
                   End If
            End If
                     
     
            
        
            If sPersona <> rsRel("cPersCod") Then
                grdCliente.AdicionaFila
                nRow = grdCliente.Rows - 1
                grdCliente.TextMatrix(nRow, 1) = UCase(PstaNombre(rsRel("Nombre")))
                grdCliente.TextMatrix(nRow, 2) = Left(UCase(rsRel("Relacion")), 2)
                grdCliente.TextMatrix(nRow, 3) = rsRel!cPersCod
                grdCliente.TextMatrix(nRow, 4) = Trim(rsRel("nPrdPersRelac"))
                sPersona = rsRel("cPersCod")
            End If
            
            rsRel.MoveNext
        Loop
        
        'Add By Gitu 23-08-2011 para cobro de comision por operacion sin tarjeta
        If sNumTarj = "" Then
            cGetValorOpe = ""
            If nmoneda = gMonedaNacional Then
                cGetValorOpe = GetMontoDescuento(2117, 1, 1)
            Else
                cGetValorOpe = GetMontoDescuento(2118, 1, 2)
            End If
            lblMonComision = Format(cGetValorOpe, "#,##0.00")
        End If
        'End Gitu
        
        Set dlsMant = Nothing
        rsRel.Close
        Set rsRel = Nothing
        txtCuenta.Enabled = False
        txtMontoCargo.Enabled = True
        'txtMontoCargo.SetFocus
        cmdCancelar.Enabled = True
        txtCuenta.Age = Mid(sCuenta, 4, 2)
        txtCuenta.Cuenta = Mid(sCuenta, 9, 10)
        fraCuentaAbono.Enabled = True
        cmdAgregar.Enabled = True
        ObtieneDatosCuenta = True
    End If
Else
    If bArchivo Then
        sError = sError & sMsg & gPrnSaltoLinea
    Else
        MsgBox sMsg, vbInformation, "Operacion"
        txtCuenta.SetFocus
    End If
    ObtieneDatosCuenta = False
End If
Set clsMant = Nothing
End Function
'ALPA 20091117***********************************************
'Private Function ObtieneDatosCuentaAbono(ByVal sCuenta As String, Optional bArchivo As Boolean = False, _
'        Optional nMonto As Double = 0, Optional ByRef bCObraITF As Boolean = True, Optional ByRef bExonerada As Boolean = False, Optional ByRef bExiste As Boolean = True, Optional ByRef bSuCuenta As Boolean = False) As Boolean
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
'ALPA 20091117******************************************
'bExonerada = fgITFVerificaExoneracion(sCuenta)
nExonerada = fgITFVerificaExoneracionInteger(sCuenta)
'*******************************************************
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
        nMonedaAbono = CLng(Mid(sCuenta, 9, 1))
        
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        sPersona = ""
        
        Dim dlsMant As COMDCaptaGenerales.DCOMCaptaGenerales  'DCapMantenimiento
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
            
'            Set dlsMant = Nothing
            
        
            If sPersona <> rsRel("cPersCod") And rsRel("nPrdPersRelac") = gCapRelPersTitular Then
                grdCuentaAbono.TextMatrix(nFila, 2) = UCase(PstaNombre(rsRel("Nombre")))
                Exit Do
            End If
            rsRel.MoveNext
        Loop
        rsRel.MoveFirst
        
        Do While Not rsRel.EOF
            
            For i = 1 To grdCliente.Rows - 1
                If grdCliente.TextMatrix(i, 3) = rsRel("cPersCod") Then
                        sCObraITF = "N"
                        LblItf.Caption = "0.00"
                        LblTotal.Caption = txtMontoCargo.Text
                        grdCuentaAbono.TextMatrix(nFila, 11) = "N"
                        GoTo sContinuar
                End If
            Next i
            
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
            cmdEliminar.Enabled = True
            cmdGrabar.Enabled = True
        End If
        ObtieneDatosCuentaAbono = True
    End If
Else
    If bArchivo Then
        sError = sError & sMsg & gPrnSaltoLinea
    Else
        MsgBox sMsg, vbInformation, "Operacion"
        cmdAgregar.SetFocus
    End If
    ObtieneDatosCuentaAbono = False
End If

If Not bArchivo Then
    txtCuentaAbo.Visible = False
End If
Set clsMant = Nothing
End Function

Private Sub LimpiaControles()

vCantITF = 0

grdCliente.Clear
grdCliente.Rows = 2
grdCliente.FormaCabecera
grdCuentaAbono.Clear
grdCuentaAbono.Rows = 2
grdCuentaAbono.FormaCabecera
txtMontoCargo.value = 0
cmdGrabar.Enabled = False
txtCuenta.Age = ""
txtCuenta.Cuenta = ""
txtCuentaAbo.Age = ""
txtCuentaAbo.Cuenta = ""
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
fraCuentaAbono.Enabled = False
txtGlosa = ""
fraGlosa.Enabled = False
txtCuenta.Enabled = True
txtCuenta.SetFocus
lblMensaje = ""
lblTotalMN = ""
lblTotalME = ""
Me.LblItf = "0.00"
Me.LblTotal = "0.00"
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


End Sub

Private Sub chkITFEfectivo_Click()
 If chkItfEfectivo.value = 1 Then
        'Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00")
        Me.LblTotal.Caption = Format(Me.txtMontoCargo.value + CCur(Me.LblItf.Caption), "#,##0.00")
        If chkVBEfectivo.value = 1 Then
            Me.LblTotal.Caption = Format(Me.txtMontoCargo.value + CCur(Me.LblItf.Caption) + CCur(Me.lblMonComision.Caption), "#,##0.00")
        End If
    Else
        If gbITFAsumidoAho Then
            Me.LblTotal.Caption = Format(txtMontoCargo.value, "#,##0.00")
        
        Else
            Me.LblTotal.Caption = Format(txtMontoCargo.value) '- CCur(Me.LblItf.Caption), "#,##0.00")
            If chkVBEfectivo.value = 1 Then
                Me.LblTotal.Caption = Format(Me.txtMontoCargo.value + CCur(Me.lblMonComision.Caption), "#,##0.00")
            End If
        End If
        
        'Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00")
    End If
End Sub

Private Sub chkVBEfectivo_Click()
    If chkVBEfectivo.value = 1 And chkItfEfectivo.value = 1 Then
        Me.LblTotal.Caption = Format(Me.txtMontoCargo.value + CCur(Me.LblItf.Caption) + CCur(Me.lblMonComision.Caption), "#,##0.00")
    ElseIf chkVBEfectivo.value = 1 And chkItfEfectivo.value = 0 Then
        Me.LblTotal.Caption = Format(txtMontoCargo.value + CCur(Me.lblMonComision.Caption), "#,##0.00")
    ElseIf chkVBEfectivo.value = 0 And chkItfEfectivo.value = 1 Then
        Me.LblTotal.Caption = Format(txtMontoCargo.value + CCur(Me.LblItf.Caption), "#,##0.00")
    Else
        Me.LblTotal.Caption = Format(txtMontoCargo.value)
    End If
End Sub

Private Sub cmdAgregar_Click()
txtCuentaAbo.Age = ""
txtCuentaAbo.Cuenta = ""
txtCuentaAbo.Visible = True
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
txtMontoCargo.Enabled = False
txtCuentaAbo.SetFocus
End Sub

Private Sub cmdCancelar_Click()
LimpiaControles
End Sub

Private Sub CmdEliminar_Click()
If MsgBox("¿Desea Eliminar la cuenta de la Relación?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    grdCuentaAbono.EliminaFila grdCuentaAbono.row
    If Trim(grdCuentaAbono.TextMatrix(1, 1)) = "" Then
        cmdEliminar.Enabled = False
        cmdGrabar.Enabled = False
    End If
    CalculaTotales
End If
End Sub
'****Agregado MPBR
Private Function ObtTitular() As String
Dim i As Integer
For i = 1 To grdCliente.Rows - 1
If Right(grdCliente.TextMatrix(i, 4), 2) = "10" Then
      ObtTitular = Trim(grdCliente.TextMatrix(i, 3))
      Exit For
  End If
Next i
End Function

Private Sub cmdGrabar_Click()
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'JUEZ 20150404
Dim nMontoCargo As Double
Dim sCuenta As String, sGlosa As String

Dim lsmensaje As String
Dim lsBoleta As String
Dim lsBoletaITF As String

Dim nFicSal As Integer

Dim Autid As Long

nMontoCargo = txtMontoCargo.value
sCuenta = txtCuenta.NroCuenta

If lblTotalMN = "" Or lblTotalME = "" Then
    MsgBox "Debe ingresar cuenta(s) para el abono", vbInformation, "Aviso"
    cmdAgregar.SetFocus
    Exit Sub
End If

If nMontoCargo = 0 Then
    MsgBox "Monto de Cargo debe ser mayor a cero", vbInformation, "Aviso"
    txtMontoCargo.SetFocus
    Exit Sub
End If
If nmoneda = gMonedaNacional Then
    If nMontoCargo <> CDbl(lblTotalMN) Then
        MsgBox "Suma total no coincide como monto de cargo", vbInformation, "Aviso"
        cmdAgregar.SetFocus
        Exit Sub
    End If
Else
    If nMontoCargo <> CDbl(lblTotalME) Then
        MsgBox "Suma total no coincide como monto de cargo", vbInformation, "Aviso"
        cmdAgregar.SetFocus
        Exit Sub
    End If
End If

'JUEZ 20150404 **************************************************
Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
If Not clsCap.ValidaSaldoCuenta(sCuenta, nMontoCargo + IIf(chkItfEfectivo.value = 0, CDbl(LblItf.Caption), 0)) Then
    MsgBox "Cuenta NO Posee SALDO SUFICIENTE", vbInformation, "Aviso"
    Set clsCap = Nothing
    Exit Sub
End If
Set clsCap = Nothing
'END JUEZ *******************************************************

If MsgBox("¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    'Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
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
    Dim nMontoLavDinero As Double, nTC As Double ', sReaPersLavDinero As String, sBenPersLavDinero As String JACA 20110224
    Dim loLavDinero As frmMovLavDinero 'JACA 20110225
    Set loLavDinero = New frmMovLavDinero 'JACA 20110225

    'Realiza la Validación para el Lavado de Dinero
    sCuenta = txtCuenta.NroCuenta
    Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
    'If clsLav.EsOperacionEfectivo(Trim(nOperacion)) Then
        Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
        If Not clsExo.EsCuentaExoneradaLavadoDinero(sCuenta) Then
            Set clsExo = Nothing
            sPersLavDinero = ""
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            If nmoneda = gMonedaNacional Then
            'cambiar
                Dim clsTC As COMDConstSistema.NCOMTipoCambio
                Set clsTC = New COMDConstSistema.NCOMTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
            Else
                nTC = 1
            End If
            If nMontoCargo >= Round(nMontoLavDinero * nTC, 2) Then
                
                'JACA 20110225
                    sPersLavDinero = IniciaLavDinero(loLavDinero)
                    'ALPA 20081030****************************************
                    'sPersLavDinero = gVarPublicas.gReaPersLavDinero
                    sPersLavDinero = loLavDinero.OrdPersLavDinero
                    '*****************************************************
'                    sReaPersLavDinero = gVarPublicas.gReaPersLavDinero 'COMENTADO X JACA 20110225
'                    sBenPersLavDinero = gVarPublicas.gBenPersLavDinero'COMENTADO X JACA 20110225
                
                'JACA END
                
                If sPersLavDinero = "" Then Exit Sub
                                
            End If
        Else
            Set clsExo = Nothing
        End If
        Set clsExo = Nothing
        Set clsLav = Nothing
    'Else
    '    Set clsLav = Nothing
   ' End If
    
      
    'If clsCap.CapTransferenciaAho(sCuenta, nMontoCargo, sMovNro, rsCtaAbo, sGlosa, gsNomAge, sLpt, sPersLavDinero, CDbl(Me.lblTCC.Caption), CDbl(Me.lblTCV.Caption), gbITFAplica, Me.LblItf.Caption, gbITFAsumidoAho, IIf(Me.chkItfEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, , , , , , , gnMovNro) Then COMENTADO X JACA 20110225
    If clsCap.CapTransferenciaAho(sCuenta, nMontoCargo, sMovNro, rsCtaAbo, sGlosa, gsNomAge, sLpt, sPersLavDinero, CDbl(Me.lblTCC.Caption), CDbl(Me.lblTCV.Caption), gbITFAplica, Me.LblItf.Caption, gbITFAsumidoAho, IIf(Me.chkItfEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), loLavDinero.BenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, , , , , , , gnMovNro) Then ' JACA 20110225
     'ALPA 20081010***************
     If gnMovNro > 0 Then
        'Call frmMovLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, sBenPersLavDinero, , , , , , gnTipoREU, gnMontoAcumulado, gsOrigen)COMENTADO X JACA 20110225
         Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110225
     End If
     '****************************
      If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
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

'Private Function IniciaLavDinero() As String ' JACA 20110225
Private Function IniciaLavDinero(ByVal loLavDinero As frmMovLavDinero) As String ' JACA 20110225

Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
'Dim nPersoneria As COMDConstantes.PersPersoneria'WIOR 20131009 COMENTÓ
Dim sOperacion As String, sTipoCuenta As String
Dim nMonto As Double
Dim sCuenta As String
Dim oDatos As COMDPersona.DCOMPersonas
Dim rsPersona As ADODB.Recordset

Set oDatos = New COMDPersona.DCOMPersonas

'WIOR 20131009 ****************************************************
For i = 1 To grdCliente.Rows - 1
    nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 4), 4)))
    If nPersoneria = gPersonaNat Then
        If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            loLavDinero.TitPersLavDinero = grdCliente.TextMatrix(i, 3)
            loLavDinero.TitPersLavDineroNom = grdCliente.TextMatrix(i, 1)
            Exit For
        End If
    Else
        If nRelacion = gCapRelPersTitular Then
            loLavDinero.TitPersLavDinero = grdCliente.TextMatrix(i, 3)
            loLavDinero.TitPersLavDineroNom = grdCliente.TextMatrix(i, 1)
        End If
        If nRelacion = gCapRelPersRepTitular Then
            loLavDinero.ReaPersLavDinero = grdCliente.TextMatrix(i, 3)
            loLavDinero.ReaPersLavDineroNom = grdCliente.TextMatrix(i, 1)
            If loLavDinero.TitPersLavDinero <> "" Then Exit For
        End If
    End If
Next i
'WIOR FIN ********************************************************


'WIOR 20131009 COMENTO TODO DE ABAJO*******************************
'For i = 1 To grdCuentaAbono.Rows - 1
'    nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
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
            'sPersCod = grdCliente.TextMatrix(i, 3)'WIOR 20131009 COMENTO
            sPersCod = loLavDinero.TitPersLavDinero 'WIOR 20131009
            'sNombre = grdCliente.TextMatrix(i, 1)'WIOR 20131009 COMENTO
            sNombre = loLavDinero.TitPersLavDineroNom 'WIOR 20131009
            sDireccion = ""
            sDocId = ""
            'Exit For
'        End If
'    End If
'Next i

nMonto = txtMontoCargo.value
sCuenta = txtCuenta.NroCuenta
'If sPersCodCMAC <> "" Then
'    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, sCuenta, sOperacion, , sTipoCuenta)
'Else
Set rsPersona = oDatos.dDatosPersonas(sPersCod)
  sDireccion = rsPersona("cPersDireccDomicilio")
  sDocId = rsPersona("cPersIdNro")
  'ALPA 20081009************************************************************************************************
  'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, True, nMonto, sCuenta, "TRANSFERENCIA AHORROS", , sTipoCuenta)
  'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, True, nMonto, sCuenta, "TRANSFERENCIA AHORROS", , sTipoCuenta, , , , , , , gnTipoREU, gnMontoAcumulado, gsOrigen) ' JACA 20110225
  'IniciaLavDinero = loLavDinero.inicia(sPersCod, sNombre, sDireccion, sDocId, True, True, nMonto, sCuenta, "TRANSFERENCIA AHORROS", , sTipoCuenta, , , , , , , gnTipoREU, gnMontoAcumulado, gsOrigen) ' JACA 20110225 'COMMENT BY MARG ERS073 ANEXO 02
  IniciaLavDinero = loLavDinero.inicia(sPersCod, sNombre, sDireccion, sDocId, True, True, nMonto, sCuenta, "TRANSFERENCIA AHORROS", , sTipoCuenta, , , , , , , gnTipoREU, gnMontoAcumulado, gsOrigen, , nOperacion) ' JACA 20110225 'ADD BY MARG ERS073 ANEXO 02
  '*************************************************************************************************************
Set oDatos = Nothing
'End If
End Function


Private Sub cmdObtDatos_Click()
Dim sArchivo As String
On Local Error Resume Next
CdlgFile.CancelError = True
'Especificar las extensiones a usar
CdlgFile.DefaultExt = "*.txt"
CdlgFile.Filter = "Textos (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
CdlgFile.ShowOpen
If err Then
    sArchivo = "" 'Cancelada la operación de abrir
Else
    sArchivo = CdlgFile.FileName
    ObtieneDatosCuentasAbonar sArchivo
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim sCuenta As String
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        
        sCuenta = frmValTarCodAnt.inicia(gCapAhorros, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    ElseIf KeyCode = vbKeyF12 And txtCuentaAbo.Enabled = True Then 'F12
       
        sCuenta = frmValTarCodAnt.inicia(gCapAhorros, False)
        If sCuenta <> "" Then
            txtCuentaAbo.NroCuenta = sCuenta
            txtCuentaAbo.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Me.Caption = "Captaciones - Ahorros - Transferencia de Cuentas"
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Prod = Trim(gCapAhorros)
txtCuentaAbo.CMAC = gsCodCMAC
txtCuentaAbo.Prod = Trim(gCapAhorros)
txtCuenta.EnabledProd = False
txtCuentaAbo.EnabledProd = False
txtCuenta.EnabledCMAC = False
txtCuentaAbo.EnabledCMAC = False
txtCuentaAbo.Visible = False
Dim clsGen As COMDConstSistema.NCOMTipoCambio
Dim rsTC As ADODB.Recordset
Set clsGen = New COMDConstSistema.NCOMTipoCambio
lblTCC = Format$(clsGen.EmiteTipoCambio(gdFecSis, TCCompra), "#0.0000")
lblTCV = Format$(clsGen.EmiteTipoCambio(gdFecSis, TCVenta), "#0.0000")
Set clsGen = Nothing
fraCuentaAbono.Enabled = False
fraGlosa.Enabled = False
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
End Sub

Private Sub grdCuentaAbono_KeyPress(KeyAscii As Integer)
'    If Mid(grdCuentaAbono.TextMatrix(grdCuentaAbono.Row, 1), 9, 1) = "1" Then
'        grdCuentaAbono.TextMatrix(grdCuentaAbono.Row, 4) = ""
'    Else
'        grdCuentaAbono.TextMatrix(grdCuentaAbono.Row, 3) = ""
'    End If
End Sub

Private Sub grdCuentaAbono_OnCellChange(pnRow As Long, pnCol As Long)
Dim nMonCta As COMDConstantes.Moneda
Dim nMonto As Double
'Dim bExonerada As Boolean

nMonCta = CLng(Mid(grdCuentaAbono.TextMatrix(pnRow, 1), 9, 1))
nMonto = CDbl(grdCuentaAbono.TextMatrix(pnRow, pnCol))

'bExonerada = fgITFVerificaExoneracion(grdCuentaAbono.TextMatrix(pnRow, 1))

Dim nRedondeoITF As Double 'RIRO20140725

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
            
            ' *** RIRO20140724
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(grdCuentaAbono.TextMatrix(pnRow, 5)))
            If nRedondeoITF > 0 Then
                grdCuentaAbono.TextMatrix(pnRow, 5) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 5)) - nRedondeoITF, "#,##0.00")
            End If
            ' *** END RIRO
            
            If grdCuentaAbono.TextMatrix(pnRow, 8) = "C" Then
          
                If grdCuentaAbono.TextMatrix(pnRow, 4) <> "" Or Val(grdCuentaAbono.TextMatrix(pnRow, 4)) > 0 Then
                    grdCuentaAbono.TextMatrix(pnRow, 6) = Format(fgITFCalculaImpuesto(grdCuentaAbono.TextMatrix(pnRow, 4)), "#,##0.00")
                End If
                If nMonCta = gMonedaNacional Then
                      grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 3)) - grdCuentaAbono.TextMatrix(pnRow, 5), "#,##0.00")
                Else
                      grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 4)) - grdCuentaAbono.TextMatrix(pnRow, 6), "#,##0.00")
                End If
                
            ElseIf grdCuentaAbono.TextMatrix(pnRow, 8) = "E" Then
                
                If grdCuentaAbono.TextMatrix(pnRow, 4) <> "" Or Val(grdCuentaAbono.TextMatrix(pnRow, 4)) > 0 Then
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
       
            ' *** RIRO20140724
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(grdCuentaAbono.TextMatrix(pnRow, 6)))
            If nRedondeoITF > 0 Then
                grdCuentaAbono.TextMatrix(pnRow, 6) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 6)) - nRedondeoITF, "#,##0.00")
            End If
            ' *** END RIRO
       
            If grdCuentaAbono.TextMatrix(pnRow, 8) = "C" Then
               ' grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 4)) - grdCuentaAbono.TextMatrix(pnRow, 6), "#,##0.00")
                If grdCuentaAbono.TextMatrix(pnRow, 3) <> "" Or Val(grdCuentaAbono.TextMatrix(pnRow, 4)) > 0 Then
                   grdCuentaAbono.TextMatrix(pnRow, 5) = Format(fgITFCalculaImpuesto(grdCuentaAbono.TextMatrix(pnRow, 3)), "#,##0.00")
                End If
                
                If nMonCta = gMonedaNacional Then
                      grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 3)) - grdCuentaAbono.TextMatrix(pnRow, 5), "#,##0.00")
                Else
                      grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 4)) - grdCuentaAbono.TextMatrix(pnRow, 6), "#,##0.00")
                End If
                
            ElseIf grdCuentaAbono.TextMatrix(pnRow, 8) = "E" Then
                'grdCuentaAbono.TextMatrix(pnRow, 9) = Format(CCur(grdCuentaAbono.TextMatrix(pnRow, 4)) + grdCuentaAbono.TextMatrix(pnRow, 6), "#,##0.00")
                If grdCuentaAbono.TextMatrix(pnRow, 3) <> "" Or Val(grdCuentaAbono.TextMatrix(pnRow, 4)) > 0 Then
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

Private Sub grdCuentaAbono_RowColChange()
If grdCuentaAbono.TextMatrix(grdCuentaAbono.row, 1) <> "" Then
    If CLng(Mid(grdCuentaAbono.TextMatrix(grdCuentaAbono.row, 1), 9, 1)) = gMonedaNacional Then
        grdCuentaAbono.BackColorControl = vbWhite
    Else
        grdCuentaAbono.BackColorControl = &HC0FFC0
    End If
End If
End Sub

Private Sub Label10_Click()

End Sub

Private Sub lblTNS_Click()

'If chkItfEfectivo.value = 1 Then
'        'Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00")
'        Me.LblTotal.Caption = Format(Me.txtMontoCargo.value + CCur(Me.LblItf.Caption), "#,##0.00")
'    Else
'        If gbITFAsumidoAho Then
'                    Me.LblTotal.Caption = Format(txtMontoCargo.value, "#,##0.00")
'
'        Else
'                    Me.LblTotal.Caption = Format(txtMontoCargo.value + CCur(Me.LblItf.Caption), "#,##0.00")
'        End If
'
'        'Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00")
'End If


End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCta As String
    sCta = txtCuenta.NroCuenta
    ObtieneDatosCuenta sCta
End If
End Sub


Private Sub txtCuentaAbo_KeyPress(KeyAscii As Integer)
'ALPA 20091117**************************************************
'Dim bCObraITF As Boolean, bExonerada As Boolean, bExiste As Boolean
Dim bCObraITF As Boolean, nExonerada As Integer, bExiste As Boolean
'***************************************************************
Dim bSuCuenta As Boolean

bCObraITF = True

If KeyAscii = 13 Then
    Dim sCta As String, sCtaCargo As String
    sCta = txtCuentaAbo.NroCuenta
    sCtaCargo = txtCuenta.NroCuenta
    If sCta = sCtaCargo Then
        MsgBox "La Cuenta de Abono no puede ser la misma cuenta de Cargo.", vbInformation, "Aviso"
        txtCuentaAbo.SetFocusCuenta
        Exit Sub
    End If
    
    If Not CuentaExisteEnLista(sCta) Then
        bExiste = True
        'ALPA 20091117*****************
        'ObtieneDatosCuentaAbono sCta, , , bCObraITF, bExonerada, bExiste, bSuCuenta
        ObtieneDatosCuentaAbono sCta, , , bCObraITF, nExonerada, bExiste, bSuCuenta
        '*****************************
        If bSuCuenta Then
            Unload Me
            Exit Sub
        End If
        
        If bExiste = False Then Exit Sub
        If nExonerada = 3 Then
            MsgBox "Cuenta de ahorro es una cuenta de haberes. Digitar otra cuenta"
            grdCuentaAbono.EliminaFila grdCuentaAbono.row
            Exit Sub
        Else
            'If bExonerada Then
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
            txtMontoCargo.Enabled = True
        End If
    Else
        MsgBox "Cuenta ya se encuentra en la lista.", vbInformation, "Aviso"
        txtCuentaAbo.SetFocusCuenta
    End If
End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAgregar.SetFocus
End If
End Sub

Private Sub txtIdAut_KeyPress(KeyAscii As Integer)

 Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Gtitular = ObtTitular
   If Gtitular = "" Then
    MsgBox "Esta cuenta no tiene titular", vbOKOnly + vbInformation, "Atención"
    Exit Sub
   End If
   nOperacion = gAhoTransferencia
      
   If KeyAscii = 13 And Trim(txtIdAut.Text) <> "" And Len(txtCuenta.NroCuenta) = 18 Then
      Dim ocapaut As COMDCaptaGenerales.COMDCaptAutorizacion
      Set ocapaut = New COMDCaptaGenerales.COMDCaptAutorizacion
            Set rs = ocapaut.SAA(Left(CStr(nOperacion), 4) & "00", Vusuario, txtCuenta.NroCuenta, Gtitular, CInt(nmoneda), CLng(Val(txtIdAut.Text)))
      Set ocapaut = Nothing
     If rs.State = 1 Then
       If rs.RecordCount > 0 Then
        txtMontoCargo.Text = rs!nMontoAprobado
      Else
          MsgBox "No Existe este Id de Autorización para esta cuenta." & vbCrLf & "Consulte las Operaciones Pendientes.", vbOKOnly + vbInformation, "Atención"
          txtIdAut.Text = ""
       End If
       
     End If
   End If

 If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 13 Or KeyAscii = 8) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtMontoCargo_GotFocus()
txtMontoCargo.MarcaTexto
End Sub
Private Sub CalcITFPorcentaje()
Dim vCargoCalc As Double, i As Integer

vCantITF = 0
 If txtMontoCargo.value > 0 And Not (vCantITF = grdCuentaAbono.Rows - 1) Then
        fraCuentaAbono.Enabled = True
        fraGlosa.Enabled = True
        cmdEliminar.Enabled = False
        txtGlosa.SetFocus
        

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
                  If txtMontoCargo.value > gnITFMontoMin Then
                        'ALPA 20091125***********************
                        'If Not lbITFCtaExonerada Then
                        If lnITFCtaExonerada = 0 Then
                        '************************************
                            Me.LblItf.Caption = Format(fgITFCalculaImpuesto(vCargoCalc), "#,##0.00")
                            
                            ' *** RIRO20140724
                            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.LblItf.Caption))
                            If nRedondeoITF > 0 Then
                                Me.LblItf.Caption = CCur(Me.LblItf.Caption) - nRedondeoITF
                            End If
                            ' *** END RIRO
                            
                        Else
                            Me.LblItf.Caption = "0.00"
                        End If
            
                        If gbITFAsumidoAho Then
                            Me.LblTotal.Caption = Format(vCargoCalc, "#,##0.00")
                            Exit Sub
                        ElseIf chkItfEfectivo.value = vbChecked Then
                            Me.LblTotal.Caption = Format(CCur(txtMontoCargo.Text) + CCur(Me.LblItf.Caption), "#,##0.00")
                            Exit Sub
                        Else
                            Me.LblTotal.Caption = Format(CCur(txtMontoCargo.Text), "#,##0.00") '+ CCur(LblItf.Caption), "#,##0.00")
                            Exit Sub
                        End If
                 End If
    
         End If
        
    End If


End Sub
Private Sub txtMontoCargo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtMontoCargo.value > 0 Then
        fraCuentaAbono.Enabled = True
        fraGlosa.Enabled = True
        cmdEliminar.Enabled = False
        txtGlosa.SetFocus
        
        If gbITFAplica Then       'Filtra para CTS
                  If txtMontoCargo.value > gnITFMontoMin Then
                        'ALPA 20091125***************************
                        'If Not lbITFCtaExonerada Then
                        If lnITFCtaExonerada = 0 Then
                        '****************************************
                            Me.LblItf.Caption = Format(fgITFCalculaImpuesto(txtMontoCargo.value), "#,##0.00")
                            
                            ' *** RIRO20140724
                            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.LblItf.Caption))
                            If nRedondeoITF > 0 Then
                                Me.LblItf.Caption = CCur(Me.LblItf.Caption) - nRedondeoITF
                            End If
                            ' *** END RIRO
                            
                        Else
                            Me.LblItf.Caption = "0.00"
                        End If
            
                        If gbITFAsumidoAho Then
                            Me.LblTotal.Caption = Format(txtMontoCargo.value, "#,##0.00")
                            Exit Sub
                        ElseIf chkItfEfectivo.value = vbChecked Then
                            Me.LblTotal.Caption = Format(CCur(txtMontoCargo.Text) + CCur(Me.LblItf.Caption), "#,##0.00")
                            If chkVBEfectivo.value = 1 Then
                                Me.LblTotal.Caption = Format(CCur(txtMontoCargo.Text) + CCur(Me.LblItf.Caption) + CCur(Me.lblMonComision.Caption), "#,##0.00")
                            End If
                            Exit Sub
                        Else
                            Me.LblTotal.Caption = Format(CCur(txtMontoCargo.Text)) ' + CCur(LblItf.Caption), "#,##0.00")
                            If chkVBEfectivo.value = 1 Then
                                Me.LblTotal.Caption = Format(CCur(txtMontoCargo.Text) + CCur(Me.lblMonComision.Caption), "#,##0.00")
                            End If
                            Exit Sub
                        End If
                 End If
    
         End If
        
    Else
        MsgBox "Monto debe ser mayor a cero", vbInformation, "Aviso"
        txtMontoCargo.SetFocus
    End If
End If
End Sub



Private Function Cargousu(ByVal NomUser As String) As String
 Dim rs As New ADODB.Recordset
 Dim oConts As COMDConstSistema.DCOMUAcceso
 Set oConts = New COMDConstSistema.DCOMUAcceso
    Set rs = oConts.Cargousu(NomUser)
 Set oConts = Nothing
 
 If Not (rs.EOF And rs.BOF) Then
    Cargousu = rs(0)
 End If
 
End Function
'ADD By GITU para el uso de las operaciones con tarjeta
Public Sub inicia()
    nOperacion = gsOpeCod
    nProducto = gCapAhorros
    If gnCodOpeTarj = 1 Then
        sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(nOperacion), sNumTarj, 232)
        If sCuenta <> "123456789" Then
            If Val(Mid(sCuenta, 6, 3)) <> nProducto And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            If sCuenta <> "" Then
                txtCuenta.NroCuenta = sCuenta
                'txtCuenta.SetFocusCuenta
                ObtieneDatosCuenta sCuenta
                Me.Show 1
            End If
        Else
            chkVBEfectivo.Visible = True
            lblMonComision.Visible = True
            Label11.Visible = True
            Me.Show
        End If
    Else
        Me.Show 1
    End If
    
End Sub
'End GITU
Private Function GetMontoDescuento(pnTipoDescuento As CaptacParametro, Optional pnCntPag As Integer = 0, _
                                   Optional pnMoneda As Integer = 1) As Double
Dim oParam As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim rsPar As New ADODB.Recordset

Set oParam = New COMNCaptaGenerales.NCOMCaptaDefinicion
'Modi By Gitu 29-08-2011
    Set rsPar = oParam.GetTarifaParametro(nOperacion, pnMoneda, pnTipoDescuento)
'End Gitu
Set oParam = Nothing

If rsPar.EOF And rsPar.BOF Then
    GetMontoDescuento = 0
Else
    GetMontoDescuento = rsPar("nParValor") * pnCntPag
End If
rsPar.Close
Set rsPar = Nothing
End Function



