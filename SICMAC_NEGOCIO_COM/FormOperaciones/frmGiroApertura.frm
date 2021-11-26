VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DB786848-D4E8-474E-A2C2-DCBC1D43DA22}#2.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmGiroApertura 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8340
   ClientLeft      =   4125
   ClientTop       =   1965
   ClientWidth     =   7890
   Icon            =   "frmGiroApertura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFormaPago 
      Caption         =   "Forma de pago"
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
      Height          =   720
      Left            =   120
      TabIndex        =   50
      Top             =   3840
      Width           =   7635
      Begin VB.ComboBox CmbForPag 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   250
         Width           =   1785
      End
      Begin SICMACT.ActXCodCta txtCuentaCargo 
         Height          =   375
         Left            =   3120
         TabIndex        =   52
         Top             =   230
         Visible         =   0   'False
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCta      =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblNroDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   3105
         TabIndex        =   55
         Top             =   250
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label LblNumDoc 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4300
         TabIndex        =   54
         Top             =   200
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblFormaPago 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
         Height          =   195
         Left            =   240
         TabIndex        =   53
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.TextBox txtMontoGiro 
      Alignment       =   1  'Right Justify
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
      Height          =   345
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   2650
      Width           =   1935
   End
   Begin OCXTarjeta.CtrlTarjeta Tarjeta 
      Height          =   375
      Left            =   2520
      TabIndex        =   46
      Top             =   7800
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   7875
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6660
      TabIndex        =   7
      Top             =   7875
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   7875
      Width           =   1095
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Remitente/Destinatario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3195
      Left            =   120
      TabIndex        =   19
      Top             =   4560
      Width           =   7635
      Begin VB.TextBox txtIdAut 
         Height          =   330
         Left            =   6075
         TabIndex        =   32
         Top             =   195
         Visible         =   0   'False
         Width           =   1380
      End
      Begin TabDlg.SSTab tabRemDest 
         Height          =   2610
         Left            =   150
         TabIndex        =   20
         Top             =   480
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   4604
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Remitente"
         TabPicture(0)   =   "frmGiroApertura.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraRem"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Destinatario"
         TabPicture(1)   =   "frmGiroApertura.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraDest"
         Tab(1).ControlCount=   1
         Begin VB.Frame fraDest 
            Height          =   2115
            Left            =   -74880
            TabIndex        =   29
            Top             =   360
            Width           =   7155
            Begin SICMACT.FlexEdit grdDest 
               Height          =   1755
               Left            =   60
               TabIndex        =   31
               Top             =   240
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   3096
               Cols0           =   5
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-Doc ID-Nombre-Referencia-cPersCod"
               EncabezadosAnchos=   "350-1000-3500-3500-0"
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
               TextStyleFixed  =   3
               ListaControles  =   "0-0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-L-L-C"
               FormatosEdit    =   "0-0-0-0-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               TipoBusqueda    =   3
               ColWidth0       =   345
               RowHeight0      =   300
               TipoBusPersona  =   1
               ForeColorFixed  =   -2147483630
            End
            Begin VB.CommandButton cmdEliminar 
               Caption         =   "&Eliminar"
               Height          =   375
               Left            =   6180
               TabIndex        =   5
               Top             =   1140
               Width           =   855
            End
            Begin VB.CommandButton cmdAgregar 
               Caption         =   "&Agregar"
               Height          =   375
               Left            =   6180
               TabIndex        =   4
               Top             =   720
               Width           =   855
            End
         End
         Begin VB.Frame fraRem 
            Height          =   2055
            Left            =   120
            TabIndex        =   21
            Top             =   420
            Width           =   7155
            Begin SICMACT.TxtBuscar txtDNI 
               Height          =   315
               Left            =   1140
               TabIndex        =   30
               Top             =   360
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
            End
            Begin VB.Label lblDireccion 
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
               Left            =   1140
               TabIndex        =   28
               Top             =   1260
               Width           =   5475
            End
            Begin VB.Label lblNombre 
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
               Left            =   1140
               TabIndex        =   27
               Top             =   840
               Width           =   5475
            End
            Begin VB.Label lblFecNac 
               Alignment       =   2  'Center
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
               Height          =   350
               Left            =   4500
               TabIndex        =   26
               Top             =   300
               Width           =   2115
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Fec. Nac. :"
               Height          =   195
               Left            =   3600
               TabIndex        =   25
               Top             =   375
               Width           =   795
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Dirección :"
               Height          =   195
               Left            =   180
               TabIndex        =   24
               Top             =   1335
               Width           =   765
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Nombre :"
               Height          =   195
               Left            =   180
               TabIndex        =   23
               Top             =   915
               Width           =   645
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Doc. ID.:"
               Height          =   195
               Left            =   180
               TabIndex        =   22
               Top             =   378
               Width           =   645
            End
         End
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
         Left            =   4740
         TabIndex        =   33
         Top             =   255
         Visible         =   0   'False
         Width           =   1290
      End
   End
   Begin VB.Frame fraDatos 
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
      ForeColor       =   &H8000000D&
      Height          =   3720
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   7635
      Begin VB.CommandButton cmdSeguridad 
         Caption         =   "&Ingresar clave de seguridad"
         Height          =   495
         Left            =   2640
         TabIndex        =   44
         Top             =   3035
         Width           =   2415
      End
      Begin VB.CheckBox chkITFEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3600
         TabIndex        =   37
         Top             =   1710
         Width           =   690
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   375
         Left            =   4875
         TabIndex        =   3
         Top             =   780
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin VB.ComboBox cboAgencia 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1665
         Width           =   2355
      End
      Begin VB.ComboBox cboTipo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1230
         Width           =   2355
      End
      Begin VB.ComboBox cboMoneda 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmGiroApertura.frx":0342
         Left            =   1095
         List            =   "frmGiroApertura.frx":0344
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   810
         Width           =   2475
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   150
         TabIndex        =   10
         Top             =   270
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   767
         Texto           =   "Giro N°"
      End
      Begin VB.Label Label13 
         Caption         =   "IMPORTE GIRO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2050
         TabIndex        =   49
         Top             =   2595
         Width           =   1455
      End
      Begin VB.Label lblClave 
         Height          =   375
         Left            =   5040
         TabIndex        =   45
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label LabelTC 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cambio:"
         Height          =   195
         Left            =   3840
         TabIndex        =   43
         Top             =   405
         Width           =   945
      End
      Begin VB.Label lblTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4875
         TabIndex        =   42
         Top             =   300
         Width           =   1965
      End
      Begin VB.Label LabelCS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comisión:                                                  S/"
         Height          =   195
         Left            =   165
         TabIndex        =   41
         Top             =   2160
         Width           =   3105
      End
      Begin VB.Label lblComSoles 
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1095
         TabIndex        =   40
         Top             =   2085
         Width           =   1965
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL :"
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
         Left            =   4035
         TabIndex        =   39
         Top             =   2100
         Width           =   750
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   4860
         TabIndex        =   38
         Top             =   2010
         Width           =   1950
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Comisión :"
         Height          =   195
         Left            =   4050
         TabIndex        =   36
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblITF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   4875
         TabIndex        =   35
         Top             =   1620
         Width           =   1935
      End
      Begin VB.Label lblSimbolo 
         AutoSize        =   -1  'True
         Caption         =   "S/"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   6975
         TabIndex        =   34
         Top             =   1695
         Width           =   270
      End
      Begin VB.Label lblSimbolo 
         AutoSize        =   -1  'True
         Caption         =   "S/"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   6990
         TabIndex        =   18
         Top             =   870
         Width           =   270
      End
      Begin VB.Label lblSimbolo 
         AutoSize        =   -1  'True
         Caption         =   "S/"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   6975
         TabIndex        =   17
         Top             =   1290
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Age. Dest.:"
         Height          =   195
         Left            =   165
         TabIndex        =   16
         Top             =   1755
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo :"
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   1290
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   4230
         TabIndex        =   14
         Top             =   930
         Width           =   555
      End
      Begin VB.Label lblComision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   4875
         TabIndex        =   13
         Top             =   1230
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ITF :"
         Height          =   195
         Left            =   4440
         TabIndex        =   12
         Top             =   1695
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   165
         TabIndex        =   11
         Top             =   870
         Width           =   675
      End
   End
   Begin VB.Label Label12 
      Caption         =   "IMPORTE GIRO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   47
      Top             =   2685
      Width           =   1455
   End
End
Attribute VB_Name = "frmGiroApertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type Tarifario
    nmoneda As COMDConstantes.Moneda
    nMontoIni As Double
    nMontoFin As Double
    nMontoMas As Double
    nValor As Double
    nTipo As COMDConstantes.GiroTarifaTipo
End Type


Dim aTarifa() As Tarifario
Dim sCodRemitente As String

'***************Variabres Agregadas********************
Dim Gtitular As String
Dim GAutNivel As String
Dim GAutMontoFinSol As Double
Dim GAutMontoFinDol As Double
Dim GMontoAprobado As Double, GNroID As Long, GPersCod As String
Dim SNroIdentidad

Dim ParamMinSoles As Double, ParamMinDolares As Double
Dim ParamMaxSoles As Double, ParamMaxDolares As Double
Dim ParamMinComSoles As Double, ParamMinComDolares As Double
Dim ParamComunicar As Double
Dim ParamNumTarjeta As String 'ADD JHCU
Dim ParamAdiComision As Double
Dim nRedondeoITF As Double 'BRGO 20110914
Dim nMontoParam As Double 'NAGL Según RFC1807260001
Dim nMontoConITF As Double 'NAGL Según RFC1807260001
Dim nMontoSinITF As Double 'NAGL Según RFC1807260001
'*****************

'Variables para la impresion de la boleta de Lavado de Dinero
Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
Dim sOperacion As String, sTipoCuenta As String
Dim fnPersRealizaOpe As Boolean 'WIOR 20121015
Dim fnPersPersoneriaGen As Integer 'WIOR 20121015
Dim fnPersCodRealiza As String 'WIOR 20121015
Dim fcPersCod As String 'WIOR 20130301

Dim lsNumTarjGir As String 'GITU 20170104
Dim lIpPuertoPinVerifyPOS As String 'GITU 20170104
Dim lsPINBlock As String 'GITU 20170104
Dim lsPINBlockConf As String 'JHCU 20190530
Private nMontoVoucher As Currency 'CTI4 ERS0112020
Dim nMovNroRVD As Long, nMovNroRVDPend As Long 'CTI4 ERS0112020
Dim sNumTarj As String 'CTI4 ERS0112020
Dim loVistoElectronico As frmVistoElectronico 'CTI4 ERS0112020
Dim nRespuesta As Integer 'CTI4 ERS0112020

Private Function IniciaLavDinero(ByVal poLavDinero As frmMovLavDinero)
Dim nMonto As Double, nmoneda As Integer

poLavDinero.TitPersLavDinero = sCodRemitente
poLavDinero.TitPersLavDineroNom = Trim(lblNombre)
poLavDinero.TitPersLavDineroDir = Trim(lblDireccion)
poLavDinero.TitPersLavDineroDoc = Trim(txtDNI.Text)
'nMonto = txtMonto.value 'COMENTADO X WIOR
nMonto = CDbl(txtMonto.value) + CDbl(lblComision.Caption)  'WIOR 20111107
sTipoCuenta = cboTipo.Text
nmoneda = Right(cboMoneda.Text, 1)

'IniciaLavDinero = frmMovLavDinero.Inicia(, , , , False, True, nMonto, "", sOperacion, False, sTipoCuenta, , , , , nmoneda)

End Function

Private Sub ClearScreen()

txtMonto.Text = ParamMinSoles
'txtMonto.value = 0
cboMoneda.ListIndex = 0
cboTipo.ListIndex = 0
cboAgencia.ListIndex = 0
lblNombre = ""
lblDireccion = ""
grdDest.FormaCabecera
grdDest.Rows = 2
lblFecNac = ""
txtDNI.Text = ""
sCodRemitente = ""
tabRemDest.Tab = 0

LabelTC.Visible = False
LabelCS.Visible = False
lblComSoles.Visible = False
lblTC.Visible = False
lblComSoles.Caption = ""
lblTC.Caption = ""
txtDNI.Text = ""
nRedondeoITF = 0
lblClave.Caption = ""
End Sub
'RECO20140520 ERS008-2014****************************************************************************
Private Function GetComision(ByVal nMonto As Double, ByVal nmoneda As COMDConstantes.Moneda) As Double
    Dim nComision As Double
    nComision = 0
    Dim oServ As New COMNCaptaServicios.NCOMCaptaServicios
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    nComision = oServ.ObtieneValorTarifarioGiro(gsCodAge, nMonto, nmoneda)
    GetComision = nComision
End Function
'Private Function GetComision(ByVal nMonto As Double, ByVal nmoneda As COMDConstantes.Moneda) As Double
'    Dim i As Long
'    Dim nComision As Double
'    Dim nParte As Double
'    'By Capi 23042008
'    Dim lnContador As Integer
'    lnContador = 1
'    '
'    nComision = 0
'    For i = 1 To UBound(aTarifa)
'        If aTarifa(i).nmoneda = nmoneda Then
'            If aTarifa(i).nMontoIni <= nMonto And nMonto <= aTarifa(i).nMontoFin Then
'                Select Case aTarifa(i).nTipo
'                    Case COMDConstantes.gGiroTarTpoPorcentaje
'                        nComision = Round(nMonto * aTarifa(i).nValor / 100, 2)
'                    Case COMDConstantes.gGiroTarTpoMonto
'                        nComision = aTarifa(i).nValor
'                End Select
'                Exit For
'            Else
'                If nMonto > aTarifa(i).nMontoIni Then
'                    nParte = nMonto - aTarifa(i).nMontoFin
'                    'By Capi 23042008
'                    Do While nParte > aTarifa(i).nMontoFin
'                        lnContador = lnContador + 1
'                        nParte = nParte - aTarifa(i).nMontoFin
'                    Loop
'
'
'                    nComision = aTarifa(i).nValor
'                    'By Capi 23042008
'                     nComision = nComision + lnContador * aTarifa(i).nMontoMas
'                    '
'                    'By Capi 23042008 comentado porque se cambio el calculo
''                    If nParte <= aTarifa(i).nMontoFin Then
''                        nComision = nComision + aTarifa(i).nMontoMas
''                    Else
''                        nComision = nComision + (Round(nParte / aTarifa(i).nMontoFin, 0) * aTarifa(i).nMontoMas)
''                    End If
'                End If
'            End If
'        End If
'    Next i
'    GetComision = nComision
'End Function
'RECO FIN
Private Sub CargaTarifario()
    Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim rs As ADODB.Recordset
    Dim i As Long
    Set rs = New ADODB.Recordset
    Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
    Set rs = clsServ.GetGiroTarifario()
    i = 0
    Do While Not rs.EOF
        i = i + 1
        ReDim Preserve aTarifa(i)
        aTarifa(i).nmoneda = rs("nMoneda")
        aTarifa(i).nMontoIni = rs("nMontoIni")
        aTarifa(i).nMontoFin = rs("nMontoFin")
        aTarifa(i).nMontoMas = rs("nMontoMas")
        aTarifa(i).nValor = rs("nValor")
        aTarifa(i).nTipo = rs("nTipoTarifa")
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set clsServ = Nothing
End Sub

Private Sub CargaParametros()
    Dim clsparam As COMDConstSistema.NCOMConstSistema
    Set clsparam = New COMDConstSistema.NCOMConstSistema

    ParamMinSoles = CDbl(clsparam.LeeConstSistema(205))
    ParamMinDolares = CDbl(clsparam.LeeConstSistema(206))
    ParamMaxSoles = CDbl(clsparam.LeeConstSistema(207))
    ParamMaxDolares = CDbl(clsparam.LeeConstSistema(208))
    ParamMinComSoles = CDbl(clsparam.LeeConstSistema(209))
    ParamMinComDolares = CDbl(clsparam.LeeConstSistema(210))
    ParamAdiComision = CDbl(clsparam.LeeConstSistema(306))
    ParamComunicar = CDbl(clsparam.LeeConstSistema(307))
    ParamNumTarjeta = clsparam.LeeConstSistema(76) 'JHCU
    Set clsparam = Nothing

End Sub
Private Sub CargaAgencias()
    Dim clsAge As COMDConstSistema.DCOMGeneral
    Dim rs As ADODB.Recordset
    Dim i As Long
    Set rs = New ADODB.Recordset
    Set clsAge = New COMDConstSistema.DCOMGeneral
    Set rs = clsAge.GetNombreAgencias(gsCodAge)
    Do While Not rs.EOF
        cboAgencia.AddItem rs("cAgeDescripcion") & space(50) & rs("cAgeCod")
        rs.MoveNext
    Loop
    cboAgencia.ListIndex = 0
    rs.Close
    Set clsAge = Nothing
    Set rs = Nothing
End Sub

Private Sub IniciaCombo(ByRef cboConst As ComboBox, ByVal nCapConst As COMDConstantes.ConstanteCabecera)
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

Private Sub cboAgencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMonto.SetFocus
End If
End Sub

Private Sub cboMoneda_Click()
Dim nmoneda As COMDConstantes.Moneda
Dim nTipoC As COMDConstSistema.DCOMGeneral
Dim vComision As Double

Set nTipoC = New COMDConstSistema.DCOMGeneral

nmoneda = CLng(Right(cboMoneda, 1))
If nmoneda = COMDConstantes.gMonedaNacional Then
    lblComision.BackColor = &HFFFFFF
    txtMonto.BackColor = &HFFFFFF
'    lblSimbolo(0).Caption = "S/."
'    lblSimbolo(1).Caption = "S/."
'    lblSimbolo(2).Caption = "S/."
    lblSimbolo(0).Caption = gcPEN_SIMBOLO
    lblSimbolo(1).Caption = gcPEN_SIMBOLO
    lblSimbolo(2).Caption = gcPEN_SIMBOLO
    LabelTC.Visible = False
    LabelCS.Visible = False
    lblComSoles.Visible = False
    lblTC.Visible = False
Else
    lblComision.BackColor = &HC0FFC0
    txtMonto.BackColor = &HC0FFC0
    lblSimbolo(0).Caption = "US$"
    lblSimbolo(1).Caption = "US$"
    lblSimbolo(2).Caption = "US$"
    LabelTC.Visible = True
    LabelCS.Visible = True
    lblComSoles.Visible = True
    lblTC.Visible = True
    'By Capi 21052008 de acuerdo a pedido de la Gerencia
    'lblTC.Caption = nTipoC.EmiteTipoCambio(gdFecSis, COMDConstantes.TCVenta)
    lblTC.Caption = nTipoC.EmiteTipoCambio(gdFecSis, COMDConstantes.TCFijoDia)
End If
Set nTipoC = Nothing

Dim sMoneda As String
        sMoneda = Right(cboMoneda.Text, 1)
        
        If sMoneda = 1 Then
                If txtMonto.value < ParamMinSoles Then
                   txtMonto.value = ParamMinSoles
                   MsgBox "El monto mínimo en giros de soles es de " & ParamMinSoles
                   
                ElseIf txtMonto.value > ParamMaxSoles Then
                
                   txtMonto.value = ParamMaxSoles
                   MsgBox "El monto máximo en giros de soles es de " & ParamMaxSoles
                End If
        Else
                If txtMonto.value < ParamMinDolares Then
                    txtMonto.value = ParamMinDolares
                    MsgBox "El monto mínimo en giros de dólares es de " & ParamMinDolares
                    
                ElseIf txtMonto.value > ParamMaxDolares Then
                
                    txtMonto.value = ParamMaxDolares
                    MsgBox "El monto máximo en giros de dólares es de " & ParamMaxDolares
                End If
       End If
        
        'By Capi 21052008 a pedido de la gerencia
        'vComision = GetComision(txtMonto.value, sMoneda)
        'vComision = GetComision(txtMonto.value, 1)
        vComision = GetComision(txtMonto.value, sMoneda) 'NAGL Según RFC1807260001
        
        If sMoneda = 1 Then
          If vComision < ParamMinComSoles Then
                vComision = ParamMinComSoles
          End If
        Else
          If vComision < ParamMinComDolares Then
                vComision = ParamMinComDolares
          End If
          If Trim(lblTC.Caption) <> "" Then
             Me.lblComSoles = Format$(vComision * Val(lblTC.Caption), "#,##0.00")
          End If
        End If
        
        lblComision.Caption = Format$(vComision, "#,##0.00")
        '***Agregado by NAGL Según RFC1807260001
        chkITFEfectivo.value = 1
        Call CalculaITF
        '***END NAGL 20181012
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'cboTipo.SetFocus
    cboAgencia.SetFocus 'MAVM 20091015
End If
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboAgencia.SetFocus
End If
End Sub


Private Sub chkITFEfectivo_Click()
nMontoConITF = CDbl(Me.txtMonto.value) + CDbl(Me.lblITF.Caption) 'NAGL 20181010
nMontoSinITF = CDbl(Me.txtMonto.value) - CDbl(Me.lblITF.Caption) 'NAGL 20181010

    If chkITFEfectivo.value = 1 Then
        'Me.lblTotal.Caption = Format(CDbl(Me.txtMonto.value) + CDbl(lblComision.Caption), "#,##0.00") 'Comentado by NAGL 20181010
        Me.lblTotal.Caption = Format(nMontoConITF + CDbl(lblComision.Caption), "#,##0.00") 'NAGL 20181010
        txtMontoGiro.Text = Format(CDbl(Me.txtMonto.value), "#,##0.00")
    Else
        'Me.lblTotal.Caption = Format(CDbl(Me.txtMonto.value) + CDbl(lblComision.Caption) - CDbl(Me.lblITF.Caption), "#,##0.00")'Comentado by NAGL 20181010
        '***ITF A CARGO
        nMontoParam = CDbl(Me.txtMonto.value)
        Me.lblTotal.Caption = Format(nMontoParam + CDbl(lblComision.Caption), "#,##0.00") 'NAGL 20181010
        txtMontoGiro.Text = Format(nMontoSinITF, "#,##0.00")
        '***NAGL 20181010
    End If
End Sub

Private Sub cmdAgregar_Click()
    
    If grdDest.Rows = 2 And grdDest.TextMatrix(1, 1) <> "" Then
        MsgBox "No se puede agregar mas de un destinatario", vbInformation + vbOKOnly
        Exit Sub
    Else
        frmGiroDestinatario.Show 1
    End If
End Sub

Private Sub cmdCancelar_Click()
    ClearScreen
End Sub

Private Sub CmdEliminar_Click()
    Dim nFila As Long
    nFila = grdDest.row
    grdDest.EliminaFila nFila
End Sub

Private Sub cmdGrabar_Click()

'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
'fin Comprobacion si es RFIII

Dim fbPersonaReaOtros As Boolean 'WIOR 20130301
Dim fnCondicion As Integer 'WIOR 20130301

    Dim nMonto As Double, nComision As Double
    Dim sCuenta As String
    Dim nmoneda As Moneda
    Dim nTipo As COMDConstantes.ProductoCuentaTipo
    Dim sAgenciaDest As String
    Dim rsDest As New ADODB.Recordset
    Dim lsClaveSeg As String
    Dim oConA As COMDConstSistema.DCOMUAcceso
    Dim nFicSal As String
    Dim lsBoleta As String
    Dim lsBoletaITF As String
    Dim sDNIRemitente As String 'NAGL 20181027
    Dim objPersona As COMDPersona.DCOMPersonas 'JACA 20110512
    Set objPersona = New COMDPersona.DCOMPersonas 'JACA 20110512
    
    'By Capi 14022008
    Dim loLavDinero As frmMovLavDinero
    Set loLavDinero = New frmMovLavDinero
    'End by
    '*****BRGO 20110914 ***********
    Dim oMov As COMDMov.DCOMMov
    Set oMov = New COMDMov.DCOMMov
    '*** End BRGO *****************
    
    Dim lsBoletaCargo  As String 'CTI4 ERS0112020
    Dim MatDatosAho(14) As String 'CTI4 ERS0112020
    Dim lsNombreClienteCargoCta As String 'CTI4 ERS0112020
    
    If Not ValidaFormaPago Then Exit Sub 'CTI4 ERS0112020
    
    Call VerSiClienteActualizoAutorizoSusDatos(fcPersCod, gServGiroApertEfec) 'FRHU ERS077-2015 20151204
    
    nmoneda = CLng(Right(cboMoneda.Text, 1))
    nMonto = CDbl(txtMontoGiro.Text) 'NAGL 20181027 Cambió de txtMonto.value a CDbl(txtMontoGiro.Text)
    nComision = CDbl(lblComision)
    lsClaveSeg = Trim(lblClave.Caption)
    sDNIRemitente = Trim(txtDNI.Text) 'NAGL 20181027
    'By Capi 05122007 para que sea obligatorio la clave de seguridad
    If lsClaveSeg = "" Then
        MsgBox "El Remitente no Ingreso Clave de Seguridad ", vbInformation, "SICMACM - Operaciones"
        cmdSeguridad.SetFocus
        Exit Sub
    End If
    
    If nmoneda = "1" Then
        If txtMonto.value < ParamMinSoles Then
            txtMonto.value = ParamMinSoles
            MsgBox "El monto mínimo en giros de soles es de " & ParamMinSoles
        ElseIf txtMonto.value > ParamMaxSoles Then
            txtMonto.value = ParamMaxSoles
            MsgBox "El monto máximo en giros de soles es de " & ParamMaxSoles
        End If
    ElseIf nmoneda = "2" Then
        If txtMonto.value < ParamMinDolares Then
            txtMonto.value = ParamMinDolares
            MsgBox "El monto mínimo en giros de dólares es de " & ParamMinDolares
        ElseIf txtMonto.value > ParamMaxDolares Then
            txtMonto.value = ParamMaxDolares
            MsgBox "El monto máximo en giros de dólares es de " & ParamMaxDolares
        End If
    End If

    If txtMonto.value = 0 Then
        MsgBox "Monto por girar debe ser mayor que cero", vbInformation, "SICMACM - Aviso"
        txtMonto.SetFocus
        Exit Sub
    End If
    If sCodRemitente = "" Then
        MsgBox "Debe registrar un remitente", vbInformation, "Aviso"
        tabRemDest.Tab = 0
        txtDNI.SetFocus
        Exit Sub
    End If
    If grdDest.TextMatrix(1, 2) = "" Then
        MsgBox "Debe registrar un destinatario", vbInformation, "Aviso"
        tabRemDest.Tab = 1
        cmdAgregar.SetFocus
        Exit Sub
    End If


    nTipo = CLng(Right(cboTipo.Text, 1))
    Set rsDest = grdDest.GetRsNew()
    
    If rsDest.RecordCount > 1 And nTipo = COMDConstantes.gPrdCtaTpoIndiv Then
        MsgBox "La cuenta no puede tener más de un destinatario si es de tipo INDIVIDUAL", vbInformation, "Aviso"
        tabRemDest.Tab = 1
        'cboTipo.SetFocus
        Exit Sub
    End If

    If rsDest.RecordCount = 1 And (nTipo = COMDConstantes.gPrdCtaTpoMancom Or nTipo = COMDConstantes.gPrdCtaTpoIndist) Then
        'MsgBox "La cuenta no puede tener un solo destinatario si es de tipo " & IIf(nTipo = COMDConstantes.gPrdCtaTpoMancom, "MANCOMUNADA", "INDISTINTA"), vbInformation, "Aviso"
        MsgBox "La cuenta no puede tener un solo destinatario si es de tipo " & IIf(nTipo = COMDConstantes.gPrdCtaTpoMancom, "MANCOMUNADA", "SOLIDARIA"), vbInformation, "Aviso" 'APRI20190109 ERS077-2018
        tabRemDest.Tab = 1
        cboTipo.SetFocus
        Exit Sub
    End If
    'JACA 20110512 *****VERIFICA SI LAS PERSONAS CUENTAN CON OCUPACION E INGRESO PROMEDIO
        Dim rsPersVerifica As Recordset
        Set rsPersVerifica = New Recordset
        
            Set rsPersVerifica = objPersona.ObtenerDatosPersona(sCodRemitente)
            If rsPersVerifica!nPersIngresoProm = 0 Or rsPersVerifica!cActiGiro1 = "" Then
                If MsgBox("Necesita Registrar la Ocupacion e Ingreso Promedio de: " + lblNombre, vbYesNo) = vbYes Then
                    'frmPersona.Inicio Me.grdCliente.TextMatrix(i, 1), PersonaActualiza
                    frmPersOcupIngreProm.Inicio sCodRemitente, lblNombre, rsPersVerifica!cActiGiro1, rsPersVerifica!nPersIngresoProm
                End If
            End If
       
    'JACA END***************************************************************************
    'RECO20140520 ERS008-2014****************************************************************************
    If txtDNI = grdDest.TextMatrix(1, 1) Then
        MsgBox "El destinatario no puede ser igual al remitente", vbCritical, "Aviso"
        Exit Sub
    End If
    If lblClave.Caption = "" Then
        MsgBox "La clave es obligatorio.", vbCritical, "Aviso"
        Exit Sub
    End If
    'RECO FIN***************
    If MsgBox("¿Desea Grabar la Operación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If

    On Error GoTo ErrGraba
    Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim sMovNro As String, sCuentaGiro As String, sPersLavDinero As String
    Dim clsMov As COMNContabilidad.NCOMContFunciones, sReaPersLavDinero As String, sBenPersLavDinero As String
    
    Dim nMontoLavDinero As Double, nTC As Double
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios
    
    'Realiza la Validación para el Lavado de Dinero
    Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
    If Not clsExo.EsPersonaExoneradaLavadoDinero(sCodRemitente) Then
        Set clsExo = Nothing
        sPersLavDinero = ""
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
        Set clsLav = Nothing
    
        If nmoneda = COMDConstantes.gMonedaNacional Then
            Dim clsTC As COMDConstSistema.DCOMGeneral
            Set clsTC = New COMDConstSistema.DCOMGeneral
            nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
            Set clsTC = Nothing
        Else
            nTC = 1
        End If
        If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
            'By Capi 1402208
            Call IniciaLavDinero(loLavDinero)
            'ALPA 20081009***********************************************************************************************
            'sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), False, sTipoCuenta, , , , , nmoneda)
            sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto + nComision, sCuenta, Mid(Me.Caption, 13), False, sTipoCuenta, , , , , nmoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen, , gServGiroApertEfec) 'WIOR 20131106 AGREGO gServGiroApertEfec
            If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
            '************************************************************************************************************
            'End
            
        End If
    Else
        Set clsExo = Nothing
    End If
    'WIOR 20130301 comento inicio ****************************************************************
    ''WIOR 20121015 *****************************************
    'If fnPersPersoneriaGen = 2 And loLavDinero.OrdPersLavDinero = "Exit" Then
    '    frmPersRealizaOperacion.Inicia "Giros", gPersRealizaGiros
    '    fnPersRealizaOpe = frmPersRealizaOperacion.PersRegistrar
    '    fnPersCodRealiza = frmPersRealizaOperacion.PersCod
    '    If Not fnPersRealizaOpe Then
    '        MsgBox "Se va a proceder a Anular la Apertura del Giro."
    '        Exit Sub
    '    End If
    'Else
    '    fnPersCodRealiza = "Exit"
    'End If
    ''WIOR FIN **********************************************
    'WIOR 20130301 comento fin **************************************************************
   'WIOR 20130301 Personas Sujetas a Procedimiento Reforzado *************************************
        fbPersonaReaOtros = False
        If (loLavDinero.OrdPersLavDinero = "Exit") Then
                
                Dim oPersonaSPR As UPersona_Cli
                Dim oPersonaU As COMDPersona.UCOMPersona
                Dim nTipoConBN As Integer
                Dim sConPersona As String
                Dim pbClienteReforzado As Boolean
                Dim rsAgeParam As Recordset
                Dim objCap As COMNCaptaGenerales.NCOMCaptaMovimiento
                Dim lnMontoX As Double, lnTC As Double
                Dim objTCX As COMDConstSistema.NCOMTipoCambio
                
                
                Set oPersonaU = New COMDPersona.UCOMPersona
                Set oPersonaSPR = New UPersona_Cli
                
                fbPersonaReaOtros = False
                pbClienteReforzado = False
                fnCondicion = 0
                
                    oPersonaSPR.RecuperaPersona fcPersCod
                                        
                    If oPersonaSPR.Personeria = 1 Then
                        If oPersonaSPR.Nacionalidad <> "04028" Then
                            sConPersona = "Extranjera"
                            fnCondicion = 1
                            pbClienteReforzado = True
                        ElseIf oPersonaSPR.Residencia <> 1 Then
                            sConPersona = "No Residente"
                            fnCondicion = 2
                            pbClienteReforzado = True
                        ElseIf oPersonaSPR.RPeps = 1 Then
                            sConPersona = "PEPS"
                            fnCondicion = 4
                            pbClienteReforzado = True
                        ElseIf oPersonaU.ValidaEnListaNegativaCondicion(IIf(Trim(oPersonaSPR.ObtenerDNI) = "", oPersonaSPR.ObtenerNumeroDoc(0), oPersonaSPR.ObtenerDNI), oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                            If nTipoConBN = 1 Or nTipoConBN = 3 Then
                                sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                                fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                                pbClienteReforzado = True
                            End If
                        End If
                    Else
                        If oPersonaU.ValidaEnListaNegativaCondicion(oPersonaSPR.ObtenerDNI, oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                            If nTipoConBN = 1 Or nTipoConBN = 3 Then
                                sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                                fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                                pbClienteReforzado = True
                            End If
                        End If
                    End If
                
                If pbClienteReforzado Then
                    MsgBox "El Cliente: " & Trim(Me.lblNombre.Caption) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
                    frmPersRealizaOpeGeneral.Inicia sOperacion & " (Persona " & sConPersona & ")", gServGiroApertEfec
                    fbPersonaReaOtros = frmPersRealizaOpeGeneral.PersRegistrar
                    
                    If Not fbPersonaReaOtros Then
                        MsgBox "Se va a proceder a Anular la Operacion ", vbInformation, "Aviso"
                        cmdGrabar.Enabled = True
                        Exit Sub
                    End If
                Else
                    fnCondicion = 0
                    lnMontoX = nMonto
                    pbClienteReforzado = False
                    
                    Set objTCX = New COMDConstSistema.NCOMTipoCambio
                    lnTC = objTCX.EmiteTipoCambio(gdFecSis, TCFijoDia)
                    Set objTCX = Nothing
                
                
                    Set objCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                    Set rsAgeParam = objCap.getCapAbonoAgeParam(gsCodAge)
                    Set objCap = Nothing
                    
                    If Trim(Right(Me.cboMoneda.Text, 5)) = 1 Then
                        lnMontoX = Round(lnMontoX / lnTC, 2)
                    End If
                
                    If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                        If lnMontoX >= rsAgeParam!nMontoMin And lnMontoX <= rsAgeParam!nMontoMax Then
                            frmPersRealizaOpeGeneral.Inicia sOperacion, gServGiroApertEfec
                            fbPersonaReaOtros = frmPersRealizaOpeGeneral.PersRegistrar
                            If Not fbPersonaReaOtros Then
                                MsgBox "Se va a proceder a Anular la Operacion", vbInformation, "Aviso"
                                cmdGrabar.Enabled = True
                                Exit Sub
                            End If
                        End If
                    End If
                    
                End If
        End If
        'WIOR FIN ***************************************************************
    'Registramos el giro
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    Dim psAgeDestino  As String
    sAgenciaDest = Right(cboAgencia.Text, 2)
    Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
    Dim oITF As New COMDConstSistema.FCOMITF
    psAgeDestino = Left(cboAgencia.Text, 30)
    Dim lsMovNro As String
    Dim loContFunct As COMNContabilidad.NCOMContFunciones
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
    lsMovNro = sMovNro
    Set loContFunct = Nothing
    Dim lsFechaHoraGrab As String
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    'By Capi 18022008
    'ALPA 20081009**********************************************************
    'sCuentaGiro = clsServ.ServGiroAperturaEfectivo(nMonto, nComision, ntipo, sAgenciaDest, sMovNro, rsDest, sCodRemitente, lblNombre, nmoneda, gsNomAge, sLpt, gsCodCMAC, , oITF.gbITFAplica, CCur(Me.lblITF.Caption), oITF.gbITFAsumidoGiros, IIf(Me.chkITFEfectivo.value = 1, COMDConstantes.gITFCobroEfectivo, COMDConstantes.gITFCobroCargo), psAgeDestino, lsBoleta, lsBoletaITF, lsClaveSeg, gbImpTMU)
    sCuentaGiro = clsServ.ServGiroAperturaEfectivo(nMonto, nComision, nTipo, sAgenciaDest, sMovNro, rsDest, sCodRemitente, lblNombre, nmoneda, gsNomAge, sLpt, gsCodCMAC, , gbITFAplica, CCur(Me.lblITF.Caption), gbITFAsumidoGiros, IIf(Me.chkITFEfectivo.value = 1, COMDConstantes.gITFCobroEfectivo, COMDConstantes.gITFCobroEfectivo), psAgeDestino, lsBoleta, lsBoletaITF, lsClaveSeg, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, gsOpeCod, sDNIRemitente, IIf(Me.chkITFEfectivo.value = 1, True, False), CInt(Trim(Right(CmbForPag.Text, 10))), nMovNroRVD, nMovNroRVDPend, txtCuentaCargo.NroCuenta, MatDatosAho, lsFechaHoraGrab)
    'NAGL 20181010 Quitó oITF. en las variables gbITFAplica / gbITFAsumidoGiros, así mismo Se reemplazo COMDConstantes.gITFCobroCargo con COMDConstantes.gITFCobroEfectivo, y se agregaron los parámetros gsOpeCod, IIf(Me.chkITFEfectivo.value = 1, True, False)
    'NAGL 20181027 Agregó sDNIRemitente
    If gnMovNro > 0 Then
        'Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
        Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
        Call oMov.InsertaMovRedondeoITF(sMovNro, 1, CCur(lblITF.Caption) + nRedondeoITF, CCur(lblITF.Caption)) 'BRGO 20110914
    End If
    '***********************************************************************
    Set clsServ = Nothing
    'CTI7 OPEv2********************************************
    Dim loImprime As COMNColoCPig.NCOMColPImpre
  
    Set loImprime = New COMNColoCPig.NCOMColPImpre
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
        Dim loGrabarRen As COMNColoCPig.NCOMColPContrato
        Set loGrabarRen = New COMNColoCPig.NCOMColPContrato
        lsNombreClienteCargoCta = PstaNombre(loGrabarRen.ObtieneNombreTitularCargoCta(txtCuentaCargo.NroCuenta))
        lsBoletaCargo = loImprime.ImprimeBoletaAhorro("RETIRO AHORROS", "CARGO A CUENTA POR GIRO.", "", CStr(nMonto + nComision + Me.lblITF.Caption), lsNombreClienteCargoCta, txtCuentaCargo.NroCuenta, "", CDbl(MatDatosAho(10)), CDbl(MatDatosAho(3)), "", 1, CDbl(MatDatosAho(11)), , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, gsCodCMAC, 0, , , gbImpTMU)
    End If
    'END CTI7
    Set loImprime = Nothing
    'JACA 20110510***********************************************************
        Dim ObjTc As COMDConstSistema.NCOMTipoCambio
        'Dim objPersona As COMDPersona.DCOMPersonas
        Dim rsPersOcu As Recordset
        Dim nAcumulado As Currency
        Dim nMontoPersOcupacion As Currency
        Dim nTpoCamb As Double
        
        Set ObjTc = New COMDConstSistema.NCOMTipoCambio
        nTpoCamb = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
        Set ObjTc = Nothing
         
        Set rsPersOcu = New Recordset
        'Set objPersona = New COMDPersona.DCOMPersonas
                        
        Set rsPersOcu = objPersona.ObtenerDatosPersona(sCodRemitente)
        nAcumulado = objPersona.ObtenerPersAcumuladoMontoOpe(nTpoCamb, Mid(Format(gdFecSis, "yyyymmdd"), 1, 6), sCodRemitente)
        nMontoPersOcupacion = objPersona.ObtenerParamPersAgeOcupacionMonto(Mid(sCodRemitente, 4, 2), CInt(Mid(rsPersOcu!cPersCIIU, 2, 2)))
    
        If nAcumulado >= nMontoPersOcupacion Then
            If Not objPersona.ObtenerPersonaAgeOcupDatos_Verificar(sCodRemitente, gdFecSis) Then
                objPersona.insertarPersonaAgeOcupacionDatos gnMovNro, sCodRemitente, IIf(nmoneda = "1", lblTotal.Caption, lblTotal.Caption * nTpoCamb), nAcumulado, gdFecSis, sMovNro
            End If
        End If
               
        
    'JACA END*****************************************************************
    


    If Trim(lsBoleta) <> "" Then
        'By Capi 20012008
        Dim lbok As Boolean
        lbok = True
        Do While lbok
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
            Print #nFicSal, lsBoleta & lsBoletaCargo
            Print #nFicSal, ""
            Close #nFicSal
            'By Capi 20012008 para reimprimir
            If MsgBox("Desea Reimprimir Boleta ??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbok = False
            End If
        Loop
        
    End If

If Trim(lsBoletaITF) <> "" Then
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
        Print #nFicSal, lsBoletaITF
        Print #nFicSal, ""
    Close #nFicSal
End If

If sPersLavDinero <> "" Then
    'Call loLavDinero.imprimirBoletaREU(sCuentaGiro, Mid(sCuentaGiro, 9, 1))' COMENTADO X JACA 20110302
    Call loLavDinero.imprimirBoletaREU(sCuentaGiro, Mid(sCuentaGiro, 9, 1), loLavDinero.OrigenPersLavDinero, loLavDinero.NroREU) 'JACA 20110302
End If
'WIOR 20130301 COMENTO INICIO **********************************
''WIOR 20121015 ************************************
'If fnPersRealizaOpe Then
'    frmPersRealizaOperacion.InsertaPersonaRealizaOperacion gnMovNro, sCuentaGiro, frmPersRealizaOperacion.PersTipoCliente, _
'    frmPersRealizaOperacion.PersCod, frmPersRealizaOperacion.PersTipoDOI, frmPersRealizaOperacion.PersDOI, frmPersRealizaOperacion.PersNombre, _
'    frmPersRealizaOperacion.TipoOperacion
'
'    fnPersRealizaOpe = False
'
'End If
'fnPersPersoneriaGen = 0
''WIOR FIN *****************************************
'WIOR 20130301 COMENTO FIN **************************************
'WIOR 20130301 ************************************************************
If fbPersonaReaOtros And gnMovNro > 0 Then
    frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(sCuentaGiro), fnCondicion
    fbPersonaReaOtros = False
End If
'WIOR FIN *****************************************************************
    
'CTI4 ERS0112020
If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
    Dim oMovOperacion As COMDMov.DCOMMov
    Dim nMovNroOperacion As Long
    Dim rsCli As New ADODB.Recordset
    Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
    Set oMovOperacion = New COMDMov.DCOMMov
    nMovNroOperacion = oMovOperacion.GetnMovNro(lsMovNro)

    loVistoElectronico.RegistraVistoElectronico nMovNroOperacion, , gsCodUser, nMovNroOperacion

    If nRespuesta = 2 Then
        Set rsCli = clsCli.GetPersonaCuenta(txtCuentaCargo.NroCuenta, gCapRelPersTitular)
        oSolicitud.ActualizarCapAutSinTarjetaVisto_nMovNro gsCodUser, gsCodAge, txtCuentaCargo.NroCuenta, rsCli!cperscod, nMovNroOperacion, CStr(gServGiroApertCargoCta)
    End If
    Set oMovOperacion = Nothing
    nRespuesta = 0
End If
'CTI4 end

 
MsgBox "Cuenta de Giro Generado N° " & sCuentaGiro, vbInformation, "Aviso"
ClearScreen
cboMoneda.SetFocus

Set loLavDinero = Nothing
'INICIO JHCU ENCUESTA 16-10-2019
Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
'FIN
Exit Sub
ErrGraba:
    MsgBox err.Description, vbExclamation, "Error"
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdSeguridad_Click() 'MODIFICACIÓN POR JHCU 30-05-2019
    
    'lblClave.Caption = InputBox("Ingrese la clave de seguridad.", "SICMACM - Giros")
    Dim bIngresoPass As Boolean 'add jhcu
    'lsNumTarjGir = "4697100000000025" 'OBSERVACIÓN TARJETA JHCU COMENTADO
    lsNumTarjGir = ParamNumTarjeta 'ADD JHCU
    lblClave.Caption = "" 'ADD JHCU
    lsPINBlock = ""
    lsPINBlockConf = ""
    bIngresoPass = False 'ADD JHCU
    If lsNumTarjGir = "" Then
        MsgBox "No existe número de Tarjeta de la CMACM, Por Favor Comunicarse con el área de TI", vbError, "SICMACM - Operaciones" 'ADD JHCU
        Exit Sub
    End If
    
    lsPINBlock = Tarjeta.PedirPinEnc(lsNumTarjGir, gNMKPOS, gWKPOS, 1, gnPinPadPuerto)
    
    If lsPINBlock <> "" And Len(lsPINBlock) = 16 Then
        'lblClave.Caption = lsPINBlock 'comentado por JHCU 30-05-2019
        'MsgBox "Clave Ingresada", vbInformation, "SICMACM - Operaciones" 'COMENTADO POR JHCU 30-05-2019
        MsgBox "Por favor ingrese nuevamente la contraseña para su confirmación", vbInformation, "SICMACM - Operaciones" 'ADD JHCU
        bIngresoPass = True
    Else
        'Me.lblClave.Caption = "NO INGRESADO"
        MsgBox lsPINBlock & "con el PINPAD", vbInformation, "SICMACM - Operaciones"
        Exit Sub 'ADD JHCU 30-05-2019
    End If
    If bIngresoPass = True Then
          lsPINBlockConf = Tarjeta.PedirPinEnc(lsNumTarjGir, gNMKPOS, gWKPOS, 1, gnPinPadPuerto)
          If lsPINBlockConf <> "" And Len(lsPINBlockConf) = 16 Then
              If lsPINBlockConf = lsPINBlock Then
              lblClave.Caption = lsPINBlock
              MsgBox "Clave Ingresada", vbInformation, "SICMACM - Operaciones"
              Else
              MsgBox "Las contraseñas no coinciden, por favor ingrese la contraseña nuevamente", vbInformation, "SICMACM - Operaciones"
              bIngresoPass = False
              cmdSeguridad_Click
              End If
          Else
              bIngresoPass = False
              MsgBox lsPINBlockConf & "con el PINPAD", vbInformation, "SICMACM - Operaciones"
          End If
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = "Servicios - Giro - Apertura"
    sOperacion = "Servicios - Giro Apertura"
    fgITFParamAsume gsCodAge, gGiro
    lblComision.Caption = "0.00" 'Cambio de Posición NAGL 20181010
    lblITF.Caption = "0.00" 'Cambio de Posición NAGL 20181010
    txtMonto.Text = "0.00"
    txtMontoGiro.Text = "0.00" 'NAGL 20181010
    nMontoConITF = 0 'NAGL 20181010
    nMontoSinITF = 0 'NAGL 20181010
    If gbITFAsumidoGiros Then
        chkITFEfectivo.value = 1
        chkITFEfectivo.Enabled = False
    Else
         If gsOpeCod = gServGiroApertEfec Then
            chkITFEfectivo.value = 1
        Else
            chkITFEfectivo.value = 0
        End If 'Condicional Agregado by NAGL Según RFC1807260001
        chkITFEfectivo.Enabled = True
    End If

    CargaParametros
    txtMonto.Text = ParamMinSoles
    CargaTarifario
    CargaAgencias
    CambiaTamañoCombo cboAgencia, 210
    IniciaCombo cboMoneda, gMoneda
    IniciaCombo cboTipo, COMDConstantes.gProductoCuentaTipo
    'By Capi 13122007 para deshabilitar el combo
    cboTipo.Enabled = False
    txtCuenta.Enabled = False
    txtCuenta.CMAC = gsCodCMAC
    txtCuenta.Prod = gGiro
    txtCuenta.Age = gsCodAge
    tabRemDest.Tab = 0
    lblClave.Caption = ""
    lblClave.Visible = False
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    fnPersRealizaOpe = False 'WIOR 20121015

    Call CargaControles 'CTI4 ERS0112020
    CmbForPag.ListIndex = 0 'CTI7 OPEv2
End Sub

Private Sub tabRemDest_Click(PreviousTab As Integer)
   txtDNI.Text = ""
   txtDNI.Text = SNroIdentidad
End Sub

Private Sub txtDNI_EmiteDatos()

    If Trim(txtDNI.sPersNroDoc) = "" Then
        MsgBox "Esta persona no tiene un documento de identidad ingresado." & vbCrLf & " Por favor actualice su información.", vbOKOnly + vbInformation, "Atención"
        txtDNI.Text = ""
        Exit Sub
    End If

    If txtDNI.Text <> "" And txtDNI.sPersNroDoc <> "" Then
        lblDireccion = txtDNI.sPersDireccion
        lblNombre = txtDNI.psDescripcion
        lblFecNac = Format$(txtDNI.dFecNac, "DD MMM YYYY")
        sCodRemitente = txtDNI.Text
        SNroIdentidad = txtDNI.sPersNroDoc
        txtDNI.Text = txtDNI.sPersNroDoc
        
    
        Dim dlsMant As COMDCaptaGenerales.DCOMCaptaGenerales
        Set dlsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
        fnPersPersoneriaGen = IIf(txtDNI.PersPersoneria > 1, 2, 1) 'WIOR 20121015
        fcPersCod = Trim(txtDNI.psCodigoPersona) 'WIOR 20121015
        
        If dlsMant.GetNroOPeradoras(gsCodAge) > 1 Then
            If txtDNI.psCodigoPersona = gsCodPersUser Then
                MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                Unload Me
                Exit Sub
            End If
        End If
    
        Set dlsMant = Nothing
    
        tabRemDest.Tab = 1
        cmdAgregar.SetFocus
    End If
End Sub

Private Sub txtIdAut_KeyPress(KeyAscii As Integer)
Dim rs As ADODB.Recordset
Dim nmoneda As COMDConstantes.Moneda
Dim oCap As COMDCaptaGenerales.COMDCaptAutorizacion
nmoneda = CLng(Right(cboMoneda.Text, 1))
   Set rs = New ADODB.Recordset
   Dim Nsoperacion As String
   Nsoperacion = COMDConstantes.gServGiroApertEfec
        
   If txtDNI.psCodigoPersona <> "" Then
       Gtitular = txtDNI.psCodigoPersona
   Else
       MsgBox "Se debe seleccionar el remitente del giro", vbOKOnly + vbInformation, "Atención"
       Exit Sub
   End If
   
   
   If KeyAscii = 13 And Trim(txtIdAut.Text) <> "" Then
      Set oCap = New COMDCaptaGenerales.COMDCaptAutorizacion
         Set rs = oCap.SAA(Left(CStr(Nsoperacion), 4) & "00", Vusuario, "", Gtitular, CInt(nmoneda), CLng(Val(txtIdAut.Text)))
      Set oCap = Nothing
      If rs.State = 1 Then
        If rs.RecordCount > 0 Then
            txtMonto.Text = rs!nMontoAprobado
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

Private Sub txtMonto_Change()
    Dim vComision As Double
    Dim sMoneda As String
    sMoneda = Right(cboMoneda.Text, 1)
    
    If sMoneda <> "" Then
    
        If txtMonto.value > 0 Then
        'vComision = Format$(GetComision(txtMonto.value, sMoneda), "#,##0.00")
            'By Capi 21052008 a pedido de la gerencia
                Dim lnTCFijoDia As Double
                Dim lnMontoGiro As Double
                lnTCFijoDia = Val(lblTC.Caption)
                lnMontoGiro = Val(txtMonto.value)
                
                 '***Comentado by NAGL Según 'RFC1807260001'
                'If sMoneda = 2 Then
                    'lnMontoGiro = Round(lnMontoGiro * lnTCFijoDia, 2)
                'End If
                'vComision = GetComision(txtMonto.value, sMoneda)
                'vComision = GetComision(lnMontoGiro, 1)
                'If sMoneda = 2 Then
                    'vComision = Round(vComision / lnTCFijoDia, 2)
                'End If
                'vComision = GetComision(txtMonto.value, sMoneda)
                '*********************************************'
            vComision = GetComision(lnMontoGiro, sMoneda) 'Agregado by NAGL Según RFC1807260001
            If sMoneda = 1 Then
                If vComision < ParamMinComSoles Then
                    vComision = ParamMinComSoles
                End If
            Else
                If vComision < ParamMinComDolares Then
                    vComision = ParamMinComDolares
                End If
                If Trim(lblTC.Caption) <> "" Then
                    Me.lblComSoles = Format$(vComision * Val(lblTC.Caption), "#,##0.00")
                End If
            End If
            lblComision.Caption = Format$(vComision, "#,##0.00")
            
            Call CalculaITF 'NAGL Según RFC1807260001
            
            '*********Comentado by NAGL para pasarlo al método CalculaITF con adecuaciones adicionales
            'ITF
            'Dim oITF As New COMDConstSistema.FCOMITF
            'If oITF.gbITFAplica Then
                'If txtMonto.value > gnITFMontoMin Then
                    'Me.lblITF.Caption = Format(fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                    'If chkITFEfectivo.value = 1 Then
                        'lblTotal.Caption = Format(CDbl(txtMonto.value) + CDbl(lblComision.Caption), "#,##0.00")
                    'Else
                        'lblTotal.Caption = Format(CDbl(txtMonto.value) + CDbl(lblComision.Caption) - CDbl(lblITF.Caption), "#,##0.00")
                    'End If
                'Else
                    'Me.lblITF.Caption = "0.00"
                    'lblTotal.Caption = Format(CDbl(txtMonto.value) + CDbl(lblComision.Caption), "0.00")
                'End If
                '*** BRGO 20110908 ************************************************
                'nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
                'If nRedondeoITF > 0 Then
                    'Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
                    'Me.lblTotal.Caption = Format(CCur(Me.lblITF.Caption) + txtMonto.value, "#,##0.00")
                'End If
                '*** END BRGO
            'Else
                'Me.lblITF.Caption = "0.00"
                'lblTotal.Caption = Format(CDbl(txtMonto.value) + CDbl(lblComision.Caption), "0.00")
            'End If
             '**************************END NAGL 20181010**********************************************
        Else
            lblComision = "0.00"
            lblITF.Caption = "0.00"
            lblTotal.Caption = "0.00"
            txtMontoGiro.Text = "0.00" 'NAGL 20181027
        End If
        
    End If
End Sub

Public Sub CalculaITF()
Dim oITF As New COMDConstSistema.FCOMITF
If gbITFAplica Then
    If txtMonto.value > gnITFMontoMin Then
        Me.lblITF.Caption = Format(fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
        '*** BRGO 20110908 ************************************************
        nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
        If nRedondeoITF > 0 Then
            Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
            'Me.lblTotal.Caption = Format(CCur(Me.lblITF.Caption) + txtMonto.value, "#,##0.00")
        End If
        '*** END BRGO
        If txtMontoGiro.Text <> nMontoConITF And txtMontoGiro.Text <> nMontoSinITF Then
            chkITFEfectivo.value = 1
        End If 'NAGL Según RFC1807260001
        If chkITFEfectivo.value = 1 Then
            lblTotal.Caption = Format(CDbl(txtMonto.value) + CDbl(lblComision.Caption) + CDbl(lblITF.Caption), "#,##0.00")
            txtMontoGiro.Text = Format(CDbl(txtMonto.value), "#,#0.00") 'NAGL 20181027
        Else
            lblTotal.Caption = Format(CDbl(txtMonto.value) + CDbl(lblComision.Caption), "#,##0.00")
            txtMontoGiro.Text = Format(CDbl(txtMonto.value) - CDbl(lblITF.Caption), "#,#0.00") 'NAGL 20181027
        End If
    Else
        Me.lblITF.Caption = "0.00"
        lblTotal.Caption = Format(CDbl(txtMonto.value) + CDbl(lblComision.Caption), "0.00")
        txtMontoGiro.Text = Format(CDbl(txtMonto.value), "#,#0.00") 'NAGL 20181027
    End If

Else
    Me.lblITF.Caption = "0.00"
    lblTotal.Caption = Format(CDbl(txtMonto.value) + CDbl(lblComision.Caption), "0.00")
End If
End Sub 'NAGL Según RFC1807260001

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtDNI.SetFocus
   Dim sMoneda As String
   sMoneda = Right(cboMoneda.Text, 1)
   If sMoneda <> "" Then
            If sMoneda = 1 Then
                 If txtMonto.value < ParamMinSoles Then
                    txtMonto.value = ParamMinSoles
                    MsgBox "El monto mínimo en giros de soles es de " & ParamMinSoles
                    
                 ElseIf txtMonto.value > ParamMaxSoles Then
                    txtMonto.value = ParamMaxSoles
                    MsgBox "El monto máximo en giros de soles es de " & ParamMaxSoles
                    
                 End If
            Else
                 If txtMonto.value < ParamMinDolares Then
                     txtMonto.value = ParamMinDolares
                     MsgBox "El monto mínimo en giros de dólares es de " & ParamMinDolares
                     
                 ElseIf txtMonto.value > ParamMaxDolares Then
                     txtMonto.value = ParamMaxDolares
                     MsgBox "El monto máximo en giros de dólares es de " & ParamMaxDolares
                     
                 End If
            End If
    End If
    If CCur(Trim(txtMonto.Text)) = "0" Then
        CmbForPag.Enabled = False
    Else
        CmbForPag.Enabled = True
    End If
End If
End Sub


Private Sub txtMonto_LostFocus()
Dim sMoneda As String
   sMoneda = Right(cboMoneda.Text, 1)
 If sMoneda <> "" Then
   If sMoneda = 1 Then
        If txtMonto.value < ParamMinSoles Then
           txtMonto.value = ParamMinSoles
           MsgBox "El monto mínimo en giros de soles es de " & ParamMinSoles
           
        ElseIf txtMonto.value > ParamMaxSoles Then
           txtMonto.value = ParamMaxSoles
           MsgBox "El monto máximo en giros de soles es de " & ParamMaxSoles
           
        End If
   Else
        If txtMonto.value < ParamMinDolares Then
            txtMonto.value = ParamMinDolares
            MsgBox "El monto mínimo en giros de dólares es de " & ParamMinDolares
            
        ElseIf txtMonto.value > ParamMaxDolares Then
            txtMonto.value = ParamMaxDolares
            MsgBox "El monto máximo en giros de dólares es de " & ParamMaxDolares
            
        End If
   End If
 End If
    If CCur(Trim(txtMonto.Text)) = "0" Then
        CmbForPag.Enabled = False
    Else
        CmbForPag.Enabled = True
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
'CTI7*********************************************************************************
Private Sub CmbForPag_Click()
    EstadoFormaPago IIf(CmbForPag.ListIndex = -1, -1, CInt(Trim(Right(IIf(CmbForPag.Text = "", "-1", CmbForPag.Text), 10))))
    If CmbForPag.ListIndex <> -1 Then
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher Then
     
            Dim oformVou As New frmCapRegVouDepBus
            Dim lnTipMot As Integer
            Dim sGlosa As String
            Dim sIF As String
            Dim sVaucher As String
            Dim sPersCod As String
            Dim sNombre As String
            Dim sDireccion As String
            Dim sDocumento As String
                        
            lnTipMot = 22 ' Giro apertura
            oformVou.iniciarFormularioDeposito CLng(Right(cboMoneda, 1)), lnTipMot, sGlosa, sIF, sVaucher, nMontoVoucher, sPersCod, nMovNroRVD, nMovNroRVDPend, sNombre, sDireccion, sDocumento, ""
            If Len(sVaucher) = 0 Then Exit Sub
            LblNumDoc.Caption = Trim(Mid(sVaucher, 1, Len(sVaucher) - 10))
            cmdGrabar.Enabled = True
        ElseIf CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            Dim sCuenta As String
            
            sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(gServGiroApertCargoCta), sNumTarj, 232, False)
            If Val(Mid(sCuenta, 6, 3)) <> "232" And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            If Len(sCuenta) = 18 Then
                If CInt(Mid(sCuenta, 9, 1)) <> CLng(Right(cboMoneda, 1)) Then
                    MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la cuenta a amortizar.", vbOKOnly + vbInformation, App.Title
                End If
            Else
                        MsgBox "La cuenta de ahorro no es válida.", vbOKOnly + vbInformation, App.Title
            End If
            If Len(sCuenta) = 0 Then EnfocaControl txtCuentaCargo: Exit Sub
            txtCuentaCargo.NroCuenta = sCuenta
            txtCuentaCargo.Enabled = False
            txtMonto_Change
            cmdGrabar.Enabled = True
            cmdGrabar.SetFocus
        End If
    End If
End Sub
Private Sub EstadoFormaPago(ByVal nFormaPago As Integer)
    LblNumDoc.Caption = ""
    txtCuentaCargo.NroCuenta = ""
    Select Case nFormaPago
        Case -1
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargo.Visible = False
            cmdGrabar.Enabled = True
        Case gColocTipoPagoEfectivo
            txtCuentaCargo.Visible = False
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            lblNroDocumento.Visible = False
            cmdGrabar.Enabled = True
        Case gColocTipoPagoCargoCta
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargo.Visible = True
            txtCuentaCargo.Enabled = True
            txtCuentaCargo.CMAC = gsCodCMAC
            txtCuentaCargo.Prod = Trim(Str(gCapAhorros))
            cmdGrabar.Enabled = False
        Case gColocTipoPagoVoucher
            LblNumDoc.Visible = True
            lblNroDocumento.Visible = True
            txtCuentaCargo.Visible = False
            cmdGrabar.Enabled = False
    End Select
End Sub
Private Function ValidaFormaPago() As Boolean
Dim clsCap As New COMNCaptaGenerales.NCOMCaptaMovimiento
    ValidaFormaPago = False
    If CmbForPag.ListIndex = -1 Then
        MsgBox "No se ha seleccionado la forma de pago. Verifique.", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher And Len(Trim(LblNumDoc.Caption)) = 0 Then
        MsgBox "No se ha seleccionado el voucher correctamente. Verifique.", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher And Len(Trim(LblNumDoc.Caption)) > 0 _
        And CCur(lblTotal.Caption) <> CCur(nMontoVoucher) Then
        MsgBox "El Monto de Transacción debe ser igual al Monto Total. Verifique.", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
    
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta And Len(txtCuentaCargo.NroCuenta) <> 18 Then
        MsgBox "Favor de ingresar la cuenta de ahorros a debitar", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
        
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
        If Not clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, CDbl(txtMonto.Text)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
            Exit Function
        End If
        
        Dim rsCli As New ADODB.Recordset
        Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
        Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
        Dim bExitoSol As Integer
        Dim nRespuesta As Integer
        Dim psPersCodTitularAhorroCargo As String
        Set rsCli = clsCli.GetPersonaCuenta(txtCuentaCargo.NroCuenta, gCapRelPersTitular)
        psPersCodTitularAhorroCargo = rsCli!cperscod
        
        If psPersCodTitularAhorroCargo = txtDNI Then
            MsgBox "El titular de la cuenta de ahorro no es la misma persona quien solicita la operación", vbInformation, "Aviso"
            Exit Function
        End If
        
    End If
    
    ValidaFormaPago = True
End Function
Private Sub txtCuentaCargo_KeyPress(KeyAscii As Integer)
    Dim sMsg As String
    If KeyAscii = 13 Then sMsg = ValidaCuentaACargo(txtCuentaCargo.NroCuenta)
    If Len(sMsg) > 0 Then
        MsgBox sMsg, vbInformation, "¡Aviso!"
        txtCuentaCargo.SetFocus
        Exit Sub
    End If
    If Len(txtCuentaCargo.NroCuenta) = 18 Then
        If CInt(Mid(txtCuentaCargo.NroCuenta, 9, 1)) <> CLng(Right(cboMoneda, 1)) Then
            MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la cuenta a amortizar.", vbOKOnly + vbInformation, App.Title
        End If
    End If
    ObtieneDatosCuenta txtCuentaCargo.NroCuenta
End Sub
Private Function ValidaCuentaACargo(ByVal psCuenta As String) As String
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sMsg As String
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    sMsg = clsCap.ValidaCuentaOperacion(psCuenta)
    ValidaCuentaACargo = sMsg
End Function
Private Sub ObtieneDatosCuenta(ByVal psCuenta As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsCta As ADODB.Recordset
    Dim rsV As ADODB.Recordset
    Dim lnTpoPrograma As Integer
    Dim lsTieneTarj As String
    Dim lbVistoVal As Boolean
    
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsV = New ADODB.Recordset
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(psCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
        If sNumTarj = "" Then
            If Mid(psCuenta, 6, 3) = "232" And lnTpoPrograma <> 1 Then
                Set rsV = clsMant.ValidaTarjetizacion(psCuenta, lsTieneTarj)
                If rsV.RecordCount > 0 Then
                    Dim tipoCta As Integer
                    tipoCta = rsCta("nPrdCtaTpo")
                    If tipoCta = 0 Or tipoCta = 2 Then
                        Dim rsCli As New ADODB.Recordset
                        Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
                        Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
                        Dim bExitoSol As Integer
                        Set rsCli = clsCli.GetPersonaCuenta(psCuenta, gCapRelPersTitular)
                        nRespuesta = oSolicitud.SolicitarVistoAtencionSinTarjeta(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gServGiroApertCargoCta))
                    
                        If nRespuesta = 1 Then '1:Tiene Visto de atencion sin tarjeta pendiente de autorizar
                             MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                             Exit Sub
                        End If
                        If nRespuesta = 2 Then '2:Tiene visto aceptado
                            MsgBox "La solicitud de atención sin tarjeta fue Aprobada, proceda con la atención", vbInformation, "Aviso"
                        End If
                        If nRespuesta = 3 Then '3:Tiene visto rechazado
                           If MsgBox("La solicitud de atención sin tarjeta fue RECHAZADA. ¿Desea realizar una nueva solicitud?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                Call oSolicitud.RegistrarVistoDeUsuario(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gServGiroApertCargoCta))
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        If nRespuesta = 4 Then '4:Se permite registrar la solicitud
                            Dim mensaje As String
                            If lsTieneTarj = "SI" Then
                                mensaje = "El Cliente posee tarjeta. Para continuar deberá registrar el Motivo de Autorización y comunicar al Coordinador o Jefe de Operaciones para su Aprobación. ¿Desea Continuar?"
                            Else
                                mensaje = "El Cliente NO posee tarjeta activa. Para continuar deberá registrar el Motivo de Autorización y comunicar al Coordinador o Jefe de Operaciones para su Aprobación. ¿Desea Continuar?"
                            End If
                        
                            If MsgBox(mensaje, vbInformation + vbYesNo, "Aviso") = vbYes Then
                        
                                bExitoSol = frmCapMotivoAutorizacion.Inicio(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gServGiroApertCargoCta))
                                If bExitoSol > 0 Then
                                    MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                                End If
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        lbVistoVal = loVistoElectronico.Inicio(5, gServGiroApertCargoCta)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema"
                            Exit Sub
                        End If
                    End If
                ElseIf lsTieneTarj = "NO" And rsV.RecordCount > 0 Then
                    If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? ", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg ers 065-2017
                        lbVistoVal = loVistoElectronico.Inicio(5, gServGiroApertCargoCta)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema" 'ADD BY MARG ERS065-2017
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            End If
        End If
        txtCuentaCargo.Enabled = False
        txtMonto_Change
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
    End If
End Sub
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R = oCons.RecuperaConstantes(gColocTipoPago, , , 4)
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, CmbForPag)
    Set loVistoElectronico = New frmVistoElectronico
    Exit Sub
ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub
'*************************************************************************************


