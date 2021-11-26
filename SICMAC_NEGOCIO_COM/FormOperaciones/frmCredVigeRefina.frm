VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredRefinancVigencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vigencia de Creditos Refinanciados"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   0
      TabIndex        =   3
      Top             =   835
      Width           =   7320
      Begin TabDlg.SSTab SSTabDatosRefina 
         Height          =   4695
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   8281
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Cliente"
         TabPicture(0)   =   "frmCredVigeRefina.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "FraCliente"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Crédito"
         TabPicture(1)   =   "frmCredVigeRefina.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "FraCredito"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Vigencia"
         TabPicture(2)   =   "frmCredVigeRefina.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).Control(1)=   "Frame1"
         Tab(2).ControlCount=   2
         Begin VB.Frame Frame3 
            Caption         =   "Créditos a Cancelar:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Left            =   -74880
            TabIndex        =   49
            Top             =   360
            Width           =   6615
            Begin SICMACT.FlexEdit FeVigenciaRefina 
               Height          =   1455
               Left            =   240
               TabIndex        =   50
               Top             =   240
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   2566
               Cols0           =   6
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "-Credito-Capital-Interes-Gastos-Total"
               EncabezadosAnchos=   "10-1800-1000-1000-1000-1300"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
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
               ColumnasAEditar =   "X-X-X-X-X-X"
               TextStyleFixed  =   4
               ListaControles  =   "0-0-0-0-0-0"
               BackColor       =   12189695
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-C-C-C-C-C"
               FormatosEdit    =   "0-0-0-0-0-0"
               lbUltimaInstancia=   -1  'True
               ColWidth0       =   15
               RowHeight0      =   300
               ForeColorFixed  =   -2147483635
               CellBackColor   =   12189695
            End
            Begin VB.Label lblTotal 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
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
               Height          =   285
               Left            =   4800
               TabIndex        =   52
               Top             =   1800
               Width           =   1695
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL :"
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
               Height          =   195
               Left            =   4080
               TabIndex        =   51
               Top             =   1800
               Width           =   630
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Procesar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   -74880
            TabIndex        =   37
            Top             =   2640
            Width           =   6615
            Begin VB.CommandButton cmdSalir 
               Caption         =   "&Salir"
               Enabled         =   0   'False
               Height          =   495
               Left            =   4680
               TabIndex        =   48
               Top             =   1320
               Width           =   1455
            End
            Begin VB.CommandButton CmdCancelar 
               Caption         =   "&Cancelar"
               Enabled         =   0   'False
               Height          =   495
               Left            =   4680
               TabIndex        =   47
               Top             =   780
               Width           =   1455
            End
            Begin VB.CommandButton CmdVigencia 
               Caption         =   "&Vigencia"
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
               Height          =   495
               Left            =   4680
               TabIndex        =   46
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Monto Total :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   300
               TabIndex        =   45
               Top             =   1440
               Width           =   1065
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Gastos :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   720
               TabIndex        =   44
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Intereses  :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   500
               TabIndex        =   43
               Top             =   600
               Width           =   825
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Capital :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   700
               TabIndex        =   42
               Top             =   240
               Width           =   630
            End
            Begin VB.Label LblMontoRefinaTotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
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
               Height          =   285
               Left            =   1680
               TabIndex        =   41
               Top             =   1440
               Width           =   1455
            End
            Begin VB.Label LblMonGastos 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
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
               Height          =   285
               Left            =   1680
               TabIndex        =   40
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label lblIntereses 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
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
               Height          =   285
               Left            =   1680
               TabIndex        =   39
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label lblCapital 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
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
               Height          =   285
               Left            =   1680
               TabIndex        =   38
               Top             =   240
               Width           =   1455
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00C00000&
               BorderWidth     =   2
               X1              =   1440
               X2              =   3180
               Y1              =   1320
               Y2              =   1320
            End
         End
         Begin VB.Frame FraCredito 
            Height          =   3615
            Left            =   360
            TabIndex        =   18
            Top             =   480
            Width           =   6180
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Apoderado :"
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
               Left            =   240
               TabIndex        =   36
               Top             =   2760
               Width           =   960
            End
            Begin VB.Label Label10 
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
               Height          =   195
               Left            =   480
               TabIndex        =   35
               Top             =   2400
               Width           =   705
            End
            Begin VB.Label lblApoderado 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   34
               Top             =   2760
               Width           =   4665
            End
            Begin VB.Label lblAnalista 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   33
               Top             =   2400
               Width           =   4665
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Aprobación :"
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
               Left            =   2880
               TabIndex        =   32
               Top             =   1845
               Width           =   1725
            End
            Begin VB.Label lblFechaApro 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   4680
               TabIndex        =   31
               Top             =   1800
               Width           =   1335
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Tasa :"
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
               Left            =   720
               TabIndex        =   30
               Top             =   1800
               Width           =   420
            End
            Begin VB.Label lblTasa 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   29
               Top             =   1800
               Width           =   1335
            End
            Begin VB.Label lblNroCuotas 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   4680
               TabIndex        =   28
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Nro. Cuotas :"
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
               Left            =   3600
               TabIndex        =   27
               Top             =   1440
               Width           =   1005
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Moneda :"
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
               Left            =   480
               TabIndex        =   26
               Top             =   1440
               Width           =   735
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Linea Credito :"
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
               Left            =   120
               TabIndex        =   25
               Top             =   1080
               Width           =   1110
            End
            Begin VB.Label LblLineaCred 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1335
               TabIndex        =   24
               Top             =   1080
               Width           =   4665
            End
            Begin VB.Label LblMoneda 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   23
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label lblTipoProd 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1335
               TabIndex        =   22
               Top             =   720
               Width           =   4665
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Producto :"
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
               Left            =   120
               TabIndex        =   21
               Top             =   720
               Width           =   1170
            End
            Begin VB.Label lblTipoCred 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1335
               TabIndex        =   20
               Top             =   360
               Width           =   4665
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Credito :"
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
               Left            =   240
               TabIndex        =   19
               Top             =   360
               Width           =   1035
            End
         End
         Begin VB.Frame FraCliente 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   5
            Top             =   480
            Width           =   6885
            Begin SICMACT.FlexEdit FECreditosVig 
               Height          =   1470
               Left            =   240
               TabIndex        =   6
               Top             =   2280
               Width           =   6465
               _ExtentX        =   11404
               _ExtentY        =   2593
               Cols0           =   6
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "--Credito-Atraso-Monto-PagAmp"
               EncabezadosAnchos=   "0-0-3000-1500-1500-0"
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
               ColumnasAEditar =   "X-1-X-X-X-X"
               TextStyleFixed  =   3
               ListaControles  =   "0-4-0-0-0-0"
               BackColor       =   14286847
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-R-R-R-R-L"
               FormatosEdit    =   "0-0-3-3-2-0"
               lbUltimaInstancia=   -1  'True
               lbPuntero       =   -1  'True
               RowHeight0      =   300
               ForeColorFixed  =   -2147483635
               CellBackColor   =   14286847
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Codigo :"
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
               Left            =   360
               TabIndex        =   17
               Top             =   360
               Width           =   660
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nombre :"
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
               Left            =   360
               TabIndex        =   16
               Top             =   720
               Width           =   705
            End
            Begin VB.Label LblCodCli 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1170
               TabIndex        =   15
               Top             =   300
               Width           =   1410
            End
            Begin VB.Label LblNomCli 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1170
               TabIndex        =   14
               Top             =   675
               Width           =   5445
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "RUC :"
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
               Left            =   600
               TabIndex        =   13
               Top             =   1095
               Width           =   420
            End
            Begin VB.Label LblDocNat 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   5160
               TabIndex        =   12
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Doc. de Identidad :"
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
               Left            =   3600
               TabIndex        =   11
               Top             =   1095
               Width           =   1470
            End
            Begin VB.Label LblDocJur 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1170
               TabIndex        =   10
               Top             =   1080
               Width           =   1440
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Créditos Vigentes en Agencia :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   360
               TabIndex        =   9
               Top             =   2040
               Width           =   2370
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Direccion :"
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
               Left            =   240
               TabIndex        =   8
               Top             =   1485
               Width           =   810
            End
            Begin VB.Label LblCliDirec 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1170
               TabIndex        =   7
               Top             =   1455
               Width           =   5445
            End
         End
      End
   End
   Begin VB.Frame fraCuenta 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton cmdExaminar 
         Caption         =   "&Examinar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   525
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   926
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmCredRefinancVigencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:               frmCredVigeRefina (Vigencia de Creditos Refinanciados)
'***     Descripcion:          Opcion que da vigencia a los creditos que han sido refinanciados, despues de ser Aprobados (estado 2030)
'***     Creado por:           LUCV
'***     Maquina:              01TIF55
'***     Fecha-Tiempo:         16/05/2016 10:52:16 AM
'***     Ultima Modificacion:  Fecha de Crecion
'***     Referencia :          ERS004-2016(Refinanciacion de creditos) - Operaciones
'*****************************************************************************************
Option Explicit
Private nMontoCapital As Double
Private nInteres As Double
Private nMontoGastos As Double
Private nTotal As Double

Private MatCredCanc(100, 2) As String
Private ContMatCredCanc As Integer
Private sCtaAho As String
Private nNroProxDesemb As Integer
Private vbDesembCC As Boolean
Private pRSRela As ADODB.Recordset
Private pnTasa As Double
Private pnPersoneria As Integer
Private pnTipoCuenta As Integer

Private pbDocumento As Boolean
Private psNroDoc As String
Private psCodIF As String
Private sOperacion As String
Private sPersCod As String

'Variables agregadas para el uso de los Componentes
Private pbOperacionEfectivo As Boolean
Private pnMontoLavDinero As Double
Private pnTC As Double
Private pbExoneradaLavado As Boolean
Private psPersCodRep As String  'Codigo del Representante del Crédito
Private psPersNombreRep As String 'Nombre del Representante de Crédito
Public rsRel As ADODB.Recordset

'Dim nTasaInt As DObjeto
Dim MatTitulares As Variant
Dim nProgAhorros As Integer
Dim nMontoAbonar As Double
Dim nPlazoAbonar As Integer
Dim sPromotorAho As String

Dim lbPuedeAperturar As Boolean
Dim nRedondeoITF As Double
Dim sTpoProdCod As String
Dim rsRelEmp As ADODB.Recordset
Dim nMontoPrestamoW As Double
Dim nDestinoCred As Integer
Dim bInstFinanc As Boolean
Dim bOnCellCheck As Boolean
Dim bRevisaDesemb As Boolean

'Variables que Quedan
Dim bVigeRefina As Boolean
Dim bLeasing As Boolean
Dim nMontoRefina As Double
Dim nPlazoRefina As Integer
Dim bCapitalInt As Boolean 'lucv
Dim nNroCalen As Integer
Dim nTotalARefinanciar As Double
Dim bRefinanc As Boolean
Dim objPista As COMManejador.Pista

'________________________________________________________________________________________________________________
'******************************************** EVENTOS - VARIOS **************************************************
Private Sub CmdVigencia_Click()
    Dim oNCredito As COMNCredito.NCOMCredito
    Dim oCredPers As COMDCredito.DCOMCredito
    Dim oDCredito As COMDCredito.DCOMCredito
    Dim rsColoEsta As ADODB.Recordset 'lucv agregado
    Dim sError As String
    Dim nTipoCuota As Integer
    Dim nPeriodoFechaFija As Integer
    Dim nPlazo As Integer
    Dim nProxMes As Integer
    Dim nTipoDesembolso As Integer
    Dim nCalendDinamico As Integer
    Dim nPeriodoGracia As Integer
    Dim nTasaGracia As Double
    Dim nTipoGracia As Integer
    Dim sTipoGasto As String
    Dim nPeriodoFechaFija2 As Integer
    Dim bIncremGraciaCap As Boolean
    Dim bGraciaEnCuotas As Boolean
    Set oNCredito = New COMNCredito.NCOMCredito
    Dim sImpreDocs As String
    Dim sMensaje As String
    
    Dim lsMensajeGrabar As String
    Dim lnDescripCuota() As String
    Dim sMensPriFecPag As String
    Dim sMensFecAprob As String
    
    Dim oPersona As COMDPersona.DCOMPersona
    Dim rsPersona As ADODB.Recordset
    
    bRefinanc = True
    sMensFecAprob = oNCredito.DevolverFechaAprobacion(ActxCta.NroCuenta)
    'tomada de desembolso***
        lnDescripCuota = oNCredito.DevolverPrimeraFechaPago(ActxCta.NroCuenta, CDbl(LblMontoRefinaTotal.Caption), gdFecSis, sMensFecAprob, "N")
        sMensPriFecPag = lnDescripCuota(2)
    '***
    If Format(lblFechaApro.Caption, "YYYY/MM/DD") < gdFecSis Then
        MsgBox "La aprobación del crédito fue realizada el " & Format(lblFechaApro.Caption, "YYYY/MM/DD") & " " & Chr(13) & "la Operación no puede continuar" & Chr(13) & "Contactarse con el area de créditos para volver a sugerir/aprobar", vbInformation, "Aviso"
        Exit Sub
    End If
    'Recuperar Datos del los Cred. Aprobados (A refinanciar - Estado Vigencia)
    Set oDCredito = New COMDCredito.DCOMCredito
    ActxCta.NroCuenta = Replace(ActxCta.NroCuenta, "'", "")
    Set rsColoEsta = oDCredito.RecuperaColocacEstado(ActxCta.NroCuenta, gColocEstAprob)
    Set oDCredito = Nothing
     If Not rsColoEsta.EOF Or rsColoEsta.BOF Then
        nTipoCuota = rsColoEsta!nColocCalendCod
        nPeriodoFechaFija = rsColoEsta!nPeriodoFechaFija
        nPeriodoGracia = rsColoEsta!nPeriodoGracia
        nPlazo = rsColoEsta!nPlazo
        nProxMes = rsColoEsta!nProxMes
        nTipoDesembolso = rsColoEsta!nTipoDesembolso
        nCalendDinamico = rsColoEsta!nCalendDinamico
        nPeriodoGracia = rsColoEsta!nPeriodoGracia
        nTasaGracia = IIf(IsNull(rsColoEsta!nTasaGracia), 0, rsColoEsta!nTasaGracia)
        nTipoGracia = rsColoEsta!nTipoGracia
        sTipoGasto = rsColoEsta!cTipoGasto
        nPeriodoFechaFija2 = IIf(IsNull(rsColoEsta!nPeriodoFechaFija2), 0, rsColoEsta!nPeriodoFechaFija2)
        bIncremGraciaCap = IIf(IsNull(rsColoEsta!bIncremGraciaCap), -1, rsColoEsta!bIncremGraciaCap)
        bGraciaEnCuotas = rsColoEsta!bGraciaEnCuotas
   End If

  'Verifica actualización Persona
   Dim oNCOMPersona As New COMNPersona.NCOMPersona
   If oNCOMPersona.NecesitaActualizarDatos(LblCodCli.Caption, gdFecSis) Then
        MsgBox "Para continuar con la Operación Ud. debe actualizar los datos del" & Chr(13) & "Titular: " & LblNomCli.Caption, vbInformation, "Aviso"
        Dim foPersona As New frmPersona
        If Not foPersona.realizarMantenimiento(LblCodCli.Caption) Then
            MsgBox "No se ha realizado la actualización de los datos de " & LblNomCli.Caption & "," & Chr(13) & "la Operación no puede continuar!", vbInformation, "Aviso"
            Exit Sub
        End If
   End If

    If Not (ActxCta.Prod = "515" Or ActxCta.Prod = "516") Then
        lsMensajeGrabar = "El Crédito fue aprobado el"
   Else
        lsMensajeGrabar = "La operación fue aprobada el"
   End If


    If MsgBox("Se Va A Grabar los Datos, Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        If MsgBox(lsMensajeGrabar & sMensFecAprob & Chr(13) & "La Primera Fecha de Pago es: " & sMensPriFecPag & Chr(13) & Chr(13) & "            Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        'oDCredito.ActualizarColocacEstadoDiasAtraso ActxCta.NroCuenta, gColocEstAprob, nPeriodoGraciaAnt
        'Set oDCredito = Nothing
        Else
        Call oNCredito.GrabarAprobacionVigenciaRefinanciar(ActxCta.NroCuenta, _
                                                            gColocEstAprob, _
                                                            gdFecSis, _
                                                            0, _
                                                            (LblLineaCred.Caption), _
                                                            CDbl(lblTasa.Caption), _
                                                            CDbl(LblMontoRefinaTotal.Caption), _
                                                            CInt(lblNroCuotas.Caption), _
                                                            CInt(nPlazoRefina), _
                                                            nTipoCuota, _
                                                            nPeriodoFechaFija, _
                                                            nProxMes, _
                                                            nTipoDesembolso, _
                                                            nCalendDinamico, _
                                                            nPeriodoGracia, _
                                                            nTasaGracia, nTipoGracia, nNroCalen, sError, sMensaje, sImpreDocs, _
                                                            gsCodAge, gsCodUser, gsNomAge, gsNomCmac, gsInstCmac, gsCodCMAC, _
                                                            lnDescripCuota, bRefinanc, _
                                                            bCapitalInt, sTipoGasto, _
                                                            nPeriodoFechaFija2, bIncremGraciaCap, _
                                                            bGraciaEnCuotas, LblCodCli.Caption)
                                                            
           
        If sError <> "" Then
               MsgBox sError, vbInformation, "Aviso"
        Else
           Set oCredPers = New COMDCredito.DCOMCredito
           Call oCredPers.EliminarDatosNuevoMIVIVIENDA(ActxCta.NroCuenta, gColocEstAprob)
           Set objPista = New COMManejador.Pista
           
           
           objPista.InsertarPista sOperacion, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, , ActxCta.NroCuenta, gCodigoCuenta
           
           Set oPersona = New COMDPersona.DCOMPersona
           Set rsPersona = oPersona.ListaVinculadosPersonaCta(Trim(Me.ActxCta.NroCuenta))
           If rsPersona.RecordCount > 0 Then
               Call oPersona.EliminarVinculadosPersona(Trim(Me.ActxCta.NroCuenta))
           End If
           
           Set oNCredito = Nothing
           MsgBox "La vigencia del Crédito N°: " & ActxCta.NroCuenta & "," & Chr(13) & "se Registro Correctamente", vbInformation, "Aviso"
                          
        '***********************************************************************************************
        'Para la Impresion de Documentos (Calendario de pagos | Comprobante de Vigencia | Hoja de Resumen)
        '***********************************************************************************************
        Set oNCredito = Nothing
            Dim clsprevio As previo.clsprevio
            If Not (ActxCta.Prod = "515" Or ActxCta.Prod = "516") Then
                If sMensaje <> "" Then
                    MsgBox sMensaje, vbInformation, "Mensaje"
                    Exit Sub
                End If
            
                Do
                    MsgBox "Coloque Papel Continuo Tamaño Carta, Para la Impresion de los Documentos de Desembolsos", vbInformation, "Aviso"
                    Set clsprevio = New previo.clsprevio
                    clsprevio.PrintSpool sLpt, oImpresora.gPrnTpoLetraSansSerif1PDef & oImpresora.gPrnTamLetra10CPIDef & sImpreDocs, False, gnLinPage
                Loop While MsgBox("Desea Reimprimir Todos los Documentos del Desembolso?", vbInformation + vbYesNo, "Aviso") = vbYes
                    
                If Not (Mid(ActxCta.NroCuenta, 6, 3) = "515" Or Mid(ActxCta.NroCuenta, 6, 3) = "516") Then
                    MsgBox "Coloque papel para Imprimir Hoja de Resumen...", vbInformation, "Aviso"
                    Call ImprimeCartillaCred(ActxCta.NroCuenta)
                End If
                gVarPublicas.LimpiaVarLavDinero
                                'INICIO JHCU ENCUESTA 16-10-2019
                                Encuestas gsCodUser, gsCodAge, "ERS0292019", sOperacion
                                'FIN
        End If
               Call LimpiaPantalla
               CmdVigencia.Enabled = False
        End If
        End If
  End If
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    CentraForm Me
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    nMontoGastos = 0
End Sub

Private Sub cmdExaminar_Click()
    bVigeRefina = True
    bLeasing = False
    ActxCta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstAprob), "Creditos a refinanciar - Vigencia", , bVigeRefina, , gsCodAge, bLeasing)
    If ActxCta.NroCuenta <> "" Then
        Call ActxCta_KeyPress(13)
    Else
        Call LimpiaPantalla
        ActxCta.CMAC = gsCodCMAC
        ActxCta.Age = gsCodAge
        ActxCta.SetFocusProd
        ActxCta.Enabled = True
    End If
End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    Dim oCredito As COMNCredito.NCOMCredito
    Dim lafirma As frmPersonaFirma
    Dim ClsPersona As COMDPersona.DCOMPersonas
    Dim Rf As ADODB.Recordset
    Dim oCredDestino As New COMDCredito.DCOMCredito
    Dim RDes As ADODB.Recordset
    Dim sError As String

    If KeyAscii = 13 Then
'LUCV20161031, Comentó
'            If ActxCta.Prod = "517" And Not vbDesembCC Then
'                MsgBox "El crédito no se puede desembolsar en efectivo. Seleccionar Desembolso con Abono a Cuenta", vbInformation, "Aviso"
'                Call cmdCancelar_Click
'                Exit Sub
'            End If
       
       'Recupera los datos de ColocacCred
        ActxCta.NroCuenta = Replace(ActxCta.NroCuenta, "'", "")
        Set RDes = oCredDestino.RecuperaColocacCred(ActxCta.NroCuenta)
        Set oCredDestino = Nothing
        If Not (RDes.EOF And RDes.BOF) Then
            nDestinoCred = IIf(IsNull(RDes!nColocDestino), 0, RDes!nColocDestino)
            RDes.Close
            Set RDes = Nothing
                If gCredDesembEfec = sOperacion Then
                   If nDestinoCred = 14 Then
                        If MsgBox("Este tipo de crédito no puede ser desembolsado por esta modalidad," & _
                           "¿Desea ser redirigido a la modalidad desembolso abono a cuenta?", vbYesNo, "Aviso") = vbYes Then
                            Unload Me
                            Dim oform As New frmCredDesembAbonoCta
                            Call oform.DesembolsoCargoCuenta(gCredDesembCtaNueva)
                            Set oform = Nothing
                            Exit Sub
                        Else
                            Call cmdCancelar_Click
                            Exit Sub
                        End If
                   End If
                End If
        End If
        Set oCredito = New COMNCredito.NCOMCredito
        
        'Valida: Niveles de Aprobacion, No se tiene Registrado un Calendario, Aprobacion del Credito
        sError = oCredito.ValidaCargaDatosDesembolso(ActxCta.NroCuenta, gdFecSis)
        If sError <> "" Then
            MsgBox sError, vbInformation, "Aviso"
            Call cmdCancelar_Click
            Exit Sub
        End If
        Set oCredito = Nothing
        If CargaDatos(ActxCta.NroCuenta) Then
            HabilitaRefinanciacion True
                Set lafirma = New frmPersonaFirma
                Set ClsPersona = New COMDPersona.DCOMPersonas
                Set Rf = ClsPersona.BuscaCliente(frmCredPersEstado.vcodper, BusquedaCodigo)
         If Not Rf.BOF And Not Rf.EOF Then
                If Rf!nPersPersoneria = 1 Then
                Call frmPersonaFirma.Inicio(Trim(frmCredPersEstado.vcodper), Mid(frmCredPersEstado.vcodper, 4, 2), True)
                End If
         End If
         Set Rf = Nothing
        Else
            HabilitaRefinanciacion False
            Call LimpiaPantalla
                If Not bRevisaDesemb Then
                    MsgBox "No pudo Encontrar  el Credito, posiblemente aun no esta Aprobado", vbInformation, "Aviso"
                End If
        End If
    End If
End Sub

'Private Sub FECreditosVig_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
'Dim i As Integer
'Dim nRedondITFCanc As Double
'Dim nITFCtaCanc As Double
'    If bOnCellCheck Then
'        If bInstFinanc Then nITFCtaCanc = 0
'        If Trim(FECreditosVig.TextMatrix(pnRow, pnCol)) <> "." Then 'Sin Check
'            Call EliminarCreditoACancelar(FECreditosVig.TextMatrix(pnRow, 2))
'            LblItf.Caption = Format(CDbl(LblItf.Caption) - nITFCtaCanc, "#0.00")
'        Else 'Con Check
'            If CDbl(FECreditosVig.TextMatrix(pnRow, 4)) > (CDbl(LblMonPrestamo.Caption) - CDbl(LblMonGastos.Caption) - CDbl(LblMonCancel.Caption)) Then
'                MsgBox "Monto a Desembolsar no es suficiente para cancelar este Credito", vbInformation, "Aviso"
'                FECreditosVig.TextMatrix(pnRow, 1) = ""
'                Exit Sub
'            End If
'            Call AdicionaCreditoACancelar(FECreditosVig.TextMatrix(pnRow, 2), CDbl(FECreditosVig.TextMatrix(pnRow, 4)))
'            LblItf.Caption = Format(CDbl(LblItf.Caption) + nITFCtaCanc, "#0.00")
'        End If
'        nMontoCredCanc = 0
'        LstCredVig.Clear
'        For i = 0 To ContMatCredCanc - 1
'            nMontoCredCanc = nMontoCredCanc + CDbl(MatCredCanc(i, 1))
'            LstCredVig.AddItem MatCredCanc(i, 0)
'        Next i
'        nMontoCredCanc = CDbl(Format(nMontoCredCanc, "#0.00"))
'        LblMonCancel.Caption = Format(nMontoCredCanc, "#0.00")
'        LblTotCred.Caption = Format(nMontoCredCanc, "#0.00")
'    Else
'        If Trim(FECreditosVig.TextMatrix(pnRow, pnCol)) = "." Then
'            FECreditosVig.TextMatrix(pnRow, 1) = ""
'        Else
'            FECreditosVig.TextMatrix(pnRow, 1) = "1"
'        End If
'    End If
'    bOnCellCheck = False
'End Sub

Private Sub cmdCancelar_Click()
    HabilitaRefinanciacion False
    Call LimpiaPantalla
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'________________________________________________________________________________________________________________
'******************************************** METODOS / FUNCIONES - VARIOS **************************************************
Public Function Inicio(ByVal pnTipoCredPers As Variant, ByVal psCaption As String, Optional ByVal pMatProd As Variant = Nothing, _
                        Optional ByVal pbRefin As Boolean = True, Optional ByVal pbMuestraTodos As Boolean = False, _
                        Optional ByVal psAgencia As String = "ALL", Optional pbLeasing As Boolean = False, _
                        Optional ByVal pbMicroMulti As Boolean = False, Optional pbInfoGas As Boolean = False) As String
    Dim i As Integer
    Dim vsSelecPers As String
    Dim vAgenciaSelect As String

    Me.Caption = psCaption
    vsSelecPers = ""
    vAgenciaSelect = IIf(psAgencia = "ALL", "", psAgencia)
    'Call CargaGridAgencia(pnTipoCredPers, pMatProd, pbRefin, pbMuestraTodos, psAgencia, pbLeasing, pbMicroMulti, pbInfoGas)
    Screen.MousePointer = 0
    Me.Show 1
    Inicio = vsSelecPers
End Function

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
    Dim oCredito As COMNCredito.NCOMCredito
    Dim oDCredito As COMDCredito.DCOMCredito
    Dim i As Integer
    Dim nTotalDesembolsos As Double
    'Variables para Carga de Datos
    Dim rsRefina As ADODB.Recordset
    Dim rsTotalRefina As ADODB.Recordset
    Dim rsCredVig As ADODB.Recordset
    Dim vTotalesVig As Variant
    Dim rsVigGra As ADODB.Recordset
    Dim rsCuentas As ADODB.Recordset
    Dim vAgencias As Variant
    Dim rsGastos As ADODB.Recordset
    Dim nMontoITF As Double
    Dim rsDesemPar As ADODB.Recordset
    Dim bAmpliacion As Boolean
    Dim rBancos As ADODB.Recordset
    Dim nTotalPrestamo As Double
    Dim sPersCodTitular As String
    Dim nDestinoCredito As Integer
    Dim oDCreditos As COMDCredito.DCOMCreditos
    Set oDCredito = New COMDCredito.DCOMCredito
 
    On Error GoTo ErrorCargaDatos
    Set oCredito = New COMNCredito.NCOMCredito
    Dim oPers  As COMDPersona.UCOMPersona
    sPersCodTitular = oDCredito.RecuperaTitularCredito(psCtaCod)
    Set oPers = New COMDPersona.UCOMPersona
        If oPers.fgVerificaEmpleado(sPersCodTitular) Or oPers.fgVerificaEmpleadoVincualdo(sPersCodTitular) Then
        If Not oCredito.ExisteAsignaSaldo(psCtaCod, 2) Then
            MsgBox "El crédito aún no tiene saldo asignado, verificar con el Departamento de Administración de Créditos.", vbInformation, "Aviso"
            CargaDatos = False
            Exit Function
        End If
        End If
    Set oPers = Nothing
        
        If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
            If gCredDesembLeasing <> sOperacion Then
                    MsgBox "Este Crédito es un Producto Leasing, favor elegir otra modalidad de Refinanciacion"
                    Exit Function
            End If
        Else
            If gCredDesembLeasing = sOperacion Then
                    MsgBox "Esta modalidad de desembolso solo es para Productos Leasing, favor elegir otro crédito"
                    Exit Function
            End If
        End If

    Call oCredito.CargarDatosRefinanciar(psCtaCod, _
                                        gdFecSis, _
                                        nNroProxDesemb, _
                                        rsRefina, _
                                        rsTotalRefina, _
                                        rsCredVig, _
                                        vTotalesVig, _
                                        rsVigGra, _
                                        rsCuentas, _
                                        vAgencias, _
                                        rsGastos, _
                                        nMontoITF, _
                                        rsDesemPar, _
                                        bAmpliacion, _
                                        sOperacion, _
                                        pbOperacionEfectivo, _
                                        pnMontoLavDinero, _
                                        pnTC, _
                                        pbExoneradaLavado, _
                                        psPersCodRep, _
                                        psPersNombreRep, _
                                        rBancos, _
                                        rsRelEmp)
 

    Set oCredito = Nothing
    If Not rsRefina.BOF And Not rsRefina.EOF Then
    
    'Rerencia de Capital Interes
     If IsNull(rsRefina!bRefCapInt) Then
            bCapitalInt = False
     Else
            bCapitalInt = IIf(rsRefina!bRefCapInt, True, False)
     End If

        sTpoProdCod = rsRefina!cTpoProdCod
        bRevisaDesemb = False
        If sTpoProdCod <> gColConsPFTpoProducto Or _
        (sTpoProdCod = gColConsPFTpoProducto And (rsRefina!nPersPersoneria <> gPersonaNat Or Left(rsRefina!cTpoCredCod, 1) = "1" Or Left(rsRefina!cTpoCredCod, 1) = "2" Or Left(rsRefina!cTpoCredCod, 1) = "3")) Then
            Set oDCreditos = New COMDCredito.DCOMCreditos
            'Valida si el credito paso por el PreDesembolso / Control Admin. Credito.
            If Not oDCreditos.VerificaRevisaDesembControlAdmCred(ActxCta.NroCuenta) Then
                MsgBox "El crédito no fue revisado o presenta observaciones para su Vigencia por parte de Administración de Créditos. Comunicar al analista respectivo para tomar las medidas del caso.", vbInformation, "Aviso"
                CargaDatos = False
                bRevisaDesemb = True
                Exit Function
            End If
            Set oDCreditos = Nothing
        End If
        
        '***** Carga Datos Tab0: "CLIENTE" *****
        CargaDatos = True
        nNroCalen = rsRefina!nNroCalen
        LblCodCli.Caption = rsRefina!cperscod
        sPersCod = rsRefina!cperscod
        LblNomCli.Caption = PstaNombre(rsRefina!cPersNombre)
        lblDocNat.Caption = IIf(IsNull(rsRefina!DNI), "", rsRefina!DNI)
        lblDocJur.Caption = IIf(IsNull(rsRefina!Ruc), "", rsRefina!Ruc)
        LblCliDirec.Caption = IIf(IsNull(rsRefina!cPersDireccDomicilio), "", rsRefina!cPersDireccDomicilio)
        '***** Carga Datos Tab1: "CREDITO" *****
        lblTipoCred.Caption = rsRefina!cTpoCredDes
        lblTipoProd.Caption = rsRefina!cTpoProdDes
        lblmoneda.Caption = IIf(IsNull(rsRefina!cMoneda), "", rsRefina!cMoneda)
        LblLineaCred.Caption = IIf(IsNull(rsRefina!cDescripcion), "", Trim(rsRefina!cDescripcion))
        lblNroCuotas.Caption = rsRefina!NroCuota
        lblTasa.Caption = rsRefina!nTasaInteres
        lblFechaApro.Caption = rsRefina!dFechaAprobacion
        lblanalista.Caption = rsRefina!cAnalista
        lblApoderado.Caption = rsRefina!cApoderado
        '***** Carga Datos Tab2: "VIGENCIA" *****
        nInteres = 0
        nMontoGastos = 0
        nTotal = 0
        nMontoCapital = 0
          'Carga Datos "feVigenciaRefina"
            LimpiaFlex FeVigenciaRefina
            If Not rsTotalRefina.BOF Or rsTotalRefina.EOF Then
              Do While Not rsTotalRefina.EOF
                    FeVigenciaRefina.AdicionaFila
                    FeVigenciaRefina.TextMatrix(rsTotalRefina.Bookmark, 1) = rsTotalRefina!cCtaCodRef
                    FeVigenciaRefina.TextMatrix(rsTotalRefina.Bookmark, 2) = Format(rsTotalRefina!nCapital, "#0.00")
                    FeVigenciaRefina.TextMatrix(rsTotalRefina.Bookmark, 3) = Format(rsTotalRefina!nInteres, "#0.00")
                    FeVigenciaRefina.TextMatrix(rsTotalRefina.Bookmark, 4) = Format(rsTotalRefina!nGastos, "#0.00")
                    FeVigenciaRefina.TextMatrix(rsTotalRefina.Bookmark, 5) = Format(rsTotalRefina!nMontoRef, "#,##0.00")
                    nMontoCapital = nMontoCapital + rsTotalRefina!nCapital
                    nInteres = nInteres + rsTotalRefina!nInteres
                    nMontoGastos = nMontoGastos + rsTotalRefina!nGastos
                    nTotal = nTotal + rsTotalRefina!nMontoRef
                    rsTotalRefina.MoveNext
                Loop
            End If
        lblCapital.Caption = Format(IIf(IsNull(nMontoCapital), 0, nMontoCapital), "#0.00")
        lblIntereses.Caption = Format(IIf(IsNull(nInteres), 0, nInteres), "#0.00")
        LblMonGastos.Caption = Format(IIf(IsNull(nMontoGastos), 0, nMontoGastos), "#0.00")
        nTotalARefinanciar = CDbl(lblCapital.Caption) + CDbl(lblIntereses.Caption) + CDbl(LblMonGastos)
        LblMontoRefinaTotal.Caption = Format(IIf(IsNull(nTotalARefinanciar), 0, nTotalARefinanciar), "#,##0.00")
        lblTotal.Caption = Format(IIf(nTotal = 0, 0, nTotal), "#,##0.00")
        nPlazoRefina = rsRefina!nPlazo
         'Creditos Datos "Vigentes"
         LimpiaFlex FECreditosVig
         If rsCredVig.RecordCount > 0 Then rsCredVig.MoveFirst
            Do While Not rsCredVig.EOF
               FECreditosVig.AdicionaFila , , True
               'FECreditosVig.TextMatrix(r.Bookmark, 0) = r.Bookmark
               FECreditosVig.TextMatrix(rsCredVig.Bookmark, 2) = rsCredVig!cCtaCod
               FECreditosVig.TextMatrix(rsCredVig.Bookmark, 3) = rsCredVig!nDiasAtraso
               FECreditosVig.TextMatrix(rsCredVig.Bookmark, 4) = vTotalesVig(rsCredVig.Bookmark)
               'FECreditosVig.TextMatrix(rsCredVig.Bookmark, 1) = IIf(rsCredVig!nCheck = 0, "", 1)
               FECreditosVig.TextMatrix(rsCredVig.Bookmark, 5) = rsCredVig!PagAmp
               rsCredVig.MoveNext
             Loop
            
            Do While Not rsVigGra.EOF
            For i = 1 To FECreditosVig.Rows - 1
                If Trim(FECreditosVig.TextMatrix(i, 2)) = Trim(rsVigGra!cCtaCodRef) Then
                    FECreditosVig.TextMatrix(i, 1) = "1"
                    bOnCellCheck = True
                    'Call FECreditosVig_OnCellCheck(i, 1)
                End If
            Next i
            rsVigGra.MoveNext
        Loop
        
        '********************Al dar check en creditos vigentes ' En observacion*****************
        Dim oDInstFinan As COMDPersona.DCOMInstFinac
        Set oDInstFinan = New COMDPersona.DCOMInstFinac
        bInstFinanc = oDInstFinan.VerificaEsInstFinanc(rsRefina!cperscod)
        Set oDInstFinan = Nothing
        '*******************************************
        
         'Carga todos los desembolsos Pp
        nTotalDesembolsos = 0
        If bAmpliacion = True Then
            MsgBox "Este credito es un credito de ampliacion " & vbCrLf & _
                   "Por favor verifique que el credito a cancelar este seleccionado", vbInformation, "AVISO"
        End If
    Else
        CargaDatos = False
    End If
    
      Set rsRefina = Nothing
    
    Exit Function
ErrorCargaDatos:
    MsgBox err.Description, vbCritical, "Aviso"
End Function

Private Sub LimpiaPantalla()
    Dim i As Integer
    ActxCta.NroCuenta = ""
    ContMatCredCanc = 0
    nMontoGastos = 0
    For i = 0 To 99
        MatCredCanc(i, 0) = ""
        MatCredCanc(i, 1) = ""
    Next i
    sCtaAho = ""
    nNroProxDesemb = 0
    psCodIF = ""
    Set pRSRela = Nothing
    pnTipoCuenta = 0
    psNroDoc = ""
    pbDocumento = False
    pnPersoneria = 0
    LimpiaControles Me, True
    LimpiaFlex FECreditosVig
    LimpiaFlex FeVigenciaRefina
    
    LblMonGastos.Caption = "0.00"
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    bInstFinanc = False
    lblCapital.Caption = "0.00"
    lblIntereses.Caption = "0.00"
    LblMonGastos.Caption = "0.00"
    LblMontoRefinaTotal.Caption = "0.00"
    lblTotal.Caption = "0.00"
End Sub

Private Sub HabilitaRefinanciacion(ByVal pbHabilita As Boolean)
    fraCliente.Enabled = pbHabilita
    fraCredito.Enabled = pbHabilita
    CmdVigencia.Enabled = pbHabilita
    CmdSalir.Enabled = pbHabilita
    cmdcancelar.Enabled = pbHabilita
    If sTpoProdCod = "517" Then
        CmdSalir.Enabled = pbHabilita
        CmdVigencia.Enabled = pbHabilita
        cmdcancelar.Enabled = pbHabilita
    End If
End Sub

Public Sub RefinanciarCredito(ByVal sCodOpe As String)
    bLeasing = False
    sOperacion = sCodOpe
    SSTabDatosRefina.TabVisible(0) = True
    SSTabDatosRefina.TabVisible(1) = True
    SSTabDatosRefina.TabVisible(2) = True
    Me.Caption = "Refinanciacion de Creditos"
    bRefinanc = True
    Me.Show 1
End Sub

