VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredDesembAbonoCta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desembolso con Abono a Cuenta"
   ClientHeight    =   7260
   ClientLeft      =   1875
   ClientTop       =   2670
   ClientWidth     =   8355
   Icon            =   "frmCredDesembAbonoCta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   6390
      Left            =   105
      TabIndex        =   3
      Top             =   825
      Width           =   8160
      Begin TabDlg.SSTab SSTabDatos 
         Height          =   6045
         Left            =   120
         TabIndex        =   4
         Top             =   210
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   10663
         _Version        =   393216
         Tabs            =   8
         Tab             =   2
         TabsPerRow      =   5
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Cliente"
         TabPicture(0)   =   "frmCredDesembAbonoCta.frx":030A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "FraCliente"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Credito"
         TabPicture(1)   =   "frmCredDesembAbonoCta.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FraCredito"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Ctas de Ahorros"
         TabPicture(2)   =   "frmCredDesembAbonoCta.frx":0342
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame5"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Desembolsos"
         TabPicture(3)   =   "frmCredDesembAbonoCta.frx":035E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "FraDatosDesemb"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Fechas Desemb."
         TabPicture(4)   =   "frmCredDesembAbonoCta.frx":037A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame1"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Cheque"
         TabPicture(5)   =   "frmCredDesembAbonoCta.frx":0396
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame3"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Desembolsos"
         TabPicture(6)   =   "frmCredDesembAbonoCta.frx":03B2
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Frame4"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "Leasing"
         TabPicture(7)   =   "frmCredDesembAbonoCta.frx":03CE
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "Frame10"
         Tab(7).Control(1)=   "Frame11"
         Tab(7).ControlCount=   2
         Begin VB.Frame Frame11 
            Caption         =   "Proveedores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   -74880
            TabIndex        =   117
            Top             =   720
            Width           =   7455
            Begin SICMACT.FlexEdit FEProveedores 
               Height          =   1455
               Left            =   480
               TabIndex        =   118
               Top             =   360
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   2566
               Cols0           =   5
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-cPersCod-Nombre-Cuenta-nMontoLeasing"
               EncabezadosAnchos=   "0-0-5000-0-1200"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X-X-X"
               ListaControles  =   "0-0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-L-L-R"
               FormatosEdit    =   "0-0-0-0-4"
               TextArray0      =   "#"
               lbEditarFlex    =   -1  'True
               lbUltimaInstancia=   -1  'True
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Pagos iniciales"
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
            TabIndex        =   115
            Top             =   2880
            Width           =   7455
            Begin SICMACT.FlexEdit FEPagosIniciales 
               Height          =   1455
               Left            =   240
               TabIndex        =   116
               Top             =   360
               Width           =   6975
               _ExtentX        =   12303
               _ExtentY        =   2566
               Cols0           =   4
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-NroConcepto-Concepto-Monto"
               EncabezadosAnchos=   "0-0-4500-2000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X-X"
               ListaControles  =   "0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-C-L-R"
               FormatosEdit    =   "0-0-0-4"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
         End
         Begin VB.Frame Frame4 
            Height          =   4980
            Left            =   -74880
            TabIndex        =   83
            Top             =   720
            Width           =   7470
            Begin VB.Frame Frame8 
               Caption         =   "Cta. Abono Recaudo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1080
               Left            =   135
               TabIndex        =   102
               Top             =   135
               Width           =   7095
               Begin VB.Label Label50 
                  AutoSize        =   -1  'True
                  Caption         =   "Cuenta :"
                  Height          =   195
                  Left            =   210
                  TabIndex        =   108
                  Top             =   270
                  Width           =   600
               End
               Begin VB.Label lblCtaAboRecaudo 
                  BackColor       =   &H00C0FFFF&
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
                  Left            =   945
                  TabIndex        =   107
                  Top             =   255
                  Width           =   2055
               End
               Begin VB.Label Label48 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo :"
                  Height          =   195
                  Left            =   3105
                  TabIndex        =   106
                  Top             =   285
                  Width           =   405
               End
               Begin VB.Label lblTipoCtaRecaudo 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000E&
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
                  Height          =   270
                  Left            =   3825
                  TabIndex        =   105
                  Top             =   240
                  Width           =   2055
               End
               Begin VB.Label Label46 
                  AutoSize        =   -1  'True
                  Caption         =   "Agencia :"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   104
                  Top             =   720
                  Width           =   675
               End
               Begin VB.Label lblAgenciaCtaRecaudo 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   270
                  Left            =   945
                  TabIndex        =   103
                  Top             =   675
                  Width           =   2055
               End
            End
            Begin VB.Frame Frame6 
               Height          =   3675
               Left            =   105
               TabIndex        =   84
               Top             =   1200
               Width           =   7125
               Begin VB.Frame Frame7 
                  BorderStyle     =   0  'None
                  Height          =   3480
                  Left            =   90
                  TabIndex        =   85
                  Top             =   165
                  Width           =   6930
                  Begin VB.CommandButton cmdSale 
                     Cancel          =   -1  'True
                     Caption         =   "&Salir"
                     Height          =   495
                     Left            =   3840
                     TabIndex        =   89
                     Top             =   1800
                     Width           =   1455
                  End
                  Begin VB.CommandButton cmdCancela 
                     Caption         =   "&Cancelar"
                     Enabled         =   0   'False
                     Height          =   495
                     Left            =   3840
                     TabIndex        =   88
                     Top             =   660
                     Width           =   1455
                  End
                  Begin VB.CommandButton cmdDesembolso 
                     Caption         =   "&Desembolsar"
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
                     Left            =   3840
                     TabIndex        =   87
                     Top             =   90
                     Width           =   1455
                  End
                  Begin VB.CommandButton cmdVisualiza 
                     Caption         =   "&Visualización"
                     Enabled         =   0   'False
                     Height          =   495
                     Left            =   3840
                     TabIndex        =   86
                     Top             =   1230
                     Visible         =   0   'False
                     Width           =   1455
                  End
                  Begin VB.Label lblampliado 
                     Caption         =   "0"
                     Height          =   255
                     Left            =   3840
                     TabIndex        =   131
                     Top             =   2400
                     Visible         =   0   'False
                     Width           =   375
                  End
                  Begin VB.Label Label32 
                     AutoSize        =   -1  'True
                     Caption         =   "Tasación           :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   285
                     TabIndex        =   114
                     Top             =   1605
                     Width           =   1515
                  End
                  Begin VB.Label lblMonTasacion 
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
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Left            =   2010
                     TabIndex        =   113
                     Top             =   1560
                     Width           =   1215
                  End
                  Begin VB.Label lblMonComision 
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
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Left            =   2010
                     TabIndex        =   112
                     Top             =   1920
                     Width           =   1215
                  End
                  Begin VB.Label label49 
                     AutoSize        =   -1  'True
                     Caption         =   "Com.Estruc.Caja :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   285
                     TabIndex        =   111
                     Top             =   1965
                     Width           =   1530
                  End
                  Begin VB.Label lblMonSeguro 
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
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Left            =   2010
                     TabIndex        =   110
                     Top             =   1200
                     Width           =   1215
                  End
                  Begin VB.Label Label44 
                     AutoSize        =   -1  'True
                     Caption         =   "Seguro (SOAT)   :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   285
                     TabIndex        =   109
                     Top             =   1245
                     Width           =   1530
                  End
                  Begin VB.Label lblTotalFinanciar 
                     Alignment       =   1  'Right Justify
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
                     Left            =   2010
                     TabIndex        =   101
                     Top             =   3075
                     Width           =   1215
                  End
                  Begin VB.Label lblMonFinanc 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00808000&
                     BackStyle       =   0  'Transparent
                     Caption         =   "A Financiar :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   100
                     Top             =   3105
                     Width           =   1095
                  End
                  Begin VB.Line Line2 
                     BorderColor     =   &H00C00000&
                     BorderWidth     =   2
                     X1              =   1770
                     X2              =   3510
                     Y1              =   3000
                     Y2              =   3000
                  End
                  Begin VB.Label lblMontoPrestamo 
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
                     Left            =   2010
                     TabIndex        =   99
                     Top             =   2625
                     Width           =   1215
                  End
                  Begin VB.Label Label40 
                     AutoSize        =   -1  'True
                     Caption         =   "Prestamo           :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   285
                     TabIndex        =   98
                     Top             =   2670
                     Width           =   1515
                  End
                  Begin VB.Label lblMonOperador 
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
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Left            =   2010
                     TabIndex        =   97
                     Top             =   480
                     Width           =   1215
                  End
                  Begin VB.Label Label38 
                     AutoSize        =   -1  'True
                     Caption         =   "Operador           :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   285
                     TabIndex        =   96
                     Top             =   525
                     Width           =   1515
                  End
                  Begin VB.Label lblMonConcesionario 
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
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Left            =   2010
                     TabIndex        =   95
                     Top             =   135
                     Width           =   1215
                  End
                  Begin VB.Label Label34 
                     AutoSize        =   -1  'True
                     Caption         =   "Concesionario    :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   285
                     TabIndex        =   94
                     Top             =   180
                     Width           =   1515
                  End
                  Begin VB.Label Label33 
                     AutoSize        =   -1  'True
                     Caption         =   "I.T.F.                :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   285
                     TabIndex        =   93
                     Top             =   2310
                     Width           =   1500
                  End
                  Begin VB.Label lblMonITF 
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
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Left            =   2010
                     TabIndex        =   92
                     Top             =   2265
                     Width           =   1215
                  End
                  Begin VB.Label lblMonNotario 
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
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Left            =   2010
                     TabIndex        =   91
                     Top             =   840
                     Width           =   1215
                  End
                  Begin VB.Label Label27 
                     AutoSize        =   -1  'True
                     Caption         =   "Notario              :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   285
                     TabIndex        =   90
                     Top             =   885
                     Width           =   1530
                  End
               End
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Datos de Cheque"
            Height          =   1515
            Left            =   -74865
            TabIndex        =   70
            Top             =   960
            Width           =   7245
            Begin VB.ComboBox CboCta 
               Height          =   315
               Left            =   1770
               Style           =   2  'Dropdown List
               TabIndex        =   75
               Top             =   630
               Width           =   3345
            End
            Begin VB.TextBox TxtCheque 
               Height          =   330
               Left            =   1770
               TabIndex        =   74
               Top             =   1005
               Width           =   2190
            End
            Begin VB.ComboBox CboBancos 
               Height          =   315
               Left            =   1770
               Style           =   2  'Dropdown List
               TabIndex        =   72
               Top             =   300
               Width           =   3345
            End
            Begin VB.Label Label26 
               Caption         =   "Cuenta                :"
               Height          =   315
               Left            =   195
               TabIndex        =   76
               Top             =   660
               Width           =   1440
            End
            Begin VB.Label Label25 
               Caption         =   "Nro. de Cheque   :"
               Height          =   315
               Left            =   180
               TabIndex        =   73
               Top             =   1050
               Width           =   1380
            End
            Begin VB.Label Label24 
               Caption         =   "Banco                 :"
               Height          =   315
               Left            =   210
               TabIndex        =   71
               Top             =   345
               Width           =   1440
            End
         End
         Begin VB.Frame Frame1 
            Height          =   3570
            Left            =   -74850
            TabIndex        =   65
            Top             =   750
            Width           =   7380
            Begin SICMACT.FlexEdit FEDesembolsos 
               Height          =   1935
               Left            =   150
               TabIndex        =   66
               Top             =   270
               Width           =   6990
               _ExtentX        =   12330
               _ExtentY        =   3413
               Cols0           =   5
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "-Nro.-Fecha-Monto-Proximo"
               EncabezadosAnchos=   "350-500-2000-2000-1000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
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
               ColumnasAEditar =   "X-X-X-X-X"
               TextStyleFixed  =   4
               ListaControles  =   "0-0-0-0-0"
               BackColor       =   16777215
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-C-C-C-C"
               FormatosEdit    =   "0-0-0-0-0"
               lbUltimaInstancia=   -1  'True
               ColWidth0       =   345
               RowHeight0      =   300
               ForeColorFixed  =   -2147483635
               CellBackColor   =   16777215
            End
            Begin VB.Label lblTotal 
               Alignment       =   1  'Right Justify
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
               Left            =   3105
               TabIndex        =   68
               Top             =   2325
               Width           =   1215
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
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
               Left            =   2325
               TabIndex        =   67
               Top             =   2340
               Width           =   630
            End
         End
         Begin VB.Frame FraDatosDesemb 
            Height          =   5220
            Left            =   -74835
            TabIndex        =   26
            Top             =   720
            Width           =   7470
            Begin VB.Frame frGastoMYPE 
               Caption         =   "Multiriesgo MYPE"
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
               Height          =   615
               Left            =   5680
               TabIndex        =   129
               Top             =   130
               Width           =   1695
               Begin VB.TextBox txtMultMype 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   130
                  Top             =   240
                  Width           =   1455
               End
            End
            Begin VB.Frame FraDesemb 
               Height          =   3555
               Left            =   120
               TabIndex        =   40
               Top             =   1575
               Width           =   7245
               Begin VB.Frame Frame9 
                  BorderStyle     =   0  'None
                  Height          =   3360
                  Left            =   90
                  TabIndex        =   41
                  Top             =   165
                  Width           =   6690
                  Begin VB.TextBox txtGlosaBloqueo 
                     Height          =   525
                     Left            =   4680
                     MultiLine       =   -1  'True
                     TabIndex        =   128
                     Top             =   2760
                     Visible         =   0   'False
                     Width           =   1935
                  End
                  Begin VB.TextBox txtMontoRetirar 
                     Alignment       =   1  'Right Justify
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
                     Height          =   285
                     Left            =   4680
                     TabIndex        =   127
                     Text            =   "0.00"
                     Top             =   2400
                     Visible         =   0   'False
                     Width           =   1215
                  End
                  Begin VB.CommandButton CmdVisualizar 
                     Caption         =   "&Visualización"
                     Enabled         =   0   'False
                     Height          =   495
                     Left            =   3480
                     TabIndex        =   69
                     Top             =   1230
                     Visible         =   0   'False
                     Width           =   1455
                  End
                  Begin VB.CommandButton CmdDesemb 
                     Caption         =   "&Desembolsar"
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
                     Left            =   3480
                     TabIndex        =   44
                     Top             =   60
                     Width           =   1455
                  End
                  Begin VB.CommandButton CmdCancelar 
                     Caption         =   "&Cancelar"
                     Enabled         =   0   'False
                     Height          =   495
                     Left            =   3480
                     TabIndex        =   43
                     Top             =   660
                     Width           =   1455
                  End
                  Begin VB.CommandButton CmdSalir 
                     Caption         =   "&Salir"
                     Height          =   495
                     Left            =   3480
                     TabIndex        =   42
                     Top             =   1770
                     Width           =   1455
                  End
                  Begin VB.Label lblMonGastoCierre 
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
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Left            =   1890
                     TabIndex        =   133
                     Top             =   1440
                     Width           =   1215
                  End
                  Begin VB.Label Label20 
                     AutoSize        =   -1  'True
                     Caption         =   "Gastos de Cierre :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   132
                     Top             =   1485
                     Width           =   1545
                  End
                  Begin VB.Label lblGlosaBloqueo 
                     Caption         =   "Glosa Bloqueo:"
                     Height          =   255
                     Left            =   3480
                     TabIndex        =   126
                     Top             =   2880
                     Visible         =   0   'False
                     Width           =   1215
                  End
                  Begin VB.Label lblMontoRetirar 
                     Caption         =   "Monto a Retirar:"
                     Height          =   255
                     Left            =   3480
                     TabIndex        =   125
                     Top             =   2400
                     Visible         =   0   'False
                     Width           =   1335
                  End
                  Begin VB.Label lblTotalIniciales 
                     Alignment       =   1  'Right Justify
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
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Left            =   1875
                     TabIndex        =   120
                     Top             =   2160
                     Width           =   1215
                  End
                  Begin VB.Label Label30 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00808000&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Pago Inicial        :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   119
                     Top             =   2160
                     Width           =   1560
                  End
                  Begin VB.Label Label28 
                     AutoSize        =   -1  'True
                     Caption         =   "Poliza                :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   78
                     Top             =   1125
                     Width           =   1545
                  End
                  Begin VB.Label lblPoliza 
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
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Left            =   1890
                     TabIndex        =   77
                     Top             =   1080
                     Width           =   1215
                  End
                  Begin VB.Label LblItf 
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
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Left            =   1890
                     TabIndex        =   64
                     Top             =   1785
                     Width           =   1215
                  End
                  Begin VB.Label Label22 
                     AutoSize        =   -1  'True
                     Caption         =   "I.T.F.                 :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   63
                     Top             =   1830
                     Width           =   1560
                  End
                  Begin VB.Label Label13 
                     AutoSize        =   -1  'True
                     Caption         =   "Gastos               :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   54
                     Top             =   420
                     Width           =   1560
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
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Left            =   1890
                     TabIndex        =   53
                     Top             =   375
                     Width           =   1215
                  End
                  Begin VB.Label Label15 
                     AutoSize        =   -1  'True
                     Caption         =   "Desembolsos      :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   52
                     Top             =   60
                     Width           =   1545
                  End
                  Begin VB.Label LblNroDesemb 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H8000000E&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "0/0"
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
                     Left            =   2490
                     TabIndex        =   51
                     Top             =   15
                     Width           =   585
                  End
                  Begin VB.Label Label16 
                     AutoSize        =   -1  'True
                     Caption         =   "Cancelaciones    :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   50
                     Top             =   765
                     Width           =   1560
                  End
                  Begin VB.Label LblMonCancel 
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
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Left            =   1890
                     TabIndex        =   49
                     Top             =   720
                     Width           =   1215
                  End
                  Begin VB.Label Label14 
                     AutoSize        =   -1  'True
                     Caption         =   "Prestamo            :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   48
                     Top             =   2550
                     Width           =   1575
                  End
                  Begin VB.Label LblMonPrestamo 
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
                     Left            =   1890
                     TabIndex        =   47
                     Top             =   2505
                     Width           =   1215
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H00C00000&
                     BorderWidth     =   2
                     X1              =   1650
                     X2              =   3390
                     Y1              =   2880
                     Y2              =   2880
                  End
                  Begin VB.Label Label17 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00808000&
                     BackStyle       =   0  'Transparent
                     Caption         =   "A Desembolsar    :"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   46
                     Top             =   2985
                     Width           =   1575
                  End
                  Begin VB.Label lblMonDesemb 
                     Alignment       =   1  'Right Justify
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
                     Left            =   1890
                     TabIndex        =   45
                     Top             =   2955
                     Width           =   1215
                  End
               End
            End
            Begin VB.Frame FraAbono 
               Caption         =   "Abono"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1395
               Left            =   2640
               TabIndex        =   30
               Top             =   135
               Width           =   3015
               Begin VB.Label LblAgeAbono 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   390
                  Left            =   825
                  TabIndex        =   38
                  Top             =   915
                  Width           =   2055
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Agencia :"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   37
                  Top             =   960
                  Width           =   675
               End
               Begin VB.Label LblTipoCta 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000E&
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
                  Height          =   270
                  Left            =   825
                  TabIndex        =   34
                  Top             =   600
                  Width           =   2055
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo :"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   33
                  Top             =   645
                  Width           =   405
               End
               Begin VB.Label LblCtaAbo 
                  BackColor       =   &H00C0FFFF&
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
                  Left            =   825
                  TabIndex        =   32
                  Top             =   255
                  Width           =   2055
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "Cuenta :"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   31
                  Top             =   270
                  Width           =   600
               End
            End
            Begin VB.Frame FraCredVig 
               Caption         =   "Creditos a Cancelar"
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
               Height          =   1395
               Left            =   120
               TabIndex        =   28
               Top             =   135
               Width           =   2535
               Begin VB.ListBox LstCredVig 
                  BackColor       =   &H00FFFFFF&
                  Height          =   450
                  Left            =   120
                  TabIndex        =   29
                  Top             =   240
                  Width           =   2325
               End
               Begin VB.Label LblTotCred 
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
                  Left            =   1200
                  TabIndex        =   36
                  Top             =   975
                  Width           =   1215
               End
               Begin VB.Label Label11 
                  Caption         =   "Total :"
                  Height          =   240
                  Left            =   195
                  TabIndex        =   35
                  Top             =   1020
                  Width           =   450
               End
            End
         End
         Begin VB.Frame Frame5 
            Height          =   4755
            Left            =   135
            TabIndex        =   20
            Top             =   810
            Width           =   7560
            Begin VB.CheckBox chkAperCtaAhorro 
               Caption         =   "APERTURAR NUEVA CUENTA"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   3360
               TabIndex        =   124
               Top             =   3360
               Width           =   2655
            End
            Begin VB.Frame fraAperturaCtaAhorro 
               Caption         =   "   "
               Enabled         =   0   'False
               Height          =   1095
               Left            =   3240
               TabIndex        =   121
               Top             =   3360
               Width           =   4215
               Begin VB.ComboBox cboPrograma 
                  Height          =   315
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   122
                  Top             =   480
                  Width           =   2775
               End
               Begin VB.Label Label35 
                  Caption         =   "Sub Producto:"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   123
                  Top             =   480
                  Width           =   1095
               End
            End
            Begin VB.CommandButton CmdSeleccionar 
               Caption         =   "&Seleccionar"
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
               Height          =   420
               Left            =   210
               TabIndex        =   25
               Top             =   3300
               Width           =   1440
            End
            Begin VB.CommandButton CmdAperturar 
               Caption         =   "&Aperturar"
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
               Height          =   420
               Left            =   135
               TabIndex        =   24
               Top             =   4260
               Visible         =   0   'False
               Width           =   1440
            End
            Begin TabDlg.SSTab SSTCtasAho 
               Height          =   2655
               Left            =   120
               TabIndex        =   21
               Top             =   495
               Width           =   7320
               _ExtentX        =   12912
               _ExtentY        =   4683
               _Version        =   393216
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   520
               TabCaption(0)   =   "Cuentas"
               TabPicture(0)   =   "frmCredDesembAbonoCta.frx":03EA
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "FECtaAhoDesemb"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Clientes"
               TabPicture(1)   =   "frmCredDesembAbonoCta.frx":0406
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "FEPersCtaAho"
               Tab(1).ControlCount=   1
               Begin SICMACT.FlexEdit FEPersCtaAho 
                  Height          =   2055
                  Left            =   -74880
                  TabIndex        =   59
                  Top             =   480
                  Width           =   7080
                  _ExtentX        =   12488
                  _ExtentY        =   3625
                  Cols0           =   3
                  HighLight       =   1
                  AllowUserResizing=   3
                  RowSizingMode   =   1
                  EncabezadosNombres=   "-Cliente-Relacion"
                  EncabezadosAnchos=   "350-5000-1200"
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
                  ColumnasAEditar =   "X-X-X"
                  TextStyleFixed  =   4
                  ListaControles  =   "0-0-0"
                  BackColor       =   13816486
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  EncabezadosAlineacion=   "C-L-C"
                  FormatosEdit    =   "0-0-0"
                  lbUltimaInstancia=   -1  'True
                  lbBuscaDuplicadoText=   -1  'True
                  ColWidth0       =   345
                  RowHeight0      =   300
                  ForeColorFixed  =   -2147483635
                  CellBackColor   =   13816486
               End
               Begin SICMACT.FlexEdit FECtaAhoDesemb 
                  Height          =   1935
                  Left            =   120
                  TabIndex        =   60
                  Top             =   480
                  Width           =   7095
                  _ExtentX        =   12330
                  _ExtentY        =   3413
                  Cols0           =   5
                  HighLight       =   1
                  AllowUserResizing=   3
                  RowSizingMode   =   1
                  EncabezadosNombres=   "-Cuenta-Agencia-ITF-TpoPrograma"
                  EncabezadosAnchos=   "350-2000-3000-1500-0"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
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
                  ColumnasAEditar =   "X-X-X-X-X"
                  TextStyleFixed  =   4
                  ListaControles  =   "0-0-0-0-0"
                  BackColor       =   16777215
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  EncabezadosAlineacion=   "C-C-C-C-C"
                  FormatosEdit    =   "0-0-0-0-0"
                  lbUltimaInstancia=   -1  'True
                  ColWidth0       =   345
                  RowHeight0      =   300
                  ForeColorFixed  =   -2147483635
                  CellBackColor   =   16777215
               End
            End
            Begin VB.CommandButton CmdDeseleccionar 
               Caption         =   "&Deseleccionar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   210
               TabIndex        =   55
               Top             =   3300
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label LblTipoAbono 
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
               Height          =   210
               Left            =   3300
               TabIndex        =   62
               Top             =   225
               Width           =   1290
            End
            Begin VB.Label Label21 
               Caption         =   "ABONO A CUENTA :"
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
               Height          =   210
               Left            =   1425
               TabIndex        =   61
               Top             =   225
               Width           =   1875
            End
         End
         Begin VB.Frame FraCredito 
            Height          =   4590
            Left            =   -74835
            TabIndex        =   15
            Top             =   750
            Width           =   7500
            Begin SICMACT.FlexEdit FECargoAutom 
               Height          =   1275
               Left            =   1500
               TabIndex        =   57
               Top             =   1320
               Width           =   5775
               _ExtentX        =   10186
               _ExtentY        =   2249
               Cols0           =   4
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "--Cuenta-Agencia"
               EncabezadosAnchos=   "350-400-2000-2500"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
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
               ColumnasAEditar =   "X-1-X-X"
               TextStyleFixed  =   4
               ListaControles  =   "0-4-0-0"
               BackColor       =   16777215
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-C-C-C"
               FormatosEdit    =   "0-0-0-0"
               SelectionMode   =   1
               lbUltimaInstancia=   -1  'True
               ColWidth0       =   345
               RowHeight0      =   300
               ForeColorFixed  =   -2147483635
            End
            Begin SICMACT.FlexEdit FEGastos 
               Height          =   1635
               Left            =   240
               TabIndex        =   58
               Top             =   2880
               Width           =   7035
               _ExtentX        =   12409
               _ExtentY        =   2884
               Cols0           =   5
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "-Nro-Gasto-Monto-CodGasto"
               EncabezadosAnchos=   "400-400-3500-1500-0"
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
               ColumnasAEditar =   "X-X-X-X-X"
               TextStyleFixed  =   4
               ListaControles  =   "0-0-0-0-0"
               BackColor       =   14811135
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-C-L-R-C"
               FormatosEdit    =   "0-0-0-2-0"
               SelectionMode   =   1
               lbUltimaInstancia=   -1  'True
               lbPuntero       =   -1  'True
               ColWidth0       =   405
               RowHeight0      =   300
               ForeColorFixed  =   -2147483635
               CellBackColor   =   14811135
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Credito :"
               Height          =   195
               Left            =   240
               TabIndex        =   82
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblTipoCred 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1335
               TabIndex        =   81
               Top             =   240
               Width           =   3345
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Producto :"
               Height          =   195
               Left            =   240
               TabIndex        =   80
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label lblTipoProd 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1335
               TabIndex        =   79
               Top             =   600
               Width           =   3345
            End
            Begin VB.Label Label18 
               Caption         =   "Gastos del Credito : "
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
               Height          =   270
               Left            =   240
               TabIndex        =   39
               Top             =   2640
               Width           =   3450
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               Caption         =   "Cargo Automatico"
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
               Height          =   570
               Left            =   165
               TabIndex        =   27
               Top             =   1320
               Width           =   1230
            End
            Begin VB.Label LblMoneda 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   5625
               TabIndex        =   19
               Top             =   960
               Width           =   1035
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Moneda :"
               Height          =   195
               Left            =   4800
               TabIndex        =   18
               Top             =   960
               Width           =   675
            End
            Begin VB.Label LblLineaCred 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1335
               TabIndex        =   17
               Top             =   960
               Width           =   3345
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Linea Credito :"
               Height          =   195
               Left            =   240
               TabIndex        =   16
               Top             =   960
               Width           =   1020
            End
         End
         Begin VB.Frame FraCliente 
            Height          =   4215
            Left            =   -74835
            TabIndex        =   5
            Top             =   645
            Width           =   7485
            Begin SICMACT.FlexEdit FECreditosVig 
               Height          =   1470
               Left            =   705
               TabIndex        =   56
               Top             =   2280
               Width           =   5865
               _ExtentX        =   10345
               _ExtentY        =   2593
               Cols0           =   6
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "--Credito-Atraso-Monto-PagAmp"
               EncabezadosAnchos=   "300-350-2000-1200-1200-0"
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
               ColWidth0       =   300
               RowHeight0      =   300
               ForeColorFixed  =   -2147483635
               CellBackColor   =   14286847
            End
            Begin VB.Label LblCliDirec 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1170
               TabIndex        =   23
               Top             =   1455
               Width           =   5445
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Direccion :"
               Height          =   195
               Left            =   225
               TabIndex        =   22
               Top             =   1485
               Width           =   765
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Creditos Vigentes en Agencia :"
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
               Left            =   210
               TabIndex        =   14
               Top             =   1875
               Width           =   2640
            End
            Begin VB.Label LblDocJur 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   5160
               TabIndex        =   13
               Top             =   1050
               Width           =   1440
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Doc. de Identidad :"
               Height          =   195
               Left            =   3465
               TabIndex        =   12
               Top             =   1095
               Width           =   1365
            End
            Begin VB.Label LblDocNat 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1170
               TabIndex        =   11
               Top             =   1050
               Width           =   1440
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "RUC:"
               Height          =   195
               Left            =   225
               TabIndex        =   10
               Top             =   1095
               Width           =   390
            End
            Begin VB.Label LblNomCli 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1170
               TabIndex        =   9
               Top             =   675
               Width           =   5445
            End
            Begin VB.Label LblCodCli 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1170
               TabIndex        =   8
               Top             =   300
               Width           =   1290
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nombre :"
               Height          =   195
               Left            =   225
               TabIndex        =   7
               Top             =   720
               Width           =   645
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Codigo :"
               Height          =   195
               Left            =   225
               TabIndex        =   6
               Top             =   375
               Width           =   585
            End
         End
      End
   End
   Begin VB.Frame FraCuenta 
      Height          =   840
      Left            =   105
      TabIndex        =   0
      Top             =   -15
      Width           =   8160
      Begin VB.CommandButton CmdExaminar 
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
         Height          =   405
         Left            =   6210
         TabIndex        =   2
         Top             =   270
         Width           =   1605
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   405
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   714
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmCredDesembAbonoCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RCtasBco As ADODB.Recordset
Private nMontoGastos As Double
Private nMontoGastoCierre As Double
Private nMontoCredCanc As Double
'Matriz que almacena Los Creditos a cancelar y sus montos
' Columna 0 : cCtaCod
' Columna 1 : Monto
Private MatCredCanc(100, 2) As String
Private ContMatCredCanc As Integer
'Matriz que almacena las Cuentas de Ahorro a Cargar el credito
' Columna 0 : cCtaCod
Private MatCargoAutom() As String
Private ContMatCargoAutom As Integer
Private sCtaAho As String
Private pnFilaSelecCtaAho As Integer
Private nNroProxDesemb As Integer
Private RPersCtaAho() As ADODB.Recordset
Private nContPersCtaAho As Integer
Private vbDesembCC As Boolean
Private vbCuentaNueva As Boolean
Private vbDesembCheque As Boolean
Private vbDesembInfoGas As Boolean 'BRGO 20111109
Private pRSRela As ADODB.Recordset
Private pnTasa As Double
Private pnPersoneria As Integer
Private pnTipoCuenta As Integer
Private pnTipoTasa As Integer
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
'CUSCO
Private psPersCodRep As String  'Codigo del Representante del Crédito
Private psPersNombreRep As String 'Nombre del Representante de Crédito

Public rsRel As ADODB.Recordset

Dim nTasaInt As DObjeto
'ARCV 12-02-2007
Dim MatTitulares As Variant
Dim nProgAhorros As Integer
Dim nMontoAbonar As Double
Dim nPlazoAbonar As Integer
Dim sPromotorAho As String
'---------
'By Capi 10042008
Dim lbPuedeAperturar As Boolean
Dim nRedondeoITF As Double 'BRGO 20111012
Dim sTpoProdCod As String  'BRGO 20111111
Dim rsRelEmp As ADODB.Recordset 'BRGO 20111111
'ALPA 20111213***********************
Dim nTotalProveedores As Integer
Dim nTotalCuotaIniciales As Integer
Dim bLeasing As Boolean
'************************************
Dim nMontoPrestamoW As Double 'WIOR 20120921
Dim nDestinoCred As Integer 'FRHU 20140225 RQ14007
Dim bInstFinanc As Boolean 'JUEZ 20140411
Dim bOnCellCheck As Boolean 'FRHU 20140424 TI-ERS015-2014
Dim bRevisaDesemb As Boolean 'JUEZ 20140730
'Dim lsCuentaCredCan As String 'VAPA 20171209 FROM 60 Comentado by NAGL 20180609
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018
Dim frsRelaFMV As ADODB.Recordset
Dim fMatTitularesFMV As Variant

Public Sub DesembolsoConCheque(ByVal sCodOpe As String)
    SSTabDatos.TabVisible(2) = False
    SSTabDatos.TabVisible(6) = False 'BRGO 201111109
    sOperacion = sCodOpe
    vbDesembCC = False
    vbCuentaNueva = False
    vbDesembCheque = True
    vbDesembInfoGas = False
    Me.Caption = "Desembolso Abono a Cuenta"
    Me.Show 1
End Sub
Public Sub DesembolsoCargoCuenta(ByVal sCodOpe As String)
    SSTabDatos.TabVisible(5) = False 'DAOR 20070213
    SSTabDatos.TabVisible(6) = False 'BRGO 201111109
    sOperacion = sCodOpe
    vbDesembCC = True
    vbCuentaNueva = False
    vbDesembCheque = False
    vbDesembInfoGas = False
    bLeasing = False
    Me.Caption = "Desembolso Abono a Cuenta"
    lblMontoRetirar.Visible = False  'FRHU 20140228 RQ 14006
    txtMontoRetirar.Visible = False  'FRHU 20140228 RQ 14006
    Me.lblGlosaBloqueo.Visible = False 'FRHU 20140228 RQ 14006
    Me.txtGlosaBloqueo.Visible = False 'FRHU 20140228 RQ 14006
    Me.Show 1
End Sub
'*** BRGO 20111109 ********************************************
'Public Sub DesembolsoInfoGas(ByVal sCodOpe As String)
'    SSTabDatos.TabVisible(5) = False
'    SSTabDatos.TabVisible(3) = False
'    sOperacion = sCodOpe
'    vbDesembCC = False
'    vbCuentaNueva = False
'    vbDesembCheque = False
'    vbDesembInfoGas = True
'    Me.Caption = "Desembolso de Producto EcoTaxi (Infogas)"
'    Me.Show 1
'End Sub
'*** END BRGO **************************************************

'By capi 06032009 Acta 022-2009
Public Function DesembolsoPigAbonoCuenta(ByVal sCodOpe As String, ByVal psPersCod As String, ByVal psCtaCod As String) As String
    
    FraCuenta.Visible = False
    SSTabDatos.TabVisible(0) = False
    SSTabDatos.TabVisible(1) = False
    'SSTabDatos.TabVisible(2) = False solo se activa cta ahorros
    SSTabDatos.TabVisible(3) = False
    SSTabDatos.TabVisible(4) = False
    SSTabDatos.TabVisible(5) = False
    SSTabDatos.TabVisible(6) = False 'BRGO 201111109
    'Set prsCuentas = oCaptacion.GetCuentasPersona(sPersCod, gCapAhorros, True, True, CInt(Mid(psCtaCod, 9, 1)))

    SSTabDatos.TabVisible(5) = False 'DAOR 20070213
    sOperacion = sCodOpe
    vbDesembCC = True
    vbCuentaNueva = False
    vbDesembCheque = False
    Me.Caption = "Desembolso Pignoraticio Abono a Cuenta"
    Me.Show 1
End Function
'ALPA 20111213
Public Sub DesembolsoCargoCuentaProveedorLeasing(ByVal sCodOpe As String)
    SSTabDatos.TabVisible(5) = False 'ALPA 20110609
    SSTabDatos.TabVisible(7) = True 'ALPA 20110609
    SSTabDatos.TabVisible(2) = False 'ALPA 20110609
    bLeasing = True
    sOperacion = sCodOpe
    vbDesembCC = False
    vbCuentaNueva = False
    vbDesembCheque = False
    Me.Caption = "Desembolso Arrendamiento Financiero"
    SSTabDatos.TabCaption(1) = "Operación"
    ActxCta.texto = "Operación"
    Label18.Caption = "Gastos de Operación"
    Label14.Caption = "Financiamiento :"
    Me.Show 1
End Sub
'*********************
Public Sub DesembolsoEfectivo(ByVal sCodOpe As String)
    bLeasing = False
    sOperacion = sCodOpe
    SSTabDatos.TabVisible(2) = False
    SSTabDatos.TabVisible(5) = False
    SSTabDatos.TabVisible(6) = False 'BRGO 201111109
    Label9.Enabled = False
    LblCtaAbo.Enabled = False
    Label10.Enabled = False
    LblTipoCta.Enabled = False
    Label12.Enabled = False
    LblAgeAbono.Enabled = False
    FraAbono.Enabled = False
    vbDesembCC = False
    vbCuentaNueva = False
    vbDesembCheque = False
    Me.Caption = "Desembolso en Efectivo"
    lblMontoRetirar.Visible = False 'FRHU20140228 RQ 14006
    txtMontoRetirar.Visible = False 'FRHU20140228 RQ 14006
    lblGlosaBloqueo.Visible = False 'FRHU20140228 RQ 14006
    txtGlosaBloqueo.Visible = False 'FRHU20140228 RQ 14006
    Me.Show 1
End Sub

Private Sub EliminarCtaAhoCargo(ByVal psCtaCod As String)
Dim i As Integer
Dim nPos As Integer

    nPos = -1
    For i = 0 To ContMatCargoAutom - 1
        If MatCargoAutom(ContMatCargoAutom - 1) = psCtaCod Then
            nPos = i
            Exit For
        End If
    Next i
    If nPos <> -1 Then
        For i = nPos To ContMatCargoAutom - 2
            MatCargoAutom(i) = MatCargoAutom(i + 1)
        Next i
        ContMatCargoAutom = ContMatCargoAutom - 1
        ReDim Preserve MatCargoAutom(ContMatCargoAutom)
    End If
        
End Sub

Private Sub AdicionaCtaAhoCargo(ByVal psCtaCod As String)
    ContMatCargoAutom = ContMatCargoAutom + 1
    ReDim Preserve MatCargoAutom(ContMatCargoAutom)
    MatCargoAutom(ContMatCargoAutom - 1) = psCtaCod
End Sub

Private Sub EliminarCreditoACancelar(ByVal psCtaCod As String)
Dim i As Integer
Dim nPos As Integer

    nPos = -1
    For i = 0 To ContMatCredCanc - 1
        If MatCredCanc(ContMatCredCanc - 1, 0) = psCtaCod Then
            nPos = i
            Exit For
        End If
    Next i
    If nPos <> -1 Then
        For i = nPos To ContMatCredCanc - 2
            MatCredCanc(i, 0) = MatCredCanc(i + 1, 0)
            MatCredCanc(i, 1) = MatCredCanc(i + 1, 1)
        Next i
        MatCredCanc(ContMatCredCanc - 1, 0) = ""
        MatCredCanc(ContMatCredCanc - 1, 1) = ""
        ContMatCredCanc = ContMatCredCanc - 1
    End If
        
End Sub

Private Sub AdicionaCreditoACancelar(ByVal psCtaCod As String, ByVal pnMonto As Double)
    ContMatCredCanc = ContMatCredCanc + 1
    MatCredCanc(ContMatCredCanc - 1, 0) = psCtaCod
    MatCredCanc(ContMatCredCanc - 1, 1) = Format(pnMonto, "#0.00")
End Sub

Private Sub CargaClientesCtaAho(ByVal psCtaCod As String)
Dim i As Integer
Dim J As Integer
Dim sPersCod As String
     LimpiaFlex FEPersCtaAho
     sPersCod = ""
     J = 0
    For i = 0 To nContPersCtaAho - 1
        RPersCtaAho(i).MoveFirst
        If RPersCtaAho(i)!cCtaCod = psCtaCod Then
            Do While Not RPersCtaAho(i).EOF
                FEPersCtaAho.AdicionaFila
                If sPersCod <> RPersCtaAho(i)!cperscod Then
                    J = J + 1
                    FEPersCtaAho.TextMatrix(J, 1) = PstaNombre(RPersCtaAho(i)!Nombre)
                    FEPersCtaAho.TextMatrix(J, 2) = Trim(Mid(RPersCtaAho(i)!Relacion, 1, 30)) 'Trim(RPersCtaAho(i)!Relacion)
                    
                End If
                sPersCod = RPersCtaAho(i)!cperscod
                RPersCtaAho(i).MoveNext
            Loop
            Exit For
        End If
    Next i
End Sub

'Private Function TotalADesembolsar() As Double
'    LblItf.Caption = Format(fgITFDesembolso(CDbl(LblMonPrestamo.Caption)), "0.00")
'    TotalADesembolsar = CDbl(LblMonPrestamo.Caption) - (CDbl(LblMonGastos.Caption) + CDbl(LblMonCancel.Caption) + CDbl(LblItf.Caption))
'    TotalADesembolsar = CDbl(Format(TotalADesembolsar, "#0.00"))
'End Function
Private Function TotalADesembolsar(ByVal pnMontoITf As Double) As Double
    Dim nITF As Double
    nITF = pnMontoITf
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("O0000001", ActxCta.Prod) Then '**ARLO20180712 ERS042 - 2018
    'If (ActxCta.Prod = "515" Or ActxCta.Prod = "516") Then
    '**ARLO20180712 ERS042 - 2018
        nITF = 0
    End If
    If bInstFinanc Then nITF = 0 'JUEZ 20140411
    '*** BRGO 20111012 ************************************************
    
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(nITF))
    If nRedondeoITF > 0 Then
        nITF = nITF - nRedondeoITF
    End If
    
    '*** END BRGO
    'ARCV 24-01-2007
    'TotalADesembolsar = CDbl(LblMonPrestamo.Caption) - (CDbl(LblMonGastos.Caption) + CDbl(LblMonCancel.Caption) + CDbl(LblItf.Caption))
    If sTpoProdCod <> "517" Then
'ALPA 20111213
'        LblItf.Caption = Format(nITF, "#,##0.00")
'        TotalADesembolsar = CDbl(LblMonPrestamo.Caption) - (CDbl(LblMonGastos.Caption) + CDbl(LblMonCancel.Caption) + CDbl(LblItf.Caption) + CDbl(lblPoliza.Caption))
'        '---------------
            '**ARLO20180712 ERS042 - 2018
            Set objProducto = New COMDCredito.DCOMCredito
            If objProducto.GetResultadoCondicionCatalogo("O0000002", ActxCta.Prod) Then
            'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
            '**ARLO20180712 ERS042 - 2018
                TotalADesembolsar = 0
                'LblItf.Caption = Format(nITF, "#,##0.00") 'ALPA 20120413
                LblItf.Caption = Format(0, "#,##0.00") 'ALPA 20120413
            Else
                LblItf.Caption = Format(nITF, "#,##0.00")
                TotalADesembolsar = CDbl(LblMonPrestamo.Caption) - (CDbl(LblMonGastos.Caption) + CDbl(LblMonCancel.Caption) + CDbl(LblItf.Caption) + CDbl(lblPoliza.Caption) + CDbl(lblMonGastoCierre.Caption))
            End If
    '---------------
    TotalADesembolsar = CDbl(Format(TotalADesembolsar, "#0.00"))

    Else
        lblMonITF.Caption = Format(nITF, "#,##0.00")
        TotalADesembolsar = CCur(Me.lblMontoPrestamo.Caption) - CCur(lblMonITF.Caption)
    End If
End Function

Private Function CargaMatrizGastosDesemb() As Variant
Dim MatGastos() As String
Dim i As Integer
Dim nTamanio As Integer
                
    nTamanio = 0
    
    If Trim(FEGastos.TextMatrix(1, 2)) = "" Then
        ReDim MatGastos(0, 0)
        CargaMatrizGastosDesemb = MatGastos
        Exit Function
    End If
    
    For i = 1 To FEGastos.Rows - 1
        If CInt(FEGastos.TextMatrix(i, 1)) = nNroProxDesemb Then
            nTamanio = nTamanio + 1
        End If
    Next i
        
    ReDim MatGastos(nTamanio, 3)
    nTamanio = 0
    For i = 1 To FEGastos.Rows - 1
        If CInt(FEGastos.TextMatrix(i, 1)) = nNroProxDesemb Then
            nTamanio = nTamanio + 1
            MatGastos(nTamanio - 1, 0) = Trim(FEGastos.TextMatrix(i, 1))
            MatGastos(nTamanio - 1, 1) = Trim(FEGastos.TextMatrix(i, 3))
            MatGastos(nTamanio - 1, 2) = Trim(FEGastos.TextMatrix(i, 4))
        End If
    Next i
    CargaMatrizGastosDesemb = MatGastos
End Function

Private Sub CargaCuentasBanco(ByVal psCodPers As String)
Dim o As COMDCredito.DCOMCredito

    Me.CboCta.Clear
    Set o = New COMDCredito.DCOMCredito
    Set RCtasBco = o.RecuperaBancosCtas(psCodPers)
    Do While Not RCtasBco.EOF
        CboCta.AddItem Trim(RCtasBco!cCtaIFDesc)
        RCtasBco.MoveNext
    Loop
    RCtasBco.Close
    Set o = Nothing

End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
'Dim oCredito As COMDCredito.DCOMCredito
'Dim oNegCred As COMNCredito.NCOMCredito
'Dim oCaptacion As COMNCaptaGenerales.NCOMCaptaGenerales
'Dim R As ADODB.Recordset
'Dim RTemp As ADODB.Recordset
'Dim oCalend As COMDCredito.DCOMCalendario
Dim nNroCalen As Integer
Dim oCredito As COMNCredito.NCOMCredito
Dim oDCredito As COMDCredito.DCOMCredito
Dim i As Integer
Dim nTotalDesembolsos As Double

'Variables para Carga de Datos
Dim rsDesemb As ADODB.Recordset
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
Dim nTotalPrestamo As Double 'BRGO 20111111
'ALPA 20111213*****************************
Dim rsProveedores As ADODB.Recordset
Dim rsComisionLeasing As ADODB.Recordset
Dim nMontoIniciales As Currency
'******************************************
Dim sPersCodTitular As String 'JUEZ 20140217
Dim nDestinoCredito As Integer 'FRHU 20140224 RQ14006
Dim oDCreditos As COMDCredito.DCOMCreditos 'JUEZ 20140730

    On Error GoTo ErrorCargaDatos
    
    'Set oCredito = New COMDCredito.DCOMCredito
    'Set R = oCredito.RecuperaDatosDesembolso(psCtaCod)
    Set oCredito = New COMNCredito.NCOMCredito
    'JUEZ 20140217 ****************************************
    Dim oPers  As COMDPersona.UCOMPersona
    Set oDCredito = New COMDCredito.DCOMCredito
    sPersCodTitular = oDCredito.RecuperaTitularCredito(psCtaCod)
    
    Set oPers = New COMDPersona.UCOMPersona
    If oPers.fgVerificaEmpleado(sPersCodTitular) Or oPers.fgVerificaEmpleadoVincualdo(sPersCodTitular) Then
    'END JUEZ *********************************************
        'WIOR 20140203 *********************
        If Not oCredito.ExisteAsignaSaldo(psCtaCod, 2) Then
            MsgBox "El crédito aún no tiene saldo asignado, verificar con el Departamento de Administración de Créditos.", vbInformation, "Aviso"
            CargaDatos = False
            Exit Function
        End If
        'WIOR FIN **************************
    End If
    Set oPers = Nothing
    
    lbPuedeAperturar = True
    'ALPA 20111213*****************
    SSTabDatos.TabVisible(7) = False
    If bLeasing = True Then
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If Not (objProducto.GetResultadoCondicionCatalogo("O0000003", ActxCta.Prod)) Then
    'If Not (ActxCta.Prod = "515" Or ActxCta.Prod = "516") Then
    '**ARLO20180712 ERS042 - 2018
        MsgBox "Este opcion es para producto Leasing, favor digitar un numero de producto Leasing"
        Exit Function
    End If
    ElseIf bLeasing = True Then
        '**ARLO20180712 ERS042 - 2018
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("O0000004", ActxCta.Prod) Then
        'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
        '**ARLO20180712 ERS042 - 2018
            MsgBox "Este Crédito es un Producto Leasing, favor elegir otra modalidad de desembolso"
            Exit Function
        End If
    End If
    
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("O0000005", ActxCta.Prod) Then
    'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
    '**ARLO20180712 ERS042 - 2018
        If gCredDesembLeasing <> sOperacion Then
                MsgBox "Este Crédito es un Producto Leasing, favor elegir otra modalidad de desembolso"
                Exit Function
        End If
    Else
        If gCredDesembLeasing = sOperacion Then
                MsgBox "Esta modalidad de desembolso solo es para Productos Leasing, favor elegir otro crédito"
                Exit Function
        End If
        If gCredDesembLeasing = sOperacion Then
            Set oDCredito = New COMDCredito.DCOMCredito
            Dim oRsLeC As ADODB.Recordset
            Set oRsLeC = New ADODB.Recordset
            Set oRsLeC = oDCredito.ObtenerProcesoLeasing(psCtaCod)
            If oRsLeC.EOF Or oRsLeC.BOF Then
                MsgBox "El crédito desembolso del proveedor"
                Exit Function
            End If
        End If
    End If
    '******************************
    Call oCredito.CargarDatosDesembolso(psCtaCod, gdFecSis, nNroProxDesemb, rsDesemb, rsCredVig, vTotalesVig, rsVigGra, _
                                        rsCuentas, vAgencias, nContPersCtaAho, RPersCtaAho, rsGastos, _
                                        nMontoITF, rsDesemPar, bAmpliacion, sOperacion, pbOperacionEfectivo, pnMontoLavDinero, pnTC, pbExoneradaLavado, _
                                        psPersCodRep, psPersNombreRep, rBancos, rsRelEmp, rsProveedores, rsComisionLeasing)  'CUSCO
    Set oCredito = Nothing
        
    'If Not R.BOF And Not R.EOF Then
    If Not rsDesemb.BOF And Not rsDesemb.EOF Then
        'By Capi 10042008 para deshabilitar boton aperturar
        sTpoProdCod = rsDesemb!cTpoProdCod
        
        'JUEZ 20140730 *********************************************************
        bRevisaDesemb = False
        If sTpoProdCod <> gColConsPFTpoProducto Or _
        (sTpoProdCod = gColConsPFTpoProducto And (rsDesemb!nPersPersoneria <> gPersonaNat Or Left(rsDesemb!cTpoCredCod, 1) = "1" Or Left(rsDesemb!cTpoCredCod, 1) = "2" Or Left(rsDesemb!cTpoCredCod, 1) = "3")) Then
            Set oDCreditos = New COMDCredito.DCOMCreditos
            If Not oDCreditos.VerificaClienteCampania(ActxCta.NroCuenta) Then 'ARLO 20170904 DESCOMENTADO BY ARLO 20171118
                If Not oDCreditos.VerificaRevisaDesembControlAdmCred(ActxCta.NroCuenta) Then
                    MsgBox "El crédito no fue revisado o presenta observaciones para su desembolso por parte de Administración de Créditos. Comunicar al analista respectivo para tomar las medidas del caso.", vbInformation, "Aviso"
                    CargaDatos = False
                    bRevisaDesemb = True
                    Exit Function
            End If 'ARLO 20170904 DESCOMENTADO BY ARLO 20171118
            End If
            Set oDCreditos = Nothing
        End If
        'END JUEZ **************************************************************
        
        'ARLO 20180326 ANEXO 01 - ERS 70******************************************
        Dim rsCompraDeuda As ADODB.Recordset
        nDestinoCredito = IIf(IsNull(rsDesemb!nColocDestino), 0, rsDesemb!nColocDestino)
        If nDestinoCredito = 14 Then
        Set oDCreditos = New COMDCredito.DCOMCreditos
        Set rsCompraDeuda = oDCreditos.VerificaPreDesemCompraDeuda(ActxCta.NroCuenta)
            If Not (rsCompraDeuda.EOF And rsCompraDeuda.BOF) Then
                txtMontoRetirar.Text = Format(rsCompraDeuda!nDesemAComprar, "#,##0.00")
                txtMontoRetirar.Enabled = False
                txtGlosaBloqueo.Enabled = False
                txtGlosaBloqueo.Text = "Por Compra de Deuda"
            Else
            MsgBox "“El supervisor de operaciones debe aprobar el pre desembolso del crédito para cargar los datos.", vbInformation, "Aviso"
            CargaDatos = False
            bRevisaDesemb = True
            Exit Function
            End If
        Set rsCompraDeuda = Nothing
        Set oDCreditos = Nothing
        End If
        'END ARLO **************************************************************
        
        '***FRHU 20140224 RQ14006
        'If rsDesemb!cCtaAhoDesembTercero <> "" Or sTpoProdCod = "517" Then 'BRGO 20111111
        '    lbPuedeAperturar = False
        'End If
        'FRHU 20140224 Para deshabilitar boton aperturar, se agrego que tambien se deshabilite cuando no es persona natural
        If rsDesemb!cCtaAhoDesembTercero <> "" Or sTpoProdCod = "517" Or rsDesemb!nPersPersoneria <> 1 Then
            lbPuedeAperturar = False
        End If
        'FRHU 20140224 Solo se mostrara el monto a retirar en caso de creditos con destino "cambio en estructura de pasivos"
        nDestinoCredito = IIf(IsNull(rsDesemb!nColocDestino), 0, rsDesemb!nColocDestino)
        If nDestinoCredito = 14 Then
            Me.lblMontoRetirar.Visible = True
            Me.txtMontoRetirar.Visible = True
            Me.lblGlosaBloqueo.Visible = True
            Me.txtGlosaBloqueo.Visible = True
        Else
            Me.lblMontoRetirar.Visible = False
            Me.txtMontoRetirar.Visible = False
            Me.lblGlosaBloqueo.Visible = False
            Me.txtGlosaBloqueo.Visible = False
        End If
        'FIN FRHU 20140224
        
        '*** BRGO 20111111 *********************************
        If sTpoProdCod = "517" Then
            vbDesembInfoGas = True
            SSTabDatos.TabVisible(6) = True
            SSTabDatos.TabVisible(3) = False
            SSTabDatos.TabVisible(1) = False
            CmdAperturar.Enabled = False
            nTotalPrestamo = 0
            While Not rsRelEmp.EOF And Not rsRelEmp.BOF
                Select Case CInt(Trim(Right(rsRelEmp!cRelacion, 4)))
                    Case 11: Me.lblMonConcesionario.Caption = Format(rsRelEmp!nMontoAbono, "#,##0.00")
                    Case 12: Me.lblMonOperador.Caption = Format(rsRelEmp!nMontoAbono, "#,##0.00")
                    Case 14: Me.lblMonNotario.Caption = Format(rsRelEmp!nMontoAbono, "#,##0.00")
                    Case 15: Me.lblMonSeguro.Caption = Format(rsRelEmp!nMontoAbono, "#,##0.00")
                    Case 16: Me.lblMonTasacion.Caption = Format(rsRelEmp!nMontoAbono, "#,##0.00")
                    Case 17: Me.lblMonComision.Caption = Format(rsRelEmp!nMontoAbono, "#,##0.00")
                End Select
                If CInt(Trim(Right(rsRelEmp!cRelacion, 4))) <> gColRelPersOperGarantia Then
                    nTotalPrestamo = nTotalPrestamo + rsRelEmp!nMontoAbono
                End If
                rsRelEmp.MoveNext
            Wend
            rsRelEmp.MoveFirst
            Me.lblMontoPrestamo.Caption = Format(nTotalPrestamo, "#,##0.00")
        Else
            vbDesembInfoGas = False
            SSTabDatos.TabVisible(6) = False
            SSTabDatos.TabVisible(3) = True
            SSTabDatos.TabVisible(1) = True
        End If
        '***************************************************

        CargaDatos = True
        'LblMonPrestamo.Caption = Format(R!nPrestamo, "#0.00")
        LblMonPrestamo.Caption = Format(IIf(IsNull(rsDesemb!nMontoADesemb), 0, rsDesemb!nMontoADesemb), "#0.00")
        nNroCalen = rsDesemb!nNroCalen
        LblCodCli.Caption = rsDesemb!cperscod
        sPersCod = rsDesemb!cperscod
        LblNomCli.Caption = PstaNombre(rsDesemb!cPersNombre)
        LblDocNat.Caption = IIf(IsNull(rsDesemb!DNI), "", rsDesemb!DNI)
        LblDocJur.Caption = IIf(IsNull(rsDesemb!Ruc), "", rsDesemb!Ruc)
        LblCliDirec.Caption = IIf(IsNull(rsDesemb!cPersDireccDomicilio), "", rsDesemb!cPersDireccDomicilio)
        LblLineaCred.Caption = IIf(IsNull(rsDesemb!cDescripcion), "", Trim(rsDesemb!cDescripcion))
        LblMoneda.Caption = IIf(IsNull(rsDesemb!cMoneda), "", rsDesemb!cMoneda)
        nNroProxDesemb = rsDesemb!nNroProxDesemb
        LblNroDesemb.Caption = rsDesemb!nNroProxDesemb & "/" & rsDesemb!nTotalDesemb
        'ALPA 20100607 B2*******************
        lblTipoCred.Caption = rsDesemb!cTpoCredDes
        lblTipoProd.Caption = rsDesemb!cTpoProdDes
        '***********************************
        nMontoPrestamoW = CDbl(LblMonPrestamo.Caption) 'WIOR 20120921
        
        'JUEZ 20140411 **************************************
        Dim oDInstFinan As COMDPersona.DCOMInstFinac
        Set oDInstFinan = New COMDPersona.DCOMInstFinac
        bInstFinanc = oDInstFinan.VerificaEsInstFinanc(rsDesemb!cperscod)
        Set oDInstFinan = Nothing
        'END JUEZ *******************************************
        
        'R.Close
        'Set R = Nothing
            
        'ARCV 24-01-2007
        'If rsDesemb!dVencAseg >= gdFecSis Then
        
            '*** PEAC 20080626
            'lblPoliza.Caption = Format(rsDesemb!nMontoPrimaTotal, "#0.00")
            'Comentado Por MAVM 20091015 requerido por IVBA
            'lblPoliza.Caption = Format(rsDesemb!nMontoPrimaTotTC * rsDesemb!NumeroAnhio, "#0.00")
            '*** FIN PEAC 20080626
            
        'Else
            lblPoliza.Caption = "0.00"
        'End If
        '-------------
        
         'Dim nMontoA As Double
         'Dim nInteresA As Double
         'Dim ntotal As Double
         
         'Creditos Vigentes
         LimpiaFlex FECreditosVig
         'Set oNegCred = New COMNCredito.NCOMCredito
         'Set R = oCredito.RecuperaCreditosVigentes(LblCodCli.Caption, , Array(gColocEstVigMor, gColocEstVigNorm, gColocEstVigVenc), Mid(ActxCta.NroCuenta, 9, 1))
         If rsCredVig.RecordCount > 0 Then rsCredVig.MoveFirst
         Do While Not rsCredVig.EOF
            FECreditosVig.AdicionaFila , , True
            'FECreditosVig.TextMatrix(r.Bookmark, 0) = r.Bookmark
                        
            FECreditosVig.TextMatrix(rsCredVig.Bookmark, 2) = rsCredVig!cCtaCod
            FECreditosVig.TextMatrix(rsCredVig.Bookmark, 3) = rsCredVig!nDiasAtraso
        '   nMontoA = R!nSaldo
        '   nInteresA = oNegCred.InteresGastosAFecha(R!cCtaCod, gdFecSis)
        '   ntotal = fgITFCalculaImpuestoNOIncluido(nMontoA + nInteresA, True)
            FECreditosVig.TextMatrix(rsCredVig.Bookmark, 4) = vTotalesVig(rsCredVig.Bookmark) 'Format(ntotal, "#0.00")
            'FECreditosVig.TextMatrix(rsCredVig.Bookmark, 1) = IIf(rsCredVig!nCheck = 0, "", 1)
            FECreditosVig.TextMatrix(rsCredVig.Bookmark, 5) = rsCredVig!PagAmp 'WIOR 20150416
            rsCredVig.MoveNext
        Loop
        'FECreditosVig.Rows = r.RecordCount + 1
        'R.Close
                        
        'Set R = Nothing
        'Set oNegCred = Nothing
            
        'Carga Cuentas de Ahorro de Clientes
        'Set oCaptacion = New COMNCaptaGenerales.NCOMCaptaGenerales
        'Set R = oCaptacion.GetCuentasPersona(LblCodCli.Caption, gCapAhorros, True, True, CInt(Mid(Me.ActxCta.NroCuenta, 9, 1)))
        'Set oCaptacion = Nothing
        LimpiaFlex FECargoAutom
        LimpiaFlex FECtaAhoDesemb
        If rsCuentas.RecordCount > 0 Then rsCuentas.MoveFirst
        Do While Not rsCuentas.EOF
            CmdSeleccionar.Enabled = True
            FECargoAutom.AdicionaFila , , True
            FECtaAhoDesemb.AdicionaFila , , True
            FECargoAutom.TextMatrix(rsCuentas.Bookmark, 0) = rsCuentas.Bookmark
            FECargoAutom.TextMatrix(rsCuentas.Bookmark, 2) = rsCuentas!cCtaCod
            FECtaAhoDesemb.TextMatrix(rsCuentas.Bookmark, 1) = rsCuentas!cCtaCod
        '   Set RTemp = oCredito.RecuperaAgencia(Mid(R!cCtaCod, 4, 2))
            FECargoAutom.TextMatrix(rsCuentas.Bookmark, 3) = vAgencias(rsCuentas.Bookmark) 'Trim(RTemp!cAgeDescripcion)
            FECtaAhoDesemb.TextMatrix(rsCuentas.Bookmark, 2) = vAgencias(rsCuentas.Bookmark) 'Trim(RTemp!cAgeDescripcion)
            'By Capi 11012008
            FECtaAhoDesemb.TextMatrix(rsCuentas.Bookmark, 3) = IIf(rsCuentas!Exonerada = "", "Afecta", "Exonerada")
            FECtaAhoDesemb.TextMatrix(rsCuentas.Bookmark, 4) = rsCuentas!nTpoPrograma 'JUEZ 20141114
            
        '   RTemp.Close
        '   Set RTemp = Nothing
            rsCuentas.MoveNext
        Loop
        FECargoAutom.Rows = rsCuentas.RecordCount + 1
        If rsCuentas.RecordCount <= 0 Then
            FECargoAutom.Enabled = False
        Else
            FECargoAutom.Enabled = True
        End If
    '   R.Close
    '   Set R = Nothing
        FECtaAhoDesemb.row = 1
            
        'Carga Clientes Relacionados con las Ctas de Ahorros
        'nContPersCtaAho = 0
    '    If Trim(FECtaAhoDesemb.TextMatrix(1, 1)) <> "" Then
    '        For i = 1 To FECtaAhoDesemb.Rows - 1
    '            nContPersCtaAho = nContPersCtaAho + 1
    '            ReDim Preserve RPersCtaAho(nContPersCtaAho)
    '            Set oCaptacion = New COMNCaptaGenerales.NCOMCaptaGenerales
    '            Set RPersCtaAho(nContPersCtaAho - 1) = oCaptacion.GetPersonaCuenta(FECtaAhoDesemb.TextMatrix(1, 1))
    '            Set oCaptacion = Nothing
    '        Next i
    '    End If
            
        LimpiaFlex FEPersCtaAho
        If nContPersCtaAho > 0 Then
            RPersCtaAho(0).MoveFirst
            'Do While Not RPersCtaAho(0).EOF
            '    FEPersCtaAho.AdicionaFila
            '    FEPersCtaAho.TextMatrix(RPersCtaAho(0).Bookmark, 1) = PstaNombre(RPersCtaAho(0)!nombre)
            '    FEPersCtaAho.TextMatrix(RPersCtaAho(0).Bookmark, 2) = Trim(Mid(RPersCtaAho(0)!Relacion, 1, 30))
            '    RPersCtaAho(0).MoveNext
            'Loop
            Call CargaClientesCtaAho(RPersCtaAho(0)!cCtaCod)
        End If
    '    Set oCredito = Nothing
    
        'Carga distribución de FMV y TP
        nMontoGastoCierre = 0
        If (sTpoProdCod = "802" Or sTpoProdCod = "806") Then
            Dim rsMiViv As ADODB.Recordset
            Set rsMiViv = oDCredito.ObtenerDatosNuevoMIVIVIENDA(ActxCta.NroCuenta, gColocEstAprob)
            
            If Not (rsMiViv.EOF And rsMiViv.BOF) Then
                nMontoGastoCierre = CDbl(rsMiViv!nGastoCierre)
            End If
            Set rsMiViv = Nothing
        End If
        nMontoGastoCierre = CDbl(Format(nMontoGastoCierre, "#0.00"))
        lblMonGastoCierre.Caption = Format(nMontoGastoCierre, "#0.00")
        
            
           'Halla el Gasto del Desembolso
           nMontoGastos = 0
           LimpiaFlex FEGastos
    '        Set oCalend = New COMDCredito.DCOMCalendario
    '        Set R = oCalend.RecuperaGastosCuotaDesemb(psCtaCod, nNroCalen, gColocCalendAplDesembolso, nNroProxDesemb)
    '        Set oCalend = Nothing
            Do While Not rsGastos.EOF
                FEGastos.AdicionaFila
                FEGastos.TextMatrix(rsGastos.Bookmark, 1) = Trim(Str(rsGastos!nCuota))
                FEGastos.TextMatrix(rsGastos.Bookmark, 2) = rsGastos!cGasto
                FEGastos.TextMatrix(rsGastos.Bookmark, 3) = Format(rsGastos!nMonto, "#0.00")
                FEGastos.TextMatrix(rsGastos.Bookmark, 4) = Trim(Str(rsGastos!nPrdConceptoCod))
                nMontoGastos = nMontoGastos + rsGastos!nMonto
                rsGastos.MoveNext
            Loop
     '       R.Close
     '       Set R = Nothing
            nMontoGastos = CDbl(Format(nMontoGastos, "#0.00"))
            LblMonGastos.Caption = Format(nMontoGastos, "#0.00")
            
      '     lblMonDesemb.Caption = Format(TotalADesembolsar, "#0.00")
            lblMonDesemb.Caption = Format(TotalADesembolsar(nMontoITF), "#0.00")
            Me.lblTotalFinanciar.Caption = Format(TotalADesembolsar(nMontoITF), "#0.00") 'BRGO 20111124
            
            'Carga todos los desembolsos Pp
            nTotalDesembolsos = 0
            LimpiaFlex FEDesembolsos
            'Set oCalend = New COMDCredito.DCOMCalendario
            'Set R = oCalend.RecuperaListaDesembolsosParciales(psCtaCod)
            'Set oCalend = Nothing
            Do While Not rsDesemPar.EOF
                FEDesembolsos.AdicionaFila
                FEDesembolsos.TextMatrix(rsDesemPar.Bookmark, 1) = Trim(Str(rsDesemPar!nCuota))
                FEDesembolsos.TextMatrix(rsDesemPar.Bookmark, 2) = Format(rsDesemPar!dVenc, "dd/MM/YYYY")
                FEDesembolsos.TextMatrix(rsDesemPar.Bookmark, 3) = Format(rsDesemPar!nMonto, "#0.00")
                FEDesembolsos.TextMatrix(rsDesemPar.Bookmark, 4) = Trim(rsDesemPar!cActual)
                nTotalDesembolsos = nTotalDesembolsos + rsDesemPar!nMonto
                rsDesemPar.MoveNext
            Loop
            'R.Close
            'Set R = Nothing
                        
            lblTotal.Caption = Format(nTotalDesembolsos, "0.00")
            '
             
           ' Mensaje que indica que el credito es una ampliacion
            'Dim oAmpliacion As COMDCredito.DCOMAmpliacion
            'Dim bAmpliacion As Boolean
            
            'Set oAmpliacion = New COMDCredito.DCOMAmpliacion
            'bAmpliacion = oAmpliacion.ValidaCreditoaAmpliar(psCtaCod)
            'Set oAmpliacion = Nothing
            
            'ARCV 12-02-2007 (AHORA ES AUTOMATICO)
            If bAmpliacion = True Then
                lblampliado.Caption = 1 'CTI3 ERS082-2018
                MsgBox "Este credito es un credito de ampliacion " & vbCrLf & _
                       "Por favor verifique que el credito a cancelar este seleccionado", vbInformation, "AVISO"
            End If
            
            CboBancos.Clear
            Do While Not rBancos.EOF
                CboBancos.AddItem rBancos!cPersNombre & space(100) & rBancos!cperscod
                rBancos.MoveNext
            Loop
                    
            'Set R = oCredito.RecuperaCreditosVigentesGrabados(ActxCta.NroCuenta)
            Do While Not rsVigGra.EOF
                For i = 1 To FECreditosVig.Rows - 1
                    If Trim(FECreditosVig.TextMatrix(i, 2)) = Trim(rsVigGra!cCtaCodRef) Then
                        FECreditosVig.TextMatrix(i, 1) = "1"
                        bOnCellCheck = True 'FRHU 20140424 TI-ERS015-2014
                        Call FECreditosVig_OnCellCheck(i, 1)
                    End If
                Next i
                rsVigGra.MoveNext
            Loop
            'VAPA 20171209 FROM 60
            'If bAmpliacion = True Then
                'ActualizaCreditoAmpliadoNew lsCuentaCredCan, psCtaCod
            'End If
            'VAPA END Comentado by NAGL 20180609
            
            
            'R.Close
            'ALPA 20111213**********************************
            Call LimpiaFlex(FEProveedores)
            Call LimpiaFlex(FEPagosIniciales)
            '***********************************************
            CmdVisualizar.Enabled = True
            cmdVisualiza.Enabled = True
            'ALPA 20110608****************************************
            nMontoIniciales = 0
            '**ARLO20180712 ERS042 - 2018
            Set objProducto = New COMDCredito.DCOMCredito
            If objProducto.GetResultadoCondicionCatalogo("O0000006", ActxCta.Prod) Then
            'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
            '**ARLO20180712 ERS042 - 2018
                SSTabDatos.TabVisible(7) = True
               
                nTotalProveedores = 0
                nTotalCuotaIniciales = 0
                LimpiaFlex FEProveedores
                Do While Not rsProveedores.EOF
                    FEProveedores.AdicionaFila
                    FEProveedores.TextMatrix(rsProveedores.Bookmark, 1) = rsProveedores!cPersCodProve
                    FEProveedores.TextMatrix(rsProveedores.Bookmark, 2) = rsProveedores!cPersNombre
                    FEProveedores.TextMatrix(rsProveedores.Bookmark, 4) = Format(rsProveedores!nValorVenta, "#0.00")
                    nTotalProveedores = nTotalProveedores + 1
                    rsProveedores.MoveNext
                Loop
                
                Do While Not rsComisionLeasing.EOF
                    FEPagosIniciales.AdicionaFila
                    FEPagosIniciales.TextMatrix(rsComisionLeasing.Bookmark, 1) = rsComisionLeasing!nPrdConceptoCod
                    FEPagosIniciales.TextMatrix(rsComisionLeasing.Bookmark, 2) = rsComisionLeasing!cDescripcion
                    FEPagosIniciales.TextMatrix(rsComisionLeasing.Bookmark, 3) = Format(rsComisionLeasing!nMonto, "#0.00")
                    nMontoIniciales = nMontoIniciales + rsComisionLeasing!nMonto
                    nTotalCuotaIniciales = nTotalCuotaIniciales + 1
                    rsComisionLeasing.MoveNext
                Loop
            'End If 'JUEZ 20130930 Se colocó el recalculo de los labels lblTotalIniciales y lblMonDesemb dentro de la condición, sólo para Leasing
                lblTotalIniciales.Caption = Format(nMontoIniciales, "#0.00")
                lblMonDesemb.Caption = Format(TotalADesembolsar(nMontoITF), "#0.00")
            End If
            '*****************************************************
            'RECO20150526 ERS023-2015*****************************
            Dim oGasto As New COMNCredito.NCOMGasto
            Dim rsGasMYPE As ADODB.Recordset
            Set rsGasMYPE = oGasto.RecuperaGastosMYPE(ActxCta.NroCuenta)
            If Not (rsGasMYPE.BOF And rsGasMYPE.EOF) Then
                frGastoMYPE.Enabled = True
            Else
                frGastoMYPE.Enabled = False
            End If
            'RECO FIN *********************************************
			
			'JOEP20211013 ERS048-2021 Restrincion de Desembolso >3500
			Dim objResPago As COMDCredito.DCOMCreditos
			Dim rsResPago As ADODB.Recordset
			Set objResPago = New COMDCredito.DCOMCreditos
				
			Set rsResPago = objResPago.ResPago(ActxCta.NroCuenta, lblMonDesemb, LblMonPrestamo, sOperacion)
			If Not (rsResPago.BOF And rsResPago.EOF) Then
				If rsResPago!cMsgBox <> "" Then
					MsgBox rsResPago!cMsgBox, vbInformation, "Aviso"
					CargaDatos = False
					bRevisaDesemb = True
				   Exit Function
				End If
			End If
			
			Set objResPago = Nothing
			RSClose rsResPago
			'JOEP20211013 ERS048-2021 Restrincion de Desembolso >3500

    Else
        CargaDatos = False
        'R.Close
        'Set R = Nothing
        'Set oCredito = Nothing
    End If
    Exit Function
    
ErrorCargaDatos:
    MsgBox err.Description, vbCritical, "Aviso"

End Function

Private Sub HabilitaDesembolso(ByVal pbHabilita As Boolean)
    FraCuenta.Enabled = Not pbHabilita
    FraCliente.Enabled = pbHabilita
    FraCredito.Enabled = pbHabilita
    SSTCtasAho.Enabled = pbHabilita
    'By Capi 10042008
    If lbPuedeAperturar = False Then
        CmdAperturar.Enabled = Not pbHabilita
        fraAperturaCtaAhorro.Enabled = Not pbHabilita 'FRHU 20140228 RQ14006
        chkAperCtaAhorro.Enabled = Not pbHabilita 'FRHU 20140228 RQ14006
    Else
        CmdAperturar.Enabled = pbHabilita
        'fraAperturaCtaAhorro.Enabled = pbHabilita 'FRHU 20140228 RQ14006
        chkAperCtaAhorro.Enabled = pbHabilita 'FRHU 20140228 RQ14006
    End If
    CmdSeleccionar.Enabled = pbHabilita
    FraDatosDesemb.Enabled = True
    CmdDesemb.Enabled = pbHabilita
    CmdCancelar.Enabled = pbHabilita
    CmdSalir.Enabled = Not pbHabilita
    FECreditosVig.lbEditarFlex = pbHabilita
    FECargoAutom.lbEditarFlex = pbHabilita
    vbCuentaNueva = False
    If sTpoProdCod = "517" Then 'BRGO 20111124
        cmdSale.Enabled = pbHabilita
        cmdDesembolso.Enabled = pbHabilita
        cmdCancela.Enabled = pbHabilita
    End If
End Sub

Private Sub LimpiaPantalla()
Dim i As Integer
    'lsCuentaCredCan = "" ' VAPA 20171211 FROM 60 Comentado by NAGL 20180609
    LblTipoAbono.Caption = ""
    ActxCta.NroCuenta = ""
    ContMatCargoAutom = 0
    ContMatCredCanc = 0
    nMontoGastos = 0
    nMontoGastoCierre = 0
    nMontoCredCanc = 0
    For i = 0 To 99
        MatCredCanc(i, 0) = ""
        MatCredCanc(i, 1) = ""
    Next i
    ReDim MatCargoAutom(0)
    sCtaAho = ""
    pnFilaSelecCtaAho = -1
    nNroProxDesemb = 0
    vbCuentaNueva = False
    ReDim RPersCtaAho(0)
    nContPersCtaAho = 0
    psCodIF = ""
    Set pRSRela = Nothing
    pnTipoCuenta = 0
    psNroDoc = ""
    pnTipoTasa = 0
    pbDocumento = False
    pnPersoneria = 0
    LimpiaControles Me, True
    LimpiaFlex FECreditosVig
    LimpiaFlex FECargoAutom
    LimpiaFlex FEGastos
    LimpiaFlex FECtaAhoDesemb
    LimpiaFlex FEPersCtaAho
    LimpiaFlex FEDesembolsos
    lblTotal.Caption = ""
    LblCtaAbo.Caption = ""
    pnFilaSelecCtaAho = -1
    LblMonGastos.Caption = "0.00"
    LblMonCancel.Caption = "0.00"
    LblMonPrestamo.Caption = "0.00"
    lblMonDesemb.Caption = "0.00"
    
    '***BRGO 20111109 *******************
    lblMonConcesionario.Caption = "0.00"
    lblMonOperador.Caption = "0.00"
    lblMonNotario.Caption = "0.00"
    lblMonSeguro.Caption = "0.00"
    lblMonComision.Caption = "0.00"
    lblMontoPrestamo.Caption = "0.00"
    Me.lblMonTasacion.Caption = "0.00"
    lblMonITF = "0.00"
    Me.SSTabDatos.TabVisible(6) = False
    Me.SSTabDatos.TabVisible(1) = True
    Me.SSTabDatos.TabVisible(3) = True
    Me.SSTabDatos.Tab = 3
    LblTipoAbono.Caption = ""
    Call FECtaAhoDesemb.BackColorRow(vbWhite)
    CmdSeleccionar.Visible = True
    CmdDeseleccionar.Visible = False
    Set rsRelEmp = Nothing
    '************************************
    
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    'ALPA 20111213***********************
    Call LimpiaFlex(FEProveedores)
    Call LimpiaFlex(FEPagosIniciales)
    'SSTabDatos.TabVisible(6) = False 'ALPA 20110608
    '************************************
    'FRHU20140228 RQ14006
    txtMontoRetirar.Text = "0.00"
    lblMontoRetirar.Visible = False
    txtMontoRetirar.Visible = False
    Me.lblGlosaBloqueo.Visible = False
    Me.txtGlosaBloqueo.Visible = False
    'FIN FRHU20140228 RQ14006
    bInstFinanc = False 'JUEZ 20140411
    
    lblampliado.Caption = 0 'CTI3
    lbPuedeAperturar = False
    fraAperturaCtaAhorro.Enabled = False
    chkAperCtaAhorro.Enabled = False
    chkAperCtaAhorro.value = 0
End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
'MARG20171205**********************
    Dim dFechaSis As Date
    dFechaSis = CDate(validarFechaSistema)

    If gdFecSis <> dFechaSis Then
        MsgBox "La Fecha de tu sesión en el Negocio no coincide con la fecha del Sistema", vbCritical, "Aviso"
        Call SalirSICMACMNegocio
        Unload Me
        End
    End If
'END MARG**************************
Dim oCredito As COMNCredito.NCOMCredito
'----- MADM
Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset
Dim rsRefinanciado As ADODB.Recordset 'LUCV20160912; Según ERS004-2016
Dim oDCOMCredito As New COMDCredito.DCOMCredito 'LUCV201609012, Según ERS004-2016
'----- MADM
'FRHU 20140225 RQ14007
Dim oCredDestino As New COMDCredito.DCOMCredito
Dim RDes As ADODB.Recordset
'FIN FRHU 20140225
Dim sError As String
    If KeyAscii = 13 Then
        'CTI520210511 ***
        If (ActxCta.Prod = "802" Or ActxCta.Prod = "806") And Not vbDesembCC Then
            MsgBox "El crédito no se puede desembolsar en efectivo. Seleccionar Desembolso con Abono a Cuenta", vbInformation, "Aviso"
            Call cmdCancelar_Click
            Exit Sub
        End If
        'END CTI5 *******
        '**ARLO20180712 ERS042 - 2018
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("O0000007", ActxCta.Prod) And Not vbDesembCC Then
        'If ActxCta.Prod = "517" And Not vbDesembCC Then 'BRGO 20111115
        '**ARLO20180712 ERS042 - 2018
            MsgBox "El crédito no se puede desembolsar en efectivo. Seleccionar Desembolso con Abono a Cuenta", vbInformation, "Aviso"
            Call cmdCancelar_Click
            Exit Sub
        End If
        '*****-> LUCV20160912, Según ERS004-2016
        Set rsRefinanciado = oDCOMCredito.RecuperaColocacRefinanc(ActxCta.NroCuenta)
        If Not (rsRefinanciado.BOF Or rsRefinanciado.EOF) Then
            MsgBox "Este crédito es refinanciado, desembolso no permitido." & Chr(13) & " -Ingresar a la Operacion: Vigencia de créditos refinanciados.", vbInformation, "Aviso"
            Exit Sub
        End If
        '<-***** LUCV20160912
        
        '*** FRHU 20140225 RQ14007
        ActxCta.NroCuenta = Replace(ActxCta.NroCuenta, "'", "") 'WIOR 20150617
        Set RDes = oCredDestino.RecuperaColocacCred(ActxCta.NroCuenta)
        Set oCredDestino = Nothing
            
        If Not (RDes.EOF And RDes.BOF) Then 'WIOR 20150617
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
        End If 'WIOR 20150617
        'FIN FRHU 20140225
        Set oCredito = New COMNCredito.NCOMCredito
        sError = oCredito.ValidaCargaDatosDesembolso(ActxCta.NroCuenta, gdFecSis)
        If sError <> "" Then
            MsgBox sError, vbInformation, "Aviso"
            Call cmdCancelar_Click
            Exit Sub
        End If
        Set oCredito = Nothing
        If CargaDatos(ActxCta.NroCuenta) Then
            HabilitaDesembolso True
            If vbDesembCC Then
                'ALPA 20111213***************************
                'SSTabDatos.Tab = 2
                '**ARLO20180712 ERS042 - 2018
                Set objProducto = New COMDCredito.DCOMCredito
                If Not (objProducto.GetResultadoCondicionCatalogo("O0000008", ActxCta.Prod)) Then
                'If Not (ActxCta.Prod = "515" Or ActxCta.Prod = "516") Then
                '**ARLO20180712 ERS042 - 2018
                    SSTabDatos.Tab = 2
                End If
                '*****************************************

                'By Capi 10042008
                If CmdAperturar.Enabled = True Then
                    'CmdAperturar.SetFocus 'Esto ya estab 10042008 'FRHU20140228 RQ14006
                Else
                    CmdSeleccionar.SetFocus
                End If
                '
            Else
                CmdDesemb.SetFocus
                SSTabDatos.Tab = 3
               
            End If
            '************ firma madm
         Set lafirma = New frmPersonaFirma
         Set ClsPersona = New COMDPersona.DCOMPersonas
        
         Set Rf = ClsPersona.BuscaCliente(frmCredPersEstado.vcodper, BusquedaCodigo)
         If Not Rf.BOF And Not Rf.EOF Then
            If Rf!nPersPersoneria = 1 Then
            Call frmPersonaFirma.Inicio(Trim(frmCredPersEstado.vcodper), Mid(frmCredPersEstado.vcodper, 4, 2), True)
            End If
         End If
         Set Rf = Nothing
             '************
        'APRI20171009 ERS028-2017
        If sOperacion = gCredDesembEfec Then
            frmSegSepelioAfiliacion.Inicio ActxCta.NroCuenta
        End If
        'END IF
        Else
            HabilitaDesembolso False
            Call LimpiaPantalla
            If Not bRevisaDesemb Then 'JUEZ 20140730
                MsgBox "No pudo Encontrar  el Credito, posiblemente aun no esta Aprobado", vbInformation, "Aviso"
            End If
        End If
    End If
End Sub

Private Sub CboBancos_Click()
    Call CargaCuentasBanco(Right(CboBancos.Text, 13))
End Sub
'JUEZ 20141114 Nuevos parámetros **********************
Private Sub cboPrograma_LostFocus()
Dim clsDef As New COMNCaptaGenerales.NCOMCaptaDefinicion
If Trim(cboPrograma.Text) <> "" Then
    If Not clsDef.GetCapParametroNew(gCapAhorros, CInt(Trim(Right(cboPrograma.Text, 1))))!bAplicaDesembCred Then
        MsgBox "El producto seleccionado no está configurado para destino de desembolso. Favor de comunicarse con el Dpto. de Ahorros", vbInformation, "Aviso"
        cboPrograma.ListIndex = -1
    End If
End If
End Sub
'END JUEZ *********************************************
'FRHU 20140226 RQ1406
Private Sub chkAperCtaAhorro_Click()
    If chkAperCtaAhorro.value = 1 Then
        If CmdDeseleccionar.Visible Then
            Call CmdDeseleccionar_Click
        End If
        fraAperturaCtaAhorro.Enabled = True
        CmdSeleccionar.Enabled = False
    Else
        fraAperturaCtaAhorro.Enabled = False
        CmdSeleccionar.Enabled = True
    End If
End Sub
'FIN FRHU 20140226 RQ1406

Private Sub CmdAperturar_Click()
    'FRHU 20140226 RQ14006
    Dim nDesembolso As Integer
    Dim sCtsAhorroNueva As String
    Dim nTasaNominal As Double
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
   
    'Dim oCaptacion As COMNCaptaGenerales.NCOMCaptaGenerales
    'Dim rsCuentas As ADODB.Recordset
    Dim sCtaCre As String
    Dim fila As Integer
    nDesembolso = 1
    sCtsAhorroNueva = ""
    sCtaCre = ActxCta.NroCuenta
    'FIN FRHU 20140226 RQ14006
    nProgAhorros = CInt(Trim(Right(cboPrograma.Text, 1))) 'JUEZ 20141114
    
    If CmdDeseleccionar.Visible Then
        Call CmdDeseleccionar_Click
    End If
    
    'vbCuentaNueva = True 'FRHU 20140226 RQ14006
    Set rsRel = New ADODB.Recordset
    
    frmCapAperturas.IniciaDesembAbonoCta gCapAhorros, gAhoApeEfec, "", CInt(Mid(Me.ActxCta.NroCuenta, 9, 1)), _
         LblCodCli.Caption, LblNomCli.Caption, CDbl(LblMonPrestamo.Caption), pRSRela, _
        pnTasa, pnPersoneria, pnTipoCuenta, pnTipoTasa, pbDocumento, psNroDoc, psCodIF, False, _
        psPersCodRep, psPersNombreRep, MatTitulares, nProgAhorros, nMontoAbonar, nPlazoAbonar, sPromotorAho
    
    'Set rsRel = pRSRela 'FRHU 20140226 RQ14006
    
    'FRHU 20140226 RQ 14006
'    LblTipoAbono.Caption = "NUEVA"
'    LblCtaAbo.Caption = "NUEVA"
'    LblCtaAbo.Alignment = 2
'    LblTipoCta.Caption = "Propia"
'    LblAgeAbono.Caption = gsNomAge
    
    
    If Not pRSRela Is Nothing Then
        'HabilitaDesembolso False
        'Call LimpiaPantalla
        'ActxCta.NroCuenta = sCtaCre
        'Call ActxCta_KeyPress(13)
        vbCuentaNueva = True
        Set rsRel = pRSRela
        LblTipoAbono.Caption = "NUEVA"
        LblCtaAbo.Caption = "NUEVA"
        LblCtaAbo.Alignment = 2
        LblTipoCta.Caption = "Propia"
        LblAgeAbono.Caption = gsNomAge
        nProgAhorros = CInt(Trim(Right(cboPrograma.Text, 1)))
        Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        nTasaNominal = clsDef.GetCapTasaInteres(gCapAhorros, CInt(Mid(Me.ActxCta.NroCuenta, 9, 1)), pnTipoTasa, , CDbl(LblMonPrestamo.Caption), gsCodAge, , nProgAhorros)
        pnTasa = nTasaNominal
    'JUEZ 20141114 *******************************
    Else
        vbCuentaNueva = False
        LblTipoAbono.Caption = ""
        LblCtaAbo.Caption = ""
        LblTipoCta.Caption = ""
        LblAgeAbono.Caption = ""
        pnTasa = 0
    'END JUEZ ************************************
    End If
    'FRHU
End Sub

Private Sub cmdCancelar_Click()
    HabilitaDesembolso False
    Call LimpiaPantalla
End Sub

Private Sub CmdDeseleccionar_Click()
    sCtaAho = ""
    CmdSeleccionar.Visible = True
    CmdDeseleccionar.Visible = False
    If pnFilaSelecCtaAho <> -1 Then
        FECtaAhoDesemb.row = pnFilaSelecCtaAho
        Call FECtaAhoDesemb.BackColorRow(vbWhite)
    End If
    pnFilaSelecCtaAho = -1
    LblCtaAbo.Caption = ""
    LblTipoCta.Caption = ""
    LblAgeAbono.Caption = ""
    LblTipoAbono.Caption = ""
End Sub

Private Sub CmdDesemb_Click()

Dim oNCredito As COMNCredito.NCOMCredito
Dim Docs As Variant
Dim sError As String
Dim MatGastos As Variant
Dim nMovNro As Long
Dim i As Integer
'Dim MatDatosAho(14) As String
Dim MatDatosAho(18) As String 'FRHU20140228 RQ14006

Dim dUltimaFecdes As Date
Dim dProxFecDes As Date
    
'ARCV 04-11-2006
Dim nITFDesembolso As Double
Dim nMontoDesembolso As Double
'-------------
    
'VARIABLES PARA EL USO DEL COMPONENTE

Dim sMensaje As String
Dim sPersLavDinero As String
Dim nMontoLavDinero As Double
Dim sImpreDocs As String
Dim sImpreBoletaAbono As String
Dim sImpreBoletaGasto As String
Dim sImpreBoletaCancel As String
Dim sImpreLavado As String
'ARCV 07-06-2006
Dim sImpreDocsBoleta As String
'----------
Dim sVisPersLavDinero As String 'DAOR 20070512
'MADM 20110504
Dim oDCredito As COMDCredito.DCOMCredito
Dim RDCredito As ADODB.Recordset
Dim nPeriodoGraciaAnt As Integer
Dim nTipoPeriodo As Integer
Dim lnDescripCuota() As String
'ALPA 20111213***************************
Dim psProveedorLeasing() As String
Dim psCuotaInicialLeasing() As String
'****************************************
'MADM 20110504
Dim nRedondITFCanc As Double 'BRGO 20111021
Dim nITFCtaCanc As Double 'BRGO 20111021
Dim lsMensajeGrabar As String 'ALPA 20120509
Dim MatCuentaGastoCierre As Variant 'CTI520210408
   'RECO20150601 ERS023-2015 ******************************
   If frGastoMYPE.Enabled = True And txtMultMype.Text = "" Then
        MsgBox "Debe ingresar el número de la póliza", vbInformation, "Alerta"
        Exit Sub
   End If
   'RECO FIN***********************************************
   
   'JOEP20170731-Valida si el credito tiene vinculado la garantia ******************************
   Dim oDValCobGaran As COMDCredito.DCOMCredito
   Dim rsValiGarantia As ADODB.Recordset
   Set rsValiGarantia = New ADODB.Recordset
   Set oDValCobGaran = New COMDCredito.DCOMCredito
   Set rsValiGarantia = oDValCobGaran.RecuperaGarantiasCredito(ActxCta.NroCuenta)
   If (rsValiGarantia.EOF And rsValiGarantia.BOF) Then
        MsgBox "Por favor, comuníquese con el Analista de Créditos, falta vincular la Garantía con el crédito.", vbInformation, "AVISO"
        Exit Sub
   End If
    Set oDValCobGaran = Nothing
    rsValiGarantia.Close
    Set rsValiGarantia = Nothing
   'JOEP20170731***********************************************
   
    'RIRO 20200913 ********************************************
    If ContMatCredCanc > 0 Then
        Dim oCreditoTmp As COMNCredito.NCOMCredito
        Set oCreditoTmp = New COMNCredito.NCOMCredito
        Dim bValidaActualizacionLiq As Boolean
        bValidaActualizacionLiq = False
    
        For i = 0 To ContMatCredCanc - 1
            bValidaActualizacionLiq = False
            bValidaActualizacionLiq = oCreditoTmp.VerificaActualizacionLiquidacion(MatCredCanc(i, 0))
            If Not bValidaActualizacionLiq Then
                MsgBox "El crédito no tiene actualizados sus datos de liquidación, no podrá realizar el desembolso " & _
                "a menos que actualice estos datos. Deberá comunicarse con el área de T.I.", vbExclamation, "Aviso"
                
                Set oCreditoTmp = Nothing
                Exit Sub
            End If
        Next i
    End If
    'RIRO 20200913 ********************************************
      
   'FRHU 20140228 RQ14006
   If Me.txtMontoRetirar.Visible = True Then
         If Me.txtMontoRetirar.Text = "" Or Me.txtMontoRetirar.Text = 0 Then
            MsgBox "Debe Ingresar el Monto a Retirar", vbInformation, "AVISO"
            Exit Sub
         Else
            If CDbl(txtMontoRetirar.Text) >= CDbl(lblMonDesemb.Caption) Then
                'MsgBox "El Monto a Retirar no debe ser mayor al Monto a Desembolsar", vbInformation, "AVISO" 'FRHU 20140807: Observacion
                MsgBox "El Monto a Retirar no debe ser mayor o igual al Monto a Desembolsar", vbInformation, "AVISO"
                Exit Sub
            End If
         End If
   Else
        Me.txtMontoRetirar.Text = 0
   End If
   If chkAperCtaAhorro.value = 1 Then
        If cboPrograma.ListIndex = -1 Then
            MsgBox "Debe Seleccionar un Producto de Ahorro", vbInformation, "AVISO"
            Exit Sub
        End If
        Call CmdAperturar_Click
   End If
   'FIN FRHU 20140228
   
   If vbDesembCC Then
       If Not vbCuentaNueva And Len(Trim(sCtaAho)) <= 0 Then
            If Not vbDesembInfoGas Then
                MsgBox "Aperture o Seleccione una Cuenta de Ahorros para Depositar el Desembolso", vbInformation, "Aviso"
            Else
                MsgBox "Seleccione una Cuenta de Ahorros Ecotaxi para Abonar el Recaudo", vbInformation, "Aviso"
            End If
            Exit Sub
       End If
   End If
    
   'CTI520210408 ***
   If (CDbl(Me.lblMonGastoCierre.Caption) > 0) Then
        MatCuentaGastoCierre = ObtenerMatrizGastoCierre()
   End If
   'END CTI5 *******
    
    'ALPA 20111213******************************************
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("O0000009", ActxCta.Prod) Then
    'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
    '**ARLO20180712 ERS042 - 2018
        Dim lsNombreProveedor As String
        For i = 1 To nTotalProveedores
            ReDim Preserve psProveedorLeasing(0 To 3, 0 To i)
            psProveedorLeasing(0, i) = FEProveedores.TextMatrix(i, 1) 'CodProve
            psProveedorLeasing(1, i) = FEProveedores.TextMatrix(i, 2) 'NomProve
            psProveedorLeasing(2, i) = FEProveedores.TextMatrix(i, 3) 'cCtaCod
            psProveedorLeasing(3, i) = FEProveedores.TextMatrix(i, 4) 'nMontoDe
        Next i
        For i = 1 To nTotalCuotaIniciales
            ReDim Preserve psCuotaInicialLeasing(0 To 2, 0 To i)
            psCuotaInicialLeasing(0, i) = FEPagosIniciales.TextMatrix(i, 1)
            psCuotaInicialLeasing(1, i) = FEPagosIniciales.TextMatrix(i, 2)
            psCuotaInicialLeasing(2, i) = FEPagosIniciales.TextMatrix(i, 3)
        Next i
   End If
   '********************************************************
   '****************************************
   
   If vbDesembCheque Then
        If Me.CboBancos.ListIndex = -1 Then
            MsgBox "Debe Seleccionar un Banco para esta operacion", vbInformation, "Aviso"
            Exit Sub
        End If
        
        If Me.CboCta.ListIndex = -1 Then
            MsgBox "Debe Seleccionar una Cuenta de Banco para esta operacion", vbInformation, "Aviso"
            Exit Sub
        End If
        
        If Trim(TxtCheque.Text) = "" Then
            MsgBox "Debe Digitar el Numero de Cheque", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
'ARCV 03-11-2006
    nTipoPeriodo = 0
    nITFDesembolso = 0
    nMontoDesembolso = 0
    nPeriodoGraciaAnt = 0
    'ALPA 20130130******************************
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("O0000010", ActxCta.Prod) Then
    'If (ActxCta.Prod = "517") Then
    '**ARLO20180712 ERS042 - 2018
        LblItf.Caption = lblMonITF.Caption
    End If
    '*******************************************
    nITFDesembolso = CDbl(LblItf.Caption) + nRedondeoITF
    nMontoDesembolso = CDbl(lblMonDesemb.Caption)
    'JUEZ 20130930 Para cobrar ITF en la Cancelacion sólo cuando es desembolso efectivo ****
    If Not vbDesembCC Then
        If ContMatCredCanc > 0 Then
           For i = 0 To ContMatCredCanc - 1
               '*** BRGO 20111012 ************************************************
                nITFCtaCanc = fgITFDesembolso(MatCredCanc(i, 1))
                nRedondITFCanc = fgDiferenciaRedondeoITF(nITFCtaCanc)
                If nRedondITFCanc > 0 Then
                    nITFCtaCanc = Format(nITFCtaCanc - nRedondITFCanc, "#,##0.00")
                End If
                nITFDesembolso = nITFDesembolso - nITFCtaCanc
                '*** END BRGO ***********************
            Next i
        End If
    End If
    'END JUEZ ******************************************************************************
'---------------
    'JUEZ 20140411 ***************
    If bInstFinanc Then
        nITFCtaCanc = 0
        nITFDesembolso = 0
    End If
    'END JUEZ ********************
    'FRHU 20140806 RQ14006: Observacion: Verifica que el saldo no quede en negativo cuando hay un monto a retirar, el itf del retiro se cobra internamente
    If Me.txtMontoRetirar.Visible = True Then
        If Not ValidarSaldoRetirar(sCtaAho, nMontoDesembolso, nITFDesembolso, txtMontoRetirar.Text) Then
           MsgBox "Importe a retirar es superior a lo aprobado, intente descontando Itf y Otras Comisiones de la operación de retiro", vbInformation, "Aviso"
           Exit Sub
        End If
    End If
    'FIN FRHU 20140806
   'EJVG20120322 Verifica actualización Persona
   Dim oPersona As New COMNPersona.NCOMPersona
   If oPersona.NecesitaActualizarDatos(LblCodCli.Caption, gdFecSis) Then
        MsgBox "Para continuar con la Operación Ud. debe actualizar los datos del" & Chr(13) & "Titular: " & LblNomCli.Caption, vbInformation, "Aviso"
        Dim foPersona As New frmPersona
        If Not foPersona.realizarMantenimiento(LblCodCli.Caption) Then
            MsgBox "No se ha realizado la actualización de los datos de " & LblNomCli.Caption & "," & Chr(13) & "la Operación no puede continuar!", vbInformation, "Aviso"
            Exit Sub
        End If
   End If
   'WIOR 20121009 Clientes Observados **************************************
        If sOperacion = gCredDesembEfec Or sOperacion = gCredDesembCtaNueva Then
            Dim oDPersona As COMDPersona.DCOMPersona
            Dim rsPersona As ADODB.Recordset
            Set oDPersona = New COMDPersona.DCOMPersona
            Set rsPersona = oDPersona.ObtenerUltimaVisita(Trim(LblCodCli.Caption))
         
            If rsPersona.RecordCount > 0 Then
                If Not (rsPersona.EOF And rsPersona.BOF) Then
                    If Trim(rsPersona!sUsual) = "3" Then
                        MsgBox "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                        Call frmPersona.Inicio(Trim(LblCodCli.Caption), PersonaActualiza)
                    End If
                End If
            End If
        End If
    'WIOR FIN ***************************************************************
    'WIOR 20150416 *****************************
    If ContMatCredCanc > 0 Then
        lsMensajeGrabar = ""
        For i = 1 To FECreditosVig.Rows - 1
            If Trim(FECreditosVig.TextMatrix(i, 1)) = "." Then
                If CInt(Trim(FECreditosVig.TextMatrix(i, 5))) = 0 Then
                    lsMensajeGrabar = lsMensajeGrabar & FECreditosVig.TextMatrix(i, 2) & ","
                End If
            End If
        Next i
        
        If lsMensajeGrabar <> "" Then
            lsMensajeGrabar = Mid(lsMensajeGrabar, 1, Len(lsMensajeGrabar) - 1)
            If Len(lsMensajeGrabar) > 18 Then
                lsMensajeGrabar = Mid(lsMensajeGrabar, 1, Len(lsMensajeGrabar) - 19) & " y " & Mid(lsMensajeGrabar, Len(lsMensajeGrabar) - 17, Len(lsMensajeGrabar)) & "."
                MsgBox "Favor de cancelar los Interes y Gastos a la Fecha de los Créditos (Preparar para la Ampliación): " & lsMensajeGrabar, vbInformation, "Aviso"
            Else
                MsgBox "Favor de cancelar los Interes y Gastos a la Fecha del Crédito (Preparar para la Ampliación): " & lsMensajeGrabar, vbInformation, "Aviso"
            End If
            Exit Sub
            
        End If
    End If
    'WIOR FIN **********************************
    
    'JOEP ERS066-20161110
        Dim oObtenerCampanaAct As COMDCredito.DCOMCredito
        Dim rsCampAct As ADODB.Recordset
        Dim rsAsigCampAct As ADODB.Recordset
        Set oObtenerCampanaAct = New COMDCredito.DCOMCredito
        Set rsCampAct = New ADODB.Recordset
        Set rsAsigCampAct = New ADODB.Recordset
        
        Set rsCampAct = oObtenerCampanaAct.ObtenerCampActiva(ActxCta.NroCuenta)
        
        If Not (rsCampAct.BOF And rsCampAct.EOF) Then
             Set rsAsigCampAct = oObtenerCampanaAct.ObtenerAsigCampActiva(rsCampAct!idCampana, gsCodAge)
             'Valida Asinacion de Campana/Agencia
                If Not (rsAsigCampAct.BOF And rsAsigCampAct.EOF) Then
                    
                Else
                    MsgBox "La Campaña no esta Asignada a la Agencia ,favor de Comunicarse con el Analista del Credito", vbInformation, "Aviso"
                Exit Sub
                End If
        Else
            MsgBox "El Credito esta  vinculado a una Campaña no Activa, favor de Comunicarse con el Analista del Credito", vbInformation, "Aviso"
            Exit Sub
        End If
    'JOEP ERS066-20161110

    
    Dim oDCreditos As COMDCredito.DCOMCreditos 'ARLO 20170904
    Set oDCreditos = New COMDCredito.DCOMCreditos 'ARLO 20170904
    
    If Not oDCreditos.VerificaClienteCampania(ActxCta.NroCuenta) Then 'ARLO 20170904 DESCOMENTADO  BY ARLO 20171118
    If TieneGarantiasPendienteMigracion(ActxCta.NroCuenta, True) Then Exit Sub 'EJVG20160322
    End If 'ARLO 20170904 DESCOMENTADO BY ARLO 20171118
    Set oDCreditos = Nothing 'ARLO 20170904
    
    'ALPA 20120509********************************************
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If Not (objProducto.GetResultadoCondicionCatalogo("O0000011", ActxCta.Prod)) Then
    'If Not (ActxCta.Prod = "515" Or ActxCta.Prod = "516") Then
    '**ARLO20180712 ERS042 - 2018
        If MsgBox("Se va a Desembolsar el Credito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Else
        If MsgBox("Se va a Desembolsar la Op.de Arrendamiento Financiero, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    End If
    '*********************************************************
   Set oNCredito = New COMNCredito.NCOMCredito
   
   'MAVM 20110301 ***
   Dim sMensPriFecPag As String
   Dim sMensFecAprob As String
   sMensFecAprob = oNCredito.DevolverFechaAprobacion(ActxCta.NroCuenta)
   'MADM 20110419 - PARAMETRO FEC APROB
   Set RDCredito = New ADODB.Recordset
   Set oDCredito = New COMDCredito.DCOMCredito
   Set RDCredito = oDCredito.RecuperaColocacEstado(ActxCta.NroCuenta, gColocEstAprob)
   If Not (RDCredito.EOF Or RDCredito.BOF) Then
        nPeriodoGraciaAnt = IIf(IsNull(RDCredito!nPeriodoGracia), 0, RDCredito!nPeriodoGracia)
        nTipoPeriodo = IIf(IsNull(RDCredito!nColocCalendCod), 0, RDCredito!nColocCalendCod)
   End If
   'MADM 20110531 - Array(2)
   'sMensPriFecPag = oNCredito.DevolverPrimeraFechaPago(ActxCta.NroCuenta, CDbl(LblMonPrestamo.Caption), gdFecSis, sMensFecAprob)
   lnDescripCuota = oNCredito.DevolverPrimeraFechaPago(ActxCta.NroCuenta, CDbl(LblMonPrestamo.Caption), gdFecSis, sMensFecAprob, "N")
   sMensPriFecPag = lnDescripCuota(2)
   
   '**ARLO20180712 ERS042 - 2018
   Set objProducto = New COMDCredito.DCOMCredito
   If Not (objProducto.GetResultadoCondicionCatalogo("O0000012", ActxCta.Prod)) Then
   'If Not (ActxCta.Prod = "515" Or ActxCta.Prod = "516") Then
   '**ARLO20180712 ERS042 - 2018
        lsMensajeGrabar = "El Crédito fue aprobado el "
   Else
        lsMensajeGrabar = "La operación fue aprobada el "
   End If
   'MADM 20110505
   If MsgBox(lsMensajeGrabar & sMensFecAprob & Chr(13) & "La Primera Fecha de Pago es: " & sMensPriFecPag & Chr(13) & Chr(13) & "            Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        oDCredito.ActualizarColocacEstadoDiasAtraso ActxCta.NroCuenta, gColocEstAprob, nPeriodoGraciaAnt
        Set oDCredito = Nothing
   Else
           MatGastos = CargaMatrizGastosDesemb
           
           'By Capi se comento por Riesgos
           
        '    If pbOperacionEfectivo Then
        '        If Not pbExoneradaLavado Then
                  sPersLavDinero = ""
        '            If CDbl(lblMonDesemb.Caption) >= Format(pnMontoLavDinero * pnTC, "#0.00") Then
        '                sPersLavDinero = IniciaLavDinero()
        '                sVisPersLavDinero = gVarPublicas.gVisPersLavDinero 'DAOR 20070512
        '                If sPersLavDinero = "" Then Exit Sub
        '            End If
        '        End If
        '    End If
           
           Dim psCtaAhoN  As String
           'ARCV 13-02-2007
           Dim MatDatosAhoNew As Variant
           ReDim MatDatosAhoNew(4)
           
           MatDatosAhoNew(0) = nProgAhorros
           MatDatosAhoNew(1) = nMontoAbonar
           MatDatosAhoNew(2) = nPlazoAbonar
           'MatDatosAhoNew(3) = sPromotorAho 'FRHU 20140228 RQ14006
           '--------------
           '*** BRGO 20111111 *****************
           Dim MatDatosLavDinero As Variant
           ReDim MatDatosLavDinero(8)
           MatDatosLavDinero(0) = LblNomCli.Caption
           MatDatosLavDinero(1) = LblDocNat.Caption
           MatDatosLavDinero(2) = LblCliDirec.Caption
           MatDatosLavDinero(3) = LblNomCli.Caption
           MatDatosLavDinero(4) = LblDocNat.Caption
           MatDatosLavDinero(5) = LblCliDirec.Caption
           MatDatosLavDinero(6) = LblNomCli.Caption
           MatDatosLavDinero(7) = LblDocNat.Caption
           '*** END BRGO ***********************
           
           
           'MADM 20110531 - MAX PARAMAETROS lnDescripCuota = LblCliDirec.Caption
           'BRGO 20111111 - Datos del REU se ingresan en una matriz
                'LblNomCli.Caption, LblDocNat.Caption, LblCliDirec.Caption, _
                'LblNomCli.Caption, LblDocNat.Caption, LblCliDirec.Caption, LblNomCli.Caption, LblDocNat.Caption, _
            'END BRGO
           
           Call oNCredito.CargarDesembolsoCredito(ActxCta.NroCuenta, sOperacion, gdFecSis, LblCodCli.Caption, vbDesembCC, vbCuentaNueva, _
                                                sCtaAho, CDbl(LblMonGastos.Caption), nITFDesembolso, CDbl(LblMonCancel.Caption), nMontoDesembolso, _
                                                CDbl(LblMonPrestamo.Caption), gsCodAge, gsCodUser, gsNomAge, gsNomCmac, gsInstCmac, gsCodCMAC, MatCargoAutom, _
                                                MatCredCanc, ContMatCredCanc, pRSRela, pnTasa, pnPersoneria, pnTipoCuenta, pnTipoTasa, pbDocumento, psNroDoc, psCodIF, _
                                                 sMensaje, sPersLavDinero, nMontoLavDinero, MatGastos, sError, sImpreDocs, sImpreBoletaAbono, sImpreBoletaGasto, _
                                                sImpreBoletaCancel, sImpreLavado, sLpt, MatDatosLavDinero, _
                                                lnDescripCuota, gsProyectoActual, vbDesembCheque, Right(CboBancos.Text, 13), Trim(TxtCheque.Text), Trim(CboCta.Text), sImpreDocsBoleta, psCtaAhoN, CDbl(lblPoliza.Caption), MatDatosAhoNew, sVisPersLavDinero, _
                                                vbDesembInfoGas, rsRelEmp, CDbl(txtMontoRetirar.Text), CStr(txtGlosaBloqueo.Text), CInt(lblampliado.Caption), MatCuentaGastoCierre) 'DAOR 20070511, sVisPersLavDinero
                                                'FRHU 20140228 RQ14006 - Se agrego CDbl(txtMontoRetirar.Text) Y CStr(txtGlosaBloqueo.Text)
                                                'BRGO 20111111 vbDesembInfoGas,rsRelEmp
          'CTI3 : lblampliado.Caption
                                                
              'vapi segun ERS082-2014 Nota: Abre el formulario de entrega de Merchandising
        If sOperacion = "100101" Or sOperacion = "100102" Then
            Call frmMkEntregaCombo.Inicio(ActxCta.NroCuenta, sOperacion, True, IIf(LblMoneda.Caption = "SOLES", 1, 2), Val(lblMonDesemb.Caption))
        End If
        '*****************************FIN VAPI***************************************
        'RECO20161020 ERS060-2016 **********************************************************
        Dim oNCOMColocEval As New NCOMColocEval
        Dim lcMovNro As String
        
        If Not ValidaExisteRegProceso(ActxCta.NroCuenta, gTpoRegCtrlDesembolso) Then
           lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
           Call oNCOMColocEval.insEstadosExpediente(ActxCta.NroCuenta, "Desembolso de Credito", lcMovNro, "", "", "", 1, 2020, gTpoRegCtrlDesembolso)
           Set oNCOMColocEval = Nothing
        End If
        'RECO FIN **************************************************************************
        'RECO20150601 ERS023-2015****************************************************
        If frGastoMYPE.Enabled = True Then
            Dim oGarant As New COMNCredito.NCOMGarantia
            Call oGarant.RegistrarNumPoliza(ActxCta.NroCuenta, txtMultMype.Text)
        End If
        'RECO FIN********************************************************************
        
            Set oNCredito = Nothing
            Dim clsprevio As previo.clsprevio
            '**ARLO20180712 ERS042 - 2018
            Set objProducto = New COMDCredito.DCOMCredito
            If Not (objProducto.GetResultadoCondicionCatalogo("O0000002", ActxCta.Prod)) Then
            'If Not (ActxCta.Prod = "515" Or ActxCta.Prod = "516") Then
            '**ARLO20180712 ERS042 - 2018
                If sMensaje <> "" Then
                    MsgBox sMensaje, vbInformation, "Mensaje"
                    Exit Sub
                End If
            
                'ARCV 03-11-2006
                If Trim(LblTipoAbono.Caption) = "NUEVA" Then
                   'GRABA DATOS ENVI0O ESTADO CTA
                   Call frmEnvioEstadoCta.GuardarRegistroEnvioEstadoCta(1, Trim(psCtaAhoN), LlenaRecordSet_EnvioEstCta, 1, 0, "") 'JUEZ 20130718
                   'IMPRIME REGISTRO DE FISMAS
                   Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
                   Dim lsCadImpFirmas As String
                   Dim lsCadImpCartilla As String
                   Dim sTipoCuenta As String
                   If pnTipoCuenta = 0 Then
                        sTipoCuenta = "INDIVIDUAL"
                   ElseIf pnTipoCuenta = 1 Then
                        sTipoCuenta = "MALCOMUNADA"
                   ElseIf pnTipoCuenta = 2 Then
                        'sTipoCuenta = "INDISTINTA"
                        sTipoCuenta = "SOLIDARIA" 'APRI20190109 ERS077-2018
                   End If
                   Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                        clsMant.IniciaImpresora gImpresora
                        lsCadImpFirmas = clsMant.GeneraRegistroFirmas(psCtaAhoN, sTipoCuenta, gdFecSis, False, rsRel, gsNomAge, gdFecSis, gsCodUser)
                   Set clsMant = Nothing
                   Set rsRel = Nothing
                   
                   'IMPRIME CARTILLA
                   'ARCV 12-02-2007
                   'lsCadImpCartilla = ImprimeCartillaLote(LblNomCli.Caption, 1, psCtaAhoN, pnTasa, LblMonPrestamo.Caption, gdFecSis) & oImpresora.gPrnSaltoPagina
                   'lsCadImpCartilla = lsCadImpCartilla & ImprimeCartillaLote(LblNomCli.Caption, 1, psCtaAhoN, pnTasa, LblMonPrestamo.Caption, gdFecSis)
                   'By Capi 07032008 convirtiendo pnTasa - TNA a TEA
                   Dim lnTasaE As Double
                   lnTasaE = Round(((1 + (pnTasa / 100 / 12) / 30) ^ 360 - 1) * 100, 2)
                   '
                    'By capi 09012009 se agrego para cuenta soñada
                    'If nProgAhorros = 0 Then
                    If nProgAhorros = 0 Or nProgAhorros = 5 Then
                    'End 09012009
                        'By Capi 07032008
                        'ImpreCartillaAhoCorriente MatTitulares, psCtaAhoN, pnTasa, CDbl(LblMonPrestamo.Caption)
                        ImpreCartillaAhoCorriente MatTitulares, psCtaAhoN, lnTasaE, CDbl(LblMonPrestamo.Caption), nProgAhorros
                       'INICIO EAAS20190523 Memorándum Nº 756-2019-GM-DI/CMACM
                        AhorroApertura_ContratosAutomaticos MatTitulares, psCtaAhoN
                        'FIN EAAS20190523 Memorándum Nº 756-2019-GM-DI/CMACM
                    ElseIf nProgAhorros = 3 Or nProgAhorros = 4 Then
                        'ImpreCartillaAhoPandero MatTitulares, psCtaAhoN, pnTasa, CDbl(LblMonPrestamo.Caption), gdFecSis, nMontoAbonar, nPlazoAbonar, nProgAhorros, ""
                        ImpreCartillaAhoPandero MatTitulares, psCtaAhoN, lnTasaE, CDbl(LblMonPrestamo.Caption), gdFecSis, nMontoAbonar, nPlazoAbonar, nProgAhorros, ""
                        'INICIO EAAS20190523 Memorándum Nº 756-2019-GM-DI/CMACM
                        AhorroApertura_ContratosAutomaticos MatTitulares, psCtaAhoN
                        'FIN EAAS20190523 Memorándum Nº 756-2019-GM-DI/CMACM
                    End If
                   '--------------
                End If
                '------------------
                
                Dim lsCadImpFirmasFMV As String
                lsCadImpFirmasFMV = ""
                If IsArray(MatCuentaGastoCierre) Then
                    GenerarDocumentosAperturaAhorro MatCuentaGastoCierre(12), MatCuentaGastoCierre(7), MatCuentaGastoCierre(11), MatCuentaGastoCierre(4), MatCuentaGastoCierre(3), MatCuentaGastoCierre(6), MatCuentaGastoCierre(0), MatCuentaGastoCierre(14), lsCadImpFirmasFMV
                End If
                 
                Do
                    MsgBox "Coloque Papel Continuo Tamaño Carta, Para la Impresion de los Documentos de Desembolsos", vbInformation, "Aviso"
                  
                    Set clsprevio = New previo.clsprevio
                    
                    '*** PEAC 20080723 ************************************
                     'clsPrevio.PrintSpool sLpt, oImpresora.gPrnCondensadaON & sImpreDocs & oImpresora.gPrnCondensadaOFF, False, gnLinPage
                     clsprevio.PrintSpool sLpt, oImpresora.gPrnTpoLetraSansSerif1PDef & oImpresora.gPrnTamLetra10CPIDef & sImpreDocs, False, gnLinPage
                    '******************************************************
                    
                    If Trim(LblTipoAbono.Caption) = "NUEVA" Then
                        clsprevio.PrintSpool sLpt, oImpresora.gPrnCondensadaON & lsCadImpFirmas & oImpresora.gPrnCondensadaOFF, False, gnLinPage   'ARCV 01-11-2006
                       
                       'IMPRIME CARTILLA
                       'ARCV 03-11-2006
                       'lsCadImp = lsCadImp & ImprimeCartillaLote(lblnomcli.Caption, 1, psCtaAhoN, pnTasa, LblMonPrestamo.Caption, gdFecSis) & oImpresora.gPrnSaltoPagina
                       'MsgBox "Cambie de Papel para imprimir Cartilla de Ahorros", vbExclamation, "Aviso"
                       '---------------
                       'clsPrevio.PrintSpool sLpt, lsCadImpCartilla, False, gnLinPage
                    End If
                    
                    If Len(Trim(lsCadImpFirmasFMV)) > 0 Then
                        clsprevio.PrintSpool sLpt, oImpresora.gPrnCondensadaON & lsCadImpFirmasFMV & oImpresora.gPrnCondensadaOFF, False, gnLinPage
                    End If
                    
                    'ARCV 07-06-2006
                    MsgBox "Cambie de Papel para imprimir las boletas de desembolsos", vbExclamation, "Aviso"
                    clsprevio.PrintSpool sLpt, Chr$(27) & Chr$(64) & sImpreDocsBoleta
                    '----------
                    If Not (sImpreBoletaAbono = "" And sImpreBoletaGasto = "" And sImpreBoletaCancel = "") Then
                        If sImpreBoletaAbono <> "" Then
                            clsprevio.PrintSpool sLpt, sImpreBoletaAbono
                        End If
                        If sImpreBoletaGasto <> "" Then
                            clsprevio.PrintSpool sLpt, sImpreBoletaGasto
                        End If
                        If sImpreBoletaCancel <> "" Then
                            clsprevio.PrintSpool sLpt, sImpreBoletaCancel
                        End If
                        Set clsprevio = Nothing
                    End If
                Loop While MsgBox("Desea Reimprimir Todos los Documentos del Desembolso?", vbInformation + vbYesNo, "Aviso") = vbYes
                
                
                If sPersLavDinero <> "" Then
                    Set clsprevio = New previo.clsprevio
                    Do
                        clsprevio.PrintSpool sLpt, sImpreLavado
                    Loop While MsgBox("Desea Reimprimir Boleta de lavado de dinero?", vbInformation + vbYesNo, "Aviso") = vbYes
                    Set clsprevio = Nothing
                End If
                'ALPA 20111213
                '**ARLO20180712 ERS042 - 2018
                Set objProducto = New COMDCredito.DCOMCredito
                If Not (objProducto.GetResultadoCondicionCatalogo("O0000013", Mid(ActxCta.NroCuenta, 6, 3))) Then
                'If Not (Mid(ActxCta.NroCuenta, 6, 3) = "515" Or Mid(ActxCta.NroCuenta, 6, 3) = "516") Then
                '**ARLO20180712 ERS042 - 2018
                    MsgBox "Coloque papel para Imprimir Hoja de Resumen...", vbInformation, "Aviso"
                    Call ImprimeCartillaCred(ActxCta.NroCuenta)
                    'INICIO EAAS20190523 Memorándum Nº 756-2019-GM-DI/CMACM
                     If (sTpoProdCod = "806" Or sTpoProdCod = "802") Then 'JATO 20210421
                    Call ContratoHipotecario(ActxCta.NroCuenta, sTpoProdCod) 'JATO 20210421
                    Else 'JATO 20210421
                    Call ContratosAutomaticosDes(ActxCta.NroCuenta)
                    End If 'JATO 20210421
                    'FIN EAAS20190523 Memorándum Nº 756-2019-GM-DI/CMACM
                End If
                '*** PEAC 20170621
                MsgBox "Se generará el Pagaré en formato PDF, por favor presione [Aceptar] para continuar.", vbInformation, "Aviso"
                Call ImpresionDePagare
                '*** FIN PEAC
                If vbDesembInfoGas Then
                    MsgBox "Coloque papel para Imprimir Cartas de Autorización del Producto Ecotaxi...", vbInformation, "Aviso"
                    Call ImprimeCartasEcotaxi(ActxCta.NroCuenta)
                End If
                
                Call ImprimeActivatePeru(ActxCta.NroCuenta) ''ANGC2020 impresion campaña reactiva
                
                gVarPublicas.LimpiaVarLavDinero 'DAOR 20070512
        Else
            'ALPA20130924
            'Leasing
               
                Do
                    Set clsprevio = New previo.clsprevio
                    MsgBox "Cambie de Papel para imprimir las boletas de desembolsos", vbExclamation, "Aviso"
                    clsprevio.PrintSpool sLpt, Chr$(27) & Chr$(64) & sImpreDocsBoleta
                    '----------
                    If Not (sImpreBoletaAbono = "" And sImpreBoletaGasto = "" And sImpreBoletaCancel = "") Then
                        If sImpreBoletaAbono <> "" Then
                            clsprevio.PrintSpool sLpt, sImpreBoletaAbono
                        End If
                        If sImpreBoletaGasto <> "" Then
                            clsprevio.PrintSpool sLpt, sImpreBoletaGasto
                        End If
                        If sImpreBoletaCancel <> "" Then
                            clsprevio.PrintSpool sLpt, sImpreBoletaCancel
                        End If
                        Set clsprevio = Nothing
                    End If
                Loop While MsgBox("Desea Reimprimir Todas las boletas del Desembolso?", vbInformation + vbYesNo, "Aviso") = vbYes
            'Fin leasing
        End If
        
        'WIOR 20120920 *****************************************************
        Dim oDCredito2 As COMDCredito.DCOMCredito
        Dim rsDCredito As ADODB.Recordset
        Dim sPrd As String
        Dim sSPrd As String
        Dim sCred As String
        Dim sSCred As String
        Dim psPerCod As String
        Dim nMontoFinal As Double
        Dim nMontoEndeud As Double
        Dim nEspecializacion As Integer
        Set oDCredito2 = New COMDCredito.DCOMCredito
        Set rsDCredito = oDCredito2.RecuperaSolicitudDatoBasicos(Trim(ActxCta.NroCuenta))

        sSPrd = Trim(rsDCredito!cTpoProdCod)
        sPrd = Mid(sSPrd, 1, 1) & "00"
        sSCred = Trim(rsDCredito!cTpoCredCod)
        sCred = Mid(sSCred, 1, 2) & "0"
        psPerCod = Trim(rsDCredito!cperscod)
        nMontoEndeud = oDCredito2.RecuperaEndeudamientoPersonal(psPerCod, gdFecSis)
        nMontoFinal = nMontoPrestamoW + nMontoEndeud
        nEspecializacion = oDCredito2.AsignarEspecializacionCred(sPrd, sSPrd, sCred, sSCred, nMontoFinal)
        
        Call oDCredito2.InsertaEspecializacionCred(Trim(ActxCta.NroCuenta), nEspecializacion)
        'WIOR FIN **********************************************************
        'APRI20171009 ERS028-2017
          If sOperacion = "100102" Or sOperacion = "100103" Or sOperacion = "100104" Or sOperacion = "100105" Then
              frmSegSepelioAfiliacion.Inicio IIf(sCtaAho = "", psCtaAhoN, sCtaAho)
          End If
          
         'Call oDCredito2.InsertaSegMultiriesgoDesgravamenTrama(Trim(ActxCta.NroCuenta)) 'COMENTADO POR APRI20181121 ERS071-2018 - MEJORA
        'END IF
        'INICIO JHCU ENCUESTA 16-10-2019
         Encuestas gsCodUser, gsCodAge, "ERS0292019", sOperacion
        'FIN
            
            Call cmdCancelar_Click
    End If
End Sub
'*** PEAC 20170621
Private Sub ImpresionDePagare()
    Dim bValor As Boolean
    Dim nTipoFormato As Integer
    Dim sFecDes As String

    If vbYes = MsgBox("Formato Nuevo [SI]  / Formato Antiguo [NO]", vbInformation + vbYesNo) Then
        nTipoFormato = 0
    Else
        nTipoFormato = 1
    End If
    
    bValor = VerificarExisteDesembolsoBcoNac(ActxCta.NroCuenta, sFecDes, 2)
    If bValor = True Then
        Call ImprimePagareCredPDF(ActxCta.NroCuenta, nTipoFormato, sFecDes)
    Else
        Call ImprimePagareCredPDF(ActxCta.NroCuenta, nTipoFormato)
    End If

End Sub

Private Function IniciaLavDinero() As String
 
Dim sPersCod As String
Dim sNombre As String
Dim sDireccion As String
Dim sDocId As String
Dim nMonto As Double

sPersCod = LblCodCli.Caption
sNombre = LblNomCli.Caption
sDireccion = LblCliDirec.Caption
sDocId = LblDocNat.Caption
 
nMonto = CDbl(lblMonDesemb.Caption)
IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, ActxCta.NroCuenta, sOperacion, False, "COLOCACIONES")
End Function

Private Sub CmdDesembolso_Click()
    Call CmdDesemb_Click
End Sub

Private Sub cmdExaminar_Click()
    'MARG20171205*********************
    Dim dFechaSis As Date
    dFechaSis = CDate(validarFechaSistema)

    If gdFecSis <> dFechaSis Then
        MsgBox "La Fecha de tu sesión en el Negocio no coincide con la fecha del Sistema", vbCritical, "Aviso"
        Call SalirSICMACMNegocio
        Unload Me
        End
    End If
    'END MARG**************************
    
    'ALPA 20111213*****************
    'ActxCta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstAprob), "Creditos para Desembolsar", , , , gsCodAge)
    ActxCta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstAprob), "Creditos para Desembolsar", , , , gsCodAge, bLeasing)
    '******************************
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
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdSeleccionar_Click()
    'By Capi 11012007 para que valide que no sea una cuenta exonerada de ITF
    If Trim(FECtaAhoDesemb.TextMatrix(FECtaAhoDesemb.row, 3)) = "Exonerada" Then
        MsgBox "No Puede Elegir Una Cuenta Exonerada ITF...", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If Trim(FECtaAhoDesemb.TextMatrix(FECtaAhoDesemb.row, 1)) <> "" Then
        'JUEZ 20141114 *********************************************
        Dim clsDef As New COMNCaptaGenerales.NCOMCaptaDefinicion
        'If Not clsDef.GetCapParametroNew(gCapAhorros, FECtaAhoDesemb.TextMatrix(FECtaAhoDesemb.row, 4))!bAplicaDesembCred Then
        If Not clsDef.GetCapParametroNew(gCapAhorros, FECtaAhoDesemb.TextMatrix(FECtaAhoDesemb.row, 4), sCtaAho)!bAplicaDesembCred Then 'APRI20190109 ERS077-2018
            MsgBox "El producto de la cuenta seleccionada no está configurado para destino de desembolso. Favor de comunicarse con el Departamento de Ahorros", vbInformation, "Aviso"
            Exit Sub
        End If
        'END JUEZ **************************************************
        sCtaAho = FECtaAhoDesemb.TextMatrix(FECtaAhoDesemb.row, 1)
        pnFilaSelecCtaAho = FECtaAhoDesemb.row
        CmdSeleccionar.Visible = False
        CmdDeseleccionar.Visible = True
        Call FECtaAhoDesemb.BackColorRow(vbYellow)
        LblCtaAbo.Caption = sCtaAho
        LblTipoCta.Caption = "Propia"
        LblAgeAbono.Caption = FECtaAhoDesemb.TextMatrix(FECtaAhoDesemb.row, 2)
        
        Me.lblCtaAboRecaudo.Caption = sCtaAho
        Me.lblTipoCtaRecaudo.Caption = "Propia"
        Me.lblAgenciaCtaRecaudo.Caption = FECtaAhoDesemb.TextMatrix(FECtaAhoDesemb.row, 2)
        
        LblTipoAbono.Caption = "EXISTENTE"
        vbCuentaNueva = False
    End If
End Sub

Private Sub CmdVisualizar_Click()
    'FrmVisualizacionDesembolsos.Inicio (ActxCta.NroCuenta)
End Sub

Private Sub FECargoAutom_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If Trim(FECargoAutom.TextMatrix(pnRow, pnCol)) <> "." Then 'Sin Check
        Call EliminarCtaAhoCargo(FECargoAutom.TextMatrix(pnRow, 2))
    Else 'Con Check
        Call AdicionaCtaAhoCargo(FECargoAutom.TextMatrix(pnRow, 2))
    End If
End Sub

Private Sub FECreditosVig_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Dim i As Integer
Dim nRedondITFCanc As Double
Dim nITFCtaCanc As Double

    If bOnCellCheck Then 'FRHU 20140424 TI-ERS015-2014 Se agrego If bOnCellCheck Then
    
        'JUEZ 20130930 Para cobrar ITF en la Cancelacion sólo cuando es desembolso efectivo ****
        If Not vbDesembCC Then
            '*** BRGO 20111012 ************************************************
            nITFCtaCanc = fgITFDesembolso(CDbl(FECreditosVig.TextMatrix(pnRow, 4)))
            nRedondITFCanc = fgDiferenciaRedondeoITF(nITFCtaCanc)
            If nRedondITFCanc > 0 Then
                nITFCtaCanc = Format(nITFCtaCanc - nRedondITFCanc, "#,##0.00")
            End If
            '*** END BRGO
        End If
        'END JUEZ *****************************************************************
        
        'CTI3 -ferimoro ERS082-2018
        'Cuando se Abona a cuenta se genera un ITF cargo a Cuenta
        If vbDesembCC Then
            '*** BRGO 20111012 ************************************************
            nITFCtaCanc = fgITFDesembolso(CDbl(FECreditosVig.TextMatrix(pnRow, 4)))
            nRedondITFCanc = fgDiferenciaRedondeoITF(nITFCtaCanc)
            If nRedondITFCanc > 0 Then
                nITFCtaCanc = Format(nITFCtaCanc - nRedondITFCanc, "#,##0.00")
            End If
            '*** END BRGO
        End If
        '--------------------------
        
        If bInstFinanc Then nITFCtaCanc = 0 'JUEZ 20140411
        
        If Trim(FECreditosVig.TextMatrix(pnRow, pnCol)) <> "." Then 'Sin Check
            Call EliminarCreditoACancelar(FECreditosVig.TextMatrix(pnRow, 2))
            'ARCV 04-11-2006
            'LblItf.Caption = Format(CDbl(LblItf.Caption) - fgITFDesembolso(CDbl(FECreditosVig.TextMatrix(pnRow, 4))), "#0.00")
            '----------
            LblItf.Caption = Format(CDbl(LblItf.Caption) - nITFCtaCanc, "#0.00") 'BRGO 20111012
            
        Else 'Con Check
            If CDbl(FECreditosVig.TextMatrix(pnRow, 4)) > (CDbl(LblMonPrestamo.Caption) - CDbl(LblMonGastos.Caption) - CDbl(LblMonCancel.Caption)) Then
                MsgBox "Monto a Desembolsar no es suficiente para cancelar este Credito", vbInformation, "Aviso"
                FECreditosVig.TextMatrix(pnRow, 1) = ""
                Exit Sub
            End If
            Call AdicionaCreditoACancelar(FECreditosVig.TextMatrix(pnRow, 2), CDbl(FECreditosVig.TextMatrix(pnRow, 4)))
            'LblItf.Caption = Format(CDbl(LblItf.Caption) + fgITFDesembolso(CDbl(FECreditosVig.TextMatrix(pnRow, 4))), "#0.00")
            LblItf.Caption = Format(CDbl(LblItf.Caption) + nITFCtaCanc, "#0.00") 'BRGO 20111012
        End If
        nMontoCredCanc = 0
        LstCredVig.Clear
        For i = 0 To ContMatCredCanc - 1
            nMontoCredCanc = nMontoCredCanc + CDbl(MatCredCanc(i, 1))
            LstCredVig.AddItem MatCredCanc(i, 0)
        Next i
        'VAPA 20171207 FROM 60
        'For i = 0 To ContMatCredCanc - 1
           'lsCuentaCredCan = lsCuentaCredCan & "," & MatCredCanc(i, 0)
        'Next
        'lsCuentaCredCan = Right(lsCuentaCredCan, Len(lsCuentaCredCan) - 1)
        'VAPA 20171207 END 'Comentado by NAGL 20180609
        nMontoCredCanc = CDbl(Format(nMontoCredCanc, "#0.00"))
        LblMonCancel.Caption = Format(nMontoCredCanc, "#0.00")
        LblTotCred.Caption = Format(nMontoCredCanc, "#0.00")
        'Se Modifico
        lblMonDesemb.Caption = Format(TotalADesembolsar(CDbl(LblItf.Caption)), "#0.00")
        
    'FRHU 20140424 TI-ERS015-2014
    Else
        If Trim(FECreditosVig.TextMatrix(pnRow, pnCol)) = "." Then
            FECreditosVig.TextMatrix(pnRow, 1) = ""
        Else
            FECreditosVig.TextMatrix(pnRow, 1) = "1"
        End If
    End If
    bOnCellCheck = False
    'FIN FRHU 20140424 TI-ERS015-2014
End Sub

Private Sub FECtaAhoDesemb_Click()
    If nContPersCtaAho > 0 Then
        Call CargaClientesCtaAho(FECtaAhoDesemb.TextMatrix(FECtaAhoDesemb.row, 1))
    End If
End Sub

Private Sub FECtaAhoDesemb_OnRowChange(pnRow As Long, pnCol As Long)
    If nContPersCtaAho > 0 Then
        Call CargaClientesCtaAho(FECtaAhoDesemb.TextMatrix(pnRow, 1))
    End If
End Sub
'ALPA 20111213******************
Private Sub FEProveedores_RowColChange()
Dim oDLeasing As COMDCredito.DCOMleasing
Dim oRs As ADODB.Recordset
Set oRs = New ADODB.Recordset
Set oDLeasing = New COMDCredito.DCOMleasing
FEProveedores.lbEditarFlex = True
FEProveedores.SetFocus
Set oRs = oDLeasing.RecuperaCtasAhorroxPersonaLeasing(FEProveedores.TextMatrix(FEProveedores.row, 1), Mid(ActxCta.Cuenta, 1, 1))
FEProveedores.CargaCombo oRs
End Sub
'*******************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColPYMEEmp, False)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    'LimpiaPantalla ' VAPA 20171211 FROM 60 Comentado by NAGL 20180609
    CentraForm Me
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    nMontoGastos = 0
    nMontoGastoCierre = 0
    nMontoCredCanc = 0
    ReDim MatCargoAutom(0)
    sCtaAho = ""
    nNroProxDesemb = 0
    ReDim RPersCtaAho(0)
    nContPersCtaAho = 0
    ContMatCredCanc = 0
    ContMatCargoAutom = 0
    pnFilaSelecCtaAho = -1
    SSTabDatos.TabVisible(7) = False 'ALPA 20111213
    fgITFParametros
    CargarSubProducto 'FRHU 20140228 RQ14006
End Sub
'*** BRGO 20111109 *************************************************
Private Sub cmdCancela_Click()
    Call cmdCancelar_Click
End Sub
Private Sub cmdVisualiza_Click()
    'FrmVisualizacionDesembolsos.Inicio (ActxCta.NroCuenta)
End Sub
Private Sub cmdSale_Click()
    Unload Me
End Sub
'*** END BRGO ******************************************************

'JUEZ 20130718 *****************************************************
Private Function LlenaRecordSet_EnvioEstCta() As ADODB.Recordset
Dim rsEnvioEstCta As ADODB.Recordset
Dim i As Integer
Set rsEnvioEstCta = New ADODB.Recordset

With rsEnvioEstCta
    .Fields.Append "codigo", adVarChar, 13
    .Fields.Append "Envio", adSmallInt
    .Fields.Append "Cuenta", adVarChar, 18
    .Fields.Append "Domicilio", adVarChar, 200
    .Open
    .AddNew
    .Fields("codigo") = LblCodCli.Caption
    .Fields("Envio") = 0
    .Fields("Cuenta") = ""
    .Fields("Domicilio") = LblCliDirec.Caption
End With
Set LlenaRecordSet_EnvioEstCta = rsEnvioEstCta
End Function
'END JUEZ **********************************************************
'FRHU 20140228 RQ14006
Private Sub CargarSubProducto()
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsConst As New ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(2030, "1,2,3,4,6,7,8", , "1")
    
    Set clsGen = Nothing
    Do While Not rsConst.EOF
            cboPrograma.AddItem rsConst("cDescripcion") & space(100) & rsConst("nConsValor")
            rsConst.MoveNext
    Loop
    cboPrograma.ListIndex = 0
End Sub
Private Sub txtMontoRetirar_KeyPress(KeyAscii As Integer)
Dim nRedondITF As Double
Dim nITFCta As Double
Dim nITFValor As Double

    KeyAscii = NumerosDecimales(txtMontoRetirar, KeyAscii, 15, 2)
    If KeyAscii = 13 Then
        If Me.txtMontoRetirar.Text = "" Then
            'Me.txtMontoRetirar.Text = Format(0, "#,##0.00")
            'nITFCta = fgITFDesembolso(LblMonPrestamo.Caption)
            'nRedondITF = fgDiferenciaRedondeoITF(nITFCta)
            'If nRedondITF > 0 Then
            '    nITFCta = Format(nITFCta - nRedondITF, "#,##0.00")
            'End If
            'lblITF.Caption = Format(nITFCta, "#0.00")
            Me.txtMontoRetirar.Text = Format(0, "#0.00")
            Exit Sub
        End If
        'nITFCta = fgITFDesembolso(Me.txtMontoRetirar.Text)
        'nRedondITF = fgDiferenciaRedondeoITF(nITFCta)
        'If nRedondITF > 0 Then
        '    nITFCta = Format(nITFCta - nRedondITF, "#,##0.00")
        'End If
        'lblITF.Caption = Format(CDbl(lblITF.Caption) + nITFCta, "#0.00")
        Me.txtMontoRetirar.Text = Format(Me.txtMontoRetirar.Text, "#,##0.00")
    End If
    
End Sub
Private Sub txtMontoRetirar_LostFocus()
    Me.txtMontoRetirar.Text = Format(Me.txtMontoRetirar.Text, "#,##0.00")
End Sub
'FIN FRHU 20140228
'FRHU 20140806 RQ14006: Observacion
Private Function ValidarSaldoRetirar(ByVal psCuenta As String, ByVal pnMontoDesembolso As Double, ByVal pnItfDesembolso As Double, ByVal pnMontoRetirar As Double) As Boolean
    Dim nITFRetiro As Double, nRedondeoITFRetiro As Double
    Dim oITF As COMDConstSistema.FCOMITF
    Dim nSaldoTemp As Double, nMontoRetiroTemp As Double
    
    Set oITF = New COMDConstSistema.FCOMITF
    oITF.fgITFParametros
    nITFRetiro = oITF.fgTruncar(oITF.fgITFCalculaImpuesto(pnMontoRetirar), 2)
    nRedondeoITFRetiro = fgDiferenciaRedondeoITF(nITFRetiro)
    nITFRetiro = IIf(nRedondeoITFRetiro > 0, nITFRetiro - nRedondeoITFRetiro, nITFRetiro)
    
    nSaldoTemp = pnMontoDesembolso
    nMontoRetiroTemp = pnMontoRetirar + nITFRetiro
    
    If vbCuentaNueva Then
        If (nSaldoTemp - nMontoRetiroTemp) >= 0 Then
            ValidarSaldoRetirar = True
        Else
            ValidarSaldoRetirar = False
        End If
    Else
        Dim clsMant As NCOMCaptaGenerales
        Dim rsCta As New ADODB.Recordset
        Dim nSaldo As Double
        
        Set clsMant = New NCOMCaptaGenerales
        Set rsCta = clsMant.GetDatosCuenta(psCuenta)
        nSaldo = rsCta("nSaldoDisp")
        
        rsCta.Close
        Set rsCta = Nothing
        Set clsMant = Nothing
        
        If (nSaldo + nSaldoTemp - nMontoRetiroTemp) >= 0 Then
            ValidarSaldoRetirar = True
        Else
            ValidarSaldoRetirar = False
        End If
    End If
End Function
'FIN FRHU 20140806
'VAPA 20171209 FROM 60
'Public Function ActualizaCreditoAmpliadoNew(ByVal psCuenta As String, psCuentaNew) As ADODB.Recordset
'On Error GoTo ActualizaCreditoAmpliadoNewErr
'   Dim oRs As ADODB.Recordset
'   Dim oConec As DConecta
'   Dim psSQL As String
'   Set oRs = New ADODB.Recordset
'   Set oConec = New DConecta
'   oConec.AbreConexion
'   psSQL = "exec stp_upd_CreAmpliado '" & psCuenta & "','" & psCuentaNew & "'"
'   Set oRs = oConec.CargaRecordSet(psSQL)
'    Set ActualizaCreditoAmpliadoNew = oRs
'   oConec.CierraConexion
'Exit Function
'ActualizaCreditoAmpliadoNewErr:
'   'Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceDiario Method")
'   err.Raise err.Number, "Actualizacion de Credito Nuevo ampliado no correcta ", err.Description
'End Function
'Private Sub LimpiMatriz()
'Dim i As Integer, contMatCan As Integer 'NAGL 20180113 FROM 60
'Dim MatCuentas(100, 2) As String 'NAGL 20180113 FROM 60
'For i = 0 To contMatCan - 1
'           MatCuentas(contMatCan - 1, 0) = ""
'           MatCuentas(contMatCan - 1, 1) = ""
'Next
'    contMatCan = 0
'End Sub
'VAPA END


'MARG20171205**********************
Sub SalirSICMACMNegocio()
    Dim oSeguridad As New COMManejador.Pista
    Call oSeguridad.InsertarPista(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, "Salida del SICMACM Operaciones" & " Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-20160312")
     If oSeguridad.ValidaAccesoPistaRF(gsCodUser) Then
            Call oSeguridad.InsertarPistaSesion(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, 0)
            Call oSeguridad.ActualizarPistaSesion(gsCodPersUser, GetMaquinaUsuario, 0) 'JUEZ 20160125
     End If
    Set oSeguridad = Nothing
End Sub
'END MARG**************************

'CTI520210408 ***
Private Function ObtenerMatrizGastoCierre() As Variant
    Dim nTasaNominal As Double
    Dim nPersoneria As Integer
    Dim nTasa As Double
    Dim nTipoCuenta As Integer
    Dim nTipoTasa As Integer
    Dim bDocumento As Boolean
    Dim sNroDoc As String
    Dim sCodIF As String
    Dim nProgAhorros As Integer
    Dim nPlazoAbono As Integer
    
    nPersoneria = PersPersoneria.gPersonaNat
    nTasa = 0
    nTipoCuenta = 0
    nTipoTasa = 0
    sNroDoc = ""
    sCodIF = ""
    nProgAhorros = CaptacSubProdAhorros.gCapSubProdAhoCorriente
    nPlazoAbono = 0

    frmCapAperturas.IniciaDesembAbonoCta gCapAhorros, gAhoApeEfec, "", CInt(Mid(Me.ActxCta.NroCuenta, 9, 1)), _
         LblCodCli.Caption, LblNomCli.Caption, CDbl(lblMonGastoCierre.Caption), frsRelaFMV, _
        nTasa, nPersoneria, nTipoCuenta, nTipoTasa, bDocumento, sNroDoc, sCodIF, False, _
        , , fMatTitularesFMV, nProgAhorros, , nPlazoAbono

    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion

    nProgAhorros = CaptacSubProdAhorros.gCapSubProdAhoCorriente
    nTasa = clsDef.GetCapTasaInteres(gCapAhorros, CInt(Mid(Me.ActxCta.NroCuenta, 9, 1)), nTipoTasa, , CDbl(lblMonGastoCierre.Caption), gsCodAge, , nProgAhorros)
    
    Dim MatApeGastoCierre As Variant
    ReDim MatApeGastoCierre(14)
    
    MatApeGastoCierre(0) = LblCodCli.Caption 'Cod Titular
    MatApeGastoCierre(1) = LblNomCli.Caption 'Nombre Titular
    MatApeGastoCierre(2) = nPersoneria 'Personería
    MatApeGastoCierre(3) = nProgAhorros 'Programa Ahorro Corriente
    MatApeGastoCierre(4) = nTipoCuenta 'Tipo de Cuenta
    MatApeGastoCierre(5) = nTipoTasa 'Tipo de Tasa
    MatApeGastoCierre(6) = nTasa 'Tasa
    MatApeGastoCierre(7) = CDbl(lblMonGastoCierre.Caption) 'Monto de apertura
    MatApeGastoCierre(8) = bDocumento 'Documento
    MatApeGastoCierre(9) = sNroDoc 'Nro de documento
    MatApeGastoCierre(10) = sCodIF 'Codigo IFI
    MatApeGastoCierre(11) = nPlazoAbono 'Plazo Abono
    MatApeGastoCierre(12) = "" 'Nro Cuenta a generar
    MatApeGastoCierre(13) = CDbl(0) 'ITF Abono
    MatApeGastoCierre(14) = LblCliDirec.Caption 'Direccion Titular
    
    ObtenerMatrizGastoCierre = MatApeGastoCierre
End Function
'END CTI5 *******
'CTI5 20210516***
Private Sub GenerarDocumentosAperturaAhorro(ByVal psCtaCod As String, ByVal pnMonto As Double, ByVal pnPlazoAbonar As Integer, _
                                            ByVal pnTipoCuenta As Integer, ByVal pnProgAhorros As Integer, ByVal pnTasa As Double, _
                                            ByVal psPersCod As String, ByVal psDireccion As String, _
                                            ByRef psCadImpFirma As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim lsCadImpFirmas As String
    Dim lsCadImpCartilla As String
    Dim sTipoCuenta As String
    
    If pnTipoCuenta = 0 Then
        sTipoCuenta = "INDIVIDUAL"
    ElseIf pnTipoCuenta = 1 Then
        sTipoCuenta = "MALCOMUNADA"
    ElseIf pnTipoCuenta = 2 Then
        sTipoCuenta = "SOLIDARIA"
    End If
    
    Dim rsEnvioEstCta As ADODB.Recordset
    Set rsEnvioEstCta = New ADODB.Recordset
    
    With rsEnvioEstCta
        .Fields.Append "codigo", adVarChar, 13
        .Fields.Append "Envio", adSmallInt
        .Fields.Append "Cuenta", adVarChar, 18
        .Fields.Append "Domicilio", adVarChar, 200
        .Open
        .AddNew
        .Fields("codigo") = psPersCod
        .Fields("Envio") = 0
        .Fields("Cuenta") = ""
        .Fields("Domicilio") = psDireccion
    End With
    
    Call frmEnvioEstadoCta.GuardarRegistroEnvioEstadoCta(1, psCtaCod, rsEnvioEstCta, 1, 0, "")
    
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    clsMant.IniciaImpresora gImpresora
    'Faltaría que el recordset "frsRelaFMV" y "fMatTitularesFMV" sea dinámico para que este método sea genérico
    psCadImpFirma = clsMant.GeneraRegistroFirmas(psCtaCod, sTipoCuenta, gdFecSis, False, frsRelaFMV, gsNomAge, gdFecSis, gsCodUser)
    Set clsMant = Nothing
    
    Dim lnTasaE As Double
    lnTasaE = Round(((1 + (pnTasa / 100 / 12) / 30) ^ 360 - 1) * 100, 2)
    
    If pnProgAhorros = 0 Or pnProgAhorros = 5 Then
        ImpreCartillaAhoCorriente fMatTitularesFMV, psCtaCod, lnTasaE, pnMonto, pnProgAhorros
        AhorroApertura_ContratosAutomaticos fMatTitularesFMV, psCtaCod
    ElseIf pnProgAhorros = 3 Or pnProgAhorros = 4 Then
        ImpreCartillaAhoPandero fMatTitularesFMV, psCtaCod, lnTasaE, pnMonto, gdFecSis, pnMonto, pnPlazoAbonar, pnProgAhorros, ""
        AhorroApertura_ContratosAutomaticos fMatTitularesFMV, psCtaCod
    End If
End Sub
'END CTI5 *******
