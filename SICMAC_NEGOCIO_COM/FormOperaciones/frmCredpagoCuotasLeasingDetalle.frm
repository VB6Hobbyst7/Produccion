VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredpagoCuotasLeasingDetalle 
   Caption         =   "Pago Cuotas Arrendamiento Financiero"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   Icon            =   "frmCredpagoCuotasLeasingDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   8265
   Begin VB.CommandButton CmdGastos 
      Caption         =   "&Gastos"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5520
      TabIndex        =   16
      Top             =   6480
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton CmdPlanPagos 
      Caption         =   "&Plan Pagos"
      Enabled         =   0   'False
      Height          =   345
      Left            =   2280
      TabIndex        =   15
      Top             =   6480
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   1200
      TabIndex        =   14
      Top             =   6480
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3360
      TabIndex        =   13
      Top             =   6480
      Width           =   1050
   End
   Begin VB.CommandButton cmdmora 
      Caption         =   "&Mora"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4440
      TabIndex        =   12
      Top             =   6480
      Width           =   1050
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   6600
      TabIndex        =   11
      Top             =   6480
      Width           =   1050
   End
   Begin TabDlg.SSTab stPagoLeasing 
      Height          =   4455
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Pago"
      TabPicture(0)   =   "frmCredpagoCuotasLeasingDetalle.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Conceptos Pago"
      TabPicture(1)   =   "frmCredpagoCuotasLeasingDetalle.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label17"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "LblNumDoc"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label26"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "LblItf"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblPagoTotal"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label28"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "FrmDetallePago"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "CmbForPag"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "TxtMonPag"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "ckOC"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      Begin VB.CheckBox ckOC 
         Caption         =   "Cancelar?"
         Enabled         =   0   'False
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   6480
         TabIndex        =   69
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox TxtMonPag 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   56
         Top             =   3840
         Width           =   1380
      End
      Begin VB.ComboBox CmbForPag 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   3480
         Width           =   1785
      End
      Begin VB.Frame FrmDetallePago 
         Caption         =   "Detalle Pago"
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
         Height          =   3015
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   7455
         Begin SICMACT.FlexEdit FEConceptosLeasing 
            Height          =   2655
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4683
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "-nConceptoCod-Concepto-Monto Pago-nNroCalend-nCuota--FechaVenc"
            EncabezadosAnchos=   "400-0-4800-1200-0-0-500-0"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-R-C-C-C-C"
            FormatosEdit    =   "0-0-0-4-0-0-0-5"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos de Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3705
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   7365
         Begin VB.Frame Frame2 
            Height          =   2175
            Left            =   240
            TabIndex        =   18
            Top             =   195
            Width           =   7410
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cuota Pendiente"
               Height          =   210
               Left            =   2835
               TabIndex        =   49
               Top             =   1740
               Width           =   1185
            End
            Begin VB.Label LblCPend 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000004&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   4290
               TabIndex        =   48
               Top             =   1720
               Width           =   495
            End
            Begin VB.Label LblMontoCuota 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000004&
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
               Height          =   270
               Left            =   4290
               TabIndex        =   37
               Top             =   1410
               Width           =   840
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Monto Cuota :"
               Height          =   195
               Left            =   2835
               TabIndex        =   36
               Top             =   1410
               Width           =   1005
            End
            Begin VB.Label LblTotCuo 
               AutoSize        =   -1  'True
               Caption         =   "Cuotas"
               Height          =   195
               Left            =   1650
               TabIndex        =   35
               Top             =   1410
               Width           =   495
            End
            Begin VB.Label LblForma 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000004&
               BorderStyle     =   1  'Fixed Single
               Height          =   270
               Left            =   1110
               TabIndex        =   34
               Top             =   1365
               Width           =   480
            End
            Begin VB.Label Lbl2 
               AutoSize        =   -1  'True
               Caption         =   "Forma Pago"
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   1395
               Width           =   870
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Deuda a la Fecha : "
               Height          =   195
               Left            =   2835
               TabIndex        =   32
               Top             =   1095
               Width           =   1410
            End
            Begin VB.Label LblTotDeuda 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000004&
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
               Height          =   270
               Left            =   4290
               TabIndex        =   31
               Top             =   1080
               Width           =   1335
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Left            =   105
               TabIndex        =   30
               Top             =   780
               Width           =   585
            End
            Begin VB.Label LblMoneda 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000004&
               BorderStyle     =   1  'Fixed Single
               Height          =   270
               Left            =   1110
               TabIndex        =   29
               Top             =   750
               Width           =   1155
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Linea Credito"
               Height          =   195
               Left            =   105
               TabIndex        =   28
               Top             =   510
               Width           =   930
            End
            Begin VB.Label LblLinCred 
               BackColor       =   &H80000004&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1110
               TabIndex        =   27
               Top             =   495
               Width           =   4950
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Saldo Capital"
               Height          =   195
               Left            =   105
               TabIndex        =   26
               Top             =   1095
               Width           =   930
            End
            Begin VB.Label LblSalCap 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000004&
               BorderStyle     =   1  'Fixed Single
               Height          =   270
               Left            =   1110
               TabIndex        =   25
               Top             =   1065
               Width           =   1155
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Monto Financiado"
               Height          =   195
               Left            =   2835
               TabIndex        =   24
               Top             =   810
               Width           =   1275
            End
            Begin VB.Label LblMonCred 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000004&
               BorderStyle     =   1  'Fixed Single
               Height          =   270
               Left            =   4290
               TabIndex        =   23
               Top             =   780
               Width           =   1335
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Cliente"
               Height          =   195
               Left            =   120
               TabIndex        =   22
               Top             =   210
               Width           =   480
            End
            Begin VB.Label LblNomCli 
               BackColor       =   &H80000004&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1110
               TabIndex        =   21
               Top             =   195
               Width           =   4950
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Calificacion :"
               Height          =   195
               Left            =   120
               TabIndex        =   20
               Top             =   1695
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.Label LblCalMiViv 
               Appearance      =   0  'Flat
               Caption         =   "Mal Pagador"
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
               Height          =   195
               Left            =   1110
               TabIndex        =   19
               Top             =   1710
               Visible         =   0   'False
               Width           =   1395
            End
         End
         Begin VB.Label LblProxfec 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1755
            TabIndex        =   68
            Top             =   3270
            Width           =   45
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Prox. fecha Pag :"
            Height          =   195
            Left            =   360
            TabIndex        =   67
            Top             =   3240
            Width           =   1230
         End
         Begin VB.Label lblMetLiq 
            Height          =   195
            Left            =   4560
            TabIndex        =   47
            Top             =   3240
            Width           =   645
         End
         Begin VB.Label Label25 
            Caption         =   "Met Liq."
            Height          =   195
            Left            =   3360
            TabIndex        =   46
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label lblCuotasMora 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   45
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Cuotas en Mora"
            Height          =   195
            Left            =   360
            TabIndex        =   44
            Top             =   2520
            Width           =   1125
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Dias Atrasados"
            Height          =   195
            Left            =   3360
            TabIndex        =   43
            Top             =   2880
            Width           =   1065
         End
         Begin VB.Label LblDiasAtraso 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4710
            TabIndex        =   42
            Top             =   3240
            Width           =   45
         End
         Begin VB.Label LblFecVec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
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
            Height          =   270
            Left            =   4515
            TabIndex        =   41
            Top             =   2520
            Width           =   1275
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Venc."
            Height          =   195
            Left            =   3360
            TabIndex        =   40
            Top             =   2520
            Width           =   915
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Monto Calen. Din:"
            Height          =   195
            Left            =   360
            TabIndex        =   39
            Top             =   2880
            Width           =   1275
         End
         Begin VB.Label LblMonCalDin 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1800
            TabIndex        =   38
            Top             =   2880
            Width           =   1155
         End
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Pag.Total. :"
         Height          =   195
         Left            =   4725
         TabIndex        =   61
         Top             =   3870
         Width           =   825
      End
      Begin VB.Label lblPagoTotal 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5580
         TabIndex        =   60
         Top             =   3840
         Width           =   1020
      End
      Begin VB.Label LblItf 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3510
         TabIndex        =   59
         Top             =   3840
         Width           =   1020
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "I.T.F. :"
         Height          =   195
         Left            =   2925
         TabIndex        =   58
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto a Pagar"
         Height          =   195
         Left            =   240
         TabIndex        =   57
         Top             =   3855
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
         Left            =   4695
         TabIndex        =   55
         Top             =   3480
         Width           =   1665
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   3480
         TabIndex        =   54
         Top             =   3510
         Width           =   1050
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
         Height          =   195
         Left            =   240
         TabIndex        =   53
         Top             =   3510
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Operación"
      Height          =   1920
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Operación"
         Height          =   960
         Left            =   4800
         TabIndex        =   2
         Top             =   150
         Width           =   3195
         Begin VB.ListBox LstCred 
            Height          =   450
            Left            =   75
            TabIndex        =   3
            Top             =   225
            Width           =   3060
         End
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3840
         TabIndex        =   1
         Top             =   315
         Width           =   900
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   435
         Left            =   180
         TabIndex        =   4
         Top             =   285
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   767
         Texto           =   "Operación :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Producto"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1515
         Width           =   1005
      End
      Begin VB.Label lblTipoProd 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1365
         TabIndex        =   8
         Top             =   1500
         Width           =   6630
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Crédito"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1215
         Width           =   855
      End
      Begin VB.Label lblTipoCred 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1365
         TabIndex        =   6
         Top             =   1200
         Width           =   6630
      End
      Begin VB.Label LblAgencia 
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
         Left            =   240
         TabIndex        =   5
         Top             =   780
         Width           =   3465
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cargo en cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   62
      Top             =   4560
      Width           =   7455
      Begin VB.ComboBox cboCtasAhorros 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   1080
         Width           =   4575
      End
      Begin SICMACT.TxtBuscar txtPersCodCargo 
         Height          =   375
         Left            =   120
         TabIndex        =   64
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
         sTitulo         =   ""
      End
      Begin VB.Label lblPersCargoNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2640
         TabIndex        =   66
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label6 
         Caption         =   "Cuenta de ahorro"
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
         Left            =   120
         TabIndex        =   65
         Top             =   1080
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmCredpagoCuotasLeasingDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public nProducto As Producto

Private oCredito As COMNCredito.NCOMCredito

'Dim nmoneda As Integer
Private nNroTransac As Long
Private bCalenDinamic As Boolean
Private bCalenCuotaLibre As Boolean
Private bRecepcionCmact As Boolean
'Private sPersCmac As String
Private vnIntPendiente As Double
Private vnIntPendientePagado As Double
Dim nCalPago As Integer
'Dim bDistrib As Boolean
Dim bPrepago As Integer
Dim nCalendDinamTipo As Integer
Dim nMiVivienda As Integer
Dim MatDatos As Variant
Dim sOperacion As String
Dim nOpcion As Integer 'ALPA20140110
Dim sPersCod As String
Dim nInteresDesagio As Double
Dim nSaldoCuenta As Double
Dim nEstado As COMDConstantes.CaptacEstado
Dim dApertura As Date
Dim nMontoPago As Double
Dim nITF As Double

Dim nPrestamo As Double
Dim bCuotaCom As Integer
Dim nCalendDinamico As Integer

Dim bRFA As Boolean

'Lavado de Dinero
Dim bExoneradaLavado As Boolean
Dim sPerscodLav As String
Dim sNombreLav As String
Dim sDireccionLav As String
Dim sDocIdLav As String

'Variables agregadas para el uso de los Componentes
Private bOperacionEfectivo As Boolean
Private nMontoLavDinero As Double
Private nTC As Double

Private bantxtmonpag As Boolean
'ARCV
Private bActualizaMontoPago As Boolean

Dim pnValorChq As Double
Dim lsTemp As String
Dim lsAgeCodAct As String
Dim lsTpoProdCod As String
Dim lsTpoCredCod As String
Dim lnTpoPrograma As Integer
Dim sCuenta As String
Dim nSalCapLeasing As Currency
Dim nGasto, nMora, nInteres, nCapital As Currency
Dim nMora1, nGasto1, nInteres1, nCapital1 As Currency
Dim ldFechaVenc As Date
Private sPersCmac As String
Dim nMonPago As Double
Dim nRedondeoITF As Double
Public Sub Inicia(sCodOpe As String, Optional lnOpcion As Integer = 1)
    bRecepcionCmact = False
    sOperacion = sCodOpe
    'ALPA201040110*****************************************************************
    nOpcion = lnOpcion
    If nOpcion = 1 Then
        Me.Caption = "Pago Cuotas Arrendamiento Financiero"
    End If
    If nOpcion = 2 Then
        Me.Caption = "Pago Pre-Cancelación Arrendamiento Financiero"
    End If
    '******************************************************************************
    Me.Show 1
End Sub

Private Sub CmbForPag_Click()

    LblNumDoc.Caption = ""
    If CmbForPag.ListIndex <> -1 Then
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCheque Then
            MatDatos = frmBuscaCheque.BuscaCheque(gChqEstEnValorizacion, CInt(Mid(ActxCta.NroCuenta, 9, 1)))
            If MatDatos(0) <> "" Then
                LblNumDoc.Caption = MatDatos(4)
                TxtMonPag.Text = MatDatos(3)
                pnValorChq = MatDatos(3)
            Else
                LblNumDoc.Caption = ""
            End If
            LblNumDoc.Visible = True
        Else
            LblNumDoc.Visible = False
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona

    
    LstCred.Clear
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Set oCredito = New COMDCredito.DCOMCredito
        Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(gColocEstVigMor, gColocEstVigVenc, gColocEstVigNorm, gColocEstRefMor, gColocEstRefVenc, gColocEstRefNorm), , True)
        Do While Not R.EOF
            LstCred.AddItem R!cCtaCod
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Set oCredito = Nothing
        'ALPA 20130108***********************************
        'FrmVerCredito.Inicio oPers.sPersCod
        Call FrmVerCredito.Inicio(oPers.sPersCod, True)
        '************************************************
        Me.ActxCta.SetFocusCuenta
        
    End If
    If LstCred.ListCount = 0 Then
        MsgBox "El Cliente No Tiene Creditos Vigentes", vbInformation, "Aviso"
    End If
    Set oPers = Nothing
End Sub
Private Sub cmdCancelar_Click()
    Call LimpiaPantalla
    Call HabilitaActualizacion(False)
    cmdGrabar.Enabled = False
    CmdPlanPagos.Enabled = False
    If Not (oCredito Is Nothing) Then Set oCredito = Nothing
End Sub
Private Sub LimpiaPantalla()
    LimpiaControles Me, True
    InicializaCombos Me
    Frame3.Enabled = False
    'LblEstado.Caption = ""
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    'LblNewSalCap.Caption = ""
    'LblProxfec.Caption = ""
    'LblNewCPend.Caption = ""
    'LblEstado.Caption = ""
    bCalenDinamic = False
    LblAgencia.Caption = ""
    Label23.Visible = False
    LblCalMiViv.Visible = False
    Call LimpiaFlex(FEConceptosLeasing)
End Sub

Private Sub cmdGrabar_Click()
Dim sArrayConceptoPago() As String
Dim i As Integer
Dim lnContador As Integer
Dim sPersLavDinero As String
Dim sError As String
Dim sImprePlanPago As String
Dim sImpreBoleta As String
Dim vPrevio As previo.clsprevio
Dim objPersona As COMDPersona.DCOMPersonas
Set objPersona = New COMDPersona.DCOMPersonas

Dim oCredD As COMDCredito.DCOMCreditos
Dim sVisPersLavDinero As String
Dim loLavDinero As frmMovLavDinero
Dim oCred As COMDCredito.DCOMCredActBD
Dim nOpcionCompra As Integer
Dim nMontoTotal As Currency
Dim nCuotaPagada As Integer
Dim nMontoAmortizada As Currency
Set loLavDinero = New frmMovLavDinero
Dim fnCondicion As Integer 'WIOR 20130301
Dim regPersonaRealizaPago As Boolean 'WIOR 20130301

'*********************************************************************
If CInt(Trim(Right(CmbForPag.Text, 2))) = gColocTipoPagoCheque Then
        If Trim(Me.LblNumDoc.Caption) = "" Then
            MsgBox "Cheque No es Valido", vbInformation, "Aviso"
            Me.CmbForPag.SetFocus
            Exit Sub
        End If
        If IsArray(MatDatos) Then
            If Trim(MatDatos(3)) = "" Then
                MatDatos(3) = "0.00"
            End If
            If Trim(TxtMonPag.Text) = "" Then
                TxtMonPag.Text = "0.00"
            End If
            If CDbl(TxtMonPag.Text) > CDbl(MatDatos(3)) Then
                MsgBox "Monto de Pago No Puede Ser Mayor que el Monto de Cheque", vbInformation, "Aviso"
                TxtMonPag.SetFocus
                Exit Sub
            End If
        End If
        
        'Validar Monto de Cheque
        Dim nValorCh As Double
        Dim nDifValorCh As Double
        Dim nDifTotalCh As Double
        Dim nPagadoTotal As Double
        Set oCredD = New COMDCredito.DCOMCreditos
        nValorCh = oCredD.ObtenerMontoCheque(LblNumDoc.Caption)
        
        nDifValorCh = Format(CDbl(MatDatos(3)), "0.00")
        
        nPagadoTotal = CDbl(lblPagoTotal.Caption)
        nDifTotalCh = (CDbl(nDifValorCh) - CDbl(nPagadoTotal))
        If nDifTotalCh < 0 Then
            MsgBox "No se puede realizar el Pago con Cheque solo dispone de: " & nDifValorCh, vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    
        Dim rsPersVerifica As Recordset
'        Dim i As Integer
        Set rsPersVerifica = New Recordset
        
            Set rsPersVerifica = objPersona.ObtenerDatosPersona(sPersCod)
            If rsPersVerifica!nPersIngresoProm = 0 Or rsPersVerifica!cActiGiro1 = "" Then
                If MsgBox("Necesita Registrar la Ocupacion e Ingreso Promedio de: " + LblNomCli, vbYesNo) = vbYes Then
                    frmPersOcupIngreProm.Inicio sPersCod, LblNomCli, rsPersVerifica!cActiGiro1, rsPersVerifica!nPersIngresoProm
                End If
            End If
     
    If MsgBox("Se va a Efectuar el Pago del Credito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
     
    Dim nMonto As Double
    Dim nmoneda As Integer
    nMonto = CDbl(TxtMonPag.Text)
'    Dim sPersLavDinero As String
    nmoneda = CLng(Mid(ActxCta.NroCuenta, 9, 1))
      
    
    sPersLavDinero = ""
    If bOperacionEfectivo Then
        If Not bExoneradaLavado Then
            If CDbl(TxtMonPag.Text) >= Round(nMontoLavDinero * nTC, 2) Then
                 Call IniciaLavDinero(loLavDinero)
                 sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, ActxCta.NroCuenta, Me.Caption, True, "", , , , , nmoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                 If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
            End If
        End If
    End If
    'WIOR 20130301 ***SEGUN TI-ERS005-2013 ************************************************************
    If loLavDinero.OrdPersLavDinero = "Exit" Then
        Dim oPersonaSPR As UPersona_Cli
        Dim oPersonaU As COMDPersona.UCOMPersona
        Dim nTipoConBN As Integer
        Dim sConPersona As String
        Dim pbClienteReforzado As Boolean
        Dim rsAgeParam As Recordset
        Dim objCred As COMNCredito.NCOMCredito
        Dim lnMonto As Double, lnTC As Double
        Dim ObjTc As COMDConstSistema.NCOMTipoCambio
        
        
        Set oPersonaU = New COMDPersona.UCOMPersona
        Set oPersonaSPR = New UPersona_Cli
        
        regPersonaRealizaPago = False
        pbClienteReforzado = False
        fnCondicion = 0

        oPersonaSPR.RecuperaPersona sPersCod
                            
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
            MsgBox "El Cliente: " & Trim(LblNomCli.Caption) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
            frmPersRealizaOpeGeneral.Inicia Me.Caption & " (Persona " & sConPersona & ")", sOperacion
            regPersonaRealizaPago = frmPersRealizaOpeGeneral.PersRegistrar
            
            If Not regPersonaRealizaPago Then
                MsgBox "Se va a proceder a Anular el Pago de la Cuota", vbInformation, "Aviso"
                Exit Sub
            End If
        Else
            fnCondicion = 0
            lnMonto = nMonto
            pbClienteReforzado = False
            
            Set ObjTc = New COMDConstSistema.NCOMTipoCambio
            lnTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
            Set ObjTc = Nothing
        
        
            Set objCred = New COMNCredito.NCOMCredito
            Set rsAgeParam = objCred.obtieneCredPagoCuotasAgeParam(gsCodAge)
            Set objCred = Nothing
            
            If Mid(ActxCta.NroCuenta, 9, 1) = 2 Then
                lnMonto = Round(lnMonto * lnTC, 2)
            End If
        
            If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                If lnMonto >= rsAgeParam!nMontoMin And lnMonto <= rsAgeParam!nMontoMax Then
                    frmPersRealizaOpeGeneral.Inicia Me.Caption, sOperacion
                    regPersonaRealizaPago = frmPersRealizaOpeGeneral.PersRegistrar
                    If Not regPersonaRealizaPago Then
                        MsgBox "Se va a proceder a Anular el Pago de la Cuota", vbInformation, "Aviso"
                        Exit Sub
                    End If
                End If
            End If
            
        End If
    End If
    'WIOR FIN ***************************************************************
    lsTemp = MDISicmact.SBBarra.Panels(1).Text
    MDISicmact.SBBarra.Panels(1).Text = "Procesando ....."
    Me.cmdGrabar.Enabled = False
'*********************************************************************
lnContador = 1
sPersLavDinero = ""
For i = 1 To FEConceptosLeasing.Rows - 1
    If FEConceptosLeasing.TextMatrix(i, 6) = "+" Then
        ReDim Preserve sArrayConceptoPago(7, 1 To lnContador)
        sArrayConceptoPago(1, lnContador) = FEConceptosLeasing.TextMatrix(i, 1) 'nConceptoCod
        sArrayConceptoPago(2, lnContador) = FEConceptosLeasing.TextMatrix(i, 2) 'Concepto
        sArrayConceptoPago(3, lnContador) = FEConceptosLeasing.TextMatrix(i, 3) 'Monto Pago
        sArrayConceptoPago(4, lnContador) = FEConceptosLeasing.TextMatrix(i, 4) 'nNroCalendario
        sArrayConceptoPago(5, lnContador) = FEConceptosLeasing.TextMatrix(i, 5) 'nCuota
        sArrayConceptoPago(7, lnContador) = FEConceptosLeasing.TextMatrix(i, 7) 'dVenc
        nMontoAmortizada = nMontoAmortizada + FEConceptosLeasing.TextMatrix(i, 3)
        lnContador = lnContador + 1
    End If
    nMontoTotal = nMontoTotal + FEConceptosLeasing.TextMatrix(i, 3)
Next i

nCuotaPagada = 0
If nOpcion = 1 Then
    If nMontoAmortizada = nMonPago Then
    'If nMontoAmortizada = nMontoTotal Then
        nCuotaPagada = 1
    End If
Else
    nCuotaPagada = 1
    nMontoPago = nMontoAmortizada
    nMonPago = nMontoAmortizada
End If
'ALPA20130919************************************
If nCuotaPagada = 0 Then
 If MsgBox("El monto de la cuota no está totalmente cubierta, Desea Continuar sin verificar los montos ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
End If
'************************************************

If ckOC.value = 1 Then
    nOpcionCompra = 1
Else
    nOpcionCompra = 0
End If

    Call oCredito.GrabarPagoOtrasCuotasLeasing(ActxCta.NroCuenta, sArrayConceptoPago, nCalPago, nMontoPago, _
                            gdFecSis, lblMetLiq.Caption, CInt(Trim(Right(CmbForPag.Text, 10))), gsCodAge, gsCodUser, gsCodCMAC, Trim(LblNumDoc.Caption), _
                            bRecepcionCmact, sPersCmac, vnIntPendiente, vnIntPendientePagado, bPrepago, sPersLavDinero, nITF, _
                            nInteresDesagio, CDbl(LblSalCap.Caption), bCalenDinamic, CDbl(LblMonCalDin.Caption), nCalendDinamTipo, gsNomAge, CInt(ActxCta.Prod), _
                            LblNomCli.Caption, LblMoneda.Caption, nNroTransac, LblProxfec.Caption, sLpt, gsInstCmac, IIf(Trim(Right(Me.CmbForPag.Text, 2)) = "2", True, False), _
                            Me.LblNumDoc.Caption, sError, sImprePlanPago, sImpreBoleta, CInt(LblDiasAtraso.Caption), gsProyectoActual, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, lsAgeCodAct, lsTpoProdCod, lsTpoCredCod, LblFecVec.Caption, nOpcionCompra, lnContador, nCuotaPagada, LblCPend.Caption, sOperacion, nOpcion)

    Set vPrevio = New clsprevio
    
    vPrevio.PrintSpool sLpt, sImpreBoleta
    
    Do While MsgBox("Desea Reimprimir el Comprobante de Pago?", vbInformation + vbYesNo, "Aviso") = vbYes
        vPrevio.PrintSpool sLpt, sImpreBoleta
    Loop
    Set vPrevio = Nothing
    
     If gnMovNro > 0 Then
     Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
     End If
    'WIOR 20130301 ***SEGUN TI-ERS005-2013 ************************************************************
    If regPersonaRealizaPago And gnMovNro > 0 Then
        frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(ActxCta.NroCuenta), fnCondicion
        regPersonaRealizaPago = False
    End If
    'WIOR FIN ******************************************************************************************
    'INICIO JHCU ENCUESTA 16-10-2019
    Encuestas gsCodUser, gsCodAge, "ERS0292019", sOperacion
    'FIN
    gVarPublicas.LimpiaVarLavDinero
    
    
    Call cmdCancelar_Click

End Sub
Public Sub RecepcionCmac(ByVal psPersCodCMAC As String)
    bRecepcionCmact = True
    sPersCmac = psPersCodCMAC
    Me.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CargaControles
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    bCalenDinamic = False
    CentraSdi Me
    bantxtmonpag = False
End Sub

Private Sub LstCred_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then
            ActxCta.NroCuenta = LstCred.Text
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub
Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'VerificaRFA (ActxCta.NroCuenta)
        If bRFA Then
            MsgBox "Este credito es RFA " & vbCrLf & _
                   "Por favor ingrese a la opción de Pagos en RFA", vbInformation, ""
        Else
            '*** CONTROLAR TIEMPO DE CARGA
            If Not CargaDatos(ActxCta.NroCuenta) Then
                HabilitaActualizacion False
                MsgBox "No se pudo encontrar el Credito, o el Credito No esta Vigente", vbInformation, "Aviso"
                CmdPlanPagos.Enabled = False
                CmdGastos.Enabled = False
            Else
                '*** CONTROLAR TIEMPO DE CARGA
                CmdPlanPagos.Enabled = True
                CmdGastos.Enabled = True
                HabilitaActualizacion True
                TxtMonPag.Enabled = False 'EJVG20120809
            End If
        End If
    End If
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean

Dim rsPers As ADODB.Recordset
Dim rsCredVig As ADODB.Recordset
Dim rsDatosCuotaLeasing As ADODB.Recordset
Dim sAgencia As String
Dim nGastos As Double
'Dim nMonPago As Double
Dim nMora As Double
Dim nCuotasMora As Integer
Dim nTotalDeuda As Currency
Dim nInteresDesagio As Double
Dim nMonCalDin As Double
Dim sMensaje As String

Dim nNewSalCap As Double
Dim nNewCPend As Integer
Dim dProxFec As Date
Dim sEstado As String

Dim nCuotaPendiente As Integer
Dim nMoraCalculada As Double
Dim dFechaVencimiento As Date

Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset

    On Error GoTo ErrorCargaDatos
    nInteresDesagio = 0
    nGasto = 0
    nMora = 0
    nInteres = 0
    nCapital = 0
    
    If Not (ActxCta.Prod = "515" Or ActxCta.Prod = "516") Then
        MsgBox "Favor digitar un número de contrato de Arrendamiento financiero", vbCritical, "Arrendamiento Financiero"
        Exit Function
    End If
    
    Set oCredito = New COMNCredito.NCOMCredito
    If oCredito.VerificaFechaLeasing(psCtaCod, gdFecSis) = 0 Then
        MsgBox "Cuota aun no se encuentra generada", vbCritical, "Arrendamiento Financiero"
        Exit Function
    End If
    Set oCredito = Nothing
    Set oCredito = New COMNCredito.NCOMCredito
    Call oCredito.CargaDatosPagoCuotasLeasing(psCtaCod, gdFecSis, bPrepago, gsCodAge, rsCredVig, sAgencia, nCalendDinamico, bCalenDinamic, bCalenCuotaLibre, _
                                    nMiVivienda, nCalPago, nGastos, nMonPago, nMora, nCuotasMora, nTotalDeuda, nInteresDesagio, _
                                    nMonCalDin, sMensaje, sPersCod, sOperacion, bExoneradaLavado, bRFA, rsPers, bOperacionEfectivo, nMontoLavDinero, nTC, _
                                    nMontoPago, nITF, vnIntPendientePagado, nNewSalCap, nNewCPend, dProxFec, sEstado, nCuotaPendiente, nMoraCalculada, dFechaVencimiento, rsDatosCuotaLeasing, nOpcion)
    nMora = 0
    If Not rsCredVig.BOF And Not rsCredVig.EOF Then
        If gCredPagLeasingCI = sOperacion And nCuotaPendiente <> 0 Then
            MsgBox "Debe seleccionar la opcion pago de cuota leasing", vbCritical, "Arrendamiento Financiero"
            Exit Function
        End If
        If gCredPagLeasingCU = sOperacion And nCuotaPendiente = 0 Then
            MsgBox "Debe seleccionar la opcion pago de cuota inicial leasing", vbCritical, "Arrendamiento Financiero"
            Exit Function
        End If
        Call LimpiaFlex(FEConceptosLeasing)
        nSalCapLeasing = 0
        Do While Not rsDatosCuotaLeasing.EOF
            FEConceptosLeasing.AdicionaFila
            FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 1) = rsDatosCuotaLeasing!nPrdConceptoCod
            FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 2) = rsDatosCuotaLeasing!cDescripcion
            FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 3) = Format(rsDatosCuotaLeasing!nMontoConcepto, "#0.00")
            FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 4) = rsDatosCuotaLeasing!nNroCalen
            FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 5) = rsDatosCuotaLeasing!nCuota
            FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 6) = "+"
            FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 7) = Format(rsDatosCuotaLeasing!dVenc, "YYYY/MM/DD")
            ldFechaVenc = Format(rsDatosCuotaLeasing!dVenc, "YYYY/MM/DD")
            FEConceptosLeasing.BackColorRow (vbGreen)
            nSalCapLeasing = nSalCapLeasing + rsDatosCuotaLeasing!nMontoConcepto
            
            If Left(FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 1), 2) = "12" Then
                nGasto = nGasto + FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 3)
            ElseIf FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 1) = "1101" Or FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 1) = "1108" Then
                nMora = nMora + FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 3)
            ElseIf Left(FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 1), 2) = "11" And Not (FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 1) = "1101" Or FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 1) = "1108") Then
                nInteres = nInteres + FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 3)
            ElseIf FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 1) = "1000" Then
                nCapital = nCapital + FEConceptosLeasing.TextMatrix(rsDatosCuotaLeasing.Bookmark, 3)
            End If
            rsDatosCuotaLeasing.MoveNext
        Loop
        nMora1 = nMora
        nGasto1 = nGasto
        nInteres1 = nInteres
        nCapital1 = nCapital
    
        LblAgencia.Caption = sAgencia
        lblMetLiq.Caption = Trim(rsCredVig!cMetLiquidacion)
        'LblMontoCuota.Caption = Format(IIf(IsNull(rsCredVig!CuotaAprobada), 0, rsCredVig!CuotaAprobada), "#0.00")
        LblMontoCuota.Caption = Format(nMora + nGasto + nInteres + nCapital, "#0.00")
        nCalendDinamTipo = rsCredVig!nCalendDinamTipo
        
        If nMiVivienda Then
            Label23.Visible = True
            LblCalMiViv.Visible = True
            If nCalPago = 1 Then
                LblCalMiViv.Caption = "Buen Pagador"
            Else
                LblCalMiViv.Caption = "Mal Pagador"
            End If
        Else
            Label23.Visible = False
            LblCalMiViv.Visible = False
        End If
                
        CargaDatos = True
        'ALPA 20100607**************************************
        lblTipoCred.Caption = rsCredVig!cTpoCredDes
        lblTipoProd.Caption = rsCredVig!cTpoProdDes
        lsAgeCodAct = rsCredVig!cAgeCodAct
        lsTpoProdCod = rsCredVig!cTpoProdCod
        lsTpoCredCod = rsCredVig!cTpoCredCod
        '***************************************************
        vnIntPendiente = IIf(IsNull(rsCredVig!nintPend), 0, rsCredVig!nintPend)
        vnIntPendientePagado = 0
        nNroTransac = IIf(IsNull(rsCredVig!nTransacc), 0, rsCredVig!nTransacc)
        LblNomCli.Caption = PstaNombre(rsCredVig!cPersNombre)
        LblLinCred.Caption = Trim(rsCredVig!cLineaCred)
        LblMoneda.Caption = Trim(rsCredVig!cmoneda)
        LblMonCred.Caption = Format(rsCredVig!nMontoCol, "#0.00")
        
        ' CMACICA_CSTS - 08/11/2003 -------------------------------------------------------------------
        nPrestamo = Format(rsCredVig!nMontoCol, "#0.00")
        bCuotaCom = IIf(IsNull(rsCredVig!bCuotaCom), 0, rsCredVig!bCuotaCom)
        '----------------------------------------------------------------------------------------------
        
        LblSalCap.Caption = Format(rsCredVig!nSaldo, "#0.00")
        LblForma.Caption = Trim(Str(rsCredVig!nCuotasApr))
        LblFecVec.Caption = dFechaVencimiento
        LblCPend.Caption = nCuotaPendiente
        '-----------------------
        lblCuotasMora.Caption = nCuotasMora
        LblDiasAtraso.Caption = Trim(Str(rsCredVig!nDiasAtraso))
        lblMetLiq.Caption = Trim(rsCredVig!cMetLiquidacion)
        TxtMonPag.Text = nMora + nGasto + nInteres + nCapital 'nMonPago
        LblTotDeuda.Caption = nTotalDeuda
        
        
        'Para Generar el calendario Dinamico
        'Si es mivivienda
        LblMonCalDin.Caption = nMonCalDin
        
        If nMiVivienda = 1 And bPrepago = 0 Then
            TxtMonPag.Locked = True
        Else
            TxtMonPag.Locked = False
        End If
        
        If sMensaje <> "" Then
            MsgBox sMensaje, vbInformation, "Mensaje"
            Exit Function
        End If
        
    bantxtmonpag = True
    TxtMonPag.Text = Format(TxtMonPag.Text, "#0.00")
    
    'EJVG20120809 ***
    Dim lnMontoPagarF As Currency
    Dim oITF As COMDConstSistema.FCOMITF
    
    Set oITF = New COMDConstSistema.FCOMITF
    oITF.fgITFParametros
    lnMontoPagarF = CCur(Me.TxtMonPag.Text)
    nITF = lnMontoPagarF * oITF.gnITFPorcent
    nITF = oITF.CortaDosITF(nITF)
    Set oITF = Nothing
    'END EJVG *******
    
    lblITF.Caption = Format(nITF, "#0.00")
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
    If nRedondeoITF > 0 Then
        Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
    End If
    lblPagoTotal.Caption = Format(Val(TxtMonPag.Text) + Val(Me.lblITF.Caption), "#0.00")
      
    bantxtmonpag = False
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus

    bActualizaMontoPago = False
    '-----------------------

        
        If Not rsPers.EOF Then
            sPerscodLav = sPersCod
            sNombreLav = rsPers!Nombre
            sDireccionLav = rsPers!Direccion
            sDocIdLav = rsPers!id & " " & rsPers![ID N°]
        End If
        
         '************ firma madm
         Set lafirma = New frmPersonaFirma
         Set ClsPersona = New COMDPersona.DCOMPersonas
        
         Set Rf = ClsPersona.BuscaCliente(sPersCod, BusquedaCodigo)
         If Not Rf.BOF And Not Rf.EOF Then
            If Rf!nPersPersoneria = 1 Then
            Call frmPersonaFirma.Inicio(Trim(sPersCod), Mid(sPersCod, 4, 2), False, True)
            End If

         Set Rf = Nothing

        '************ firma madm
        End If
    Else
        CargaDatos = False
    End If
    
    Exit Function

ErrorCargaDatos:
    MsgBox err.Description, vbCritical, "Aviso"

End Function

Private Function HabilitaActualizacion(ByVal pbHabilita As Boolean) As Boolean
    cmdmora.Enabled = pbHabilita
    Frame4.Enabled = Not pbHabilita
    CmbForPag.Enabled = pbHabilita
    LblNumDoc.Enabled = pbHabilita
    TxtMonPag.Enabled = pbHabilita
    If Mid(ActxCta.NroCuenta, 9, 1) = "1" Or Trim(Mid(ActxCta.NroCuenta, 9, 1)) = "" Then
        TxtMonPag.BackColor = vbWhite
        lblITF.BackColor = vbWhite
    Else
        TxtMonPag.BackColor = vbGreen
        lblITF.BackColor = vbGreen
    End If
    Frame3.Enabled = pbHabilita
    If CmbForPag.ListCount > 0 Then
        CmbForPag.ListIndex = 0
    End If
    If pbHabilita Then
        If TxtMonPag.Enabled And TxtMonPag.Visible Then
            TxtMonPag.SetFocus
        End If
    End If
End Function

Private Sub MostrarDetalleActivoDesactivo(ByRef nPos As Integer)
    Dim nCol As Integer
    nCol = FEConceptosLeasing.Col
    nPos = FEConceptosLeasing.row
    If Trim(FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1)) <> "" Then
        FEConceptosLeasing.Col = 2
        If FEConceptosLeasing.CellBackColor = vbGreen Then
            FEConceptosLeasing.BackColorRow (vbWhite)
            FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 6) = "-"
            nSalCapLeasing = nSalCapLeasing - FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
        Else
            FEConceptosLeasing.BackColorRow (vbGreen)
            FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 6) = "+"
            nSalCapLeasing = nSalCapLeasing + FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
        End If
    End If
    FEConceptosLeasing.Col = nCol
    
     Dim oITF As New COMDConstSistema.FCOMITF
     Dim lnValor As Double
     Dim pnITF As Double
     Set oITF = New COMDConstSistema.FCOMITF
     oITF.fgITFParametros
     If ActxCta.Prod = "801" Then
     Else
        lnValor = nSalCapLeasing * oITF.gnITFPorcent
        lnValor = oITF.CortaDosITF(lnValor)
        pnITF = lnValor
     End If
     TxtMonPag.Text = nSalCapLeasing
     lblPagoTotal.Caption = Format(nSalCapLeasing + pnITF, "#0.00")
     lblITF.Caption = Format(pnITF, "#0.00")
     
     nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
     If nRedondeoITF > 0 Then
        Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
     End If
     
End Sub
'private sub
'lblMetLiq

Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R = oCons.RecuperaConstantes(gColocTipoPago)
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, CmbForPag)
    Exit Sub

ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
    
End Sub
Private Sub FEConceptosLeasing_Click()
    Dim nPos As Integer
    Dim nLogico As Integer
    nLogico = 0
    Call VerificarMetodoLiquidacion(lblMetLiq.Caption, nLogico)
    If nLogico = 1 Then
        Call MostrarDetalleActivoDesactivo(nPos)
    End If
End Sub
Private Sub VerificarMetodoLiquidacion(ByVal psMetLiq As String, ByRef pnActivar As Integer)
        Dim i As Integer
        Dim c1, c2, c3, c4 As String
        Dim GMIC As Currency
'        Dim nMora1, nGasto1, nInteres1, nCapital1 As Currency
        
        c1 = Mid(psMetLiq, 1, 1)
        c2 = Mid(psMetLiq, 2, 1)
        c3 = Mid(psMetLiq, 3, 1)
        c4 = Mid(psMetLiq, 4, 1)
        
        GMIC = nMora + nGasto + nInteres + nCapital
        
'        nMora1 = nMora
'        nGasto1 = nGasto
'        nInteres1 = nInteres
'        nCapital1 = nCapital
        
        
        If Left(FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1), 2) = "12" Then
                If FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 6) = "-" Then
                    nGasto1 = nGasto1 + FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                Else
                    nGasto1 = nGasto1 - FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                End If
        ElseIf FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1) = "1101" Or FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1) = "1108" Then
                If FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 6) = "-" Then
                    nMora1 = nMora1 + FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                Else
                    nMora1 = nMora1 - FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                End If
        ElseIf Left(FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1), 2) = "11" And Not (FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1) = "1101" Or FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1) = "1108") Then
                If FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 6) = "-" Then
                    nInteres1 = nInteres1 + FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                Else
                    nInteres1 = nInteres1 - FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                End If
        ElseIf FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1) = "1000" Then
                If FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 6) = "-" Then
                    nCapital1 = nCapital1 + FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                Else
                    nCapital1 = nCapital1 - FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                End If
        End If

        If c4 = "C" Then
            If FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1) = "1000" Then
                If FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 6) = "-" Then
                     If (nMora1 + nGasto1 + nInteres1) = 0 And (nMora + nGasto + nInteres) > 0 Then
                        Exit Sub
                     ElseIf ((nMora + nGasto + nInteres) - (nMora1 + nGasto1 + nInteres1)) > 0 Then
                        nCapital1 = nCapital1 - FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                        Exit Sub
                     End If
                End If
            End If
        End If
        If c3 = "I" Then
            If Left(FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1), 2) = "11" And Not (FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1) = "1101" Or FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1) = "1108") Then
                If FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 6) = "-" Then
                    If ((nMora1 + nGasto1) - (nMora + nGasto)) <> 0 And nCapital1 = 0 Then
                        nInteres1 = nInteres1 - FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                        Exit Sub
                     ElseIf (nMora1 + nGasto1) = 0 And nCapital1 > 0 Then
                        nInteres1 = nInteres1 - FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                        Exit Sub
                     ElseIf (nMora1 + nGasto1) = 0 And nCapital1 = 0 Then
                        nInteres1 = nInteres1 - FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                        Exit Sub
                     End If
                ElseIf FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 6) = "+" Then
                    If nCapital1 > 0 Then
                        nInteres1 = nInteres1 + FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                        Exit Sub
                    End If
                End If
            End If
        End If
        If c2 = "M" Then 'Probar con una Mora
            If FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1) = "1101" Or FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1) = "1108" Then
                If FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 6) = "-" Then
                     If (((nMora + nGasto) - (nMora1 + nGasto1)) > 0) And nMora1 = 0 Then
                        nMora1 = nMora1 - FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                        Exit Sub
                     End If
                ElseIf FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 6) = "+" Then
                    If nCapital1 > 0 Then
                        nMora1 = nMora1 + FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                        Exit Sub
                    End If
                End If
            End If
        End If
        If c1 = "G" Then
            If Left(FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 1), 2) = "12" Then
                If FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 6) = "-" Then
                     If (nCapital1 + nInteres1 + nMora1) > 0 Then
                        nGasto1 = nGasto1 - FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                        Exit Sub
                     End If
                ElseIf FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 6) = "+" Then
                    If (nCapital1 + nMora1 + nInteres1) > 0 Then
                        nGasto1 = nGasto1 + FEConceptosLeasing.TextMatrix(FEConceptosLeasing.row, 3)
                        Exit Sub
                    'ElseIf (nCapital1 + nMora1 + nInteres1) = 0 Then
                    End If
                End If
            End If
        End If
        pnActivar = 1
End Sub

Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero)

    Dim nMonto As Double
    
    nMonto = CDbl(TxtMonPag.Text)
    poLavDinero.TitPersLavDinero = sPersCod
    poLavDinero.OrdPersLavDinero = sPersCod

End Sub
