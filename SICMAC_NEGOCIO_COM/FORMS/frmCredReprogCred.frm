VERSION 5.00
Begin VB.Form frmCredReprogCred 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reprogramacion de Credito"
   ClientHeight    =   8580
   ClientLeft      =   1575
   ClientTop       =   2400
   ClientWidth     =   11055
   Icon            =   "frmCredReprogCred.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frMontCuotaCovid 
      Caption         =   "Monto Cuota"
      Height          =   615
      Left            =   9240
      TabIndex        =   70
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
      Begin SICMACT.EditMoney txtMontoCuota 
         Height          =   255
         Left            =   120
         TabIndex        =   71
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
   Begin VB.Frame Frame5 
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
      Left            =   120
      TabIndex        =   54
      Top             =   6480
      Width           =   10815
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
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
         Left            =   6000
         TabIndex        =   74
         Top             =   275
         Width           =   1290
      End
      Begin VB.ComboBox cmbCovidReprog 
         Height          =   315
         ItemData        =   "frmCredReprogCred.frx":030A
         Left            =   3120
         List            =   "frmCredReprogCred.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox CmbReprogNatEspeciales 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   240
         Width           =   2895
      End
      Begin VB.OptionButton optReprogNatEspeciales 
         Caption         =   "Zona Afectada por Inundación"
         Enabled         =   0   'False
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
         Index           =   5
         Left            =   120
         TabIndex        =   61
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdSalir 
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
         TabIndex        =   60
         Top             =   275
         Width           =   1050
      End
      Begin VB.CommandButton cmdGastos 
         Caption         =   "Gastos"
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
         Left            =   7320
         TabIndex        =   59
         Top             =   275
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
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
         Left            =   8400
         TabIndex        =   58
         Top             =   275
         Width           =   1170
      End
      Begin VB.CommandButton cmdReprogramar 
         Caption         =   "Reprogramar"
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
         Left            =   5880
         TabIndex        =   57
         Top             =   275
         Width           =   1410
      End
      Begin VB.OptionButton optReprogNatEspeciales 
         Caption         =   "Emergencia Sanitaria - Covid-19"
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   3240
         TabIndex        =   56
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.OptionButton optReprogNatEspeciales 
         Caption         =   "Desastre natural o conflicto social"
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   1560
         TabIndex        =   55
         Top             =   600
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Glosa "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   33
      Top             =   7320
      Width           =   10815
      Begin VB.TextBox TxtGlosa 
         Height          =   840
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   10575
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1650
      Left            =   120
      TabIndex        =   6
      Top             =   6480
      Width           =   10815
      Begin VB.OptionButton optReprogNatEspeciales 
         Caption         =   "Desastre natural o conflicto social"
         Enabled         =   0   'False
         Height          =   195
         Index           =   4
         Left            =   6360
         TabIndex        =   51
         Top             =   1320
         Width           =   2775
      End
      Begin VB.OptionButton optReprogNatEspeciales 
         Caption         =   "Emergencia Café (Roya Amarilla)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   50
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CommandButton cmdReprog 
         Caption         =   "Reprogramar"
         Height          =   375
         Left            =   6960
         TabIndex        =   45
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtDiasporReprog 
         Height          =   375
         Left            =   2880
         TabIndex        =   43
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdGastos_ 
         Caption         =   "&Gastos"
         Enabled         =   0   'False
         Height          =   345
         Left            =   9240
         TabIndex        =   42
         ToolTipText     =   "Generar Plan Gastos"
         Top             =   1200
         Width           =   1320
      End
      Begin VB.CheckBox chkcalendOrig 
         Caption         =   "De acuerdo a condiciones Originales del Crédito."
         Enabled         =   0   'False
         Height          =   375
         Left            =   6630
         TabIndex        =   36
         Top             =   1140
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.Frame FraOptRepro 
         Height          =   465
         Left            =   165
         TabIndex        =   30
         Top             =   480
         Width           =   3690
         Begin VB.OptionButton OptRepro 
            Caption         =   "REPROGRAMAR"
            Height          =   195
            Index           =   1
            Left            =   1035
            TabIndex        =   32
            Top             =   195
            Width           =   1665
         End
         Begin VB.OptionButton OptRepro 
            Caption         =   "Reprogramar"
            Height          =   195
            Index           =   0
            Left            =   2310
            TabIndex        =   31
            Top             =   195
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1290
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Renovar A "
         Height          =   960
         Left            =   6600
         TabIndex        =   27
         Top             =   480
         Width           =   2460
         Begin VB.CommandButton CmdGenCal 
            Caption         =   "Generar Calendario"
            Enabled         =   0   'False
            Height          =   450
            Left            =   240
            TabIndex        =   41
            Top             =   240
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.OptionButton OptRenov 
            Caption         =   "Menor Cuota Mismo Plazo"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   29
            Top             =   660
            Visible         =   0   'False
            Width           =   2160
         End
         Begin VB.OptionButton OptRenov 
            Caption         =   "Misma Cuota Menor Plazo"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   28
            Top             =   315
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   2160
         End
      End
      Begin VB.CommandButton CmdReprogramar_ 
         Caption         =   "Renovar Credito"
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
         Height          =   390
         Left            =   705
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   360
         Left            =   2040
         TabIndex        =   14
         Top             =   675
         Width           =   1170
      End
      Begin VB.CommandButton CmdSalir_ 
         Caption         =   "&Salir"
         Enabled         =   0   'False
         Height          =   360
         Left            =   810
         TabIndex        =   13
         Top             =   675
         Width           =   1170
      End
      Begin VB.CommandButton CmdCancelar_ 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   405
         Left            =   9240
         TabIndex        =   12
         Top             =   720
         Width           =   1305
      End
      Begin VB.CommandButton CmdAceptar_ 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   405
         Left            =   9240
         TabIndex        =   11
         Top             =   720
         Width           =   1290
      End
      Begin VB.Frame fraIntRep 
         Caption         =   "Interes Reprogramado"
         Height          =   1080
         Left            =   4155
         TabIndex        =   8
         Top             =   480
         Width           =   2400
         Begin VB.OptionButton OptTipoRep 
            Caption         =   "Segun CMAC ICA"
            Height          =   360
            Index           =   3
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.OptionButton OptTipoRep 
            Caption         =   "Reprogramacion Especial"
            Height          =   360
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   2235
         End
         Begin VB.OptionButton OptTipoRep 
            Caption         =   "Adicionar una Ultima Cuota"
            Height          =   360
            Index           =   1
            Left            =   105
            TabIndex        =   10
            Top             =   450
            Width           =   2235
         End
         Begin VB.OptionButton OptTipoRep 
            Caption         =   "Proratear"
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   9
            Top             =   270
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
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
         Left            =   195
         TabIndex        =   7
         Top             =   675
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label7 
         Caption         =   "Dias por Reprogramar :"
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
         Left            =   480
         TabIndex        =   44
         Top             =   1155
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   10815
      Begin SICMACT.FlexEdit FECalend 
         Height          =   2625
         Left            =   120
         TabIndex        =   5
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
   Begin VB.Frame FraDatos 
      Height          =   3225
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.Frame fr_TasaEspecial 
         Caption         =   "Tasa Especial"
         Height          =   615
         Left            =   7800
         TabIndex        =   72
         Top             =   2520
         Width           =   1215
         Begin VB.Label lbl_TasaEspecial 
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
            TabIndex        =   73
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox txtTCEAReprog 
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
         TabIndex        =   67
         Text            =   "0%"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox txtTCEADes 
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
         TabIndex        =   66
         Text            =   "0%"
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtTEA 
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
         TabIndex        =   65
         Text            =   "0%"
         Top             =   2880
         Width           =   975
      End
      Begin VB.ListBox LstCtas 
         Height          =   645
         Left            =   8325
         TabIndex        =   4
         Top             =   210
         Width           =   2235
      End
      Begin VB.CommandButton CmdBuscar 
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
         TabIndex        =   2
         Top             =   260
         Width           =   1230
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   480
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   847
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
         CMAC            =   "108"
      End
      Begin VB.Label Label13 
         Caption         =   "TEA. :"
         Height          =   255
         Left            =   2640
         TabIndex        =   64
         Top             =   2930
         Width           =   495
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
         TabIndex        =   63
         Top             =   2930
         Width           =   1095
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
         TabIndex        =   62
         Top             =   2930
         Width           =   975
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
         TabIndex        =   53
         Top             =   1365
         Width           =   1485
      End
      Begin VB.Label lblDiasReprog 
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
         TabIndex        =   52
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label lblTipoProducto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   49
         Top             =   1320
         Width           =   5985
      End
      Begin VB.Label lblTipoCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   48
         Top             =   1725
         Width           =   5985
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
         TabIndex        =   47
         Top             =   1755
         Width           =   1095
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
         TabIndex        =   46
         Top             =   1365
         Width           =   1215
      End
      Begin VB.Label lblfecUltCuota 
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
         TabIndex        =   40
         Top             =   960
         Width           =   1185
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
         TabIndex        =   39
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label lblSaldoRep 
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
         TabIndex        =   38
         Top             =   1680
         Width           =   1200
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
         TabIndex        =   37
         Top             =   1755
         Width           =   1575
      End
      Begin VB.Label LblSaldo 
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
         TabIndex        =   25
         Top             =   2520
         Width           =   1035
      End
      Begin VB.Label Saldo 
         AutoSize        =   -1  'True
         Caption         =   "Saldo :"
         Height          =   285
         Left            =   2640
         TabIndex        =   24
         Top             =   2565
         Width           =   495
      End
      Begin VB.Label LblPrestamo 
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
         TabIndex        =   22
         Top             =   2565
         Width           =   765
      End
      Begin VB.Label LblAnalista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Top             =   2115
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
         TabIndex        =   20
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label lblTitular 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   915
         Width           =   5985
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
         TabIndex        =   18
         Top             =   960
         Width           =   525
      End
      Begin VB.Label LblTasa 
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
         TabIndex        =   16
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Interes : "
         Height          =   285
         Left            =   4440
         TabIndex        =   15
         Top             =   2565
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmCredReprogCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'REDISEÑADO X JUEZ 20160316

Option Explicit
 
Private nFilaEditar As Integer
Private dFecTemp As Date
Private nMontoApr As Double
Private fnTasaInteres As Double
Private MatCalend As Variant
Private nTipoReprogCred As Integer
Dim ldVigencia As Date
Dim dFecUltCuota As Date

'Add By gitu 14-08-08
Dim nCuoPag As Integer
Dim nCuoNoPag As Integer
Dim lnCapital As Double
Dim lnIntComp As Double
Dim lnIntGra As Double
Dim loVistoElectronico As frmVistoElectronico
Dim lbVistoVal As Boolean
Dim lsPersCod As String
'End GITU

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

Dim bValidaActualizacionLiq As Boolean 'RIRO 20200911

Private Sub HabilitaControlesReprog(ByVal pbHabilita As Boolean)
    FECalend.lbEditarFlex = pbHabilita
    CmdEditar.Enabled = Not pbHabilita
    'CmdReprogramar.Enabled = Not pbHabilita
    CmdNuevo.Enabled = Not pbHabilita
    'cmdSalir.Enabled = Not pbHabilita
    fraIntRep.Enabled = Not pbHabilita
    'cmdAceptar.Enabled = pbHabilita
    'cmdCancelar.Enabled = pbHabilita
    FraOptRepro.Enabled = Not pbHabilita
    If Me.OptTipoRep(2).value Then
        FECalend.lbEditarFlex = False
    End If
    txtDiasporReprog.Enabled = Not pbHabilita 'EJVG20121211
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
    LimpiaFlex FECalend
    MatCalend = ""
    
    Call CargaCombo 'JOEP20200706 Mantener Cuota
    
    'Obtiene los datos del crédito a reprogramar
    Set oCredito = New COMDCredito.DCOMCredito
    Call oCredito.CargarDatosReprogramacion(psCtaCod, rsCred, rsCal, rsReprogApr, rsDatosReprog, bAutorizado)
    Set oCredito = Nothing
    
    'Validaciones de los registros cargados
    If rsCred.EOF Or rsCred.BOF Then
        MsgBox "El crédito no existe, Favor de verificar", vbInformation, "Aviso"
        rsCred.Close
        Exit Function
    End If
    
    If rsDatosReprog.EOF Or rsDatosReprog.BOF Then
        MsgBox "El crédito no tiene solicitud", vbInformation, "Aviso"
        rsReprogApr.Close
        Exit Function
    End If
    
    If rsReprogApr.EOF Or rsReprogApr.BOF Then
        MsgBox "El crédito no está aprobado para ser reprogramado", vbInformation, "Aviso"
        rsReprogApr.Close
        Exit Function
    End If
       
    If CInt(rsDatosReprog!nDiasAtraso) <= 0 And CInt(rsCred!nDiasAtraso) > 0 And Not bAutorizado Then
        MsgBox "El crédito actualmente ya posee " & CStr(rsCred!nDiasAtraso) & IIf(CInt(rsCred!nDiasAtraso) = 1, " día", " días") & " de atraso, será necesario rechazar esta solicitud y registrar una nueva", vbInformation, "Aviso"
        rsCred.Close
        Exit Function
    End If
       
    If rsCal.BOF Or rsCal.EOF Then
        MsgBox "No se puede obtener el calendario del crédito", vbInformation, "Aviso"
        rsCal.Close
        Exit Function
    End If
        
    'RIRO20200911 VALIDA LIQUIDACION ***************
    Dim oCreditoTmp As COMNCredito.NCOMCredito
    Set oCreditoTmp = New COMNCredito.NCOMCredito
    bValidaActualizacionLiq = oCreditoTmp.VerificaActualizacionLiquidacion(psCtaCod)
    If Not bValidaActualizacionLiq Then
        MsgBox "El crédito no tiene actualizados sus datos de liquidación, no podrá realizar " & _
        "la reprogramación mientras no se actualicen estos datos. Deberá comunicarse con el área de T.I.", vbExclamation, "Aviso"
        Exit Function
    End If
    'END RIRO **************************************
    
'JOEP20210211 Garantia covid
    Dim oDCred As COMDCredito.DCOMCredito
    Dim rsvalida As ADODB.Recordset
    Set oDCred = New COMDCredito.DCOMCredito
    Set rsvalida = oDCred.ReprogramacionPropuestaMsgBox(psCtaCod, 5)
    If Not (rsvalida.EOF And rsvalida.BOF) Then
        If rsvalida!cMgsBox <> "" Then
            MsgBox rsvalida!cMgsBox, vbInformation, "Aviso"
            Exit Function
        End If
    End If
    Set oDCred = Nothing
    RSClose rsvalida
'JOEP20210211 Garantia covid
    
    'Asignación de valores según tablas de memoria
    nPrdEstado = rsCred!nPrdEstado
    fnTasaInteres = CDbl(Format(rsCred!nTasaInteres, "#,##0.0000"))
    lblTitular.Caption = PstaNombre(rsCred!cTitular)
    LblAnalista.Caption = PstaNombre(rsCred!cAnalista)
    LblSaldo.Caption = Format(rsCred!nSaldo, "#,##0.00")
    LblPrestamo.Caption = Format(rsCred!nMontoCol, "#,##0.00")
    LblTasa.Caption = Format(fnTasaInteres, "#0.0000")
    ldVigencia = Format(rsCred!dVigencia, "dd/mm/yyyy")
    fnTasaSegDes = Format(rsCred!nTasaSegDesg, "#0.0000")
    lsPersCod = rsCred!cPersCod
    'ALPA 20100606***************
    lblTipoCredito.Caption = rsCred!cTpoCredDes
    lblTipoProducto.Caption = rsCred!cTpoProdDes
    '****************************
    fdFechaCuotaPend = Format(rsCred!dVenc_Cuota, "dd/mm/yyyy") 'LUCV20180601, Según ERS022-2018
    lblDiasReprog.Caption = DateDiff("d", CDate(rsDatosReprog!dFecCuotaVenc), CDate(rsDatosReprog!dFecNuevaCuotaVenc))
    txtDiasporReprog.Text = DateDiff("d", CDate(rsDatosReprog!dFecCuotaVenc), CDate(rsDatosReprog!dFecNuevaCuotaVenc))
        
    'JOEP20200425 cuota igual covid (en Reprogramados)
    txtTEA = rsCred!TEA
    txtTCEADes = rsCred!TCEA
    'txtMontoCuota = rsCred!nMontoCuota
    'JOEP20200425 cuota igual covid (en Reprogramados)
    
    rsCred.Close
    Set rsCred = Nothing
    
'    If nPrdEstado <> gColocEstVigNorm And nPrdEstado <> gColocEstVigVenc And nPrdEstado <> gColocEstVigMor _
'        And nPrdEstado <> gColocEstRefNorm And nPrdEstado <> gColocEstRefMor And nPrdEstado <> gColocEstRefVenc Then
'        CargaDatos = False
'        R.Close
'        Exit Function
'    End If
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
        
        FECalend.AdicionaFila
        FECalend.TextMatrix(rsCal.Bookmark, 1) = Format(rsCal!dVenc, "dd/mm/yyyy")
        FECalend.TextMatrix(rsCal.Bookmark, 2) = Trim(str(rsCal!nCuota))
        FECalend.TextMatrix(rsCal.Bookmark, 3) = Format(IIf(IsNull(lnCapital), 0, lnCapital) + _
                                        IIf(IsNull(lnIntComp), 0, lnIntComp) + _
                                        IIf(IsNull(lnIntGra), 0, lnIntGra) + _
                                        IIf(IsNull(rsCal!nIntMor), 0, rsCal!nIntMor) + _
                                        IIf(IsNull(rsCal!nIntReprog), 0, rsCal!nIntReprog) + _
                                        IIf(IsNull(rsCal!nGasto), 0, rsCal!nGasto), "#0.00")
        FECalend.TextMatrix(rsCal.Bookmark, 4) = Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")
        FECalend.TextMatrix(rsCal.Bookmark, 5) = Format(IIf(IsNull(lnIntComp), 0, lnIntComp), "#0.00")
        FECalend.TextMatrix(rsCal.Bookmark, 6) = Format(IIf(IsNull(rsCal!nIntMor), 0, rsCal!nIntMor), "#0.00")
        FECalend.TextMatrix(rsCal.Bookmark, 7) = Format(IIf(IsNull(rsCal!nIntReprog), 0, rsCal!nIntReprog), "#0.00")
        FECalend.TextMatrix(rsCal.Bookmark, 8) = Format(IIf(IsNull(lnIntGra), 0, lnIntGra), "#0.00")
        FECalend.TextMatrix(rsCal.Bookmark, 9) = Format(IIf(IsNull(rsCal!nGasto), 0, rsCal!nGasto), "#0.00")
        nMontoApr = nMontoApr - IIf(IsNull(rsCal!nCapital), 0, rsCal!nCapital)
        nMontoApr = CDbl(Format(nMontoApr, "#0.00"))
        FECalend.TextMatrix(rsCal.Bookmark, 10) = Format(nMontoApr, "#0.00")
        FECalend.TextMatrix(rsCal.Bookmark, 11) = Trim(str(rsCal!nColocCalendEstado))
        FECalend.TextMatrix(rsCal.Bookmark, 12) = Format(IIf(IsNull(rsCal!nCapitalPag), 0, rsCal!nCapitalPag), "#0.00")
        
        lnSaldoNew = lnSaldoNew + IIf(IsNull(rsCal!nCapital), 0, rsCal!nCapital) - IIf(IsNull(rsCal!nCapitalPag), 0, rsCal!nCapitalPag)
        'End Gitu
        
        If rsCal!nColocCalendEstado = gColocCalendEstadoPagado Then
            FECalend.row = rsCal.Bookmark
            Call FECalend.ForeColorRow(vbRed)
            nCuoNoPag = nCuoNoPag - 1
        End If
        If rsCal.RecordCount = rsCal.Bookmark Then
            lblfecUltCuota = Format(rsCal!dVenc, "dd/mm/yyyy")
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
    lblSaldoRep = Format(lnSaldoNew, "#,##0.00")
    
    If bAutorizado Then
        CmbReprogNatEspeciales.Visible = True 'JOEP20200428 Covid Cuota Iguales
        'optReprogNatEspeciales(1).Enabled = IIf(val(txtDiasporReprog.Text) >= 29, True, False) 'JOEP20200324 Comento Covid-19
        
        'comento JOEP20200428 covid cuotas iguales
        'optReprogNatEspeciales(1).Enabled = IIf(val(txtDiasporReprog.Text) > 0, True, False) 'JOEP20200324 Covid-19 add
        'optReprogNatEspeciales(2).Enabled = IIf(val(txtDiasporReprog.Text) >= 29, True, False)
        'optReprogNatEspeciales(5).Enabled = IIf(val(txtDiasporReprog.Text) >= 29, True, False) 'PASI20170405
        'comento JOEP20200428 covid cuotas iguales
    End If
    
    fnIntGraciaPend = nMontIntGraciaTotal - nMontIntGracCuotaPag 'JOEP
    
    CargaDatos = True
    'ARCV 24-10-2006
'    Dim MatCalend_2 As Variant
'    Dim psIntCompFecha As Double
'    Dim psIntMorFecha As Double
'    Dim oNCredito As COMNCredito.NCOMCredito
'    Set oNCredito = New COMNCredito.NCOMCredito
'    MatCalend_2 = oNCredito.RecuperaMatrizCalendarioPendiente(psCtaCod)
'
'    psIntCompFecha = Format(oNCredito.MatrizInteresTotalesAFechaSinMora(psCtaCod, MatCalend_2, gdFecSis), "#0.00")
'    psIntMorFecha = Format(oNCredito.ObtenerMoraVencida(gdFecSis, MatCalend_2), "#0.00")
'
'    lblSaldoRep.Caption = Format(CDbl(LblSaldo.Caption) + psIntCompFecha + psIntMorFecha, "#0.00")
'
'    Set oNCredito = Nothing
    '-----------------------
    Exit Function
ErrorCargaDatos:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Private Sub Carga_Mat_A_Flex(ByVal MatCalend As Variant)
Dim i As Integer
Dim lnSaldoNew As Double
    lnSaldoNew = 0
    For i = 0 To UBound(MatCalend) - 1
        FECalend.AdicionaFila
        FECalend.TextMatrix(i + 1, 1) = MatCalend(i, 0)
        FECalend.TextMatrix(i + 1, 2) = MatCalend(i, 1)
        FECalend.TextMatrix(i + 1, 3) = Format(CDbl(MatCalend(i, 2)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 10)) + CDbl(MatCalend(i, 14)), "#0.00")
        FECalend.TextMatrix(i + 1, 4) = Format(CDbl(MatCalend(i, 2)), "#0.00")
        FECalend.TextMatrix(i + 1, 5) = Format(CDbl(MatCalend(i, 4)), "#0.00")
        FECalend.TextMatrix(i + 1, 6) = Format(CDbl(MatCalend(i, 6)), "#0.00")
        FECalend.TextMatrix(i + 1, 7) = Format(CDbl(MatCalend(i, 8)), "#0.00")
        FECalend.TextMatrix(i + 1, 8) = Format(CDbl(MatCalend(i, 10)), "#0.00")
        FECalend.TextMatrix(i + 1, 9) = Format(CDbl(MatCalend(i, 14)), "#0.00")
        FECalend.TextMatrix(i + 1, 10) = MatCalend(i, 16)
        FECalend.TextMatrix(i + 1, 11) = MatCalend(i, 17)
        FECalend.TextMatrix(i + 1, 12) = Format(CDbl(MatCalend(i, 3)), "#0.00")
        FECalend.row = i + 1
        
        If CInt(FECalend.TextMatrix(i + 1, 11)) = gColocCalendEstadoPagado Then
            FECalend.row = i + 1
            Call FECalend.ForeColorRow(vbRed)
        Else
            FECalend.ForeColorRow (vbBlack)
        End If
        lnSaldoNew = lnSaldoNew + CDbl(MatCalend(i, 2)) - CDbl(MatCalend(i, 3))
    Next i
    lblSaldoRep = Format(lnSaldoNew, "#,##0.00")
End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    Dim oDCredito As New COMDCredito.DCOMCredito 'JUEZ 20131104
    Dim bCredito As Boolean
    Dim bTieneDiasAtraso As Boolean  'JUEZ 20131104
    If KeyAscii = 13 Then
        'JUEZ 20131114 ****************************************
        'Set oDCredito = New COMDCredito.DCOMCredito
        'If oDCredito.RecuperaDatosComunes(Me.ActxCta.NroCuenta).RecordCount <= 0 Then
        '    MsgBox "El crédito no existe, Favor de verificar", vbInformation, "Aviso"
        '    Exit Sub
        'End If
        ''bCredito = odCredito.VerificarVencimientoCredito(Me.ActxCta.NroCuenta)
        'bTieneDiasAtraso = IIf(oDCredito.VerificarVencimientoCredito(Me.ActxCta.NroCuenta) = False, True, False)
        'Set oDCredito = Nothing
        'If bCredito = True Then
        'If bTieneDiasAtraso Then
        '    If MsgBox("El credito tiene dias de atraso. Sólo podrá reprogramar créditos por motivo Emergencia Café (Roya Amarilla), Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        '        fbReprogDiasAtraso = True
        '    Else
        '        Call cmdNuevo_Click
        '        Exit Sub
        '    End If
        'End If
        'END JUEZ *********************************************
        If Len(ActxCta.NroCuenta) = 18 Then 'JUEZ 20130412
            If CargaDatos(ActxCta.NroCuenta) Then
                'JUEZ 20130412 ******************************************************************
                'Dim odCredito As COMDCredito.DCOMCredito 'JUEZ 20131104
                Set oDCredito = New COMDCredito.DCOMCredito
                If oDCredito.ExisteComisionVigente(ActxCta.NroCuenta, gComisionReprogCredito) Then
                    fnTipoComision = 1 'Pagado
                Else
                    Dim R As ADODB.Recordset
                    Dim lsPrdConceptoCod As Integer
                    Set oDCredito = New COMDCredito.DCOMCredito
                    Set R = oDCredito.RecuperaDatosComision(ActxCta.NroCuenta, 1)
                    fnPersoneria = R!nPersoneria
                    lsPrdConceptoCod = IIf(R!nPersoneria = 1, gColocConceptoCodGastoComisionReprogNat, gColocConceptoCodGastoComisionReprogJur)
                    
                    Set R = oDCredito.RecuperaProductoConcepto(lsPrdConceptoCod)
                    'JUEZ 20151229 ************************************************
                    Dim lnTCVenta As Double
                    Dim oDGeneral As COMDConstSistema.NCOMTipoCambio
                    Set oDGeneral = New COMDConstSistema.NCOMTipoCambio
                        lnTCVenta = oDGeneral.EmiteTipoCambio(gdFecSis, TCVenta)
                    Set oDGeneral = Nothing
                    lnValor = CDbl(R!nValor) / IIf(Mid(ActxCta.NroCuenta, 9, 1) = "1", 1, lnTCVenta)
                    'END JUEZ *****************************************************
                    'If MsgBox("Para reprogramar un crédito el cliente debe realizar un pago de S/." & Format(R!nValor, "#,##0.00") & " por concepto de Comisión por Reprogramación de Créditos; en caso de que se desee realizar esta operación sin el pago previo se cargará en la siguiente cuota y requerirá el VºBº del Jefe de Agencia. Para continuar presione Aceptar", vbOKCancel, "Aviso") = vbCancel Then
                    
                    'If MsgBox("Para reprogramar un crédito el cliente debe realizar un pago de " & IIf(Mid(ActxCta.NroCuenta, 9, 1) = "1", "S/ ", "$ ") & Format(lnValor, "#,##0.00") & " por concepto de Comisión por Reprogramación de Créditos; en caso de que se desee realizar esta operación sin el pago previo se cargará en la siguiente cuota y requerirá el VºBº del Jefe de Agencia. Para continuar presione Si", vbYesNo, "Aviso") = vbYes Then 'JUEZ 20151229 'JOEP
                    If MsgBox("Se aplicará el pago de " & IIf(Mid(ActxCta.NroCuenta, 9, 1) = "1", "S/ ", "$ ") & Format(lnValor, "#,##0.00") & " por concepto de Comisión de Reprogramación de Créditos, en la siguiente cuota." & Chr(10) & "Para continuar presione Si", vbYesNo, "Aviso") = vbYes Then 'JUEZ 20151229 'JOEP
                                                                        
                        fnTipoComision = 2 'Cobrar en proxima cuota
                        
                        'Comento JOEP20171214 ERS082-2017
'                        Set loVistoElectronico = New frmVistoElectronico
'                        lbVistoVal = loVistoElectronico.Inicio(2, "", lsPersCod)
'
'                        If lbVistoVal Then
'                            fnTipoComision = 2 'Cobrar en proxima cuota
'                        Else
'                            Call cmdCancelar_Click
'                            Exit Sub
'                        End If
                        'Comento JOEP20171214 ERS082-2017
                        
                    Else
                        Call cmdCancelar_Click
                        Exit Sub
                    End If
                    
                    
                End If
                'END JUEZ ***********************************************************************
                FraDatos.Enabled = False
                
                'cmdEditar.Enabled = True
                cmdReprogramar.Enabled = True
                
                'CmdGastos.Enabled = True 'DAOR 20070410
                HabilitarReprogramar True
                
                'joep20201002 Tasa especial y reduccion de monto
                Dim obOCM As COMDCredito.DCOMCredito
                Dim rsRepgOCM As ADODB.Recordset
                Set obOCM = New COMDCredito.DCOMCredito
                Set rsRepgOCM = obOCM.ReprogramacionOCM(ActxCta.NroCuenta)
                If Not (rsRepgOCM.BOF And rsRepgOCM.EOF) Then
                    Call OCMControl(rsRepgOCM!nModalidad)
                End If
                Set obOCM = Nothing
                RSClose rsRepgOCM
                'joep20201002 Tasa especial y reduccion de monto
            Else
                cmdCancelar_Click
            End If
        End If
    End If
End Sub

Private Sub Impresion()
Dim sCad As String
Dim i As Integer
Dim oPrev As previo.clsprevio

    sCad = Chr$(10)
    sCad = sCad & Space(40) & "Reprogramacion de Credito" & Chr$(10)
    sCad = sCad & Space(38) & String(30, "-") & Chr$(10)
    sCad = sCad & Chr$(10) & Chr$(10)
    sCad = sCad & Space(2) & "Credito : " & Me.ActxCta.NroCuenta
    sCad = sCad & Space(2) & "Titular : " & lblTitular.Caption & Chr$(10)
    sCad = sCad & Space(2) & "Analista : " & LblPrestamo.Caption & Chr$(10)
    sCad = sCad & Space(2) & "Saldo Capital: " & LblSaldo.Caption
    sCad = sCad & Space(2) & "Tasa : " & LblTasa.Caption & Chr$(10) & Chr$(10) & Chr$(10)
    
    sCad = sCad & Space(2) & "Justificacion : " & Chr$(10) & Chr$(10)
    sCad = sCad & Space(15) & Trim(Me.TxtGlosa.Text) & Chr$(10) & Chr$(10) & Chr$(10)
    sCad = sCad & Space(2) & "Calendario Nuevo : " & Chr$(10)
    
    sCad = sCad & Space(2) & ImpreFormat("Fecha", 10) & ImpreFormat("Nro", 3) & ImpreFormat("Monto", 7)
    sCad = sCad & Space(2) & ImpreFormat("Capital", 10) & ImpreFormat("Interes", 10) & ImpreFormat("Mora", 8)
    sCad = sCad & Space(2) & ImpreFormat("Int.Rep", 10) & ImpreFormat("Int.Gra", 10) & ImpreFormat("Gastos", 10) & ImpreFormat("Saldo", 10) & Chr$(10)
    sCad = sCad & Space(2) & String(110, "-") & Chr$(10)
    
    For i = 1 To FECalend.rows - 1
        sCad = sCad & Space(2) & ImpreFormat(FECalend.TextMatrix(i, 1), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 2), 3)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 3), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 4), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 5), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 6), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 7), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 8), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 9), 10)
        sCad = sCad & ImpreFormat(FECalend.TextMatrix(i, 10), 10) & IIf(MatCalend(i - 1, 8) = 1, "C", "P") & Chr$(10)
                
    Next i
    
    sCad = sCad & Space(2) & String(110, "-") & Chr$(10)
    
    Set oPrev = New previo.clsprevio
    oPrev.Show sCad, "Reprogramacion Credito", True
    Set oPrev = Nothing
    
    
End Sub

Private Sub CmdAceptar_Click()
Dim oNCredito As COMNCredito.NCOMCredito
Dim oDCredito As COMDCredito.DCOMCredito 'EJVG20121210
Dim i As Integer
Dim pnMovMro As Long
Dim lsOpeCod As String
Dim bReprogxDesNat As Boolean 'EJVG20121210
Dim bReprogRoyaAmarilla As Boolean 'JUEZ 20131022
Dim bReprogZonaInundada As Boolean 'PASI 20170327

    On Error GoTo ErrorCmdAceptar_Click
    
    'JUEZ 20131104 *************************************************************
'    If fbReprogDiasAtraso Then
'        If Not optReprogNatEspeciales(1).Enabled Or Not optReprogNatEspeciales(1).value Then
'            MsgBox "Sólo los crèditos con motivo Emergencia Café (Roya Amarilla) se puede reprogramar con dias de atraso", vbInformation, "Aviso"
'            Exit Sub
'        End If
'    End If
    'END JUEZ ******************************************************************
        
'    If dFecTemp = CDate(Me.FECalend.TextMatrix(FECalend.row, 1)) And Not OptTipoRep(2).value Then
'        MsgBox "La Fecha de Reprogramacion No Puede ser la misma", vbInformation, "Aviso"
'        Exit Sub
'    End If

'JOEP20200926 Tasa ESpecial Reduccion de Monto
    If ValidaDatos(1) = True Then
        Exit Sub
    End If
'JOEP20200926 Tasa ESpecial Reduccion de Monto
'Comento JOEP20200926 Tasa ESpecial Reduccion de Monto
'    If txtGlosa.Text = "" Then
'        MsgBox "Deber ingresar una glosa", vbInformation, "Aviso"
'        txtGlosa.SetFocus
'        Exit Sub
'    End If
'Comento JOEP20200926 Tasa ESpecial Reduccion de Monto
    
    'EJVG20121210 ***
    Set oDCredito = New COMDCredito.DCOMCredito
    'If chkDesastreNat.Enabled = True And chkDesastreNat.value = 1 Then
    'If optReprogNatEspeciales(2).Enabled = True And optReprogNatEspeciales(2).value Then 'JUEZ 20131022 'comento JOEP20200428 covid cuotas iguales
    If CmbReprogNatEspeciales.Visible = True And CmbReprogNatEspeciales.Enabled = False And Right(CmbReprogNatEspeciales, 1) = 2 Then 'Add JOEP20200428 covid cuotas iguales
        If oDCredito.RecuperaMovReprogxDesastreNatxMes(ActxCta.NroCuenta, Year(gdFecSis), Month(gdFecSis), gCredReprogDesastreNat).RecordCount > 0 Then
            MsgBox "El Crédito ya fue reprogramado por Desastre Natural o Conflicto Social en el presente mes." & Chr(10) & "Solo una vez al mes se podría reprogramar con esta condición", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    'JUEZ 20131022 ************************************************
    'If optReprogNatEspeciales(1).Enabled = True And optReprogNatEspeciales(1).value Then'comento JOEP20200428 covid cuotas iguales
    If CmbReprogNatEspeciales.Visible = True And CmbReprogNatEspeciales.Enabled = False And Right(CmbReprogNatEspeciales, 1) = 1 Then  'Add JOEP20200428 covid cuotas iguales
        If oDCredito.RecuperaMovReprogxDesastreNatxMes(ActxCta.NroCuenta, Year(gdFecSis), Month(gdFecSis), gCredReprogRoyaAmarilla).RecordCount > 0 Then
            'MsgBox "El Crédito ya fue reprogramado por Emergencia Café (Roya Amarilla) en el presente mes." & Chr(10) & "Solo una vez al mes se podría reprogramar con esta condición", vbInformation, "Aviso"
            MsgBox "El Crédito ya fue reprogramado por Covid-19 en el presente mes." & Chr(10) & "Solo una vez al mes se podría reprogramar con esta condición", vbInformation, "Aviso" 'JOEP 20200323 Covid-19
            Exit Sub
        End If
    End If
    'END JUEZ *****************************************************
    'PASI 20170327 ************************************************
    'If optReprogNatEspeciales(5).Enabled = True And optReprogNatEspeciales(5).value Then 'comento JOEP20200428 covid cuotas iguales
    If CmbReprogNatEspeciales.Visible = True And CmbReprogNatEspeciales.Enabled = False And Right(CmbReprogNatEspeciales, 1) = 5 Then 'Add JOEP20200428 covid cuotas iguales
        If oDCredito.RecuperaMovReprogxDesastreNatxMes(ActxCta.NroCuenta, Year(gdFecSis), Month(gdFecSis), gCredReprogZonaInundada).RecordCount > 0 Then
            MsgBox "El Crédito ya fue reprogramado por Zona Afectada por Inundación en el presente mes." & Chr(10) & "Solo una vez al mes se podría reprogramar con esta condición", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    'END PASI *****************************************************
    
    Set oDCredito = Nothing
    'END EJVG *******
    
    'Add By GITU 2008-08-25
    lsOpeCod = "100902"
    
'    Set loVistoElectronico = New frmVistoElectronico
'    If val(txtDiasporReprog.Text) > 29 Then
'        lbVistoVal = loVistoElectronico.Inicio(2, lsOpeCod, lsPersCod)
'
'        If Not lbVistoVal Then
'            If CDate(FECalend.TextMatrix(FECalend.Rows - 1, 1)) > CDate(lblfecUltCuota) + 29 Then
'                MsgBox "La Fecha de Vencimiento de la ultima cuota no puede ser mayor a : " & lblfecUltCuota & "Necesita autorizacion", vbInformation
'                'txtDiasporReprog.SetFocus
'                If Me.txtDiasporReprog.Enabled Then txtDiasporReprog.SetFocus 'EJVG20121211
'                cmdReprog.Enabled = False
'                Exit Sub
'            End If
'        End If
'    End If
    
    If CCur(lblSaldoRep) <> CCur(LblSaldo) Then
        'MsgBox "Saldo de Capital Original no coincide con Saldo Reprogramado. Por favor Verificar", vbInformation, "Aviso"
        'Austar para que cuadre al cartera
            Set oNCredito = New COMNCredito.NCOMCredito
            Call oNCredito.AjustaCalendarioACapital(MatCalend, CDbl(LblSaldo.Caption))
            Set oNCredito = Nothing
    End If
    If Not IsArray(MatCalend) Then
        MsgBox "Debe reprogramar el Calendario", vbInformation, "Mensaje"
        Call cmdCancelar_Click
        Exit Sub
    End If
       
    'EJVG20121210 ***
    Set oDCredito = New COMDCredito.DCOMCredito 'JUEZ 20131022
    'If chkDesastreNat.Enabled = True And chkDesastreNat.value = 1 Then
    'If optReprogNatEspeciales(2).Enabled And optReprogNatEspeciales(2).value Then 'JUEZ 20131022
    If CmbReprogNatEspeciales.Visible = True And CmbReprogNatEspeciales.Enabled = False And Right(CmbReprogNatEspeciales, 1) = 2 Then  'Add JOEP20200428 covid cuotas iguales
        bReprogxDesNat = True
    'End If
    'END EJVG *******
    'JUEZ 20131022 *****************************************************************
    'ElseIf optReprogNatEspeciales(1).Enabled And optReprogNatEspeciales(1).value Then 'Emergencia Café (Roya Amarilla)
    ElseIf CmbReprogNatEspeciales.Visible = True And CmbReprogNatEspeciales.Enabled = False And Right(CmbReprogNatEspeciales, 1) = 1 Then  'Add JOEP20200428 covid cuotas iguales
        bReprogRoyaAmarilla = True
    'ElseIf optReprogNatEspeciales(5).Enabled And optReprogNatEspeciales(5).value Then 'PASI 20170327
    ElseIf CmbReprogNatEspeciales.Visible = True And CmbReprogNatEspeciales.Enabled = False And Right(CmbReprogNatEspeciales, 1) = 5 Then 'PASI 20170327
        bReprogZonaInundada = True
    End If
    Set oDCredito = Nothing
    'END JUEZ **********************************************************************
    
    If MsgBox("Se va a Reprogramar el Credito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
        Set oNCredito = New COMNCredito.NCOMCredito
        'Call oNCredito.ReprogramarCredito(ActxCta.NroCuenta, MatCalend, nTipoReprogCred, , , gdFecSis, , gsCodUser, gsCodAge, nTipoPeriodo, nPlazo, True, pnMovMro, lnDiaFijoColocEstado, bReprogxDesNat, IIf(bReprogZonaInundada, False, True), fnTipoComision, fnPersoneria, bReprogRoyaAmarilla, , RTrim(LTrim(TxtGlosa.Text)), NewTCEA, bReprogZonaInundada, , val(Right(cmbCovidReprog.Text, 1)))     'JOEP202003223 comento Covid-19/ RIRO ADD val(Right(cmbCovidReprog.Text, 1))
        Call oNCredito.ReprogramarCredito(ActxCta.NroCuenta, MatCalend, nTipoReprogCred, , , gdFecSis, , gsCodUser, gsCodAge, nTipoPeriodo, nPlazo, True, pnMovMro, lnDiaFijoColocEstado, bReprogxDesNat, IIf(bReprogZonaInundada, False, True), fnTipoComision, fnPersoneria, bReprogRoyaAmarilla, , RTrim(LTrim(TxtGlosa.Text)), NewTCEA, bReprogZonaInundada, , val(Right(cmbCovidReprog.Text, 1)), 1, txtMontoCuota)    'RIRO 20200903 Se añadió parámetro nuevo
        
        Call oNCredito.EliminarCuotaSobranteRep(ActxCta.NroCuenta, CInt(FECalend.rows) - 1) 'LARI 2021-06-05 elimina las cuotas sobrantes
        'PASI 20170327 agrego bReprogZonaInundada
    Set oNCredito = Nothing
    
    MsgBox "- Se ha reprogramado el crédito. Proceda a generar su nuevo plan de pagos en la opción Duplicados" & Chr(13) & Chr(13) & _
           "- No olvidar entregar los documentos de la reprogramación a Administración de Créditos para la revisión respectiva", vbInformation, "Aviso"
    
    ''*** PEAC 20090126 descomentar gitu
    'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, , ActxCta.NroCuenta, gCodigoCuenta

    
    '*** PEAC 20080807
    
    'loVistoElectronico.RegistraVistoElectronico (pnMovMro)'Comento JOEP20171214 ERS082-2017
        
    '*** FIN PEAC
    
    'Call Impresion

    'Call cmdNuevo_Click
    Call cmdCancelar_Click
    
    Exit Sub

ErrorCmdAceptar_Click:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub cmdBuscar_Click()
Dim oCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim sPersCod As String
Dim oPers As COMDPersona.UCOMPersona

    On Error GoTo ErrorCmdBuscar_Click
    Set oPers = frmBuscaPersona.Inicio()
    If oPers Is Nothing Then
        'MsgBox "No se Selecciono Ninguna Persona", vbInformation, "Aviso"
        Exit Sub
    Else
        sPersCod = oPers.sPersCod
    End If
    Set oPers = Nothing
    
    Set oCred = New COMDCredito.DCOMCredito
    Set R = oCred.RecuperaCreditosVigentes(sPersCod, , Array(gColocEstVigMor, gColocEstVigNorm, gColocEstVigVenc, gColocEstRefMor, gColocEstRefNorm, gColocEstRefVenc))
    Set oCred = Nothing
    LstCtas.Clear
    If R.BOF And R.EOF Then
        MsgBox "No Existen Creditos Vigentes", vbInformation, "Aviso"
    Else
        Do While Not R.EOF
            LstCtas.AddItem R!cCtaCod
            R.MoveNext
        Loop
        LstCtas.SetFocus
    End If
    
    R.Close
    Set R = Nothing
    Exit Sub

ErrorCmdBuscar_Click:
        MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub cmdCancelar_Click()
'JUEZ 20160316 *****************************************
'    If nTipoReprogCred = 1 Then
'        FECalend.row = nFilaEditar
'        Call FECalend.BackColorRow(vbWhite)
'        nFilaEditar = -1
'    End If
'    bCalendGen = False ' DAOR 20070411
'    HabilitaControlesReprog False
'    nTipoReprogCred = -1
'    MatCalend = ""
'    txtDiasporReprog.Text = "" 'Add by Gitu 22-08-08
'    'JUEZ 20131022 *************************
'    'chkDesastreNat.value = 0 'EJVG20121211
'    'optReprogNatEspeciales(1).value = 0 'JUEZ 20140415 MEMO 903-2014
'    optReprogNatEspeciales(2).value = 0
'    'END JUEZ ******************************
'    cmdReprog.Enabled = False
'    Call ActxCta_KeyPress(13)
Call cmdNuevo_Click
'END JUEZ **********************************************

Call VisibleBotones(0, False) 'JOEP20200626 Mantener cuota Covid

End Sub

Private Sub cmdEditar_Click()
    If Trim(FECalend.TextMatrix(1, 1)) = "" Then
        Exit Sub
    End If
    If CInt(FECalend.TextMatrix(FECalend.row, 11)) = gColocCalendEstadoPagado Then
        MsgBox "No se Puede Reprogramar Cuotas Canceladas", vbInformation, "Aviso"
        Exit Sub
    End If
    nFilaEditar = FECalend.row
    dFecTemp = CDate(FECalend.TextMatrix(FECalend.row, 1))
    'Call FECalend.BackColorRow(&HC0FFFF)
    HabilitaControlesReprog True
    nTipoReprogCred = 1
    
    If Me.OptTipoRep(2).value Then
        Call FECalend_OnValidate(1, 1, False)
    End If
End Sub

'**DAOR 20070409, Permite obtener los gastos.
'**Función obtenida del formulario frmCredAprobacion
Private Sub CmdGastos_Click()
Dim oNGasto As COMNCredito.NCOMGasto
Dim bErrorCargaCalendario As Boolean 'ARCV

On Error GoTo ErrorCmdGastos_Click
    If Not bCalendGen Then
        MsgBox "Genere el Calendario de Pagos para poder Definir los Gastos", vbInformation, "Aviso"
        'CmdGenCal.SetFocus
        Exit Sub
    End If
    
    bGastosGen = True
    Set oNGasto = Nothing
    
    Call frmCredGastos.MuestraGastos(MatGastos, nNumGastos, True)
    Exit Sub

ErrorCmdGastos_Click:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub CmdGenCal_Click()
Dim i As Integer
Dim nInteresGraciaPend As Double 'DAOR 20071120

'   ARCV 24-10-2006
    Dim MatCalend_2 As Variant
    Dim nInteresAFecha As Double
    Dim oNCredito As COMNCredito.NCOMCredito
    Set oNCredito = New COMNCredito.NCOMCredito
    MatCalend_2 = oNCredito.RecuperaMatrizCalendarioPendiente(ActxCta.NroCuenta)

    nInteresAFecha = Format(oNCredito.MatrizInteresTotalesAFechaSinMora(ActxCta.NroCuenta, MatCalend_2, gdFecSis) + _
                        oNCredito.ObtenerMoraVencida(gdFecSis, MatCalend_2), "#0.00")

    Set oNCredito = Nothing
    '-----------------------

    nTipoReprogCred = 2
    
    'Comentado por DAOR 20070410
    'MatCalend = frmCredCalendPagos.Renovar(CDbl(Me.lblSaldoRep.Caption), gdFecSis, fnTasaInteres, nInteresAFecha)
    
    '**DAOR 20070410*****************************
    MatCalend = frmCredCalendPagos.Renovar(CDbl(Me.lblSaldoRep.Caption), gdFecSis, fnTasaInteres, nInteresAFecha, ActxCta.NroCuenta)
    bCalendGen = True
    Call CmdGastos_Click
    '********************************************
    If IsArray(MatCalend) Then
        HabilitaControlesReprog True
        LimpiaFlex FECalend

        For i = 0 To UBound(MatCalend) - 1
            FECalend.AdicionaFila
            FECalend.TextMatrix(i + 1, 1) = MatCalend(i, 0)
            FECalend.TextMatrix(i + 1, 2) = MatCalend(i, 1)
            FECalend.TextMatrix(i + 1, 3) = MatCalend(i, 2)
            FECalend.TextMatrix(i + 1, 4) = MatCalend(i, 3)
            FECalend.TextMatrix(i + 1, 5) = MatCalend(i, 4)
            FECalend.TextMatrix(i + 1, 6) = "0.00"
            FECalend.TextMatrix(i + 1, 7) = "0.00"
            FECalend.TextMatrix(i + 1, 8) = MatCalend(i, 5)
            FECalend.TextMatrix(i + 1, 9) = MatCalend(i, 6)
            FECalend.TextMatrix(i + 1, 10) = MatCalend(i, 7)
            FECalend.row = i + 1
            Call FECalend.ForeColorRow(vbBlack)
        Next i
    Else
    '    Call cmdCancelar_Click
    End If
        
    FECalend.lbEditarFlex = False

End Sub

Private Sub cmdNuevo_Click()
    nTipoReprogCred = -1
    FraDatos.Enabled = True
    LimpiaFlex FECalend
    CmdEditar.Enabled = False
    nFilaEditar = -1
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    cmdSalir.Enabled = True
    HabilitaControlesReprog False
    CmdEditar.Enabled = False
    'CmdReprogramar.Enabled = False
    LblAnalista.Caption = ""
    LblPrestamo.Caption = "0.00"
    LblSaldo.Caption = "0.00"
    lblTitular.Caption = ""
    LblTasa.Caption = "0.0000"
    Me.TxtGlosa.Text = ""
    MatCalend = ""
    lblSaldoRep = "0.00"
    'Add by gitu 13-08-08
    OptRepro(1).value = False
    txtDiasporReprog.Text = 0
    cmdReprog.Enabled = False
    Label7.Visible = False
    txtDiasporReprog.Visible = False
    'End Gitu
    LstCtas.Clear   'Se agrego
    'ALPA 20100606***************
    lblTipoCredito.Caption = ""
    lblTipoProducto.Caption = ""
    '****************************
    'JUEZ 20131022 *************************
    'chkDesastreNat.value = 0 'EJVG20121210
    'optReprogNatEspeciales(1).value = 0 'JUEZ 20140415 MEMO 903-2014
    'optReprogNatEspeciales(2).value = 0'comento JOEP20200428 covid cuotas iguales
    CmbReprogNatEspeciales.ListIndex = -1 'JOEP20200428 Covid cuota iguales
    cmbCovidReprog.Visible = False 'JOEP20200428 Covid cuota iguales
    CmbReprogNatEspeciales.Visible = False 'JOEP20200428 Covid cuota iguales
    'END JUEZ ******************************
    fbReprogDiasAtraso = False
    lblfecUltCuota.Caption = ""
    Me.lblDiasReprog.Caption = ""
    cmdReprogramar.Enabled = False
    HabilitarReprogramar True
    lnValor = 0
    'JOEP20200626 Mantener Cuota covid
    txtMontoCuota.Text = "0.00"
    txtTCEAReprog.Text = "0%"
    txtTEA.Text = "0%"
    txtTCEADes.Text = "0%"
    'JOEP20200626 Mantener Cuota covid
    bValidaActualizacionLiq = False
    nMontoCuota = 0
End Sub

Private Sub HabilitarReprogramar(ByVal pbHabilita As Boolean)
    cmdReprogramar.Visible = pbHabilita
    cmdAceptar.Visible = Not pbHabilita
    
    CmbReprogNatEspeciales.Enabled = pbHabilita 'JOEP20200728 covid mantener cuota
    cmbCovidReprog.Enabled = pbHabilita 'JOEP20200728 covid mantener cuota
    txtMontoCuota.Enabled = pbHabilita 'ALAN- JOEP20200925
End Sub

'->***** LUCV20180601, Comentó según ERS022-2018
'***************************************************'
'** Nuevo proceso de Reprogramacion GITU 15-08-08 **'
'***************************************************'
'Private Sub cmdReprog_Click()
'    Dim oCredito As COMDCredito.DCOMCredito
'    Dim oCred As COMNCredito.NCOMCredito
'    Dim oCon As COMConecta.DCOMConecta
'    Dim rsCal As ADODB.Recordset
'    Dim rsCred As ADODB.Recordset
'    Dim rsEst As ADODB.Recordset
'    Dim rsReprogApr As ADODB.Recordset, rsDatosReprog As ADODB.Recordset
'    Dim bAutorizacion As Boolean
'    Dim bCredito As Boolean
'    Dim nPrdEstado As Integer
'    Dim lnSaldoNew As Double
'    Dim nFec As Integer
'    Dim nDiaFijo As Integer
'    Dim nDiaTemp As Integer
'    Dim nMes As Integer
'    Dim nAnio As Integer
'    Dim dFechaReprog As Date
'    Dim nInteres As Double
'    Dim nInterespro As Double
'    Dim nSaldoCapital As Double
'    Dim i As Integer
'    Dim lnPerFechaFija As Integer
'    Dim lnSaldoMat As Double
'    Dim sqlFF As String
'    'WIOR 20160226 ***
'    Dim nNumMesesPorDia As Integer
'    Dim nNumMesesPorMes As Integer
'    Dim nGastoIncendio As Currency
'    Dim nGastoSegDesg As Currency
'    Dim nGastos As Currency
'    'WIOR FIN ********
'
'    'JOEP INICIO
'    Dim P As Integer
'    Dim nUltPago As Date
'    Dim nUltPagoNoComp As Date
'    Dim fnCantAfiliadosSegDes As Integer
'    Dim nTasaSegDes As Double
'    Dim nTotalCapital As Double
'    Dim nAjuste As Double
'    Dim nValCuoataAju As Double
'    Dim nCuotaAjustada As Double
'    Dim rsTipPeriodo As ADODB.Recordset
'    'JOEP FIN
'
'    lnPerFechaFijaAct = 0
'    lnDiaFijoColocEstado = 0
'
'    LimpiaFlex FECalend
'    MatCalend = ""
'    HabilitaControlesReprog True
'
'    lnSaldoNew = 0
'    nFec = 1
'    i = 0
'    P = 0
'    lnCapital = 0
'    lnIntComp = 0
'    lnIntGra = 0
'
'    '1. Obtiene datos de reprogramación
'    Set oCredito = New COMDCredito.DCOMCredito
'    Set oCred = New COMNCredito.NCOMCredito
'    Call oCredito.CargarDatosReprogramacion(ActxCta.NroCuenta, rsCred, rsCal, rsReprogApr, rsDatosReprog, bAutorizacion)
'    Set oCredito = Nothing
'
'    Set rsEst = New ADODB.Recordset
'    Set oCon = New COMConecta.DCOMConecta
'    oCon.AbreConexion
'
'    sqlFF = "Select nPeriodoFechaFija from ColocacEstado CE2 Where CE2.cCtaCod = '" & ActxCta.NroCuenta & "' And CE2.nPrdEstado In (" & COMDConstantes.gColocEstVigNorm & "," & COMDConstantes.gColocEstRefNorm & ")" ' Add By GITU 15-08-08"
'    Set rsEst = oCon.CargaRecordSet(sqlFF)
'    If Not rsEst.EOF Then
'        nPrdEstado = rsCred!nPrdEstado
'        fnTasaInteres = CDbl(Format(rsCred!nTasaInteres, "#0.00"))
'        ldVigencia = Format(rsCred!dVigencia, "dd/mm/yyyy")
'        nMontoApr = rsCal!nSaldoPactado
'        'nMontoApr = rsCred!nMontoCol 'Comentado JOEP
'        lnPerFechaFija = rsEst!nPeriodoFechaFija
'        nSaldoCapital = val(Replace(LblSaldo.Caption, ",", "")) 'JOEP 20170210
'    End If
'
'    'JOEP Identificar Tipo de Periodo
'    Set oCredito = New COMDCredito.DCOMCredito
'    Set rsTipPeriodo = oCredito.IdentificarTipoPeriodo(ActxCta.NroCuenta)
'    If Not (rsTipPeriodo.EOF And rsTipPeriodo.BOF) Then
'        nTipoPeriodo = rsTipPeriodo!nTpPeriodo
'    End If
'
'    'Inicio Tasa de Seguro Desgravamen JOEP
'    nTasaSegDes = oCred.ObtenerTasaSeguroDesg(ActxCta.NroCuenta, gdFecSis, fnCantAfiliadosSegDes)
'
'    'Inicio Identificar si la Gracia es Proratiada JOEP
'    Dim K As Integer
'    Dim nPrimeraIntGracia As Double
'    Dim nProrrateo As Double
'    Dim rsValidProrateado As ADODB.Recordset
'    Dim nInteresGraciaPendientes As Double
'    'Dim nInteresGraciaCuotaRepr As Double
'    'nProrrateo=0 --No es Prorateado
'    'nProrrateo=1 --Es Prorateado
'
'    Set rsValidProrateado = rsCal.Clone
'    For K = 1 To rsValidProrateado.RecordCount
'        If K = 1 Then
'            nPrimeraIntGracia = rsValidProrateado!nIntGracia
'        End If
'
'        If K <> (rsValidProrateado.RecordCount) Then
'            If (nPrimeraIntGracia) = (rsValidProrateado!nIntGracia) Then
'               nProrrateo = 1
'            Else
'               nProrrateo = 0
'            End If
'        End If
'
'        'Sumar Interes de Gracia de las Cuotas Pendientes
'        If rsValidProrateado!nColocCalendEstado = 0 Then
'            nInteresGraciaPendientes = nInteresGraciaPendientes + rsValidProrateado!nIntGracia
'        End If
'
'        rsValidProrateado.MoveNext
'    Next K
'     'Fin Identificar si la Gracia es Proratiada JOEP
'
'    'Generación del Calendario
'    ReDim MatCalend(nCuoPag, 11) 'WIOR 20160226
'    Do While Not rsCal.EOF
'        FECalend.AdicionaFila
'        nGastos = IIf(IsNull(rsCal!nGasto), 0, rsCal!nGasto)
'        MatCalend(i, 10) = 0 'nGastoIncendio
'        MatCalend(i, 11) = 0 'nGastoSegDesg
'
'        'Cuotas Pagadas
'        If rsCal!nColocCalendEstado = gColocCalendEstadoPagado Then
'            FECalend.TextMatrix(rsCal.Bookmark, 1) = Format(rsCal!dVenc, "dd/mm/yyyy") 'Fecha Venc.
'            MatCalend(i, 0) = Format(rsCal!dVenc, "dd/mm/yyyy") 'Fecha Venc.
'            lnCapital = rsCal!nCapital
'            lnIntComp = rsCal!nIntComp
'            lnIntGra = rsCal!nIntGracia
'            nUltPago = rsCal!dVenc 'Para Obtener el Ultimo Pago Completo fecha Venc
'        'Cuotas Pendientes
'        Else
'            If nFec = 1 Then
'                nMontoApr = LblSaldo 'Obtener el saldo a Reprogramar
'                'Inicio joep
'                Dim nDiaInteres As Double
'                Dim nDias As Integer
'                Dim nTotalDiasMes As Integer
'
'                Dim nNewCuotaComp As Double
'                Dim nNewCuotaCap As Double
'                Dim nNewCuotaGra As Double
'                Dim nInteresGracia As Double
'                Dim nInteresCompCapital As Double
'                Dim nIntDias As Double
'                Dim nSaldoGracia As Double
'
'                Dim nMontDias As Double
'                Dim nMontoDiasTotal As Double
'                Dim nFactDiario As Double
'                Dim nContDiasMes As Integer
'
'                Dim dFechaPago As Date
'                Dim nPagoInteres As Double
'                Dim nIntresCompPendi As Double
'                Dim nSaldoPendi As Double
'                Dim nPagRealizado As Double
'                Dim nPagNoRealizado As Double
'                Dim nInteresCompPendiente As Double
'
'                Dim dFecCont As Date
'                Dim lnIntGraAnt As Double
'                Dim lnSaldoIntGraAnt As Double
'
'                Dim oOpe As COMDCredito.DCOMCredito
'                Set oOpe = New COMDCredito.DCOMCredito
'                Dim rsDiasMes As ADODB.Recordset
'                Set rsDiasMes = New ADODB.Recordset
'                Dim rsValidaProrateo As ADODB.Recordset
'                Set rsValidaProrateo = New ADODB.Recordset
'
'                FECalend.TextMatrix(rsCal.Bookmark, 1) = Format(rsCal!dVenc + txtDiasporReprog.Text, "dd/mm/yyyy") 'Fecha Venc. (Considerando DiasReprog.)
'                dFechaReprog = rsCal!dVenc + txtDiasporReprog.Text
'
'                'INICIO Creditos a Una Solo Cuota - JOEP
'                If rsCal.RecordCount = 1 Then
'                    nPagRealizado = rsCal!nCapitalPag + rsCal!nIntCompPag + rsCal!nIntGraciaPag
'                    nPagNoRealizado = rsCal!nCapital + rsCal!nIntComp + rsCal!nIntGracia
'                    nInteresCompPendiente = rsCal!nIntComp
'
'                    'Verifica si hay Pago realizados.
'                    If (rsCal!nCapitalPag <> 0 Or rsCal!nIntCompPag <> 0 Or rsCal!nIntGraciaPag) <> 0 Then
'
'                        dFechaPago = Format(rsCal!dPago, "dd/mm/yyyy")
'                        nDias = (dFechaPago - ldVigencia)
'
'                        nFactDiario = ((1 + (fnTasaInteres / 100)) ^ (1 / 30) - 1) * 100 'Factor Diario
'                        nNewCuotaCap = (nSaldoCapital / (1 / (1 + (nFactDiario / 100)) ^ (nDias))) 'Nueva Cuota fecha Vencimiento
'
'                        nInteresCompCapital = (((1 + (fnTasaInteres / 100)) ^ ((nDias / 30)) - 1) * nSaldoCapital)
'
'                        nIntresCompPendi = nInteresCompCapital - rsCal!nIntCompPag
'                        nSaldoPendi = nSaldoCapital + nIntresCompPendi
'
'                        nDias = (dFechaReprog - dFechaPago)
'                        nNewCuotaCap = (nSaldoPendi / (1 / (1 + (nFactDiario / 100)) ^ (nDias))) 'Nueva Cuota fecha Reprogramacion
'
'                        lnIntComp = ((((1 + (fnTasaInteres / 100)) ^ ((nDias / 30)) - 1) * nSaldoPendi) + nIntresCompPendi)
'                        lnCapital = (Round(nNewCuotaCap, 2) - Round(lnIntComp, 2))
'
'                        'Seguro Desgravamen
'                        MatCalend(i, 0) = Format(rsCal!dVenc + val(txtDiasporReprog.Text), "dd/mm/yyyy") 'Fecha de Reprogramacion
'
'                        nGastoIncendio = 0
'                        nGastoSegDesg = 0
'
'                        'ldVigencia
'                        nGastoIncendio = IIf(IsNull(rsCal!nGastoPolizaIncendio), 0, rsCal!nGastoPolizaIncendio)
'                        nGastoSegDesg = IIf(IsNull(rsCal!nGastoSegDes), 0, rsCal!nGastoSegDes)
'                        nGastos = nGastos - nGastoIncendio - nGastoSegDesg
'
'                        If i = 0 Then
'                            nNumMesesPorDia = Round(DateDiff("d", dFechaPago, MatCalend(i, 0)) / 30, 0)
'                            nNumMesesPorMes = DateDiff("m", dFechaPago, MatCalend(i, 0))
'                        End If
'
'                        nGastoIncendio = nGastoIncendio * IIf(nNumMesesPorDia >= nNumMesesPorMes, nNumMesesPorDia, nNumMesesPorMes)
'                        nGastoSegDesg = (nSaldoCapital * (nTasaSegDes / 100) * IIf(nNumMesesPorDia >= nNumMesesPorMes, nNumMesesPorDia, nNumMesesPorMes)) 'Calculo Seguro Desgravamen
'                        nGastos = nGastos + nGastoIncendio + nGastoSegDesg
'
'                        MatCalend(i, 10) = nGastoIncendio
'                        MatCalend(i, 11) = nGastoSegDesg
'                    Else
'                        'No hizo pago
'                        MatCalend(i, 0) = Format(rsCal!dVenc + val(txtDiasporReprog.Text), "dd/mm/yyyy") 'Fecha de Reprogramacion
'
'                        nDias = (dFechaReprog - rsCal!dVenc) 'Dias Fecha Reprogramacion
'                        nInteresCompCapital = (((1 + (fnTasaInteres / 100)) ^ ((nDias / 30)) - 1)) * (nPagNoRealizado) + (nInteresCompPendiente) 'Interes compensatorio del Monto a Pagar (original)
'
'                        lnCapital = nSaldoCapital 'Capital (Mostrar en Formulario)
'                        lnIntComp = nInteresCompCapital 'Interes Compensatorio (Mostrar en Formulario)
'
'                        nGastoIncendio = 0
'                        nGastoSegDesg = 0
'
'                        nGastoIncendio = IIf(IsNull(rsCal!nGastoPolizaIncendio), 0, rsCal!nGastoPolizaIncendio) 'Seguro Incendio
'                        nGastoSegDesg = IIf(IsNull(rsCal!nGastoSegDes), 0, rsCal!nGastoSegDes) 'Seguro Desgravamen
'                        nGastos = nGastos - nGastoIncendio - nGastoSegDesg
'
'                        If i = 0 Then
'                            nNumMesesPorDia = Round(DateDiff("d", ldVigencia, MatCalend(i, 0)) / 30, 0) 'Dias para Seguro Desgravamen
'                            nNumMesesPorMes = DateDiff("m", ldVigencia, MatCalend(i, 0)) 'Mes para Seguro Desgravamen
'                        End If
'
'                        nGastoIncendio = nGastoIncendio * IIf(nNumMesesPorDia >= nNumMesesPorMes, nNumMesesPorDia, nNumMesesPorMes) 'Calculo Incendio
'                        nGastoSegDesg = (nSaldoCapital * (nTasaSegDes / 100) * IIf(nNumMesesPorDia >= nNumMesesPorMes, nNumMesesPorDia, nNumMesesPorMes)) 'Calculo Seguro Desgravamen
'                        nGastos = nGastos + nGastoIncendio + nGastoSegDesg 'Gastos (Mostrar en Formulario)
'
'
'                        MatCalend(i, 10) = nGastoIncendio
'                        MatCalend(i, 11) = nGastoSegDesg
'                    End If
'                'FIN Creditos a Una Solo Cuota - JOEP
'                Else
'                    'Asignar Fecha de Desembolso para la primera cuota si se va a reprogramar JOEP
'                    If ((rsCal!nCuota = 1) And (rsCal!nCapitalPag <> 0 Or rsCal!nIntCompPag <> 0 Or rsCal!nIntGraciaPag <> 0 Or rsCal!nGastoPag <> 0)) Or ((rsCal!nCuota = 1) And (rsCal!nCapitalPag = 0 Or rsCal!nIntCompPag = 0 Or rsCal!nIntGraciaPag = 0 Or rsCal!nGastoPag = 0)) Then
'                        nUltPago = ldVigencia
'                    End If
'
'                    'Asignar Fecha de Desembolso para la primera cuota si se va a reprogramar JOEP
'                    Set rsDiasMes = oOpe.ObtenerDiasMes(FECalend.TextMatrix(rsCal.Bookmark, 1))
'                    nTotalDiasMes = rsDiasMes!nDia
'                    nUltPagoNoComp = rsCal!dVenc 'Para Obtener el Ultimo No Pago Completo
'
'                    'Inicio si la Gracia es Proratiada JOEP
'                    If nProrrateo = 1 Then
'                        lnSaldoIntGraAnt = nInteresGraciaPendientes
'                    Else
'                        lnSaldoIntGraAnt = fnIntGraciaPend 'nInteresGraciaCuotaRepr
'                    End If
'                    'Fin si la Gracia es Proratiada JOEP
'
'                    'INICIO Plazo de Gracia en Dias
'                    If lnPerFechaFija > 0 Then
'                        nDias = ((dFechaReprog - 30) - nUltPago)
'                    Else
'                        'nDias = ((dFechaReprog - 30) - nUltPago)--Joep
'                        nDias = ((dFechaReprog - (nUltPagoNoComp - nUltPago)) - nUltPago)
'                    End If
'                    'FIN Plazo de Gracia en Dias
'
'                    lnIntGraAnt = oCred.MontoIntPerDias(fnTasaInteres, nDias, lnSaldoIntGraAnt)
'                    nInteresGracia = oCred.MontoIntPerDias(fnTasaInteres, nDias, nSaldoCapital) + lnIntGraAnt + lnSaldoIntGraAnt
'
'                    'INICIO Si la Fecha de la cuota que se va a reprogramar cae Febrero "nUltPagoNoComp" - Nueva Cuota Gracia y Interes Gracia
'                    If lnPerFechaFija > 0 Then
'                        nNewCuotaGra = (((1 + fnTasaInteres / 100) ^ (30 / 30)) - 1) * nInteresGracia
'                        nInteresCompCapital = nSaldoCapital * ((1 + (fnTasaInteres / 100)) ^ (30 / 30) - 1)
'                    Else
'                        nNewCuotaGra = (((1 + fnTasaInteres / 100) ^ ((nUltPagoNoComp - nUltPago) / 30)) - 1) * nInteresGracia
'                        nInteresCompCapital = nSaldoCapital * ((1 + (fnTasaInteres / 100)) ^ ((nUltPagoNoComp - nUltPago) / 30) - 1)
'                    End If
'                    'FIN Si la Fecha de la cuota que se va a reprogramar cae Febrero "nUltPagoNoComp" - Nueva Cuota Gracia y Interes Gracia
'
'                    'INICIO Para Calcular la Cuota del Capital - JOEP
'                    nMontoDiasTotal = 0
'                    nMontDias = 0
'                    nContDiasMes = 0
'                    nFactDiario = ((1 + (fnTasaInteres / 100)) ^ (1 / 30) - 1) * 100 'Factor Diario
'
'                    For P = 1 To nCuoNoPag
'                        If P = 1 Then
'                            'Calculo solo para la primera cuota par los dias de gracia
'                             If lnPerFechaFija > 0 Then
'                                 nContDiasMes = nContDiasMes + 30
'                             Else
'                                 nContDiasMes = nContDiasMes + (nUltPagoNoComp - nUltPago)
'                             End If
'                         'Calculo solo para la primera cuota par los dias de gracia
'                         Else
'                             nContDiasMes = nContDiasMes + nTotalDiasMes
'                         End If
'
'                         nMontDias = (1 / (1 + (nFactDiario / 100)) ^ (nContDiasMes))
'                         nMontoDiasTotal = nMontoDiasTotal + nMontDias
'
'                         nDiaTemp = Day(rsCal!dVenc + txtDiasporReprog.Text)
'                         lnPerFechaFijaAct = nDiaTemp
'
'                         If lnPerFechaFija > 0 Then
'                             nDiaFijo = nDiaTemp
'                             nMes = Month(dFechaReprog)
'                             nAnio = Year(dFechaReprog)
'
'                             nMes = nMes + 1
'
'                             If nMes > 12 Then
'                                 nAnio = nAnio + 1
'                                 nMes = 1
'                             End If
'
'                             If nMes = 2 Then
'                                 If nDiaFijo > 28 Then
'                                     If nAnio Mod 4 = 0 Then
'                                         nDiaFijo = 29
'                                     Else
'                                         nDiaFijo = 28
'                                     End If
'                                 End If
'                             Else
'                                 If nDiaFijo > 30 Then
'                                     If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
'                                         nDiaFijo = 30
'                                     End If
'                                 End If
'                             End If
'
'                             If nDiaFijo <> 0 And lnDiaFijoColocEstado = 0 Then
'                                 lnDiaFijoColocEstado = nDiaFijo
'                             End If
'
'                             dFechaReprog = CDate(Right("0" & Trim(str(nDiaFijo)), 2) & "/" & Right("0" & Trim(str(nMes)), 2) & "/" & Trim(str(nAnio)))
'                             dFecCont = Format(dFechaReprog, "dd/mm/yyyy")
'
'                             Set rsDiasMes = oOpe.ObtenerDiasMes(dFecCont)
'                             nTotalDiasMes = rsDiasMes!nDia
'
'                         Else
'                             If P = 1 Then
'                                 MatCalend(i, 0) = Format(rsCal!dVenc + txtDiasporReprog.Text, "dd/mm/yyyy")
'                             End If
'
'                             nTotalDiasMes = (dFechaReprog - ((dFechaReprog - (nUltPagoNoComp - nUltPago))))
'                             dFechaReprog = dFechaReprog + ((nUltPagoNoComp - nUltPago))
'                         End If
'                    Next P
'                    nNewCuotaCap = lblSaldoRep / nMontoDiasTotal
'                    'FIN Para Calcular la Cuota del Capital - JOEP
'
'                    nNewCuotaComp = nInteresGracia / nMontoDiasTotal
'
'                    lnCapital = (nNewCuotaCap - nInteresCompCapital) ' - rsCal!nCapitalPag
'                    lnIntGra = (nNewCuotaComp - nNewCuotaGra) '- rsCal!nIntGraciaPag
'                    lnIntComp = (nInteresCompCapital + nNewCuotaGra) '- rsCal!nIntCompPag
'                    'Fin joep
'
'                    'FECalend.TextMatrix(rsCal.Bookmark, 1) = Format(rsCal!dVenc + txtDiasporReprog.Text, "dd/mm/yyyy")'comentadp por joep
'                    MatCalend(i, 0) = Format(rsCal!dVenc + val(txtDiasporReprog.Text), "dd/mm/yyyy")
'                    dFechaReprog = rsCal!dVenc + txtDiasporReprog.Text
'                    'nInteres = oCred.MontoIntPerDias(fnTasaInteres, DateDiff("d", rsCal!dVenc, dFechaReprog), rsCal!nCapital + rsCal!nIntComp)
'                    'nInteres = oCred.MontoIntPerDias(fnTasaInteres, DateDiff("d", rsCal!dVenc, dFechaReprog), nSaldoCapital)
'                    'nInteres = oCred.MontoIntPerDias(fnTasaInteres, 16, nSaldoCapital)'comentadp por joep
'                    'nInterespro = Round(nInteres / nCuoNoPag, 2) 'comentadp por joep
'
'                    nDiaTemp = Day(rsCal!dVenc + txtDiasporReprog.Text)
'                    lnPerFechaFijaAct = nDiaTemp
'
'                    'WIOR 20160226 ***
'                    nGastoIncendio = 0
'                    nGastoSegDesg = 0
'
'                    'ldVigencia
'                    nGastoIncendio = IIf(IsNull(rsCal!nGastoPolizaIncendio), 0, rsCal!nGastoPolizaIncendio)
'                    nGastoSegDesg = IIf(IsNull(rsCal!nGastoSegDes), 0, rsCal!nGastoSegDes)
'                    nGastos = nGastos - nGastoIncendio - nGastoSegDesg
'
'                    If i = 0 Then
'                        nNumMesesPorDia = Round(DateDiff("d", ldVigencia, MatCalend(i, 0)) / 30, 0)
'                        nNumMesesPorMes = DateDiff("m", ldVigencia, MatCalend(i, 0))
'                    Else
'                        nNumMesesPorDia = Round(DateDiff("d", MatCalend(i - 1, 0), MatCalend(i, 0)) / 30, 0)
'                        nNumMesesPorMes = DateDiff("m", MatCalend(i - 1, 0), MatCalend(i, 0))
'                    End If
'
'                    nGastoIncendio = nGastoIncendio * IIf(nNumMesesPorDia >= nNumMesesPorMes, nNumMesesPorDia, nNumMesesPorMes)
'                    'nGastoSegDesg = nGastoSegDesg * IIf(nNumMesesPorDia >= nNumMesesPorMes, nNumMesesPorDia, nNumMesesPorMes)
'                    nGastoSegDesg = (nSaldoCapital * (nTasaSegDes / 100) * IIf(nNumMesesPorDia >= nNumMesesPorMes, nNumMesesPorDia, nNumMesesPorMes)) 'Calculo Seguro Desgravamen
'                    nGastos = nGastos + nGastoIncendio + nGastoSegDesg
'
'                    MatCalend(i, 10) = nGastoIncendio
'                    MatCalend(i, 11) = nGastoSegDesg
'                    'WIOR FIN ********
'                    nValCuoataAju = i 'JOEP para la cuota que se va ajustar
'                    nSaldoGracia = nInteresGracia  'joep
'                    End If
'            Else 'Fecha (2)
'                If lnPerFechaFija > 0 Then 'Preguntamos Si es a fecha Fija
'                    If lnPerFechaFija > 0 Then
'                        nDiaFijo = nDiaTemp
'                        nMes = Month(dFechaReprog)
'                        nAnio = Year(dFechaReprog)
'
'                        nMes = nMes + 1
'                        If nMes > 12 Then
'                            nAnio = nAnio + 1
'                            nMes = 1
'                        End If
'
'                        If nMes = 2 Then
'                            If nDiaFijo > 28 Then
'                                If nAnio Mod 4 = 0 Then
'                                    nDiaFijo = 29
'                                Else
'                                    nDiaFijo = 28
'                                End If
'                            End If
'                        Else
'                            If nDiaFijo > 30 Then
'                                If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
'                                    nDiaFijo = 30
'                                End If
'                            End If
'                        End If
'                        If nDiaFijo <> 0 And lnDiaFijoColocEstado = 0 Then
'                            lnDiaFijoColocEstado = nDiaFijo
'                        End If
'
'                        dFechaReprog = CDate(Right("0" & Trim(str(nDiaFijo)), 2) & "/" & Right("0" & Trim(str(nMes)), 2) & "/" & Trim(str(nAnio)))
'                        FECalend.TextMatrix(rsCal.Bookmark, 1) = Format(dFechaReprog, "dd/mm/yyyy")
'                        MatCalend(i, 0) = Format(dFechaReprog, "dd/mm/yyyy")
'
'                    Else
'                        FECalend.TextMatrix(rsCal.Bookmark, 1) = Format(rsCal!dVenc + txtDiasporReprog.Text, "dd/mm/yyyy")
'                        MatCalend(i, 0) = Format(rsCal!dVenc + txtDiasporReprog.Text, "dd/mm/yyyy")
'                    End If
'                Else
'                    'Cambio2
'                    dFechaReprog = dFechaReprog + ((nUltPagoNoComp - nUltPago))
'                    FECalend.TextMatrix(rsCal.Bookmark, 1) = Format(dFechaReprog, "dd/mm/yyyy")
'                    MatCalend(i, 0) = dFechaReprog
'                End If
'
'            End If
'
'            nFec = nFec + 1
'
'            'Inicio JOEP 'Fecha (3)
'            If nFec <> 2 Then
'                'Fin Para obtener el Saldo de Gracia que va quedando JOEP
'                nSaldoGracia = nSaldoGracia - (nNewCuotaComp - nNewCuotaGra)
'                'Inicio Para obtener el Saldo de Gracia que va quedando JOEP
'                If lnPerFechaFija > 0 Then
'                    Set rsDiasMes = oOpe.ObtenerDiasMes(FECalend.TextMatrix(rsCal.Bookmark, 1))
'                    nTotalDiasMes = rsDiasMes!nDia
'                    nInteresCompCapital = nMontoApr * ((1 + (fnTasaInteres / 100)) ^ (nTotalDiasMes / 30) - 1)
'                    nNewCuotaGra = (((1 + fnTasaInteres / 100) ^ (nTotalDiasMes / 30)) - 1) * nSaldoGracia
'
'                    lnCapital = nNewCuotaCap - nInteresCompCapital
'                    lnIntGra = nNewCuotaComp - nNewCuotaGra
'                    lnIntComp = nInteresCompCapital + nNewCuotaGra
'                    nGastos = (nMontoApr * (nTasaSegDes / 100)) + IIf(IsNull(rsCal!nGastoPolizaIncendio), 0, rsCal!nGastoPolizaIncendio)
'
'                    MatCalend(i, 10) = IIf(IsNull(rsCal!nGastoPolizaIncendio), 0, rsCal!nGastoPolizaIncendio)
'                    MatCalend(i, 11) = (nMontoApr * (nTasaSegDes / 100))
'                Else
'                    'Cambio4
'                    nTotalDiasMes = (dFechaReprog - ((dFechaReprog - (nUltPagoNoComp - nUltPago))))
'                    nInteresCompCapital = nMontoApr * ((1 + (fnTasaInteres / 100)) ^ (nTotalDiasMes / 30) - 1)
'                    nNewCuotaGra = (((1 + fnTasaInteres / 100) ^ (nTotalDiasMes / 30)) - 1) * nSaldoGracia
'
'                    lnCapital = nNewCuotaCap - nInteresCompCapital
'                    lnIntGra = nNewCuotaComp - nNewCuotaGra
'                    lnIntComp = nInteresCompCapital + nNewCuotaGra
'                    nGastos = (nMontoApr * (nTasaSegDes / 100)) + IIf(IsNull(rsCal!nGastoPolizaIncendio), 0, rsCal!nGastoPolizaIncendio)
'
'                    MatCalend(i, 10) = IIf(IsNull(rsCal!nGastoPolizaIncendio), 0, rsCal!nGastoPolizaIncendio)
'                    MatCalend(i, 11) = (nMontoApr * (nTasaSegDes / 100))
'                End If
'            End If
'            'Fin JOEP
'        End If
'
'        'Asignamos valores en el Flex del calendario
'        FECalend.TextMatrix(rsCal.Bookmark, 2) = Trim(str(rsCal!nCuota))                            'NroCuota
'        FECalend.TextMatrix(rsCal.Bookmark, 3) = Format(IIf(IsNull(lnCapital), 0, lnCapital) + _
'                                                IIf(IsNull(lnIntComp), 0, lnIntComp) + _
'                                                IIf(IsNull(lnIntGra), 0, lnIntGra) + _
'                                                IIf(IsNull(rsCal!nIntMor), 0, rsCal!nIntMor) + _
'                                                IIf(IsNull(rsCal!nIntReprog), 0, rsCal!nIntReprog) + _
'                                                nGastos, "#0.00") 'Importe Cuota
'        FECalend.TextMatrix(rsCal.Bookmark, 4) = Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00") 'Capital
'        FECalend.TextMatrix(rsCal.Bookmark, 5) = Format(IIf(IsNull(lnIntComp), 0, lnIntComp), "#0.00") 'Interés Compensatorio
'        FECalend.TextMatrix(rsCal.Bookmark, 6) = Format(IIf(IsNull(rsCal!nIntMor), 0, rsCal!nIntMor), "#0.00") 'Interés Moratorio
'        FECalend.TextMatrix(rsCal.Bookmark, 7) = Format(IIf(IsNull(rsCal!nIntReprog), 0, rsCal!nIntReprog), "#0.00")
'        FECalend.TextMatrix(rsCal.Bookmark, 8) = Format(IIf(IsNull(lnIntGra), 0, lnIntGra), "#0.00")
'        FECalend.TextMatrix(rsCal.Bookmark, 9) = Format(nGastos, "#0.00") 'WIOR 20160226
'        nMontoApr = nMontoApr - IIf(IsNull(lnCapital), 0, lnCapital)
'        nMontoApr = CDbl(Format(nMontoApr, "#0.0000"))
'        FECalend.TextMatrix(rsCal.Bookmark, 10) = Format(nMontoApr, "#0.00")
'        FECalend.TextMatrix(rsCal.Bookmark, 11) = Trim(str(rsCal!nColocCalendEstado))
'        FECalend.TextMatrix(rsCal.Bookmark, 12) = Format(IIf(IsNull(rsCal!nCapitalPag), 0, rsCal!nCapitalPag), "#0.00")
'
'        'lnSaldoNew = lnSaldoNew + IIf(IsNull(rsCal!nCapital), 0, rsCal!nCapital) - IIf(IsNull(rsCal!nCapitalPag), 0, rsCal!nCapitalPag)'cOMENTADO joep
'        lnSaldoNew = lnSaldoNew + IIf(IsNull(lnCapital), 0, lnCapital) - IIf(IsNull(rsCal!nCapitalPag), 0, rsCal!nCapitalPag)
'
'        'Asignación de valores en Matriz del Calendario
'        MatCalend(i, 1) = Trim(str(rsCal!nCuota))
'        MatCalend(i, 2) = Format(IIf(IsNull(lnCapital), 0, lnCapital) + _
'                                        IIf(IsNull(lnIntComp), 0, lnIntComp) + _
'                                        IIf(IsNull(lnIntGra), 0, lnIntGra) + _
'                                        IIf(IsNull(rsCal!nIntMor), 0, rsCal!nIntMor) + _
'                                        IIf(IsNull(rsCal!nIntReprog), 0, rsCal!nIntReprog) + _
'                                        nGastos, "#0.00")
'                                        'IIf(IsNull(rsCal!nGasto), 0, rsCal!nGasto), "#0.00") 'WIOR 20160226 COMENTADO
'                                        'WIOR 20160226 AGREGO nGastos
'
'        MatCalend(i, 3) = Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")
'        MatCalend(i, 4) = Format(IIf(IsNull(lnIntComp), 0, lnIntComp), "#0.00")
'        MatCalend(i, 5) = Format(IIf(IsNull(lnIntGra), 0, lnIntGra), "#0.00")
'        MatCalend(i, 6) = Format(nGastos, "#0.00") 'WIOR 20160226
'        MatCalend(i, 7) = Format(lnSaldoNew, "#0.00")
'        MatCalend(i, 8) = rsCal!nColocCalendEstado
'        'MAVM 20100407 ***
'        MatCalend(i, 9) = rsCal!nIntCompPag + lnIntComp 'nInteres1 + nInterespro
'        ' ***
'
'        If rsCal!nColocCalendEstado = gColocCalendEstadoPagado Then
'            FECalend.row = rsCal.Bookmark
'            Call FECalend.ForeColorRow(vbRed)
'        End If
'
'        If rsCal.RecordCount = rsCal.Bookmark Then
'            lblfecUltCuota = Format(rsCal!dVenc, "dd/mm/yyyy")
'        End If
'        nInteres = 0
'        i = i + 1
'        rsCal.MoveNext
'
'        nTotalCapital = nTotalCapital + Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00") 'JOEP total capital
'    Loop
'
'    'JOEP Ajuste para que cuadre el capital
'    If LblPrestamo <> nTotalCapital Then
'        nAjuste = nTotalCapital - LblPrestamo
'        nCuotaAjustada = MatCalend(nValCuoataAju, 3) - nAjuste
'        MatCalend(nValCuoataAju, 3) = Format(nCuotaAjustada, "#0.00")
'    End If
'    'JOEP Ajuste para que cuadre el capital
'
'    NewTCEA = oCred.GeneraTasaCostoEfectivoAnual(CDate(Format(ldVigencia, "dd/mm/yyyy")), CDbl(LblPrestamo), MatCalend, CDbl(fnTasaInteres), ActxCta.NroCuenta, nTipoPeriodo)  'Para calcular la TCEA
'
'    nInterespro = 0
'    rsCal.Close
'    Set rsCal = Nothing
'    lblSaldoRep = Format(lnSaldoNew, "#,##0.00")
'End Sub
'<-***** Fin LUCV20180601

Private Sub CmdReprogramar_Click()
'Dim oCredito As COMNCredito.NCOMCredito
'Dim i As Integer
'
'    nTipoReprogCred = 2
'    Set oCredito = New COMNCredito.NCOMCredito
'    MatCalend = oCredito.ReprogramarCreditoenMemoriaTotal(ActxCta.NroCuenta, gdFecSis, IIf(OptRenov(0).value, True, False))
'    Set oCredito = Nothing
'
'    HabilitaControlesReprog True
'
'    LimpiaFlex FECalend
'    For i = 0 To UBound(MatCalend) - 1
'        FECalend.AdicionaFila
'        FECalend.TextMatrix(i + 1, 1) = MatCalend(i, 0)
'        FECalend.TextMatrix(i + 1, 2) = MatCalend(i, 1)
'        FECalend.TextMatrix(i + 1, 3) = MatCalend(i, 2)
'        FECalend.TextMatrix(i + 1, 4) = MatCalend(i, 3)
'        FECalend.TextMatrix(i + 1, 5) = MatCalend(i, 4)
'        FECalend.TextMatrix(i + 1, 6) = "0.00"
'        FECalend.TextMatrix(i + 1, 7) = "0.00"
'        FECalend.TextMatrix(i + 1, 8) = MatCalend(i, 5)
'        FECalend.TextMatrix(i + 1, 9) = MatCalend(i, 6)
'        FECalend.TextMatrix(i + 1, 10) = MatCalend(i, 7)
'        FECalend.row = i + 1
'        Call FECalend.ForeColorRow(vbBlack)
'    Next i
's
'    FECalend.lbEditarFlex = False
'
'

'cmdReprog_Click 'comento JOEP20200428 covid cuotas iguales

If ValidaDatos(0) = True Then
    Exit Sub
End If

'If CmbReprogNatEspeciales.Visible = True Then
'    If Right(cmbCovidReprog.Text, 1) = 2 Or Right(cmbCovidReprog.Text, 1) = 4 Then
'        If txtMontoCuota.Text = "" Or txtMontoCuota.Text = "0" Or txtMontoCuota.Text = "0.00" Then
'            MsgBox "Ingrese el monto de la cuota", vbInformation, "Aviso"
'            txtMontoCuota.SetFocus
'            Exit Sub
'        ElseIf txtMontoCuota > nMontoCuota Then
'            MsgBox "La cuota ingresada tiene que ser menor a la cuota original " & Format(nMontoCuota, "#,#0.00"), vbInformation, "Aviso"
'            txtMontoCuota.SetFocus
'            Exit Sub
'        End If
'    ElseIf Right(cmbCovidReprog.Text, 1) = 3 Then
'        If CCur(lbl_TasaEspecial) = 0 Then
'            MsgBox "No existe configuracion para la [Tasa Especial] ", vbInformation, "Aviso"
'            Exit Sub
'        End If
'        If CCur(lbl_TasaEspecial) >= CCur(LblTasa) Then
'            MsgBox "La [Tasa Especial] tiene que ser menor a la [Tasa Interes]", vbInformation, "Aviso"
'            Exit Sub
'        End If
'    End If
'End If

'Add JOEP20200428 covid cuotas iguales
If CmbReprogNatEspeciales.Visible = True And (CmbReprogNatEspeciales.Text) = "" Then
    MsgBox "Seleccione Opción de reprogramación", vbInformation, "Aviso"
    CmbReprogNatEspeciales.SetFocus
    Exit Sub
End If
If cmbCovidReprog.Visible = True And (cmbCovidReprog.Text) = "" Then
    MsgBox "Seleccione Opción de reprogramación", vbInformation, "Aviso"
    cmbCovidReprog.SetFocus
    Exit Sub
End If



If CmbReprogNatEspeciales.Visible = True And Right(CmbReprogNatEspeciales.Text, 1) = 1 Then
    'Normal=1,Mnatener cuota=2,Tasa especial=3
    If cmbCovidReprog.Visible = True And (Right(cmbCovidReprog.Text, 1) = 1 Or Right(cmbCovidReprog.Text, 1) = 3) Then
        cmdReprog_Click
    ElseIf cmbCovidReprog.Visible = True And (Right(cmbCovidReprog.Text, 1) = 2 Or Right(cmbCovidReprog.Text, 1) = 4) Then
        CovidCuotasIguales
    End If
Else
    cmdReprog_Click
End If

EnableBotones 0, False

'Add JOEP20200428 covid cuotas iguales

HabilitarReprogramar False
If cmdAceptar.Visible And cmdAceptar.Enabled Then cmdAceptar.SetFocus
'optReprogNatEspeciales(2).value = 0 'comento JOEP20200428 covid cuotas iguales
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub FECalend_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim oNCredito As COMNCredito.NCOMCredito
    
    'On Error GoTo ErrorFECalend_OnValidate
    
    If CDbl(FECalend.TextMatrix(nFilaEditar, 3)) < CDbl(FECalend.TextMatrix(nFilaEditar, 5)) + CDbl(FECalend.TextMatrix(nFilaEditar, 6)) + CDbl(FECalend.TextMatrix(nFilaEditar, 7)) + CDbl(FECalend.TextMatrix(nFilaEditar, 8)) + CDbl(FECalend.TextMatrix(nFilaEditar, 9)) Then
        MsgBox "Monto de Cuota debe ser mayor a la suma de los intereses ", vbInformation, "Aviso"
        Exit Sub
    End If
    If OptTipoRep(3).value = False Then
        If Format(CDbl(FECalend.TextMatrix(nFilaEditar, 3)), "#.00") > Format(CDbl(FECalend.TextMatrix(nFilaEditar, 4)) + CDbl(FECalend.TextMatrix(nFilaEditar, 5)) + CDbl(FECalend.TextMatrix(nFilaEditar, 6)) + CDbl(FECalend.TextMatrix(nFilaEditar, 7)) + CDbl(FECalend.TextMatrix(nFilaEditar, 8)) + CDbl(FECalend.TextMatrix(nFilaEditar, 9)), "#.00") Then
            MsgBox "Monto de Cuota debe ser menor a la cuota anterior", vbInformation, "Aviso"
            FECalend.TextMatrix(nFilaEditar, 3) = Format(CDbl(FECalend.TextMatrix(nFilaEditar, 4)) + CDbl(FECalend.TextMatrix(nFilaEditar, 5)) + CDbl(FECalend.TextMatrix(nFilaEditar, 6)) + CDbl(FECalend.TextMatrix(nFilaEditar, 7)) + CDbl(FECalend.TextMatrix(nFilaEditar, 8)) + CDbl(FECalend.TextMatrix(nFilaEditar, 9)), "#0.00")
            Exit Sub
        End If
    End If
    
    If OptTipoRep(3).value = False Then
        If CDate(FECalend.TextMatrix(nFilaEditar, 1)) < dFecTemp And Not Me.OptTipoRep(2).value Then
            MsgBox "La Fecha de Reprogramación debe ser Mayor a la Anterior", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        End If
    End If
    
    If CDate(FECalend.TextMatrix(nFilaEditar, 1)) > CDate(lblfecUltCuota) And FECalend.row > 1 Then 'ARCV 17-07
        MsgBox "La Fecha de Reprogramación no puede ser Mayor al vencimiento de la ultima cuota", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
    If CDate(FECalend.TextMatrix(nFilaEditar, 1)) < gdFecSis Then
        MsgBox "Fecha no puede ser menor a la fecha actual del sistema", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
    
    Set oNCredito = New COMNCredito.NCOMCredito
    If OptTipoRep(0).value Then
        MatCalend = oNCredito.ReprogramarCreditoenMemoria(ActxCta.NroCuenta, fnTasaInteres, dFecTemp, CDate(FECalend.TextMatrix(nFilaEditar, 1)), nFilaEditar - 1, 1, True, MatCalend)
    Else
        If OptTipoRep(1).value Then
            MatCalend = oNCredito.ReprogramarCreditoenMemoria(ActxCta.NroCuenta, fnTasaInteres, dFecTemp, CDate(FECalend.TextMatrix(nFilaEditar, 1)), nFilaEditar - 1, 2, False, MatCalend)
        Else
            If OptTipoRep(2).value Then
                MatCalend = oNCredito.ReprogramarCreditoenMemoria(ActxCta.NroCuenta, fnTasaInteres, dFecTemp, CDate(FECalend.TextMatrix(nFilaEditar, 1)), nFilaEditar - 1, 3, False, MatCalend)
            Else
                If val(FECalend.TextMatrix(nFilaEditar, 12)) > 0 Then
                    MsgBox "Cuota posee pagado capital no podrá continuar", vbInformation, "Aviso"
                    Cancel = False
                    Exit Sub
                End If
                
                Dim lnMonto As Double
                If (CDbl(FECalend.TextMatrix(nFilaEditar, 3)) <> CDbl(FECalend.TextMatrix(nFilaEditar, 4)) + CDbl(FECalend.TextMatrix(nFilaEditar, 5)) + CDbl(FECalend.TextMatrix(nFilaEditar, 6)) + CDbl(FECalend.TextMatrix(nFilaEditar, 7)) + CDbl(FECalend.TextMatrix(nFilaEditar, 8)) + CDbl(FECalend.TextMatrix(nFilaEditar, 9))) And (nFilaEditar <> FECalend.rows - 1) Then
                    'MatCalend = oNCredito.ReprogramarCreditoMonto(MatCalend, nFilaEditar - 1, CDbl(FECalend.TextMatrix(FECalend.Row, 3)))
                    lnMonto = CDbl(FECalend.TextMatrix(FECalend.row, 3))
                Else
                    lnMonto = 0
                End If
                If IsDate(FECalend.TextMatrix(nFilaEditar - 1, 1)) = True Then
                    dFecTemp = FECalend.TextMatrix(nFilaEditar - 1, 1)
                End If
                MatCalend = oNCredito.ReprogramarCreditoenMemoria(ActxCta.NroCuenta, fnTasaInteres, dFecTemp, CDate(FECalend.TextMatrix(nFilaEditar, 1)), nFilaEditar - 1, 4, False, MatCalend, lnMonto, ldVigencia, chkcalendOrig.value, CDate(lblfecUltCuota))
                
                'If (CDbl(FECalend.TextMatrix(nFilaEditar, 3)) - (CDbl(MatCalend(nFilaEditar, 4)) + CDbl(MatCalend(nFilaEditar - 1, 6)) + CDbl(MatCalend(nFilaEditar - 1, 8)) + CDbl(MatCalend(nFilaEditar - 1, 10)) + CDbl(MatCalend(nFilaEditar - 1, 12)) + CDbl(MatCalend(nFilaEditar - 1, 14)))) < CDbl(MatCalend(nFilaEditar - 1, 3)) Then
                '        MsgBox "El monto de la cuota debe ser " & Format((CDbl(MatCalend(nFilaEditar - 1, 4)) + CDbl(MatCalend(nFilaEditar - 1, 6)) + CDbl(MatCalend(nFilaEditar - 1, 8)) + CDbl(MatCalend(nFilaEditar - 1, 10)) + CDbl(MatCalend(nFilaEditar - 1, 12)) + CDbl(MatCalend(nFilaEditar - 1, 14)) + CDbl(MatCalend(nFilaEditar - 1, 3))), "#0.00") & " por tener  capital pagado ", vbInformation, "Aviso"
                '        Exit Sub
                'End If
            End If
        End If
    End If
    
    Set oNCredito = Nothing
    
    If (CDbl(FECalend.TextMatrix(nFilaEditar, 3)) - (CDbl(MatCalend(nFilaEditar - 1, 4)) + CDbl(MatCalend(nFilaEditar - 1, 6)) + CDbl(MatCalend(nFilaEditar - 1, 8)) + CDbl(MatCalend(nFilaEditar - 1, 10)) + CDbl(MatCalend(nFilaEditar - 1, 12)) + CDbl(MatCalend(nFilaEditar - 1, 14)))) < CDbl(MatCalend(nFilaEditar - 1, 3)) Then
        MsgBox "El monto de la cuota debe ser " & Format((CDbl(MatCalend(nFilaEditar - 1, 4)) + CDbl(MatCalend(nFilaEditar - 1, 6)) + CDbl(MatCalend(nFilaEditar - 1, 8)) + CDbl(MatCalend(nFilaEditar - 1, 10)) + CDbl(MatCalend(nFilaEditar - 1, 12)) + CDbl(MatCalend(nFilaEditar - 1, 14)) + CDbl(MatCalend(nFilaEditar - 1, 3))), "#0.00") & " por tener  capital pagado ", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If (nFilaEditar = FECalend.rows - 1) And FECalend.Col = 3 Then
        MsgBox "No se puede Actualizar el monto de la Ultima Cuota", vbInformation, "Aviso"
        FECalend.TextMatrix(nFilaEditar, 3) = Format(CDbl(FECalend.TextMatrix(nFilaEditar, 4)) + CDbl(FECalend.TextMatrix(nFilaEditar, 5)) + CDbl(FECalend.TextMatrix(nFilaEditar, 6)) + CDbl(FECalend.TextMatrix(nFilaEditar, 7)) + CDbl(FECalend.TextMatrix(nFilaEditar, 8)) + CDbl(FECalend.TextMatrix(nFilaEditar, 9)), "#0.00")
        Exit Sub
    End If
    'Set oNCredito = Nothing
    LimpiaFlex FECalend
    Call Carga_Mat_A_Flex(MatCalend)
    
    '*********** Restriccion de la Ultima Cuota **************
    If OptTipoRep(0).value Then
        Dim i As Integer
        
        If nReprogUltimaCuotaFija = 0 Then
        
            For i = 0 To UBound(MatCalend) - 2 'Hasta la penultima cuota
                If CDate(MatCalend(i, 0)) > CDate(lblfecUltCuota.Caption) Then
                    'MsgBox "La Fecha de vencimiento de la Cuota " & MatCalend(i, 1) & " no puede ser mayor a la de la Ultima Cuota", vbInformation, "Mensaje"
                    MsgBox "La Fecha de vencimiento de la última cuota no puede exceder el Plazo  - 1 dias", vbInformation, "Mensaje"
                    Call cmdCancelar_Click
                    Exit Sub
                End If
            Next
        Else    'Ultima cuota fija
            For i = 0 To UBound(MatCalend) - 2 'Hasta la penultima cuota
                If CDate(MatCalend(i, 0)) > CDate(lblfecUltCuota.Caption) Then
                    MsgBox "La Fecha de vencimiento de la Cuota " & MatCalend(i, 1) & " no puede ser mayor a la de la Ultima Cuota", vbInformation, "Mensaje"
                    Call cmdCancelar_Click
                    Exit Sub
                End If
            Next
            'La ultima cuota debe ser fija
            FECalend.TextMatrix(UBound(MatCalend), 1) = CDate(lblfecUltCuota.Caption)
            MatCalend(UBound(MatCalend) - 1, 0) = CDate(lblfecUltCuota.Caption)
        End If
    End If
    ' *********************************************************
    
    
    If nFilaEditar > 0 Then
        FECalend.row = nFilaEditar
        FECalend.BackColorRow vbYellow
        FECalend.SetFocus
    End If
    
    'cmdAceptar.Enabled = True
    'cmdCancelar.Enabled = True
    'cmdNuevo.Enabled = True
    
    
    Exit Sub

ErrorFECalend_OnValidate:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub FECalend_RowColChange()
    'If nFilaEditar <> -1 Then
    '    FECalend.Row = nFilaEditar
    '    FECalend.Col = 1
    'End If
    If nFilaEditar <> -1 Then
        nFilaEditar = FECalend.row
        'FECalend.Col = 1
    End If
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then 'F12
'And ActxCta.Enabled = True
        Dim bRetSinTarjeta As Boolean
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.inicia(gColComercEmp, bRetSinTarjeta)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    CentraForm Me
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    nFilaEditar = -1
    ValidarFechaActual
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredReprogramarCredito
    fnPersoneria = 0
    fbReprogDiasAtraso = False
    HabilitarReprogramar True
    bValidaActualizacionLiq = False
    fr_TasaEspecial.Visible = False 'JOEP20200926 tasa especial y reduccion de monto
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



Private Sub LstCtas_Click()
'    If LstCtas.ListCount > 0 And LstCtas.ListIndex <> -1 Then
'        ActxCta.NroCuenta = LstCtas.Text
'        ActxCta.SetFocusCuenta
'    End If
End Sub

Private Sub LstCtas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If LstCtas.ListCount > 0 And LstCtas.ListIndex <> -1 Then
            ActxCta.NroCuenta = LstCtas.Text
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub OptRepro_Click(Index As Integer)
    If Index = 0 Then
        'CmdReprogramar.Enabled = False
        OptRenov(0).Enabled = False
        OptRenov(1).Enabled = False
        CmdEditar.Enabled = True
        OptTipoRep(0).Enabled = True
        OptTipoRep(1).Enabled = True
        OptTipoRep(2).Enabled = True
        CmdGenCal.Enabled = False
        Label7.Visible = False 'Add by gitu 12-08-08
        txtDiasporReprog.Visible = False 'Add by gitu 12-08-08
    Else
        CmdGenCal.Enabled = True
        'CmdReprogramar.Enabled = True
        Label7.Visible = True 'Add by gitu 12-08-08
        txtDiasporReprog.Visible = True 'Add by gitu 12-08-08
        cmdReprog.Enabled = True 'Add by gitu 12-08-08
        OptRenov(0).Enabled = True
        OptRenov(1).Enabled = True
        CmdEditar.Enabled = False
        OptTipoRep(0).Enabled = False
        OptTipoRep(1).Enabled = False
        OptTipoRep(2).Enabled = False
    End If
End Sub
Private Sub OptTipoRep_Click(Index As Integer)
chkcalendOrig.Enabled = False
If OptTipoRep(3).value = True Then
    chkcalendOrig.Enabled = True
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


Private Sub txtDiasporReprog_KeyPress(KeyAscii As Integer)
'    KeyAscii = SoloNumeros(KeyAscii)
'    'EJVG20121210 ***
'    If KeyAscii = 13 Then
'        'JUEZ 20131022 ******************
'        'If chkDesastreNat.Enabled Then
'        '    chkDesastreNat.SetFocus
'        If optReprogNatEspeciales(1).Enabled Then
'            optReprogNatEspeciales(1).SetFocus
'            optReprogNatEspeciales(1).value = False
'        End If
'        'END JUEZ ***********************
'    End If
'    'END EJVG *******
End Sub
Private Sub txtDiasporReprog_Change()
'    'JUEZ 20131022 ************************************************************
'    'chkDesastreNat.Enabled = IIf(val(txtDiasporReprog.Text) >= 29, True, False)
'    'optReprogNatEspeciales(1).Enabled = IIf(val(txtDiasporReprog.Text) >= 29, True, False) 'JUEZ 20140415 MEMO 903-2014
'    optReprogNatEspeciales(2).Enabled = IIf(val(txtDiasporReprog.Text) >= 29, True, False)
'    'optReprogNatEspeciales(1).value = False 'JUEZ 20140415 MEMO 903-2014
'    optReprogNatEspeciales(2).value = False
'    'END JUEZ ****************************************************************
End Sub

Private Sub txtDiasporReprog_LostFocus()
    cmdReprog.Enabled = True
End Sub
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
    LimpiaFlex FECalend
    MatCalend = ""
    lnSaldoNew = 0: i = 0: j = 0: lnCapital = 0: lnIntComp = 0: lnIntGra = 0: nMontoPoliza = 0: nDiasPorReprogramacion = 0
    nMontoApr = rsCalend!nSaldoPactado
    nSaldoCapital = val(Replace(LblSaldo.Caption, ",", ""))
    nDiasPorReprogramacion = CInt(txtDiasporReprog.Text)
    dFechaReprog = fdFechaCuotaPend + nDiasPorReprogramacion
     nTasaEspCovid = 0 'Joep20200910 Tasa Especial covid
    nOpCovid = 0 'Joep20200910 Tasa Especial covid
    
    'Joep20200910 Tasa Especial covid
    'If cmbCovidReprog.Visible = True And cmbCovidReprog.Enabled = True Then
    If cmbCovidReprog.Visible = True Then
        nOpCovid = Right(cmbCovidReprog.Text, 1)
    End If
    'Joep20200910 Tasa Especial covid
    
    'Estados datos del crédito con estado Aprobado
    Set oDCOMCredito = New COMDCredito.DCOMCredito
    Set rsDatosAprob = oDCOMCredito.RecuperaColocacEstado(ActxCta.NroCuenta, gColocEstAprob)
    Set oDCOMCredito = Nothing
    
     'Joep20200910 Tasa Especial covid
    Set oDCOMCredito = New COMDCredito.DCOMCredito
    Set rsTasaEspecialCovid = oDCOMCredito.ReprogramacionObtTasaEspecial(ActxCta.NroCuenta, nSaldoCapital)
        If Not (rsTasaEspecialCovid.BOF And rsTasaEspecialCovid.EOF) Then
            nTasaEspCovid = rsTasaEspecialCovid!nTasaInteres
        End If
    Set oDCOMCredito = Nothing
    'Joep20200910 Tasa Especial covid
    
    'Seguro Desgravamen
    Set oNCOMCalendario = New COMNCredito.NCOMCalendario
    If fnTasaSegDes <> 0 Then
        nTasaSegDes = fnTasaSegDes 'Tasa SegDes. Desembolso
    Else
        Set oNCOMCredito = New COMNCredito.NCOMCredito 'JOEP20200317 Mejora
        nTasaSegDes = oNCOMCredito.ObtenerTasaSeguroDesg(ActxCta.NroCuenta, gdFecSis, fnCantAfiliadosSegDes) 'Tasa SegDes. Actual
    End If
    
    'Seguro Incendio del desembolso
    'Comento JOEP20200414 Covid
'    nMontoPoliza = oNGasto.RecuperaMontoPoliza(ActxCta.NroCuenta, _
'                                                nCuoNoPag, _
'                                                gColocConceptoCodGastoPolizaIncendioHipoteca, _
'                                                nTasaSegInc)
    'Comento JOEP20200414 Covid
    'Add JOEP20200414 Covid
    
    'RIRO 20210215 COMENTADO ***************************************************************************
    'nMontoPoliza = oNGasto.RecuperaMontoPoliza(ActxCta.NroCuenta, _
    '                                            nCuoNoPag, _
    '                                            gColocConceptoCodGastoPolizaIncendioHipoteca, _
    '                                            nTasaSegInc, , , , 1)
    'END RIRO COMENTADO *******************************************************************************
    
    'RIRO 20210211 Se realiza de esta manera por mergencia y evitar el pase *****
    Dim oConPoliza As COMConecta.DCOMConecta
    Dim rsPoliza As ADODB.Recordset
    Dim ssql As String
    Set rsPoliza = New ADODB.Recordset
    
    ssql = "exec stp_sel_ObtieneSeguroIncendio '" & ActxCta.NroCuenta & "'," & nCuoNoPag & "," & nCuoNoPag & ", " & _
            IIf(Trim(Right(cmbCovidReprog.Text, 1)) = "", 1, Right(cmbCovidReprog.Text, 1)) & ", " & CCur(txtMontoCuota) & ", " & nDiasPorReprogramacion
            
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
                End If
            End If
        End If
    End If
    If dFechaCorte = "01/01/1900" Then
        MsgBox "Se han presentado inconvenientes al validar la póliza contra incendios, favor de comunicarse con T.I.", vbInformation, "Validación Póliza"
        Exit Sub
    End If
    'END RIRO *******************************************************************
        
    'Add JOEP20200414 Covid
    'Liquidación de la deuda:
    Dim MatCalendIni As Variant          'Matriz del Calendario Pend. a pagar
    Dim vArrayDatos As Variant           'Array de parametros de la liquidación de la deuda
    Dim nCapital As Double               'Saldo Capital
    Dim nInteresCompAFecha As Double     'Interés Compensatorio (Hasta la Fecha Reprogramación)
    Dim nInteresGraciaAFecha As Double   'Interés Gracia pendiente
    Dim nInteresCompVencAFecha As Double 'Interés Compensatorio Vencido
    Dim nInteresMoratorio As Double      'Interés Moratorio (de todas las cuotas)
    Dim nSegDesgAnt As Double
    Dim nSegIncAnt As Double
    Dim nSegIncGraciaAnt As Double
    
    'Calendario de pagos pendiente
    Set oNCOMCredito = New COMNCredito.NCOMCredito
    MatCalendIni = oNCOMCredito.RecuperaMatrizCalendarioPendiente(ActxCta.NroCuenta)
    
    'Capital
    nCapital = oNCOMCredito.MatrizCapitalAFecha(ActxCta.NroCuenta, MatCalendIni)
    
     'Comento Add JOEP20200414 Covid
     'JOEP_RIRO_20200914
'    Set rsLiquidacionConceptos = oNCOMCredito.RecuperaDatosLiquidacionReprogramacion(ActxCta.NroCuenta)
'    If Not (rsLiquidacionConceptos.BOF And rsLiquidacionConceptos.EOF) Then
'        'Liq. Interes Compensatorio.
'        nInteresCompAFecha = rsLiquidacionConceptos!LiqIntComp
'        nSegDesgAnt = rsLiquidacionConceptos!LiqSegDesg
'    End If
    'Comento JOEP_RIRO_20200914
    'Add JOEP20200414 Covid
    
    'Liq. Interes Compensatorio.
     nInteresCompAFecha = oNCOMCredito.MatrizInteresCompAFecha(ActxCta.NroCuenta, MatCalendIni, dFechaCorte)  'ADD RIRO 20210214
    'nInteresCompAFecha = oNCOMCredito.MatrizInteresCompAFecha(ActxCta.NroCuenta, MatCalendIni, gdFecSis) 'Comento JOEP20200414 Covid 'Cumple cuando la cuota no tiene días de atraso'DesComento JOEP_RIRO_20200914 / RIRO 20210215 COMENTADO
    'nInteresCompAFecha = oNCOMCredito.MatrizIntCompCalendario(MatCalendIni) 'Sumatoria de Total Int. Comp. Calend. Pendiente
    'nInteresCompAFecha = oNCOMCredito.MatrizInteresGastosAFecha(ActxCta.NroCuenta, MatCalendIni, gdFecSis) 'Interés Compensatorio + Gastos
    
    'Liq. Interés de Gracia.
    nInteresGraciaAFecha = oNCOMCredito.MatrizInteresGraciaFecha(ActxCta.NroCuenta, MatCalendIni, dFechaCorte) 'ADD RIRO 20210214
    'nInteresGraciaAFecha = fnIntGraciaPend 'RIRO 20210215 COMENTADO
    'nInteresGraciaAFecha = oNCOMCredito.MatrizInteresGraciaFecha(ActxCta.NroCuenta, MatCalendIni, gdFecSis)
    
    'Liq. Interés Moratorio
    nInteresMoratorio = oNCOMCredito.MatrizIntMoratorioCalendario(MatCalendIni)
    
    'Liq. Interés Compensatorio Vencido. (Este proceso esta en proceso de implementación)
    nInteresCompVencAFecha = oNCOMCredito.MatrizInteresCompVencidoFecha(ActxCta.NroCuenta, MatCalendIni)
    
    'Liq. de Gastos
    nSegDesgAnt = oNCOMCredito.TotalGastosAFecha(ActxCta.NroCuenta, dFechaCorte, gColocConceptoCodGastoSeguro7)  'RIRO20210214
    'nSegDesgAnt = oNCOMCredito.TotalGastosAFecha(ActxCta.NroCuenta, gdFecSis, gColocConceptoCodGastoSeguro7) 'Comento JOEP20200414 Covid 'Descomento JOEP_RIRO_20200914 / RIRO 20210215 COMENTADO
    'Comento JOEP20200414 Covid
    'nSegIncAnt = oNCOMCredito.TotalGastosAFecha(ActxCta.NroCuenta, Format(gdFecSis, "mm/dd/yyyy"), gColocConceptoCodGastoPolizaIncendioHipoteca) 'DesComento JOEP_RIRO_20200914 / RIRO 20210215 COMENTADO
    'Comento JOEP20200414 Covid
    'nSegIncAnt = 0 'Add JOEP20200414 Covid 'Comento JOEP_RIRO_20200914
    'nSegIncGraciaAnt = oNCOMCredito.TotalGastosAFecha(ActxCta.NroCuenta, Format(gdFecSis, "mm/dd/yyyy"), gColocConceptoCodGastoPolizaIncendioHipotecaGracia) / RIRO 20210215
    
    'Agrupación de importes liquidados.
    ReDim vArrayDatos(10) ' / RIRO 20210212, Se cambió de 09 a 10 length
    vArrayDatos(0) = nInteresCompAFecha
    vArrayDatos(1) = nInteresGraciaAFecha
    vArrayDatos(2) = nInteresMoratorio
    vArrayDatos(3) = nInteresCompVencAFecha
    vArrayDatos(4) = nSegDesgAnt
    vArrayDatos(5) = nPolizaMen 'nSegIncAnt / RIRO 20210212
    vArrayDatos(6) = nPolizaCuotReprog 'nSegIncGraciaAnt / RIRO 20210212
    vArrayDatos(7) = nTasaEspCovid 'Joep20200910 Tasa Especial covid
    vArrayDatos(8) = nOpCovid 'Joep20200910 Tasa Especial covid
    vArrayDatos(9) = nPolizaProrrateo '/ RIRO 20210212
    'Fin Liquidación
    
    'Generacion del calendario de pagos de las cuotas no pagadas
    ReDim MatCalendReprogramado(nCuoNoPag)
          
    MatCalendReprogramado = oNCOMCalendario.GeneraCalendario(CDbl(LblSaldo), _
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
                                                            ActxCta.NroCuenta, , , _
                                                            nInteresGraciaAFecha, _
                                                            , , , , nTasaSegDes, _
                                                            MatCalendSegDes, , _
                                                            nMontoPoliza, _
                                                            nTasaSegInc, _
                                                            vArrayDatos)
    
    'Generacion del calendario Reprogramado
    ReDim MatCalend(nCuoPag, 23) 'LUCV20180601, Modificó 11 por 17 / RIRO 20200829 de 17 a 23
    Do While Not rsCalend.EOF
        FECalend.AdicionaFila
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
            FECalend.TextMatrix(rsCalend.Bookmark, 1) = Format(rsCalend!dVenc, "dd/mm/yyyy")
            MatCalend(i, 0) = Format(rsCalend!dVenc, "dd/mm/yyyy")
            lnCapital = rsCalend!nCapital
            lnIntComp = rsCalend!nIntComp
            lnIntGra = rsCalend!nIntGracia
        'Cuotas Pendientes
        Else
            FECalend.TextMatrix(rsCalend.Bookmark, 1) = MatCalendReprogramado(j, 0)  'FechaCuota (Fila, Colum)
            MatCalend(i, 0) = MatCalendReprogramado(j, 0) 'FechaVenc.
            'MatCalend(i, 1) = MatCalendReprogramado(j, 1) 'NroCuota
            'MatCalend(i, 2) = MatCalendReprogramado(j, 2) 'MontoCuota
            lnCapital = MatCalendReprogramado(j, 3) 'Capital
            lnIntComp = MatCalendReprogramado(j, 4) 'IntComp
            lnIntGra = MatCalendReprogramado(j, 5) 'IntGrac
            'MatCalend(i, 6) = MatCalendReprogramado(j, 6) 'Gasto
            'MatCalend(i, 7) = MatCalendReprogramado(j, 7) 'SaldoCap
            'MatCalend(i, 7) = MatCalendReprogramado(j, 8) 'EstadoCuota
            
            'MatCalend(i, 11) = MatCalendReprogramado(j, 8) 'SegDesg
            'MatCalend(i, 10) = CDbl(MatCalendReprogramado(j, 15)) 'SegInc
            'MatCalend(i, 12) = CDbl(MatCalendReprogramado(j, 16)) 'SegInc por días de gracia
            
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
            nGastos = CDbl(CDbl(MatCalendReprogramado(j, 8)) + nMontoPoliza) '+ CDbl(MatCalendReprogramado(j, 16))) 'Add JOEP20200414 Covid / RIRO20210301
            j = j + 1
        End If

        'Asignación de valores en el Flex del calendario
        FECalend.TextMatrix(rsCalend.Bookmark, 2) = Trim(str(rsCalend!nCuota))                                      'Nro. Cuota
        FECalend.TextMatrix(rsCalend.Bookmark, 3) = Format(IIf(IsNull(lnCapital), 0, lnCapital) + _
                                                    IIf(IsNull(lnIntComp), 0, lnIntComp) + _
                                                    IIf(IsNull(lnIntGra), 0, lnIntGra) + _
                                                    IIf(rsCalend!nColocCalendEstado = gColocCalendEstadoPagado, IIf(IsNull(rsCalend!nIntMor), 0, rsCalend!nIntMor), 0) + _
                                                    IIf(IsNull(rsCalend!nIntReprog), 0, rsCalend!nIntReprog) + _
                                                    nGastos, "#0.00")                                               'Importe Cuota
        FECalend.TextMatrix(rsCalend.Bookmark, 4) = Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")           'Capital
        FECalend.TextMatrix(rsCalend.Bookmark, 5) = Format(IIf(IsNull(lnIntComp), 0, lnIntComp), "#0.00")           'Interés Compensatorio
        'LUCV20180601, Agregó según ERS022-2018
        If rsCalend!nColocCalendEstado = gColocCalendEstadoPagado Then
            FECalend.TextMatrix(rsCalend.Bookmark, 6) = Format(IIf(IsNull(rsCalend!nIntMor), 0, rsCalend!nIntMor), "#0.00") 'Interés Moratorio
        Else
            FECalend.TextMatrix(rsCalend.Bookmark, 6) = Format(0, "#0.00") 'Interés Moratorio
        End If
        'FIN LUCV20180601.
        
        FECalend.TextMatrix(rsCalend.Bookmark, 7) = Format(IIf(IsNull(rsCalend!nIntReprog), 0, rsCalend!nIntReprog), "#0.00")
        FECalend.TextMatrix(rsCalend.Bookmark, 8) = Format(IIf(IsNull(lnIntGra), 0, lnIntGra), "#0.00")                 'Interés de Gracia
        FECalend.TextMatrix(rsCalend.Bookmark, 9) = Format(nGastos, "#0.00")
        nMontoApr = nMontoApr - IIf(IsNull(lnCapital), 0, lnCapital)
        nMontoApr = CDbl(Format(nMontoApr, "#0.0000"))
        FECalend.TextMatrix(rsCalend.Bookmark, 10) = Format(nMontoApr, "#0.00")
        FECalend.TextMatrix(rsCalend.Bookmark, 11) = Trim(str(rsCalend!nColocCalendEstado))                             'Estado Cuota
        FECalend.TextMatrix(rsCalend.Bookmark, 12) = Format(IIf(IsNull(rsCalend!nCapitalPag), 0, rsCalend!nCapitalPag), "#0.00")
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
                
        If rsCalend!nColocCalendEstado = gColocCalendEstadoPagado Then
            FECalend.row = rsCalend.Bookmark
            Call FECalend.ForeColorRow(vbRed)
        End If
        If rsCalend.RecordCount = rsCalend.Bookmark Then
            lblfecUltCuota = Format(rsCalend!dVenc, "dd/mm/yyyy")
        End If

        i = i + 1
        rsCalend.MoveNext

        nTotalCapital = nTotalCapital + Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")
    Loop

    'JOEP Ajuste para que cuadre el capital
    'If LblPrestamo <> nTotalCapital Then 'JOEP20200327 comento Covid-19
    'If LblSaldo <> nTotalCapital Then 'JOEP20200327 Add Covid-19
'    If RTrim(Replace(LblSaldo, ",", "")) <> RTrim(nTotalCapital) Then 'JOEP20200330 Add Covid-19
'        'nAjuste = nTotalCapital - LblPrestamo 'JOEP20200327 comento Covid-19
'        nAjuste = LblPrestamo - nTotalCapital 'JOEP20200327 Add Covid-19
'        nCuotaAjustada = MatCalend(nValCuoataAju, 3) - nAjuste
'        MatCalend(nValCuoataAju, 3) = Format(nCuotaAjustada, "#0.00")
'        lnSaldoNew = LblSaldo
'    End If
    'JOEP Ajuste para que cuadre el capital

    'JOEP Identificar Tipo de Periodo(Para calcular TCEA)
    Set oDCOMCredito = New COMDCredito.DCOMCredito
    Set rsTipPeriodo = oDCOMCredito.IdentificarTipoPeriodo(ActxCta.NroCuenta)
    If Not (rsTipPeriodo.EOF And rsTipPeriodo.BOF) Then
        nTipoPeriodo = rsTipPeriodo!nTpPeriodo
    End If
    
    NewTCEA = oNCOMCredito.GeneraTasaCostoEfectivoAnual(CDate(Format(ldVigencia, "dd/mm/yyyy")), CDbl(LblPrestamo), MatCalend, CDbl(fnTasaInteres), ActxCta.NroCuenta, nTipoPeriodo)  'Para calcular la TCEA
    txtTCEAReprog = NewTCEA & " %" 'Add JOEP20200425 Cuota Igual
    rsCalend.Close
    Set rsCalend = Nothing
    lblSaldoRep = Format(lnSaldoNew, "#,##0.00")
End Sub
'<-*****Fin LUCV20180601

'JOEP20200428 covid cuotas iguales
Private Sub CovidCuotasIguales()
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
    LimpiaFlex FECalend
    MatCalend = ""
    lnSaldoNew = 0: i = 0: j = 0: lnCapital = 0: lnIntComp = 0: lnIntGra = 0: nMontoPoliza = 0: nDiasPorReprogramacion = 0
    nMontoApr = rsCalend!nSaldoPactado
    nSaldoCapital = val(Replace(LblSaldo.Caption, ",", ""))
    nDiasPorReprogramacion = CInt(txtDiasporReprog.Text)
    dFechaReprog = fdFechaCuotaPend + nDiasPorReprogramacion
    
    'Estados datos del crédito con estado Aprobado
    Set oDCOMCredito = New COMDCredito.DCOMCredito
    Set rsDatosAprob = oDCOMCredito.RecuperaColocacEstado(ActxCta.NroCuenta, gColocEstAprob)
    Set oDCOMCredito = Nothing
    
    'Seguro Desgravamen
    Set oNCOMCalendario = New COMNCredito.NCOMCalendario
    If fnTasaSegDes <> 0 Then
        nTasaSegDes = fnTasaSegDes 'Tasa SegDes. Desembolso
    Else
        Set oNCOMCredito = New COMNCredito.NCOMCredito 'JOEP20200317 Mejora
        nTasaSegDes = oNCOMCredito.ObtenerTasaSeguroDesg(ActxCta.NroCuenta, gdFecSis, fnCantAfiliadosSegDes) 'Tasa SegDes. Actual
    End If
        
    'Add JOEP20200414 Covid
    
    'RIRO 20210212 Comentado ********************
    'nMontoPoliza = oNGasto.RecuperaMontoPoliza(ActxCta.NroCuenta, _
    '                                            nCuoNoPag, _
    '                                            gColocConceptoCodGastoPolizaIncendioHipoteca, _
    '                                            nTasaSegInc, , , , 1)
    'End RIRO ***********************************
    
    'RIRO 20210211 Se realiza de esta manera por mergencia y evitar el pase *****
    Dim oConPoliza As COMConecta.DCOMConecta
    Dim rsPoliza As ADODB.Recordset
    Dim ssql As String
    Set rsPoliza = New ADODB.Recordset
    
    ssql = "exec stp_sel_ObtieneSeguroIncendio '" & ActxCta.NroCuenta & "'," & nCuoNoPag & "," & nCuoNoPag & ", " & _
            IIf(Trim(Right(cmbCovidReprog.Text, 1)) = "", 1, Right(cmbCovidReprog.Text, 1)) & ", " & CCur(txtMontoCuota) & ", " & nDiasPorReprogramacion
            
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
                End If
            End If
        End If
    End If
    If dFechaCorte = "01/01/1900" Then
        MsgBox "Se han presentado inconvenientes al validar la póliza contra incendios, favor de comunicarse con T.I.", vbInformation, "Validación Póliza"
        Exit Sub
    End If
    'END RIRO *******************************************************************
                                                
    'Add JOEP20200414 Covid
    'Liquidación de la deuda:
    Dim MatCalendIni As Variant          'Matriz del Calendario Pend. a pagar
    Dim vArrayDatos As Variant           'Array de parametros de la liquidación de la deuda
    Dim nCapital As Double               'Saldo Capital
    Dim nInteresCompAFecha As Double     'Interés Compensatorio (Hasta la Fecha Reprogramación)
    Dim nInteresGraciaAFecha As Double   'Interés Gracia pendiente
    Dim nInteresCompVencAFecha As Double 'Interés Compensatorio Vencido
    Dim nInteresMoratorio As Double      'Interés Moratorio (de todas las cuotas)
    Dim nSegDesgAnt As Double
    Dim nSegIncAnt As Double
    Dim nSegIncGraciaAnt As Double
    
    'Calendario de pagos pendiente
    Set oNCOMCredito = New COMNCredito.NCOMCredito
    MatCalendIni = oNCOMCredito.RecuperaMatrizCalendarioPendiente(ActxCta.NroCuenta)
    
    'Capital
    nCapital = oNCOMCredito.MatrizCapitalAFecha(ActxCta.NroCuenta, MatCalendIni)
    
     'Add JOEP20200414 Covid
     'Comento RIRO_JOEP-2020
'    Set rsLiquidacionConceptos = oNCOMCredito.RecuperaDatosLiquidacionReprogramacion(ActxCta.NroCuenta)
'    If Not (rsLiquidacionConceptos.BOF And rsLiquidacionConceptos.EOF) Then
'        'Liq. Interes Compensatorio.
'        nInteresCompAFecha = rsLiquidacionConceptos!LiqIntComp
'        nSegDesgAnt = rsLiquidacionConceptos!LiqSegDesg
'    End If
    'Comento RIRO_JOEP-2020
'    'Add JOEP20200414 Covid
    
    'Liq. Interes Compensatorio.
    'nInteresCompAFecha = oNCOMCredito.MatrizInteresCompAFecha(ActxCta.NroCuenta, MatCalendIni, gdFecSis) 'Comento JOEP20200414 Covid 'Cumple cuando la cuota no tiene días de atraso'DesComento JOEP_RIRO_20200914
    nInteresCompAFecha = oNCOMCredito.MatrizInteresCompAFecha(ActxCta.NroCuenta, MatCalendIni, dFechaCorte) 'ADD RIRO 20210214
    
    'Liq. Interés de Gracia.
    'nInteresGraciaAFecha = fnIntGraciaPend RIRO20210214 Comentado
    nInteresGraciaAFecha = oNCOMCredito.MatrizInteresGraciaFecha(ActxCta.NroCuenta, MatCalendIni, dFechaCorte) 'ADD RIRO 20210214
    
    'Liq. Interés Moratorio
    nInteresMoratorio = oNCOMCredito.MatrizIntMoratorioCalendario(MatCalendIni)
    
    'Liq. Interés Compensatorio Vencido. (Este proceso esta en proceso de implementación)
    nInteresCompVencAFecha = oNCOMCredito.MatrizInteresCompVencidoFecha(ActxCta.NroCuenta, MatCalendIni)
    
    'Liq. de Gastos
    'nSegDesgAnt = oNCOMCredito.TotalGastosAFecha(ActxCta.NroCuenta, gdFecSis, gColocConceptoCodGastoSeguro7)  'Descomento JOEP_RIRO_20200914 / RIRO 20210214 Comentado
    nSegDesgAnt = oNCOMCredito.TotalGastosAFecha(ActxCta.NroCuenta, dFechaCorte, gColocConceptoCodGastoSeguro7)  'RIRO20210214 ADD
    
    'nSegIncAnt = oNCOMCredito.TotalGastosAFecha(ActxCta.NroCuenta, Format(gdFecSis, "mm/dd/yyyy"), gColocConceptoCodGastoPolizaIncendioHipoteca) 'DesComento JOEP_RIRO_20200914 /* RIRO 20210212 COMENTADO */
    
    'Liq. de Gastos
    'nSegIncAnt = 0 'Add JOEP20200414 Covid 'Comento RIRO_JOEP_20200914
    'nSegIncGraciaAnt = oNCOMCredito.TotalGastosAFecha(ActxCta.NroCuenta, Format(gdFecSis, "mm/dd/yyyy"), gColocConceptoCodGastoPolizaIncendioHipotecaGracia) /* RIRO 20210212 COMENTADO */
    
    'Agrupación de importes liquidados.
    ReDim vArrayDatos(10) ' / RIRO 20210212, Se cambió de 09 a 10 length
    vArrayDatos(0) = nInteresCompAFecha
    vArrayDatos(1) = nInteresGraciaAFecha
    vArrayDatos(2) = nInteresMoratorio
    vArrayDatos(3) = nInteresCompVencAFecha
    vArrayDatos(4) = nSegDesgAnt
    vArrayDatos(5) = nPolizaMen 'nSegIncAnt / RIRO 20210212
    vArrayDatos(6) = nPolizaCuotReprog 'nSegIncGraciaAnt / RIRO 20210212
    vArrayDatos(7) = CCur(txtMontoCuota) - nMontoPoliza
    vArrayDatos(8) = 0 'Add JOEP-LUCV
    vArrayDatos(9) = nPolizaProrrateo '/ RIRO 20210212
    'Fin Liquidación
        
    'Generacion del calendario de pagos de las cuotas no pagadas
    ReDim MatCalendReprogramado(nCuoNoPag)
    
    MatCalendReprogramado = oNCOMCalendario.CovidCuotasIgualesGeneraCalendario(CDbl(LblSaldo), _
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
                                                            ActxCta.NroCuenta, , , _
                                                            nInteresGraciaAFecha, _
                                                            , , , , nTasaSegDes, _
                                                            MatCalendSegDes, , _
                                                            nMontoPoliza, _
                                                            nTasaSegInc, _
                                                            vArrayDatos)
    
  'JOEP-Lucv coivd
    Dim nCuotasPag As Integer
    Dim X As Integer
    nCuotasPag = nCuoPag - nCuoNoPag
    nCuotasPag = nCuotasPag + vArrayDatos(8)
    
    'Generacion del calendario Reprogramado
    'ReDim MatCalend(nCuoPag, 17) 'LUCV20180601, Modificó 11 por 17
    ReDim MatCalend(nCuotasPag, 23) 'LUCV20180601, Modificó 11 por 17/ RIRO Se modificó a 23
    'JOEP-Lucv coivd
  
    'Do While Not rsCalend.EOF'JOEP-Lucv coivd
    For X = 1 To nCuotasPag '- 1 'JOEP-Lucv coivd
        FECalend.AdicionaFila
        
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
            FECalend.TextMatrix(X, 1) = MatCalendReprogramado(j, 0)  'FechaCuota (Fila, Colum)
            MatCalend(i, 0) = MatCalendReprogramado(j, 0) 'FechaVenc.
            lnCapital = MatCalendReprogramado(j, 3) 'Capital
            lnIntComp = MatCalendReprogramado(j, 4) 'IntComp
            lnIntGra = MatCalendReprogramado(j, 5) 'IntGrac
                    
            nGastoSegDesg = MatCalendReprogramado(j, 8)
            nGastoIncendio = CDbl(MatCalendReprogramado(j, 15))
            nGastoIncendioGracia = CDbl(MatCalendReprogramado(j, 16)) 'SegInc por días de gracia
                                
            nGastos = CDbl(CDbl(MatCalendReprogramado(j, 8)) + nMontoPoliza) '+ CDbl(MatCalendReprogramado(j, 16))) 'Add JOEP20200414 Covid / RIRO 20210212
            'j = j + 1 RIRO 20210214
                        
            'Asignación de valores en el Flex del calendario
            FECalend.TextMatrix(X, 2) = X                                      'Nro. Cuota
            FECalend.TextMatrix(X, 3) = Format(IIf(IsNull(lnCapital), 0, lnCapital) + _
                                                        IIf(IsNull(lnIntComp), 0, lnIntComp) + _
                                                        IIf(IsNull(lnIntGra), 0, lnIntGra) + _
                                                        0 + _
                                                        0 + _
                                                        nGastos, "#0.00") 'Importe Cuota
            FECalend.TextMatrix(X, 4) = Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")           'Capital
            FECalend.TextMatrix(X, 5) = Format(IIf(IsNull(lnIntComp), 0, lnIntComp), "#0.00")           'Interés Compensatorio
            'LUCV20180601, Agregó según ERS022-2018
            FECalend.TextMatrix(X, 6) = Format(0, "#0.00") 'Interés Moratorio
            'FIN LUCV20180601.
            
            FECalend.TextMatrix(X, 7) = Format(0, "#0.00")
            FECalend.TextMatrix(X, 8) = Format(IIf(IsNull(lnIntGra), 0, lnIntGra), "#0.00")                 'Interés de Gracia
            FECalend.TextMatrix(X, 9) = Format(nGastos, "#0.00")
            nMontoApr = nMontoApr - IIf(IsNull(lnCapital), 0, lnCapital)
            nMontoApr = CDbl(Format(nMontoApr, "#0.0000"))
            FECalend.TextMatrix(X, 10) = Format(nMontoApr, "#0.00")
            FECalend.TextMatrix(X, 11) = 0                            'Estado Cuota
            FECalend.TextMatrix(X, 12) = Format(0, "#0.00")
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
            
            'RIRO 20210214 Liquidación *******************************
            nIntCompCalculado = CDbl(MatCalendReprogramado(j, 17))
            nDiasCalculo = CInt(MatCalendReprogramado(j, 18))
            nIntCompDiferenciaCapitalizado = CDbl(MatCalendReprogramado(j, 19))
            nIntGraciaGenerado = CDbl(MatCalendReprogramado(j, 20))
            nIntGraciaCapitalizado = CDbl(MatCalendReprogramado(j, 21))
            nIntGraciaAsignado = CDbl(MatCalendReprogramado(j, 22))
            
            MatCalend(i, 18) = nIntCompCalculado
            MatCalend(i, 19) = nDiasCalculo
            MatCalend(i, 20) = nIntCompDiferenciaCapitalizado
            MatCalend(i, 21) = nIntGraciaGenerado
            MatCalend(i, 22) = nIntGraciaCapitalizado
            MatCalend(i, 23) = nIntGraciaAsignado
            'RIRO 20210214 Corrección Liquidación *********************
            
'            If rsCalend.RecordCount = rsCalend.Bookmark Then
'                lblfecUltCuota = Format(rsCalend!dVenc, "dd/mm/yyyy")
'            End If
    
            i = i + 1
            j = j + 1 'RIRO 20210214
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
                FECalend.TextMatrix(rsCalend.Bookmark, 1) = Format(rsCalend!dVenc, "dd/mm/yyyy")
                MatCalend(i, 0) = Format(rsCalend!dVenc, "dd/mm/yyyy")
                lnCapital = rsCalend!nCapital
                lnIntComp = rsCalend!nIntComp
                lnIntGra = rsCalend!nIntGracia
            'Cuotas Pendientes
            Else
                FECalend.TextMatrix(rsCalend.Bookmark, 1) = MatCalendReprogramado(j, 0)  'FechaCuota (Fila, Colum)
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
                                
                nGastos = CDbl(CDbl(MatCalendReprogramado(j, 8)) + nMontoPoliza) '+ CDbl(MatCalendReprogramado(j, 16))) 'Add JOEP20200414 Covid / RIRO 20210212
                j = j + 1
            End If
            
            'Asignación de valores en el Flex del calendario
            FECalend.TextMatrix(rsCalend.Bookmark, 2) = Trim(str(rsCalend!nCuota))                                      'Nro. Cuota
            FECalend.TextMatrix(rsCalend.Bookmark, 3) = Format(IIf(IsNull(lnCapital), 0, lnCapital) + _
                                                        IIf(IsNull(lnIntComp), 0, lnIntComp) + _
                                                        IIf(IsNull(lnIntGra), 0, lnIntGra) + _
                                                        IIf(rsCalend!nColocCalendEstado = gColocCalendEstadoPagado, IIf(IsNull(rsCalend!nIntMor), 0, rsCalend!nIntMor), 0) + _
                                                        IIf(IsNull(rsCalend!nIntReprog), 0, rsCalend!nIntReprog) + _
                                                        nGastos, "#0.00")                                               'Importe Cuota
            FECalend.TextMatrix(rsCalend.Bookmark, 4) = Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")           'Capital
            FECalend.TextMatrix(rsCalend.Bookmark, 5) = Format(IIf(IsNull(lnIntComp), 0, lnIntComp), "#0.00")           'Interés Compensatorio
            'LUCV20180601, Agregó según ERS022-2018
            If rsCalend!nColocCalendEstado = gColocCalendEstadoPagado Then
                FECalend.TextMatrix(rsCalend.Bookmark, 6) = Format(IIf(IsNull(rsCalend!nIntMor), 0, rsCalend!nIntMor), "#0.00") 'Interés Moratorio
            Else
                FECalend.TextMatrix(rsCalend.Bookmark, 6) = Format(0, "#0.00") 'Interés Moratorio
            End If
            'FIN LUCV20180601.
            
            FECalend.TextMatrix(rsCalend.Bookmark, 7) = Format(IIf(IsNull(rsCalend!nIntReprog), 0, rsCalend!nIntReprog), "#0.00")
            FECalend.TextMatrix(rsCalend.Bookmark, 8) = Format(IIf(IsNull(lnIntGra), 0, lnIntGra), "#0.00")                 'Interés de Gracia
            FECalend.TextMatrix(rsCalend.Bookmark, 9) = Format(nGastos, "#0.00")
            
            'JOEP20210803
            If rsCalend!nColocCalendEstado = gColocCalendEstadoPagado Then
                nMontoApr = nMontoApr - IIf(IsNull(lnCapital), 0, lnCapital)
            Else
                nMontoApr = (nMontoApr - IIf(IsNull(rsCalend!nCapitalPag), 0, rsCalend!nCapitalPag)) - IIf(IsNull(lnCapital), 0, lnCapital)
            End If
            'JOEP20210803
            
            nMontoApr = CDbl(Format(nMontoApr, "#0.0000"))
            FECalend.TextMatrix(rsCalend.Bookmark, 10) = Format(nMontoApr, "#0.00")
            FECalend.TextMatrix(rsCalend.Bookmark, 11) = Trim(str(rsCalend!nColocCalendEstado))                             'Estado Cuota
            FECalend.TextMatrix(rsCalend.Bookmark, 12) = Format(IIf(IsNull(rsCalend!nCapitalPag), 0, rsCalend!nCapitalPag), "#0.00")
                        
            'JOEP20210803
            If rsCalend!nColocCalendEstado = gColocCalendEstadoPagado Then
                lnSaldoNew = lnSaldoNew + IIf(IsNull(lnCapital), 0, lnCapital) - IIf(IsNull(rsCalend!nCapitalPag), 0, rsCalend!nCapitalPag)
            Else
                lnSaldoNew = lnSaldoNew + IIf(IsNull(lnCapital), 0, lnCapital)
            End If
            'JOEP20210803
            
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
            
            If rsCalend!nColocCalendEstado = gColocCalendEstadoPagado Then
                FECalend.row = rsCalend.Bookmark
                Call FECalend.ForeColorRow(vbRed)
            End If
            If rsCalend.RecordCount = rsCalend.Bookmark Then
                lblfecUltCuota = Format(rsCalend!dVenc, "dd/mm/yyyy")
            End If
    
            i = i + 1
            rsCalend.MoveNext
        End If
            nTotalCapital = nTotalCapital + Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")
    'Loop'JOEP-Lucv coivd
    Next X 'JOEP-Lucv coivd
    
    'JOEP Identificar Tipo de Periodo(Para calcular TCEA)
    Set oDCOMCredito = New COMDCredito.DCOMCredito
    Set rsTipPeriodo = oDCOMCredito.IdentificarTipoPeriodo(ActxCta.NroCuenta)
    If Not (rsTipPeriodo.EOF And rsTipPeriodo.BOF) Then
        nTipoPeriodo = rsTipPeriodo!nTpPeriodo
    End If
    
    NewTCEA = oNCOMCredito.GeneraTasaCostoEfectivoAnual(CDate(Format(ldVigencia, "dd/mm/yyyy")), CDbl(LblPrestamo), MatCalend, CDbl(fnTasaInteres), ActxCta.NroCuenta, nTipoPeriodo)  'Para calcular la TCEA
    txtTCEAReprog = NewTCEA & " %" 'Add JOEP20200425 Cuota Igual
    rsCalend.Close
    Set rsCalend = Nothing
    lblSaldoRep = Format(lnSaldoNew, "#,##0.00")
End Sub

Private Sub CmbReprogNatEspeciales_Click()
Dim rsCovidOpciones As ADODB.Recordset
Dim oDCOMCred As COMDConstantes.DCOMConstantes
Set oDCOMCred = New COMDConstantes.DCOMConstantes

Dim rsMsgbox As ADODB.Recordset
Dim objMsg As COMDCredito.DCOMCredito

If CmbReprogNatEspeciales.Visible = True And Trim(Right(CmbReprogNatEspeciales.Text, 2)) <> "" Then
    Set objMsg = New COMDCredito.DCOMCredito
    Set rsMsgbox = objMsg.MsgBoxReprog(ActxCta.NroCuenta, Trim(Right(CmbReprogNatEspeciales.Text, 2)))
    If Not (rsMsgbox.BOF And rsMsgbox.EOF) Then
        If rsMsgbox!MsgBox <> "" Then
            MsgBox rsMsgbox!MsgBox, vbInformation, "Aviso"
        End If
    End If
End If

    If CmbReprogNatEspeciales.Visible = True And Right(CmbReprogNatEspeciales.Text, 2) = 1 Then
        Set rsCovidOpciones = oDCOMCred.RecuperaConstantes(2090)
        If Not (rsCovidOpciones.BOF And rsCovidOpciones.EOF) Then
            Call Llenar_Combo_con_Recordset(rsCovidOpciones, cmbCovidReprog)
            Call CambiaTamañoCombo(cmbCovidReprog, 100)
        End If
        cmbCovidReprog.Visible = True
    Else
        cmbCovidReprog.ListIndex = -1
        cmbCovidReprog.Visible = False
    End If
Set oDCOMCred = Nothing
RSClose rsCovidOpciones
End Sub

Private Sub EnableBotones(ByVal Opcion As Integer, Optional ByVal bValor As Boolean)
    If Opcion = 0 Then
        If CmbReprogNatEspeciales.Visible = True And Right(cmbCovidReprog.Text, 1) = "2" Then
            CmbReprogNatEspeciales.Enabled = bValor
            cmbCovidReprog.Enabled = bValor
        End If
    End If
End Sub

Private Sub VisibleBotones(ByVal Opcion As Integer, Optional ByVal bValor As Boolean)
    If Opcion = 0 Then
        If CmbReprogNatEspeciales.Visible = True And Right(cmbCovidReprog.Text, 1) = "2" Then
            frMontCuotaCovid.Visible = bValor
            txtMontoCuota.Enabled = False
            fr_TasaEspecial.Visible = False
        ElseIf CmbReprogNatEspeciales.Visible = True And Right(cmbCovidReprog.Text, 1) = "4" Then
            frMontCuotaCovid.Visible = bValor
            txtMontoCuota.Enabled = bValor
            frMontCuotaCovid.Enabled = bValor
            txtMontoCuota.Text = Format(0#, "#,#0.00")
            fr_TasaEspecial.Visible = False
        ElseIf CmbReprogNatEspeciales.Visible = True And Right(cmbCovidReprog.Text, 1) = "3" Then
            fr_TasaEspecial.Visible = bValor
            frMontCuotaCovid.Visible = False
        Else
            frMontCuotaCovid.Visible = False
            txtMontoCuota.Text = Format(0#, "#,#0.00")
            fr_TasaEspecial.Visible = False
            'CmbReprogNatEspeciales.Visible = False
        End If
    End If
End Sub

Private Sub cmbCovidReprog_Click()
Dim rsMC As ADODB.Recordset
Dim rsTasaEspecialCovid As ADODB.Recordset 'Joep20200910 Tasa Especial covid
Dim obMc As COMDCredito.DCOMCredito
Set obMc = New COMDCredito.DCOMCredito
nMontoCuota = 0

If CmbReprogNatEspeciales.Visible = True And cmbCovidReprog.Visible = True And cmbCovidReprog.Enabled = True Then
    If Right(cmbCovidReprog.Text, 1) = "2" Or Right(cmbCovidReprog.Text, 1) = "4" Then
        Set rsMC = obMc.ReprogramacionObtMantenerCuota(ActxCta.NroCuenta)
            If Not (rsMC.BOF And rsMC.EOF) Then
                txtMontoCuota.Text = Format(rsMC!nMontoCuota, "#,#0.00")
                nMontoCuota = Format(txtMontoCuota, "#,#0.00")
                Call VisibleBotones(0, True)
            End If
    ElseIf Right(cmbCovidReprog.Text, 1) = "3" Then
        Set rsTasaEspecialCovid = obMc.ReprogramacionObtTasaEspecial(ActxCta.NroCuenta, LblSaldo)
            If Not (rsTasaEspecialCovid.BOF And rsTasaEspecialCovid.EOF) Then
                lbl_TasaEspecial = Format(rsTasaEspecialCovid!nTasaInteres, "0.00")
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
    Dim rsRepgOCM As ADODB.Recordset 'JOEP20200928 Reprogramacion OCM
    Dim obOCM As COMDCredito.DCOMCredito 'JOEP20200928 Reprogramacion OCM
    Dim oDCOMCred As COMDConstantes.DCOMConstantes
    Set oDCOMCred = New COMDConstantes.DCOMConstantes
    Set obOCM = New COMDCredito.DCOMCredito 'JOEP20200928 Reprogramacion OCM
    Dim bOCM As Integer 'JOEP20200928 Reprogramacion OCM
    bOCM = 0 'JOEP20200928 Reprogramacion OCM
    
    'JOEP20200928 Reprogramacion OCM
    Set rsRepgOCM = obOCM.ReprogramacionOCM(ActxCta.NroCuenta)
    If Not (rsRepgOCM.BOF And rsRepgOCM.EOF) Then
        bOCM = rsRepgOCM!nModalidad
    End If
    'JOEP20200928 Reprogramacion OCM
    
    Set rsReprogEmergencia = oDCOMCred.RecuperaConstanteReprogaramacion(2080, bOCM)
    If Not (rsReprogEmergencia.BOF And rsReprogEmergencia.EOF) Then
        Call Llenar_Combo_con_Recordset(rsReprogEmergencia, CmbReprogNatEspeciales)
        Call CambiaTamañoCombo(CmbReprogNatEspeciales, 200)
    End If
    
    Set oDCOMCred = Nothing
    Set obOCM = Nothing
    RSClose rsRepgOCM
    RSClose rsReprogEmergencia
    
    CmbReprogNatEspeciales.Visible = False
        
End Sub
'JOEP20200428 covid cuotas iguales

'JOEP20200926 Tasa Especial y Reduccion de monto
Private Function ValidaDatos(ByVal pnBoton As Integer) As Boolean
Dim rsVD As ADODB.Recordset
Dim obVD As COMDCredito.DCOMCredito
Set obVD = New COMDCredito.DCOMCredito
ValidaDatos = False

Dim nCapital As Currency
Dim dFechaVencFinReprg As String
nCapital = 0
dFechaVencFinReprg = ""
Dim i As Integer
    For i = 1 To FECalend.rows - 1
        If FECalend.TextMatrix(i, 11) = 0 Then
            nCapital = nCapital + FECalend.TextMatrix(i, 4)
            dFechaVencFinReprg = FECalend.TextMatrix(i, 1)
        End If
    Next i

Set rsVD = obVD.ReprogramacionValidDatos(ActxCta.NroCuenta, TxtGlosa.Text, CCur(txtMontoCuota.Text), nCapital, IIf(CmbReprogNatEspeciales.Visible, 1, 0), IIf(cmbCovidReprog.Visible, 1, 0), IIf(cmbCovidReprog.Visible = True, Right(cmbCovidReprog.Text, 1), -1), nMontoCuota, CCur(lbl_TasaEspecial), CCur(LblTasa), dFechaVencFinReprg, pnBoton)


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
    CmbReprogNatEspeciales.Visible = True
    CmbReprogNatEspeciales.ListIndex = 0
    cmbCovidReprog.ListIndex = pnModalidad - 1
    CmbReprogNatEspeciales.Enabled = False
    cmbCovidReprog.Enabled = False
End Sub
'JOEP20200926 Tasa Especial y Reduccion de monto
