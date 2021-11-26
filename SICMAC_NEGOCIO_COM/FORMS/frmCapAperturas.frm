VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmCapAperturas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9375
   ClientLeft      =   3615
   ClientTop       =   645
   ClientWidth     =   9405
   Icon            =   "frmCapAperturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraCargoCta 
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
      Height          =   615
      Left            =   90
      TabIndex        =   127
      Top             =   5880
      Width           =   9210
      Begin SICMACT.ActXCodCta txtCuentaCargo 
         Height          =   375
         Left            =   0
         TabIndex        =   128
         Top             =   220
         Width           =   3630
         _extentx        =   6403
         _extenty        =   661
         texto           =   "Cuenta N°:"
         enabledcta      =   -1
         enabledage      =   -1
      End
      Begin VB.Label lblTitularCargoCta 
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4560
         TabIndex        =   130
         Top             =   240
         Width           =   4545
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Titular :"
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
         Left            =   3840
         TabIndex        =   129
         Top             =   280
         Width           =   675
      End
   End
   Begin VB.Frame fraPromotor 
      Height          =   615
      Left            =   90
      TabIndex        =   104
      Top             =   5880
      Width           =   9210
      Begin VB.ComboBox cboPromotor 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   195
         Width           =   6495
      End
      Begin VB.CheckBox chkPromotor 
         Caption         =   "Gestor de Cartera"
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
         Left            =   120
         TabIndex        =   107
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label Label28 
         Caption         =   "Nombre Gestor de Cartera:"
         Height          =   255
         Left            =   120
         TabIndex        =   108
         Top             =   270
         Width           =   1935
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
      Height          =   2145
      Left            =   4650
      TabIndex        =   36
      Top             =   6600
      Width           =   4650
      Begin VB.Frame FraITFAsume 
         Height          =   615
         Left            =   600
         TabIndex        =   110
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
         Begin VB.OptionButton OptAsuITF 
            Caption         =   "No Asume ITF"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   112
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton OptAsuITF 
            Caption         =   "Asume ITF"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   111
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fraespecialCTS 
         BorderStyle     =   0  'None
         Height          =   1860
         Left            =   120
         TabIndex        =   77
         Top             =   240
         Visible         =   0   'False
         Width           =   4125
         Begin VB.TextBox txtInta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Height          =   315
            Left            =   675
            TabIndex        =   80
            Text            =   "0.00"
            Top             =   525
            Width           =   1305
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Height          =   315
            Left            =   675
            TabIndex        =   79
            Text            =   "0.00"
            Top             =   960
            Width           =   1305
         End
         Begin VB.TextBox txtDU 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Height          =   315
            Left            =   675
            TabIndex        =   78
            Text            =   "0.00"
            Top             =   1875
            Width           =   1305
         End
         Begin VB.Label lblTotTran 
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
            Left            =   2355
            TabIndex        =   92
            Top             =   150
            Width           =   1665
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Monto Transacción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   30
            TabIndex        =   91
            Top             =   210
            Width           =   1650
         End
         Begin VB.Label lblDu 
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
            ForeColor       =   &H80000001&
            Height          =   270
            Left            =   3195
            TabIndex        =   90
            Top             =   1890
            Width           =   855
         End
         Begin VB.Label lblDisp 
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
            ForeColor       =   &H80000002&
            Height          =   285
            Left            =   3195
            TabIndex        =   89
            Top             =   1035
            Width           =   855
         End
         Begin VB.Label lblInta 
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
            ForeColor       =   &H80000002&
            Height          =   300
            Left            =   3195
            TabIndex        =   88
            Top             =   510
            Width           =   855
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Intang.(%)"
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
            Left            =   2175
            TabIndex        =   87
            Top             =   555
            Width           =   870
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Dispon.(%) "
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
            Left            =   2175
            TabIndex        =   86
            Top             =   1065
            Width           =   975
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "D.U. (%)"
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
            Left            =   2175
            TabIndex        =   85
            Top             =   1920
            Width           =   720
         End
         Begin VB.Label lblCMon 
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
            Left            =   2025
            TabIndex        =   84
            Top             =   150
            Width           =   300
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Intang."
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
            Left            =   30
            TabIndex        =   83
            Top             =   570
            Width           =   615
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Dispon. "
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
            Left            =   30
            TabIndex        =   82
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "D.U."
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
            Left            =   30
            TabIndex        =   81
            Top             =   1920
            Width           =   405
         End
      End
      Begin VB.CheckBox chkITFEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1350
         TabIndex        =   37
         Top             =   1485
         Width           =   705
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   375
         Left            =   1350
         TabIndex        =   38
         Top             =   1020
         Width           =   2025
         _extentx        =   3572
         _extenty        =   661
         font            =   "frmCapAperturas.frx":030A
         backcolor       =   12648447
         forecolor       =   12582912
         text            =   "0"
         enabled         =   -1
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
         Left            =   1350
         TabIndex        =   39
         Top             =   1755
         Width           =   2025
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   660
         TabIndex        =   46
         Top             =   1095
         Width           =   540
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3420
         TabIndex        =   45
         Top             =   1080
         Width           =   315
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
         ForeColor       =   &H80000002&
         Height          =   390
         Left            =   2400
         TabIndex        =   44
         Top             =   315
         Width           =   795
      End
      Begin VB.Label lblCTS 
         AutoSize        =   -1  'True
         Caption         =   "Disponib.Excedente (%) :"
         Height          =   195
         Left            =   435
         TabIndex        =   43
         Top             =   375
         Width           =   1920
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "ITF :"
         Height          =   195
         Left            =   660
         TabIndex        =   42
         Top             =   1470
         Width           =   330
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         Height          =   195
         Left            =   660
         TabIndex        =   41
         Top             =   1845
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
         Left            =   2130
         TabIndex        =   40
         Top             =   1425
         Width           =   1245
      End
   End
   Begin VB.CheckBox chkEspecial 
      Caption         =   "Especial"
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
      Left            =   7440
      TabIndex        =   93
      Top             =   885
      Visible         =   0   'False
      Width           =   1080
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
      Height          =   3015
      Left            =   90
      TabIndex        =   65
      Top             =   2880
      Width           =   9210
      Begin VB.Frame fraReglasPorderes 
         Caption         =   "Regla de Poderes"
         Height          =   1455
         Left            =   6600
         TabIndex        =   122
         Top             =   240
         Width           =   2445
         Begin VB.ListBox lsLetras 
            Height          =   1185
            Left            =   105
            Style           =   1  'Checkbox
            TabIndex        =   125
            Top             =   240
            Width           =   795
         End
         Begin VB.CommandButton cmdQuitarRega 
            Caption         =   "&Quitar"
            Height          =   375
            Left            =   1470
            TabIndex        =   124
            Top             =   1575
            Width           =   735
         End
         Begin VB.CommandButton cmdAgregarRegla 
            Caption         =   "Ag&regar"
            Height          =   375
            Left            =   525
            TabIndex        =   123
            Top             =   1575
            Width           =   735
         End
         Begin SICMACT.FlexEdit grdReglas 
            Height          =   1245
            Left            =   945
            TabIndex        =   126
            Top             =   240
            Width           =   1440
            _extentx        =   2540
            _extenty        =   2196
            highlight       =   1
            allowuserresizing=   3
            encabezadosnombres=   "#-Regla"
            encabezadosanchos=   "300-960"
            font            =   "frmCapAperturas.frx":0336
            font            =   "frmCapAperturas.frx":0362
            font            =   "frmCapAperturas.frx":038E
            font            =   "frmCapAperturas.frx":03BA
            font            =   "frmCapAperturas.frx":03E6
            fontfixed       =   "frmCapAperturas.frx":0412
            columnasaeditar =   "X-X"
            textstylefixed  =   4
            listacontroles  =   "0-0"
            encabezadosalineacion=   "C-C"
            formatosedit    =   "0-0"
            textarray0      =   "#"
            colwidth0       =   300
            rowheight0      =   300
         End
      End
      Begin VB.Frame fraPreferencial 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Enabled         =   0   'False
         Height          =   225
         Left            =   2115
         TabIndex        =   105
         Top             =   2640
         Width           =   2235
         Begin VB.CheckBox chkPermanente 
            Caption         =   "Es Tasa Permanente"
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
            Left            =   0
            TabIndex        =   106
            Top             =   0
            Visible         =   0   'False
            Width           =   2295
         End
      End
      Begin VB.CheckBox chkTasaPreferencial 
         Caption         =   " Tasa Preferencial"
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
         Left            =   120
         TabIndex        =   97
         Top             =   2640
         Width           =   1920
      End
      Begin VB.TextBox txtAlias 
         Height          =   330
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   76
         Top             =   2280
         Width           =   7215
      End
      Begin VB.TextBox txtNumSolicitud 
         Height          =   285
         Left            =   5640
         TabIndex        =   95
         Top             =   2640
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox TxtMinFirmas 
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
         Height          =   240
         Left            =   6120
         MaxLength       =   3
         TabIndex        =   74
         Top             =   1800
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   3200
         TabIndex        =   69
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   4200
         TabIndex        =   68
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtNumFirmas 
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
         Height          =   240
         Left            =   6360
         MaxLength       =   3
         TabIndex        =   67
         Top             =   1800
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.ComboBox cboTipoCuenta 
         Height          =   315
         Left            =   1155
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   1800
         Width           =   1920
      End
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1485
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   8925
         _extentx        =   15743
         _extenty        =   2619
         cols0           =   11
         highlight       =   1
         allowuserresizing=   3
         visiblepopmenu  =   -1
         encabezadosnombres=   "#-Codigo-Nombre-Relacion-P-Cta-Firma Oblig-Documento-Direccion-Grupo-otro"
         encabezadosanchos=   "250-1700-3500-1500-0-0-0-0-0-1000-0"
         font            =   "frmCapAperturas.frx":0440
         font            =   "frmCapAperturas.frx":0468
         font            =   "frmCapAperturas.frx":0490
         font            =   "frmCapAperturas.frx":04B8
         font            =   "frmCapAperturas.frx":04E0
         fontfixed       =   "frmCapAperturas.frx":0508
         lbultimainstancia=   -1
         tipobusqueda    =   3
         columnasaeditar =   "X-1-X-3-X-5-6-X-X-X-X"
         textstylefixed  =   4
         listacontroles  =   "0-1-0-3-0-4-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-L-C-C-C-L-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-0-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbflexduplicados=   0
         colwidth0       =   255
         rowheight0      =   300
      End
      Begin VB.Label lblEstadoSol 
         Alignment       =   2  'Center
         Caption         =   " ESTADO SOLICITUD"
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
         Height          =   270
         Left            =   6600
         TabIndex        =   96
         Top             =   2640
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label lblTitSol 
         AutoSize        =   -1  'True
         Caption         =   "Nro Solicitud"
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
         Left            =   4440
         TabIndex        =   94
         Top             =   2640
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Alias de Cuenta"
         Height          =   195
         Left            =   150
         TabIndex        =   75
         Top             =   2280
         Width           =   1110
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "N° Min. Firmas :"
         Height          =   315
         Left            =   6600
         TabIndex        =   73
         Top             =   1800
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "N° Firmas :"
         Height          =   195
         Left            =   7080
         TabIndex        =   72
         Top             =   1800
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   1800
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
      Height          =   2220
      Left            =   90
      TabIndex        =   47
      Top             =   -15
      Width           =   9210
      Begin VB.ComboBox cboPrograma 
         Height          =   315
         Left            =   3915
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   3345
      End
      Begin VB.Frame fraDatos 
         Height          =   1335
         Left            =   120
         TabIndex        =   48
         Top             =   795
         Width           =   8565
         Begin VB.CheckBox chkRelConv 
            Alignment       =   1  'Right Justify
            Caption         =   "Relacion con Convenio"
            Height          =   255
            Left            =   5565
            TabIndex        =   120
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.ComboBox cboInstConvDep 
            Height          =   315
            Left            =   4800
            Style           =   2  'Dropdown List
            TabIndex        =   119
            Top             =   600
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.CheckBox chkSubasta 
            Caption         =   "SUBASTA"
            Height          =   195
            Left            =   6705
            TabIndex        =   118
            Top             =   960
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkDepGar 
            Caption         =   "Deposito en Garantia"
            Height          =   255
            Left            =   6705
            TabIndex        =   114
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox chkAbonIntCta 
            Caption         =   "Abono Int.Cta.Aho."
            Height          =   255
            Left            =   3240
            TabIndex        =   113
            Top             =   960
            Width           =   1695
         End
         Begin SICMACT.EditMoney txtMontoAbonar 
            Height          =   300
            Left            =   6720
            TabIndex        =   102
            Top             =   600
            Width           =   1410
            _extentx        =   2487
            _extenty        =   529
            font            =   "frmCapAperturas.frx":052E
            forecolor       =   8388608
            text            =   "0.00"
            enabled         =   -1
         End
         Begin VB.Frame fraTasa 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   450
            Left            =   105
            TabIndex        =   98
            Top             =   540
            Width           =   2970
            Begin VB.ComboBox cboTipoTasa 
               Height          =   315
               Left            =   945
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   45
               Width           =   1920
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Tasa :"
               Height          =   195
               Left            =   0
               TabIndex        =   99
               Top             =   105
               Width           =   810
            End
         End
         Begin VB.ComboBox cboPeriodo 
            Height          =   315
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   240
            Width           =   2670
         End
         Begin VB.TextBox txtPlazo 
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
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   4350
            MaxLength       =   5
            TabIndex        =   53
            Text            =   "0"
            Top             =   585
            Width           =   1425
         End
         Begin VB.ComboBox cboFormaRetiro 
            Height          =   315
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   240
            Width           =   2355
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   180
            Width           =   1920
         End
         Begin VB.CheckBox chkOrdenPago 
            Alignment       =   1  'Right Justify
            Caption         =   "Orden Pago"
            Height          =   345
            Left            =   3210
            TabIndex        =   51
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox chkEmpCMACT 
            Alignment       =   1  'Right Justify
            Caption         =   "Emp. CMACC:"
            Height          =   315
            Left            =   3240
            TabIndex        =   49
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin SICMACT.TxtBuscar txtCtaAhoAboInt 
            Height          =   300
            Left            =   4920
            TabIndex        =   50
            Top             =   960
            Width           =   1755
            _extentx        =   3096
            _extenty        =   529
            appearance      =   1
            appearance      =   1
            font            =   "frmCapAperturas.frx":055A
            appearance      =   1
         End
         Begin SICMACT.TxtBuscar txtInstitucion 
            Height          =   330
            Left            =   5640
            TabIndex        =   121
            Top             =   600
            Width           =   1980
            _extentx        =   3493
            _extenty        =   582
            appearance      =   1
            appearance      =   1
            font            =   "frmCapAperturas.frx":0586
            appearance      =   1
            tipobusqueda    =   3
            tipobuspers     =   1
         End
         Begin VB.Label label14 
            AutoSize        =   -1  'True
            Caption         =   "Abono:"
            Height          =   315
            Left            =   5880
            TabIndex        =   103
            Top             =   600
            Width           =   510
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Left            =   105
            TabIndex        =   63
            Top             =   240
            Width           =   675
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
            Left            =   1050
            TabIndex        =   62
            Top             =   990
            Width           =   1905
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Forma Retiro :"
            Height          =   195
            Left            =   3210
            TabIndex        =   61
            Top             =   270
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (días) :"
            Height          =   195
            Left            =   3240
            TabIndex        =   60
            Top             =   645
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tasa EA (%) :"
            Height          =   195
            Left            =   105
            TabIndex        =   59
            Top             =   1050
            Width           =   960
         End
         Begin VB.Label lblPeriodo 
            AutoSize        =   -1  'True
            Caption         =   "Período"
            Height          =   195
            Left            =   3225
            TabIndex        =   58
            Top             =   270
            Width           =   570
         End
         Begin VB.Label lblInstitucion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3225
            TabIndex        =   57
            Top             =   945
            Width           =   4395
         End
         Begin VB.Label lblInst 
            AutoSize        =   -1  'True
            Caption         =   "Institución :"
            Height          =   195
            Left            =   4740
            TabIndex        =   56
            Top             =   645
            Width           =   810
         End
         Begin VB.Label lblCuentaAbo 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Abono :"
            Height          =   195
            Left            =   3915
            TabIndex        =   55
            Top             =   1005
            Width           =   1110
         End
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   150
         TabIndex        =   64
         Top             =   330
         Width           =   3630
         _extentx        =   6403
         _extenty        =   661
         texto           =   "Cuenta N°:"
      End
      Begin VB.Label lblCampana 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   7320
         TabIndex        =   131
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Sub Producto:"
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
         Left            =   3945
         TabIndex        =   100
         Top             =   135
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin VB.Frame fraITF 
      Height          =   615
      Left            =   90
      TabIndex        =   32
      Top             =   2220
      Width           =   9210
      Begin VB.CheckBox chkExoITF 
         Caption         =   "Exonerado ITF"
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
         Height          =   225
         Left            =   120
         TabIndex        =   34
         Top             =   0
         Width           =   1590
      End
      Begin VB.ComboBox cboTipoExoneracion 
         Height          =   315
         Left            =   1995
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   210
         Width           =   6600
      End
      Begin VB.Label Label17 
         Caption         =   "Tipo de Exoneracion :"
         Height          =   225
         Left            =   315
         TabIndex        =   35
         Top             =   270
         Width           =   1620
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   90
      TabIndex        =   8
      Top             =   8880
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   8880
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   8880
      Width           =   1000
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
      Height          =   2145
      Left            =   120
      TabIndex        =   18
      Top             =   6600
      Width           =   4530
      Begin VB.TextBox txtTransferGlosa 
         Height          =   360
         Left            =   825
         MaxLength       =   255
         TabIndex        =   21
         Top             =   1290
         Width           =   3465
      End
      Begin VB.CommandButton cmdTranfer 
         Height          =   350
         Left            =   2520
         Picture         =   "frmCapAperturas.frx":05B2
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   480
         Width           =   475
      End
      Begin VB.ComboBox cboTransferMoneda 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   165
         Width           =   1575
      End
      Begin VB.Label lblEtiMonTra 
         AutoSize        =   -1  'True
         Caption         =   "Monto Transacción"
         Height          =   195
         Left            =   960
         TabIndex        =   117
         Top             =   1740
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lblSimTra 
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
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   2400
         TabIndex        =   116
         Top             =   1710
         Visible         =   0   'False
         Width           =   240
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
         TabIndex        =   115
         Top             =   1680
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblTTCVD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   31
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblTTCCD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   30
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "TCV"
         Height          =   285
         Left            =   3120
         TabIndex        =   29
         Top             =   480
         Width           =   390
      End
      Begin VB.Label lblTTCC 
         Caption         =   "TCC"
         Height          =   285
         Left            =   3120
         TabIndex        =   28
         Top             =   180
         Width           =   390
      End
      Begin VB.Label lblTransferGlosa 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   1290
         Width           =   495
      End
      Begin VB.Label lblTransferMoneda 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   30
         TabIndex        =   26
         Top             =   225
         Width           =   585
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
         TabIndex        =   25
         Top             =   525
         Width           =   1575
      End
      Begin VB.Label lbltransferBcol 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   90
         TabIndex        =   23
         Top             =   975
         Width           =   555
      End
      Begin VB.Label lbltransferN 
         AutoSize        =   -1  'True
         Caption         =   "Nro Doc :"
         Height          =   195
         Left            =   60
         TabIndex        =   22
         Top             =   600
         Width           =   690
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
         TabIndex        =   24
         Top             =   900
         Width           =   3465
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
      Height          =   2145
      Left            =   90
      TabIndex        =   9
      Top             =   6600
      Width           =   4530
      Begin VB.ComboBox cboUsuRef 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1740
         Visible         =   0   'False
         Width           =   3060
      End
      Begin VB.CommandButton cmdDocumento 
         Height          =   350
         Left            =   2745
         Picture         =   "frmCapAperturas.frx":09F4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   475
      End
      Begin VB.TextBox txtGlosa 
         Height          =   600
         Left            =   825
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1050
         Width           =   3585
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nro Doc :"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   675
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
         Left            =   840
         TabIndex        =   15
         Top             =   600
         Width           =   3375
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
         Left            =   840
         TabIndex        =   14
         Top             =   225
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Usuario Ref."
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1050
         Width           =   495
      End
   End
   Begin VB.PictureBox pctCheque 
      Height          =   390
      Left            =   3195
      Picture         =   "frmCapAperturas.frx":0E36
      ScaleHeight     =   330
      ScaleWidth      =   375
      TabIndex        =   12
      Top             =   7800
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox pctNotaAbono 
      Height          =   360
      Left            =   2520
      Picture         =   "frmCapAperturas.frx":1508
      ScaleHeight     =   300
      ScaleWidth      =   330
      TabIndex        =   13
      Top             =   7800
      Visible         =   0   'False
      Width           =   390
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   330
      Left            =   0
      TabIndex        =   101
      Top             =   1080
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   582
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmCapAperturas.frx":1A8A
   End
End
Attribute VB_Name = "frmCapAperturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nProducto As COMDConstantes.Producto
Dim nTitular As Integer
Dim nRepresentante As Integer
Dim nClientes As Integer
Dim nTipoCuenta As COMDConstantes.ProductoCuentaTipo
Dim nmoneda As Moneda
Dim nTipoTasa As COMDConstantes.CaptacTipoTasa
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim bDocumento As Boolean
Dim nDocumento As COMDConstantes.TpoDoc
Dim nPersoneria As COMDConstantes.PersPersoneria
Public dFechaValorizacion As Date
Public sCodIF As String
Dim vbDesembolso As Boolean
'Para Desembolso Con Apertura de Cuenta
Dim vMatRela As ADODB.Recordset
Dim vnTasa As Integer
Dim vnPersoneria As Integer
Dim vnTipoCuenta As Integer
Dim vnTipoTasa As Integer
Dim vbDocumento As Boolean
Dim vsNroDoc As String
Dim vsCodIF As String
Dim lbImpRegFirma As Byte
Dim sOperacion As String
Dim nTasaNominal As Double
Dim cPersTasaEspecial As String
Dim vnMontoDOC As Double
Dim vSperscod As String
Dim sPerSolicitud As String 'MODIFICADO POR "RIRO" EL 27/11/2012


'Transferencia
Dim lnMovNroTransfer As Long
Dim lnTransferSaldo As Currency
Dim fsPersCodTransfer As String '***Agregado por ELRO el 20120706, según OYP-RFC024-2012
Dim fsOpeCod As String '***Agregado por ELRO el 20120706, según OYP-RFC024-2012
Dim fnMovNroRVD As Long '***Agregado por ELRO el 20120706, según OYP-RFC024-2012

'***************Variabres Agregadas********************
Dim Gtitular As String
Dim GAutNivel As String
Dim GAutMontoFinSol As Double
Dim GAutMontoFinDol As Double
Dim GMontoAprobado As Double, GNroID As Long, GPersCod As String

'Variables para la impresion de la boleta de Lavado de Dinero

Dim sPersCod As String, sDocId As String, sDireccion As String
Dim sPersCodRea As String, sNombreRea As String, sDocIdRea As String, sDireccionRea As String
Dim sTipoCuenta As String
Dim sNombre As String

'Variables ITF
Dim lbITFCtaExonerada As Boolean

'Variable de Impresion
Dim nFicSal As Integer

'Dias de valorizacion
Public lnDValoriza As Integer
Dim nPlazoVal As Long

Dim sPersVistoCod As String '*** PEAC 20080807
Dim sPersVistoCom As String '*** PEAC 20080807

Dim lnValOpePF As Integer 'Add By GITU 20100806
Dim lnTitularPJ As Integer 'Add By GITU 20100809

'ARCV 12-02-2007
Dim vMatTitular As Variant
Dim vnPrograma  As Integer
Dim vnMontoAbonar As Double
Dim vnPlazoAbono As Integer
Dim vsPromotor As String
'----------------
'By Capi 20042008
Dim lnTotIntMes As Double
' Brgo 20110908
Dim nRedondeoITF As Double
Dim nTpoProgramaCTS As Integer
            
Dim nTasaNominalTemp As Double 'MADM 20111022
Dim nTasaEfectivaTemp As Double 'MADM 20111022
Dim nMontoMinimoPFPremium As Double ' BRGO 20111219
Dim lnTpoPrograma As Integer
Dim fnTpoCtaCargo As Integer 'JUEZ 20131212
Dim rsRelPersCtaCargo As ADODB.Recordset 'JUEZ 20131312
Dim sMovNroAut As String 'JUEZ 20131212
Dim bInstFinanc As Boolean 'JUEZ 20140414

'RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
Dim intPunteroPJ_NA As Integer 'Si intPunteroPJ_NA=0 --> No tiene PJ ; intPunteroPJ_NA>=1 ----> Si tiene PJ
Dim oDocRec As UDocRec 'EJVG20140408

'JUEZ 20141008 Nuevos Parametros **********
Dim bParPersJur As Boolean
Dim bParPersNat As Boolean
Dim bParMonedaSol As Boolean
Dim bParMonedaDol As Boolean
Dim nParMontoMinSol As Double
Dim nParMontoMinDol As Double
Dim nParOrdPag As Integer
Dim nParPlazoMin As Long
Dim nParPlazoMax As Long
Dim bParFormaRetFinPlazo As Boolean
Dim bParFormaRetMensual As Boolean
Dim bParFormaRetIniPlazo As Boolean
Dim nParAumCapMinSol As Double
Dim nParAumCapMinDol As Double
'END JUEZ *********************************
'JUEZ 20160420 ********************
Dim fnCampanaCod As Long
Dim fsCampanaDesc As String
'END JUEZ *************************

Private Sub EmiteCalendarioRetiroIntPFMensual(ByVal nCapital As Double, ByVal nTasa As Double, ByVal nPlazo As Long, _
            ByVal dApertura As Date, ByVal nmoneda As Moneda, Optional ByVal nDiasVal As Integer = 0, Optional sCuenta As String = "", Optional nCostoMan As Currency = 0)

Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim nIntMens As Double, nIntFinal As Double
Dim dFecVenc As Date, dFecVal As Date
    
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
dFecVenc = DateAdd("d", nPlazo + nDiasVal, dApertura)
dFecVal = DateAdd("d", nDiasVal, dApertura)
nIntMens = clsMant.GetInteresPF(nTasa, nCapital, 30)
nIntFinal = clsMant.GetInteresPF(nTasa, nCapital, nPlazo)

Set clsMant = Nothing

Dim clsPrev As previo.clsprevio
Dim sCad As String
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
'By Capi 20042008
'sCad = clsMant.GetPFPlanRetInt(dApertura, Round(nIntMens, 2), nPlazo, nmoneda, Round(nIntFinal, 2), nCapital, nTasa, nDiasVal, dFecVal)
sCad = clsMant.GetPFPlanRetInt(dApertura, Round(nIntMens, 2), nPlazo, nmoneda, Round(nIntFinal, 2), nCapital, nTasa, nDiasVal, dFecVal, lnTotIntMes, sCuenta, nCostoMan)
    
Set clsMant = Nothing

Set clsPrev = New previo.clsprevio
    'ALPA 20100202***********************
    'clsPrev.Show sCad, "Plazo Fijo"
    clsPrev.Show sCad, "Plazo Fijo", True, , gImpresora
Set clsPrev = Nothing
End Sub

Private Function ValidaTasaInteres()
Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales 'BRGO 20111020
Dim bOrdPag As Boolean
Dim nMonto As Double
Dim nPlazo As Long
Dim nTpoPrograma As Integer
Dim sTitular As String
Dim rsCamp As ADODB.Recordset 'JUEZ 20160420

Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales 'BRGO 20111020
bOrdPag = IIf(chkOrdenPago.value = 1, True, False)
nMonto = txtMonto.value
nTpoPrograma = 0

If cboPrograma.Visible Then
    nTpoPrograma = CInt(Right(Trim(cboPrograma.Text), 2))
End If

If chkTasaPreferencial.value = vbUnchecked Then
    nTipoTasa = gCapTasaNormal
    If nProducto <> gCapCTS Then 'JUEZ 20160420
        fnCampanaCod = 0
        nTasaNominal = clsDef.GetCapTasaInteresCamp(nProducto, nTpoPrograma, nmoneda, IIf(txtPlazo <> "", CLng(txtPlazo), 0), nMonto, gsCodAge, gdFecSis, IIf(nPersoneria <> gPersonaNat Or lnTitularPJ = 1, True, False), bOrdPag, fnCampanaCod, fsCampanaDesc)
    End If
    If nProducto = gCapPlazoFijo Then
        If txtPlazo <> "" Then
            nPlazo = CLng(txtPlazo)
            'JUEZ 20160420 TASAS CAMPAÑA ************************************************
            If nTasaNominal = 0 Then
                'Add by Gitu 2010-08-06
                If lnValOpePF = 1 And (nPersoneria <> gPersonaNat Or lnTitularPJ = 1) Then
                    If chkDepGar.value = 1 Then 'MADM 20111022
                        'nTasaNominal = (clsDef.GetCapTasaInteresPF(gCapPlazoFijo, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma) / 2)
                        nTasaNominalTemp = clsDef.GetCapTasaInteresPF(gCapPlazoFijo, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma)
                        nTasaEfectivaTemp = Format$(ConvierteTNAaTEA(nTasaNominalTemp), "#,##0.00") / 2
                        nTasaNominal = Format$(ConvierteTEAaTNA(nTasaEfectivaTemp), "#,##0.00")
                        'Add By Gitu 18-04-2013
                        If nTasaEfectivaTemp > 1 Then
                            nTasaEfectivaTemp = 1#
                            nTasaNominal = Format$(ConvierteTEAaTNA(nTasaEfectivaTemp), "#,##0.0000")
                        End If
                        'End GITU
                    Else
                         nTasaNominal = clsDef.GetCapTasaInteresPF(gCapPlazoFijo, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma)
                    End If
                Else
                    If chkDepGar.value = 1 Then 'MADM 20111022
                        'nTasaNominal = (clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma) / 2)
                        nTasaNominalTemp = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma)
                        nTasaEfectivaTemp = Format$(ConvierteTNAaTEA(nTasaNominalTemp), "#,##0.00") / 2
                        nTasaNominal = Format$(ConvierteTEAaTNA(nTasaEfectivaTemp), "#,##0.00")
                        'Add By Gitu 18-04-2013
                        If nTasaEfectivaTemp > 1 Then
                            nTasaEfectivaTemp = Format$(1, "#,##0.00")
                            nTasaNominal = Format$(ConvierteTEAaTNA(nTasaEfectivaTemp), "#,##0.0000")
                        End If
                        'End GITU
                    Else
                        nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma)
                    End If
                End If
                'End Gitu
            End If
            'END JUEZ TASAS CAMPAÑA *****************************************************
            lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
        End If
    ElseIf nProducto = gCapAhorros Then
        'JUEZ 20160420 TASAS CAMPAÑA ************************************************
        If nTasaNominal = 0 Then
            nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, bOrdPag, nTpoPrograma)
        End If
        'END JUEZ TASAS CAMPAÑA *****************************************************
        lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
    Else
        nTpoProgramaCTS = 1 'Por defecto se asigna Tasa de CTS sin Cta Sueldo
        sTitular = ObtTitular
        If sTitular <> "" Then
            'JUEZ 20140319 *****************************
            If Trim(txtInstitucion.Text) = "" Then
                txtInstitucion.SetFocus
                lblTasa.Caption = "0.00"
                Exit Function
            End If
            'END JUEZ **********************************
            If clsCap.TieneCuentasCaptacxSubProducto(sTitular, gCapAhorros, 6, Trim(txtInstitucion.Text)) Then 'Verifica si Cliente tiene Cta Sueldo 'JUEZ 20140319 Se agregó Trim(txtInstitucion.Text)
                nTpoProgramaCTS = 0 ' Si tiene CtaSueldo se cambia de Sub Producto
            End If
        End If
        'JUEZ 20160420 TASAS CAMPAÑA ************************************************
        fnCampanaCod = 0
        nTasaNominal = clsDef.GetCapTasaInteresCamp(nProducto, nTpoProgramaCTS, nmoneda, nPlazo, nMonto, gsCodAge, gdFecSis, IIf(nPersoneria <> gPersonaNat Or lnTitularPJ = 1, True, False), bOrdPag, fnCampanaCod, fsCampanaDesc)
        If nTasaNominal = 0 Then
            nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoProgramaCTS)
        End If
        'END JUEZ TASAS CAMPAÑA *****************************************************
        lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
    End If
End If
lblCampana.Caption = "" 'JUEZ 20160420
If fnCampanaCod <> 0 Then lblCampana.Caption = "TASA CAMPAÑA: " & fsCampanaDesc 'JUEZ 20160420
Set clsDef = Nothing
End Function

Private Function EsExoneradaLavadoDinero() As Boolean
Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim bExito As Boolean
Dim clsExo As COMNCaptaServicios.NCOMCaptaServicios 'NCapServicios
bExito = True
Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
For i = 1 To grdCliente.Rows - 1
    nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
    If nRelacion = gCapRelPersTitular Then
        sPersCod = grdCliente.TextMatrix(i, 1)
        If Not clsExo.EsPersonaExoneradaLavadoDinero(sPersCod) Then
            bExito = False
            Exit For
        End If
    End If
Next i
Set clsExo = Nothing
EsExoneradaLavadoDinero = bExito
End Function

Private Sub MarcaSoloUnaFila(ByVal nFilaMarcada As Long)
Dim i As Long
For i = 1 To grdCliente.Rows - 1
    If i <> nFilaMarcada Then
        grdCliente.TextMatrix(i, 5) = ""
    End If
Next i
End Sub

Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero)
Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim nMonto As Double
Dim oPersona As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsPers As New ADODB.Recordset

For i = 1 To grdCliente.Rows - 1
    nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
    If nPersoneria = gPersonaNat Then
        If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            poLavDinero.TitPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDinero.TitPersLavDineroNom = grdCliente.TextMatrix(i, 2)
            Exit For
        End If
    Else
        If nRelacion = gCapRelPersTitular Then
            poLavDinero.TitPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDinero.TitPersLavDineroNom = grdCliente.TextMatrix(i, 2)
        End If
        If nRelacion = gCapRelPersRepTitular Then
            poLavDinero.TitPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDinero.ReaPersLavDineroNom = grdCliente.TextMatrix(i, 2)
            If poLavDinero.TitPersLavDinero <> "" Then Exit For
        End If
    End If
Next i
nMonto = txtMonto.value
sTipoCuenta = cboTipoCuenta.Text
End Sub

Private Function ValidaRelaciones() As Boolean
Dim i As Long
For i = 1 To grdCliente.Rows - 1
    If Trim(grdCliente.TextMatrix(i, 3)) = "" Then
        MsgBox "Debe registrar la relación con el cliente.", vbInformation, "Aviso"
        grdCliente.row = i
        grdCliente.Col = 3
        grdCliente.SetFocus
        SendKeys "{Enter}"
        ValidaRelaciones = False
        Exit Function
    End If
Next i
ValidaRelaciones = True
End Function

Private Function ValidaUsuarios() As Boolean
Dim i As Long
Dim dlsMant As COMDCaptaGenerales.DCOMCaptaGenerales 'DCapMantenimiento
Set dlsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
If dlsMant.GetNroOPeradoras(gsCodAge) > 1 Then

        For i = 1 To grdCliente.Rows - 1
            If Trim(grdCliente.TextMatrix(i, 1)) = gsCodPersUser Then
                 ValidaUsuarios = False
                 Exit Function
            End If
        Next i
        ValidaUsuarios = True
Else
        ValidaUsuarios = True
End If
Set dlsMant = Nothing
End Function

Private Function GetDireccionCliente() As String
Dim sDireccion As String
Dim i As Integer
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsPers As New ADODB.Recordset
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
sDireccion = ""
For i = 1 To grdCliente.Rows - 1
    If CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4))) = gCapRelPersTitular Then
        Set rsPers = clsMant.GetDatosPersona(grdCliente.TextMatrix(i, 1))
        sDireccion = Trim(rsPers("Direccion"))
        Exit For
    End If
Next i
Set clsMant = Nothing
GetDireccionCliente = sDireccion
End Function
'***Modificado por ELRO el 20120124, según Acta N° 006-2012/TI-D
'Private Sub EmiteCertificadoPlazoFijo(ByVal sCuenta As String,
'                                      ByVal Rsr As ADODB.Recordset,
'                                      Optional nCostoMan As Currency = 0)
Private Sub EmiteCertificadoPlazoFijo(ByVal sCuenta As String, _
                                      ByVal Rsr As ADODB.Recordset, _
                                      Optional nCostoMan As Currency = 0, _
                                      Optional pnFormaRetiro As Long = 0)
'***Fin por ELRO************************************************

Dim bReImp As Boolean
Dim sNomTit As String, sDirCli As String
Dim nMonto As Double
Dim lsCadImp As String
Dim lcCapImp As COMNCaptaGenerales.NCOMCaptaImpresion
Dim sFormaRetiro As String, rsPer As New ADODB.Recordset
Dim i As Integer

'Set rsPer = New ADODB.Recordset

If MsgBox("¿ Desea Imprimir el Certificado de PF ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim sLetras As String
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    'sNomTit = clsMant.GetNombreTitulares(sCuenta, True, 3, 0)
    Set rsPer = clsMant.GetPersonaCuenta(sCuenta, 10)
        sNomTit = ""
    While Not rsPer.EOF
        sNomTit = sNomTit & Left(Trim(rsPer.Fields("NOMBRE")), 60) & Space(2) & Left(Trim(rsPer.Fields("ID N°")), 15) & Space(2) & Left(Trim(rsPer.Fields("Direccion")), 25) & Chr(10)
        rsPer.MoveNext
    Wend
    Set rsPer = Nothing
    
    Set clsMant = Nothing
    sDirCli = GetDireccionCliente
    
    nMonto = IIf(chkITFEfectivo.value = 1, txtMonto.value, txtMonto.value - CDbl(LblItf.Caption))
    sLetras = ConversNL(nmoneda, nMonto)
    '***Modificado por ELRO el 20120124, según Acta N° 006-2012/TI-D
    'sFormaRetiro = Trim(Left(cboFormaRetiro.Text, 25))
    If pnFormaRetiro = 0 Then
        sFormaRetiro = ""
    Else
    
        For i = 0 To CInt(cboFormaRetiro.ListCount) - 1
            cboFormaRetiro.ListIndex = i
            If CLng(Trim(Right(cboFormaRetiro.Text, 4))) = pnFormaRetiro Then
                sFormaRetiro = Trim(Left(cboFormaRetiro.Text, 25))
                Exit For
            End If
        Next i
    End If
    '***Fin por ELRO************************************************
    bReImp = False
    Set lcCapImp = New COMNCaptaGenerales.NCOMCaptaImpresion
    lcCapImp.IniciaImpresora gImpresora
    '***Modificado por ELRO el 20120124, según Acta N° 006-2012/TI-D
    'lsCadImp = lcCapImp.ImprimeCertificadoPlazoFijo(gdFecSis, sNomTit, sDirCli, sCuenta, "1", CLng(val(nPlazoVal)), nMonto, nTasaNominal, sFormaRetiro, sLetras, , False, Trim(Left(cboTipoCuenta.Text, 20)), Rsr, gsNomAge, cboFormaRetiro.Text, nCostoMan)
     lsCadImp = lcCapImp.ImprimeCertificadoPlazoFijo(gdFecSis, sNomTit, _
                                                     sDirCli, sCuenta, _
                                                     "1", CLng(Val(nPlazoVal)), _
                                                     nMonto, nTasaNominal, _
                                                     sFormaRetiro, sLetras, , _
                                                     False, _
                                                     Trim(Left(cboTipoCuenta.Text, 20)), _
                                                     Rsr, gsNomAge, _
                                                     sFormaRetiro, nCostoMan)
    '***Fin por ELRO************************************************
    Set lcCapImp = Nothing
    Do
         If Trim(lsCadImp) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
               Print #nFicSal, lsCadImp & Chr$(12)
               Print #nFicSal, ""
            Close #nFicSal
         End If
         
         If MsgBox("Desea reimprimir Certificado PF?? ", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            bReImp = True
         Else
            bReImp = False
         End If
    Loop Until Not bReImp
    
End Sub

Private Sub EmiteBoleta(ByVal sCuenta As String, ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double, ByVal psBoletaCargo As String)
Dim bReImp As Boolean
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String, sNroDoc As String
Dim bProd As Boolean
Dim lsCadImp As String
Dim nTipoPag As Integer
Dim nMontoAper As Double

'PASI20140530
Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
'end PASI


nMontoAper = nSaldoCnt
sTipDep = Trim(Left(cboMoneda.Text, 15))
sCodOpe = Trim(nOperacion)
If bDocumento Then
    sModDep = "Depósito Cheque"
    sNroDoc = lblNroDoc.Caption
Else
    If nOperacion = gAhoApeTransf Or nOperacion = gPFApeTransf Or nOperacion = gCTSApeTransf Then
        sModDep = "Depósito Transferencia"
    'JUEZ 20131212 **************************************************
    ElseIf nOperacion = gAhoApeCargoCta Or nOperacion = gPFApeCargoCta Then
        sModDep = "Depósito Cargo a Cuenta"
    'END JUEZ *******************************************************
    Else
        sModDep = "Depósito Efectivo"
    End If
End If
Select Case nProducto
    Case gCapAhorros
        bProd = gITF.gbITFAsumidoAho
        If chkOrdenPago.value = 1 Then
            sTipApe = "APERTURA AHORROS CON OP"
        Else
            sTipApe = "APERTURA AHORROS"
        End If
    Case gCapPlazoFijo
        bProd = gITF.gbITFAsumidoPF
        sTipApe = "APERTURA PLAZO FIJO"
    Case gCapCTS
        bProd = True
        sTipApe = "APERTURA CTS"
End Select


'************MODIFICACION MPBR ****************
Dim MontoItf As Double

MontoItf = 0

'If gbITFAplica And nProducto <> gCapCTS And bProd = False Then        'Filtra para CTS
If gbITFAplica And nProducto <> gCapCTS And bProd = False And sCodOpe <> gAhoApeCargoCta And sCodOpe <> gPFApeCargoCta Then 'JUEZ 20131212 Para exonerar de ITF para aperturas con cargo a cuenta
    If txtMonto.value > gnITFMontoMin And chkExoITF.value = 0 Then
        If chkITFEfectivo.value = vbUnchecked Then
            MontoItf = LblItf.Caption
        End If
    End If
    nSaldoCnt = nSaldoCnt - MontoItf
    nSaldoDisp = nSaldoDisp - MontoItf
End If
'*********************MODIF MPBR

Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    sNomTit = ImpreCarEsp(clsMant.GetNombreTitulares(sCuenta))
Set clsMant = Nothing
Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
bReImp = False
'****** validacion CTS *****
If nProducto = gCapCTS Then
    If chkEspecial.value = 1 Then
        nSaldoDisp = txtDisp.Text
    Else
        nSaldoDisp = txtMonto.Text * CDbl(lblDispCTS) / 100
    End If
End If
'***************************

'Obtener el Tipo de pago de ITF
If chkITFEfectivo.value = 1 Then
    nTipoPag = 1
Else
    nTipoPag = 2
End If

    If bDocumento Then
        If nDocumento = TpoDocCheque Then
            dFechaValorizacion = oNCapMov.ObtenerFechaValorizaCheque(oDocRec.fsNroDoc, oDocRec.fsPersCod, oDocRec.fsIFTpo, oDocRec.fsIFCta) 'PASI20140530
            lsCadImp = clsCap.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(txtMonto.Text), sNomTit, sCuenta, Format$(dFechaValorizacion, "dd/mm/yyyy"), nSaldoDisp, 0, "Fecha Valor", 1, nSaldoCnt, , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, False, gsCodCMAC, , , , , , , , , , , , True, nTipoPag, CDbl(LblItf.Caption), True, , , gbImpTMU)
        ElseIf nDocumento = TpoDocNotaAbono Then
            lsCadImp = clsCap.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(txtMonto.Text), sNomTit, sCuenta, "", nSaldoDisp, 0, "", 1, nSaldoCnt, , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodCMAC, , , , , , , , , , , , True, nTipoPag, CDbl(LblItf.Caption), True, , , gbImpTMU)
        End If
    Else
        lsCadImp = clsCap.ImprimeBoleta(sTipApe, ImpreCarEsp(sModDep), sCodOpe, Format$(nMontoAper, "#,##0.00"), sNomTit, sCuenta, "", nSaldoDisp, 0, "", 1, nSaldoCnt, , , , , , , , , , gdFecSis, Trim(gsNomAge), gsCodUser, sLpt, False, gsCodCMAC, , , , , , , , , , , , True, nTipoPag, CDbl(LblItf.Caption), True, , , gbImpTMU)
    End If
    
If nOperacion = gAhoApeCargoCta Or nOperacion = gPFApeCargoCta Then lsCadImp = lsCadImp & psBoletaCargo 'JUEZ 20131212

Do
    If Trim(lsCadImp) <> "" Then
       nFicSal = FreeFile
       Open sLpt For Output As nFicSal
          Print #nFicSal, lsCadImp
          Print #nFicSal, ""
       Close #nFicSal
     End If
Loop Until MsgBox("Desea reimprimir Boleta?? ", vbQuestion + vbYesNo, Me.Caption) = vbNo

Set clsCap = Nothing
End Sub

Private Sub LimpiaControles()
    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral
    '***Modificado por ELRO el 20121015, según OYP-RFC024-2012
    'Me.cboTransferMoneda.Enabled = True
    cboTransferMoneda.Enabled = False
    '***Fin Modificado por ELRO el 20121015*******************
    Me.txtTransferGlosa.Text = ""
    Me.lbltransferBco.Caption = ""
    Me.lblTrasferND.Caption = ""
    lnMovNroTransfer = -1
    lnTransferSaldo = 0
    '***Agregado por ELRO el 20120821, según OYP-RFC024-2012
    fsPersCodTransfer = ""
    fsOpeCod = ""
    fnMovNroRVD = 0
    '***Fin Agregado por ELRO 20120821**********************
    txtCtaAhoAboInt.Visible = False
    lblCuentaAbo.Visible = False
    grdCliente.Clear
    grdCliente.Rows = 2
    grdCliente.FormaCabecera
    txtGlosa = ""
    txtPlazo = "0"
    txtNumFirmas = "0"
    txtMonto.Enabled = True
    txtMonto.value = 0
    'chkOrdenPago.value = 0
    chkOrdenPago.value = IIf(nParOrdPag = 1, 1, 0) 'JUEZ 20141008
    cboMoneda.Enabled = True
    cboMoneda.ListIndex = 0
    cboTipoTasa.ListIndex = 0
    cboTipoCuenta.ListIndex = 0
    txtInstitucion.Text = ""
    lblInstitucion.Caption = ""
    'cboUsuRef.ListIndex = 0
    txtAlias.Text = ""
    TxtMinFirmas.Text = "0"
    Me.chkTasaPreferencial.value = vbUnchecked
    vSperscod = ""
    lblEstadoSol.Caption = "ESTADO SOLICITUD"
    cboPromotor.ListIndex = 0
    cboPromotor.Enabled = False
    chkPromotor.value = 0
    OptAsuITF(0).value = False
    OptAsuITF(1).value = False
    chkAbonIntCta.Visible = False
    '***Agregado por ELRO el 20120326, por incidente INC1203200006
    lblTasa.Caption = "0.00"
    '***Fin Agregado por ELRO*************************************
    '***Agregado por ELRO el 20120403
    chkEmpCMACT.value = 0
    '***Fin Agregado por ELRO********
    Select Case nProducto
        Case gCapAhorros
            txtCuenta.Prod = Trim(Str(gCapAhorros))
            grdCliente.ColWidth(5) = 0
            'grdCliente.ColWidth(6) = 1200 Comentado por RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            txtAlias.Visible = True
            'TxtMinFirmas.Visible = True Comentado por RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            txtMontoAbonar.Text = "0"
            chkRelConv.value = 0
            cboInstConvDep.ListIndex = 0
            If nOperacion = gAhoApeCargoCta Then LimpiaControlesCargoCta 'JUEZ 20131212
        Case gCapPlazoFijo
            txtCuenta.Prod = Trim(Str(gCapPlazoFijo))
            cboFormaRetiro.ListIndex = 0
            txtCtaAhoAboInt.Text = ""
            'grdCliente.ColWidth(6) = 1200 Comentado por RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            txtAlias.Visible = True
            'TxtMinFirmas.Visible = True Comentado por RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            lnTitularPJ = 0
            chkAbonIntCta.Visible = True
            chkAbonIntCta.value = vbUnchecked
            grdCliente.ColWidth(5) = 0
            If cboPrograma.ListIndex = 1 Then
                'txtPlazo.Text = loRs.GetParametro(2000, 2120)
                txtPlazo.Text = nParPlazoMin 'JUEZ 20141008 Nuevos Parámetros
                chkAbonIntCta.Visible = False
            End If
            txtMontoAbonar.Text = "0.00"
            If nOperacion = gPFApeCargoCta Then LimpiaControlesCargoCta 'JUEZ 20131212
        Case gCapCTS
            txtCuenta.Prod = Trim(Str(gCapCTS))
            'cboPeriodo.ListIndex = 0
            cboPeriodo.ListIndex = 4
            grdCliente.ColWidth(5) = 0
            grdCliente.ColWidth(6) = 0
            txtAlias.Visible = False
            TxtMinFirmas.Visible = False
            'chkEspecial.value = vbUnchecked 'JUEZ 20130814
            Me.fraespecialCTS.Visible = False
            Me.lblTotTran.Caption = "0.00"
            Me.txtDisp.Text = "0.00"
            Me.txtInta.Text = "0.00"
            Me.txtDU.Text = "0.00"
            Me.lblDisp.Caption = "0.00"
            Me.lblDu.Caption = "0.00"
            Me.lblInta.Caption = "0.00"
            nTpoProgramaCTS = 0
            cmdAgregar.Enabled = True 'Agregado por RIRO 20130411
            
    End Select
    If bDocumento Then
        lblNroDoc.Caption = ""
        lblNombreIF.Caption = ""
    End If
    cmdEliminar.Enabled = False
    nPersoneria = gPersonaNat
    nTitular = 0
    nClientes = 0
    txtCuenta.Cuenta = ""
    sTipoCuenta = ""
    
    'RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
    limpiarReglas
    sinReglas
    chkDepGar.value = 0
    'Fin RIRO
    
    ' Por RIRO el 20130411 **
    cboPrograma.Enabled = True
    If cboPrograma.ListIndex <> -1 Then
        cboPrograma.ListIndex = 0
    End If
    sPerSolicitud = ""
    chkITFEfectivo.value = 0 'EJVG20130914
    sMovNroAut = "" 'JUEZ 20131212
    bInstFinanc = False
    'JUEZ 20160420 *********
    fnCampanaCod = 0
    fsCampanaDesc = ""
    ValidaTasaInteres
    'END JUEZ **************
End Sub

Private Sub GetUsuariosReferencia()
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim rsTrab As New ADODB.Recordset
Set clsGen = New COMDConstSistema.DCOMGeneral
Set rsTrab = clsGen.GetTrabajadorCMACT()
If Not (rsTrab.EOF And rsTrab.EOF) Then
    Do While Not rsTrab.EOF
        cboUsuRef.AddItem PstaNombre(rsTrab("cPersNombre")) & Space(100) & rsTrab("cPersCod")
        rsTrab.MoveNext
    Loop
End If
rsTrab.Close
Set rsTrab = Nothing
cboUsuRef.ListIndex = 0
Set clsGen = Nothing
CambiaTamañoCombo cboUsuRef, 280
End Sub

Private Function CuentaTitular() As Integer
Dim i As Integer, nFila As Integer, nCol As Integer

Dim nPers As COMDConstantes.PersPersoneria

nFila = grdCliente.row
nCol = grdCliente.Col
nTitular = 0
nClientes = 0
nRepresentante = 0
nPersoneria = gPersonaNat

For i = 1 To grdCliente.Rows - 1
    If grdCliente.TextMatrix(i, 1) = "" Then Exit For
    If grdCliente.TextMatrix(i, 3) <> "" Then
        If nPersoneria = gPersonaNat Then
            If CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4))) = gCapRelPersTitular Then 'Or CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4))) = gCapRelPersRepTitular Then
                nTitular = nTitular + 1
                nPers = CLng(grdCliente.TextMatrix(i, 4))
                If nPers > nPersoneria Then
                    nPersoneria = nPers
                End If
            ElseIf CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4))) = gCapRelPersRepTitular Then
                nRepresentante = nRepresentante + 1
            End If
        Else
            If CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4))) = gCapRelPersTitular Then
                nTitular = nTitular + 1
                nPers = CLng(grdCliente.TextMatrix(i, 4))
                If nPers > nPersoneria Then
                    nPersoneria = nPers
                End If
            ElseIf CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4))) = gCapRelPersRepTitular Then
                nRepresentante = nRepresentante + 1
            End If
        End If
    End If
    If grdCliente.TextMatrix(i, 1) <> "" Then
        nClientes = nClientes + 1
    End If
Next i
grdCliente.row = nFila
grdCliente.Col = nCol
If nClientes = 0 Then
    cmdEliminar.Enabled = False
    cmdAgregar.Enabled = True
Else
    If nClientes = 1 And nProducto = gCapCTS Then
        cmdAgregar.Enabled = False
        ValidaTasaInteres 'Evalua la tasa según el cliente ingresado (Si tiene o no tiene Cta Sueldo)
        
'        Dim clsCap As NCapMantenimiento
'        Set clsCap = New NCapMantenimiento
'        Dim rsCta As Recordset, x As Integer
'        Dim scadena As String
'         x = 0
'        Set rsCta = clsCap.GetListadoCuentasCTS(Trim(grdCliente.TextMatrix(1, 1)))
'       If Not (rsCta Is Nothing) Then
'
'         scadena = ""
'                Do While Not (rsCta.EOF)
'                        scadena = scadena & rsCta("cctacod") & " " & Trim(Left(rsCta("Cinstitucion"), 40)) & Space(41 - Len(Trim(Left(rsCta("Cinstitucion"), 40)))) & rsCta("cEstado") & vbCrLf
'
'                        If (x = 4 Or x = rsCta.RecordCount - 1) Then
'                            scadena = scadena & " ... APERTURAR CUENTA ????? "
'                            Exit Do
'                        End If
'                        x = x + 1
'                        rsCta.MoveNext
'                Loop
'
'                Font.Bold = True
'                    If MsgBox("CLIENTE POSEE CUENTAS DE CTS ACTUALMENTE:" & vbCrLf & "CUENTA                   INSTITUCION                             ESTADO" & vbCrLf & scadena, vbYesNo + vbDefaultButton2 + vbQuestion, "AVISO") = vbNo Then
'                        cmdCancelar_Click
'                    End If
'
'                Font.Bold = False
'
'        End If
'
'        Set clsCap = Nothing
'        Set rsCta = Nothing
        
    End If
    cmdEliminar.Enabled = True
End If

End Function

Private Sub EvaluaTitular()
Dim i As Integer
'If nClientes > 1 Then
'    For I = 0 To cboTipoCuenta.ListCount - 1
'        If CLng(Trim(Right(cboTipoCuenta.List(I), 4))) = gPrdCtaTpoIndist Or CLng(Trim(Right(cboTipoCuenta.List(I), 4))) = gPrdCtaTpoMancom Then
'            cboTipoCuenta.ListIndex = I
'            Exit For
'        End If
'    Next I
'Else
'    For I = 0 To cboTipoCuenta.ListCount - 1
'        If CLng(Trim(Right(cboTipoCuenta.List(I), 4))) = gPrdCtaTpoIndiv Then
'            cboTipoCuenta.ListIndex = I
'            Exit For
'        End If
'    Next I
'End If


If nPersoneria = gPersonaNat Then
    txtNumFirmas = Format$(nTitular, "#0")
Else
    txtNumFirmas = Format$(nRepresentante, "#0")
End If
End Sub

Private Sub IniciaComboCTSPeriodo()
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsConst As New ADODB.Recordset
Dim sCodigo As String * 2
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsConst = clsMant.GetCTSPeriodo()
Set clsMant = Nothing
Do While Not rsConst.EOF
    sCodigo = rsConst("nItem")
    cboPeriodo.AddItem sCodigo & Space(2) & UCase(rsConst("cDescripcion")) & Space(100) & rsConst("nPorcentaje")
    rsConst.MoveNext
Loop
cboPeriodo.ListIndex = 4
End Sub

Private Sub IniciaCombo(ByRef cboConst As ComboBox, ByVal nCapConst As ConstanteCabecera)
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim rsConst As New ADODB.Recordset
Set clsGen = New COMDConstSistema.DCOMGeneral
' RIRO20131102 comentado
'Set rsConst = clsGen.GetConstante(nCapConst, , , "1") 'MADM 20110630

' RIRO20131102 Agregado por RIRO
If nCapConst = gCaptacSubProdAhorros Then
    ' Filtrando Ahorro destino
    Set rsConst = clsGen.GetConstante(nCapConst, "4", , "1")
Else
    Set rsConst = clsGen.GetConstante(nCapConst, , , "1")
End If

Set clsGen = Nothing
Do While Not rsConst.EOF
    'AVMM -- VALIDAR TIPO DE CUENTA PARA CTS -- 16-06-2006
    
    If nCapConst = gProductoCuentaTipo Then
        If txtCuenta.Prod = "234" Then
            If rsConst("nConsValor") = 0 Then
                cboConst.AddItem rsConst("cDescripcion") & Space(100) & rsConst("nConsValor")
            End If
            rsConst.MoveNext
        Else
           'If Trim(Right(cboPrograma, 1)) = 1 Then
           '***Modificado por ELRO el 20120123, según Acta N° 005-2012/TI-D
           'If Trim(Right(cboPrograma, 1)) = 1 Or Trim(Right(cboPrograma, 1)) = 6 Then
           If txtCuenta.Prod = "232" And (Trim(Right(cboPrograma, 1)) = 1 Or Trim(Right(cboPrograma, 1)) = 6) Then
           '***Fin Modificado por ELRO*************************************
                'ALPA 20091123********************************
                If rsConst("nConsValor") = 0 Then
                '*********************************************
                   cboConst.AddItem rsConst("cDescripcion") & Space(100) & rsConst("nConsValor")
                   
                End If
                rsConst.MoveNext
           Else
                cboConst.AddItem rsConst("cDescripcion") & Space(100) & rsConst("nConsValor")
                rsConst.MoveNext
           End If
        End If
    Else
        cboConst.AddItem rsConst("cDescripcion") & Space(100) & rsConst("nConsValor")
        rsConst.MoveNext
    End If
Loop
cboConst.ListIndex = 0
End Sub

Public Sub IniciaDesembAbonoCta(ByVal nProd As Producto, ByVal nOpe As CaptacOperacion, ByVal sPersona As String, _
    ByVal nMonedaCtaCred As Moneda, ByVal sPersCodTitular As String, sPersNombre As String, ByVal nMontoDesemb As Double, _
    ByRef pMatRela As ADODB.Recordset, ByRef pnTasa As Double, ByRef pnPersoneria As Integer, _
    ByRef pnTipoCuenta As Integer, ByRef pnTipoTasa As Integer, ByRef pbDocumento As Boolean, ByRef psNroDoc As String, _
    ByRef psCodIF As String, Optional ByVal pbTransf = False, _
    Optional ByVal sPersCodRep As String = "", Optional ByVal sPersNombreRep As String = "", _
    Optional ByRef pMatTitulares As Variant, Optional ByRef pnPrograma As Integer = 0, _
    Optional ByRef pnMontoAbonar As Double = 0, Optional ByVal pnPlazoAbono As Integer = 0, _
    Optional ByRef psPromotor As String)

        cmdgrabar.Enabled = False
        cmdCancelar.Enabled = False
        cmdAgregar.Enabled = False
        grdCliente.Enabled = False
        FraCliente.Enabled = False
        txtMonto.Enabled = False
        'Me.BorderStyle = 3
        
        Label5.Visible = False
        Label6.Visible = False
        txtPlazo.Visible = False
        cboFormaRetiro.Visible = False
        lblInst.Visible = False
        txtInstitucion.Visible = False
        lblInstitucion.Visible = False
        'chkOrdenPago.Visible = True 'JUEZ 20141008
        txtCuenta.Prod = Trim(Str(gCapAhorros))
        Me.Caption = "Captaciones - Apertura - Ahorros"
        lblPeriodo.Visible = False
        cboPeriodo.Visible = False
        lblCTS.Visible = False
        lblDispCTS.Visible = False
        txtCtaAhoAboInt.Visible = False
        grdCliente.ColWidth(5) = 0
        chkEmpCMACT.Visible = False
        fraTranferecia.Visible = pbTransf
        
        nProducto = nProd 'FRHU 20140421
        nOperacion = nOpe 'FRHU 20140421
        
'Verifica si la operacion necesita algun documento
If nOperacion = gAhoApeChq Or nOperacion = gPFApeChq Or nOperacion = gCTSApeChq Then
    lblNroDoc.Visible = True
    lblNombreIF.Visible = True
    cmdDocumento.Visible = True
    fraDocumento.Caption = "Cheque"
    bDocumento = True
    txtMonto.Enabled = False
    Label12.Visible = True
    Label13.Visible = True
Else
    lblNroDoc.Visible = False
    lblNombreIF.Visible = False
    cmdDocumento.Visible = False
    bDocumento = False
    Label12.Visible = False
    Label13.Visible = False
End If
cboPrograma.Visible = True
'JUEZ 20141008 *******************************
If nProducto = gCapAhorros Then
    IniciaCombo cboPrograma, 2030
ElseIf nProducto = gCapPlazoFijo Then
    IniciaCombo cboPrograma, 2032
End If
cboPrograma.Enabled = True
cboPrograma.ListIndex = IndiceListaCombo(cboPrograma, pnPrograma)
'END JUEZ ************************************
IniciaCombo cboMoneda, gMoneda
'JUEZ 20141008 *******************************
If nMonedaCtaCred <> CInt(Right(cboMoneda.Text, 1)) Then
    Set pMatRela = Nothing
    Exit Sub
End If
'END JUEZ ************************************
IniciaCombo cboTipoTasa, gCaptacTipoTasa
IniciaCombo cboTipoCuenta, gProductoCuentaTipo
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Age = Right(gsCodAge, 2)
txtCuenta.Enabled = False
cmdAgregar.Enabled = True
cmdEliminar.Enabled = False
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim rsRel As New ADODB.Recordset
Set clsGen = New COMDConstSistema.DCOMGeneral
Set rsRel = clsGen.GetConstante(gCaptacRelacPersona)
Set clsGen = Nothing
grdCliente.CargaCombo rsRel
Set rsRel = Nothing
nPersoneria = gPersonaNat
nTitular = 0
nClientes = 0
nTasaNominal = 0
nTasaNominalTemp = 0
nTasaEfectivaTemp = 0
'Para Desembolso Abono a Cuenta Nueva

    vbDesembolso = True
    chkOrdenPago.Enabled = False
    cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, nMonedaCtaCred)
    cboMoneda.Enabled = False
    txtCuenta.Enabled = False
    cboTipoCuenta.ListIndex = IndiceListaCombo(cboTipoCuenta, gPrdCtaTpoIndiv)
    cboTipoCuenta.Enabled = False
    txtNumFirmas.Text = "1"
    txtNumFirmas.Enabled = False
    grdCliente.AdicionaFila
    grdCliente.TextMatrix(1, 1) = sPersCodTitular
    grdCliente.TextMatrix(1, 2) = sPersNombre
    grdCliente.TextMatrix(1, 3) = "TITULAR" & Space(50) & gCapRelPersTitular
    grdCliente.TextMatrix(1, 9) = "A" 'FRHU 20140228 RQ14006
    
    'ARCV 13-02-2007
    Dim ClsPersona As COMDPersona.DCOMPersonas
    Dim R As New ADODB.Recordset
    Set ClsPersona = New COMDPersona.DCOMPersonas
    Set R = ClsPersona.BuscaCliente(sPersCodTitular, BusquedaCodigo)
        If Not (R.EOF And R.BOF) Then
           grdCliente.TextMatrix(grdCliente.row, 4) = R!nPersPersoneria
           grdCliente.TextMatrix(grdCliente.row, 7) = IIf(R!cPersIDnroDNI = "", R!cPersIDnroRUC, R!cPersIDnroDNI)
           grdCliente.TextMatrix(grdCliente.row, 8) = R!cPersDireccDomicilio
        End If
    Set ClsPersona = Nothing
    '-----------
    
    'CUSCO: AGREGAR LOS REPRESENTANTES
    If sPersCodRep <> "" Then
        grdCliente.AdicionaFila
        grdCliente.TextMatrix(2, 1) = sPersCodRep
        grdCliente.TextMatrix(2, 2) = sPersNombreRep
        grdCliente.TextMatrix(2, 3) = "REP. LEGAL TITULAR" & Space(50) & gCapRelPersRepTitular
    End If
    '******************
    nTitular = 1
    txtMonto.Text = Format(nMontoDesemb, "#0.00")
    txtMonto.value = Format(nMontoDesemb, "#0.00")
    
    cboTipoTasa.ListIndex = IndiceListaCombo(cboTipoTasa, gCapTasaNormal)
    
    'CAAU
    Dim oCons As COMDConstantes.DCOMConstantes
    Set oCons = New COMDConstantes.DCOMConstantes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Label20.Visible = True
'    cboPrograma.Visible = True
'    If nProducto = gCapAhorros Then
'        IniciaCombo cboPrograma, 2030
'    ElseIf nProducto = gCapPlazoFijo Then 'BRGO 20111217 Se agregó subproductos de PlazoFijo
'        IniciaCombo cboPrograma, 2032
'    End If
    Set oCons = Nothing
'    cboPrograma.Enabled = True
    
    Call cboTipoTasa_Click
    'cboTipoTasa.Enabled = False
    
    'ARCV 13-02-2007
    lblCuentaAbo.Visible = False
    chkEspecial.Visible = False
    '---
    Call cmdGrabar_Click
    
    'Me.Show 1
    Call cmdsalir_Click   'FRHU 20140228 RQ14006
    
    Set pMatRela = vMatRela
    pnTasa = nTasaNominal
    pnPersoneria = vnPersoneria
    pnTipoCuenta = vnTipoCuenta
    pnTipoTasa = vnTipoTasa
    pbDocumento = vbDocumento
    psNroDoc = vsNroDoc
    psCodIF = vsCodIF
    
    'ARCV 13-02-2007
    pMatTitulares = vMatTitular
    pnPrograma = vnPrograma
    pnMontoAbonar = vnMontoAbonar
    pnPlazoAbono = vnPlazoAbono
    psPromotor = vsPromotor
    '------------
End Sub

Public Sub Inicia(ByVal nProd As Producto, ByVal nOpe As CaptacOperacion, Optional sPersona As String = "", _
        Optional sDescOperacion As String = "")
    nProducto = nProd
    nOperacion = nOpe

    fgITFParamAsume gsCodAge, CStr(nProd)
    cPersTasaEspecial = ""
    chkAbonIntCta.Visible = False
    chkDepGar.Visible = False
    Select Case nProd
        Case gCapAhorros
            Label5.Visible = True
            Label14.Visible = True
            txtMontoAbonar.Visible = True
            txtMontoAbonar.value = 0
        
            Label6.Visible = False
            txtPlazo.Visible = True
            txtPlazo.Text = "0"
        
            cboFormaRetiro.Visible = False
            lblInst.Visible = False
            txtInstitucion.Visible = False
            lblInstitucion.Visible = False
        
            'Label18.Visible = True RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            Label18.Visible = False 'RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            Label19.Visible = True
            txtAlias.Visible = True
            'TxtMinFirmas.Visible = True RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            TxtMinFirmas.Visible = False 'RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            'grdCliente.ColWidth(6) = 1200 'RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            grdCliente.ColWidth(6) = 0 'RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            chkOrdenPago.Visible = True
            txtCuenta.Prod = Trim(Str(gCapAhorros))
            Me.Caption = "Captaciones - Ahorros " & sDescOperacion
            lblPeriodo.Visible = False
            cboPeriodo.Visible = False
            lblCTS.Visible = False
            lblDispCTS.Visible = False
            txtCtaAhoAboInt.Visible = False
            grdCliente.ColWidth(5) = 0
            chkEmpCMACT.Visible = False
            lblCuentaAbo.Visible = False
            Me.fraITF.Visible = True
            Me.chkExoITF.value = 0
            chkExoITF_Click
        
            If gbITFAsumidoAho Then
                chkITFEfectivo.value = 1
                chkITFEfectivo.Visible = False
            Else
                chkITFEfectivo.value = 0
                chkITFEfectivo.Visible = True
            End If
        
            Dim oCons As COMDConstantes.DCOMConstantes
            Set oCons = New COMDConstantes.DCOMConstantes
            Dim rs As ADODB.Recordset
            Set rs = New ADODB.Recordset
            Label20.Visible = True
            cboPrograma.Visible = True
            IniciaCombo cboPrograma, 2030
            Set oCons = Nothing
            chkEspecial.Visible = False
            FraITFAsume.Visible = False
            
            'Add By Gitu 22-10-2012
            chkRelConv.Visible = True
            cboPrograma_Click ' RIRO20131102
            IniciaComboConvDep 9
            fraPromotor.Visible = IIf(nOperacion = gAhoApeCargoCta, False, True) 'JUEZ 20131212
            FraCargoCta.Visible = IIf(nOperacion = gAhoApeCargoCta, True, False) 'JUEZ 20131212
            
    Case gCapPlazoFijo
            Label5.Visible = True
            Label6.Visible = True
            txtPlazo.Visible = True
            cboFormaRetiro.Visible = True
            lblInst.Visible = False
        
            Label14.Visible = False
            txtMontoAbonar.Visible = False
        
            'Label18.Visible = True RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            Label18.Visible = False 'RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            Label19.Visible = True
            txtAlias.Visible = True
            'TxtMinFirmas.Visible = True RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            TxtMinFirmas.Visible = False 'RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
        
            'grdCliente.ColWidth(6) = 1200 RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            grdCliente.ColWidth(6) = 0 'RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            chkOrdenPago.Visible = False
            txtInstitucion.Visible = False
            lblInstitucion.Visible = False
            txtCuenta.Prod = Trim(Str(gCapPlazoFijo))
            lblPeriodo.Visible = False
            cboPeriodo.Visible = False
            lblCTS.Visible = False
            lblDispCTS.Visible = False
            txtCtaAhoAboInt.Visible = False
            Me.Caption = "Captaciones - Plazo Fijo - " & sDescOperacion
            IniciaCombo cboFormaRetiro, gCaptacPFFormaRetiro
            chkEmpCMACT.Visible = False
            txtPlazo.Text = "0"
            lblCuentaAbo.Visible = False
            Me.fraITF.Visible = True
            Me.chkExoITF.value = 0
            chkExoITF_Click
        
            If gbITFAsumidoPF Then
                chkITFEfectivo.value = 0
            Else
                chkITFEfectivo.value = 1
            End If
            chkITFEfectivo.Visible = 1
            '***Modificado por ELRO el 20110912, según Acta 245-2011/TI-D
            'chkITFEfectivo.Enabled = False 'comentado por ELRO el 20110912
            chkITFEfectivo.Enabled = True
            '***End Modificado por ELRO**********************
            chkEspecial.Visible = False
            '***Modificado por ELRO el 20110912, según Acta 245-2011/TI-D
            'FraITFAsume.Visible = True 'comentado por ELRO el 20110912
            FraITFAsume.Visible = False
            '***End Modificado por ELRO**********************
            OptAsuITF(0).value = False
            OptAsuITF(1).value = False
            chkAbonIntCta.Visible = True
            chkDepGar.Visible = True 'MADM 20111022
            grdCliente.ColWidth(5) = 0
            lnValOpePF = 1  'Add GITU 20100806 para realizar validacion a las aperturas de PF
            cboPrograma.Visible = True 'BRGO 20111217
            IniciaCombo cboPrograma, 2032 'BRGO 20111217
            Label20.Visible = True
            fraPromotor.Visible = IIf(nOperacion = gPFApeCargoCta, False, True) 'JUEZ 20131212
            FraCargoCta.Visible = IIf(nOperacion = gPFApeCargoCta, True, False) 'JUEZ 20131212
        Case gCapCTS
            Label5.Visible = False
            Label6.Visible = False
            txtPlazo.Visible = False
            Label14.Visible = False
            txtMontoAbonar.Visible = False
            cboFormaRetiro.Visible = False
            chkOrdenPago.Visible = False
            txtNumFirmas.Enabled = False
            txtCuenta.Prod = Trim(Str(gCapCTS))
            lblPeriodo.Visible = True
            cboPeriodo.Visible = True
            lblCTS.Visible = True
            lblDispCTS.Visible = True
            lblInst.Visible = True
            txtInstitucion.Visible = True
            lblInstitucion.Visible = True
        
            Label18.Visible = False
            Label19.Visible = False
            txtAlias.Visible = False
            TxtMinFirmas.Visible = False
            chkEspecial.Visible = False 'True 'JUEZ 20130814
        
            grdCliente.ColWidth(6) = 0
        
            IniciaComboCTSPeriodo
            Me.Caption = "Captaciones - CTS - " & sDescOperacion
            txtCtaAhoAboInt.Visible = False
            grdCliente.ColWidth(5) = 0
            chkEmpCMACT.Visible = True
            lblCuentaAbo.Visible = False
            Me.fraITF.Visible = False
            Me.chkExoITF.value = 0
            chkITFEfectivo.Visible = False
            chkEspecial.Visible = False 'True 'JUEZ 20130814
            FraITFAsume.Visible = False
            FraCargoCta.Visible = False 'JUEZ 20131212
    End Select
    Me.chkPromotor.value = 0
    Me.cboPromotor.Enabled = False
    'Verifica si la operacion necesita algun documento
    If nOperacion = gAhoApeChq Or nOperacion = gPFApeChq Or nOperacion = gCTSApeChq Then
        Me.fraDocumento.Visible = True
        Me.fraTranferecia.Visible = False
        nDocumento = TpoDocCheque
        lblNroDoc.Visible = True
        lblNombreIF.Visible = True
        cmdDocumento.Visible = True
        fraDocumento.Caption = "Cheque"
        bDocumento = True
        txtMonto.Enabled = False
        Label12.Visible = True
        Label13.Visible = True
        chkITFEfectivo.Enabled = False
        chkITFEfectivo.value = vbChecked
        'EJVG20140408 ***
        cmdAgregar.Visible = False
        cmdEliminar.Visible = False
        chkITFEfectivo.Enabled = True
        'END EJVG *******
    ElseIf nOperacion = gAhoApeTransf Or nOperacion = gPFApeTransf Or nOperacion = gCTSApeTransf Then
        lblNroDoc.Visible = False
        lblNombreIF.Visible = False
        cmdDocumento.Visible = False
        bDocumento = False
        Label12.Visible = False
        Label13.Visible = False
        Me.fraDocumento.Visible = False
        Me.fraTranferecia.Visible = True
        
        '***Modificado por ELRO el 20120725, según OYP-RFC024-2012
        'chkITFEfectivo.Visible = False
        'chkITFEfectivo.value = 1
        chkITFEfectivo.value = 0
        chkITFEfectivo.Visible = True
        'chkITFEfectivo.Enabled = False
        chkITFEfectivo.Enabled = True 'EJVG20130913
        '***Fin Modificado por ELRO el 20120725*******************
        '***Agregado por ELRO el 20120706, según OYP-RFC024-2012
        txtMonto.value = 0
        txtMonto.Enabled = False
        LblItf.Caption = "0.00"
        lblTotal.Caption = "0.00"
        lblEtiMonTra.Visible = True
        lblSimTra.Visible = True
        lblMonTra.Visible = True
        chkExoITF.value = 0
        chkExoITF_Click
        cboTransferMoneda.Enabled = False
        If nOperacion <> gCTSApeTransf Then
            fraITF.Visible = True
        End If
        chkEspecial.Visible = False
        '***Fin Agregado por ELRO*******************************
        'EJVG20130912 ***
        cmdAgregar.Visible = False
        cmdEliminar.Visible = False
        'END EJVG *******
    Else
        Me.fraDocumento.Visible = True
        Me.fraTranferecia.Visible = False
        lblNroDoc.Visible = False
        lblNombreIF.Visible = False
        cmdDocumento.Visible = False
        bDocumento = False
        Label12.Visible = False
        Label13.Visible = False
    End If
    nTasaNominal = 0
    vbDesembolso = False
    nPersoneria = gPersonaNat
    IniciaCombo cboMoneda, gMoneda
    IniciaCombo cboTransferMoneda, gMoneda

    IniciaCombo cboTipoTasa, gCaptacTipoTasa
    cboTipoCuenta.Clear
    IniciaCombo cboTipoCuenta, gProductoCuentaTipo
    txtCuenta.CMAC = gsCodCMAC
    txtCuenta.Age = Right(gsCodAge, 2)
    txtCuenta.Enabled = False
    cmdAgregar.Enabled = True
    cmdEliminar.Enabled = False
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsRel As New ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsRel = clsGen.GetConstante(gCaptacRelacPersona, "13,14")
    Set clsGen = Nothing
    grdCliente.CargaCombo rsRel
    Set rsRel = Nothing
    nTitular = 0
    nClientes = 0
    'GetUsuariosReferencia

    Dim oGen As COMDConstSistema.DCOMGeneral
    Set oGen = New COMDConstSistema.DCOMGeneral

    lbImpRegFirma = CInt(oGen.LeeConstSistema(100))
    Set oGen = Nothing
    'JUEZ 20131212 ********************
    txtCuentaCargo.CMAC = gsCodCMAC
    txtCuentaCargo.Age = gsCodAge
    txtCuentaCargo.Prod = gCapAhorros
    sMovNroAut = ""
    'END JUEZ *************************
    bInstFinanc = False 'JUEZ 20140414
    'JUEZ 20160420 ************
    fnCampanaCod = 0
    fsCampanaDesc = ""
    'END JUEZ *****************
    Me.Show 1
End Sub
'MIOL 20121011, según OYP-RFC098-2012 *****************
Private Sub cboFormaRetiro_Click()
    If CLng(Trim(Right(cboFormaRetiro.Text, 4))) = 2 Then
        chkSubasta.Visible = True
    Else
        chkSubasta.Visible = False
    End If
End Sub
'END MIOL *********************************************

Private Sub cboFormaRetiro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPlazo.SetFocus
End If
End Sub

Private Sub cboMoneda_Click()
nmoneda = CLng(Right(cboMoneda.Text, 1))
'JUEZ 20141008 VERIFICAR PARAMETRO MONEDA *****************
If nProducto <> gCapCTS Then
    If nmoneda = gMonedaNacional And Not bParMonedaSol Then
        '''MsgBox "El producto no permite apertura de cuentas en soles", vbInformation, "Aviso" 'MARG ERS 044-2016
        MsgBox "El producto no permite apertura de cuentas en " & StrConv(gcPEN_PLURAL, vbLowerCase), vbInformation, "Aviso" 'MARG ERS 044-2016
        cboMoneda.ListIndex = 1
    End If
    If nmoneda = gMonedaExtranjera And Not bParMonedaDol Then
        MsgBox "El producto no permite apertura de cuentas en dólares", vbInformation, "Aviso"
        cboMoneda.ListIndex = 0
    End If
End If
'END JUEZ *************************************************
If nmoneda = gMonedaNacional Then
    txtMonto.BackColor = &HC0FFFF
    '''lblMon.Caption = "S/." 'MARG ERS044-2016
    lblMon.Caption = gcPEN_SIMBOLO 'MARG ERS044-2016
    '''lblCMon.Caption = "S/." 'MARG ERS044-2016
    lblCMon.Caption = gcPEN_SIMBOLO 'MARG ERS044-2016
    txtInta.BackColor = &HC0FFFF
    txtDisp.BackColor = &HC0FFFF
    txtDU.BackColor = &HC0FFFF
    lblTotTran.BackColor = &HC0FFFF
    
ElseIf nmoneda = gMonedaExtranjera Then
    txtMonto.BackColor = &HC0FFC0
    lblMon.Caption = "$"
    
    lblCMon.Caption = "$"
    txtInta.BackColor = &HC0FFC0
    txtDisp.BackColor = &HC0FFC0
    txtDU.BackColor = &HC0FFC0
    lblTotTran.BackColor = &HC0FFC0

ElseIf nmoneda = 3 Then
    txtMonto.BackColor = &HC0C0FF
    lblMon.Caption = "Eu."
    
    lblCMon.Caption = "Eu."
    txtInta.BackColor = &HC0C0FF
    txtDisp.BackColor = &HC0C0FF
    txtDU.BackColor = &HC0C0FF
    lblTotTran.BackColor = &HC0C0FF
    
End If

Me.LblItf.BackColor = txtMonto.BackColor
Me.lblTotal.BackColor = txtMonto.BackColor


If nOperacion = gAhoApeChq Or nOperacion = gPFApeChq Or nOperacion = gCTSApeChq Then
    txtMonto.value = 0
    lblNroDoc.Caption = ""
    Me.lblNombreIF.Caption = ""
    SetDatosCheque 'EJVG20140408
End If

'Función que valida la tasa de interes
ValidaTasaInteres

'Verifica el tipo de cuenta para el abono de intereses de plazo fijo
'** BRGO 20110701 *****************************
If nProducto = gCapPlazoFijo Then
    Dim i As Long
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsCta As New ADODB.Recordset
    
    '*** BRGO 20111220 ***********************************
    If nmoneda = gMonedaExtranjera And cboPrograma.ListIndex = 1 Then 'Condiciona que PF Premium aplica sólo en MN
        cboMoneda.ListIndex = 0
        Set clsMant = Nothing
        Set rsCta = Nothing
        Exit Sub
    End If
    '*** END BRGO ****************************************
    
    If Me.chkAbonIntCta.value = vbChecked Then
        Set rsCta = clsMant.GetCuentaAhorroTitularesPF(nTipoCuenta, ObtTodosTitulares, nmoneda)
        txtCtaAhoAboInt.rs = rsCta
        Set rsCta = Nothing
    End If
'*** END BRGO
'Comentado por BRGO 20110701
'    For I = 1 To grdCliente.Rows - 1
'        If grdCliente.TextMatrix(I, 5) = "." Then
'            Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
'            Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
'            Set rsCta = clsMant.GetCuentasPersona(grdCliente.TextMatrix(I, 1), gCapAhorros, True, , nmoneda)
'            Set clsMant = Nothing
'            txtCtaAhoAboInt.rs = rsCta
'            Set rsCta = Nothing
'            Exit For
'        End If
'    Next I
    
End If

If nOperacion = gAhoApeTransf Or nOperacion = gPFApeTransf Or nOperacion = gCTSApeTransf Then
    SetDatosTransferencia "", "", "", 0, -1, "" 'EJVG20130912
    '***Agregado por ELRO el 20121015, según OYP-RFC024-2012
    If cboTransferMoneda.Visible Then
        cboTransferMoneda.ListIndex = cboMoneda.ListIndex
    End If
    '***Fin Agregado por ELRO el 20121015********************
    'If Right(cboMoneda, 3) = Moneda.gMonedaNacional Then
    '    If Right(cboTransferMoneda, 3) = Moneda.gMonedaNacional Then
    '        Me.txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")
    '    Else
    '        Me.txtMonto.Text = Format(lnTransferSaldo * CCur(Me.lblTTCCD.Caption), "#,##0.00")
    '    End If
    'Else
    '    If Right(cboTransferMoneda, 3) = Moneda.gMonedaNacional Then
    '        Me.txtMonto.Text = Format(lnTransferSaldo / CCur(Me.lblTTCVD.Caption), "#,##0.00")
    '    Else
    '        Me.txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")
    '    End If
    'End If
End If

'JUEZ 20131212 ********************************************************
If nOperacion = gAhoApeCargoCta Or nOperacion = gPFApeCargoCta Then
    If Len(txtCuentaCargo.NroCuenta) = 18 Then
        If Trim(Right(cboMoneda, 2)) <> Mid(txtCuentaCargo.NroCuenta, 9, 1) Then
            MsgBox "La moneda debe ser igual a la moneda de la cuenta a debitar", vbInformation, "Aviso"
            Me.cboMoneda.ListIndex = IIf(Mid(txtCuentaCargo.NroCuenta, 9, 1) = gMonedaNacional, 0, 1)
            Exit Sub
        End If
    End If
End If
'END JUEZ *************************************************************

End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cboTipoTasa.Enabled Then cboTipoTasa.SetFocus
End If
End Sub

Private Sub cboPeriodo_Click()
lblDispCTS = Format$(CDbl(Trim(Right(cboPeriodo.Text, 5))) * 100, "#,##0.00")
End Sub

Private Sub cboPeriodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    chkEmpCMACT.SetFocus
End If
End Sub

Private Sub cboPrograma_Click()
Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim bOrdPag As Boolean
Dim nMonto As Double
Dim nPlazo As Long
Dim nTpoPrograma As Integer

'JUEZ 20141008 ******************************************************
Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
If nProducto <> gCapCTS Then
    Dim rsPar As ADODB.Recordset
    Set rsPar = clsDef.GetCapParametroNew(nProducto, CInt(Trim(Right(cboPrograma.Text, 2))))
    bParPersNat = rsPar!bPersNat
    bParPersJur = rsPar!bPersJur
    bParMonedaSol = rsPar!bMonSol
    bParMonedaDol = rsPar!bMonDol
    nParMontoMinSol = rsPar!nMontoMinApertSol
    nParMontoMinDol = rsPar!nMontoMinApertDol
    If nProducto = gCapAhorros Then
        nParOrdPag = rsPar!nOrdPago
    ElseIf nProducto = gCapPlazoFijo Then
        nParPlazoMin = rsPar!nPlazoMin
        nParPlazoMax = rsPar!nPlazoMax
        bParFormaRetFinPlazo = rsPar!bFormaRetFinPlazo
        bParFormaRetMensual = rsPar!bFormaRetMensual
        bParFormaRetIniPlazo = rsPar!bFormaRetInicioPlazo
        nParAumCapMinSol = rsPar!nAumCapMinSol
        nParAumCapMinDol = rsPar!nAumCapMinDol
    End If
    If (Trim(Right(cboMoneda.Text, 2))) <> "" Then cboMoneda_Click
End If
'END JUEZ ***********************************************************
'JUEZ 20131212 *************************************************
If nProducto = gCapAhorros And nOperacion = gAhoApeCargoCta Or nOperacion = gPFApeCargoCta Then
    If Val(Right(cboPrograma.Text, 2)) = 7 Or Val(Right(cboPrograma.Text, 2)) = 6 Then
        MsgBox "La cuenta a Aperturar no puede ser Ecotaxi ni Caja Sueldo", vbInformation, "Aviso"
        cboPrograma.ListIndex = 0
        Exit Sub
    End If
End If
'END JUEZ ******************************************************

Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
bOrdPag = IIf(chkOrdenPago.value = 1, True, False)
nMonto = txtMonto.value
nTpoPrograma = 1

txtCtaAhoAboInt.Visible = False
If nProducto = gCapAhorros Then
Me.chkRelConv.Visible = False ' RIRO20131102
    If Trim(Right(cboPrograma.Text, 1)) = 0 Then
        Me.chkOrdenPago.Visible = True
        Me.Label5.Visible = False
        Me.Label14.Visible = False
        Me.txtPlazo.Text = "0"
        Me.txtMontoAbonar.Text = "0"
        Me.txtPlazo.Visible = False
        Me.txtMontoAbonar.Visible = False
        txtInstitucion.Visible = False
        lblInstitucion.Visible = False
        lblInst.Visible = False
        '***Agregado por ELRO el 20120403
        chkEmpCMACT.Visible = False
        chkEmpCMACT.value = 0
        '***Fin Agregado por ELRO********
        
        '*** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES" ***
            'cboTipoCuenta.Clear
            'IniciaCombo cboTipoCuenta, gProductoCuentaTipo
    ElseIf Trim(Right(cboPrograma.Text, 1)) = 1 Then
        Me.chkOrdenPago.Visible = False
        'Me.Label5.Visible = True RIRO20131102
        Me.Label14.Visible = False
        'Me.txtPlazo.Text = "0" RIRO20131102
        Me.txtMontoAbonar.Text = "0"
        'Me.txtPlazo.Visible = True RIRO20131102
        Me.txtMontoAbonar.Visible = False
        txtInstitucion.Visible = False
        lblInstitucion.Visible = False
        lblInst.Visible = False
        
        '*** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES" ***
            'cboTipoCuenta.Clear
            'IniciaCombo cboTipoCuenta, gProductoCuentaTipo
    ElseIf Trim(Right(cboPrograma.Text, 1)) = 7 Then
        Me.chkOrdenPago.Visible = False
        chkAbonIntCta.Visible = False
        txtCtaAhoAboInt.Visible = False
        cboMoneda.ListIndex = 0
        lblInstitucion.Visible = False
        txtInstitucion.Visible = False
        lblInst.Visible = False
        chkEmpCMACT.Visible = False
        txtPlazo.Visible = False
        Label5.Visible = False
        Me.txtMontoAbonar.Visible = False
        Me.Label14.Visible = False
    Else
        Me.chkOrdenPago.Visible = False
        Me.Label5.Visible = False
        Me.Label14.Visible = False
        Me.txtPlazo.Text = "0"
        Me.txtMontoAbonar.Text = "0"
        Me.txtPlazo.Visible = False
        Me.txtMontoAbonar.Visible = False
        chkEmpCMACT.Visible = False
        txtInstitucion.Visible = False
        lblInstitucion.Visible = False
        lblInst.Visible = False
        If Trim(Right(cboPrograma.Text, 1)) = 4 Then
           txtInstitucion.Visible = True
           lblInstitucion.Visible = True
           lblInst.Visible = True
           Label14.Top = 240
           txtMontoAbonar.Top = 240
           txtPlazo.Top = 240
           Label5.Top = 240
        End If
        
        '*** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES" ***
            'cboTipoCuenta.Clear
            'IniciaCombo cboTipoCuenta, gProductoCuentaTipo
    End If
    If Trim(Right(cboPrograma.Text, 1)) = 6 Then
        chkEmpCMACT.Visible = True
        '***Agregado por ELRO el 20120403
        chkEmpCMACT.value = 0
        '***Fin Agregado por ELRO********
        txtInstitucion.Visible = True
        lblInstitucion.Visible = True
        lblInst.Visible = True
        
        Me.Label5.Visible = False
        Me.Label14.Visible = False
        Me.txtPlazo.Text = "0"
        Me.txtMontoAbonar.Text = "0"
        Me.txtPlazo.Visible = False
        Me.txtMontoAbonar.Visible = False
        chkExoITF.value = 1
        Me.chkITFEfectivo.value = 1
        cboTipoExoneracion.ListIndex = IndiceListaCombo(cboTipoExoneracion, 3)
        Me.chkRelConv.Visible = False
    Else
        chkExoITF.value = 0
        Me.chkITFEfectivo.value = 0
        cboTipoExoneracion.ListIndex = -1
        'Me.chkRelConv.Visible = True RIRO20131102
    End If
    If Trim(Right(cboPrograma.Text, 1)) = 8 Then
        Me.chkRelConv.Visible = True
    End If
'**** BRGO 20111219 ********************************************
ElseIf nProducto = gCapPlazoFijo Then
    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral
    
    Me.chkOrdenPago.Visible = False
    Me.Label5.Visible = True
    Me.txtPlazo.Text = "0"
    Me.txtPlazo.Visible = True
    Me.txtMontoAbonar.Visible = True
    txtInstitucion.Visible = False
    lblInstitucion.Visible = False
    lblInst.Visible = False
    
    '*** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES" ***
        'cboTipoCuenta.Clear
        'IniciaCombo cboTipoCuenta, gProductoCuentaTipo

    chkDepGar.Visible = False
    chkDepGar.value = 0
    chkSubasta.Visible = False 'Agregado por MIOL 20121011, según OYP-RFC098-2012
    chkAbonIntCta.Visible = False
    txtCtaAhoAboInt.Visible = False
    txtMontoAbonar.Visible = False
    Label14.Visible = False
    cboFormaRetiro.Visible = False
    Label6.Visible = False
    chkTasaPreferencial.Visible = True
    txtPlazo.Locked = False
    If cboPrograma.ListIndex = 0 Then
        chkDepGar.Visible = True
        chkAbonIntCta.Visible = True
        'txtCtaAhoAboInt.Visible = True
        cboFormaRetiro.Visible = True
        Label6.Visible = True
    End If
    If cboPrograma.ListIndex = 1 Then
        'txtPlazo.Text = loRs.GetParametro(2000, 2120)
        txtPlazo.Text = nParPlazoMin 'JUEZ 20141008 Nuevos Parámetros
        chkTasaPreferencial.Visible = False
        txtPlazo.Locked = True
        'nMontoMinimoPFPremium = loRs.GetParametro(2000, 2119)
        nMontoMinimoPFPremium = IIf(CInt(Trim(Right(cboMoneda.Text, 1))) = gMonedaNacional, nParMontoMinSol, nParMontoMinDol) 'JUEZ 20141008 Nuevos Parámetros
    End If
    If cboPrograma.ListIndex = 2 Or cboPrograma.ListIndex = 3 Then
        txtMontoAbonar.Visible = True
        Label14.Visible = True
    End If
End If
'*** END BRGO **********************************************************

'JUEZ 20141008 ************************************
If nProducto = gCapAhorros Then
    chkOrdenPago.Visible = IIf(nParOrdPag = 0, False, True)
    chkOrdenPago.value = IIf(nParOrdPag = 1, 1, 0)
    chkOrdenPago.Enabled = IIf(nParOrdPag = 1, False, True)
End If
'END JUEZ *****************************************
'By Capi 19082008 para no visualizar plazo y monto minimo
If cboPrograma.ListIndex = 5 Then
    Me.txtPlazo.Visible = False
    Me.txtMontoAbonar.Visible = False
End If

If cboPrograma.Visible Then
    nTpoPrograma = CInt(Right(Trim(cboPrograma.Text), 2))
End If

If chkTasaPreferencial.value = vbUnchecked Then
    'JUEZ 20160420 ***************************************************
    'If nProducto = gCapPlazoFijo Then
    '    If txtPlazo <> "" Then
    '        nPlazo = CLng(txtPlazo)
    '        If chkDepGar.value = 1 Then 'MADM 20111022
    '            'nTasaNominal = (clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma) / 2)
    '            nTasaNominalTemp = clsDef.GetCapTasaInteres(nProducto, nMoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma)
    '            nTasaEfectivaTemp = Format$(ConvierteTNAaTEA(nTasaNominalTemp), "#,##0.00") / 2
    '            nTasaNominal = Format$(ConvierteTEAaTNA(nTasaEfectivaTemp), "#,##0.00")
    '        Else
    '            nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nMoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma)
    '        End If
    '
    '        lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
    '    End If
    'ElseIf nProducto = gCapAhorros Then
    '    nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nMoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, bOrdPag, nTpoPrograma)
    '    lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
    '
    'Else
    '    nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nMoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma)
    '    lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
    'End If
    ValidaTasaInteres
    'END JUEZ ********************************************************
End If

If nProducto = gCapAhorros Then
    If nTpoPrograma = 0 Or nTpoPrograma = 5 Or nTpoPrograma = 6 Or nTpoPrograma = 8 Then
    '***Se agreggó la condición nTpoPrograma = 8 por ELRO, el 20130130 según TI-ERS020-2013
        chkRelConv.Visible = True
        If chkRelConv.value = 1 And nTpoPrograma <> 6 Then
            Me.cboInstConvDep.Visible = True
        ElseIf nTpoPrograma = 5 Then
            Me.txtInstitucion.Visible = False
            Me.lblInstitucion.Visible = False
            Me.lblInst.Visible = False
        Else
            Me.cboInstConvDep.Visible = False
'            Me.txtInstitucion.Visible = False
'            Me.lblInstitucion.Visible = False
'            Me.lblInst.Visible = False
        End If
        ' RIRO20131102
        If nTpoPrograma = 8 Then
            chkRelConv.Visible = True
        Else
            chkRelConv.Visible = False
        End If
    Else
        chkRelConv.Visible = False
        chkRelConv.value = 0
        Me.cboInstConvDep.Visible = False
    End If
    lnTpoPrograma = nTpoPrograma
End If

Set clsDef = Nothing
End Sub

Private Sub cboTipoCuenta_Click()
    If ValidarFirmas = False Then
        Exit Sub
    Else
        nTipoCuenta = CLng(Trim(Right(cboTipoCuenta.Text, 4)))
    End If
End Sub

Private Sub cboTipoCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not txtNumFirmas.Enabled Then
        txtNumFirmas.Enabled = True
    End If
    txtNumFirmas.SetFocus
End If
End Sub




Private Sub cboTipoTasa_Click()
'Función que valida la tasa de interes
ValidaTasaInteres
End Sub
Private Sub cboTipoTasa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If chkOrdenPago.Visible Then
        chkOrdenPago.SetFocus
    Else
        'cmdAgregar.SetFocus
        If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
    End If
    If cboFormaRetiro.Visible Then
        cboFormaRetiro.SetFocus
    End If
    If cboPeriodo.Visible Then
        cboPeriodo.SetFocus
    End If
End If
End Sub
'***Agregado por ELRO el 20120823, según OYP-RFC024-2012
Private Sub cboTransferMoneda_Click()
    If Right(cboTransferMoneda, 3) = Moneda.gMonedaNacional Then
        '''lblSimTra.Caption = "S/." 'marg ers044-2016
        lblSimTra.Caption = gcPEN_SIMBOLO 'marg ers044-2016
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

Private Sub cboUsuRef_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGlosa.SetFocus
End If
End Sub

Private Sub chkAbonIntCta_Click()
    Dim rsCta As ADODB.Recordset
    Set rsCta = New ADODB.Recordset
    txtCtaAhoAboInt.rs = rsCta
    If chkAbonIntCta.value = vbChecked Then
        txtCtaAhoAboInt.Visible = True
        If nProducto = gCapPlazoFijo Then
            Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
            Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
            Set rsCta = clsMant.GetCuentaAhorroTitularesPF(nTipoCuenta, ObtTodosTitulares, nmoneda)
            Set clsMant = Nothing
            txtCtaAhoAboInt.rs = rsCta
            Set rsCta = Nothing
            lblCuentaAbo.Visible = True
            txtCtaAhoAboInt.Visible = True
        End If
    Else
        txtCtaAhoAboInt.Visible = False
        txtCtaAhoAboInt.Text = ""
    End If
End Sub
'MADM 20111022
Private Sub chkDepGar_Click()
   ValidaTasaInteres
   'MIOL 20121012, SEGUN OYP-RFC098-2012
   If chkDepGar.value = 1 Then
        Me.chkSubasta.Enabled = False
   Else
        Me.chkSubasta.Enabled = True
   End If
   'END IF *****************************
End Sub

Private Sub chkEmpCMACT_Click()
If chkEmpCMACT.value = 1 Then
    If Not ValidaInstConv(txtInstitucion.Text) And chkRelConv.value = 1 Then
        MsgBox "La Institucion no esta para convenio de Depositos", vbInformation, "SISTEMA"
        txtInstitucion.Text = ""
        chkEmpCMACT.value = 0
    Else
        lblInstitucion.Enabled = False
        lblInst.Enabled = False
        txtInstitucion.Text = gsCodPersCMACT
        txtInstitucion.Enabled = False
        lblInstitucion = UCase(gsNomCmac)
    End If
Else
    lblInstitucion = ""
    txtInstitucion.Text = ""
    lblInstitucion.Enabled = True
    lblInst.Enabled = True
    txtInstitucion.Enabled = True
End If
ValidaTasaInteres 'JUEZ 20140319
End Sub

Private Sub chkEmpCMACT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtInstitucion.Enabled And txtInstitucion.Visible Then Me.txtInstitucion.SetFocus
    End If
End Sub

Private Sub chkEspecial_Click()
If chkEspecial.value = vbChecked Then
    fraespecialCTS.Visible = True
    txtDisp.Text = "0.00"
    txtInta.Text = "0.00"
    txtDU.Text = "0.00"
'    lblTotTran.Caption = "0.00"
Else
    fraespecialCTS.Visible = False
    txtDisp.Text = "0.00"
    txtInta.Text = "0.00"
    txtDU.Text = "0.00"
'    lblTotTran.Caption = "0.00"
End If

End Sub

Private Sub chkExoITF_Click()
'***Agregado por ELRO el 20120825, según OYP-RFC024-2012
If (gsOpeCod = CStr(gAhoApeTransf) Or _
   gsOpeCod = CStr(gPFApeTransf) Or _
   gsOpeCod = CStr(gCTSApeTransf)) And _
   chkExoITF.value = 1 Then

    Dim lbResultadoVisto As Boolean
    Dim loVistoElectronico As frmVistoElectronico
    Set loVistoElectronico = New frmVistoElectronico
     
    lbResultadoVisto = loVistoElectronico.Inicio(3, gsOpeCod)
    If Not lbResultadoVisto Then
       chkExoITF.value = 0
       Exit Sub
    End If

End If
'***Fin Agregado por ELRO el 20120825*******************


    Me.cboTipoExoneracion.ListIndex = -1
    If chkExoITF.value = 1 Then
        Me.cboTipoExoneracion.Enabled = True
    Else
        Me.cboTipoExoneracion.Enabled = False
    End If
    txtMonto_Change
End Sub

Private Sub chkOrdenPago_Click()
'Función que valida la tasa de interes
ValidaTasaInteres
End Sub

Private Sub chkOrdenPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'cmdAgregar.SetFocus
    If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
End If
End Sub

Private Sub chkPromotor_Click()
    Me.cboPromotor.ListIndex = -1
    If chkPromotor.value = 1 Then
        Me.cboPromotor.Enabled = True
    Else
        Me.cboPromotor.Enabled = False
    End If
    txtMonto_Change
End Sub

Private Sub chkRelConv_Click()
    If lnTpoPrograma <> 6 Then
        If chkRelConv.value = 1 Then
            cboInstConvDep.Visible = True
        Else
            cboInstConvDep.Visible = False
        End If
    End If
End Sub

Private Sub chkSubasta_Click()
   'MIOL 20121011, SEGUN OYP-RFC098-2012
   If chkSubasta.value = 1 Then
        Me.chkDepGar.Enabled = False
   Else
        Me.chkDepGar.Enabled = True
   End If
   'END IF *****************************
End Sub

'Private Sub chkTasaEspecial_Click()
'Dim i As Integer, J As Integer, rsTemp As ADODB.Recordset
'Dim bEncontro As Boolean
'
''cPersTasaEspecial = ""
''bEncontro = False
''
''If chkTasaEspecial.value = vbChecked Then
''    cboTasaEspecial.Clear
''    lblTasaEspecial.Caption = ""
''    For i = 1 To grdCliente.Rows - 1
''        If Left(grdCliente.TextMatrix(i, 3), 7) = "TITULAR" Then
''            Set rsTemp = New ADODB.Recordset
''            Set rsTemp = GetDataInteresEspecial(grdCliente.TextMatrix(i, 1), 1, nProducto, nMoneda, gdFecSis)
''            If rsTemp.State = 1 Then
''                If rsTemp.RecordCount >= 1 Then
''                    bEncontro = True
''                    cPersTasaEspecial = grdCliente.TextMatrix(i, 1)
''                    Exit For
''                End If
''            End If
''        End If
''    Next i
''
''    If bEncontro Then
''        Label20.Visible = True
''        If rsTemp.RecordCount > 1 Then
''            cboTasaEspecial.Visible = True
''            j = 0
''            While Not rsTemp.EOF
''                cboTasaEspecial.AddItem Format$(ConvierteTNAaTEA(rsTemp!nTasa), "#0.00") & Space(50) & ":" & Trim(rsTemp!nMonto) & "/" & CStr(rsTemp!nPlazo)
''                cboTasaEspecial.ItemData(j) = rsTemp!nNumSolicitud
''                j = j + 1
''                rsTemp.MoveNext
''            Wend
''            cboTasaEspecial.ListIndex = 0
''            txtMonto.Text = CCur(Mid(cboTasaEspecial.Text, InStr(1, cboTasaEspecial.Text, ":", vbTextCompare) + 1, InStr(1, cboTasaEspecial.Text, "/", vbTextCompare) - InStr(1, cboTasaEspecial.Text, ":", vbTextCompare) - 1))
''            txtPlazo.Text = CCur(Mid(cboTasaEspecial.Text, InStr(1, cboTasaEspecial.Text, "/", vbTextCompare) + 1, Len(cboTasaEspecial.Text) - InStr(1, cboTasaEspecial.Text, "/", vbTextCompare)))
''        Else
''             lblTasaEspecial.Visible = True
''             nTasaNominal = rsTemp!nTasa
''             lblTasaEspecial.Caption = Format$(ConvierteTNAaTEA(rsTemp!nTasa), "#0.00")
''             lblTasaEspecial.Tag = CStr(rsTemp!nNumSolicitud)
''             txtMonto.Text = rsTemp!nMonto
''             txtPlazo.Text = rsTemp!nPlazo
''
''        End If
''    Else
''        MsgBox "No se encontro Tasas Especiales Solicitadas por Titulares"
''    End If
''
''    Set rsTemp = Nothing
''Else
''    Label20.Visible = False
''    cboTasaEspecial.Clear
''    lblTasaEspecial.Caption = ""
''    lblTasaEspecial.Visible = False
''    cboTasaEspecial.Visible = False
''End If
'End Sub



Private Sub chkTasaPreferencial_Click()
   If chkTasaPreferencial.value = vbChecked Then
        lblTitSol.Visible = True
        txtNumSolicitud.Visible = True
        lblEstadoSol.Visible = True
        'chkPermanente.Visible = True
        txtPlazo.Enabled = True
        txtMonto.Enabled = False
        cboMoneda.Enabled = False
        txtNumSolicitud.Text = ""
        
   Else
        lblTitSol.Visible = False
        txtNumSolicitud.Visible = False
        lblEstadoSol.Visible = False
        
        lblEstadoSol.Caption = "ESTADO SOLICITUD"
        txtNumSolicitud.Text = ""
        
        txtPlazo.Enabled = True
        txtMonto.Enabled = True
        txtMonto.value = Format(0, "0.00")
        cboMoneda.Enabled = True
        cboMoneda.ListIndex = 0
        
        cboTipoTasa.ListIndex = 0
        cboTipoTasa_Click
        
        chkPermanente.Visible = False
        
        LimpiaControles
        
   End If
End Sub

Private Sub cmdAgregar_Click()
'***Agreado por ELRO por 20120313, según Acta N° 044-2012/TI-D
If txtCuenta.Prod = "233" And (Trim(Right(cboPrograma, 1)) = "2" Or Trim(Right(cboPrograma, 1)) = "3") Then
 Dim lsMsg As String
 lsMsg = ValidarMontoAbonar
    If lsMsg <> "" Then
        MsgBox lsMsg, vbInformation, "Aviso"
        txtMontoAbonar.SetFocus
        Exit Sub
    End If
lsMsg = ""
End If
'***Fin Agreado por ELRO

'Validacion para cuenta Individual
If grdCliente.TextMatrix(grdCliente.Rows - 1, 3) = "" And nClientes >= 1 Then
    MsgBox "Debe seleccionar la relacion con la cuenta", vbInformation, "Aviso"
    grdCliente.Col = 3
    grdCliente.SetFocus
    SendKeys "1|{ENTER}"
    Exit Sub
End If

' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES" ***

    ''ALPA 20091123***********************************
    ''If Trim(Right(cboTipoCuenta, 4)) = 0 Then
    'If Trim(Right(cboTipoCuenta, 4)) = 0 And Trim(Right(cboPrograma, 1)) <> 1 Then
    ''************************************************
    '  If Trim(grdCliente.TextMatrix(1, 3)) <> "" Then
    '    If CLng(Trim(Right(grdCliente.TextMatrix(1, 3), 4))) = gCapRelPersTitular Then
    '       MsgBox "Cuenta Individual solo permite un Participante", vbInformation, "Aviso"
    '       Exit Sub
    '    End If
    '   End If
    'End If
    '
    ''***Agregado por ELRO el 20120124, según Acta N° 005-2012/TI-D
    'If txtCuenta.Prod = "233" And Trim(Right(cboTipoCuenta, 4)) = 0 And Trim(Right(cboPrograma, 1)) = 1 Then
    '  If Trim(grdCliente.TextMatrix(1, 3)) <> "" Then
    '    If CLng(Trim(Right(grdCliente.TextMatrix(1, 3), 4))) = gCapRelPersTitular Then
    '       MsgBox "Cuenta Individual solo permite un Participante", vbInformation, "Aviso"
    '       Exit Sub
    '    End If
    '   End If
    'End If
    ''***Fin Modificado por ELRO*************************************
    'If ValidarFirmas = False Then Exit Sub
    
' *** FIN RIRO

' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES" ***

    grdCliente.AdicionaFila
    Dim i As Integer
    Dim intPJ As Integer
    intPJ = -1
    
    'Verifica si hay personas jurídicas dentro del grid
    For i = 1 To grdCliente.Rows - 1
        If grdCliente.TextMatrix(i, 9) = "PJ" Then
            intPJ = i
        End If
    Next
    
    'Aplica el grid de reglas para las PN que cumplan la condicion.
    If intPJ = -1 Then
        intPunteroPJ_NA = 0
        If grdCliente.Rows > 2 Then
            'cboTipoCuenta.ListIndex = 2
            grdCliente.ColumnasAEditar = "X-1-X-3-X-5-6-X-X-9"
            conReglas
        Else
            'cboTipoCuenta.ListIndex = 0
            grdCliente.ColumnasAEditar = "X-1-X-3-X-5-6-X-X-X"
            sinReglas
        End If
    Else
        sinReglas
    End If
    seleccionarTipoCuentaXregla

' *** FIN RIRO ***

'grdCliente.AdicionaFila
grdCliente.SetFocus
SendKeys "{ENTER}"
End Sub

Private Sub cmdCancelar_Click()
    LimpiaControles
    cboMoneda.SetFocus
End Sub

Private Sub cmdDocumento_Click()
    'frmCapAperturaListaChq.inicia frmCapAperturas, nOperacion, nmoneda, nProducto
    'Me.lblTotal.Caption = Format(txtMonto.value + CCur(Me.lblITF.Caption), "#,##0.00")
    Dim oform As New frmChequeBusqueda
    Dim lnOperacion As TipoOperacionCheque

    On Error GoTo ErrCargaDocumento
    If gsOpeCod = gPFApeChq Then
        lnOperacion = DPF_Apertura
    Else
        lnOperacion = Ninguno
    End If
    
    SetDatosCheque
    Set oDocRec = oform.Iniciar(nmoneda, lnOperacion)
    Set oform = Nothing
    If Not ValidaSeleccionCheque() Then
        Exit Sub
    End If
    SetDatosCheque oDocRec.fsNroDoc, oDocRec.fsPersNombre, oDocRec.fsDetalle, oDocRec.fsGlosa, oDocRec.fnMonto
    grdCliente.row = 1
    grdCliente.Col = 3
    grdCliente_OnEnterTextBuscar grdCliente.TextMatrix(1, 1), 1, 1, False
    Exit Sub
ErrCargaDocumento:
    MsgBox "Ha sucedido un error al cargar los datos del Documento", vbCritical, "Aviso"
End Sub

Private Sub cmdEliminar_Click()
'***Agregado por ELRO el 20120707, según OYP-RFC024-2012
If nOperacion = gAhoApeTransf Or nOperacion = gPFApeTransf Or nOperacion = gCTSApeTransf Then
    If fsPersCodTransfer = grdCliente.TextMatrix(grdCliente.row, 1) Then
        MsgBox "No se puede eliminar al Titular del Voucher.", vbInformation, "Aviso"
        Exit Sub
    End If
End If

' *** RIRO 20130411 ***
If grdCliente.TextMatrix(grdCliente.row, 1) = sPerSolicitud And sPerSolicitud <> "" Then
    Call MsgBox("No puedes eliminar un cliente asociado a la solicitud de tasa preferencial", vbOKOnly + vbExclamation, "AVISO")
    Exit Sub
End If
' *** Fin RIRO ***

'***Fin Agregado por ELRO*******************************
If MsgBox("¿¿Está seguro de eliminar a la persona de la relación??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    If grdCliente.TextMatrix(grdCliente.row, 4) <> "" Then
        Dim nPers As COMDConstantes.PersPersoneria
        nPers = CLng(grdCliente.TextMatrix(grdCliente.row, 4))
        If nPers <> gPersonaNat Then
            nPersoneria = gPersonaNat
            'MIOL 20121011, SEGUN OYP-RFC098-2012 ************
            chkSubasta.Visible = False
            chkSubasta.value = 0
            'END MIOL ****************************************
        End If
        grdCliente.EliminaFila grdCliente.row
    Else
        grdCliente.EliminaFila grdCliente.row
    End If
    
    seleccionarTipoCuentaXregla ' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES" ***
    
    'If Not txtPlazo.Enabled Then
    '    txtPlazo.Enabled = True
        lnTitularPJ = 0
    'End If
    CuentaTitular
    EvaluaTitular
    ValidaTasaInteres 'JUEZ 20160420
End If
End Sub

'****Agregado MPBR
Private Function ObtTitular() As String
Dim i As Integer
For i = 1 To grdCliente.Rows - 1
  If Right(grdCliente.TextMatrix(i, 3), 2) = "10" Then
      ObtTitular = Trim(grdCliente.TextMatrix(i, 1))
      Exit For
  End If
Next i
End Function

Private Function ObtTodosTitulares() As String
    Dim i As Integer
    Dim cTitulares As String
    For i = 1 To grdCliente.Rows - 1
      If Right(grdCliente.TextMatrix(i, 3), 2) = "10" Then
          cTitulares = cTitulares & Trim(grdCliente.TextMatrix(i, 1)) & ","
      End If
    Next i
    If Len(cTitulares) > 1 Then
        cTitulares = Mid(cTitulares, 1, Len(cTitulares) - 1)
    End If
    ObtTodosTitulares = cTitulares
End Function

Private Sub cmdGrabar_Click()
    Dim nProgramAhorro As Integer 'WIOR 20131106
    Dim nFirmas As Long
    Dim sInstitucion As String, sErrDesc As String, sCtaAbono As String
    Dim bOrdPag As Boolean, bCtaAboInt As Boolean
    Dim nMonto As Double
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sNroDoc As String
    Dim rsRel As New ADODB.Recordset
    Dim nTasa As Double
    Dim nPlazo As Long
    Dim sCodPromotor() As String
    Dim loLavDinero As frmMovLavDinero
    Set loLavDinero = New frmMovLavDinero
    
    '*** PEAC 20080811
    Dim lbResultadoVisto As Boolean
    Dim lbResultadoPersoneria As Boolean 'MIOL 20121113, SEGUN RFC098-2012
    Dim sPersVistoCod  As String
    Dim sPersVistoCom As String
    Dim loVistoElectronico As frmVistoElectronico
    Set loVistoElectronico = New frmVistoElectronico
    
    Dim objPersona As COMDPersona.DCOMPersonas 'JACA 20110512
    Set objPersona = New COMDPersona.DCOMPersonas 'JACA 20110512
    
    Dim objPers As COMDPersona.DCOMPersona 'MIOL 20121006, SEGUN RQ12272
    Set objPers = New COMDPersona.DCOMPersona 'MIOL 20121006, SEGUN RQ12272
    
    Dim loMov As COMDMov.DCOMMov 'BRGO 20110908
    Set loMov = New COMDMov.DCOMMov 'BRGO 20110908
    
    Dim oNCOMContImprimir As COMNContabilidad.NCOMContImprimir '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
    Set oNCOMContImprimir = New COMNContabilidad.NCOMContImprimir '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
    Dim lsPersNombreCVME As String '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
    Dim lsPersDireccionCVME As String '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
    Dim lsdocumentoCVME As String '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
    Dim lsPersCodConv As String 'Add By gitu 22-10-2012 aperturas en lote con convenios
    Dim lsDireccionActualizada As String '***Agregado por ELRO el 20130219, según INC1302150010
    'WIOR 20130301 **************************
    Dim fbPersonaReaAhorros As Boolean
    Dim fnCondicion As Integer
    Dim nI As Integer
    nI = 0
    'WIOR FIN *******************************
    Dim lsBoletaCargo As String 'JUEZ 20131212
    
    nMonto = txtMonto.value
    Gtitular = ObtTitular
    nTasa = nTasaNominal
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    bOrdPag = chkOrdenPago.value
    bCtaAboInt = txtCtaAhoAboInt.Visible
    grdCliente.ColWidth(5) = 0
    
    If ValidaFlexVacio Then
        MsgBox "Ingrese un Cliente", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'JUEZ 20131212 *****************************************************
    If nOperacion = gAhoApeCargoCta Or nOperacion = gPFApeCargoCta Then
        If Len(txtCuentaCargo.NroCuenta) <> 18 Then
            MsgBox "Debe ingresar la cuenta de ahorros a la que se va a debitar el monto de apertura", vbInformation, "Aviso"
            txtCuentaCargo.SetFocusCuenta
            Exit Sub
        End If
    End If
    'END JUEZ **********************************************************
    
    'JUEZ 20141008 Nuevos Parámetros *****************************
    Dim x As Integer
    If nProducto <> gCapCTS Then
        For x = 1 To grdCliente.Rows - 1
            If grdCliente.TextMatrix(x, 4) = gPersonaNat And Not bParPersNat Then
                MsgBox "El producto no permite ingresar personas naturales", vbInformation, "Aviso"
                'grdCliente.EliminaFila X
                'nClientes = nClientes - 1
                'If nClientes = 0 Then cmdEliminar.Enabled = False
                Exit Sub
            End If
            If grdCliente.TextMatrix(x, 4) <> gPersonaNat And Not bParPersJur Then
                MsgBox "El producto no permite ingresar personas jurídicas", vbInformation, "Aviso"
                'grdCliente.EliminaFila X
                'nClientes = nClientes - 1
                'If nClientes = 0 Then cmdEliminar.Enabled = False
                Exit Sub
            End If
        Next x
    End If
    'END JUEZ ****************************************************
    
     'Agregado por RIRO EL 20130411 *****
    If sPerSolicitud <> "" Then
        Dim n As Integer
        Dim b As Boolean
        b = False
        For n = 1 To grdCliente.Rows - 1
            If grdCliente.TextMatrix(n, 1) = sPerSolicitud Then
                b = True
            End If
        Next
        If Not b Then
            MsgBox "No puede concluir la operacion sin antes contar con el cliente asociado a la tasa especial", vbOKOnly + vbExclamation, "AVISO"
            Exit Sub
        End If
    End If
    ' Fin RIRO **

    If Me.chkExoITF.value = 1 And fraITF.Visible Then
        If Me.cboTipoExoneracion.Text = "" Then
            MsgBox "Debe ingresa un tipo de Exoneracion.", vbInformation, "Aviso"
            cboTipoExoneracion.SetFocus
            Exit Sub
        End If
    End If

    If chkTasaPreferencial.value = vbChecked Then
        If Trim(lblEstadoSol.Caption) <> "APROBADA" Then
            MsgBox "NO PUEDE APERTURAR ESTA CUENTA CON UNA TASA ESPECIAL SIN APROBAR.", vbOKOnly + vbExclamation, "AVISO"
            Exit Sub
        End If
    End If
    
    If chkRelConv.value = 1 Then 'Add By Gitu 22-10-2012
        If Trim(Right(cboPrograma, 3)) <> "1" And Trim(Right(cboPrograma, 3)) <> "4" And _
            Trim(Right(cboPrograma, 3)) <> "7" Then
        '***Condición cboPrograma.ListIndex <> 6 modificado por ELRO el 20130201
            lsPersCodConv = Trim(Right(Me.cboInstConvDep.Text, 13))
        Else
            lsPersCodConv = txtInstitucion.Text
        End If
    End If
    
    If Not ValidaRelaciones Then Exit Sub

    If Not ValidaUsuarios Then
        MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
        Unload Me
        Exit Sub
    End If
    
    If Not (nProducto = gCapAhorros And Me.cboPrograma.ListIndex = 7) Then 'BRGO 20111116 Para que no valide cuando es SubProducto Ahorro Ecotaxi
        If nTasa <= 0 Then
            MsgBox "Tasa tiene que ser mayor que cero.", vbInformation, "Aviso"
            cboMoneda.SetFocus
            Exit Sub
        End If
    End If
    '*** BRGO 20111219 Valida el monto mínimo del PF Premium *************************
'    If nProducto = gCapPlazoFijo And Me.cboPrograma.ListIndex = 1 Then
'        If Me.txtMonto.Text < nMontoMinimoPFPremium Then
'            MsgBox "Monto de Apertura debe ser > ó = a " & Format(nMontoMinimoPFPremium, "#,##0.00"), vbInformation, "Aviso"
'            txtMonto.SetFocus
'            Exit Sub
'        End If
'    End If
    '*** END BRGO *********************************************************************
    
    'Validacion de ITF
    If Me.chkExoITF.value = 1 And Me.cboTipoExoneracion.Text = "" Then
        MsgBox "Debe Elegir un tipo de exoneracion Valido.", vbInformation, "Aviso"
        Exit Sub
    End If
        
    If nOperacion <> gAhoApeCargoCta And nOperacion <> gPFApeCargoCta Then 'JUEZ 20131212
        'Validacion por Gestor Add GITU 22-09-2009
        If Me.cboPromotor.Text = "" And txtCuenta.Prod <> "232" Then
            If MsgBox("No elgio el Gestor de Cartera, ¿Desea agregar algun Gestor?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                Exit Sub
            End If
        End If
    End If
    
    'Si el producto es PlazoFijo verificar el plazo
    If nProducto = gCapPlazoFijo Then
        'JUEZ 20141008 Nuevos Parámetros **************************************
        'If CLng(txtPlazo) < 30 And Right(Me.cboFormaRetiro.Text, 1) = 1 Then
        '    MsgBox "Si la forma de Retiro es mensual, el Plazo debe ser mayor o igual a Treinta dias.", vbInformation, "Aviso"
        '    txtPlazo.SetFocus
        '    Exit Sub
        'ElseIf CLng(txtPlazo) <= 0 Then
        '    MsgBox "El Plazo debe ser mayor a 0 dias.", vbInformation, "Aviso"
        '    txtPlazo.SetFocus
        '    Exit Sub
        'End If
        If chkSubasta.value = 0 And Not bInstFinanc Then
            If Not ValidarPlazoPF Then Exit Sub
        End If
        If Not ValidarMedioRetiroPF Then Exit Sub
        'END JUEZ *************************************************************
        If bCtaAboInt Then
            If Trim(txtCtaAhoAboInt.Text) = "" Then
                MsgBox "Debe seleccionar la cuenta de ahorros para el abono de lo intereses.", vbInformation, "Aviso"
                txtCtaAhoAboInt.SetFocus
                Exit Sub
            Else
                sCtaAbono = txtCtaAhoAboInt.Text
            End If
        End If
    End If

    If chkEspecial.value = vbChecked Then
        Dim nMontoCTSVal As Double
        nMontoCTSVal = CDbl((Val(txtInta.Text) + Val(txtDisp.Text) + Val(txtDU.Text)))
        If nOperacion = gCTSApeChq Or nOperacion = gCTSApeTransf Then
        
            If Format(nMontoCTSVal, "0.00") > Format(vnMontoDOC, "0.00") Then
                MsgBox "El valor distribuido para la apertura no debe ser mayor al monto de la transacción", vbOKOnly + vbExclamation, "AVISO"
                Exit Sub
            ElseIf Format(nMontoCTSVal, "0.00") > Format(vnMontoDOC, "0.00") Then
                MsgBox "El valor distribuido para la apertura no debe ser menor al monto de la transacción", vbOKOnly + vbExclamation, "AVISO"
                Exit Sub
            End If
        End If
    End If

    'Verifica si el numero de firmas corresponde al numero de titulares
    If txtNumFirmas = "" Then
        MsgBox "Número de firmas no válido", vbInformation, "Aviso"
        txtNumFirmas.SetFocus
        Exit Sub
    End If

    If nProducto = gCapCTS Then
        sInstitucion = Trim(txtInstitucion.Text)
        If sInstitucion = "" Then
            MsgBox "Institución No Válida", vbInformation, "Aviso"
            txtInstitucion.SetFocus
            Exit Sub
        End If
        Dim i As Integer, nTpoCta As COMDConstantes.ProductoCuentaTipo
        For i = 0 To cboTipoCuenta.ListCount - 1
            nTpoCta = CLng(Right(cboTipoCuenta.List(i), 4))
            If nTpoCta = gPrdCtaTpoIndiv Then
                cboTipoCuenta.ListIndex = i
                Exit For
            End If
        Next i
        nTipoCuenta = gPrdCtaTpoIndiv
    Else
        nTipoCuenta = CLng(Trim(Right(cboTipoCuenta.Text, 4)))
        If Trim(Right(cboPrograma.Text, 1)) = 4 Then
          sInstitucion = Trim(txtInstitucion.Text)
          If sInstitucion = "" Then
                MsgBox "Institución No Válida", vbInformation, "Aviso"
                txtInstitucion.SetFocus
                Exit Sub
          End If
          'JUEZ 20141008 *********************************************
          'If val(txtMontoAbonar.Text) <> val(txtMonto.Text) Then
          '      MsgBox "El Monto de Apertura Tiene que ser Igual al de Abono", vbInformation, "Aviso"
          '
          '      If txtMonto.Enabled Then
          '          txtMonto.SetFocus
          '      End If
          '
          '      Exit Sub
          'End If
          'END JUEZ **************************************************
        Else
          sInstitucion = ""
        End If
    End If
    If Trim(Right(cboPrograma, 1)) = "6" Then
         sInstitucion = Trim(txtInstitucion.Text)
          If sInstitucion = "" Then
                MsgBox "Institución No Válida", vbInformation, "Aviso"
                txtInstitucion.SetFocus
                Exit Sub
          End If
    End If
    nFirmas = CLng(txtNumFirmas)
    
    ' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES" ***
    
            'If nFirmas = 0 Then
            '    MsgBox "Número de Firmas no puede ser 0.", vbInformation, "Aviso"
            '    txtNumFirmas.SetFocus
            '    Exit Sub
            'End If
        
    ' *** FIN RIRO ***
    
    'By Capi 19082008 para que valide la persona natural de cuenta soñada
    If nProducto = gCapAhorros And cboPrograma.ListIndex = 5 Then
        If nPersoneria <> gPersonaNat Then
            MsgBox "Cuenta Soñada, solo para personas naturales", vbInformation, "Aviso"
            grdCliente.SetFocus
            Exit Sub
        End If
        
    End If
    '
    If nPersoneria = gPersonaNat Then
        If nTitular = 0 Then
            MsgBox "No existen titulares en la cuenta.", vbInformation, "Aviso"
            grdCliente.SetFocus
            Exit Sub
        End If
        
    ' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES" ***
    
            'If nTipoCuenta = gPrdCtaTpoMancom Or nTipoCuenta = gPrdCtaTpoIndiv Then
            '    If nTitular <> nFirmas Then 'Valida # de Firmas
            '        MsgBox "Número de firmas difiere del número de titulares.", vbInformation, "Aviso"
            '        txtNumFirmas.SetFocus
            '        Exit Sub
            '    End If
            'ElseIf nTipoCuenta = gPrdCtaTpoIndiv Then
            '    If nTitular = 1 And val(TxtMinFirmas.Text) > 1 Then 'Valida # de Firmas
            '        MsgBox "Número de firmas no corresponde con el tipo de cuenta.", vbInformation, "Aviso"
            '        txtNumFirmas.SetFocus
            '        Exit Sub
            '    End If
            'End If
    
    ' *** FIN RIRO
    Else
        If nTitular > 1 Then
            MsgBox "La cuenta con personería jurídica no debe tener mas de un titular.", vbInformation, "Aviso"
            grdCliente.SetFocus
            Exit Sub
        End If
        If nRepresentante = 0 Then
            MsgBox "No existen representantes en la cuenta.", vbInformation, "Aviso"
            grdCliente.SetFocus
            Exit Sub
        End If
        
    ' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES" ***
    
        'If nTipoCuenta = gPrdCtaTpoMancom Then
        '    If nRepresentante < 2 Then
        '        MsgBox "Número de Representantes a firmar es menor a 2.", vbInformation, "Aviso"
        '        txtNumFirmas.SetFocus
        '        Exit Sub
        '    Else
        '        If nRepresentante < nFirmas Then  'Valida # de Firmas
        '            MsgBox "Número de firmas excede del número de posibles Representantes a firmar.", vbInformation, "Aviso"
        '            txtNumFirmas.SetFocus
        '            Exit Sub
        '        End If
        '    End If
        'Else
        '    If nFirmas > nRepresentante Then   'Valida # de Firmas
        '        MsgBox "Número de firmas no corresponde con el tipo de cuenta.", vbInformation, "Aviso"
        '        txtNumFirmas.SetFocus
        '        Exit Sub
        '    End If
        'End If
    
    ' *** FIN RIRO ***
               
    End If

    If nTitular = 1 Then 'Valida Tipo de Cuenta
        If nTipoCuenta <> gPrdCtaTpoIndiv And nPersoneria = gPersonaNat Then
            
            ' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES" ***
            
                'MsgBox "La cuenta posee un solo titular por lo que debe ser Individual.", vbInformation, "Aviso"
                'cboTipoCuenta.SetFocus
                'Exit Sub
                
            ' *** FIN RIRO ***
        End If
    Else
        If nTipoCuenta <> gPrdCtaTpoIndist And nTipoCuenta <> gPrdCtaTpoMancom Then
            MsgBox "La cuenta posee más de un titular por lo que no puede ser Individual.", vbInformation, "Aviso"
            cboTipoCuenta.SetFocus
            Exit Sub
        End If
    End If

    'Valida si la operacion es con documentos y si el documento es valido
    If bDocumento Then
        If Trim(lblNroDoc.Caption) = "" Then
            MsgBox "Debe seleccionar un documento (" & fraDocumento.Caption & ") válido para la operacion.", vbInformation, "Aviso"
            cmdDocumento.SetFocus
            Exit Sub
        End If
        sNroDoc = Trim(lblNroDoc.Caption)
    End If

    'Para Desembolso Abono A Cuenta
    '***************************************************************************************
    If vbDesembolso Then
        'ARCV 13-02-2007
        'Set rsRel = grdCliente.GetRsNew
        'Set vMatRela = rsRel
        vnTasa = nTasa
        vnPersoneria = nPersoneria
        vnTipoCuenta = nTipoCuenta
        vnTipoTasa = nTipoTasa
        vbDocumento = bDocumento
        vsNroDoc = sNroDoc
        vsCodIF = sCodIF
        'Call cmdSalir_Click
        Exit Sub
    End If
    '***************************************************************************************
    
    'Valida Promotor
    If chkPromotor.value = vbChecked Then
        If cboPromotor.Text = "" Then
            MsgBox "Debe seleccionar un promotor.", vbInformation, "Aviso"
            'cboPromotor.SetFocus
            Exit Sub
        Else
             sCodPromotor = Split(cboPromotor.Text, "|")
        End If
    End If
    
    'Valida Montos y Plazos para los tipo de Apertura de Ahorros AVMM -- 21-02-2007
    If nProducto = gCapAhorros Then
        If ValidaMonPlazoAho = False Then Exit Sub
    End If

    'Valida Monto mínimo de apertura
    'JUEZ 20141008 Nuevos parámetros **************************
    'If nProducto = gCapAhorros Then
    '    If cboPrograma.ListIndex = -1 Then
    '        MsgBox "Debe de Seleccionar un tipo de Sub Producto para AHORRO", vbInformation
    '    End If
    '    sErrDesc = clsCap.ValidaMontoApertura(nProducto, nPersoneria, nMonto, nMoneda, chkOrdenPago.value)
    '
    'Else
        '*** BRGO 20111219 ****************************************************
        'If nProducto = gCapPlazoFijo And cboPrograma.ListIndex = 1 Then
    '        sErrDesc = clsCap.ValidaMontoApertura(nProducto, nPersoneria, nMonto, nMoneda, , cboPrograma.ListIndex)
        'Else
        '    sErrDesc = clsCap.ValidaMontoApertura(nProducto, nPersoneria, nMonto, nMoneda)
        'End If
        '*** END BRGO *********************************************************
    If nProducto <> gCapCTS Then If Not ValidaMontoMinimoApertura Then Exit Sub
    'END JUEZ *************************************************

    If sErrDesc <> "" Then
        MsgBox sErrDesc, vbInformation, "Aviso"
        If txtMonto.Enabled Then
            txtMonto.SetFocus
        Else
            '***Agregado por ELRO el 20120823, según OYP-RFC024-2012
            'cmdDocumento.SetFocus
            If cmdDocumento.Visible And cmdDocumento.Enabled Then
                cmdDocumento.SetFocus
            End If
            '***Fin Agregado por ELRO el 20120823*******************
        End If
        Exit Sub
    End If

    If nOperacion = gAhoApeTransf Or nOperacion = gPFApeTransf Or nOperacion = gCTSApeTransf Then
        If lblTrasferND.Caption = "" Then
            MsgBox "Debe ingresar un numero de transferencia.", vbInformation, "Aviso"
            cmdTranfer.SetFocus
            Exit Sub
        End If
    End If
    
    'JUEZ 20130723 ******************************************************************
    If nProducto = gCapCTS Then
        Dim lsPersCodTitular As String
        Dim oCap As New COMNCaptaGenerales.NCOMCaptaGenerales
        For nI = 1 To grdCliente.Rows - 1
            If Trim(Right(grdCliente.TextMatrix(nI, 3), 2)) = "10" Then
                lsPersCodTitular = grdCliente.TextMatrix(nI, 1)
                Exit For
            End If
        Next nI
        If oCap.VerificarExisteCuentaCTS(lsPersCodTitular, txtInstitucion.Text, CInt(Trim(Right(cboMoneda.Text, 1)))) Then
            MsgBox "No es posible realizar la operación debido a que ya existe otra cuenta CTS del cliente con el mismo empleador y con la misma moneda", vbInformation, "Aviso"
            Exit Sub
        End If
        Set oCap = Nothing
    End If
    'END JUEZ ***********************************************************************
    
    'JUEZ 20131212 **************************************************************
    If nOperacion = gAhoApeCargoCta Or nOperacion = gPFApeCargoCta Then
        If nTipoCuenta <> fnTpoCtaCargo Then
            MsgBox "Cuenta a debitar debe tener el mismo Tipo de Cuenta de la apertura", vbInformation, "Aviso"
            Exit Sub
        End If
        If Not ValidaRelPersonasCtaCargo Then
            MsgBox "Las personas y relaciones de la cuenta a debitar deben ser las mismas que las de la apertura", vbInformation, "Aviso"
            'txtCuentaCargo.SetFocusCuenta
            LimpiaControlesCargoCta
            Exit Sub
        End If
        If nMonto <= 0 Then
            MsgBox "Monto debe ser mayor a 0", vbInformation, "Aviso"
            Exit Sub
        End If
        If Not clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, nMonto) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "Aviso"
            Exit Sub
        End If
        
        If VerificarAutorizacion = False Then Exit Sub
    End If
    'END JUEZ *******************************************************************
    
    '*** PEAC 20080811 ******************************************************
   
    For i = 1 To grdCliente.Rows - 1
        lbResultadoVisto = loVistoElectronico.Inicio(1, nProducto, grdCliente.TextMatrix(i, 1))
        If Not lbResultadoVisto Then
            Exit Sub
        End If
    Next i
    '*** FIN PEAC ************************************************************
    
    'EJVG20120322 Verifica actualización Persona
    For i = 1 To grdCliente.Rows - 1
        Dim oPersona As New COMNPersona.NCOMPersona
        Call VerSiClienteActualizoAutorizoSusDatos(grdCliente.TextMatrix(i, 1), nOperacion) 'FRHU ERS077-2015 20151204
        If oPersona.NecesitaActualizarDatos(grdCliente.TextMatrix(i, 1), gdFecSis) Then
             MsgBox "Para continuar con la Operación Ud. debe actualizar los datos del" & Chr(13) & Trim(Left(grdCliente.TextMatrix(i, 3), 50)) & ": " & grdCliente.TextMatrix(i, 2), vbInformation, "Aviso"
             Dim foPersona As New frmPersona
             If Not foPersona.realizarMantenimiento(grdCliente.TextMatrix(i, 1), lsDireccionActualizada) Then
                 MsgBox "No se ha realizado la actualización de los datos de " & grdCliente.TextMatrix(i, 2) & "," & Chr(13) & "la Operación no puede continuar!", vbInformation, "Aviso"
                 Exit Sub
             End If
             '***Agregado por ELRO el 20130219, según INC1302150010
             If Trim(lsDireccionActualizada) <> "" Then
                grdCliente.TextMatrix(i, 8) = lsDireccionActualizada
             End If
             '***Fin Agregado por ELRO el 20130219*****************
        End If
        lsDireccionActualizada = ""
    Next
    
    'JACA 20110512 *****VERIFICA SI LAS PERSONAS CUENTAN CON OCUPACION E INGRESO PROMEDIO
        Dim rsPersVerifica As Recordset
        Set rsPersVerifica = New Recordset
        For i = 1 To grdCliente.Rows - 1
            Set rsPersVerifica = objPersona.ObtenerDatosPersona(Me.grdCliente.TextMatrix(i, 1))
            If rsPersVerifica!nPersIngresoProm = 0 Or rsPersVerifica!cActiGiro1 = "" Then
                If MsgBox("Necesita Registrar la Ocupacion e Ingreso Promedio de " + Me.grdCliente.TextMatrix(i, 2), vbYesNo) = vbYes Then
                    'frmPersona.Inicio Me.grdCliente.TextMatrix(i, 1), PersonaActualiza
                    frmPersOcupIngreProm.Inicio Me.grdCliente.TextMatrix(i, 1), Me.grdCliente.TextMatrix(i, 2), rsPersVerifica!cActiGiro1, rsPersVerifica!nPersIngresoProm
                End If
            End If
        Next i
    'JACA END***************************************************************************

    '***Agreado por ELRO por 20120313, según Acta N° 044-2012/TI-D
    If txtCuenta.Prod = "233" And (Trim(Right(cboPrograma, 1)) = "2" Or Trim(Right(cboPrograma, 1)) = "3") Then
     Dim lsMsg As String
     lsMsg = ValidarMontoAbonar
        If lsMsg <> "" Then
            MsgBox lsMsg, vbInformation, "Aviso"
            txtMontoAbonar.SetFocus
            Exit Sub
        End If
    lsMsg = ""
    End If
        '***Fin Agreado por ELRO
    'EJVG20130913 ***
    '***Agregado por ELRO el 20120713, según OYP-RFC024-2012
    'If nOperacion = gAhoApeTransf Or nOperacion = gPFApeTransf Or nOperacion = gCTSApeTransf Then
    '    For i = 1 To grdCliente.Rows - 1
    '         If fsPersCodTransfer = grdCliente.TextMatrix(i, 1) Then
    '            lsPersNombreCVME = grdCliente.TextMatrix(i, 2)
    '            lsdocumentoCVME = grdCliente.TextMatrix(i, 7)
    '            lsPersDireccionCVME = grdCliente.TextMatrix(i, 8)
    '            Exit For
    '         End If
    '         If i = grdCliente.Rows - 1 Then
    '            MsgBox "Debe ingresar el Tilular del Voucher.", vbOKOnly + vbInformation, "AVISO"
    '            Exit Sub
    '         End If
    '    Next i
    'End If
    '***Fin Agregado por ELRO*******************************
    'END EJVG *******
    '***Agregado por ELRO el 20120919, según OYP-RFC087-2012
    If nOperacion = gAhoApeEfec Or nOperacion = gPFApeEfec Or nOperacion = gCTSApeEfec Or _
       nOperacion = gAhoApeChq Or nOperacion = gPFApeChq Or nOperacion = gCTSApeChq Or _
       nOperacion = gAhoApeTransf Or nOperacion = gPFApeTransf Or nOperacion = gCTSApeTransf Then
       Dim lnIdVoBoConCli As Long
       Dim lnfilaIndConCli, lnEstadoVoBo As Integer
       Dim cMovNroVoBoConCli As String
       Dim lnIndicadorInterno As Currency
       Dim lnIndicadorPrimerCliente As Currency
       Dim lnIndicadorDiezCliente As Currency
       Dim lnIndicadorVeinteCliente As Currency
       Dim lnSaldoTotalDepositos As Currency

       cMovNroVoBoConCli = loMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)


            For lnfilaIndConCli = 1 To grdCliente.Rows - 1
                If UCase(Left(grdCliente.TextMatrix(lnfilaIndConCli, 3), 7)) = "TITULAR" Then
                    clsCap.devolverIndicadorCliente grdCliente.TextMatrix(lnfilaIndConCli, 1), _
                                                    gnTipCambio, _
                                                    Format(DateAdd("d", -1, gdFecSis), "yyyyMMdd"), _
                                                    nMonto, _
                                                    Right(gsCodAge, 2), _
                                                    CStr(txtCuenta.Prod), _
                                                    nTasa, _
                                                    cMovNroVoBoConCli, _
                                                    lnIndicadorInterno, _
                                                    lnIndicadorPrimerCliente, _
                                                    lnIndicadorDiezCliente, _
                                                    lnIndicadorVeinteCliente, _
                                                    lnSaldoTotalDepositos, _
                                                    lnIdVoBoConCli

                                                    
                    If lnIndicadorPrimerCliente > 0 Then
                         MsgBox "No se puede realizar esta operación porque supero el Indicador 1 de la Concentración de Clientes." & Chr(10) & "Coordinar con el Departamento de Ahorros y Servicios.", vbInformation, "Aviso"
                         lnIndicadorInterno = 0
                         lnIndicadorPrimerCliente = 0
                         lnIndicadorDiezCliente = 0
                         lnIndicadorVeinteCliente = 0
                         lnSaldoTotalDepositos = 0
                         lnIdVoBoConCli = 0
                         cMovNroVoBoConCli = ""
                         Exit Sub
                    End If

                   If (lnIndicadorDiezCliente > 0 Or lnIndicadorVeinteCliente > 0) And lnIdVoBoConCli > 0 Then
                        While lnEstadoVoBo = 0
                            MsgBox "En espera de Aprobación/Rechazo de la Jefatura de Ahorros y Servicios.", vbInformation, "Aviso"
                            lnEstadoVoBo = clsCap.DevolverVoBoConcentracionCliente(lnIdVoBoConCli)
                        Wend

                         If lnEstadoVoBo = 2 Then
                            MsgBox "La Jefatura de Ahorros ha rechazado la solicitud de apertura y la operación no puede continuar.", vbInformation, "Aviso"
                            lnIndicadorInterno = 0
                            lnIndicadorPrimerCliente = 0
                            lnIndicadorDiezCliente = 0
                            lnIndicadorVeinteCliente = 0
                            lnSaldoTotalDepositos = 0
                            lnIdVoBoConCli = 0
                            cMovNroVoBoConCli = ""
                            Exit Sub
                        End If
                        'JUEZ 20140807 *********************************************
                        Dim oNcapRep As New COMNCaptaGenerales.NCOMCaptaReportes
                            lnIdVoBoConCli = oNcapRep.modificarVoBo(lnIdVoBoConCli, lnEstadoVoBo, cMovNroVoBoConCli)
                        Set oNcapRep = Nothing
                        'END JUEZ **************************************************
                   End If

                End If
            Next lnfilaIndConCli

            lnIndicadorInterno = 0
            lnIndicadorPrimerCliente = 0
            lnIndicadorDiezCliente = 0
            lnIndicadorVeinteCliente = 0
            lnSaldoTotalDepositos = 0
            lnIdVoBoConCli = 0
            cMovNroVoBoConCli = ""
    End If
    '***Fin Agregado por ELRO el 20120919*******************
    'MIOL 20121006, SEGUN RQ12272 **************************
    If nProducto = gCapAhorros And cboPrograma.ListIndex = 0 And chkOrdenPago.value = 1 Then
        If objPers.ObtenerPersonaBloqueadaxSobreGiro_OrdPag(Gtitular) Then
            MsgBox "La emisión de orden de pago no puede proceder debido a que el cliente ha sido penalizado por sobregiro en el último año", vbInformation, "Aviso"
            Unload Me
            Exit Sub
        End If
    End If
    'END MIOL **********************************************
    'WIOR 20121009 Clientes Observados ****************************************************
    If nOperacion = gAhoApeEfec Or nOperacion = gAhoApeChq Or nOperacion = gPFApeEfec Or nOperacion = gPFApeChq Or nOperacion = gCTSApeEfec Or nOperacion = gCTSApeChq Then
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
    'WIOR FIN ********************************************************************************
    
    'MIOL 20121113, SEGUN RFC098-2012 ***********************************************
    If nProducto = gCapPlazoFijo And cboFormaRetiro.ListIndex = 1 And chkSubasta.value = 1 Then
        Dim oDPersoneria As COMDPersona.DCOMInstFinac
        Set oDPersoneria = New COMDPersona.DCOMInstFinac
        Dim nPerSubJur As Integer
        For i = 1 To grdCliente.Rows - 1
            lbResultadoPersoneria = oDPersoneria.PersoneriaSubasta(grdCliente.TextMatrix(i, 1))
            If lbResultadoPersoneria Then
                nPerSubJur = nPerSubJur + 1
            End If
        Next i
        If nPerSubJur <> 1 Then
            MsgBox "Verificar Datos del cliente para Subasta - Incorrectos", vbInformation, "Aviso"
            nPerSubJur = 0
            Exit Sub
        End If
        If chkSubasta.Visible = True And chkSubasta.value = 1 Then
            If Not ((CLng(Trim(Right(cboTipoCuenta.Text, 4))) = 1 Or CLng(Trim(Right(cboTipoCuenta.Text, 4))) = 2) And CLng(Trim(Right(cboFormaRetiro.Text, 4))) = 2) Then
                MsgBox "Verificar el Tipo Cuenta, la Apertura para subasta debe ser Mancomunada o Indistinta", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    End If
    'END MIOL ***********************************************************************
    
    ' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES" ***
    
    Dim J, nLetra, nContar, nLetraMax, nRelacion, nPJuridica As Integer
    Dim bOrden As Boolean
        
    ' Agregado Por RIRO el 20130501
    If Val(Right(Trim(cboPrograma.Text), 1)) = 6 Then
        If grdCliente.Rows > 2 Or cboTipoCuenta.ListIndex <> 0 Then
            MsgBox "El tipo de cuenta debe ser individual, registrar SOLO un titular", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    ' Fin RIRO
    
    ' Validando Mayoria de edad
    Dim oPersonaTemp As COMNPersona.NCOMPersona
    Dim iTemp, nMenorEdad, nIntervinientes As Integer
    
    Set oPersonaTemp = New COMNPersona.NCOMPersona
    
    For iTemp = 1 To grdCliente.Rows - 1
        If Val(Trim(grdCliente.TextMatrix(iTemp, 4))) <= 1 And Trim(grdCliente.TextMatrix(iTemp, 1)) <> "" Then
            If oPersonaTemp.validarPersonaMayorEdad(grdCliente.TextMatrix(iTemp, 1), Format(gdFecSis, "dd/mm/yyyy")) = False _
               And grdCliente.TextMatrix(iTemp, 9) <> "PJ" Then
                
                If Right(grdCliente.TextMatrix(iTemp, 3), 2) = "11" Or _
                   Right(grdCliente.TextMatrix(iTemp, 3), 2) = "12" Or _
                   Right(grdCliente.TextMatrix(iTemp, 3), 2) = "13" Then
                   
                   MsgBox "Un menor de edad no debe tener la relacion de Apoderado ni de Representante Legal ", vbInformation, "Aviso"
                   Exit Sub
                   
                End If
                
                nMenorEdad = nMenorEdad + 1
                i = iTemp
            Else
                If grdCliente.TextMatrix(iTemp, 9) <> "PJ" Then
                    nIntervinientes = nIntervinientes + 1
                End If
            End If
        End If
    Next
    
    If nMenorEdad >= 1 Then
    
        If nMenorEdad > 1 Then
            MsgBox "No es posible aperturar una cuenta con mas de un menor de edad", vbInformation, "Aviso"
            Exit Sub
        ElseIf nProducto = gCapCTS Then
            MsgBox "No es posible agregar menores de edad en cuentas CTS", vbInformation, "Aviso"
            Exit Sub
        ElseIf nProducto = gCapAhorros And Val(Right(cboPrograma.Text, 2)) = 6 Then
            MsgBox "No es posible agregar menores de edad en cuentas Caja Sueldo", vbInformation, "Aviso"
            Exit Sub
        ElseIf nProducto = gCapAhorros And Val(Right(cboPrograma.Text, 2)) = 7 Then
            MsgBox "No es posible agregar menores de edad en cuentas Ahorro Ecotaxi", vbInformation, "Aviso"
            Exit Sub
        ElseIf nProducto = gCapAhorros And Val(Right(cboPrograma.Text, 2)) = 8 Then
            MsgBox "No es posible agregar menores de edad en cuentas Ahorro Convenio", vbInformation, "Aviso"
            Exit Sub
        ElseIf intPunteroPJ_NA = 1 Then
            MsgBox "No es posible agregar menores de edad en cuentas con personería jurídica", vbInformation, "Aviso"
            Exit Sub
        ElseIf nIntervinientes = 0 Then
            MsgBox "No es posible crear cuentas donde sólo intervengan menores de edad", vbInformation, "Aviso"
            Exit Sub
        Else
            If Val(Trim(Right(grdCliente.TextMatrix(i, 3), 2))) <> 10 Then
                MsgBox "Se agregó un menor de edad, el cual debe tener la relacion: Titular", vbInformation, "Aviso"
                grdCliente.SetFocus
                Exit Sub
            End If
          
        End If
    Else
                        
    End If
    
    ' Validando cuentas de tipo Ahorro Ñañito
    If nProducto = gCapAhorros Then
        If Val(Trim(Right(cboPrograma.Text, 5))) = 1 Then
            If Val(Trim(Right(cboTipoCuenta.Text, 5))) <> gPrdCtaTpoIndiv Then
                MsgBox "La cuenta que se está aperturando debe ser Individual, solo un menor de edad debe ser titular y los demás intervinientes deberán ser apoderados", vbInformation, "Aviso"
                Exit Sub
            End If
            If nMenorEdad = 0 Then
                MsgBox "La cuenta debe ser aperturada por un menor de edad y uno o mas apoderados", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    End If
        
    'Validando que todos los clientes tengan un grupo
    For i = 1 To Me.grdCliente.Rows - 1
        If Me.grdCliente.TextMatrix(i, 9) = "" Then
            MsgBox "Debe asignar un grupo a cada cliente", vbExclamation, "Mensaje"
            Exit Sub
        End If
    Next
    If grdCliente.Rows = 2 Then
        If nPersoneria <> gPersonaNat Then
            MsgBox "Persona Jurídica requiere como mínimo un representante", vbCritical, "Mensaje"
            Exit Sub
        End If
    ElseIf validaExistenciaReglas = False And nProducto <> gCapCTS Then
        MsgBox "Verificar si los grupos asignados forman parte de alguna regla o si las reglas contienen los grupos asignados", vbExclamation, "Mensaje"
        Exit Sub
    End If
    For i = 1 To grdCliente.Rows - 1
        If Val(Trim(grdCliente.TextMatrix(i, 4))) > 1 Then
            nPJuridica = nPJuridica + 1
            If Trim(Left(grdCliente.TextMatrix(i, 3), 10)) <> "TITULAR" Then
                nRelacion = nRelacion + 1
            End If
            For J = 1 To grdCliente.Rows - 1
                If Trim(Left(grdCliente.TextMatrix(J, 3), 10)) = "TITULAR" And Val(Trim(grdCliente.TextMatrix(J, 4))) = 1 Then
                    grdCliente.TextMatrix(J, 3) = ""
                    nRelacion = nRelacion + 1
                End If
            Next
        End If
    Next
    If nRelacion > 0 Then
        MsgBox "Solo la persona jurídica debe ser el titular de la cuenta", vbExclamation, "Aviso"
        grdCliente.SetFocus
        Exit Sub
    End If
    If nPJuridica > 1 Then
        MsgBox "No es posible relacionar dos personas jurídicas en una misma cuenta", vbInformation, "Aviso"
        grdCliente.SetFocus
        Exit Sub
    End If
    'Si en la apertura solo intervienen personas naturales, hay que validar que los grupos sean correlativos.
    If intPunteroPJ_NA = 0 Then
        nLetraMax = 65
        'Obteniendo la letra mayor
        For J = 1 To grdCliente.Rows - 1
            If Trim(grdCliente.TextMatrix(J, 9)) <> "AP" Then
                If CInt(AscW(grdCliente.TextMatrix(J, 9))) > nLetraMax Then
                    nLetraMax = CInt(AscW(grdCliente.TextMatrix(J, 9)))
                End If
            End If
        Next
        'Verificado que las letras (Sean consecutivas)
        For nLetra = 65 To nLetraMax
            nContar = 0
            For J = 1 To grdCliente.Rows - 1
                If Chr(nLetra) = grdCliente.TextMatrix(J, 9) Then
                    nContar = nContar + 1
                End If
            Next
            If nContar = 0 Then
                MsgBox "La secuencia de los grupos deben ser: A, B, C, D ...", _
                vbExclamation, "Aviso"
                Exit Sub
            End If
        Next
    End If
    seleccionarTipoCuentaXregla
    
    '*** Fin RIRO ***
    
    'AMDO 20130702 TI-ERS063-2013 ****************************************************
    If nOperacion = gAhoApeEfec Or nOperacion = gPFApeEfec Or nOperacion = gCTSApeEfec Then
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
    
    'JUEZ 20131212 *******************************************************************
    If nOperacion = gAhoApeCargoCta Or nOperacion = gPFApeCargoCta Then
        If nTipoCuenta = gPrdCtaTpoIndist Or nTipoCuenta = gPrdCtaTpoMancom Then
            If Not frmCapConfirmPoderes.Inicia(txtCuentaCargo.NroCuenta, gCapAhorros, nOperacion, "Débito Apertura") Then
                Exit Sub
            End If
        End If
    End If
    'END JUEZ ************************************************************************

    'WIOR 20131106 ***************************
     If Trim(cboPrograma.Text) = "" Then
        nProgramAhorro = 0
    Else
        nProgramAhorro = CInt(Right(Trim(cboPrograma.Text), 1))
    End If
    'WIOR FIN ********************************
    'EJVG20140408 ***
    If nOperacion = gAhoApeChq Or nOperacion = gPFApeChq Or nOperacion = gCTSApeChq Then
        If Not ValidaSeleccionCheque Then
            MsgBox "Ud. debe seleccionar un Cheque para continuar", vbInformation, "Aviso"
            If cmdDocumento.Visible And cmdDocumento.Enabled Then cmdDocumento.SetFocus
            Exit Sub
        End If
    End If
    'END EJVG *******

    If MsgBox("¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Dim nformaretiro As COMDConstantes.CaptacPFFormaRetiro
        Dim sMovNro As String, sCuenta As String, sPersLavDinero As String, sReaPersLavDinero As String, sBenPersLavDinero As String
        Dim clsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
        Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
        Dim clsprevio As previo.clsprevio
        Dim clsLibret As previo.clsprevio
        Dim nPorcRetCTS As Double, nMontoLavDinero As Double, nTC As Double
        Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion 'nCapDefinicion
        Dim valCombo As Integer
        Dim lsCadImp As String
        Dim lbOk As Boolean
        Dim vnTpoProg As Integer
        Dim pnMovNro As Long '*** PEAC 20080812
        Dim lsBoletaCVME As String '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
        'JUEZ 20130520 *********************
        Dim rsEnvEstCta As ADODB.Recordset
        Dim nModoEnvioEstCta As Integer
        Dim nDebitoMismaCta As Integer
        'EN JUEZ ***************************
        Dim ArrDatos As Variant 'EJVG20140408
        
    '    valCombo = 0
    '    If Right(cboTipoTasa.Text, 3) = 101 Then
    '        valCombo = cboTasaEspecial.ItemData(cboTasaEspecial.ListIndex)
    '    End If
        
        'JUEZ 20130520 *********************************************
        frmEnvioEstadoCta.InicioCap nProducto, grdCliente.GetRsNew()
        If frmEnvioEstadoCta.RegistraEnvio = False Then Exit Sub
        Set rsEnvEstCta = frmEnvioEstadoCta.RecordSetDatos
        nModoEnvioEstCta = frmEnvioEstadoCta.ModoEnvioEstCta
        nDebitoMismaCta = frmEnvioEstadoCta.DebitoMismaCta
        'END JUEZ **************************************************
    
        'Realiza la Validación para el Lavado de Dinero
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        'If clsLav.EsOperacionEfectivo(Trim(nOperacion)) Then
            If Not EsExoneradaLavadoDinero() Then
                sPersLavDinero = ""
                nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
                Set clsLav = Nothing
                If nmoneda = gMonedaNacional Then
                    Dim clsTC As COMDConstSistema.NCOMTipoCambio
                    Set clsTC = New COMDConstSistema.NCOMTipoCambio
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
                    sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), False, sTipoCuenta, , , , , nmoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen, , nOperacion, nProgramAhorro) 'WIOR 20131106 nOperacion, nProgramAhorro
                    'ALPA
                    'If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                    If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                    'End
                End If
            End If
        'Else
        '    Set clsLav = Nothing
        'End If
        'WIOR 20130301 ***SEGUN TI-ERS005-2013 ************************************************************
        fbPersonaReaAhorros = False
        If loLavDinero.OrdPersLavDinero = "Exit" _
            And (nOperacion = gAhoApeEfec Or nOperacion = gAhoApeChq Or nOperacion = gAhoApeTransf _
            Or nOperacion = gPFApeEfec Or nOperacion = gPFApeChq Or nOperacion = gPFApeTransf _
            Or nOperacion = gCTSApeEfec Or nOperacion = gCTSApeChq Or nOperacion = gCTSApeTransf) Then
            
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
                frmPersRealizaOpeGeneral.Inicia Me.Caption & " (Persona " & sConPersona & ")", nOperacion
                fbPersonaReaAhorros = frmPersRealizaOpeGeneral.PersRegistrar
                
                If Not fbPersonaReaAhorros Then
                    MsgBox "Se va a proceder a Anular la Apertura de la cuenta", vbInformation, "Aviso"
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
                
                If Trim(Right(cboMoneda.Text, 5)) = 1 Then
                    lnMonto = Round(lnMonto / lnTC, 2)
                End If
            
                If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                    If lnMonto >= rsAgeParam!nMontoMin And lnMonto <= rsAgeParam!nMontoMax Then
                        frmPersRealizaOpeGeneral.Inicia Me.Caption, nOperacion
                        fbPersonaReaAhorros = frmPersRealizaOpeGeneral.PersRegistrar
                        If Not fbPersonaReaAhorros Then
                            MsgBox "Se va a proceder a Anular la Apertura de la cuenta", vbInformation, "Aviso"
                            cmdgrabar.Enabled = True
                            Exit Sub
                        End If
                    End If
                End If
                
            End If
        End If
        'WIOR FIN ***************************************************************
    
        Set clsMov = New COMNContabilidad.NCOMContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        Set clsMov = Nothing
        ' On Error GoTo ErrGraba
        
        Set rsRel = grdCliente.GetRsNew()
        Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    
        Set clsprevio = New previo.clsprevio
        Set clsLibret = New previo.clsprevio
        
        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    
        Dim lscartilla As String
        'Dim MatTitular() As String  'ARCV 12-02-2007
        '*** cargar tituales para cartilla - AVMM 11-08-2006 ***
            CargaTitulares vMatTitular
        '*******************************************************
        
        ' *** RIRO 20131102 SEGUN "CAMBIOS EN PODERES" ***
        Dim sGrupo As String
        If nProducto = gCapCTS Then
            sGrupo = "A"
        Else
            sGrupo = prepararGrupoPersona()
        End If
        ' *** FIN RIRO ***
        'EJVG20140408 ***
        If oDocRec Is Nothing Then Set oDocRec = New UDocRec
        'ReDim ArrDatos(5)
        ReDim ArrDatos(6) 'JUEZ 20160420
        ArrDatos(0) = sLpt
        ArrDatos(1) = oDocRec.fnTpoDoc
        ArrDatos(2) = oDocRec.fsNroDoc
        ArrDatos(3) = oDocRec.fsPersCod
        ArrDatos(4) = oDocRec.fsIFTpo
        ArrDatos(5) = oDocRec.fsIFCta
        'END EJVG *******
        ArrDatos(6) = fnCampanaCod 'JUEZ 20160420
                        
        Select Case nProducto
            Case gCapAhorros
             
                vnTpoProg = 0
                If cboPrograma.ListCount > 0 Then
                    vnTpoProg = Val(Right(Trim(cboPrograma.Text), 1))
                End If
                If chkPromotor.value = 1 Then
                'ALPA 20081009*******************************************************
                    sCuenta = clsCap.CapAperturaCuenta(gCapAhorros, nmoneda, rsRel, gsCodAge, nOperacion, nTasa, nMonto, gdFecSis, nFirmas, nPersoneria, _
                              IIf(nOperacion = gAhoApeTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), nTipoCuenta, sMovNro, nTipoTasa, bDocumento, sNroDoc, sCodIF, bOrdPag, , , , , , sInstitucion, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, lnMovNroTransfer, CInt(Right(Me.cboTransferMoneda.Text, 3)), gbITFAplica, IIf(Me.chkExoITF.value = 1, Right(Me.cboTipoExoneracion.Text, 3), "0"), CCur(Me.LblItf.Caption), gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 1, gITFCobroEfectivo, gITFCobroCargo), Trim(txtAlias.Text), Val(TxtMinFirmas.Text), , , vnTpoProg, txtMontoAbonar.Text, txtPlazo.Text, Trim(sCodPromotor(1)), pnMovNro, gnMovNro, _
                              , fnMovNroRVD, CCur(lblMonTra), lsPersCodConv, , prepararRegla, sGrupo, txtCuentaCargo.NroCuenta, lsBoletaCargo, gbImpTMU, gsNomCmac, gsNomAge, ArrDatos)
                              '***Parametro fnMovNroRVD, CCur(lblMonTra) agregado por ELRO el 20120717, según OYP-RFC024-2012
                              '*** RIRO 20131102, Se agregó: prepararRegla, sGrupo
                              'JUEZ 20131212 Se agregó txtCuentaCargo.NroCuenta, lsBoletaCargo, gbImpTMU, gsNomCmac, gsNomAge, sLpt
                              'EJVG20140408 Se agregó ArrDatos como ultimo parametro reemplazando sLpt
                Else
                'ALPA 20081009*******************************************************
                    sCuenta = clsCap.CapAperturaCuenta(gCapAhorros, nmoneda, rsRel, gsCodAge, nOperacion, nTasa, nMonto, gdFecSis, nFirmas, nPersoneria, _
                              IIf(nOperacion = gAhoApeTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), nTipoCuenta, sMovNro, nTipoTasa, bDocumento, sNroDoc, sCodIF, bOrdPag, , , , , , sInstitucion, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, lnMovNroTransfer, CInt(Right(Me.cboTransferMoneda.Text, 3)), gbITFAplica, IIf(Me.chkExoITF.value = 1, Right(Me.cboTipoExoneracion.Text, 3), "0"), CCur(Me.LblItf.Caption), gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 1, gITFCobroEfectivo, gITFCobroCargo), Trim(txtAlias.Text), Val(TxtMinFirmas.Text), , , vnTpoProg, txtMontoAbonar.Text, txtPlazo.Text, , pnMovNro, gnMovNro, _
                              , fnMovNroRVD, CCur(lblMonTra), lsPersCodConv, , prepararRegla, sGrupo, txtCuentaCargo.NroCuenta, lsBoletaCargo, gbImpTMU, gsNomCmac, gsNomAge, ArrDatos)
                              '***Parametro fnMovNroRVD, CCur(lblMonTra) agregado por ELRO el 20120717, según OYP-RFC024-2012
                              '*** RIRO 20131102, Se agregó: prepararRegla, sGrupo
                              'JUEZ 20131212 Se agregó txtCuentaCargo.NroCuenta, lsBoletaCargo, gbImpTMU, gsNomCmac, gsNomAge, sLpt
                              'EJVG20140408 Se agregó ArrDatos como ultimo parametro reemplazando sLpt

                End If
                '*** CADENA PARA IMPRIMIR CARTILLA - NO PROCEDE PARA MAYNAS ******
                'lscartilla = lscartilla & ImprimeCartilla(MatTitular(), 1, sCuenta, lblTasa, txtMonto.Text, gdFecSis, , , chkITFEfectivo.value, CDbl(lblITF.Caption)) & oImpresora.gPrnSaltoPagina
                'lscartilla = lscartilla & ImprimeCartilla(MatTitular(), 1, sCuenta, lblTasa, txtMonto.Text, gdFecSis, , , chkITFEfectivo.value, CDbl(lblITF.Caption)) & oImpresora.gPrnSaltoPagina
                '*****************************************************************
            Case gCapPlazoFijo
            Dim sOpeITFPlazoFijo As String
                
            vnTpoProg = 0
            If cboPrograma.ListCount > 0 Then
                vnTpoProg = Val(Right(Trim(cboPrograma.Text), 1))
            End If
            If chkITFEfectivo.value = 1 Then
                If OptAsuITF(0).value = True Then
                    sOpeITFPlazoFijo = gITFCobroEfectivoAsumidoPF
                Else
                    sOpeITFPlazoFijo = gITFCobroEfectivo
                End If
            End If
            If Trim(lblNroDoc.Caption) <> "" Then
                Dim oCapPf As COMNCaptaGenerales.NCOMCaptaDefinicion
                Set oCapPf = New COMNCaptaGenerales.NCOMCaptaDefinicion
                    lnDValoriza = oCapPf.ObtenerDiasValoriza(lblNroDoc)
                Set oCapPf = Nothing
            Else
                lnDValoriza = 0
            End If
            nPlazo = IIf(txtPlazo = "", 0, CLng(txtPlazo)) + lnDValoriza
            nPlazoVal = nPlazo
            
            '*** BRGO 20111219 **********************************************
            If cboPrograma.ListIndex = 0 Then
                nformaretiro = CLng(Trim(Right(cboFormaRetiro.Text, 4)))
            Else
                nformaretiro = gCapPFFormRetFinalPlazo
            End If
            '*** END BRGO ***************************************************
            
            'By Capi 20012008
            If chkPromotor.value = 1 Then
                'MADM 20111022 - chkDepGar - ALPA 20081009*******************************************************
                '***Modificado por ELRO el 20120206, según Acta N° 245-2011/TI-D
                'sCuenta = clsCap.CapAperturaCuenta(gCapPlazoFijo, nmoneda, rsRel, gsCodAge, nOperacion, nTasa, nMonto, gdFecSis, nFirmas, nPersoneria, _
                '          IIf(nOperacion = gPFApeTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), nTipoCuenta, sMovNro, nTipoTasa, bDocumento, sNroDoc, sCodIF, , nPlazo, nformaretiro, bCtaAboInt, sCtaAbono, , , loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, lnMovNroTransfer, CInt(Right(Me.cboTransferMoneda.Text, 3)), gbITFAplica, IIf(Me.chkExoITF.value = 1, Right(Me.cboTipoExoneracion.Text, 3), "0"), CCur(Me.lblITF.Caption), gbITFAsumidoPF, IIf(Me.chkITFEfectivo.value = 1, sOpeITFPlazoFijo, gITFCobroCargo), Trim(txtAlias.Text), val(TxtMinFirmas.Text) _
                '          , IIf(chkPermanente.Visible And chkPermanente.value = vbChecked, cPersTasaEspecial, ""), IIf(chkPermanente.Visible And chkPermanente.value = vbChecked, True, False), vnTpoProg, txtMontoAbonar.Text, txtPlazo.Text, Trim(sCodPromotor(1)), pnMovNro, gnMovNro, IIf(chkDepGar.Visible And chkDepGar.value = vbChecked, 1, -1))
                'MIOL 20121109, SEGUN RFC098-2012-A ******************************************************************
                sCuenta = clsCap.CapAperturaCuenta(gCapPlazoFijo, nmoneda, rsRel, gsCodAge, nOperacion, nTasa, nMonto, gdFecSis, nFirmas, nPersoneria, _
                          IIf(nOperacion = gPFApeTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), nTipoCuenta, sMovNro, nTipoTasa, bDocumento, sNroDoc, sCodIF, , nPlazo, nformaretiro, bCtaAboInt, sCtaAbono, , , loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, lnMovNroTransfer, CInt(Right(Me.cboTransferMoneda.Text, 3)), gbITFAplica, IIf(Me.chkExoITF.value = 1, Right(Me.cboTipoExoneracion.Text, 3), "0"), CCur(Me.LblItf.Caption), gbITFAsumidoPF, IIf(Me.chkITFEfectivo.value = 1, sOpeITFPlazoFijo, gITFCobroCargoPF), Trim(txtAlias.Text), Val(TxtMinFirmas.Text) _
                          , IIf(chkPermanente.Visible And chkPermanente.value = vbChecked, cPersTasaEspecial, ""), IIf(chkPermanente.Visible And chkPermanente.value = vbChecked, True, False), vnTpoProg, txtMontoAbonar.Text, txtPlazo.Text, Trim(sCodPromotor(1)), pnMovNro, gnMovNro, IIf(chkDepGar.Visible And chkDepGar.value = vbChecked, 1, -1), _
                          fnMovNroRVD, CCur(lblMonTra), , IIf(chkSubasta.Visible And chkSubasta.value = vbChecked, 1, -1), prepararRegla, sGrupo, txtCuentaCargo.NroCuenta, lsBoletaCargo, gbImpTMU, gsNomCmac, gsNomAge, ArrDatos)
                          '*** RIRO 20131102, Se agregó: prepararRegla, sGrupo
                          'JUEZ 20131212 Se agregó txtCuentaCargo.NroCuenta, lsBoletaCargo, gbImpTMU, gsNomCmac, gsNomAge, sLpt
                          'EJVG20140408 Se agregó ArrDatos como ultimo parametro reemplazando sLpt
                          
'                sCuenta = clsCap.CapAperturaCuenta(gCapPlazoFijo, nmoneda, rsRel, gsCodAge, nOperacion, nTasa, nMonto, gdFecSis, nFirmas, nPersoneria, _
'                          IIf(nOperacion = gPFApeTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), nTipoCuenta, sMovNro, nTipoTasa, bDocumento, sNroDoc, sCodIF, , nPlazo, nformaretiro, bCtaAboInt, sCtaAbono, , , loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, lnMovNroTransfer, CInt(Right(Me.cboTransferMoneda.Text, 3)), gbITFAplica, IIf(Me.chkExoITF.value = 1, Right(Me.cboTipoExoneracion.Text, 3), "0"), CCur(Me.lblITF.Caption), gbITFAsumidoPF, IIf(Me.chkItfEfectivo.value = 1, sOpeITFPlazoFijo, gITFCobroCargoPF), Trim(txtAlias.Text), val(txtMinFirmas.Text) _
'                          , IIf(chkPermanente.Visible And chkPermanente.value = vbChecked, cPersTasaEspecial, ""), IIf(chkPermanente.Visible And chkPermanente.value = vbChecked, True, False), vnTpoProg, txtMontoAbonar.Text, txtPlazo.Text, Trim(sCodPromotor(1)), pnMovNro, gnMovNro, IIf(chkDepGar.Visible And chkDepGar.value = vbChecked, 1, -1), _
'                          fnMovNroRVD, CCur(lblMonTra))
                'END MIOL ********************************************************************************************
                '***Fin Modificado por ELRO el 20120206*************************
                '***Parametro fnMovNroRVD, CCur(lblMonTra) agregado por ELRO el 20120717, según OYP-RFC024-2012
            
            
            Else
                'ALPA 20081009*******************************************************
                '***Modificado por ELRO el 20120206, según Acta N° 245-2011/TI-D
                'sCuenta = clsCap.CapAperturaCuenta(gCapPlazoFijo, nmoneda, rsRel, gsCodAge, nOperacion, nTasa, nMonto, gdFecSis, nFirmas, nPersoneria, _
                '          IIf(nOperacion = gPFApeTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), nTipoCuenta, sMovNro, nTipoTasa, bDocumento, sNroDoc, sCodIF, , nPlazo, nformaretiro, bCtaAboInt, sCtaAbono, , , loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, lnMovNroTransfer, CInt(Right(Me.cboTransferMoneda.Text, 3)), gbITFAplica, IIf(Me.chkExoITF.value = 1, Right(Me.cboTipoExoneracion.Text, 3), "0"), CCur(Me.lblITF.Caption), gbITFAsumidoPF, IIf(Me.chkITFEfectivo.value = 1, sOpeITFPlazoFijo, gITFCobroCargo), Trim(txtAlias.Text), val(TxtMinFirmas.Text) _
                '          , IIf(chkPermanente.Visible And chkPermanente.value = vbChecked, cPersTasaEspecial, ""), IIf(chkPermanente.Visible And chkPermanente.value = vbChecked, True, False), vnTpoProg, txtMontoAbonar.Text, txtPlazo.Text, , pnMovNro, gnMovNro, IIf(chkDepGar.Visible And chkDepGar.value = vbChecked, 1, -1))
                'MIOL 20121109, SEGUN RFC098-2012-A ******************************************************************
                sCuenta = clsCap.CapAperturaCuenta(gCapPlazoFijo, nmoneda, rsRel, gsCodAge, nOperacion, nTasa, nMonto, gdFecSis, nFirmas, nPersoneria, _
                          IIf(nOperacion = gPFApeTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), nTipoCuenta, sMovNro, nTipoTasa, bDocumento, sNroDoc, sCodIF, , nPlazo, nformaretiro, bCtaAboInt, sCtaAbono, , , loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, lnMovNroTransfer, CInt(Right(Me.cboTransferMoneda.Text, 3)), gbITFAplica, IIf(Me.chkExoITF.value = 1, Right(Me.cboTipoExoneracion.Text, 3), "0"), CCur(Me.LblItf.Caption), gbITFAsumidoPF, IIf(Me.chkITFEfectivo.value = 1, sOpeITFPlazoFijo, gITFCobroCargoPF), Trim(txtAlias.Text), Val(TxtMinFirmas.Text) _
                          , IIf(chkPermanente.Visible And chkPermanente.value = vbChecked, cPersTasaEspecial, ""), IIf(chkPermanente.Visible And chkPermanente.value = vbChecked, True, False), vnTpoProg, txtMontoAbonar.Text, txtPlazo.Text, , pnMovNro, gnMovNro, IIf(chkDepGar.Visible And chkDepGar.value = vbChecked, 1, -1), _
                          fnMovNroRVD, CCur(lblMonTra), , IIf(chkSubasta.Visible And chkSubasta.value = vbChecked, 1, -1), prepararRegla, sGrupo, txtCuentaCargo.NroCuenta, lsBoletaCargo, gbImpTMU, gsNomCmac, gsNomAge, ArrDatos)
                          '*** RIRO 20131102, Se agregó: prepararRegla, sGrupo
                          'JUEZ 20131212 Se agregó txtCuentaCargo.NroCuenta, lsBoletaCargo, gbImpTMU, gsNomCmac, gsNomAge, sLpt
                          'EJVG20140408 Se agregó ArrDatos como ultimo parametro reemplazando sLpt
                          
'                sCuenta = clsCap.CapAperturaCuenta(gCapPlazoFijo, nmoneda, rsRel, gsCodAge, nOperacion, nTasa, nMonto, gdFecSis, nFirmas, nPersoneria, _
'                          IIf(nOperacion = gPFApeTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), nTipoCuenta, sMovNro, nTipoTasa, bDocumento, sNroDoc, sCodIF, , nPlazo, nformaretiro, bCtaAboInt, sCtaAbono, , , loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, lnMovNroTransfer, CInt(Right(Me.cboTransferMoneda.Text, 3)), gbITFAplica, IIf(Me.chkExoITF.value = 1, Right(Me.cboTipoExoneracion.Text, 3), "0"), CCur(Me.lblITF.Caption), gbITFAsumidoPF, IIf(Me.chkItfEfectivo.value = 1, sOpeITFPlazoFijo, gITFCobroCargoPF), Trim(txtAlias.Text), val(txtMinFirmas.Text) _
'                          , IIf(chkPermanente.Visible And chkPermanente.value = vbChecked, cPersTasaEspecial, ""), IIf(chkPermanente.Visible And chkPermanente.value = vbChecked, True, False), vnTpoProg, txtMontoAbonar.Text, txtPlazo.Text, , pnMovNro, gnMovNro, IIf(chkDepGar.Visible And chkDepGar.value = vbChecked, 1, -1), _
'                          fnMovNroRVD, CCur(lblMonTra))
                'END MIOL ********************************************************************************************
                '***Fin Modificado por ELRO el 20120206*************************
                '***Parametro fnMovNroRVD, CCur(lblMonTra) agregado por ELRO el 20120717, según OYP-RFC024-2012
            End If
            '*** CADENA PARA IMPRIMIR CARTILLA - NO PROCEDE PARA MAYNAS ****
            'lscartilla = lscartilla & ImprimeCartilla(MatTitular(), 2, sCuenta, nTasaNominal, txtMonto.Text, gdFecSis, txtPlazo.Text, , chkITFEfectivo.value, CDbl(lblITF.Caption)) & oImpresora.gPrnSaltoPagina
            'lscartilla = lscartilla & ImprimeCartilla(MatTitular(), 2, sCuenta, nTasaNominal, txtMonto.Text, gdFecSis, txtPlazo.Text, , chkITFEfectivo.value, CDbl(lblITF.Caption)) & oImpresora.gPrnSaltoPagina
            '***************************************************************
            lnTotIntMes = 0
            If nformaretiro = gCapPFFormRetMensual Then
                
                'ALPA 20100112********************************************
                If nOperacion = gPFApeChq Or nOperacion = gPFApeLoteChq Then
                    '***Modificado por ELRO el 20110912, según Acta 245-2011/TI-D
                    'EmiteCalendarioRetiroIntPFMensual nMonto, nTasa, CInt(txtPlazo) + 4, gdFecSis, nmoneda, 0, sCuenta 'comentado por ELRO el 20110912
                    EmiteCalendarioRetiroIntPFMensual IIf(Me.chkITFEfectivo = 1, nMonto, nMonto - CDbl(Me.LblItf.Caption)), nTasa, CInt(txtPlazo) + 4, gdFecSis, nmoneda, 0, sCuenta
                    '***Fin Modificado por ELRO**********************
                Else
                    '***Modificado por ELRO el 20110912, según Acta 245-2011/TI-D
                    'EmiteCalendarioRetiroIntPFMensual nMonto, nTasa, CInt(txtPlazo), gdFecSis, nmoneda, lnDValoriza, sCuenta   'comentado por ELRO el 20110912
                    EmiteCalendarioRetiroIntPFMensual IIf(Me.chkITFEfectivo = 1, nMonto, nMonto - CDbl(Me.LblItf.Caption)), nTasa, CInt(txtPlazo), gdFecSis, nmoneda, lnDValoriza, sCuenta
                    '***Fin Modificado por ELRO**********************
                End If
                '*********************************************************
            End If
        
            Case gCapCTS
                vnTpoProg = nTpoProgramaCTS '**JUEZ 20120216
                If chkEspecial.value = vbChecked Then
                    nPorcRetCTS = CDbl(lblDisp.Caption)
                    nMonto = CDbl(txtDisp.Text) + CDbl(txtInta.Text) + CDbl(txtDU.Text)
                Else
                    nPorcRetCTS = 0 'CDbl(lblDispCTS.Caption)
                End If
                'capi1
                If chkPromotor.value = 1 Then
                    'ALPA 20081009*******************************************************
                    sCuenta = clsCap.CapAperturaCuenta(gCapCTS, nmoneda, rsRel, gsCodAge, nOperacion, nTasa, nMonto, gdFecSis, nFirmas, nPersoneria, _
                              IIf(nOperacion = gCTSApeTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), nTipoCuenta, sMovNro, nTipoTasa, bDocumento, sNroDoc, sCodIF, , , , , , nPorcRetCTS, sInstitucion, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, lnMovNroTransfer, CInt(Right(Me.cboTransferMoneda.Text, 3)), False, , , , , , , , , nTpoProgramaCTS, , , Trim(sCodPromotor(1)), pnMovNro, gnMovNro, _
                              , fnMovNroRVD, CCur(lblMonTra), , , prepararRegla, sGrupo, , , , , , ArrDatos)
                              '***Parametro fnMovNroRVD, CCur(lblMonTra) agregado por ELRO el 20120717, según OYP-RFC024-2012
                              '*** RIRO 20131102, Se agregó: prepararRegla, sGrupo

                Else
                    'ALPA 20081009*******************************************************
                    sCuenta = clsCap.CapAperturaCuenta(gCapCTS, nmoneda, rsRel, gsCodAge, nOperacion, nTasa, nMonto, gdFecSis, nFirmas, nPersoneria, _
                              IIf(nOperacion = gCTSApeTransf, Trim(txtTransferGlosa.Text), Trim(txtGlosa.Text)), nTipoCuenta, sMovNro, nTipoTasa, bDocumento, sNroDoc, sCodIF, , , , , , nPorcRetCTS, sInstitucion, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, lnMovNroTransfer, CInt(Right(Me.cboTransferMoneda.Text, 3)), False, , , , , , , , , nTpoProgramaCTS, , , , pnMovNro, gnMovNro, _
                              , fnMovNroRVD, CCur(lblMonTra), , , prepararRegla, sGrupo, , , , , , ArrDatos)
                              '***Parametro fnMovNroRVD, CCur(lblMonTra) agregado por ELRO el 20120717, según OYP-RFC024-2012
                    '********************************************************************
                    '*** RIRO 20131102, Se agregó: prepararRegla, sGrupo
                    
                End If
                '*** CADENA PARA IMPRIMIR CARTILLA - NO PROCEDE PARA MAYNAS ***
                'If txtMonto.Text <> 0 Then
                '    lscartilla = lscartilla & ImprimeCartilla(MatTitular(), 3, sCuenta, lblTasa, txtMonto.Text, gdFecSis, , , chkITFEfectivo.value, CDbl(lblITF.Caption)) & oImpresora.gPrnSaltoPagina
                '    lscartilla = lscartilla & ImprimeCartilla(MatTitular(), 3, sCuenta, lblTasa, txtMonto.Text, gdFecSis, , , chkITFEfectivo.value, CDbl(lblITF.Caption)) & oImpresora.gPrnSaltoPagina
                'Else
                '    lscartilla = lscartilla & ImprimeCartilla(MatTitular(), 3, sCuenta, lblTasa, lblTotTran.Caption, gdFecSis, , , chkITFEfectivo.value, CDbl(lblITF.Caption)) & oImpresora.gPrnSaltoPagina
                '    lscartilla = lscartilla & ImprimeCartilla(MatTitular(), 3, sCuenta, lblTasa, lblTotTran.Caption, gdFecSis, , , chkITFEfectivo.value, CDbl(lblITF.Caption)) & oImpresora.gPrnSaltoPagina
                'End If
                '**************************************************************
    End Select
    '*** BRGO 20110906 ***************************
    If gITF.gbITFAplica And CCur(Me.LblItf) > 0 Then
        Call loMov.InsertaMovRedondeoITF(sMovNro, 1, CCur(Me.LblItf) + nRedondeoITF, CCur(Me.LblItf))
    End If
    Set loMov = Nothing
    '*** BRGO
    If gnMovNro > 0 Then
        'Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen) COMENTADO X JACA 20110225
         Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110225
    End If
    
    'WIOR 20130301 ************************************************************
    If fbPersonaReaAhorros And gnMovNro > 0 Then
        frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(sCuenta), fnCondicion
        fbPersonaReaAhorros = False
    End If
    'WIOR FIN *****************************************************************
        
    Set clsCap = Nothing
    Set clsMant = Nothing

    Call frmEnvioEstadoCta.GuardarRegistroEnvioEstadoCta(1, Trim(sCuenta), rsEnvEstCta, nModoEnvioEstCta, nDebitoMismaCta, sMovNro) 'JUEZ 20130520
    'JUEZ 20150121 ************************************
    Dim rsDir As ADODB.Recordset
    nI = 0
    For nI = 1 To UBound(vMatTitular) - 1
        Set objPers = New COMDPersona.DCOMPersona
        Set rsDir = objPers.RecuperaPersonaEnvioEstadoCtaDoc(vMatTitular(nI, 2))
        Set objPers = Nothing
        vMatTitular(nI, 3) = rsDir!cPersDireccDomicilio
    Next nI
    'END JUEZ *****************************************
    '*** PEAC 20080807
    
    loVistoElectronico.RegistraVistoElectronico (pnMovNro)
    
    '*** FIN PEAC

    If chkTasaPreferencial.value = vbChecked Then
        If Trim(sCuenta) <> "" Then
            Dim oServ As COMDCaptaServicios.DCOMCaptaServicios
            Set oServ = New COMDCaptaServicios.DCOMCaptaServicios
            Dim nNumSolicitud As Long, bPermanente As Boolean
            
            
            nNumSolicitud = CLng(Me.txtNumSolicitud.Text)
            
            bPermanente = IIf(chkPermanente.value = vbUnchecked, False, True)
            'sPersCod, nProducto, nMoneda
            
            oServ.AgregaCapTasaEspecial nNumSolicitud, vSperscod, nProducto, nmoneda, 2, sMovNro, Val(lblTasa.Caption), "APERTURA DE CUENTA CON BTASA PREFERENCIAL", nMonto, sCuenta, nPlazo, , bPermanente
            
            
            Set oServ = Nothing
        End If
    End If
    'JACA 20110512***********************************************************
                
        'Dim objPersona As COMDPersona.DCOMPersonas
        Dim rsPersOcu As Recordset
        Dim nAcumulado As Currency
        Dim nMontoPersOcupacion As Currency
        
        Set rsPersOcu = New Recordset
        'Set objPersona = New COMDPersona.DCOMPersonas
        
        Set clsTC = New COMDConstSistema.NCOMTipoCambio
        nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
        Set clsTC = Nothing
                
        For i = 1 To grdCliente.Rows - 1
            If CLng(Trim(Right(Me.grdCliente.TextMatrix(i, 3), 4))) = 10 Then
                Set rsPersOcu = objPersona.ObtenerDatosPersona(Me.grdCliente.TextMatrix(i, 1))
               Exit For
            End If
        Next i
        
        nAcumulado = objPersona.ObtenerPersAcumuladoMontoOpe(nTC, Mid(Format(gdFecSis, "yyyymmdd"), 1, 6), rsPersOcu!cPersCod)
        'nMontoPersOcupacion = objPersona.ObtenerParamPersAgeOcupacionMonto(Mid(rsPersOcu!cPersCod, 4, 2), CInt(Mid(rsPersOcu!cPersCIIU, 2, 2)))  'RIRO20140911 INC1409110010 Comentado
        nMontoPersOcupacion = objPersona.ObtenerParamPersAgeOcupacionMonto(Mid(rsPersOcu!cPersCod, 4, 2), Val(Mid(IIf(IsNull(rsPersOcu!cPersCIIU), "", rsPersOcu!cPersCIIU), 2, 2))) 'RIRO20140911 INC1409110010 ADD
    
        If nAcumulado >= nMontoPersOcupacion Then
            If Not objPersona.ObtenerPersonaAgeOcupDatos_Verificar(rsPersOcu!cPersCod, gdFecSis) Then
                objPersona.insertarPersonaAgeOcupacionDatos gnMovNro, rsPersOcu!cPersCod, IIf(nmoneda = gMonedaNacional, Me.lblTotal, Me.lblTotal / nTC), nAcumulado, gdFecSis, sMovNro
            End If
        End If
       
        Set objPersona = Nothing
    'JACA END*****************************************************************
    'FRHU 20140926 ERS099-2014: SE QUITO EL REGISTRO DE FIRMAS Y LA LIBRETA DE AHORROS: PARA PLAZO FIJO SE QUITO EL CERTIFICADO DE PLAZO FIJO (RESUMEN DE DEPOSITO)
    'By Capi Acta 014-2007
'    If nProducto <> gCapAhorros Then
'        Dim rsCta As New ADODB.Recordset
'        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
'        Set rsCta = clsMant.GetProductoPersona(sCuenta)
'        If Not (rsCta.EOF And rsCta.BOF) Then
'            Set grdCliente.Recordset = rsCta
'        End If
'        Set rsRel = grdCliente.GetRsNew()
'        Set rsCta = Nothing
'        grdCliente.ColWidth(5) = 0
'    End If
        
    
'    If lbImpRegFirma = 1 Then
'        MsgBox "Coloque papel para el registro de firmas", vbInformation, "Aviso"
'        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
'        clsMant.IniciaImpresora gImpresora
'        lsCadImp = clsMant.GeneraRegistroFirmas(sCuenta, Trim(Left(cboTipoCuenta, 25)), gdFecSis, bOrdPag, rsRel, gsNomAge, gdFecSis, gsCodUser, vnTpoProg)
'        'ALPA 20100202*******************************
'        'clsprevio.Show lscadimp, "Registro Firmas", True
'        clsprevio.Show lsCadImp, "Registro Firmas", True, , gImpresora
'
'        'By Capi Acta 014-2007 Impresion Libretas de Ahorro
'            Dim lsCad As String
'            MsgBox "Coloque Libreta para Impresion", vbInformation, "Aviso"
'            Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
'            clsMant.IniciaImpresora gImpresora
'            lsCad = clsMant.GeneraRegistroLibretas(sCuenta, Trim(Left(cboTipoCuenta, 25)), gdFecSis, bOrdPag, rsRel, gsNomAge, gdFecSis, gsCodUser, vnTpoProg)
'            'ALPA 20100202
'            'clsLibret.Show lsCad, "Registro Libretas", True
'            clsLibret.Show lsCad, "Registro Libretas", True, , gImpresora
'            Set clsMant = Nothing
'            Set clsprevio = Nothing
'        'End By
'
'
'    End If
    
'    If nProducto = gCapPlazoFijo Then
'        MsgBox "Coloque papel para el Certificado de Plazo Fijo", vbInformation, "Aviso"
'        '***Modificado por ELRO el 20120124, según Acta N° 006-2012/TI-D
'        'EmiteCertificadoPlazoFijo sCuenta, rsRel
'        EmiteCertificadoPlazoFijo sCuenta, rsRel, , nformaretiro
'        '***Fin Modificado por ELRO*************************************
'    End If

    MsgBox "Coloque papel para la Solicitud de Apertura", vbInformation, "Aviso"
    Set clsprevio = New previo.clsprevio
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    clsMant.IniciaImpresora gImpresora
    If nPersoneria <> 1 Then
        lsCadImp = clsMant.GeneraSolicitudAperturaPersJuridica(sCuenta, Trim(Left(cboTipoCuenta, 25)), gdFecSis, bOrdPag, rsRel, gsNomAge, gdFecSis, gsCodUser, vnTpoProg, , prepararRegla)
    Else
        lsCadImp = clsMant.GeneraSolicitudAperturaPersNatural(sCuenta, Trim(Left(cboTipoCuenta, 25)), gdFecSis, bOrdPag, rsRel, gsNomAge, gdFecSis, gsCodUser, vnTpoProg, , prepararRegla)
    End If
    clsprevio.Show lsCadImp, "Solicitud de Apertura", True, , gImpresora
    Set clsMant = Nothing
    Set clsprevio = Nothing
    'FIN FRHU 20140926 ERS099-2014
    
    On Error GoTo ErrImp
    
   
   
    If sPersLavDinero <> "" Then
       'By Capi 28022008 el mensaje debe estar dentro
       MsgBox "Coloque papel para la Boleta de Lavado de Dinero", vbInformation, "Aviso"
       'Call loLavDinero.imprimirBoletaREU(sCuenta, Mid(sCuenta, 9, 1), loLavDinero.OrigenPersLavDinero) 'COMENTADO X JACA 20110302
       Call loLavDinero.imprimirBoletaREU(sCuenta, Mid(sCuenta, 9, 1), loLavDinero.OrigenPersLavDinero, loLavDinero.NroREU) 'JACA 20110302
    End If

    
    'impresion de cartillas
    MsgBox "Coloque papel para imprimir Cartilla", vbInformation, "Aviso"
    

    '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
    ' ---- Funcion en Word para imprimir Cartillas --- avmm -- 06-02-2007
    'If nOperacion = gAhoApeChq Or nOperacion = gAhoApeEfec Then
    If nOperacion = gAhoApeChq Or nOperacion = gAhoApeEfec Or nOperacion = gAhoApeTransf Or nOperacion = gAhoApeCargoCta Then 'JUEZ 20131212 gAhoApeCargoCta
    '***Fin Agregado por ELRO el 20120717*******************
        'by capi 20082008 se agrego en condicion para cuenta soñada
        'If Trim(Right(cboPrograma.Text, 1)) = 0 Or Trim(Right(cboPrograma.Text, 1)) = 2 Then
        
        'By capi 03112008 porque ahora panderito tiene su propia cartilla
        'If Trim(Right(cboPrograma.Text, 1)) = 0 Or Trim(Right(cboPrograma.Text, 1)) = 2 Or Trim(Right(cboPrograma.Text, 1)) = 5 Then
        If Trim(Right(cboPrograma.Text, 1)) = 0 Or Trim(Right(cboPrograma.Text, 1)) = 5 Or Trim(Right(cboPrograma.Text, 1)) = 6 Or Trim(Right(cboPrograma.Text, 1)) = 8 Then
        '***Condición Trim(Right(cboPrograma.Text, 1)) = 8  agregada por ELRO el 20130130, según TI-ERS020-2013
            If chkOrdenPago.value = 0 Then
                '***Modificado por ELRO el 20111025, según Acta 245-2011/TI-D
                'ImpreCartillaAhoCorriente vMatTitular, sCuenta, lblTasa, txtMonto.Text, Trim(Right(cboPrograma.Text, 1))
                '***Modificado por ELRO el 20130130, según TI-ERS020-2013
                'ImpreCartillaAhoCorriente vMatTitular, sCuenta, LblTasa, IIf(Me.chkITFEfectivo = 1, nMonto, nMonto - CDbl(Me.LblItf.Caption)), Trim(Right(cboPrograma.Text, 1))
                If Trim(Right(cboPrograma.Text, 1)) = 8 Then
                    If Trim(cboInstConvDep.Text) <> "" Then
                        ImpreCartillaAhoCorriente vMatTitular, sCuenta, lblTasa, IIf(Me.chkITFEfectivo = 1, nMonto, nMonto - CDbl(Me.LblItf.Caption)), Trim(Right(cboPrograma.Text, 1)), , , , Left(cboInstConvDep.Text, Len(cboInstConvDep.Text) - 20)
                    Else
                        ImpreCartillaAhoCorriente vMatTitular, sCuenta, lblTasa, IIf(Me.chkITFEfectivo = 1, nMonto, nMonto - CDbl(Me.LblItf.Caption)), Trim(Right(cboPrograma.Text, 1))
                    End If
                Else
                    ImpreCartillaAhoCorriente vMatTitular, sCuenta, lblTasa, IIf(Me.chkITFEfectivo = 1, nMonto, nMonto - CDbl(Me.LblItf.Caption)), Trim(Right(cboPrograma.Text, 1))
                End If
                '***Fin Modificado por ELRO el 20130130******************
                '***Fin Modificado por ELRO**********************************
            Else
                '***Modificado por ELRO el 20111025, según Acta 245-2011/TI-D
                'ImpreCartillaAhoCorrienteOP vMatTitular, sCuenta, lblTasa, txtMonto.Text, nPersoneria
                ImpreCartillaAhoCorrienteOP vMatTitular, sCuenta, lblTasa, IIf(Me.chkITFEfectivo = 1, nMonto, nMonto - CDbl(Me.LblItf.Caption)), nPersoneria
                '***Fin Modificado por ELRO**********************************
            End If
        'By Capi 03112008 para ahorro panderito
        ElseIf Trim(Right(cboPrograma.Text, 1)) = 2 Or Trim(Right(cboPrograma.Text, 1)) = 7 Then
            '***Modificado por ELRO el 20111025, según Acta 245-2011/TI-D
            'ImpreCartillaAhoCorriente vMatTitular, sCuenta, lblTasa, txtMonto.Text, Trim(Right(cboPrograma.Text, 1)), val(txtPlazo.Text), gdFecSis
            ImpreCartillaAhoCorriente vMatTitular, sCuenta, lblTasa, IIf(Me.chkITFEfectivo = 1, nMonto, nMonto - CDbl(Me.LblItf.Caption)), Trim(Right(cboPrograma.Text, 1)), Val(txtPlazo.Text), gdFecSis
            '***Fin Modificado por ELRO**********************************
        '
        ElseIf Trim(Right(cboPrograma.Text, 1)) = 1 Then
            'ALPA 20091119**********************************************
            'ImpreCartillaAhoNanito vMatTitular, sCuenta, lblTasa, txtMonto.Text
            '***Modificado por ELRO el 20111025, según Acta 245-2011/TI-D
            'ImpreCartillaAhoNanito vMatTitular, sCuenta, lblTasa, txtMonto.Text, val(txtPlazo.Text)
            ImpreCartillaAhoNanito vMatTitular, sCuenta, lblTasa, IIf(Me.chkITFEfectivo = 1, nMonto, nMonto - CDbl(Me.LblItf.Caption)), Val(txtPlazo.Text)
            '***Fin Modificado por ELRO**********************************
            '***********************************************************
        ElseIf Trim(Right(cboPrograma.Text, 1)) = 3 Or Trim(Right(cboPrograma.Text, 1)) = 4 Then
            '***Modificado por ELRO el 20111025, según Acta 245-2011/TI-D
            'ImpreCartillaAhoPandero vMatTitular, sCuenta, lblTasa, txtMonto.Text, gdFecSis, txtMontoAbonar, txtPlazo, Trim(Right(cboPrograma.Text, 1)), Trim(lblInstitucion.Caption)
            ImpreCartillaAhoPandero vMatTitular, sCuenta, lblTasa, IIf(Me.chkITFEfectivo = 1, nMonto, nMonto - CDbl(Me.LblItf.Caption)), gdFecSis, txtMontoAbonar, txtPlazo, Trim(Right(cboPrograma.Text, 1)), Trim(lblInstitucion.Caption)
            '***Fin Modificado por ELRO**********************************
        End If
    '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
    'ElseIf nOperacion = gPFApeChq Or nOperacion = gPFApeEfec Then
    ElseIf nOperacion = gPFApeChq Or nOperacion = gPFApeEfec Or nOperacion = gPFApeTransf Or nOperacion = gPFApeCargoCta Then 'JUEZ 20131212 gPFApeCargoCta
    '***Fin Agregado por ELRO el 20120717*******************
        '*** Modificado por BRGO **************************************
        If Trim(Right(cboPrograma.Text, 1)) = 2 Then
            ImpreCartillaAhoCorriente vMatTitular, sCuenta, lblTasa, IIf(Me.chkITFEfectivo = 1, nMonto, nMonto - CDbl(Me.LblItf.Caption)), Trim(Right(cboPrograma.Text, 1)), Val(txtPlazo.Text), gdFecSis
        ElseIf Trim(Right(cboPrograma.Text, 1)) = 3 Then
            ImpreCartillaAhoPandero vMatTitular, sCuenta, lblTasa, IIf(Me.chkITFEfectivo = 1, nMonto, nMonto - CDbl(Me.LblItf.Caption)), gdFecSis, txtMontoAbonar, txtPlazo, Trim(Right(cboPrograma.Text, 1)), Trim(lblInstitucion.Caption)
        Else
            '***Modificado por ELRO el 20111025, según Acta 245-2011/TI-D
            'ImpreCartillaAhoCorriente vMatTitular, sCuenta, lblTasa, txtMonto.Text, Trim(Right(cboPrograma.Text, 1)), val(txtPlazo.Text), gdFecSis
            ImpreCartillaPlazoFijo vMatTitular, sCuenta, lblTasa, IIf(Me.chkITFEfectivo = 1, nMonto, nMonto - CDbl(Me.LblItf.Caption)), nPlazo, gdFecSis, nformaretiro, lnTotIntMes, Trim(Right(cboPrograma.Text, 1))
            '***Fin Modificado por ELRO**********************************
        End If
    '***Modificado por ELRO el 20120717, según OYP-RFC024-2012
    'ElseIf nOperacion = gCTSApeChq Or nOperacion = gCTSApeEfec Then
    ElseIf nOperacion = gCTSApeChq Or nOperacion = gCTSApeEfec Or nOperacion = gCTSApeTransf Then
    '***Fin Modificado por ELRO el 20120717*******************
        ImpreCartillaCTS vMatTitular, sCuenta, lblTasa, gdFecSis, txtMonto.Text
    End If
    '----------------------------------------------------------------------
'    lbOk = True
'    Do While lbOk
'          nFicSal = FreeFile
'          Open sLpt For Output As nFicSal
'              Print #nFicSal, lscartilla
'              Print #nFicSal, ""
'          Close #nFicSal
'          If MsgBox("Desea Reimprimir Cartilla ??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
'              lbOk = False
'          End If
'    Loop
    
    'impresion de boletas
    MsgBox "Coloque papel para la Boleta", vbInformation, "Aviso"
    If bDocumento And nDocumento = TpoDocCheque Then
        EmiteBoleta sCuenta, 0, nMonto, "" 'JUEZ 20131212
    Else
        EmiteBoleta sCuenta, nMonto, nMonto, lsBoletaCargo 'JUEZ 20131212
    End If
        
    '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
    If nOperacion = gAhoApeTransf Or nOperacion = gPFApeTransf Or nOperacion = gCTSApeTransf Then
        If Trim(Right(cboMoneda, 3)) <> Trim(Right(cboTransferMoneda, 3)) Then
          MsgBox "Coloque papel para la Boleta de Compra/Venta Moneda Extranjera.", vbInformation, "Aviso"
          lsBoletaCVME = oNCOMContImprimir.ImprimeBoletaCompraVentaME("Compra/Venta Moneda Extranjera", "", _
                                                                      lsPersNombreCVME, _
                                                                      lsPersDireccionCVME, _
                                                                      lsdocumentoCVME, _
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
  End If
  '***Fin Agregado por ELRO el 20120717*******************
       
    rsRel.Close
    Set rsRel = Nothing
    lnTitularPJ = 0
    txtCuenta.Cuenta = Right(sCuenta, 10)
        
    MsgBox "Cuenta Generada : " & sCuenta, vbInformation, "Nueva Cuenta"
    
    'vapi segun ERS082-2014 Nota: Abre el formulario de entrega de Merchandising
    If nOperacion = "200101" Or nOperacion = "200103" Or nOperacion = "210101" Or nOperacion = "210102" Or nOperacion = "210103" Then
       Call frmMkEntregaCombo.Inicio(sCuenta, nOperacion, False, nmoneda, nMonto)
    End If
    '*****************************FIN VAPI***************************************
    
    gVarPublicas.LimpiaVarLavDinero
     
    Set clsLav = Nothing
    Set loLavDinero = Nothing
    cmdCancelar_Click
    grdCliente.ColWidth(5) = 0
    '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
    fnMovNroRVD = 0
    lblMonTra = "0.00"
    '***Fin Agregado por ELRO el 20120717*******************
        
    ' *** RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
    If intPunteroPJ_NA = 1 Then
        Sleep (1000)
        Set clsMov = New COMNContabilidad.NCOMContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set clsMov = Nothing
        
        Dim clsBloq As COMNCaptaGenerales.NCOMCaptaGenerales
        Set clsBloq = New COMNCaptaGenerales.NCOMCaptaGenerales
        MsgBox "Esta cuenta será bloqueada hasta que se de mantenimiento a sus poderes", vbInformation, "Aviso"
        Call clsBloq.BloqueCuentaTotal(sCuenta, gCapMotBlqTotFaltanFirmas, "BLOQUEO EN APERTURA DE CUENTA", sMovNro)
        Set clsBloq = Nothing
    End If
    ' *** Fin RIRO
    
End If
Set oNCOMContImprimir = Nothing '***Agregado por ELRO el 20120717, según OYP-RFC024-2012

Exit Sub
ErrGraba:
    MsgBox err.Description, vbExclamation, "Error"
    Exit Sub
ErrImp:
    MsgBox err.Description, vbExclamation, "Error de Impresion"
    cmdCancelar_Click
End Sub

Private Sub cmdQuitarRega_Click()

    If grdReglas.Rows = 2 And Trim(grdReglas.TextMatrix(1, 1)) = "" Then
        Exit Sub
    End If
    
    If MsgBox("¿¿Está seguro de eliminar la regla creada??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        grdReglas.EliminaFila grdReglas.row
    End If
        
    seleccionarTipoCuentaXregla
    
End Sub

Private Sub cmdsalir_Click()
'ARCV 13-02-2007
If vbDesembolso Then
Dim rsRel As ADODB.Recordset
    
    'FRHU 20140228 RQ14006
    'VALIDACION DE NUEVOS CAMPOS
    'If cboPromotor.Text = "" Then
    '    MsgBox "Debe seleccionar un promotor.", vbInformation, "Aviso"
    '    cboPromotor.SetFocus
    '    Exit Sub
    'End If
    'FIN FRHU 20140228
    
    Set rsRel = grdCliente.GetRsNew
    Set vMatRela = rsRel
    Call CargaTitulares(vMatTitular)
    vnPrograma = CInt(Trim(Right(cboPrograma.Text, 1)))
    vnMontoAbonar = CDbl(txtMontoAbonar.Text)
    vnPlazoAbono = CInt(txtPlazo.Text)
    vsPromotor = Right(Trim(cboPromotor.Text), 13)
    lnTitularPJ = 0
End If



Unload Me
End Sub

'EJVG20130912 ***
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
'    '***Agregado por ELRO el 20120706, según OYP-RFC024-2012
'    If Trim(grdCliente.TextMatrix(1, 1)) = "" Then
'        MsgBox "Debe ingresar el(los) Titular(es) de la Cuenta.", vbOKOnly + vbInformation, "AVISO"
'        Exit Sub
'    End If
'
'    If gsOpeCod = gAhoApeTransf Then
'        lnTipMot = 1
'    ElseIf gsOpeCod = gPFApeTransf Then
'        lnTipMot = 3
'    ElseIf gsOpeCod = gCTSApeTransf Then
'        lnTipMot = 5
'    End If
'    '***Fin Agregado por ELRO*******************************
'
'    '***Modificado por ELRO el 20120706, según OYP-RFC024-2012
'    'lnMovNroTransfer = frmTransfpendientes.Ini(Right(Me.cboTransferMoneda.Text, 2), lnTransferSaldo, lsGlosa, lsInstit, lsDoc)
'    oform.iniciarFormulario Trim(Right(cboTransferMoneda, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer
'    '***Fin Modificado por ELRO el 20120706*******************
'
'    '***Agregado por ELRO el 20120706, según OYP-RFC024-2012
'    For i = 1 To grdCliente.Rows - 1
'         If fsPersCodTransfer = grdCliente.TextMatrix(i, 1) Then
'            Exit For
'         End If
'
'         If i = grdCliente.Rows - 1 Then
'            MsgBox "Debe ingresar el Tilular del Voucher.", vbOKOnly + vbInformation, "AVISO"
'            Exit Sub
'         End If
'    Next i
'    '***Fin Agregado por ELRO el 20120706*******************
'
'    '***Comentado por ELRO el 20120706, según OYP-RFC024-2012
'    'If lnMovNroTransfer = -1 Then
'    '    Me.cboTransferMoneda.Enabled = True
'    '    lnTransferSaldo = 0
'    'Else
'    '    Me.cboTransferMoneda.Enabled = False
'    'End If
'    '***Fin Comentado por ELRO*******************************
'
'    Me.txtTransferGlosa.Text = lsGlosa
'    Me.lbltransferBco.Caption = lsInstit
'    Me.lblTrasferND.Caption = lsDoc
'    'sNroDoc = lsDoc
'
'    'Me.txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")'***Comentado por ELRO el 20120706, según OYP-RFC024-2012
'
'    If Right(cboMoneda, 3) = Moneda.gMonedaNacional Then
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
'    If txtCuenta.Prod = "234" Then
'        vnMontoDOC = CDbl(txtMonto.Text)
'        lblTotTran.Caption = vnMontoDOC
'    End If
'
'    txtMonto_Change '***Agregado por ELRO el 20120726, según OYP-RFC024-2012
'
'   'Me.LblTotal.Caption = Format(txtMonto.value + CCur(Me.LblItf.Caption), "#,##0.00")
'
'    If lnMovNroTransfer <> -1 Then
'        Me.txtTransferGlosa.SetFocus
'    End If
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
'
'End Sub
Private Sub cmdTranfer_Click()
    Dim lsGlosa As String
    Dim lsDoc As String
    Dim lsInstit As String
    Dim oform As frmCapRegVouDepBus
    Dim lnTipMot As Integer
    Dim lsDetalle As String

    On Error GoTo ErrTransfer
    If cboTransferMoneda.Text = "" Then
        MsgBox "Debe escoger la moneda de la transferencia.", vbInformation, "Aviso"
        If cboTransferMoneda.Visible And cboTransferMoneda.Enabled Then cboTransferMoneda.SetFocus
        Exit Sub
    End If
       
    If gsOpeCod = gAhoApeTransf Then
        lnTipMot = 1
    ElseIf gsOpeCod = gPFApeTransf Then
        lnTipMot = 3
    ElseIf gsOpeCod = gCTSApeTransf Then
        lnTipMot = 5
    End If
    
    fnMovNroRVD = 0
    Set oform = New frmCapRegVouDepBus
    sinReglas 'EJVG20140408
    SetDatosTransferencia "", "", "", 0, -1, "" 'Limpiamos datos y variables globales
    oform.iniciarFormulario Trim(Right(cboTransferMoneda, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer, lsDetalle
    If fnMovNroRVD = 0 Then
        Exit Sub
    End If
    SetDatosTransferencia lsGlosa, lsInstit, lsDoc, lnTransferSaldo, lnMovNroTransfer, lsDetalle
    grdCliente.row = 1
    grdCliente.Col = 3
    grdCliente_OnEnterTextBuscar grdCliente.TextMatrix(1, 1), 1, 1, False
    Set oform = Nothing
    Exit Sub
ErrTransfer:
    MsgBox "Ha sucedido un error al cargar los datos de la Transferencia", vbCritical, "Aviso"
End Sub
Private Sub SetDatosTransferencia(ByVal psGlosa As String, ByVal psInstit As String, ByVal psDoc As String, ByVal pnTransferSaldo As Currency, ByVal pnMovNroTransfer As Long, ByVal psDetalle As String)
    Dim oPersona As New DCOMPersona
    Dim rsPersona As New ADODB.Recordset
    Dim row As Integer
    
    txtTransferGlosa.Text = psGlosa
    lbltransferBco.Caption = psInstit
    lblTrasferND.Caption = psDoc
    
    LimpiaFlex grdCliente
    If psDetalle <> "" Then
        Set rsPersona = oPersona.RecuperaPersonaxCapRegVouDep(psDetalle)
        Do While Not rsPersona.EOF
            grdCliente.AdicionaFila
            row = grdCliente.row
            grdCliente.TextMatrix(row, 1) = rsPersona!cPersCod
            grdCliente.TextMatrix(row, 2) = rsPersona!cPersNombre
            grdCliente.TextMatrix(row, 4) = rsPersona!nPersPersoneria
            rsPersona.MoveNext
        Loop
    End If
    
    If Right(cboMoneda, 3) = Moneda.gMonedaNacional Then
        If Right(cboTransferMoneda, 3) = Moneda.gMonedaNacional Then
            txtMonto.Text = Format(pnTransferSaldo, "#,##0.00")
        Else
            txtMonto.Text = Format(pnTransferSaldo * CCur(lblTTCCD.Caption), "#,##0.00")
        End If
    Else
        If Right(cboTransferMoneda, 3) = Moneda.gMonedaNacional Then
            txtMonto.Text = Format(pnTransferSaldo / CCur(lblTTCVD.Caption), "#,##0.00")
        Else
            txtMonto.Text = Format(pnTransferSaldo, "#,##0.00")
        End If
    End If
    
    If txtCuenta.Prod = "234" Then
        vnMontoDOC = CDbl(txtMonto.Text)
        lblTotTran.Caption = vnMontoDOC
    End If
    
    txtMonto_Change

    If pnMovNroTransfer <> -1 Then
        txtTransferGlosa.SetFocus
    End If
    
    txtTransferGlosa.Locked = True
    txtMonto.Enabled = False
    lblMonTra = Format(pnTransferSaldo, "#,##0.00")
    
    Set rsPersona = Nothing
    Set oPersona = Nothing
End Sub
'END EJVG *******

Private Sub Form_Activate()
  If txtCuenta.Prod = "234" And nOperacion = gCTSApeChq And lblNroDoc.Caption <> "" Then
        vnMontoDOC = CDbl(txtMonto.Text)
        lblTotTran.Caption = vnMontoDOC
   End If
End Sub

Private Sub Form_Load()

    Dim objCapta As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set objCapta = New COMNCaptaGenerales.NCOMCaptaDefinicion

    nTitular = 0
    nClientes = 0
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    GetTipCambio gdFecSis
    '***Modificado por ELRO el 20121828, según OYP-RFC024-2012
    'lblTTCCD.Caption = Format(gnTipCambioC, "#.00")
    lblTTCCD.Caption = Format(gnTipCambioC, "#,#0.0000")
    'lblTTCVD.Caption = Format(gnTipCambioV, "#.00")
    lblTTCVD.Caption = Format(gnTipCambioV, "#,#0.0000")
    '***Fin Modificado por ELRO el 20121828******************
    lnMovNroTransfer = -1
    
    ' *** Agregado por RIRO el 20130501, Proyecto Ahorro - Poderes ***
    'grdCliente.CeldaPegar (False)
    grdCliente.VisiblePopMenu = False
    fraReglasPorderes.Left = 6720
    fraReglasPorderes.Top = 150
    fraReglasPorderes.Height = 2055
    fraReglasPorderes.Width = 2460
    chkPromotor.Enabled = False
    chkEspecial.Enabled = False
    ' *** Fin RIRO ***
    
    Dim oCons As COMDConstantes.DCOMConstantes
    Set oCons = New COMDConstantes.DCOMConstantes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Set rs = oCons.GetConstante(1044, , , True, , "0','1044")
    Me.cboTipoExoneracion.Clear
    While Not rs.EOF
        cboTipoExoneracion.AddItem rs.Fields(1) & Space(100) & rs.Fields(0)
        rs.MoveNext
    Wend
    Set oCons = Nothing
    rs.Close
    Set rs = Nothing
    
    'Cambio GRVA Promotores
    Dim oCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim rsPromotor As ADODB.Recordset
    
    Set oCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsPromotor = New ADODB.Recordset
    
    cboPromotor.Clear
    Set rsPromotor = oCapGen.GetPromotores(gsCodAge)
    While Not rsPromotor.EOF
        cboPromotor.AddItem rsPromotor.Fields("cUser") & Space(3) & rsPromotor.Fields("cPersNombre") & Space(100) & "|" & rsPromotor.Fields("cPersCod")
        rsPromotor.MoveNext
    Wend
    rsPromotor.Close
    Set rsPromotor = Nothing
    
    'Agregado por RIRO 20130501 Poderes en Aperturas
    Dim x As Integer
    For x = 65 To 90 Step 1
       lsLetras.AddItem UCase(Chr(x))
    Next x
    sinReglas
    'Fin RIRO
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oDocRec = Nothing 'EJVG20140408
'***Agregado por ELRO el 20120823, según OYP-RFC024-2012
If nOperacion = gAhoApeTransf Or nOperacion = gPFApeTransf Or nOperacion = gCTSApeTransf Then
    LimpiaControles
End If
'***Agregado por ELRO el 20120823***********************
End Sub

Private Sub grdCliente_KeyPress(KeyAscii As Integer)
    
    If grdCliente.Col = 6 Then
    
    ' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
    
            'If UCase(Chr(KeyAscii)) = "S" Or UCase(Chr(KeyAscii)) = "SI" Then
            '    grdCliente.TextMatrix(grdCliente.row, 6) = "SI"
            'ElseIf UCase(Chr(KeyAscii)) = "N" Or UCase(Chr(KeyAscii)) = "NO" Then
            '    grdCliente.TextMatrix(grdCliente.row, 6) = "NO"
            'ElseIf UCase(Chr(KeyAscii)) = "O" Or UCase(Chr(KeyAscii)) = "OPCIONAL" Then
            '    grdCliente.TextMatrix(grdCliente.row, 6) = "OPCIONAL"
            'Else
            '    MsgBox "PRESIONE LA TECLA S SI LA FIRMA DE ESTE CLIENTE ES OBLIGATORIA." & vbCrLf & "PRESIONE LA TECLA N SI LA FIRMA DE ESTE CLIENTE NO ES OBLIGATORIA." & vbCrLf & "PRESIONE LA TECLA O SI LA FIRMA DE ESTE CLIENTE  ES OPCIONAL.", vbOKOnly + vbExclamation, App.Title
            'End If
    
    ' *** FIN RIRO
    
    ElseIf grdCliente.Col = 9 Then
    
        Dim i, nContar, nContarGrupos As Integer
        Dim sLetra As String
        Dim sReglas() As String
        Dim sD As Variant
        
        ' Bloqueando columnas no editables
        Dim sColumnas() As String
        sColumnas = Split(grdCliente.ColumnasAEditar, "-")
        
        If sColumnas(grdCliente.Col) = "X" Or _
           Val(Trim(grdCliente.TextMatrix(grdCliente.row, 4))) > 1 Or _
           Val(Trim(Right(grdCliente.TextMatrix(grdCliente.row, 3), 2))) = 11 Then
            Exit Sub
        End If
        ' Fin Bloqueo
        
        If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then
            grdCliente.TextMatrix(grdCliente.row, 9) = UCase(Chr(KeyAscii))
        Else
            grdCliente.TextMatrix(grdCliente.row, 9) = ""
        End If
        seleccionarTipoCuentaXregla
    End If

End Sub

Private Sub grdCliente_LostFocus()
        
    'Se debe considerar que en la apertuba, debe intervenir solo una persona juridica
    Dim nRelacion, nPJuridica, i, J As Integer
    Dim sMensaje As String
    For i = 1 To grdCliente.Rows - 1
        If Val(Trim(grdCliente.TextMatrix(i, 4))) > 1 Then
            If Trim(Left(grdCliente.TextMatrix(i, 3), 10)) <> "TITULAR" Then
                grdCliente.TextMatrix(i, 3) = ""
                nPJuridica = nPJuridica + 1
            End If
            For J = 1 To grdCliente.Rows - 1
                If Trim(Left(grdCliente.TextMatrix(J, 3), 10)) = "TITULAR" And Val(Trim(grdCliente.TextMatrix(J, 4))) = 1 Then
                    grdCliente.TextMatrix(J, 3) = ""
                    nRelacion = nRelacion + 1
                End If
            Next
            If nPJuridica > 0 Then
                sMensaje = "* La relacion de una persona jurídica debe ser: Titular" & vbNewLine
            End If
            If nRelacion > 0 Then
                sMensaje = sMensaje & "* Solo la persona jurídica debe ser el titular de la cuenta"
            End If
            If nPJuridica > 0 Or nRelacion > 0 Then
                MsgBox "Observaciones: " & vbNewLine & vbNewLine & sMensaje, vbExclamation, "Aviso"
                Unload frmBuscaPersona
                grdCliente.SetFocus
                Exit Sub
            End If
        End If
    Next
    
    'JUEZ 20140414 ******************************************
    If nOperacion = gAhoApeEfec Or nOperacion = gAhoApeChq Or nOperacion = gPFApeEfec Or nOperacion = gPFApeChq Then
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
    'END JUEZ ***********************************************
    
    ' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
      
            'Dim i As Integer, numfirmas As Integer
            'numfirmas = 0
            'For i = 1 To grdCliente.Rows - 1
            '    If grdCliente.TextMatrix(i, 6) = "SI" Then
            '        numfirmas = numfirmas + 1
            '    End If
            'Next i
            '
            'Label18.Tag = numfirmas
            '
            'TxtMinFirmas.Text = CStr(numfirmas)
            ''TXTALIAS.SetFocus
            
    ' *** FIN RIRO
    
End Sub

Private Sub grdCliente_OnCellChange(pnRow As Long, pnCol As Long)

If grdCliente.TextMatrix(pnRow, 1) = "" Then Exit Sub

If pnCol = 3 And grdCliente.TextMatrix(pnRow, pnCol) <> "" Then
    
    ' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
    
            '    Dim nPers As PersPersoneria
            '    Dim nRelacion As CaptacRelacPersona
            '
            ''    Dim clsGen As DGeneral
            ''    Dim rsRel As Recordset
            ''    Set clsGen = New DGeneral
            ''    Set rsRel = clsGen.GetConstante(gCaptacRelacPersona)
            ''    grdCliente.CargaCombo rsRel
            ''    rsRel.Close
            ''    Set rsRel = Nothing
            '
            '    nPers = CLng(grdCliente.TextMatrix(pnRow, 4))
            '    nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(pnRow, pnCol), 4)))
            '
            '    If nPers <> gPersonaNat And nRelacion <> gCapRelPersTitular Then
            '        MsgBox "Las Personas Jurídicas deben ser titulares de la cuenta", vbInformation, "Aviso"
            '        grdCliente.TextMatrix(pnRow, 3) = ""
            '        'cmdAgregar.SetFocus
            '        If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
            '        Exit Sub
            '    End If
            
    ' *** FIN RIRO

ElseIf pnCol = 6 Then
    
    ' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

            ''    If UCase(grdCliente.TextMatrix(grdCliente.Row, 6)) = "S" Or UCase(grdCliente.TextMatrix(grdCliente.Row, 6)) = "SI" Then
            ''        grdCliente.TextMatrix(grdCliente.Row, 6) = "SI"
            ''    ElseIf UCase(grdCliente.TextMatrix(grdCliente.Row, 6)) = "N" Or UCase(grdCliente.TextMatrix(grdCliente.Row, 6)) = "NO" Then
            ''        grdCliente.TextMatrix(grdCliente.Row, 6) = "NO"
            ''    ElseIf UCase(grdCliente.TextMatrix(grdCliente.Row, 6)) <> "S" And UCase(grdCliente.TextMatrix(grdCliente.Row, 6)) <> "SI" And UCase(grdCliente.TextMatrix(grdCliente.Row, 6)) <> "N" And UCase(grdCliente.TextMatrix(grdCliente.Row, 6)) <> "NO" Then
            ''        MsgBox "PRESIONE LA TECLA S SI LA FIRMA DE ESTE CLIENTE ES OBLIGATORIA." & vbCrLf & "PRESIONE LA TECLA N SI LA FIRMA DE ESTE CLIENTE NO ES OBLIGATORIA.", vbOKOnly + vbExclamation, App.Title
            ''        grdCliente.TextMatrix(grdCliente.Row, 6) = ""
            ''    End If
            '
            '    If UCase(Trim(grdCliente.TextMatrix(grdCliente.row, 6))) = "S" Or UCase(Trim(grdCliente.TextMatrix(grdCliente.row, 6))) = "SI" Then
            '        grdCliente.TextMatrix(grdCliente.row, 6) = "SI"
            '    ElseIf UCase(Trim(grdCliente.TextMatrix(grdCliente.row, 6))) = "N" Or UCase(Trim(grdCliente.TextMatrix(grdCliente.row, 6))) = "NO" Then
            '        grdCliente.TextMatrix(grdCliente.row, 6) = "NO"
            '    ElseIf UCase(Trim(grdCliente.TextMatrix(grdCliente.row, 6))) = "O" Or UCase(Trim(grdCliente.TextMatrix(grdCliente.row, 6))) = "OPCIONAL" Then
            '        grdCliente.TextMatrix(grdCliente.row, 6) = "OPCIONAL"
            '    Else
            '        MsgBox "PRESIONE LA TECLA S SI LA FIRMA DE ESTE CLIENTE ES OBLIGATORIA." & vbCrLf & "PRESIONE LA TECLA N SI LA FIRMA DE ESTE CLIENTE NO ES OBLIGATORIA." & vbCrLf & "PRESIONE LA TECLA O SI LA FIRMA DE ESTE CLIENTE  ES OPCIONAL.", vbOKOnly + vbExclamation, App.Title
            '    End If

    ' *** FIN RIRO

ElseIf pnCol = 9 Then
    
'Agregado por RIRO el 20130501, Proyecto Ahorro - Poderes
    grdCliente.row = IIf(pnRow + 1 = grdCliente.Rows, pnRow, pnRow + 1)
    grdCliente.Col = 3
    grdCliente.SetFocus
    seleccionarTipoCuentaXregla
    
ElseIf pnCol = 1 Then
' Condicion Agregada por RIRO el 20130501
' Se adicionó para evitar que se ejecuten los procedimientos: CuentaTitular, EvaluaTitular
' Estos procedimientos se ejecutarán de todas maneras en el evento 'OnEnterTextBuscar'.
    
    Exit Sub

End If


'ElseIf pnCol = 6 Then
'    Dim rsObl As Recordset
'    Set rsObl = New Recordset
'
'        rsObl.Fields.Append "Valor", adVarChar, 2
'        rsObl.Open
'        rsObl.AddNew "Valor", "SI"
'        rsObl.AddNew "Valor", "NO"
'
'
'    grdCliente.CargaCombo rsObl
'        If rsObl.State = 1 Then rsObl.Close
'        Set rsObl = Nothing
'
'End If

CuentaTitular
EvaluaTitular
End Sub

Private Sub grdCliente_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Dim rsCta As ADODB.Recordset
If grdCliente.TextMatrix(pnRow, pnCol) = "." Then
    If nProducto = gCapPlazoFijo Then
        Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        'Set rsCta = clsMant.GetCuentasPersona(grdCliente.TextMatrix(pnRow, 1), gCapAhorros, True, , nmoneda)
        Set rsCta = clsMant.GetCuentaAhorroTitularesPF(nTipoCuenta, ObtTodosTitulares, nmoneda)
        Set clsMant = Nothing
        txtCtaAhoAboInt.rs = rsCta
        Set rsCta = Nothing
        lblCuentaAbo.Visible = True
        txtCtaAhoAboInt.Visible = True
        MarcaSoloUnaFila pnRow
    End If
Else
    lblCuentaAbo.Visible = False
    txtCtaAhoAboInt.Visible = False
    Set rsCta = New ADODB.Recordset
    txtCtaAhoAboInt.rs = rsCta
End If
End Sub

Private Sub grdCliente_OnChangeCombo()

'If grdCliente.Col = 3 Then
' Dim nRelacion As CaptacRelacPersona
'
''    Dim clsGen As DGeneral
''    Dim rsRel As Recordset
''    Set clsGen = New DGeneral
''    Set rsRel = clsGen.GetConstante(gCaptacRelacPersona)
''    grdCliente.CargaCombo rsRel
''    If rsRel.State = 1 Then rsRel.Close
''    Set rsRel = Nothing
'
'Else
'     Dim rsObl As Recordset
'    Set rsObl = New Recordset
'
'        rsObl.Fields.Append "Valor", adVarChar, 2
'        rsObl.Open
'        rsObl.AddNew "Valor", "SI"
'        rsObl.AddNew "Valor", "NO"
'        rsObl.Update
'    grdCliente.CargaCombo rsObl
'        If rsObl.State = 1 Then rsObl.Close
'        Set rsObl = Nothing
'
'End If

' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

        'If grdCliente.col = 3 Then
        '    If Right(grdCliente.TextMatrix(grdCliente.row, 3), 2) = "10" Then
        '        grdCliente.TextMatrix(grdCliente.row, 6) = "SI"
        '    Else
        '        grdCliente.TextMatrix(grdCliente.row, 6) = "NO"
        '    End If
        'End If
        
' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

If grdCliente.Col = 3 Then

    If Val(Trim(Right(grdCliente.TextMatrix(grdCliente.row, 3), 2))) = 11 And _
       Val(Trim(grdCliente.TextMatrix(grdCliente.row, 4))) <= 1 Then
            grdCliente.TextMatrix(grdCliente.row, 9) = "AP"
    ElseIf Val(Trim(grdCliente.TextMatrix(grdCliente.row, 4))) <= 1 Then
        If intPunteroPJ_NA = 0 Then
            grdCliente.TextMatrix(grdCliente.row, 9) = "A"
        Else
            grdCliente.TextMatrix(grdCliente.row, 9) = "N/A"
        End If
    End If
    seleccionarTipoCuentaXregla
    
End If

' *** END RIRO

End Sub


Private Sub grdCliente_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim nNuevaPersoneria As PersPersoneria

If pbEsDuplicado And psDataCod <> "" Then 'RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
    MsgBox "Persona ya esta registrada en la relación.", vbInformation, "Aviso"
    grdCliente.EliminaFila grdCliente.row
ElseIf psDataCod = "" Then

    ' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
    If Not pbEsDuplicado Then
        grdCliente.TextMatrix(pnRow, 3) = ""
        grdCliente.TextMatrix(pnRow, 4) = 0
        grdCliente.TextMatrix(pnRow, 9) = ""
    End If
    ' *** END RIRO
    
    Exit Sub
    'grdCliente.EliminaFila grdCliente.Row
ElseIf psDataCod = gsCodPersUser Then
    MsgBox "No se puede aperturar cuenta en si mismo.", vbInformation, "Aviso"
    grdCliente.EliminaFila grdCliente.row
Else
    nNuevaPersoneria = grdCliente.PersPersoneria
    If nNuevaPersoneria <> gPersonaNat Then
        '*** BRGO 20111219 *********************************************
        'JUEZ 20141008 Comentado, esto se manejará por parámetros
        'If nProducto = gCapPlazoFijo And cboPrograma.ListIndex = 1 Then
        '    MsgBox "El subproducto de Plazo fijo no permite el registro de una Persona Jurídica"
        '    grdCliente.EliminaFila grdCliente.row
        '    nClientes = nClientes - 1
        '    If nClientes = 0 Then
        '        cmdEliminar.Enabled = False
        '    End If
        '    Exit Sub
        'End If
        '*** END BRGO **************************************************
        If nPersoneria <> gPersonaNat Then
            MsgBox "No es posible relacionar dos personas jurídicas en una misma cuenta.", vbInformation, "Aviso"
            grdCliente.EliminaFila grdCliente.row
            CuentaTitular
            EvaluaTitular
            Exit Sub
        Else
            nPersoneria = nNuevaPersoneria
        End If
        'MIOL 20121011, según OYP-RFC098-2012 *******
        If nProducto = gCapPlazoFijo Then
            If CLng(Trim(Right(cboFormaRetiro.Text, 4))) = 2 Then
                chkSubasta.Visible = True
            Else
                chkSubasta.Visible = False
            End If
        End If
        'END MIOL ***********************************
    Else
        nPersoneria = gPersonaNat
    End If
    
    'JUEZ 20141008 VERIFICAR PARAMETRO PERSONERIA *************
    If nProducto <> gCapCTS Then
        If nNuevaPersoneria = gPersonaNat And Not bParPersNat Then
            MsgBox "El producto no permite ingresar personas naturales", vbInformation, "Aviso"
            grdCliente.EliminaFila grdCliente.row
            nClientes = nClientes - 1
            If nClientes = 0 Then
                cmdEliminar.Enabled = False
            End If
            Exit Sub
        End If
        If nNuevaPersoneria <> gPersonaNat And Not bParPersJur Then
            MsgBox "El producto no permite ingresar personas jurídicas", vbInformation, "Aviso"
            grdCliente.EliminaFila grdCliente.row
            nClientes = nClientes - 1
            If nClientes = 0 Then
                cmdEliminar.Enabled = False
            End If
            Exit Sub
        End If
    End If
    'END JUEZ *************************************************
    
    grdCliente.TextMatrix(grdCliente.row, 4) = Trim(nPersoneria)
    
' *** RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

    Dim nContar, nContarNA, i As Integer
    Dim rsRel As New ADODB.Recordset
    Dim clsGen As COMDConstSistema.DCOMGeneral
    
    For i = 1 To grdCliente.Rows - 1
        If grdCliente.TextMatrix(i, 1) <> "" Then
            If Val(Trim(grdCliente.TextMatrix(i, 4))) > 1 Then
                nContar = nContar + 1
            Else
                If Trim(grdCliente.TextMatrix(i, 9)) = "N/A" Then
                    nContarNA = nContarNA + 1
                End If
            End If
        End If
    Next
    If nContar = 0 Then
        intPunteroPJ_NA = 0
    ElseIf nContar = 1 Then
        intPunteroPJ_NA = 1
    Else
        intPunteroPJ_NA = 1
        MsgBox "No es posible relacionar dos personas jurídicas en una misma cuenta", vbInformation, "Aviso"
        grdCliente.EliminaFila grdCliente.row
    End If
    
    Set clsGen = New COMDConstSistema.DCOMGeneral
    
    If intPunteroPJ_NA = 0 Then

        If Val(Trim(Right(grdCliente.TextMatrix(grdCliente.row, 3), 5))) <> gCapRelPersApoderado Then
            grdCliente.TextMatrix(grdCliente.row, 9) = "A"
        Else
            grdCliente.TextMatrix(grdCliente.row, 9) = "AP"
        End If
        
        If grdCliente.Rows > 2 Then
            grdCliente.ColumnasAEditar = "X-1-X-3-X-5-6-X-X-9"
            If nContarNA > 0 Then
                Set rsRel = clsGen.GetConstante(gCaptacRelacPersona, "13,14")
                Set clsGen = Nothing
                grdCliente.CargaCombo rsRel
                Set rsRel = Nothing
                For i = 1 To grdCliente.Rows - 1
                    grdCliente.TextMatrix(i, 3) = ""
                    grdCliente.TextMatrix(i, 9) = "A"
                Next
            End If

            conReglas
            
        Else
            grdCliente.ColumnasAEditar = "X-1-X-3-X-5-6-X-X-X"
            Set rsRel = clsGen.GetConstante(gCaptacRelacPersona, "13,14")
            Set clsGen = Nothing
            grdCliente.CargaCombo rsRel
            grdCliente.TextMatrix(1, 3) = ""
            Set rsRel = Nothing
        End If
    Else
        
        grdCliente.ColumnasAEditar = "X-1-X-3-X-5-6-X-X-X"
        
        If nContarNA = 0 Then
            
            Set clsGen = New COMDConstSistema.DCOMGeneral
            Set rsRel = clsGen.GetConstante(gCaptacRelacPersona, "13,14,11")
            Set clsGen = Nothing
            grdCliente.CargaCombo rsRel
            Set rsRel = Nothing
    
            For i = 1 To grdCliente.Rows - 1
                grdCliente.TextMatrix(i, 3) = ""
            Next
        
        End If
        
        For i = 1 To grdCliente.Rows - 1
           If Val(Trim(grdCliente.TextMatrix(i, 4))) > 1 Then
                grdCliente.TextMatrix(i, 9) = "PJ"
           Else
                If Val(Trim(Right(grdCliente.TextMatrix(i, 3), 5))) <> gCapRelPersApoderado Then
                    grdCliente.TextMatrix(i, 9) = "N/A"
                Else
                    grdCliente.TextMatrix(i, 9) = "AP"
                End If
           End If
        Next
        
        sinReglas
    End If
    seleccionarTipoCuentaXregla
    
'*** END RIRO
    
'-- Para agregar  La Direccion y el documento que se Utilizara para la Cartilla -- AVMM -- 06-02-2006
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim R As New ADODB.Recordset
Set ClsPersona = New COMDPersona.DCOMPersonas
Set R = ClsPersona.BuscaCliente(grdCliente.TextMatrix(grdCliente.row, 1), BusquedaCodigo)
    If Not (R.EOF And R.BOF) Then
       'grdCliente.TextMatrix(grdCliente.row, 7) = IIf(IsNull(R!cPersIDnroDNI), IIf(IsNull(R!cPersIDnroRUC), "", R!cPersIDnroRUC), R!cPersIDnroDNI)
       grdCliente.TextMatrix(grdCliente.row, 7) = IIf(R!cPersIDnroDNI <> "", R!cPersIDnroDNI, IIf(R!cPersIDnroRUC <> "", R!cPersIDnroRUC, R!cPersIDnroOtro)) 'JUEZ 20131002
       grdCliente.TextMatrix(grdCliente.row, 8) = R!cPersDireccDomicilio
    End If
Set ClsPersona = Nothing
'----------------------------------------------------------------------------------------------------
End If
'Add By Gitu 2010-08-06
lnTitularPJ = 0
If (grdCliente.TextMatrix(grdCliente.row, 3) = "" Or Left(grdCliente.TextMatrix(grdCliente.row, 3), 7) = "TITULAR") And (nPersoneria <> gPersonaNat) Then
    ValidaTasaInteres
    'txtPlazo.Enabled = False
    lnTitularPJ = 1
Else
    If grdCliente.TextMatrix(1, 3) = "" Or Left(grdCliente.TextMatrix(grdCliente.row, 3), 7) = "TITULAR" Then
        ValidaTasaInteres
        'txtPlazo.Enabled = True
    End If
End If
'End Gitu
CuentaTitular
EvaluaTitular
End Sub

' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

Private Sub grdCliente_OnRowDelete()

 Dim i As Integer
    Dim intPJ As Integer
    intPJ = -1
    
    'Verificando si dentro del grid, hay personas juridicas
    For i = 1 To grdCliente.Rows - 1
        'If grdCliente.TextMatrix(i, 9) = "PJ" Then
        If Val(Trim(grdCliente.TextMatrix(i, 4))) > 1 Then
            intPJ = i
        End If
    Next
    
    'Verifica la cantidad de clientes intervinientes en la apertura
    If grdCliente.Rows <= 2 Then
        cboTipoCuenta.ListIndex = 0
        sinReglas
        If Trim(grdCliente.TextMatrix(1, 1)) = "" Then
            grdCliente.Clear
            grdCliente.FormaCabecera
            intPunteroPJ_NA = 0
        Else
            grdCliente.ColumnasAEditar = "X-1-X-3-X-5-6-X-X-X"
            If intPJ = -1 Then
                If Val(Trim(Right(grdCliente.TextMatrix(grdCliente.row, 3), 5))) <> gCapRelPersApoderado Then
                    grdCliente.TextMatrix(grdCliente.row, 9) = "A"
                Else
                    grdCliente.TextMatrix(grdCliente.row, 9) = "AP"
                End If
                intPunteroPJ_NA = 0
            Else
                grdCliente.TextMatrix(grdCliente.row, 9) = "PJ"
                intPunteroPJ_NA = 1
            End If
            Exit Sub
        End If
    Else
        If intPJ = -1 Then
            conReglas
            grdCliente.ColumnasAEditar = "X-1-X-3-X-5-6-X-X-9"
            intPunteroPJ_NA = 0
            For i = 1 To grdCliente.Rows - 1
                If Val(Trim(Right(grdCliente.TextMatrix(i, 3), 5))) <> gCapRelPersApoderado Then
                    grdCliente.TextMatrix(i, 9) = "A"
                Else
                    grdCliente.TextMatrix(i, 9) = "AP"
                End If
            Next
        Else
            sinReglas
            intPunteroPJ_NA = 1
            grdCliente.ColumnasAEditar = "X-1-X-3-X-5-6-X-X-X"
                        
        End If
    End If
    
    ' Recargar el combo del grid segun sea el caso
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsRel As ADODB.Recordset
    
    Set clsGen = New COMDConstSistema.DCOMGeneral
    
    If intPunteroPJ_NA = 0 Then
        Set rsRel = clsGen.GetConstante(gCaptacRelacPersona, "13,14")
        Set clsGen = Nothing
        grdCliente.CargaCombo rsRel
    Else
        Set rsRel = clsGen.GetConstante(gCaptacRelacPersona, "11,13,14")
        Set clsGen = Nothing
        grdCliente.CargaCombo rsRel
    End If

    Set rsRel = Nothing

    For i = 1 To grdCliente.Rows - 1
        grdCliente.TextMatrix(i, 3) = ""
    Next

    seleccionarTipoCuentaXregla

End Sub

' *** END RIRO

Private Sub grdCliente_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)

' COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

        'If pnCol > 4 And grdCliente.TextMatrix(pnRow, 6) = "" Then
        ' MsgBox "DEBE INDICAR SI ES OBLIGATORIA O NO LA FIRMA DEL CLIENTE", vbOKOnly + vbInformation, App.Title
        'End If
        
' AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

    Dim sColumnas() As String
    sColumnas = Split(grdCliente.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
    'Si la validacion es en la columna Nº 9
    If pnCol = 9 Then
        grdCliente.TextMatrix(pnRow, 9) = UCase(grdCliente.TextMatrix(pnRow, 9))
        If intPunteroPJ_NA = 0 Then
            'Se compara el tamaño de la cadena de la celda grupos
            If Len(Trim(grdCliente.TextMatrix(pnRow, 9))) = 1 Then
                Dim blPuntero As Boolean
                blPuntero = False
                Dim x As Integer
                For x = 65 To 90 Step 1
                    If UCase(Chr(x)) = UCase(grdCliente.TextMatrix(pnRow, 9)) Then
                        blPuntero = True
                        Exit For
                    End If
                Next x
                If blPuntero = False Then
                     grdCliente.TextMatrix(grdCliente.row, 9) = ""
                End If
            Else
                grdCliente.TextMatrix(pnRow, 9) = ""
                Exit Sub
            End If
        Else
            'En caso de las personas juridicas, esta columna es de solo lectura
        End If
    End If

End Sub

Private Sub grdCliente_RowColChange()
Dim nFila As Long, nCol As Long
nFila = grdCliente.row
nCol = grdCliente.Col

' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

    If nCol = 9 Then
        If Val(Trim(Right(grdCliente.TextMatrix(nFila, 3), 2))) = 11 Then
           grdCliente.lbEditarFlex = False
        Else
           grdCliente.lbEditarFlex = True
        End If
    Else
        grdCliente.lbEditarFlex = True
    End If
    
' *** END RIRO

'EJVG20130912 *** Bloqueo del Código de Persona y Nombres
If nOperacion = gAhoApeTransf Or nOperacion = gPFApeTransf Or nOperacion = gCTSApeTransf _
    Or nOperacion = gAhoApeChq Or nOperacion = gPFApeChq Or nOperacion = gCTSApeChq Then
    If nCol = 1 Or nCol = 2 Then
        grdCliente.lbEditarFlex = False
    Else
        grdCliente.lbEditarFlex = True
    End If
End If
'END EJVG *******

If nCol = 3 Then
End If
End Sub


'Private Sub lblTasaEspecial_Change()
'If lblTasaEspecial.Visible Then
'  lblTasa.Caption = Format$(ConvierteTNAaTEA(Val(lblTasaEspecial.Caption)), "#,##0.00")
'  nTasaNominal = Val(lblTasaEspecial.Caption)
'End If
'End Sub

Private Sub TXTALIAS_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    If Me.fraTranferecia.Visible Then
        '***Modificado por ELRO el 20121015, según OYP-RFC024-2012
        'Me.cboTransferMoneda.SetFocus
        If cboTransferMoneda.Enabled Then
            'Me.cboTransferMoneda.SetFocus
            If cboTransferMoneda.Visible And cboTransferMoneda.Enabled Then cboTransferMoneda.SetFocus
        End If
        '***Fin Modificado por ELRO el 20121015*******************
    ElseIf fraDocumento.Visible Then
        txtGlosa.SetFocus
    End If
End If
End Sub

Private Sub txtCtaAhoAboInt_EmiteDatos()
If Not txtCtaAhoAboInt.rs Is Nothing Then
    If nClientes = 0 Then
        MsgBox "Debe seleccionar un Cliente que posea cuentas de ahorros.", vbInformation, "Aviso"
    Else
        If (txtCtaAhoAboInt.rs.EOF And txtCtaAhoAboInt.rs.BOF) Then
            MsgBox "Los titulares no poseen cuentas con el mismo tipo.", vbInformation, "Aviso"
        End If
    End If
Else
    If nClientes = 0 Then
        MsgBox "Debe Agregar al menos una persona y seleccionar la opción de Cuenta", vbInformation, "Aviso"
    Else
        MsgBox "Debe seleccionar un Cliente que posea cuentas de ahorros.", vbInformation, "Aviso"
    End If
End If
End Sub

Private Sub txtCtaAhoAboInt_Click(psCodigo As String, psDescripcion As String)
    Dim rsCta As New ADODB.Recordset
    If nProducto = gCapPlazoFijo Then 'Modificado por BRGO 20111115 para implementación de Ahorro Ecotaxi
        nTipoCuenta = CLng(Trim(Right(cboTipoCuenta.Text, 4)))
            Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
            Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
            Set rsCta = clsMant.GetCuentaAhorroTitularesPF(nTipoCuenta, ObtTodosTitulares, nmoneda)
            Set clsMant = Nothing
            If Not rsCta.EOF Then
                txtCtaAhoAboInt.rs = rsCta
            End If
            Set rsCta = Nothing
    ElseIf nProducto = gCapAhorros Then
        
    End If
End Sub

Private Sub txtCtaAhoAboInt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'cmdAgregar.SetFocus
    If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
End If
End Sub

Private Sub txtCuentaCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGlosa.SetFocus
    End If
End Sub

'JUEZ 20131212 *************************
Private Sub txtCuentaCargo_LostFocus()
    ValidaCargoCta
End Sub
'END JUEZ ******************************

Private Sub txtDisp_Change()
    Dim ntmpTotal As Double
    
    ntmpTotal = Val(txtInta.Text) + Val(txtDisp.Text) + Val(txtDU.Text)
    
    If nOperacion <> gCTSApeChq And nOperacion <> gCTSApeTransf Then
        lblTotTran.Caption = Format(ntmpTotal, "#0.00")
    End If
    
    If ntmpTotal > 0 Then
        lblInta.Caption = Format((Val(txtInta.Text) / ntmpTotal) * 100, "#0.0000")
        lblDisp.Caption = Format((Val(txtDisp.Text) / ntmpTotal) * 100, "#0.0000")
        lblDu.Caption = Format((Val(txtDU.Text) / ntmpTotal) * 100, "#0.0000")
    End If
    
End Sub

Private Sub txtDisp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If txtDisp = "" Then
        txtDisp.Text = "0.00"
     Else
         txtDisp.Text = Format(txtDisp.Text, "#0.00")
     End If
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim bOrdPag As Boolean
    Dim nMonto As Double

    If cboTipoTasa.Text <> "" Then nTipoTasa = CLng(Right(cboTipoTasa.Text, 4))
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    bOrdPag = IIf(chkOrdenPago.value = 1, True, False)
    nMonto = txtMonto.value

    If chkTasaPreferencial.value = vbUnchecked Then
           If nProducto = gCapPlazoFijo Then
               If txtPlazo <> "" Then
                   nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, CLng(Val(txtPlazo.Text)), lblTotTran, gsCodAge)
                   lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
               End If
           Else
               nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, , lblTotTran, gsCodAge, bOrdPag)
               lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
           End If
     End If
     
     txtDU.SetFocus
ElseIf KeyAscii <> 13 And Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46) Then
    KeyAscii = 0
End If
End Sub

Private Sub txtDisp_LostFocus()
     If txtDisp = "" Then
        txtDisp.Text = "0.00"
     Else
         txtDisp.Text = Format(txtDisp.Text, "#0.00")
     End If
End Sub

Private Sub txtDU_Change()
 Dim ntmpTotal As Double
    ntmpTotal = Val(txtInta.Text) + Val(txtDisp.Text) + Val(txtDU.Text)
    
    If nOperacion <> gCTSApeChq And nOperacion <> gCTSApeTransf Then
        lblTotTran.Caption = Format(ntmpTotal, "#0.00")
    End If
    
    If ntmpTotal > 0 Then
        lblInta.Caption = Format((Val(txtInta.Text) / ntmpTotal) * 100, "#0.0000")
        lblDisp.Caption = Format((Val(txtDisp.Text) / ntmpTotal) * 100, "#0.0000")
        lblDu.Caption = Format((Val(txtDU.Text) / ntmpTotal) * 100, "#0.0000")
    End If
End Sub

Private Sub txtDU_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If txtDU = "" Then
        txtDU.Text = "0.00"
    Else
        txtDU.Text = Format(txtDU.Text, "#0.00")
    End If
    
    If cmdgrabar.Enabled = True Then
        cmdgrabar.SetFocus
    End If
End If
End Sub

Private Sub txtDU_LostFocus()
If txtDU = "" Then
    txtDU.Text = "0.00"
Else
    txtDU.Text = Format(txtDU.Text, "#0.00")
End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    If txtMonto.Enabled Then
        txtMonto.SetFocus
    Else
        If cmdgrabar.Enabled = True Then
            cmdgrabar.SetFocus
        End If
    End If
End If
End Sub

Private Sub txtInstitucion_EmiteDatos()
If txtInstitucion.Text <> "" Then
    If chkRelConv.value = 0 Then
        lblInstitucion = txtInstitucion.psDescripcion
        If cboMoneda.Enabled Then
            cboMoneda.SetFocus
        End If
    Else
        If Not ValidaInstConv(txtInstitucion.Text) Then
            txtInstitucion.Text = ""
            lblInstitucion.Caption = ""
            MsgBox "La Institucion no esta para convenio de Depositos", vbInformation, "SISTEMA"
        Else
            lblInstitucion.Caption = txtInstitucion.psDescripcion
            cmdgrabar.SetFocus
        End If
    End If
End If
ValidaTasaInteres 'JUEZ 20140319
End Sub

Private Sub txtinstitucion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdAgregar.Enabled And cmdAgregar.Visible Then
        'cmdAgregar.SetFocus
        If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
    Else
        If cmdEliminar.Enabled And cmdEliminar.Visible Then cmdEliminar.SetFocus
    End If
End If
End Sub

Private Sub txtInta_Change()
Dim ntmpTotal As Double
    ntmpTotal = Val(txtInta.Text) + Val(txtDisp.Text) + Val(txtDU.Text)
    If nOperacion <> gCTSApeChq And nOperacion <> gCTSApeTransf Then
        lblTotTran.Caption = Format(ntmpTotal, "#0.00")
    End If
    
    If ntmpTotal > 0 Then
        lblInta.Caption = Format((Val(txtInta.Text) / ntmpTotal) * 100, "#0.0000")
        lblDisp.Caption = Format((Val(txtDisp.Text) / ntmpTotal) * 100, "#0.0000")
        lblDu.Caption = Format((Val(txtDU.Text) / ntmpTotal) * 100, "#0.0000")
    End If
End Sub

Private Sub txtInta_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
     If txtInta = "" Then
        txtInta.Text = "0.00"
     Else
        txtInta.Text = Format(txtInta.Text, "#0.00")
     End If
     
     txtDisp.SetFocus
     
ElseIf KeyAscii <> 13 And Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46) Then
    KeyAscii = 0
End If

End Sub

Private Sub txtInta_LostFocus()
     If txtInta = "" Then
        txtInta.Text = "0.00"
     Else
        txtInta.Text = Format(txtInta.Text, "#0.00")
     End If
     
End Sub
Private Sub TxtMinFirmas_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    If ValidarFirmas = False Then Exit Sub
    txtAlias.SetFocus
 ElseIf KeyAscii <> 13 And Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
 End If
End Sub

Private Sub txtMinFirmas_LostFocus()
 If Val(TxtMinFirmas.Text) < Val(Label18.Tag) Then
   If Trim(Right(cboTipoCuenta.Text, 1)) = 1 Then
    MsgBox "EL NRO MINIMO DE FIRMAS OBLIGATORIAS DEBEN SER " & CStr(Label18.Tag), vbOKOnly + vbInformation, "AVISO"
    TxtMinFirmas.Text = CStr(Label18.Tag)
   End If
 End If
End Sub

Private Sub txtMonto_Change()

' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

If Trim(txtMonto.Text) = "." Then
    txtMonto.Text = 0
    Exit Sub
End If

' *** END RIRO

Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim nMonto As Double

    nMonto = nMonto
    nMonto = txtMonto
    ValidaTasaInteres
    'If gbITFAplica And nProducto <> gCapCTS Then       'Filtra para CTS
    If gbITFAplica And nProducto <> gCapCTS And nOperacion <> gAhoApeCargoCta And nOperacion <> gPFApeCargoCta Then 'JUEZ 20131212 Para exonerar ITF en aperturas con cargo a cuenta
        If nMonto > gnITFMontoMin Then
            If Me.chkExoITF.value = 0 Then
                If nOperacion = gAhoApeTransf Or nOperacion = gPFApeTransf Then
                    '***Modificado por ELRO el 20120725, según OYP-RFC024-2012
                    'Me.lblITF.Caption = Format(0, "#,##0.00")
                    Me.LblItf.Caption = Format(fgITFCalculaImpuesto(nMonto), "#,##0.00")
                    '***Fin Modificado por ELRO el 20120725*******************
                ElseIf nProducto = gCapAhorros And (nOperacion <> gAhoApeChq) Then
                    Me.LblItf.Caption = Format(fgITFCalculaImpuesto(nMonto), "#,##0.00")
                Else
                    Me.LblItf.Caption = Format(fgITFCalculaImpuesto(nMonto), "#,##0.00")
                End If
                '*** BRGO 20110908 ************************************************
                nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.LblItf.Caption))
                If nRedondeoITF > 0 Then
                   Me.LblItf.Caption = Format(CCur(Me.LblItf.Caption) - nRedondeoITF, "#,##0.00")
                End If
                '*** END BRGO
                If bInstFinanc Then LblItf.Caption = "0.00" 'JUEZ 20140414
            Else
            
                Me.LblItf.Caption = "0.00"
            End If
            If nOperacion = gAhoRetOPCanje Or nOperacion = gAhoRetOPCertCanje Or nOperacion = gAhoRetFondoFijoCanje Then
                Me.lblTotal.Caption = Format(0, "#,##0.00")
            ElseIf nOperacion = gPFApeChq Or nOperacion = gAhoApeChq Then
                If nProducto = gCapAhorros And gbITFAsumidoAho Then
                    Me.lblTotal.Caption = Format(0, "#,##0.00")
                ElseIf nProducto = gCapPlazoFijo And gbITFAsumidoPF Then
                    Me.lblTotal.Caption = Format(0, "#,##0.00")
                Else
                    Me.lblTotal.Caption = Format(CCur(Me.LblItf.Caption), "#,##0.00")
                End If
            Else
                If nProducto = gCapAhorros And gbITFAsumidoAho Then
                    Me.lblTotal.Caption = Format(nMonto, "#,##0.00")
                ElseIf nProducto = gCapPlazoFijo And gbITFAsumidoPF Then
                    Me.lblTotal.Caption = Format(nMonto, "#,##0.00")
                Else
                    If Me.chkITFEfectivo.value = 1 Then
                        Me.lblTotal.Caption = Format(nMonto + CCur(Me.LblItf.Caption), "#,##0.00")
                    Else
                        Me.lblTotal.Caption = Format(nMonto, "#,##0.00")
                    End If
                End If
            End If
        End If
        If bInstFinanc Then LblItf.Caption = "0.00" 'JUEZ 20140414
    Else
        Me.LblItf.Caption = Format(0, "#,##0.00")
        
        If nOperacion = gCTSDepChq Then
            Me.lblTotal.Caption = Format(0, "#,##0.00")
        Else
            Me.lblTotal.Caption = Format(nMonto, "#,##0.00")
        End If
        chkITFEfectivo_Click
    End If
    
    If nMonto = 0 Then
        Me.LblItf.Caption = "0.00"
        Me.lblTotal.Caption = "0.00"
        chkITFEfectivo_Click
    End If
    
End Sub

Private Sub chkITFEfectivo_Click()
    If nProducto = gCapAhorros And gbITFAsumidoAho Then
    
    ElseIf nProducto = gCapPlazoFijo And gbITFAsumidoPF Then
    
    Else
        If chkITFEfectivo.value = 1 Then
            Me.lblTotal.Caption = Format(Me.txtMonto.value + CCur(Me.LblItf.Caption), "#,##0.00")
        Else
            Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00")
        End If
    End If
End Sub

Private Sub txtMonto_GotFocus()
txtMonto.MarcaTexto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdgrabar.Enabled = True Then
        cmdgrabar.SetFocus
    End If
End If
End Sub
'JUEZ 20141008 ***********************
Private Sub txtMonto_LostFocus()
    If nProducto <> gCapCTS Then ValidaMontoMinimoApertura
End Sub
'END JUEZ ****************************
Private Sub txtMontoAbonar_KeyPress(KeyAscii As Integer)
Dim loRs As COMDConstSistema.DCOMGeneral
Set loRs = New COMDConstSistema.DCOMGeneral
Dim nMontoMin As Double
    If KeyAscii = 13 Then
        If Trim(Right(cboPrograma.Text, 1)) = 2 Or Trim(Right(cboPrograma.Text, 1)) = 3 Or Trim(Right(cboPrograma.Text, 1)) = 4 Then
           If Trim(Right(cboPrograma.Text, 1)) = 3 Or Trim(Right(cboPrograma.Text, 1)) = 4 Then
                nMontoMin = loRs.GetParametro(2000, 2093)
                If txtMontoAbonar.Text < nMontoMin Then
                  MsgBox "El Monto de Abono debe de ser Igual o Mayor a " & nMontoMin, vbInformation, "Aviso"
                  Exit Sub
                Else
                  'cmdAgregar.SetFocus
                  If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
                  Exit Sub
                End If
            ElseIf Trim(Right(cboPrograma.Text, 1)) = 2 Then
                nMontoMin = loRs.GetParametro(2000, 2094)
                If txtMontoAbonar.Text < nMontoMin Then
                  MsgBox "El Monto de Abono debe de ser Igual o Mayor a " & nMontoMin, vbInformation, "Aviso"
                  Exit Sub
                Else
                  'cmdAgregar.SetFocus
                  If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
                  Exit Sub
                End If
            End If
        Else
           'cmdAgregar.SetFocus
           If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
           Exit Sub
        End If
    End If
End Sub



Private Sub txtNumFirmas_GotFocus()
With txtNumFirmas
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtNumFirmas_KeyPress(KeyAscii As Integer)
Dim i As Integer
Dim nCon As Integer
nCon = 0
If KeyAscii = 13 Then
'    If txtNumFirmas.Text <> grdCliente.Rows - 1 Then
'        MsgBox "Numero de Firmas debe de ser " & grdCliente.Rows - 1, vbInformation, "Aviso"
'        txtNumFirmas.Text = grdCliente.Rows - 1
'    Else
        If ValidarFirmas = False Then Exit Sub
'    End If
    If cmdDocumento.Visible Then
        cmdDocumento.SetFocus
    End If
    Exit Sub
End If
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtNumSolicitud_KeyPress(KeyAscii As Integer)

Dim i As Integer
If KeyAscii = 13 And Trim(grdCliente.TextMatrix(1, 1)) = "" Then
    MsgBox "DEBE INGRESAR LOS TITULARES PARA LA APERTURA DE ESTA CUENTA", vbOKOnly + vbInformation, "AVISO"
    Exit Sub
End If

If KeyAscii = 13 And Trim(grdCliente.TextMatrix(1, 1)) <> "" Then
    Dim oServ As COMDCaptaServicios.DCOMCaptaServicios
    Dim rsTasa As New ADODB.Recordset
    Dim sPersCod As String
        
    Set oServ = New COMDCaptaServicios.DCOMCaptaServicios
    i = 1
    
    While i <= grdCliente.Rows - 1
      If UCase(Left(grdCliente.TextMatrix(i, 3), 7)) = "TITULAR" Then
          sPersCod = grdCliente.TextMatrix(i, 1)
          Set oServ = New COMDCaptaServicios.DCOMCaptaServicios
          Set rsTasa = oServ.GetUltEstadoTP(CLng(Me.txtNumSolicitud.Text), sPersCod, nProducto, nmoneda)
          
          If Not rsTasa.EOF Then
            If rsTasa.RecordCount > 0 Then GoTo Continua
          End If
          
      End If
        Set oServ = Nothing
        i = i + 1
        
        If i > grdCliente.Rows - 1 Then
            MsgBox "SOLICITUD NO ENCONTRADA." & vbCrLf & "VERIFIQUE LA INFORMACION DE LOS TITULARES", vbOKOnly + vbInformation, "AVISO"
            Set rsTasa = Nothing
            Exit Sub
        End If
        
    Wend
    
Continua:
 If Not rsTasa.EOF Then
    If rsTasa!nEstado = 1 Then
        vSperscod = sPersCod
        cboTipoTasa.ListIndex = 1
        
        ' Agregado por RIRO el 20130411
        For i = 0 To cboPrograma.ListCount - 1
            If Trim(Left(cboPrograma.List(i), 40)) = rsTasa!cConsDescripcion Then
                cboPrograma.ListIndex = i
                Exit For
            End If
        Next
        sPerSolicitud = rsTasa!cPersCod
        cboPrograma.Enabled = False
        ' Fin RIRO
        
        nTasaNominal = rsTasa!nTasa
        lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.0000")
        lblEstadoSol.Caption = "APROBADA"
        lblEstadoSol.Tag = rsTasa!nEstado
        txtMonto.Enabled = True
        txtMonto.Text = Format(rsTasa!nMonto, "#,##0.00")
        txtMonto.Enabled = False
        chkPermanente.value = IIf(rsTasa!bPermanente = 0, vbUnchecked, vbChecked)
        txtPlazo.Enabled = True
        txtPlazo.Text = rsTasa!nPlazo
        If nProducto = gCapPlazoFijo Then
            txtPlazo.Enabled = False
        End If
    ElseIf rsTasa!nEstado = 2 Then
         MsgBox "ESTA SOLICITUD YA FUE ATENDIDA"
         Set rsTasa = Nothing
         Set oServ = Nothing
         
         Exit Sub
    ElseIf rsTasa!nEstado = 3 Then
        Dim saux As String
        
        If rsTasa!NEXTORNO = 0 Then
            saux = " POR SOLICITUD"
        ElseIf rsTasa!NEXTORNO = 1 Then
            saux = " POR APROBACION"
        ElseIf rsTasa!NEXTORNO = 2 Then
            saux = " POR APERTURA"
        ElseIf rsTasa!NEXTORNO = 4 Then
            saux = " POR RECHAZO"
        End If
        
        MsgBox "ESTA SOLICITUD SE ENCUENTRA EXTORNADA" & saux
        Set rsTasa = Nothing
        Set oServ = Nothing
        
        Exit Sub
        
    Else
        lblEstadoSol.Caption = IIf(rsTasa!nEstado = 0, "SOLICITADO", "RECHAZADA")
        
    End If
Else
    MsgBox "SOLICITUD NO ENCONTRADA", vbOKOnly + vbInformation, "AVISO"
End If

    Set rsTasa = Nothing
    Set oServ = Nothing
End If

If KeyAscii <> 13 And Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46) Then
        KeyAscii = 0
End If

End Sub

Private Sub txtPlazo_Change()
ValidaTasaInteres
End Sub

Private Sub txtPlazo_GotFocus()
With txtPlazo
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtPlazo_KeyPress(KeyAscii As Integer)
Dim loRs As COMDConstSistema.DCOMGeneral
Set loRs = New COMDConstSistema.DCOMGeneral
Dim nMinPlazo As Double
Dim nMaxPlazo As Double
If KeyAscii = 13 Then
    If nOperacion = gAhoApeChq Or nOperacion = gAhoApeEfec Then  'BRGO 20111220 Se Agregó gPFApeEfec y gPFApeChq
        If Trim(Right(cboPrograma.Text, 1)) <> 0 Then
           If Trim(Right(cboPrograma.Text, 1)) = 4 Then 'Or Trim(Right(cboPrograma.Text, 1)) = 4 Then
                'nMinPlazo = loRs.GetParametro(2000, gPlazoMinCancelPandero)
                nMinPlazo = nParPlazoMin 'JUEZ 20141008 Nuevos Parámetros
                If txtPlazo.Text < nMinPlazo Then
                  MsgBox "El Plazo debe de ser Igual o Mayor a " & nMinPlazo, vbInformation, "Aviso"
                  Exit Sub
                Else
                  txtMontoAbonar.SetFocus
                  Exit Sub
                End If
'*** Comentado por BRGO 20111220 ******************************************
'            ElseIf Trim(Right(cboPrograma.Text, 1)) = 2 Then
'                nMinPlazo = loRs.GetParametro(2000, 2095)
'                nMaxPlazo = loRs.GetParametro(2000, 2096)
'                If txtPlazo.Text < nMinPlazo Or txtPlazo.Text > nMaxPlazo Then
'                  MsgBox "El Plazo debe estar entre " & nMinPlazo & " y " & nMaxPlazo, vbInformation, "Aviso"
'                  Exit Sub
'                Else
'                  txtMontoAbonar.SetFocus
'                  Exit Sub
'                End If
            ElseIf Trim(Right(cboPrograma.Text, 1)) = 1 Then
                'nMinPlazo = loRs.GetParametro(2000, gPlazoMinCancelNañito)
                nMinPlazo = nParPlazoMin 'JUEZ 20141008 Nuevos Parámetros
                If txtPlazo.Text < nMinPlazo Then
                    MsgBox "El Plazo debe de ser Igual o Mayor a " & nMinPlazo, vbInformation, "Aviso"
                    Exit Sub
                Else
                     'cmdAgregar.SetFocus
                     If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
                     Exit Sub
                End If
            End If
        Else
           'cmdAgregar.SetFocus
           If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
           Exit Sub
        End If
    ElseIf nOperacion = gPFApeEfec Or nOperacion = gPFApeChq Then
        If Trim(Right(cboPrograma.Text, 1)) = 3 Then
            'nMinPlazo = loRs.GetParametro(2000, gPlazoMinCancelPandero)
            nMinPlazo = nParPlazoMin 'JUEZ 20141008 Nuevos Parámetros
            If txtPlazo.Text < nMinPlazo Then
                MsgBox "El Plazo debe de ser Igual o Mayor a " & nMinPlazo, vbInformation, "Aviso"
                Exit Sub
            Else
                txtMontoAbonar.SetFocus
                Exit Sub
            End If
        ElseIf Trim(Right(cboPrograma.Text, 1)) = 2 Then
            'nMinPlazo = loRs.GetParametro(2000, 2095)
            'nMaxPlazo = loRs.GetParametro(2000, 2096)
            nMinPlazo = nParPlazoMin 'JUEZ 20141008 Nuevos Parámetros
            nMaxPlazo = nParPlazoMax 'JUEZ 20141008 Nuevos Parámetros
            If txtPlazo.Text < nMinPlazo Or txtPlazo.Text > nMaxPlazo Then
                MsgBox "El Plazo debe estar entre " & nMinPlazo & " y " & nMaxPlazo, vbInformation, "Aviso"
                Exit Sub
            Else
                txtMontoAbonar.SetFocus
                Exit Sub
            End If
        Else
            'cmdAgregar.SetFocus
            If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
            Exit Sub
        End If
    Else
        'cmdAgregar.SetFocus
        If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
        Exit Sub
    End If
End If
Set loRs = Nothing
KeyAscii = NumerosEnteros(KeyAscii)
End Sub


Private Sub txtTransferGlosa_GotFocus()
    txtTransferGlosa.SelStart = 0
    txtTransferGlosa.SelLength = 500
End Sub

Private Sub txtTransferGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    If txtMonto.Enabled Then
        txtMonto.SetFocus
    Else
        If cmdgrabar.Enabled = True Then
            cmdgrabar.SetFocus
        End If
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

Public Sub CargaTitulares(ByRef MatTitular As Variant)
    Dim nNumTit As Integer
    Dim i As Integer

    'ARCV 13-02-2007
    'nNumTit = grdCliente.Rows - 1
    
    'ReDim MatTitular(nNumTit, 3)
    
    'For i = 1 To grdCliente.Rows - 1
    '    If UCase(Left(grdCliente.TextMatrix(i, 3), 7)) = "TITULAR" Then
    '        MatTitular(i, 1) = grdCliente.TextMatrix(i, 2)
    '        MatTitular(i, 2) = grdCliente.TextMatrix(i, 7)
    '        MatTitular(i, 3) = grdCliente.TextMatrix(i, 8)
    '    End If
    'Next i
    
    Dim MatIndices As Variant
    ReDim MatIndices(0)
    
    For i = 1 To grdCliente.Rows - 1
        If nProducto = gCapAhorros And Trim(Right(cboPrograma.Text, 2)) = 1 Then 'JUEZ 20150121
            ReDim Preserve MatIndices(UBound(MatIndices) + 1)
            MatIndices(UBound(MatIndices) - 1) = i
        Else
            If UCase(Left(grdCliente.TextMatrix(i, 3), 7)) = "TITULAR" Then
                ReDim Preserve MatIndices(UBound(MatIndices) + 1)
                MatIndices(UBound(MatIndices) - 1) = i
            End If
        End If
    Next i
    
    'ReDim MatTitular(UBound(MatIndices) + 1, 3)
    ReDim MatTitular(UBound(MatIndices) + 1, 4) 'JUEZ 20150121
    
    For i = 1 To UBound(MatIndices)
        MatTitular(i, 1) = grdCliente.TextMatrix(MatIndices(i - 1), 2)
        MatTitular(i, 2) = grdCliente.TextMatrix(MatIndices(i - 1), 7)
        MatTitular(i, 3) = grdCliente.TextMatrix(MatIndices(i - 1), 8)
        MatTitular(i, 4) = UCase(Left(grdCliente.TextMatrix(i, 3), 20)) 'JUEZ 20150121
    Next i
    '-----------------
    
End Sub

Public Function NroTitularesObligatorio() As Integer
    Dim nNumTit As Integer
    Dim i As Integer
  
    For i = 1 To grdCliente.Rows - 1
        If UCase(Left(grdCliente.TextMatrix(i, 6), 7)) = "SI" Then
            nNumTit = nNumTit + 1
        End If
    Next i
    NroTitularesObligatorio = nNumTit + 1
End Function

Public Function NroTitularesRelacion() As Integer
    Dim nNumTit As Integer
    Dim nNumRep As Integer
    Dim i As Integer
  
    For i = 1 To grdCliente.Rows - 1
        If UCase(Left(grdCliente.TextMatrix(i, 3), 7)) = "TITULAR" Then
            nNumTit = nNumTit + 1
        End If
    Next i
    
    For i = 1 To grdCliente.Rows - 1
        If UCase(Left(grdCliente.TextMatrix(i, 3), 7)) = "REP. LE" Then
            nNumRep = nNumRep + 1
        End If
    Next i
    
    If nNumRep <> 0 Then
       If nNumRep <> 1 Then
          NroTitularesRelacion = nNumRep - nNumTit
       Else
          NroTitularesRelacion = 0
       End If
    Else
        NroTitularesRelacion = nNumTit - 1
    End If
    
End Function

Public Function ValidarFirmas() As Boolean
 'Firma Individual
   Dim i As Integer
   ValidarFirmas = True
   If grdCliente.TextMatrix(1, 3) = "" Then Exit Function
   If Trim(Right(cboTipoCuenta, 4)) = 0 Then
       i = 1
     If grdCliente.Rows - 1 > 1 Then
        If Trim(Right(grdCliente.TextMatrix(i, 3), 4)) = gCapRelPersTitular And Trim(Right(grdCliente.TextMatrix(i + 1, 3), 4)) = gCapRelPersApoderado Then
            If grdCliente.Rows - 1 >= 2 Then
                'MsgBox "Cuenta Individual solo permite un Participante", vbInformation, "Aviso"
                'cboTipoCuenta.ListIndex = 1
                TxtMinFirmas.Text = 1
                txtNumFirmas.Text = 1
                ValidarFirmas = False
                Exit Function
            Else
                TxtMinFirmas.Text = 1
                txtNumFirmas.Text = 1
            End If
       ElseIf Trim(Right(grdCliente.TextMatrix(i, 3), 4)) = gCapRelPersApoderado And Trim(Right(grdCliente.TextMatrix(i + 1, 3), 4)) = gCapRelPersTitular Then
             If grdCliente.Rows - 1 >= 2 Then
                'MsgBox "Cuenta Individual solo permite un Participante", vbInformation, "Aviso"
                'cboTipoCuenta.ListIndex = 1
                TxtMinFirmas.Text = 1
                txtNumFirmas.Text = 1
                ValidarFirmas = False
                Exit Function
            Else
                TxtMinFirmas.Text = 1
                txtNumFirmas.Text = 1
            End If
       End If
    Else
             
             If grdCliente.Rows - 1 > 1 Then
                  'MsgBox "Cuenta Individual solo permite un Participante", vbInformation, "Aviso"
                  'cboTipoCuenta.ListIndex = 1
                  TxtMinFirmas.Text = 1
                  txtNumFirmas.Text = 1
                  ValidarFirmas = False
                  Exit Function
'            Else
'                If CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4))) = gCapRelPersTitular Then
'                   MsgBox "Cuenta Individual solo permite un Participante", vbInformation, "Aviso"
'                   TxtMinFirmas.Text = 1
'                   txtNumFirmas.Text = 1
'                   ValidarFirmas = False
'                   Exit Function
'                End If
            End If
    End If
      
    ElseIf Trim(Right(cboTipoCuenta, 4)) = 1 Then
        
'        If TxtMinFirmas.Text <> NroTitularesObligatorio Then
'            MsgBox "El Nro de Firmas debe de ser " & NroTitularesObligatorio, vbInformation, "Aviso"
'            TxtMinFirmas.Text = NroTitularesObligatorio
'            ValidarFirmas = False
'           Exit Function
'        Else
'            TxtMinFirmas.Text = NroTitularesObligatorio
'        End If
    ElseIf Trim(Right(cboTipoCuenta.Text, 1)) = 2 Then
        For i = 1 To grdCliente.Rows - 1
            If grdCliente.TextMatrix(i, 6) = "SI" Then
                'MsgBox "No debe de existir Firmas Obligatorias", vbInformation, "Aviso"
                ValidarFirmas = False
                Exit Function
            End If
        Next i
'        If TxtMinFirmas.Text > NroTitularesRelacion Then
'            MsgBox "El Nro de Firmas no coincide para una cuenta Indistinta", vbInformation, "Aviso"
'            TxtMinFirmas.Text = NroTitularesRelacion
'            ValidarFirmas = False
'            Exit Function
'        Else
'             TxtMinFirmas.Text = NroTitularesRelacion
'        End If
    End If
          
End Function

Public Function ValidaFlexVacio() As Boolean
    Dim i As Integer
    For i = 1 To grdCliente.Rows - 1
        If grdCliente.TextMatrix(i, 1) = "" Then
            ValidaFlexVacio = True
            Exit Function
        End If
    Next i
End Function
Sub ImprimeCartillaAhorro(MatTitular() As String, ByVal psCtaCod As String, ByVal pnTasa As Double, ByVal pnMonto As Double)
    If Right(cboPrograma.Text, 1) = 0 Then
        ImpreCartillaAhoCorriente MatTitular, psCtaCod, pnTasa, pnMonto
    End If
End Sub

Private Function ValidaMonPlazoAho() As Boolean
    Dim loRs As COMDConstSistema.DCOMGeneral
    Set loRs = New COMDConstSistema.DCOMGeneral
    Dim nMinPlazo As Double
    Dim nMaxPlazo As Double
    Dim nMontoMin  As Double
    ValidaMonPlazoAho = True
    If nOperacion = gAhoApeChq Or nOperacion = gAhoApeEfec Then
        If Trim(Right(cboPrograma.Text, 1)) <> 0 Then
            If Trim(Right(cboPrograma.Text, 1)) = 3 Or Trim(Right(cboPrograma.Text, 1)) = 4 Then
                'nMinPlazo = loRs.GetParametro(2000, gPlazoMinCancelPandero)
                nMinPlazo = nParPlazoMin 'JUEZ 20141008 Nuevos Parámetros
                If txtPlazo.Text < nMinPlazo Then
                  MsgBox "El Plazo debe de ser Igual o Mayor a " & nMinPlazo, vbInformation, "Aviso"
                  txtPlazo.SetFocus
                  ValidaMonPlazoAho = False
                  Exit Function
                End If
            ElseIf Trim(Right(cboPrograma.Text, 1)) = 2 Then
                'nMinPlazo = loRs.GetParametro(2000, 2095)
                'nMaxPlazo = loRs.GetParametro(2000, 2096)
                nMinPlazo = nParPlazoMin 'JUEZ 20141008 Nuevos Parámetros
                nMaxPlazo = nParPlazoMax 'JUEZ 20141008 Nuevos Parámetros
                If txtPlazo.Text < nMinPlazo Or txtPlazo.Text > nMaxPlazo Then
                  MsgBox "El Plazo debe estar entre " & nMinPlazo & " y " & nMaxPlazo, vbInformation, "Aviso"
                  txtPlazo.SetFocus
                  ValidaMonPlazoAho = False
                  Exit Function
                End If
            ElseIf Trim(Right(cboPrograma.Text, 1)) = 1 Then
                'RIRO20131102 Comentado
                'nMinPlazo = loRs.GetParametro(2000, gPlazoMinCancelNañito)
                'If txtPlazo.Text < nMinPlazo Then                '
                '    MsgBox "El Plazo debe de ser Igual o Mayor a " & nMinPlazo, vbInformation, "Aviso"
                '    ValidaMonPlazoAho = False
                '    Exit Function
                'Else
                '    cmdAgregar.SetFocus
                '    If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
                '    Exit Function
                'End If
                If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
                Exit Function
            End If
           
            If Trim(Right(cboPrograma.Text, 1)) = 3 Or Trim(Right(cboPrograma.Text, 1)) = 4 Then
                nMontoMin = loRs.GetParametro(2000, 2093)
                If txtMontoAbonar.Text < nMontoMin Then
                  MsgBox "El Monto de Abono debe de ser Igual o Mayor a " & nMontoMin, vbInformation, "Aviso"
                  txtMontoAbonar.SetFocus
                  ValidaMonPlazoAho = False
                  Exit Function
                End If
            ElseIf Trim(Right(cboPrograma.Text, 1)) = 2 Then
                nMontoMin = loRs.GetParametro(2000, 2094)
                If txtMontoAbonar.Text < nMontoMin Then
                  MsgBox "El Monto de Abono debe de ser Igual o Mayor a " & nMontoMin, vbInformation, "Aviso"
                  txtMontoAbonar.SetFocus
                  ValidaMonPlazoAho = False
                  Exit Function
                End If
            End If
        End If

    End If
End Function
'***Agregado por ELRO el 20120313, según Acta N° 044-2012/TI-D
Private Function ValidarMontoAbonar() As String
Dim loRs As COMDConstSistema.DCOMGeneral
Set loRs = New COMDConstSistema.DCOMGeneral
Dim nMontoMin As Double
    If Trim(Right(cboPrograma.Text, 1)) = 2 Or Trim(Right(cboPrograma.Text, 1)) = 3 Or Trim(Right(cboPrograma.Text, 1)) = 4 Then
        If Trim(Right(cboPrograma.Text, 1)) = 3 Or Trim(Right(cboPrograma.Text, 1)) = 4 Then
             'nMontoMin = loRs.GetParametro(2000, 2093)
             nMontoMin = IIf(CInt(Right(cboMoneda.Text, 1)) = gMonedaNacional, nParAumCapMinSol, nParAumCapMinDol) 'JUEZ 20141008 Nuevos parámetros
             If txtMontoAbonar.Text < nMontoMin Then
               ValidarMontoAbonar = "El Monto de Abono debe de ser Igual o Mayor a " & nMontoMin
               Exit Function
             Else
               ValidarMontoAbonar = ""
               Exit Function
             End If
         ElseIf Trim(Right(cboPrograma.Text, 1)) = 2 Then
             'nMontoMin = loRs.GetParametro(2000, 2094)
             nMontoMin = IIf(CInt(Right(cboMoneda.Text, 1)) = gMonedaNacional, nParAumCapMinSol, nParAumCapMinDol) 'JUEZ 20141008 Nuevos parámetros
             If txtMontoAbonar.Text < nMontoMin Then
               ValidarMontoAbonar = "El Monto de Abono debe de ser Igual o Mayor a " & nMontoMin
               Exit Function
             Else
               ValidarMontoAbonar = ""
               Exit Function
             End If
         End If
    End If
'***Fin Agregado por ELRO*************************************

End Function

'**Create By GITU 11-09-2012
Private Sub IniciaComboConvDep(ByVal pnTipoRol As Integer)
Dim lRegPers As New ADODB.Recordset
Dim oPers As COMDPersona.DCOMRoles

    Set oPers = New COMDPersona.DCOMRoles
    Set lRegPers = oPers.CargaPersonas(pnTipoRol)
    Set oPers = Nothing

    If Not lRegPers.BOF And Not lRegPers.EOF Then
        Do While Not lRegPers.EOF
            If lRegPers("PersEstado") = "ACTIVO" Then
                cboInstConvDep.AddItem lRegPers("cPersNombre") & Space(100) & lRegPers("cPersCod")
            End If
            lRegPers.MoveNext
        Loop
        cboInstConvDep.ListIndex = 0
    End If
    lRegPers.Close
    Set lRegPers = Nothing
 End Sub
 
Private Function ValidaInstConv(ByVal psCodPers As String) As Boolean
    Dim oPers As COMDPersona.DCOMRoles

    Set oPers = New COMDPersona.DCOMRoles
    If oPers.ExistePersonaRol(psCodPers, 9) Then
        ValidaInstConv = True
    Else
        ValidaInstConv = False
    End If
    Set oPers = Nothing
End Function

' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES" ***

Private Function prepararRegla() As String
    
    Dim i As Integer
    Dim strCadena As String
    
    For i = 1 To grdReglas.Rows - 1
        If i = 1 Then
            strCadena = grdReglas.TextMatrix(i, 1)
        Else
            strCadena = strCadena & "-" & grdReglas.TextMatrix(i, 1)
        End If
    Next
    
    If Trim(strCadena) = "" Then
        If intPunteroPJ_NA = 0 Then
            strCadena = "A"
        End If
    End If
        
    prepararRegla = strCadena
    
End Function


Private Function prepararGrupoPersona() As String
    
    Dim i As Integer
    Dim strGrupo As String
    
    For i = 1 To grdCliente.Rows - 1
        If i = 1 Then
            strGrupo = grdCliente.TextMatrix(i, 9)
        Else
            strGrupo = strGrupo & "-" & grdCliente.TextMatrix(i, 9)
        End If
    Next
    prepararGrupoPersona = strGrupo
End Function

Private Function validaExistenciaReglas() As Boolean
    'Verifica que los participantes de la cuenta sean parte de  algun grupo
    'Y que no exista una regla demas que no sea parte de algun participante
    
    If intPunteroPJ_NA <> 0 Then
        validaExistenciaReglas = True
        Exit Function
    End If
    Dim i, J, nContar, nContarAsociados As Integer
    Dim sReglas() As String
    Dim sLetra As Variant
    Dim lbValida As Boolean
    lbValida = True
    For i = 1 To grdCliente.Rows - 1
        nContar = 0
        If Trim(grdCliente.TextMatrix(i, 9)) <> "AP" Then
            nContarAsociados = nContarAsociados + 1
            For J = 1 To grdReglas.Rows - 1
                sReglas = Split(grdReglas.TextMatrix(J, 1), "+")
                For Each sLetra In sReglas
                    If Trim(grdCliente.TextMatrix(i, 9)) = sLetra Then
                        nContar = nContar + 1
                    End If
                Next
            Next
            If nContar = 0 Then
                lbValida = False
            End If
        End If
    Next
    nContar = 0
    If nContarAsociados = 0 Then
        lbValida = False
    End If
    For i = 1 To grdReglas.Rows - 1
        nContar = 0
        sReglas = Split(grdReglas.TextMatrix(i, 1), "+")
        For Each sLetra In sReglas
            nContar = 0
            For J = 1 To grdCliente.Rows - 1
                If Trim(grdCliente.TextMatrix(J, 9)) <> "AP" Then
                    nContarAsociados = nContarAsociados + 1
                    If sLetra = Trim(grdCliente.TextMatrix(J, 9)) Then
                        nContar = nContar + 1
                    End If
                End If
            Next
            If nContar = 0 Then
                lbValida = False
            End If
        Next
    Next
    If nContarAsociados = 0 Then
        lbValida = False
    End If
    
    validaExistenciaReglas = lbValida
     
End Function

Private Sub cmdAgregarRegla_Click()

    Dim x As Integer
    Dim i As Integer
    Dim strRegla As String
    'Verificando que cada interviniente en la apertura de la cuenta tenga una relacion: Titular, apoderado
    For i = 1 To grdCliente.Rows - 1
        If Me.grdCliente.TextMatrix(i, 3) = "" Then
            MsgBox "Seleccione la relación de cada persona", vbInformation, "Aviso"
            Exit Sub
        End If
    Next
    i = 0
    'Verifica que las letras marcadas en el control listchek, esten dentro del 'Grid cliente'
    For x = 0 To lsLetras.ListCount - 1
        If lsLetras.Selected(x) = True Then
          lsLetras.Selected(x) = False
            If existeLetraEnSocio(UCase(Chr(65 + x))) = True Then
                If strRegla = "" Then
                    strRegla = UCase(Chr(65 + x))
                Else
                    strRegla = strRegla & "+" & UCase(Chr(65 + x))
                End If
            Else
                MsgBox "Verificar los grupos asignados a cada persona", vbExclamation, "Mensaje"
                Exit Sub
            End If
            i = i + 1
        End If
    Next
    If i = 0 Then
        MsgBox "Debe seleccionar uno o mas grupos antes de presionar este boton", vbExclamation, "Mensaje"
        Exit Sub
    End If
    If existeRegla(strRegla) Then
        MsgBox "Regla ya existe", vbExclamation, "Mensaje"
        Exit Sub
    End If
    grdReglas.AdicionaFila
    grdReglas.SetFocus
    grdReglas.TextMatrix(grdReglas.row, 1) = strRegla
    seleccionarTipoCuentaXregla
End Sub

Private Sub conReglas()
        
    fraReglasPorderes.Visible = True
    grdCliente.Width = 6615
    grdCliente.Height = 1485
    grdCliente.ColWidth(1) = 1300
    grdCliente.ColWidth(2) = 2900
    grdCliente.ColWidth(3) = 1100
    grdCliente.ColWidth(9) = 700
    limpiarReglas
             
End Sub

Private Sub sinReglas()
    
    fraReglasPorderes.Visible = False
    grdCliente.Width = 8925
    grdCliente.Height = 1485
    grdCliente.ColWidth(1) = 1700
    grdCliente.ColWidth(2) = 3500
    grdCliente.ColWidth(3) = 1500
    grdCliente.ColWidth(9) = 1000
    limpiarReglas
End Sub

Private Sub limpiarReglas()
    Dim i As Integer
    For i = 1 To grdReglas.Rows - 1
        grdReglas.EliminaFila grdReglas.row
    Next
    For i = 0 To lsLetras.ListCount - 1
        lsLetras.Selected(i) = False
    Next
End Sub

Private Sub seleccionarTipoCuentaXregla()
    
    Dim nContar, nFirmantes, nTem, x, Y, i, J As Integer
    Dim sReglas(), sGruposPersonas() As String
    Dim sLetra, sValor As Variant
    Dim sGrupo As String
    Dim lbEsMancomunada, lbRepiteGrupo As Boolean
    lbEsMancomunada = True
    lbRepiteGrupo = True
    For i = 1 To grdCliente.Rows - 1
        If Trim(grdCliente.TextMatrix(i, 1)) <> "" Then
            nTem = Val(Trim(Right(grdCliente.TextMatrix(i, 3), 3)))
            If intPunteroPJ_NA = 0 Then
                'If nTem = 10 Or nTem = 11 Or nTem = 12 Then
                If nTem = 10 Or nTem = 12 Then
                    nFirmantes = nFirmantes + 1
                    ReDim Preserve sGruposPersonas(J)
                    sGruposPersonas(J) = grdCliente.TextMatrix(i, 9)
                    J = J + 1
                End If
            Else
                'If nTem = 11 Or nTem = 12 Then
                If nTem = 12 Then
                    nFirmantes = nFirmantes + 1
                    ReDim Preserve sGruposPersonas(J)
                    sGruposPersonas(J) = grdCliente.TextMatrix(i, 9)
                    J = J + 1
                End If
            End If
        End If
    Next
    ' Solo intervienen personas naturales
    If nFirmantes <= 1 Then
        cboTipoCuenta.ListIndex = 0
        Exit Sub
    Else
        nTem = 0
        For Each sLetra In sGruposPersonas
            nTem = 0
            For Each sValor In sGruposPersonas
                If CStr(sLetra) = CStr(sValor) Then
                    nTem = nTem + 1
                End If
            Next
            If nTem <= 1 Then
                lbRepiteGrupo = False
            Else
                lbRepiteGrupo = True
                Exit For
            End If
        Next
        For i = 1 To grdReglas.Rows - 1
            sGrupo = Trim(grdReglas.TextMatrix(i, 1))
            For Each sLetra In sGruposPersonas
                If InStr(sGrupo, CStr(sLetra)) = 0 Then
                    lbEsMancomunada = False
                    Exit For
                Else
                    lbEsMancomunada = True
                End If
            Next
            If Not lbEsMancomunada Then
                Exit For
            End If
        Next
        'Mancomunado
        If lbEsMancomunada And lbRepiteGrupo = False Then
            cboTipoCuenta.ListIndex = 1
        'Indistinta
        Else
            cboTipoCuenta.ListIndex = 2
        End If
    End If
End Sub

Private Function existeLetraEnSocio(letra As String) As Boolean
    Dim i As Integer
    Dim blPuntero As Boolean
    blPuntero = False
    For i = 0 To grdCliente.Rows - 1
        If Trim(UCase(letra)) = Trim(UCase(grdCliente.TextMatrix(i, 9))) Then
            blPuntero = True
        End If
    Next
    existeLetraEnSocio = blPuntero
End Function

Private Function existeRegla(strRegla As String) As Boolean
    
    Dim blReglaExiste As Boolean
    Dim i As Integer
    blReglaExiste = False
    For i = 0 To grdReglas.Rows - 1
        If Trim(grdReglas.TextMatrix(i, 1)) = Trim(strRegla) Then
            blReglaExiste = True
         End If
    Next
    existeRegla = blReglaExiste
    
End Function

' *** END RIRO ***
'JUEZ 20131212 ****************************************************
Private Sub ValidaCargoCta()
Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
Dim rs As ADODB.Recordset
Dim sMsg As String
    
    If ValidaFlexVacio Then
        MsgBox "Ingrese un Cliente", vbInformation, "Aviso"
        cmdAgregar.SetFocus
        LimpiaControlesCargoCta
        Exit Sub
    End If
    
    Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
    sMsg = oNCapMov.ValidaCuentaOperacion(txtCuentaCargo.NroCuenta)
    If sMsg <> "" Then
        MsgBox sMsg, vbInformation, "Aviso"
        txtCuentaCargo.SetFocusCuenta
        LimpiaControlesCargoCta
        Exit Sub
    End If
    If Trim(Right(cboMoneda, 2)) <> Mid(txtCuentaCargo.NroCuenta, 9, 1) Then
        MsgBox "Cuenta debe ser de la misma moneda que la apertura", vbInformation, "Aviso"
        txtCuentaCargo.SetFocusCuenta
        LimpiaControlesCargoCta
        Exit Sub
    End If
    
    Set rs = oDCapGen.GetDatosCuentaAho(txtCuentaCargo.NroCuenta)
    fnTpoCtaCargo = rs("nPrdCtaTpo")
    
    If Trim(Right(cboTipoCuenta, 2)) <> fnTpoCtaCargo Then
        MsgBox "Cuenta debe ser del mismo tipo de cuenta de la apertura", vbInformation, "Aviso"
'        txtCuentaCargo.SetFocusCuenta
        LimpiaControlesCargoCta
        Exit Sub
    End If
    Set rs = Nothing
    
    Set rsRelPersCtaCargo = oDCapGen.GetPersonaCuenta(txtCuentaCargo.NroCuenta)
    Set oDCapGen = Nothing
    If Not ValidaRelPersonasCtaCargo Then
        MsgBox "La personas y relaciones de la cuenta a debitar deben ser las mismas que las de la apertura", vbInformation, "Aviso"
        'txtCuentaCargo.SetFocusCuenta
        LimpiaControlesCargoCta
        Exit Sub
    End If
    rsRelPersCtaCargo.MoveFirst
    lblTitularCargoCta.Caption = UCase(PstaNombre(rsRelPersCtaCargo("Nombre")))
    
End Sub

Private Sub LimpiaControlesCargoCta()
    txtCuentaCargo.Age = gsCodAge
    txtCuentaCargo.Cuenta = ""
    lblTitularCargoCta.Caption = ""
    Set rsRelPersCtaCargo = Nothing
    fnTpoCtaCargo = 0
End Sub

Private Function ValidaRelPersonasCtaCargo() As Boolean
    Dim bExisteRelPers As Boolean
    Dim i As Integer
    
    ValidaRelPersonasCtaCargo = False
    
    rsRelPersCtaCargo.MoveFirst
    Do While Not rsRelPersCtaCargo.EOF
        bExisteRelPers = False
        For i = 1 To grdCliente.Rows - 1
            If grdCliente.TextMatrix(i, 1) = rsRelPersCtaCargo("cPersCod") And Trim(Right(grdCliente.TextMatrix(i, 3), 2)) = Trim(Right(rsRelPersCtaCargo("Relacion"), 2)) Then
                bExisteRelPers = True
                Exit For
            End If
        Next i
        If Not bExisteRelPers Then Exit Function
        rsRelPersCtaCargo.MoveNext
    Loop
    
    ValidaRelPersonasCtaCargo = True
End Function

Private Function VerificarAutorizacion() As Boolean
Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
Dim oCapAutN  As COMNCaptaGenerales.NCOMCaptAutorizacion
Dim oPers As COMDPersona.UCOMAcceso
Dim rs As New ADODB.Recordset
Dim lnMonTopD As Double
Dim lnMonTopS As Double
Dim lsmensaje As String
Dim gsGrupo As String
Dim sCuenta As String, sNivel As String
Dim lbEstadoApr As Boolean
Dim nMonto As Double
Dim nmoneda As Moneda

sCuenta = txtCuentaCargo.NroCuenta
nMonto = txtMonto.value
nmoneda = CLng(Mid(sCuenta, 9, 1))
'Obtiene los grupos al cual pertenece el usuario
Set oPers = New COMDPersona.UCOMAcceso
    gsGrupo = oPers.CargaUsuarioGrupo(gsCodUser, gsDominio)
Set oPers = Nothing
 
'Verificar Montos
Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
    'Set rs = ocapaut.ObtenerMontoTopNivAutRetCan(gsGrupo, "3", gsCodAge)
    Set rs = oCapAut.ObtenerMontoTopNivAutRetCan(gsGrupo, "3", gsCodAge, gsCodPersUser) 'RIRO ERS159
Set oCapAut = Nothing
 
If Not (rs.EOF And rs.BOF) Then
    lnMonTopD = rs("nTopDol")
    lnMonTopS = rs("nTopSol")
    sNivel = rs("cNivCod")
Else
    MsgBox "Usuario no Autorizado para realizar Operacion", vbInformation, "Aviso"
    VerificarAutorizacion = False
    Exit Function
End If

If nmoneda = gMonedaNacional Then
    If nMonto <= lnMonTopS Then
        VerificarAutorizacion = True
        Exit Function
    End If
Else
    If nMonto <= lnMonTopD Then
        VerificarAutorizacion = True
        Exit Function
    End If
End If
   
Set oCapAutN = New COMNCaptaGenerales.NCOMCaptAutorizacion
If sMovNroAut = "" Then 'Si es nueva, registra
    oCapAutN.NuevaSolicitudAutorizacion sCuenta, "3", nMonto, gdFecSis, gsCodAge, gsCodUser, nmoneda, gOpeAutorizacionCargoCuenta, sNivel, sMovNroAut
    MsgBox "Solicitud Registrada, comunique a su AdmInistrador para la Aprobación..." & Chr$(10) & _
        " No salir de esta operación mientras se realice el proceso..." & Chr$(10) & _
        " Porque sino se procedera a grabar otra Solicitud...", vbInformation, "Aviso"
    VerificarAutorizacion = False
Else
    'Valida el estado de la Solicitud
    If Not oCapAutN.VerificarAutorizacion(sCuenta, "3", nMonto, sMovNroAut, lsmensaje) Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        VerificarAutorizacion = False
    Else
        VerificarAutorizacion = True
    End If
End If
Set oCapAutN = Nothing
End Function
'END JUEZ *********************************************************
'EJVG20140203 ***
Private Sub SetDatosCheque(Optional ByVal psNroDoc As String = "", Optional ByVal psNombreIFi As String = "", Optional ByVal psDetalle As String = "", Optional ByVal psGlosa As String = "", Optional ByVal pnMonto As Currency = 0#)
    Dim oPersona As New DCOMPersona
    Dim rsPersona As New ADODB.Recordset
    Dim row As Integer
    
    lblNroDoc.Caption = psNroDoc
    lblNombreIF.Caption = psNombreIFi
    txtGlosa.Text = psGlosa
    
    FormateaFlex grdCliente
    If psDetalle <> "" Then
        Set rsPersona = oPersona.RecuperaPersonaxCheque(psDetalle)
        Do While Not rsPersona.EOF
            grdCliente.AdicionaFila
            row = grdCliente.row
            grdCliente.TextMatrix(row, 1) = rsPersona!cPersCod
            grdCliente.TextMatrix(row, 2) = rsPersona!cPersNombre
            grdCliente.TextMatrix(row, 4) = rsPersona!nPersPersoneria
            rsPersona.MoveNext
        Loop
        grdCliente.Col = 5
    End If
    txtMonto.Text = Format(pnMonto, gsFormatoNumeroView)
    If txtCuenta.Prod = "234" Then
        vnMontoDOC = CDbl(txtMonto.Text)
    End If
    
    txtMonto_Change

    If psDetalle <> "" Then
        txtGlosa.SetFocus
    End If
    txtGlosa.Locked = True
    txtMonto.Enabled = False
    lblTotal.Caption = Format(txtMonto.value + CCur(Me.LblItf.Caption), gsFormatoNumeroView)

    Set rsPersona = Nothing
    Set oPersona = Nothing
End Sub
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
'JUEZ 20141008 *************************************************
Private Function ValidaMontoMinimoApertura() As Boolean
Dim nMontoMinProd As Double
ValidaMontoMinimoApertura = True
nMontoMinProd = IIf(CInt(Right(cboMoneda.Text, 1)) = gMonedaNacional, nParMontoMinSol, nParMontoMinDol)
If nProducto = gCapAhorros Or nProducto = gCapPlazoFijo Then
    If txtMonto.value < nMontoMinProd Then
        MsgBox "El monto mínimo de Apertura no debe ser menor de " & Format(nMontoMinProd, "#,##0.00"), vbInformation, "Aviso"
        txtMonto.Text = Format(nMontoMinProd, "#,##0.00")
        txtMonto.SetFocus
        ValidaMontoMinimoApertura = False
    End If
End If
End Function
Private Function ValidarPlazoPF() As Boolean
ValidarPlazoPF = True
If CInt(txtPlazo.Text) < nParPlazoMin Or CInt(txtPlazo.Text) > nParPlazoMax Then
    If nParPlazoMin = nParPlazoMax Then
        MsgBox "El plazo debe ser " & nParPlazoMin & " días", vbInformation, "Aviso"
    Else
        MsgBox "El plazo debe estar entre " & nParPlazoMin & " y " & nParPlazoMax & " días", vbInformation, "Aviso"
    End If
    txtPlazo.Text = nParPlazoMin
    ValidarPlazoPF = False
    txtPlazo.SetFocus
End If
End Function
Private Function ValidarMedioRetiroPF() As Boolean
Dim nFormaRet As Integer
ValidarMedioRetiroPF = True
If cboPrograma.ListIndex = 0 Then
    nFormaRet = CLng(Trim(Right(cboFormaRetiro.Text, 4)))
Else
    nFormaRet = gCapPFFormRetFinalPlazo
End If
If nFormaRet = gCapPFFormRetMensual And Not bParFormaRetMensual Then
    MsgBox "El producto no permite la forma de retiro seleccionada", vbInformation, "Aviso"
    cboFormaRetiro.SetFocus
    ValidarMedioRetiroPF = False
End If
If nFormaRet = gCapPFFormRetFinalPlazo And Not bParFormaRetFinPlazo Then
    MsgBox "El producto no permite la forma de retiro seleccionada", vbInformation, "Aviso"
    cboFormaRetiro.SetFocus
    ValidarMedioRetiroPF = False
End If
If nFormaRet = gCapPFFormRetAdelantado And Not bParFormaRetIniPlazo Then
    MsgBox "El producto no permite la forma de retiro seleccionada", vbInformation, "Aviso"
    cboFormaRetiro.SetFocus
    ValidarMedioRetiroPF = False
End If
End Function
'END JUEZ ******************************************************
