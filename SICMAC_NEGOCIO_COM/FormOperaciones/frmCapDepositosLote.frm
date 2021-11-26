VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCapDepositosLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12600
   Icon            =   "frmCapDepositosLote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFormaPago 
      Height          =   600
      Left            =   135
      TabIndex        =   62
      Top             =   5640
      Visible         =   0   'False
      Width           =   7785
      Begin VB.ComboBox CmbForPag 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   200
         Width           =   1785
      End
      Begin SICMACT.ActXCodCta txtCuentaCargo 
         Height          =   375
         Left            =   3120
         TabIndex        =   38
         Top             =   200
         Visible         =   0   'False
         Width           =   3630
         _extentx        =   6403
         _extenty        =   661
         texto           =   "Cuenta N°:"
         enabledcta      =   -1
         enabledage      =   -1
      End
      Begin VB.Label lblNroDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   3105
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   200
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblFormaPago 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   250
         Width           =   855
      End
   End
   Begin VB.Frame FRGlosa 
      Caption         =   "Glosa"
      Height          =   2535
      Left            =   4800
      TabIndex        =   30
      Top             =   5640
      Width           =   4215
      Begin VB.TextBox txtGlosa 
         Height          =   2175
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame FRCheque 
      Caption         =   "Cheque"
      Height          =   2535
      Left            =   120
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   4455
      Begin VB.ComboBox cboCheque 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   240
         Width           =   1575
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
         Picture         =   "frmCapDepositosLote.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lblEtiMonChe 
         AutoSize        =   -1  'True
         Caption         =   "Monto Cheque"
         Height          =   195
         Left            =   960
         TabIndex        =   63
         Top             =   1330
         Width           =   1050
      End
      Begin VB.Label lblSimChe 
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
         Left            =   2160
         TabIndex        =   33
         Top             =   1330
         Width           =   300
      End
      Begin VB.Label lblMonChe 
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
         Left            =   2520
         TabIndex        =   32
         Top             =   1330
         Width           =   1755
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nro Doc"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   465
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
         TabIndex        =   27
         Top             =   960
         Width           =   3465
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
         TabIndex        =   26
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame FRTransferencia 
      Caption         =   "Transferencia"
      Height          =   2535
      Left            =   120
      TabIndex        =   11
      Top             =   5640
      Visible         =   0   'False
      Width           =   4455
      Begin VB.ComboBox cboTransferMoneda 
         Height          =   315
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdTranfer 
         Height          =   350
         Left            =   2490
         Picture         =   "frmCapDepositosLote.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   555
         Width           =   475
      End
      Begin VB.TextBox txtTransferGlosa 
         Appearance      =   0  'Flat
         Height          =   720
         Left            =   825
         MaxLength       =   255
         TabIndex        =   12
         Top             =   1365
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
         Height          =   345
         Left            =   825
         TabIndex        =   23
         Top             =   975
         Width           =   3465
      End
      Begin VB.Label lbltransferN 
         AutoSize        =   -1  'True
         Caption         =   "Nro Doc"
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   675
         Width           =   600
      End
      Begin VB.Label lbltransferBcol 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   300
         TabIndex        =   21
         Top             =   1050
         Width           =   465
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
         Left            =   825
         TabIndex        =   20
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblTransferMoneda 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lblTransferGlosa 
         AutoSize        =   -1  'True
         Caption         =   "Glosa"
         Height          =   195
         Left            =   330
         TabIndex        =   18
         Top             =   1485
         Width           =   405
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
         Left            =   2730
         TabIndex        =   17
         Top             =   2115
         Width           =   1545
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
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2370
         TabIndex        =   16
         Top             =   2145
         Width           =   300
      End
      Begin VB.Label lblEtiMonTra 
         AutoSize        =   -1  'True
         Caption         =   "Monto Transacción"
         Height          =   195
         Left            =   930
         TabIndex        =   15
         Top             =   2175
         Width           =   1380
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   11040
      TabIndex        =   10
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   9600
      TabIndex        =   9
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Frame FRTotal 
      Caption         =   "Total"
      Height          =   2535
      Left            =   9120
      TabIndex        =   2
      Top             =   5640
      Width           =   3375
      Begin VB.Frame FREfectivo 
         Height          =   495
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
         Begin VB.ComboBox cboEfectivo 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   180
            Width           =   585
         End
      End
      Begin VB.Frame FRPeriodoCTS 
         Height          =   975
         Left            =   120
         TabIndex        =   64
         Top             =   720
         Visible         =   0   'False
         Width           =   3135
         Begin VB.ComboBox cboPeriodo 
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   240
            Width           =   2295
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
            TabIndex        =   39
            Top             =   600
            Width           =   795
         End
         Begin VB.Label lblDisponible 
            AutoSize        =   -1  'True
            Caption         =   "Dispon.del Excedente (%) :"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   630
            Width           =   1905
         End
         Begin VB.Label lblPeriodo 
            Caption         =   "Periodo"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Label lblTotalMN 
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
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1560
         TabIndex        =   49
         Top             =   1800
         Width           =   1665
      End
      Begin VB.Label Label14 
         Caption         =   "Total S/."
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
         Height          =   255
         Left            =   600
         TabIndex        =   48
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblTotalME 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1560
         TabIndex        =   5
         Top             =   2160
         Width           =   1665
      End
      Begin VB.Label Label5 
         Caption         =   "Total $"
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
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   2160
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   12375
      Begin VB.TextBox txtRuta 
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   4200
         Width           =   4935
      End
      Begin VB.CommandButton cmdRuta 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   59
         Top             =   4200
         Width           =   495
      End
      Begin VB.CommandButton cmdFormato 
         Caption         =   "&Formato"
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   10920
         TabIndex        =   8
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   9480
         TabIndex        =   7
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "&Cargar"
         Height          =   375
         Left            =   7680
         TabIndex        =   6
         Top             =   4200
         Width           =   1335
      End
      Begin SICMACT.FlexEdit FEClientes 
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   12135
         _extentx        =   21405
         _extenty        =   5741
         cols0           =   9
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Codigo-Nombre-CTS-Monto S/.-Monto $-ITF S/.-ITF $-Salto"
         encabezadosanchos=   "500-1500-4000-2000-1200-1200-600-600-0"
         font            =   "frmCapDepositosLote.frx":0B8E
         font            =   "frmCapDepositosLote.frx":0BBA
         font            =   "frmCapDepositosLote.frx":0BE6
         font            =   "frmCapDepositosLote.frx":0C12
         font            =   "frmCapDepositosLote.frx":0C3E
         fontfixed       =   "frmCapDepositosLote.frx":0C6A
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         tipobusqueda    =   3
         columnasaeditar =   "X-1-X-3-4-5-X-X-X"
         textstylefixed  =   4
         listacontroles  =   "0-1-0-3-0-0-0-0-0"
         encabezadosalineacion=   "C-R-L-R-R-R-R-R-C"
         formatosedit    =   "0-0-0-0-2-2-2-2-0"
         textarray0      =   "#"
         lbbuscaduplicadotext=   -1
         colwidth0       =   495
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin MSComDlg.CommonDialog CdlgFile 
         Left            =   9000
         Top             =   4200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label8 
         Caption         =   "Archivo"
         Height          =   255
         Left            =   1440
         TabIndex        =   61
         Top             =   4275
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "SubTotal S/."
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
         Height          =   255
         Left            =   6720
         TabIndex        =   57
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label lblSubTotalITFME 
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
         Left            =   11160
         TabIndex        =   47
         Top             =   3720
         Width           =   555
      End
      Begin VB.Label lblSubTotalITFMN 
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
         Left            =   10560
         TabIndex        =   46
         Top             =   3720
         Width           =   555
      End
      Begin VB.Label lblSubTotalME 
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
         Left            =   9360
         TabIndex        =   45
         Top             =   3720
         Width           =   1155
      End
      Begin VB.Label lblSubTotalMN 
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
         Left            =   8160
         TabIndex        =   44
         Top             =   3720
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   16815
      Begin VB.Label lblTCV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   43
         Top             =   555
         Width           =   735
      End
      Begin VB.Label lblTCC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   42
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "TCV"
         Height          =   285
         Left            =   120
         TabIndex        =   41
         Top             =   555
         Width           =   390
      End
      Begin VB.Label lblTTCC 
         Caption         =   "TCC"
         Height          =   285
         Left            =   120
         TabIndex        =   40
         Top             =   255
         Width           =   390
      End
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   375
      Left            =   1440
      TabIndex        =   55
      Top             =   8280
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblProcesando 
      Caption         =   "Procesando"
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
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   8380
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmCapDepositosLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************************
'***Nombre      : frmCapDepositosLote
'***Descripción : Formulario para Depósito de CTS en Lote.
'***Creación    : ELRO el 20121019, según OYP-RFC101-2012
'***************************************************************************************************

Private fnProducto As Producto
Private fnOpeCod As CaptacOperacion
Private fsDescOperacion As String
Private fnMoneda As COMDConstantes.Moneda

Private fnMovNroTransfer As Long
Private fnTransferSaldo As Currency
Private fnMovNroRVD As Long
Private fsPersCodTransfer As String

Private fnValorChq As Currency
Public fsCodIF As String
Public fdFechaValorizacion As Date
Private fnNroClientesTransf As Integer 'EJVG20130916
Dim oDocRec As UDocRec 'EJVG20140408
'JUEZ 20141014 Nuevos parámetros **************
Dim bValidaCantDep As Boolean
Dim nParCantDepAnio As Integer
Dim nParDiasVerifRegSueldo As Integer
Dim loVistoElectronico As frmVistoElectronico
Dim lbVistoVal As Boolean
Dim nCantDepCta As Integer
'END JUEZ *************************************
Private nMontoVoucher As Currency 'CTI4 ERS0112020
Dim nMovNroRVD As Long, nMovNroRVDPend As Long 'CTI4 ERS0112020
Dim sNumTarj As String 'CTI4 ERS0112020
Dim loVistoElectronico2 As frmVistoElectronico 'CTI4 ERS0112020
Dim nRespuesta As Integer 'CTI4 ERS0112020

Public Sub iniciarFormulario(ByVal pnProducto As Producto, ByVal pnOpeCod As CaptacOperacion, _
                             Optional psDescOperacion As String = "")
fnProducto = pnProducto
fnOpeCod = pnOpeCod
fsDescOperacion = psDescOperacion

Select Case fnProducto
    Case gCapCTS
        FRPeriodoCTS.Visible = True
        Me.Caption = "Captaciones - CTS - " & psDescOperacion
        If fnOpeCod = gCTSDepLotEfec Then
            FREfectivo.Visible = True
            fraFormaPago.Visible = True 'CTI4 ERS0112020
            CmbForPag.Enabled = True 'CTI4 ERS0112020
        ElseIf fnOpeCod = gCTSDepLotChq Then
            FRCheque.Visible = True
            FRGlosa.Visible = False
        ElseIf fnOpeCod = gCTSDepLotTransf Then
            FRTransferencia.Visible = True
            FRGlosa.Visible = False
        End If
End Select

'JUEZ 20141014 Verificar si operación valida cantidad de depositos en mes ****
Dim oCapDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
bValidaCantDep = oCapDef.ValidaCantOperaciones(fnOpeCod, fnProducto, gCapMovDeposito)
Set oCapDef = Nothing
'END JUEZ ********************************************************************
'FRHU 20141222: Por el momento no se mostraran estos controles
'If fnOpeCod = gCTSDepLotEfec Or fnOpeCod = gCTSDepLotChq Or fnOpeCod = gCTSDepLotTransf Then
'    Label8.Visible = False
'    txtRuta.Visible = False
'    cmdRuta.Visible = False
'    cmdCargar.Visible = False
'End If
'FIN FRHU 20141222
Me.Show 1
End Sub

Private Function ValidarDatos() As Boolean
Dim oDCOMCaptaMovimiento As DCOMCaptaMovimiento
Dim oDCOMPersonas As COMDPersona.DCOMPersonas

Dim rsUltActSueldo As ADODB.Recordset
Dim rsPersVerifica As Recordset
Dim ldFecha As Date
Dim i, lnFila As Integer

lnFila = FEClientes.Rows

If CCur(lblTCC) <= 0 Or CCur(lblTCV) <= 0 Then
    MsgBox "No se a definido el Tipo de Cambio", vbCritical, "!AViso¡"
End If

For i = 1 To lnFila - 1
    Set oDCOMPersonas = New COMDPersona.DCOMPersonas
    Set oDCOMCaptaMovimiento = New DCOMCaptaMovimiento
    Set rsUltActSueldo = New ADODB.Recordset
    Set rsPersVerifica = New Recordset
    
    '***Verificar el Titular****
    If Trim(FEClientes.TextMatrix(i, 1)) = "" Then
        MsgBox "Debe ingresar a una Persona.", vbInformation, "!Aviso¡"
        FEClientes.row = i
        FEClientes.Col = 1
        FEClientes.SetFocus
        ValidarDatos = False
        Exit Function
    End If
    '***Fin Verificar el Titular
    
    '***Verificar la Cuenta****
    If Trim(FEClientes.TextMatrix(i, 3)) = "" Then
        MsgBox "Debe ingresar la Cuenta.", vbInformation, "!Aviso¡"
        FEClientes.row = i
        FEClientes.Col = 3
        FEClientes.SetFocus
        ValidarDatos = False
        Exit Function
    End If
    '***Fin Verificar la Cuenta
    
    '***Verificar el monto a depositar****
    If fnMoneda = gMonedaNacional Then
        If Trim(FEClientes.TextMatrix(i, 4)) = "" Then
            MsgBox "Debe ingresar el monto a depositar.", vbInformation, "!Aviso¡"
            FEClientes.row = i
            FEClientes.Col = 4
            FEClientes.SetFocus
            ValidarDatos = False
            Exit Function
        End If
   
        If CCur(Trim(FEClientes.TextMatrix(i, 4))) <= 0 Then
            MsgBox "El monto a depositar no debe ser cero o negativo .", vbInformation, "!Aviso¡"
            FEClientes.row = i
            FEClientes.Col = 4
            FEClientes.SetFocus
            ValidarDatos = False
            Exit Function
        End If
    Else
        If Trim(FEClientes.TextMatrix(i, 5)) = "" Then
            MsgBox "Debe ingresar el monto a depositar.", vbInformation, "!Aviso¡"
            FEClientes.row = i
            FEClientes.Col = 5
            FEClientes.SetFocus
            ValidarDatos = False
            Exit Function
        End If
        
        If CCur(Trim(FEClientes.TextMatrix(i, 5))) <= 0 Then
            MsgBox "El monto a depositar no debe ser cero o negativo .", vbInformation, "!Aviso¡"
            FEClientes.row = i
            FEClientes.Col = 5
            FEClientes.SetFocus
            ValidarDatos = False
            Exit Function
        End If
    End If
    '***Fin Verificar el monto a depositar
    
    
    '***Verifica la Actualización de las 06 Últimas Remuneaciones Brutas****
    Set rsUltActSueldo = oDCOMCaptaMovimiento.ObtenerFecUltimaActSueldosCTS(Trim(FEClientes.TextMatrix(i, 3)))
    If rsUltActSueldo.BOF Or rsUltActSueldo.EOF Then
        MsgBox "No se encontraron registros de los 6 Últimos Sueldos del Titular del Nro. de Cuenta " & Trim(FEClientes.TextMatrix(i, 3)) & Chr(10) & _
               "Debe registrar el Total de los 6 Últimos Sueldos para proceder."
        ValidarDatos = False
        cmdSalir.SetFocus
        Exit Function
    Else
        ldFecha = rsUltActSueldo!FechaAct
        'If DateDiff("d", ldFecha, gdFecSis) > 30 Then
        If DateDiff("d", ldFecha, gdFecSis) > nParDiasVerifRegSueldo Then 'JUEZ 20141014
            MsgBox "La última actualización ha caducado." & Chr(10) & _
                   "Favor actualice su registro de los 6 Últimos Sueldos del Nro. de Cuenta " & Trim(FEClientes.TextMatrix(i, 3))
            ValidarDatos = False
             cmdSalir.SetFocus
            Exit Function
        End If
    End If
    '***Fin Verifica la Actualización de las 06 Últimas Remuneaciones Brutas
    
    '***Verifica si las personas cuentan con ocupacion e ingreso promedio****
    Set rsPersVerifica = oDCOMPersonas.ObtenerDatosPersona(Trim(FEClientes.TextMatrix(i, 1)))
    If Not (rsPersVerifica.BOF And rsPersVerifica.EOF) Then
        If rsPersVerifica!nPersIngresoProm = 0 Or rsPersVerifica!cActiGiro1 = "" Then
            If MsgBox("Necesita Registrar la Ocupación e Ingreso Promedio de: " + Trim(FEClientes.TextMatrix(i, 2)), vbYesNo) = vbYes Then
                frmPersOcupIngreProm.Inicio Trim(FEClientes.TextMatrix(i, 1)), Trim(FEClientes.TextMatrix(i, 2)), rsPersVerifica!cActiGiro1, rsPersVerifica!nPersIngresoProm
            End If
        End If
    End If
    '***Fin Verifica si las personas cuentan con ocupacion e ingreso promedio
   
    Set oDCOMPersonas = Nothing
    Set oDCOMCaptaMovimiento = Nothing
    Set rsUltActSueldo = Nothing
    Set rsPersVerifica = Nothing
    
Next i

If fnOpeCod = gCTSDepLotChq Or fnOpeCod = gCTSDepLotTransf Then
    If fnOpeCod = gCTSDepLotChq Then
        If CCur(lblMonChe) < IIf(fnMoneda = gMonedaNacional, (lblTotalMN), (lblTotalME)) Or _
           CCur(lblMonChe) > IIf(fnMoneda = gMonedaNacional, (lblTotalMN), (lblTotalME)) Then
            MsgBox "El monto de la operación no es igual al monto del cheque.", vbInformation, "Aviso"
            ValidarDatos = False
            cmdGuardar.SetFocus
            Exit Function
        End If
        If lblNroDoc = "" Then
            MsgBox "Debe seleccionar un cheque válido para la operación.", vbInformation, "Aviso"
            cmdDocumento.SetFocus
            Exit Function
        End If
        'EJVG20140408 ***
        If oDocRec.fnNroCliLote <> lnFila - 1 Then 'Verifica cantidad cuentas a depositar
            MsgBox "No puede continuar ya que  el Nro. de Depositos (" & lnFila - 1 & ") es diferente al registrado en el cheque de (" & oDocRec.fnNroCliLote & ")", vbInformation, "Aviso"
            ValidarDatos = False
            If cmdGuardar.Visible And cmdGuardar.Enabled Then cmdGuardar.SetFocus
            Exit Function
        End If
        'END EJVG *******
    Else
'        If CCur(lblMonTra) < IIf(fnMoneda = gMonedaNacional, (lblTotalMN), (lblTotalME)) Or _
'           CCur(lblMonTra) > IIf(fnMoneda = gMonedaNacional, (lblTotalMN), (lblTotalME)) Then
'            MsgBox "El monto de la operación no es igual al monto de la tranferencia."
'            ValidarDatos = False
'            cmdGuardar.SetFocus
'            Exit Function
'        End If
        If CCur(lblMonTra) <> IIf(fnMoneda = gMonedaNacional, (lblTotalMN), (lblTotalME)) Then
            MsgBox "El monto de la operación no es igual al monto de la tranferencia."
            ValidarDatos = False
            cmdGuardar.SetFocus
            Exit Function
        End If
        If IIf(fnMoneda = gMonedaNacional, (lblSubTotalMN), (lblSubTotalME)) <> IIf(fnMoneda = gMonedaNacional, (lblTotalMN), (lblTotalME)) Then
            MsgBox "El monto de la operación no es igual al monto de la tranferencia."
            ValidarDatos = False
            cmdGuardar.SetFocus
            Exit Function
        End If
        If lblTrasferND = "" Then
            MsgBox "Debe seleccionar un voucher válido para la operación.", vbInformation, "Aviso"
            cmdDocumento.SetFocus
            Exit Function
        End If
        'EJVG20130917 *** Verifica la Cantidad de Intervinientes sea igual que ingresaron en Voucher
        If fnNroClientesTransf <> lnFila - 1 Then
            MsgBox "La cantidad de Clientes no coincide con los ingresados en el Voucher", vbInformation, "Aviso"
            ValidarDatos = False
            cmdGuardar.SetFocus
            Exit Function
        End If
        'END EJVG *******
    End If
End If

ValidarDatos = True
End Function

Private Sub cargarCTSPeriodo()
Dim oNCOMCaptaGenerales As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsConst As New ADODB.Recordset
Dim sCodigo As String * 2
Set oNCOMCaptaGenerales = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsConst = oNCOMCaptaGenerales.GetCTSPeriodo()
Set oNCOMCaptaGenerales = Nothing
Do While Not rsConst.EOF
    sCodigo = rsConst("nItem")
    cboPeriodo.AddItem sCodigo & space(2) & UCase(rsConst("cDescripcion")) & space(100) & rsConst("nPorcentaje")
    rsConst.MoveNext
Loop
cboPeriodo.ListIndex = 4
End Sub

Private Sub IniciaCombo(ByRef cboConst As ComboBox, ByVal pnCapConst As ConstanteCabecera)
Dim oDCOMGeneral As New COMDConstSistema.DCOMGeneral
Dim rsConst As New ADODB.Recordset
Set rsConst = oDCOMGeneral.GetConstante(pnCapConst, , , "1")
Set oDCOMGeneral = Nothing

Do While Not rsConst.EOF
    cboConst.AddItem rsConst("cDescripcion") & space(100) & rsConst("nConsValor")
    rsConst.MoveNext
Loop

cboConst.ListIndex = 0
End Sub

Private Sub devolverCTSPersona(ByVal pcPersCod As String, Optional ByVal psCtaCod As String = "", Optional ByRef pbExiste As Boolean = False)
Dim oNCOMCaptaGenerales As New COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsCtaCTS As New ADODB.Recordset

Set rsCtaCTS = Nothing
Set rsCtaCTS = oNCOMCaptaGenerales.obtenerListadoCuentasCTS(pcPersCod)

pbExiste = False 'JUEZ 20150922

If RSVacio(rsCtaCTS) Then
    MsgBox "La persona no tiene Cuenta de CTS Activa.", vbInformation, "!Aviso¡"
    FEClientes.TextMatrix(FEClientes.row, 1) = ""
    FEClientes.TextMatrix(FEClientes.row, 2) = ""
    FEClientes.TextMatrix(FEClientes.row, 3) = ""
    FEClientes.TextMatrix(FEClientes.row, 4) = "0.00"
    FEClientes.TextMatrix(FEClientes.row, 5) = "0.00"
    FEClientes.TextMatrix(FEClientes.row, 6) = ""
    Exit Sub
ElseIf Not (rsCtaCTS.BOF And rsCtaCTS.EOF) Then
    If psCtaCod <> "" Then
        Do While Not rsCtaCTS.EOF
            If rsCtaCTS!cCtaCod = psCtaCod Then
                pbExiste = True
                Exit Do
            End If
            rsCtaCTS.MoveNext
        Loop
        rsCtaCTS.MoveFirst
    End If
    
    FEClientes.CargaCombo rsCtaCTS
    FEClientes.TextMatrix(FEClientes.row, 3) = ""
End If

If pbExiste Then
    FEClientes.TextMatrix(FEClientes.row, 3) = psCtaCod
End If

Set rsCtaCTS = Nothing
Set oNCOMCaptaGenerales = Nothing
End Sub


Private Sub ImprimeBoleta(ByVal psBoleta As String, Optional ByVal psMensaje As String = "Boleta Operación")
Dim lnFicSal As Integer
Do
    lnFicSal = FreeFile
    Open sLpt For Output As lnFicSal
    If fnProducto = gCapCTS Then
        psBoleta = psBoleta & oImpresora.gPrnSaltoLinea
    End If
    Print #lnFicSal, psBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    Print #lnFicSal, ""
    Print #lnFicSal, ""
    Close #lnFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & psMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub

Private Sub limpiarCamposChq()
    lblNroDoc = ""
    lblNombreIF = ""
    lblMonChe = "0.00"
    fnValorChq = 0
    fsCodIF = ""
    fdFechaValorizacion = "01/01/1900"
End Sub

Private Sub limpiarCamposTranf()
    lblTrasferND = ""
    lbltransferBco = ""
    txtTransferGlosa = ""
    lblMonTra = "0.00"
    fnMovNroTransfer = 0
    fnTransferSaldo = 0
    fnMovNroRVD = 0
    fsPersCodTransfer = ""
End Sub


Private Sub limpiarCampos()
 LimpiaFlex FEClientes
 txtGlosa = ""
 lblSubTotalMN = "0.00"
 lblSubTotalME = "0.00"
 lblSubTotalITFMN = "0.00"
 lblSubTotalITFME = "0.00"
 lblTotalMN = "0.00"
 lblTotalME = "0.00"
 txtRuta = ""
 
 If FRTransferencia.Visible Then
    cboTransferMoneda.ListIndex = 0
    limpiarCamposTranf
 End If
 If FRCheque.Visible Then
    cboCheque.ListIndex = 0
    limpiarCamposChq
 End If
 If fraFormaPago.Visible Then 'CTI4 ERS0112020
    CmbForPag.ListIndex = 0
    txtCuentaCargo.NroCuenta = ""
    LblNumDoc.Caption = ""
    sNumTarj = ""
 End If
End Sub

Private Sub cargarTipoCambio()
    Dim oDCOMTipoCambioEsp As New COMDConstSistema.DCOMTipoCambioEsp
    Dim rsTipoCambio As New ADODB.Recordset
    
    Set rsTipoCambio = oDCOMTipoCambioEsp.GetTipoCambioCV(0)
        
    If Not (rsTipoCambio.BOF And rsTipoCambio.EOF) Then
        Do While Not rsTipoCambio.EOF
           If 0 <= Val(rsTipoCambio!nHasta) Then
                lblTCC = rsTipoCambio!nCompra
                lblTCV = rsTipoCambio!nVenta
                Exit Do
            End If
            rsTipoCambio.MoveNext
        Loop
    Else
        MsgBox "No se a definido el Tipo de Cambio", vbCritical, "AVISO"
        Set rsTipoCambio = Nothing
        Set oDCOMTipoCambioEsp = Nothing
        Exit Sub
    End If

End Sub

Private Sub cargarTotales()
    lblSubTotalMN = Format$(FEClientes.SumaRow(4), "#,##0.00")
    lblSubTotalME = Format$(FEClientes.SumaRow(5), "#,##0.00")
    lblSubTotalITFMN = Format$(FEClientes.SumaRow(6), "#,##0.00")
    lblSubTotalITFME = Format$(FEClientes.SumaRow(7), "#,##0.00")
    lblTotalMN = Format$(FEClientes.SumaRow(4) + FEClientes.SumaRow(6), "#,##0.00")
    lblTotalME = Format$(FEClientes.SumaRow(5) + FEClientes.SumaRow(7), "#,##0.00")
End Sub

Private Sub actualizarSeisUltimasRemuneraciones(ByVal pcCtaCod As String, ByVal pnMoneda As Integer, ByVal pnSueldos As Currency)

Dim oNCOMCaptaMovimiento As New COMNCaptaGenerales.NCOMCaptaMovimiento
Dim oNCOMCaptaDefinicion As New COMNCaptaGenerales.NCOMCaptaDefinicion
Dim oDCOMCaptaMovimiento As New COMDCaptaGenerales.DCOMCaptaMovimiento
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Dim oNCOMContFunciones As New COMNContabilidad.NCOMContFunciones
Dim oDCOMGeneral As New COMDConstSistema.DCOMGeneral
 
Dim rsCta As New ADODB.Recordset

Dim lsMovNro As String
Dim lnPorcDisp As Double
Dim lnExcedente As Double
Dim lnIntSaldo As Double
Dim ldUltMov As Date
Dim lnTasa As Double
Dim lnDiasTranscurridos As Integer
Dim lnSaldoRetiro As Currency

Set rsCta = oDCOMCaptaGenerales.GetDatosCuentaCTS(pcCtaCod)
lnSaldoRetiro = rsCta("nSaldRetiro")
lnTasa = rsCta("nTasaInteres")
ldUltMov = rsCta("dUltCierre")

lnDiasTranscurridos = DateDiff("d", ldUltMov, gdFecSis) - 1
If lnDiasTranscurridos < 0 Then
    lnDiasTranscurridos = 0
End If
lnIntSaldo = oNCOMCaptaMovimiento.GetInteres(lnSaldoRetiro, lnTasa, lnDiasTranscurridos, TpoCalcIntSimple)

lnPorcDisp = oNCOMCaptaDefinicion.GetCapParametro(gPorRetCTS)
lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
lnExcedente = 0


oDCOMCaptaMovimiento.AgregaDatosSueldosClientesCTS lsMovNro, pcCtaCod, pnMoneda, pnSueldos

Set rsCta = oDCOMCaptaMovimiento.ObtenerCapSaldosCuentasCTS(pcCtaCod, oDCOMGeneral.GetTipCambio(gdFecSis, TCFijoMes))
lnExcedente = rsCta!TotalSaldos - rsCta!TotalSueldos
If lnExcedente > 0 Then
    lnSaldoRetiro = lnExcedente * lnPorcDisp / 100
Else
    lnSaldoRetiro = 0
End If
oDCOMCaptaMovimiento.ActualizaSaldoRetiroCTS pcCtaCod, lnSaldoRetiro, lnIntSaldo
        
Set oNCOMCaptaDefinicion = Nothing
Set oNCOMCaptaMovimiento = Nothing
Set oNCOMContFunciones = Nothing
Set oDCOMCaptaMovimiento = Nothing
Set oDCOMCaptaGenerales = Nothing

End Sub

Private Function EsExoneradaLavadoDinero(ByVal pnFila As Integer) As Boolean
Dim oNCOMCaptaServicios As COMNCaptaServicios.NCOMCaptaServicios
Dim lsPersCod As String
Dim lbExito As Boolean

Set oNCOMCaptaServicios = New COMNCaptaServicios.NCOMCaptaServicios

lsPersCod = FEClientes.TextMatrix(pnFila, 1)
If Not oNCOMCaptaServicios.EsPersonaExoneradaLavadoDinero(lsPersCod) Then
    lbExito = False
    Exit Function
End If

Set oNCOMCaptaServicios = Nothing
EsExoneradaLavadoDinero = lbExito
End Function

Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero, ByVal pnFila As Integer)
Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim nMonto As Double
Dim oPersona As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsPers As New ADODB.Recordset

poLavDinero.TitPersLavDinero = FEClientes.TextMatrix(pnFila, 1)
poLavDinero.TitPersLavDineroNom = FEClientes.TextMatrix(pnFila, 2)
End Sub

Private Sub cboCheque_Click()

    If Right(cboCheque, 3) = Moneda.gMonedaNacional Then
        lblSimChe.Caption = "S/."
        lblMonChe.BackColor = &HC0FFFF
         fnMoneda = Moneda.gMonedaNacional
         limpiarCamposChq
    Else
        lblSimChe.Caption = "$"
        lblMonChe.BackColor = &HC0FFC0
        fnMoneda = Moneda.gMonedaExtranjera
        limpiarCamposChq
    End If

End Sub

Private Sub cboEfectivo_Click()
Dim i As Long
Dim lnFilas As Long
lnFilas = FEClientes.Rows - 1

    If Right(cboEfectivo, 3) = Moneda.gMonedaNacional Then
        FEClientes.BackColor = &HC0FFFF
        fnMoneda = Moneda.gMonedaNacional
        If lnFilas >= 1 And FEClientes.TextMatrix(1, 0) <> "" Then
            For i = 1 To lnFilas
                FEClientes_OnCellChange i, 4
            Next i
        End If
    Else
        lblSimTra.Caption = "$"
        FEClientes.BackColor = &HC0FFC0
        fnMoneda = Moneda.gMonedaExtranjera
        If lnFilas >= 1 And FEClientes.TextMatrix(1, 0) <> "" Then
            For i = 1 To lnFilas
                FEClientes_OnCellChange i, 5
            Next i
        End If
    End If
End Sub

Private Sub cboPeriodo_Click()
    lblDispCTS.Caption = Format$(CDbl(Trim(Right(cboPeriodo.Text, 5))) * 100, "#,##0.00")
End Sub

Private Sub cboTransferMoneda_Click()
    If Right(cboTransferMoneda, 3) = Moneda.gMonedaNacional Then
        lblSimTra.Caption = "S/."
        lblMonTra.BackColor = &HC0FFFF
        fnMoneda = Moneda.gMonedaNacional
        limpiarCamposTranf
    Else
        lblSimTra.Caption = "$"
        lblMonTra.BackColor = &HC0FFC0
        fnMoneda = Moneda.gMonedaExtranjera
        limpiarCamposTranf
    End If
    
End Sub

Private Sub cmdAgregar_Click()
Dim i, lnFilas As Integer

lnFilas = FEClientes.Rows

For i = 1 To lnFilas - 2
 If FEClientes.TextMatrix(i, 1) = FEClientes.TextMatrix(FEClientes.row, 1) Then
    If FEClientes.TextMatrix(i, 3) = FEClientes.TextMatrix(FEClientes.row, 3) Then
        MsgBox "El Cliente " & FEClientes.TextMatrix(i, 2) & " con el Nro de Cuenta " & FEClientes.TextMatrix(i, 3) & " ya fue agregado en la fila " & i & " de la relación." & Chr(10) & _
               "No se debe agregar dos veces un Cliente con una misma cuenta."
        FEClientes.TextMatrix(FEClientes.row, 3) = ""
        FEClientes.SetFocus
        Exit Sub
    End If
 End If
Next i

    If fnOpeCod = gCTSDepLotChq Then
        If lblNroDoc = "" Then
            MsgBox "Debe seleccionar un Cheque.", vbInformation, "!Aviso¡"
            If cmdDocumento.Visible And cmdDocumento.Enabled Then cmdDocumento.SetFocus
            Exit Sub
        End If
        'EJVG20140408 *** Verifica cantidad cuentas a depositar
        If IIf(FEClientes.TextMatrix(1, 0) = "", 0, FEClientes.Rows - 1) >= oDocRec.fnNroCliLote Then
            MsgBox "No puede agregar ya que la cantidad máxima de clientes registradas en el cheque es de " & oDocRec.fnNroCliLote, vbInformation, "Aviso"
            If cmdAgregar.Visible And cmdAgregar.Enabled Then cmdAgregar.SetFocus
            Exit Sub
        End If
        'END EJVG *******
    ElseIf fnOpeCod = gCTSDepLotTransf Then
        If lblTrasferND = "" Then
            MsgBox "Debe seleccionar un Voucher.", vbInformation, "!Aviso¡"
            Exit Sub
        End If
    ElseIf fnOpeCod = gCTSDepLotCargoCta Then
        If Len(txtCuentaCargo.NroCuenta) < 18 Then
            MsgBox "Debe ingresar la cuenta de ahorro a cargar.", vbInformation, "!Aviso¡"
            Exit Sub
        End If
    End If

    FEClientes.lbEditarFlex = True
    FEClientes.AdicionaFila
    FEClientes.TextMatrix(FEClientes.row, 1) = ""
    FEClientes.TextMatrix(FEClientes.row, 2) = ""
    FEClientes.TextMatrix(FEClientes.row, 3) = ""
    FEClientes.TextMatrix(FEClientes.row, 4) = "0.00"
    FEClientes.TextMatrix(FEClientes.row, 5) = "0.00"
    FEClientes.TextMatrix(FEClientes.row, 6) = "0.00"
    FEClientes.TextMatrix(FEClientes.row, 7) = "0.00"
    FEClientes.SetFocus
    SendKeys "{ENTER}"
End Sub

Private Sub cmdCargar_Click()
If InStr(Trim(txtRuta), ".xls") = 0 And InStr(Trim(txtRuta), ".xlsx") = 0 Then
    MsgBox "Debe seleccionar el archivo .xls o .xlsx para que sea cargado los datos."
    cmdRuta.SetFocus
    Exit Sub
End If

LimpiaFlex FEClientes
cboEfectivo.Enabled = False
lbVistoVal = False

Dim oUCOMPersona As COMDPersona.UCOMPersona
Dim rsPersona As ADODB.Recordset
'Variable de tipo Aplicación de Excel
Dim oExcel As Excel.Application
Dim lnTipoDOI, lnFila1, lnFila2, lnFilasFormato As Integer
Dim lsDOI As String
Dim lsMoneda As String
Dim lbExisteCTS As Boolean
Dim lbExisteError As Boolean
Dim bSuperaDepAnio As Boolean
Dim lnInicioFila As Integer 'JUEZ 20150922

'Una variable de tipo Libro de Excel
Dim oLibro As Excel.Workbook
Dim oHoja As Excel.Worksheet

'creamos un nuevo objeto excel
Set oExcel = New Excel.Application

lnInicioFila = 36 'JUEZ 20150922
lnFilasFormato = 3525
 
PB1.Min = 0
PB1.Max = lnFilasFormato
PB1.value = 0
lblProcesando.Visible = True
PB1.Visible = True
 
'Usamos el método open para abrir el archivo que está en el directorio del programa llamado archivo.xls
Set oLibro = oExcel.Workbooks.Open(txtRuta)

'Hacemos referencia a la Hoja
Set oHoja = oLibro.Sheets(1)

'Hacemos el Excel Visible
'oLibro.Visible = False

FEClientes.lbEditarFlex = True

lsMoneda = oHoja.Cells(27, 4)

cboEfectivo.ListIndex = 0
If Trim(Left(cboEfectivo.Text, 10)) <> UCase(lsMoneda) Then
    cboEfectivo.ListIndex = 1
End If

'FRHU 20141218 OBSERVACION
Dim oConSis As New COMDConstSistema.DCOMGeneral
Dim sPass As String
sPass = oConSis.LeeConstSistema(190)
oHoja.Unprotect (sPass)
'FIN FRHU 2041218
With oHoja
    PB1.value = lnInicioFila
    For lnFila1 = lnInicioFila To lnFilasFormato
        lnTipoDOI = .Cells(lnFila1, 4)
        lsDOI = .Cells(lnFila1, 3)
        If lsDOI <> "" Then lsDOI = IIf(Len(lsDOI) < 8, String(8 - Len(lsDOI), "0"), "") & lsDOI 'JUEZ 20150922
        If Len(lsDOI) > 0 Then
            '***Agrega nueva fila****
            FEClientes.AdicionaFila
            lnFila2 = FEClientes.row
            FEClientes.TextMatrix(lnFila2, 1) = ""
            FEClientes.TextMatrix(lnFila2, 2) = ""
            FEClientes.TextMatrix(lnFila2, 3) = ""
            FEClientes.TextMatrix(lnFila2, 4) = "0.00"
            FEClientes.TextMatrix(lnFila2, 5) = "0.00"
            FEClientes.TextMatrix(lnFila2, 6) = "0.00"
            FEClientes.TextMatrix(lnFila2, 7) = "0.00"
            '***Fin Agrega nueva fila
            
            '***Verifica si la persona esta registrada****
            Set oUCOMPersona = New COMDPersona.UCOMPersona
            Set rsPersona = New ADODB.Recordset
            Set rsPersona = oUCOMPersona.devolverDatosPersona(lnTipoDOI, lsDOI)
            '***Fin Verifica si la persona esta registrada
                        
            If Not (rsPersona.BOF And rsPersona.EOF) Then
                FEClientes.TextMatrix(lnFila2, 1) = rsPersona!cperscod
                FEClientes.TextMatrix(lnFila2, 2) = rsPersona!cPersNombre

                devolverCTSPersona rsPersona!cperscod, .Cells(lnFila1, 9), lbExisteCTS
                If lbExisteCTS Then
                    actualizarSeisUltimasRemuneraciones Trim(FEClientes.TextMatrix(lnFila2, 3)), IIf(.Cells(lnFila1, 11) = "Soles", 1, 2), CCur(.Cells(lnFila1, 12))
                Else
                    .Range("A" & lnFila1, "M" & lnFila1).Interior.Color = RGB(255, 255, 0)
                    .Cells(lnFila1, 13) = "EL NRO CUENTA CTS NO RELACIONADO CON EL CLIENTE O NO EXISTE EN EL SISTEMA."
                    lbExisteError = True
                End If
                If lsMoneda = "Soles" Then
                    FEClientes.TextMatrix(lnFila2, 4) = Format$(.Cells(lnFila1, 10), "#,##0.00")
                    Call FEClientes_OnCellChange(CLng(lnFila2), 4)
                    
                Else
                    FEClientes.TextMatrix(lnFila2, 5) = Format$(.Cells(lnFila1, 10), "#,##0.00")
                    Call FEClientes_OnCellChange(CLng(lnFila2), 5)
                End If
                Call FEClientes_OnCellChange(CLng(lnFila2), 3) 'JUEZ 20150922
                'JUEZ 20141014 Nuevos parametros **************
                Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
                Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                    nCantDepCta = clsMant.ObtenerCantidadOperaciones(.Cells(lnFila1, 9), gCapMovDeposito, gdFecSis)
                Set clsMant = Nothing
                If nCantDepCta >= nParCantDepAnio Then
                    bSuperaDepAnio = True
                    lbExisteError = True
                    .Range("A" & lnFila1, "M" & lnFila1).Interior.Color = RGB(255, 255, 0)
                    .Cells(lnFila1, 13) = "SUPERA NUMERO MAXIMO DE DEPOSITOS CTS"
                End If
                'END JUEZ *************************************
            Else
                .Range("A" & lnFila1, "M" & lnFila1).Interior.Color = RGB(255, 255, 0)
                .Cells(lnFila1, 13) = "LA PERSONA NO EXISTE EN EL SISTEMA."
                lbExisteError = True
            End If
        Else 'JUEZ 20150922
            PB1.value = PB1.Max
            Exit For
        End If
        PB1.value = lnFila1
    Next lnFila1
End With

oHoja.Protect (sPass) 'FRHU 20141218 OBSERVACION
FEClientes.ColumnasAEditar = "X-X-X-X-X-X-X-X-X"
FEClientes.lbEditarFlex = False 'EJVG20130917
lblProcesando.Visible = False
PB1.Visible = False

If lbExisteError = False Then
    'oLibro.Close
    cmdGuardar.Enabled = True
Else
    oExcel.Visible = True
    'Exit Sub
End If
'JUEZ 20141014 **********************************************
If bSuperaDepAnio Then
    If MsgBox("Existen CTS que superaron el límete de depósitos por año, se requiere de VB del supervisor para grabar la operación. Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Set loVistoElectronico = New frmVistoElectronico
                
    lbVistoVal = loVistoElectronico.Inicio(3, fnOpeCod)
    If Not lbVistoVal Then Exit Sub
    cmdGuardar.Enabled = True
End If
'END JUEZ ***************************************************
Set oHoja = Nothing
Set oLibro = Nothing
oExcel.Quit
Set oExcel = Nothing

End Sub

Private Sub cmdDocumento_Click()
    'EJVG20140219 ***
'frmCapAperturaListaChq.Inicia frmCapDepositosLote, fnOpeCod, fnMoneda, fnProducto
'fnValorChq = CCur(lblMonChe)
    Dim oForm As New frmChequeBusqueda
    Dim lnOperacion As TipoOperacionCheque

    On Error GoTo ErrCargaDocumento
    If fnOpeCod = gCTSDepLotChq Then
        lnOperacion = CTS_DepositoLote
    Else
        lnOperacion = Ninguno
    End If

    Set oDocRec = oForm.Iniciar(fnMoneda, lnOperacion)
    Set oForm = Nothing
    
    FormateaFlex FEClientes
    txtGlosa.Text = oDocRec.fsGlosa
    lblNombreIF.Caption = oDocRec.fsPersNombre
    lblNroDoc.Caption = oDocRec.fsNroDoc
    fsCodIF = oDocRec.fsPersCod
    lblMonChe.Caption = Format(oDocRec.fnMonto, gsFormatoNumeroView)
    Exit Sub
ErrCargaDocumento:
    MsgBox "Ha sucedido un error al cargar los datos del Documento", vbCritical, "Aviso"
'END EJVG *******
End Sub

Private Sub CmdEliminar_Click()
    If MsgBox("¿Esta seguro que desea eliminar el Nro. de Cuenta " & FEClientes.TextMatrix(FEClientes.row, 3) & "?", vbYesNo + vbInformation, "Aviso") = vbYes Then
        FEClientes.EliminaFila FEClientes.row
        cargarTotales
    End If
End Sub

'JUEZ 20150922 **************************************
Private Sub cmdFormato_Click()
Dim xlsAplicacion As Excel.Application
Dim lsFile As String

    lsFile = "FormatoDepositoCTSLote"
    
    Set xlsAplicacion = New Excel.Application
    xlsAplicacion.Workbooks.Open (App.Path & "\FormatoCarta\" & lsFile & ".xls")
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
End Sub
'END JUEZ *******************************************

Private Sub cmdGuardar_Click()
'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
'fin Comprobacion si es RFIII

Dim lsMovNro As String, lsBoletaImp As String
Dim lsmensaje As String

Dim lsBoletaCargo  As String 'CTI4 ERS0112020
Dim MatDatosAho(14) As String 'CTI4 ERS0112020
Dim lsNombreClienteCargoCta As String 'CTI4 ERS0112020
Dim oNCOMColpContrato As COMNColoCPig.NCOMColPContrato 'CTI4 ERS0112020

If ValidarDatos = False Then Exit Sub
If fraFormaPago.Visible Then 'CTI4 ERS0112020
    If Not ValidaFormaPago Then Exit Sub
End If

'If MsgBox("¿Esta seguro que desea guardar?", vbYesNo, "!Aviso¡") = vbYes Then
If MsgBox("¿Esta seguro que desea guardar?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    Dim oNCOMContFunciones As New COMNContabilidad.NCOMContFunciones
    Dim oNCOMCaptaMovimiento As New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oNCOMTipoCambio As New COMDConstSistema.NCOMTipoCambio
    Dim oNCOMContImprimir As COMNContabilidad.NCOMContImprimir
    Dim oDCOMPersonas As New COMDPersona.DCOMPersonas
    Dim oNCOMCaptaDefinicion As COMNCaptaGenerales.NCOMCaptaDefinicion
    
    'PASI20140530
    Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    'end PASI
    
    Dim rsCuentas As New ADODB.Recordset
    Dim rsPersOcu As ADODB.Recordset
    Dim lnMontoPersOcupacion As Currency
    Dim lnAcumulado As Currency
    Dim lnPorcDisp As Double
    
    Dim ofrmMovLavDinero()  As New frmMovLavDinero
    Dim lsPersLavadoDinero() As String
    Dim lnMontoLavadoDinero As Double
    Dim lnTCLavadoDinero As Double
    Dim lnMovNroLavadoDinero As Long
    Dim lnMonedaLavadoDinero As Integer
    Dim lsTipoCuentaLavadoDinero As String
    Dim lnMontoDepositadoLavadoDinero As Double
    Dim lsCuentaLavadoDinero() As String
    
    Dim lnTC As Double
    Dim i, J, K, a As Integer
    
    '****************APRI20170523 RFC1705230001************************************
    Dim PersCIIU() As Variant
    Dim PersCod() As String
    Dim PersNombre() As String
    Dim count, counter, B As Integer
    count = 0
    counter = 0
    ReDim Preserve PersCod(count)
    
    For a = 1 To FEClientes.Rows - 1
        Set rsPersOcu = New ADODB.Recordset
        Set rsPersOcu = oDCOMPersonas.ObtenerDatosPersona(FEClientes.TextMatrix(a, 1))
        If IsNull(rsPersOcu!cPersCIIU) Or rsPersOcu!cPersCIIU = "" Then
            count = count + 1
            
            ReDim Preserve PersCod(count)
            ReDim Preserve PersNombre(count)
            
            PersCod(counter) = rsPersOcu!cperscod
            PersNombre(counter) = rsPersOcu!cPersNombre
            
            counter = counter + 1
            
        End If
    Next a
    
    If UBound(PersCod) > 0 Then
    counter = 0
    ReDim Preserve PersCIIU(UBound(PersCod), 2)
    For B = 1 To UBound(PersCod)
        PersCIIU(B, 1) = PersCod(counter)
        PersCIIU(B, 2) = PersNombre(counter)
        
        counter = counter + 1
    Next B
    
    If UBound(PersCIIU) = 1 And PersCIIU(1, 1) <> "" Then
        MsgBox "Para continuar con la Operación Ud. debe actualizar los datos del" & Chr(13) & "Cliente: " & PersCIIU(1, 2), vbInformation, "Aviso"
        Dim foPersona As New frmPersona
        Do While Not foPersona.realizarMantenimiento(PersCIIU(1, 1))
            MsgBox "Aun no ha realizado la actualización de los datos del cliente: " & PersCIIU(1, 2) & ". Para continuar actualice." & Chr(13), vbInformation, "Aviso"
        Loop
        MsgBox "Se ha actualizado correctamente los datos del cliente: " & PersCIIU(1, 2) & "." & Chr(13), vbInformation, "Aviso"
     ElseIf UBound(PersCIIU) > 1 Then
        MsgBox "Se ha encontrado clientes que requieren que sus datos sean actualizados. Para continuar favor de actualizar.", vbInformation, "Aviso"
        PersonaActualizarCIIUExcel PersCIIU
        Exit Sub
    End If
    End If
    '****************END APRI RFC1705230001*****************************************
    
    'Realiza la Validación para el Lavado de Dinero
    ReDim ofrmMovLavDinero(FEClientes.Rows - 1)
    ReDim lsPersLavadoDinero(FEClientes.Rows - 1)
    ReDim lsCuentaLavadoDinero(FEClientes.Rows - 1)
    If oDocRec Is Nothing Then Set oDocRec = New UDocRec 'EJVG20140408
    
    For i = 1 To FEClientes.Rows - 1
    Set oNCOMCaptaDefinicion = New COMNCaptaGenerales.NCOMCaptaDefinicion
        If Not EsExoneradaLavadoDinero(i - 1) Then
            lsPersLavadoDinero(i - 1) = ""
            lnMontoLavadoDinero = oNCOMCaptaDefinicion.GetCapParametro(gMonOpeLavDineroME)
            Set oNCOMCaptaDefinicion = Nothing
            
            lsTipoCuentaLavadoDinero = "INDIVIDUAL"
            lnMonedaLavadoDinero = CInt(Mid(FEClientes.TextMatrix(i, 3), 9, 1))
            lnMontoDepositadoLavadoDinero = IIf(lnMonedaLavadoDinero = gMonedaNacional, FEClientes.TextMatrix(i, 4), FEClientes.TextMatrix(i, 5))
            lsCuentaLavadoDinero(i - 1) = Trim(FEClientes.TextMatrix(i, 3))
            
            If lnMonedaLavadoDinero = gMonedaNacional Then
                Set oNCOMTipoCambio = New COMDConstSistema.NCOMTipoCambio
                lnTCLavadoDinero = oNCOMTipoCambio.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set oNCOMTipoCambio = Nothing
            Else
                lnTCLavadoDinero = 1
            End If
            If lnMontoDepositadoLavadoDinero >= Round(lnMontoLavadoDinero * lnTCLavadoDinero, 2) Then
                Call IniciaLavDinero(ofrmMovLavDinero(i - 1), i)
                lsPersLavadoDinero(i - 1) = ofrmMovLavDinero(i - 1).Inicia(, , , , False, True, lnMontoDepositadoLavadoDinero, lsCuentaLavadoDinero(i - 1), Mid(Me.Caption, 15), False, lsTipoCuentaLavadoDinero, , , , , lnMonedaLavadoDinero, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                If ofrmMovLavDinero(i - 1).OrdPersLavDinero = "" Then Exit Sub
            End If
        End If
      Next i

    Set rsCuentas = FEClientes.GetRsNew()
    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    lnPorcDisp = CDbl(lblDispCTS)
    PB1.Min = 0
    PB1.Max = FEClientes.Rows - 1
    PB1.value = 0
    lblProcesando.Visible = True
    PB1.Visible = True
    
    Set oNCOMColpContrato = New COMNColoCPig.NCOMColPContrato 'CTI4 ERS0112020
    
    If fnOpeCod = gCTSDepLotCargoCta Then lsNombreClienteCargoCta = PstaNombre(oNCOMColpContrato.ObtieneNombreTitularCargoCta(txtCuentaCargo.NroCuenta)) 'CTI ERS0112020
    
    If fnOpeCod = gCTSDepLotEfec Then
        lnMovNroLavadoDinero = oNCOMCaptaMovimiento.CapAbonoCuentaCTSLote(rsCuentas, fnOpeCod, lsMovNro, txtGlosa, lnPorcDisp, , , , , gsNomAge, sLpt, , , gsCodCMAC, lsmensaje, lsBoletaImp, gbImpTMU, , , fnMoneda)
    ElseIf fnOpeCod = gCTSDepLotCargoCta Then
        lnMovNroLavadoDinero = oNCOMCaptaMovimiento.CapAbonoCuentaCTSLote(rsCuentas, fnOpeCod, lsMovNro, txtGlosa, lnPorcDisp, , , , , gsNomAge, sLpt, , , gsCodCMAC, lsmensaje, lsBoletaImp, gbImpTMU, , , , , , , fnMoneda, txtCuentaCargo.NroCuenta, MatDatosAho)
    ElseIf fnOpeCod = gCTSDepLotChq Then
        'lnMovNroLavadoDinero = oNCOMCaptaMovimiento.CapAbonoCuentaCTSLote(rsCuentas, fnOpeCod, lsMovNro, txtGlosa, lnPorcDisp, TpoDocCheque, lblNroDoc, fsCodIF, fdFechaValorizacion, gsNomAge, sLpt, , , gsCodCMAC, lsmensaje, lsBoletaImp, gbImpTMU, , , , fnMoneda)
        fdFechaValorizacion = oNCapMov.ObtenerFechaValorizaCheque(oDocRec.fsNroDoc, oDocRec.fsPersCod, oDocRec.fsIFTpo, oDocRec.fsIFCta) 'PASI20140530
        lnMovNroLavadoDinero = oNCOMCaptaMovimiento.CapAbonoCuentaCTSLote(rsCuentas, fnOpeCod, lsMovNro, txtGlosa, lnPorcDisp, oDocRec.fnTpoDoc, oDocRec.fsNroDoc, oDocRec.fsPersCod, fdFechaValorizacion, gsNomAge, sLpt, , , gsCodCMAC, lsmensaje, lsBoletaImp, gbImpTMU, , , , fnMoneda, oDocRec.fsIFTpo, oDocRec.fsIFCta) 'EJVG20140408
      ElseIf fnOpeCod = gCTSDepLotTransf Then
        lnMovNroLavadoDinero = oNCOMCaptaMovimiento.CapAbonoCuentaCTSLote(rsCuentas, fnOpeCod, lsMovNro, txtTransferGlosa, lnPorcDisp, , , , , gsNomAge, sLpt, fnMovNroTransfer, fnMoneda, gsCodCMAC, lsmensaje, lsBoletaImp, gbImpTMU, fnMovNroRVD, fnTransferSaldo)
    End If
    
    If lnMovNroLavadoDinero > 0 Then
        For J = 0 To UBound(lsPersLavadoDinero)
            If lsPersLavadoDinero(J) <> "" Then
               Call ofrmMovLavDinero(J).InsertarLavDinero(ofrmMovLavDinero(J).TitPersLavDinero, , , lnMovNroLavadoDinero, ofrmMovLavDinero(J).BenPersLavDinero, ofrmMovLavDinero(J).TitPersLavDinero, ofrmMovLavDinero(J).OrdPersLavDinero, ofrmMovLavDinero(J).ReaPersLavDinero, ofrmMovLavDinero(J).BenPersLavDinero, ofrmMovLavDinero(J).VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, ofrmMovLavDinero(J).BenPersLavDinero2, ofrmMovLavDinero(J).BenPersLavDinero3, ofrmMovLavDinero(J).BenPersLavDinero4)
               MsgBox "Coloque papel para la Boleta de Lavado de Dinero", vbInformation, "Aviso"
               Call ofrmMovLavDinero(J).imprimirBoletaREU(lsCuentaLavadoDinero(J), Mid(lsCuentaLavadoDinero(J), 9, 1), ofrmMovLavDinero(J).OrigenPersLavDinero, ofrmMovLavDinero(J).NroREU)
            
            End If
         Next J
    End If
    
    lnTC = oNCOMTipoCambio.EmiteTipoCambio(gdFecSis, TCFijoDia)
               
    For K = 1 To FEClientes.Rows - 1
        Set rsPersOcu = New ADODB.Recordset
        Set rsPersOcu = oDCOMPersonas.ObtenerDatosPersona(FEClientes.TextMatrix(K, 1))
        
        lnAcumulado = oDCOMPersonas.ObtenerPersAcumuladoMontoOpe(lnTC, Mid(Format(gdFecSis, "yyyymmdd"), 1, 6), rsPersOcu!cperscod)
        lnMontoPersOcupacion = oDCOMPersonas.ObtenerParamPersAgeOcupacionMonto(Mid(rsPersOcu!cperscod, 4, 2), CInt(Mid(rsPersOcu!cPersCIIU, 2, 2)))
        
        If lnAcumulado >= lnMontoPersOcupacion Then
           If Not oDCOMPersonas.ObtenerPersonaAgeOcupDatos_Verificar(rsPersOcu!cperscod, gdFecSis) Then
               oDCOMPersonas.insertarPersonaAgeOcupacionDatos gnMovNro, rsPersOcu!cperscod, IIf(fnMoneda = gMonedaNacional, lblTotalMN, lblTotalMN * lnTC), lnAcumulado, gdFecSis, lsMovNro
           End If
        End If
               
        PB1.value = K
    Next K
    
    If lbVistoVal Then loVistoElectronico.RegistraVistoElectronico (gnMovNro) 'JUEZ 20141014
    
    'CTI4 ERS0112020
    If fraFormaPago.Visible Then
        If (fnOpeCod = gCTSDepLotCargoCta) Then
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            Dim oMovOperacion As COMDMov.DCOMMov
            Dim nMovNroOperacion As Long
            Dim rsCli As New ADODB.Recordset
            Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
            Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
            Set oMovOperacion = New COMDMov.DCOMMov
            nMovNroOperacion = oMovOperacion.GetnMovNro(lsMovNro)

            loVistoElectronico2.RegistraVistoElectronico nMovNroOperacion, , gsCodUser, nMovNroOperacion

            If nRespuesta = 2 Then
                Set rsCli = clsCli.GetPersonaCuenta(txtCuentaCargo.NroCuenta, gCapRelPersTitular)
                oSolicitud.ActualizarCapAutSinTarjetaVisto_nMovNro gsCodUser, gsCodAge, txtCuentaCargo.NroCuenta, rsCli!cperscod, nMovNroOperacion, CStr(gAhoCargoCtaDepositoLoteCTS)
            End If
            Set oMovOperacion = Nothing
            nRespuesta = 0
        End If
        End If
    End If
    'CTI4 end

    If Trim(lsmensaje) <> "" Then
       MsgBox lsmensaje, vbInformation
    End If

    If fnOpeCod = gCTSDepLotCargoCta Then 'CTI4 ERS0112020
        Dim oNCOMCaptaImpresion As New COMNCaptaGenerales.NCOMCaptaImpresion
        lsBoletaCargo = oNCOMCaptaImpresion.ImprimeBoletaAhorro(IIf(fnOpeCod <> gCTSDepLotCargoCta, "RETIRO AHORROS", "CARGO A CUENTA EMPRESA"), "CARGO A CUENTA POR DEP. CTS.", "", CStr(CDbl(IIf(fnMoneda = gMonedaNacional, (lblTotalMN), (lblTotalME)))), lsNombreClienteCargoCta, txtCuentaCargo.NroCuenta, "", CDbl(MatDatosAho(10)), CDbl(MatDatosAho(3)), "", 1, CDbl(MatDatosAho(11)), , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, gsCodCMAC, 0, , , gbImpTMU, IIf(fnOpeCod = gCTSDepLotCargoCta, True, False))
        Set oNCOMCaptaImpresion = Nothing
    End If
    
    If Trim(lsBoletaImp) <> "" Then ImprimeBoleta lsBoletaImp & lsBoletaCargo

    Set oNCOMContImprimir = Nothing
    Set oNCOMTipoCambio = Nothing
    Set oNCOMCaptaMovimiento = Nothing
    fnMovNroRVD = 0
    lblMonTra = "0.00"
    PB1.Visible = False
    lblProcesando.Visible = False
    
'End If
If cboEfectivo.Enabled = False Then
    cboEfectivo.Enabled = True
End If
limpiarCampos
Exit Sub
ErrGraba:
    MsgBox err.Description, vbExclamation, "Error"
    Exit Sub
End Sub

Private Sub cmdRuta_Click()

If fnOpeCod = gCTSDepLotChq Then
    If lblNroDoc = "" Then
        MsgBox "Debe seleccionar un Cheque.", vbInformation, "!Aviso¡"
        Exit Sub
    End If
ElseIf fnOpeCod = gCTSDepLotTransf Then
    If lblTrasferND = "" Then
        MsgBox "Debe seleccionar un Voucher.", vbInformation, "!Aviso¡"
        Exit Sub
    End If
ElseIf fnOpeCod = gCTSDepLotCargoCta Then
    If Len(txtCuentaCargo.texto) < 18 Then
        MsgBox "Debe ingresar la cuenta de ahorro a cargar.", vbInformation, "!Aviso¡"
        Exit Sub
    End If
End If

txtRuta.Text = Empty
    
CdlgFile.InitDir = "C:\"
CdlgFile.Filter = "Archivos de Excel (*.xls)|*.xls| Archivos de Excel (*.xlsx)|*.xlsx"

CdlgFile.ShowOpen
 
If CdlgFile.FileName <> Empty Then
    txtRuta = CdlgFile.FileName
    cmdAgregar.Enabled = False
    cmdEliminar.Enabled = False
    cmdGuardar.Enabled = False
Else
    txtRuta = ""
    MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
    Exit Sub
End If
End Sub

Private Sub cmdsalir_Click()
limpiarCampos
Unload Me
End Sub



Private Sub cmdTranfer_Click()
    Dim lsGlosa As String
    Dim lsVoucher As String
    Dim lsIF As String
    Dim oForm As New frmCapRegVouDepBus
    Dim lnTipMot As Integer
    Dim i As Integer
    Dim lsDetalle As String
    
    If Me.cboTransferMoneda.Text = "" Then
        MsgBox "Debe escoger la moneda de la transferencia.", vbInformation, "Aviso"
        cboTransferMoneda.SetFocus
        Exit Sub
    End If
        
    If fnOpeCod = gCTSDepLotTransf Then
        lnTipMot = 7
    End If
    'EJVG20130917 ***
    fnNroClientesTransf = 0
    'oform.iniciarFormulario fnMoneda, lnTipMot, lsGlosa, lsIF, lsVoucher, fnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, fnMovNroTransfer
    oForm.iniciarFormulario fnMoneda, lnTipMot, lsGlosa, lsIF, lsVoucher, fnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, fnMovNroTransfer, lsDetalle
    fnNroClientesTransf = Val(lsDetalle)
    'END EJVG *******
    txtTransferGlosa.Text = lsGlosa
    lbltransferBco.Caption = lsIF
    lblTrasferND.Caption = lsVoucher
    
    If fnMovNroTransfer <> -1 Then
        'cmdAgregar.SetFocus
        If cmdAgregar.Visible And cmdAgregar.Enabled Then cmdAgregar.SetFocus
    End If
    
    txtTransferGlosa.Locked = True
    lblMonTra = Format(fnTransferSaldo, "#,##0.00")
    
End Sub

Private Sub FEClientes_Click()

If Trim(txtRuta) = "" Then
    If FEClientes.TextMatrix(FEClientes.row, 1) <> "" Then
        If FEClientes.Col = 3 Then
            devolverCTSPersona FEClientes.TextMatrix(FEClientes.row, 1)
        End If
    End If
End If
End Sub

Private Sub FEClientes_EnterCell()
If Trim(txtRuta) = "" Then
    If FEClientes.TextMatrix(FEClientes.row, 1) <> "" Then
        If FEClientes.Col = 3 Then
            devolverCTSPersona FEClientes.TextMatrix(FEClientes.row, 1)
        End If
    End If
End If
End Sub

Private Sub FEClientes_OnCellChange(pnRow As Long, pnCol As Long)
Dim lnMonedaCuenta As Moneda

Dim clsCap As New COMNCaptaGenerales.NCOMCaptaMovimiento 'CTI4 ERS0112020

If pnCol = 3 Then
    If Trim(FEClientes.TextMatrix(pnRow, 3)) = "" Then Exit Sub
    
    lnMonedaCuenta = CInt(Mid(FEClientes.TextMatrix(pnRow, 3), 9, 1))
    If lnMonedaCuenta = gMonedaNacional Then
        FEClientes.BackColorRow &HC0FFFF
    Else
        FEClientes.BackColorRow &HC0FFC0
    End If
    'CTI4 ERS0112020
    If fraFormaPago.Visible = True Then
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            If fnMoneda <> CInt(Mid(FEClientes.TextMatrix(pnRow, 3), 9, 1)) Then
                MsgBox "La tipo de moneda de la cuenta no es la correcta.", vbInformation, "!Aviso¡"
                FEClientes.EliminaFila FEClientes.row
                Exit Sub
            End If
        End If
    End If
    'END CTI4
    If Not ValidaDepCTS Then Exit Sub
End If

If pnCol = 4 Or pnCol = 5 Then

    Dim lnMontoDepositar As Currency
    
    If Trim(FEClientes.TextMatrix(pnRow, 3)) = "" Then Exit Sub
    
    lnMonedaCuenta = CInt(Mid(FEClientes.TextMatrix(pnRow, 3), 9, 1))
    lnMontoDepositar = CCur(FEClientes.TextMatrix(pnRow, pnCol))

    If fnMoneda = gMonedaNacional Then
        If lnMonedaCuenta = fnMoneda Then
            FEClientes.TextMatrix(pnRow, 5) = "0.00"
        Else
            If pnCol = 5 Then
                FEClientes.TextMatrix(pnRow, 4) = Format$(lnMontoDepositar * CDbl(lblTCV), "#0.00")
            Else
                FEClientes.TextMatrix(pnRow, 5) = Format$(lnMontoDepositar / CDbl(lblTCV), "#0.00")
            End If
        End If
    ElseIf fnMoneda = gMonedaExtranjera Then
        If lnMonedaCuenta = fnMoneda Then
            FEClientes.TextMatrix(pnRow, 4) = "0.00"
        Else
            If pnCol = 5 Then
                FEClientes.TextMatrix(pnRow, 4) = Format$(lnMontoDepositar * CDbl(lblTCC), "#0.00")
            Else
                FEClientes.TextMatrix(pnRow, 5) = Format$(lnMontoDepositar / CDbl(lblTCC), "#0.00")
            End If
        End If
    End If
    
    cargarTotales
    
    
    If fnMoneda = gMonedaNacional Then
            
        If fnOpeCod = gCTSDepLotChq Then
            
            If CCur(lblMonChe) < CCur(FEClientes.TextMatrix(pnRow, 4)) Then
                MsgBox "EL DEPÓSITO supera al monto establecido por el Cheque.", vbInformation, "!Aviso¡"
                FEClientes.TextMatrix(pnRow, 4) = "0.00"
                cargarTotales
                Exit Sub
            ElseIf CCur(lblMonChe) < CCur(lblTotalMN) Then
                MsgBox "SUMA TOTAL supera al monto establecido por el Cheque.", vbInformation, "!Aviso¡"
                FEClientes.TextMatrix(pnRow, 4) = "0.00"
                cargarTotales
                Exit Sub
            ElseIf CCur(lblMonChe) = CCur(lblTotalMN) Then
                cmdAgregar.Enabled = False
                cmdRuta.Enabled = False
                cmdCargar.Enabled = False
            ElseIf CCur(lblMonChe) > CCur(lblTotalMN) Then
                cmdAgregar.Enabled = True
                cmdRuta.Enabled = True
                cmdCargar.Enabled = True
            End If
        ElseIf fnOpeCod = gCTSDepLotTransf Then
            If CCur(lblMonTra) < CCur(FEClientes.TextMatrix(pnRow, 4)) Then
                MsgBox "EL DEPÓSITO supera al monto establecido por el Cheque.", vbInformation, "!Aviso¡"
                FEClientes.TextMatrix(pnRow, 4) = "0.00"
                cargarTotales
                Exit Sub
            ElseIf CCur(lblMonTra) < CCur(lblTotalMN) Then
                MsgBox "SUMA TOTAL supera al monto establecido por el Voucher.", vbInformation, "!Aviso¡"
                FEClientes.TextMatrix(pnRow, 4) = "0.00"
                cargarTotales
                Exit Sub
            ElseIf CCur(lblMonTra) = CCur(lblTotalMN) Then
                cmdAgregar.Enabled = False
                cmdRuta.Enabled = False
                cmdCargar.Enabled = False
            ElseIf CCur(lblMonTra) > CCur(lblTotalMN) Then
                cmdAgregar.Enabled = True
                cmdRuta.Enabled = True
                cmdCargar.Enabled = True
            End If
        ElseIf fnOpeCod = gCTSDepLotCargoCta Then 'CTI4 ERS0112020
            If Not clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, CCur(FEClientes.TextMatrix(pnRow, 4))) Then
                MsgBox "EL DEPÓSITO supera al saldo de la cuenta a cargar.", vbInformation, "!Aviso¡"
                FEClientes.TextMatrix(pnRow, 4) = "0.00"
                cargarTotales
                Exit Sub
            ElseIf Not clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, CCur(lblTotalMN)) Then
                MsgBox "SUMA TOTAL supera al saldo de la cuenta a cargar.", vbInformation, "!Aviso¡"
                FEClientes.TextMatrix(pnRow, 4) = "0.00"
                cargarTotales
                Exit Sub
            ElseIf clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, CCur(lblTotalMN)) Then
                cmdAgregar.Enabled = True
                cmdRuta.Enabled = True
                cmdCargar.Enabled = True
            Else
                cmdAgregar.Enabled = False
                cmdRuta.Enabled = False
                cmdCargar.Enabled = False
            End If
        End If
        
        
        
    ElseIf fnMoneda = gMonedaExtranjera Then
    
        If fnOpeCod = gCTSDepLotChq Then
            If CCur(lblMonChe) < CCur(FEClientes.TextMatrix(pnRow, 5)) Then
                MsgBox "EL DEPÓSITO supera al monto establecido por el Cheque.", vbInformation, "!Aviso¡"
                FEClientes.TextMatrix(pnRow, 5) = "0.00"
                cargarTotales
                Exit Sub
            ElseIf CCur(lblMonChe) < CCur(lblTotalME) Then
                MsgBox "SUMA TOTAL supera al monto establecido por el Cheque.", vbInformation, "!Aviso¡"
                FEClientes.TextMatrix(pnRow, 5) = "0.00"
                cargarTotales
                Exit Sub
            ElseIf CCur(lblMonChe) = CCur(lblTotalME) Then
                cmdAgregar.Enabled = False
                cmdRuta.Enabled = False
                cmdCargar.Enabled = False
            ElseIf CCur(lblMonChe) > CCur(lblTotalME) Then
                cmdAgregar.Enabled = True
                cmdRuta.Enabled = True
                cmdCargar.Enabled = True
            End If
        ElseIf fnOpeCod = gCTSDepLotTransf Then
            If CCur(lblMonTra) < CCur(FEClientes.TextMatrix(pnRow, 5)) Then
                MsgBox "EL DEPÓSITO supera al monto establecido por el Cheque.", vbInformation, "!Aviso¡"
                FEClientes.TextMatrix(pnRow, 5) = "0.00"
                cargarTotales
                Exit Sub
            ElseIf CCur(lblMonTra) < CCur(lblTotalME) Then
                MsgBox "SUMA TOTAL supera al monto establecido por el Voucher.", vbInformation, "!Aviso¡"
                FEClientes.TextMatrix(pnRow, 5) = "0.00"
                cargarTotales
                Exit Sub
            ElseIf CCur(lblMonTra) = CCur(lblTotalME) Then
                cmdAgregar.Enabled = False
                cmdRuta.Enabled = False
                cmdCargar.Enabled = False
            ElseIf CCur(lblMonTra) > CCur(lblTotalME) Then
                cmdAgregar.Enabled = True
                cmdRuta.Enabled = True
                cmdCargar.Enabled = True
            End If
        ElseIf fnOpeCod = gCTSDepLotCargoCta Then 'CTI4 ERS0112020
            If Not clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, CCur(FEClientes.TextMatrix(pnRow, 5))) Then
                MsgBox "EL DEPÓSITO supera al saldo de la cuenta a cargar.", vbInformation, "!Aviso¡"
                FEClientes.TextMatrix(pnRow, 5) = "0.00"
                cargarTotales
                Exit Sub
            ElseIf Not clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, CCur(lblTotalME)) Then
                MsgBox "SUMA TOTAL supera al saldo de la cuenta a cargar.", vbInformation, "!Aviso¡"
                FEClientes.TextMatrix(pnRow, 5) = "0.00"
                cargarTotales
                Exit Sub
            ElseIf clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, CCur(lblTotalMN)) Then
                cmdAgregar.Enabled = True
                cmdRuta.Enabled = True
                cmdCargar.Enabled = True
            Else
                cmdAgregar.Enabled = False
                cmdRuta.Enabled = False
                cmdCargar.Enabled = False
            End If
        End If
    
    End If

End If
    

End Sub
'JUEZ 20141014 *************************************
Private Function ValidaDepCTS() As Boolean
Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales
Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim rsPar As ADODB.Recordset, rs As ADODB.Recordset
Dim sCtaCod As String
Dim bValida As Boolean
    
    ValidaDepCTS = True
    lbVistoVal = False
    sCtaCod = FEClientes.TextMatrix(FEClientes.row, 3)
    
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rs = clsCap.GetDatosCuenta(sCtaCod)

    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set rsPar = clsDef.GetCapParametroNew(fnProducto, IIf(rs("nTpoPrograma"), 0, rs("nTpoPrograma")))
    If fnProducto = gCapCTS Then
        nParCantDepAnio = rsPar!nCantOpeDepAnio
        nParDiasVerifRegSueldo = rsPar!nDiasVerifUltRegSueldo
    End If
    Set rsPar = Nothing

    If bValidaCantDep Then
        Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
            nCantDepCta = clsCap.ObtenerCantidadOperaciones(sCtaCod, gCapMovDeposito, gdFecSis)
        Set clsCap = Nothing

        If nCantDepCta >= nParCantDepAnio And Not lbVistoVal Then
            If fnProducto = gCapCTS Then
                If MsgBox("Se ha realizado el número máximo de depósitos CTS, se requiere de VB del supervisor para grabar la operación. Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                    FEClientes.EliminaFila FEClientes.row
                    ValidaDepCTS = False
                    Exit Function
                End If
                Set loVistoElectronico = New frmVistoElectronico

                lbVistoVal = loVistoElectronico.Inicio(3, fnOpeCod)

                If Not lbVistoVal Then
                    FEClientes.EliminaFila FEClientes.row
                    ValidaDepCTS = False
                    Exit Function
                End If
            End If
        End If
    End If
End Function
'END JUEZ ******************************************

Private Sub FEClientes_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    devolverCTSPersona psDataCod
End Sub

Private Sub Form_Load()
cargarTipoCambio
cargarCTSPeriodo
IniciaCombo cboTransferMoneda, gMoneda
IniciaCombo cboCheque, gMoneda
IniciaCombo cboEfectivo, gMoneda
If fnOpeCod = gCTSDepLotEfec Then CargaControles 'CTI4 ERS0112020
End Sub

Private Sub FEClientes_OnChangeCombo() 'JUEZ 20160509
    If FEClientes.Col = 3 Then
        FEClientes_OnCellChange FEClientes.row, FEClientes.Col
    End If
End Sub
'CTI4 ERS0112020 *****************
Private Sub CmbForPag_Click()
    EstadoFormaPago IIf(CmbForPag.ListIndex = -1, -1, CInt(Trim(Right(IIf(CmbForPag.Text = "", "-1", CmbForPag.Text), 10))))
    If CmbForPag.ListIndex <> -1 Then
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher Then
     
'            Dim oformVou As New frmCapRegVouDepBus
'            Dim lnTipMot As Integer
'            Dim sGlosa As String
'            Dim sIF As String
'            Dim sVaucher As String
'            Dim sPersCod As String
'            Dim sNombre As String
'            Dim sDireccion As String
'            Dim sDocumento As String
'
'            lnTipMot = 15 ' Cancelacion Credito Pignoraticio
'            oformVou.iniciarFormularioDeposito CInt(Mid(AXCodCta.NroCuenta, 9, 1)), lnTipMot, sGlosa, sIF, sVaucher, nMontoVoucher, sPersCod, nMovNroRVD, nMovNroRVDPend, sNombre, sDireccion, sDocumento, AXCodCta.NroCuenta
'            If Len(sVaucher) = 0 Then Exit Sub
'            LblNumDoc.Caption = Trim(Mid(sVaucher, 1, Len(sVaucher) - 10))
'            cmdGrabar.Enabled = True
        ElseIf CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            Dim sCuenta As String
            fnOpeCod = gCTSDepLotCargoCta
            FREfectivo.Visible = False
            FRGlosa.Visible = False
            cmdCargar.Enabled = False
            cmdAgregar.Enabled = False
            fraFormaPago.Width = 7785
            If fnOpeCod <> gCTSDepLotCargoCta Then
                sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(gCTSDepLotCargoCta), sNumTarj, 232, False)
                If Val(Mid(sCuenta, 6, 3)) <> "232" And sCuenta <> "" Then
                    MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                    Exit Sub
                End If
            End If
            If Len(sCuenta) = 0 And fnOpeCod <> gCTSDepLotCargoCta Then txtCuentaCargo.SetFocusAge: Exit Sub
            If Len(Trim(txtCuentaCargo.NroCuenta)) = 18 Then
                txtCuentaCargo.NroCuenta = sCuenta
                txtCuentaCargo.Enabled = False
                cmdCargar.Enabled = True
                cmdAgregar.Enabled = True
                fnMoneda = CInt(Mid(sCuenta, 9, 1))
            Else
                txtCuentaCargo.Prod = "232"
            End If
            If Len(sCuenta) = 0 Then txtCuentaCargo.SetFocusAge: Exit Sub
            txtCuentaCargo.NroCuenta = sCuenta
            txtCuentaCargo.Enabled = False
            cmdCargar.Enabled = True
            cmdAgregar.Enabled = True
            fnMoneda = CInt(Mid(sCuenta, 9, 1))
        ElseIf CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoEfectivo Then
            fnOpeCod = gCTSDepLotEfec
            FREfectivo.Visible = True
            FRGlosa.Visible = True
            cmdCargar.Enabled = True
            cmdAgregar.Enabled = True
            fnMoneda = Right(cboEfectivo, 3)
            fraFormaPago.Width = 4455
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
            
        Case gColocTipoPagoEfectivo
            txtCuentaCargo.Visible = False
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            lblNroDocumento.Visible = False
            
        Case gColocTipoPagoCargoCta
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargo.Visible = True
            txtCuentaCargo.Enabled = True
            txtCuentaCargo.CMAC = gsCodCMAC
            txtCuentaCargo.Prod = Trim(Str(gCapAhorros))
            
        Case gColocTipoPagoVoucher
            LblNumDoc.Visible = True
            lblNroDocumento.Visible = True
            txtCuentaCargo.Visible = False
            
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
'    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher And Len(Trim(LblNumDoc.Caption)) = 0 Then
'        MsgBox "No se ha seleccionado el voucher correctamente. Verifique.", vbInformation, "¡Aviso!"
'        CmbForPag.SetFocus
'        Exit Function
'    End If
'    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher And Len(Trim(LblNumDoc.Caption)) > 0 _
'        And CCur(LblMontoPagar.Caption) > CCur(nMontoVoucher) Then
'        MsgBox "No se puede realizar el Pago con Voucher solo dispone de: " & Format(nMontoVoucher, "#0.00") & ". Verifique.", vbInformation, "¡Aviso!"
'        CmbForPag.SetFocus
'        Exit Function
'    End If
    
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta And Len(txtCuentaCargo.NroCuenta) <> 18 Then
        MsgBox "Favor de ingresar la cuenta de ahorros a debitar", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
        
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
        If Not clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, CDbl(IIf(fnMoneda = gMonedaNacional, (lblTotalMN), (lblTotalME)))) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
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
                        nRespuesta = oSolicitud.SolicitarVistoAtencionSinTarjeta(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaDepositoLoteCTS))
                    
                        If nRespuesta = 1 Then '1:Tiene Visto de atencion sin tarjeta pendiente de autorizar
                             MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                             cmdGuardar.Enabled = False
                             Exit Sub
                        End If
                        If nRespuesta = 2 Then '2:Tiene visto aceptado
                            MsgBox "La solicitud de atención sin tarjeta fue Aprobada, proceda con la atención", vbInformation, "Aviso"
                        End If
                        If nRespuesta = 3 Then '3:Tiene visto rechazado
                           If MsgBox("La solicitud de atención sin tarjeta fue RECHAZADA. ¿Desea realizar una nueva solicitud?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                Call oSolicitud.RegistrarVistoDeUsuario(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaDepositoLoteCTS))
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
                        
                                bExitoSol = frmCapMotivoAutorizacion.Inicio(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaDepositoLoteCTS))
                                If bExitoSol > 0 Then
                                    MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                                    cmdGuardar.Enabled = False
                                End If
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        lbVistoVal = loVistoElectronico2.Inicio(5, gAhoCargoCtaDepositoLoteCTS)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema"
                            Exit Sub
                        End If
                    End If
                ElseIf lsTieneTarj = "NO" And rsV.RecordCount > 0 Then
                    If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? ", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg ers 065-2017
                        lbVistoVal = loVistoElectronico2.Inicio(5, gAhoCargoCtaDepositoLoteCTS)
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
        fnMoneda = CInt(Mid(txtCuentaCargo.NroCuenta, 9, 1))
        cmdCargar.Enabled = True
        cmdAgregar.Enabled = True
        cmdCargar.SetFocus
    End If
End Sub
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R = oCons.RecuperaConstantes(gColocTipoPago, , , 3)
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, CmbForPag)
    Set loVistoElectronico2 = New frmVistoElectronico
    CmbForPag.ListIndex = 0
    Exit Sub
ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub
'END CTI4


