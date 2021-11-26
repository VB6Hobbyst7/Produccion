VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmCuadreInteres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuadre de Interes"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   705
      Left            =   0
      TabIndex        =   14
      Top             =   4080
      Width           =   7635
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   315
         Left            =   6480
         TabIndex        =   17
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   1230
         TabIndex        =   16
         Top             =   270
         Width           =   945
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   150
         TabIndex        =   15
         Top             =   270
         Width           =   945
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2355
      Left            =   0
      TabIndex        =   13
      Top             =   1740
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   4154
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      Height          =   1725
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7605
      Begin VB.Frame Frame4 
         ClipControls    =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   3615
         Begin VB.OptionButton ChkDolares 
            Caption         =   "Dolares"
            Height          =   255
            Left            =   1440
            TabIndex        =   20
            Top             =   150
            Width           =   1095
         End
         Begin VB.OptionButton ChkSoles 
            Caption         =   "Soles"
            Height          =   195
            Left            =   150
            TabIndex        =   19
            Top             =   180
            Width           =   1065
         End
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Migrar a Excel"
         Height          =   315
         Left            =   6060
         TabIndex        =   12
         Top             =   990
         Width           =   1425
      End
      Begin VB.Frame Frame2 
         Height          =   765
         Left            =   6060
         TabIndex        =   9
         Top             =   180
         Width           =   1425
         Begin VB.OptionButton OptNormal 
            Caption         =   "Normal"
            Height          =   315
            Left            =   150
            TabIndex        =   11
            Top             =   420
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptRFA 
            Caption         =   "RFA"
            Height          =   225
            Left            =   150
            TabIndex        =   10
            Top             =   210
            Width           =   855
         End
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   345
         Left            =   4080
         TabIndex        =   8
         Top             =   810
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   1110
         TabIndex        =   6
         Top             =   810
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboPago 
         Height          =   315
         Left            =   4050
         TabIndex        =   4
         Top             =   270
         Width           =   1875
      End
      Begin VB.ComboBox CboProducto 
         Height          =   315
         Left            =   930
         TabIndex        =   2
         Top             =   270
         Width           =   1875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   3030
         TabIndex        =   7
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   870
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Pago"
         Height          =   195
         Left            =   2970
         TabIndex        =   3
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   645
      End
   End
End
Attribute VB_Name = "FrmCuadreInteres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CargarControles()
    'Adicionando los productos
    CboProducto.AddItem "Comercial Empresarial"
    CboProducto.AddItem "Comercial Agricola"
    CboProducto.AddItem "Mes Empresarial"
    CboProducto.AddItem "Mes Agricola"
    CboProducto.AddItem "Consumo"
    
    CboPago.AddItem "Pagos Normales"
    CboPago.AddItem "Refinanciados Normales"
    CboPago.AddItem "Pagos Vencidos"
    CboPago.AddItem "Judiciales"
    
End Sub
