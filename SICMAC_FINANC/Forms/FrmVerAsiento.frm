VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmVerAsiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Asiento"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   645
      Left            =   90
      TabIndex        =   13
      Top             =   5220
      Width           =   7185
   End
   Begin Sicmact.FlexEdit FlexEdit1 
      Height          =   3675
      Left            =   60
      TabIndex        =   12
      Top             =   1560
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   6482
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
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
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   -1
      RowHeight0      =   240
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7245
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   345
         Left            =   6060
         TabIndex        =   11
         Top             =   990
         Width           =   1125
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   345
         Left            =   6060
         TabIndex        =   10
         Top             =   600
         Width           =   1125
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   345
         Left            =   6030
         TabIndex        =   9
         Top             =   210
         Width           =   1125
      End
      Begin Sicmact.ActXCodCta ActXCodCta1 
         Height          =   375
         Left            =   2220
         TabIndex        =   8
         Top             =   930
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   661
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   3060
         TabIndex        =   7
         Top             =   180
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   609
         _Version        =   393216
         Format          =   60817409
         CurrentDate     =   38476
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3060
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   570
         Width           =   2715
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1875
         Begin VB.CheckBox ChkProducto 
            Appearance      =   0  'Flat
            Caption         =   "Producto"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   150
            TabIndex        =   3
            Top             =   570
            Width           =   1485
         End
         Begin VB.CheckBox ChkCuentaContable 
            Appearance      =   0  'Flat
            Caption         =   "Cuenta Contable"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   150
            Width           =   1515
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2220
         TabIndex        =   6
         Top             =   180
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2190
         TabIndex        =   5
         Top             =   540
         Width           =   825
      End
   End
End
Attribute VB_Name = "FrmVerAsiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
