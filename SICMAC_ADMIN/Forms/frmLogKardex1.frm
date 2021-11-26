VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogKardex1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "frmLogKardex1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   5850
      TabIndex        =   10
      Top             =   5520
      Width           =   1170
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   7095
      TabIndex        =   9
      Top             =   5520
      Width           =   1170
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   360
      Left            =   4590
      TabIndex        =   8
      Top             =   5520
      Width           =   1170
   End
   Begin SicmactAdmin.FlexEdit FlexEdit1 
      Height          =   4305
      Left            =   30
      TabIndex        =   1
      Top             =   1155
      Width           =   8205
      _ExtentX        =   14526
      _ExtentY        =   7435
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
      Appearance      =   0
      RowHeight0      =   240
   End
   Begin VB.Frame fraProducto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Producto"
      ForeColor       =   &H00800000&
      Height          =   1080
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   8205
      Begin VB.ComboBox cmdAlmacen 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000006&
         Height          =   315
         Left            =   150
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   630
         Width           =   3015
      End
      Begin SicmactAdmin.TxtBuscar txtProducto 
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   300
         Left            =   4290
         TabIndex        =   2
         Top             =   637
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFecFin 
         Height          =   300
         Left            =   6780
         TabIndex        =   5
         Top             =   637
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecFin 
         Caption         =   "Fecha Fin."
         Height          =   195
         Left            =   5760
         TabIndex        =   7
         Top             =   690
         Width           =   915
      End
      Begin VB.Label lblFecIni 
         Caption         =   "Fecha Ini."
         Height          =   195
         Left            =   3315
         TabIndex        =   6
         Top             =   690
         Width           =   915
      End
      Begin VB.Label lblProducto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2025
         TabIndex        =   4
         Top             =   270
         Width           =   6060
      End
   End
End
Attribute VB_Name = "frmLogKardex1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdProcesar_Click()
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen

    Me.FlexDetalle.rsTextBuscar = oAlmacen.GetBienesAlmacen()
End Sub

Private Sub Form_Load()
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    
    Me.FlexDetalle.rsTextBuscar = oAlmacen.GetBienesAlmacen()
End Sub

Private Sub txtProducto_EmiteDatos()
    lblProducto.Caption = txtProducto.psDescripcion
End Sub
