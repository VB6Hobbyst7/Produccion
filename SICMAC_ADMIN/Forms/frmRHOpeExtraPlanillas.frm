VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRHOpeExtraPlanillas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "frmRHOpeExtraPlanillas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   30
      TabIndex        =   8
      Top             =   4965
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1230
      TabIndex        =   7
      Top             =   4965
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2430
      TabIndex        =   6
      Top             =   4965
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3630
      TabIndex        =   5
      Top             =   4965
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7710
      TabIndex        =   4
      Top             =   4965
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4485
      Left            =   45
      TabIndex        =   2
      Top             =   435
      Width           =   8730
      Begin Sicmact.FlexEdit FlexEdit1 
         Height          =   4185
         Left            =   105
         TabIndex        =   3
         Top             =   225
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   7382
         HighLight       =   1
         AllowUserResizing=   3
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
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1485
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   135
      Width           =   3405
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   7005
      TabIndex        =   0
      Top             =   120
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   582
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   30
      TabIndex        =   9
      Top             =   4965
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1230
      TabIndex        =   10
      Top             =   4965
      Width           =   1095
   End
End
Attribute VB_Name = "frmRHOpeExtraPlanillas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSalir_Click()
    Unload Me
End Sub
