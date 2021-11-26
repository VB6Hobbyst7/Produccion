VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNotaCargoAbono 
   Caption         =   "Form2"
   ClientHeight    =   5325
   ClientLeft      =   705
   ClientTop       =   1920
   ClientWidth     =   8730
   LinkTopic       =   "Form2"
   ScaleHeight     =   5325
   ScaleWidth      =   8730
   Begin VB.Frame fraIngNotaCargo 
      Caption         =   "Nota de Cargo"
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
      Height          =   1815
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   8385
      Begin VB.TextBox txtNotaCargo 
         Height          =   315
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   315
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtFechaNC 
         Height          =   345
         Left            =   6870
         TabIndex        =   2
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label13 
         Caption         =   "Número"
         Height          =   225
         Left            =   4020
         TabIndex        =   7
         Top             =   345
         Width           =   585
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   6330
         TabIndex        =   6
         Top             =   345
         Width           =   555
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Datos Titular :"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   923
         Width           =   990
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3000
         TabIndex        =   4
         Top             =   855
         Width           =   5010
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1245
         TabIndex        =   3
         Top             =   855
         Width           =   1680
      End
   End
End
Attribute VB_Name = "frmNotaCargoAbono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

