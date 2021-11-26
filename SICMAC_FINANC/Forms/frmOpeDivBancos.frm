VERSION 5.00
Begin VB.Form frmOpeDivBancos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5760
   ClientLeft      =   1380
   ClientTop       =   2085
   ClientWidth     =   9195
   Icon            =   "frmOpeDivBancos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraOrigen 
      Caption         =   "Origen"
      Height          =   1560
      Left            =   60
      TabIndex        =   5
      Top             =   765
      Width           =   9045
   End
   Begin VB.Frame FrameTipCambio 
      Caption         =   "Tipo de Cambio"
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
      Height          =   675
      Left            =   6405
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   2685
      Begin VB.TextBox txtTipCambio 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   435
         TabIndex        =   2
         Top             =   240
         Width           =   690
      End
      Begin VB.TextBox txtTCBanco 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1785
         TabIndex        =   1
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lblTipCambio 
         Caption         =   "Fijo"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   435
      End
      Begin VB.Label lblvariable 
         Caption         =   "Banco"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   300
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmOpeDivBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
CentraForm Me
Me.Caption = gsOpeDesc

End Sub
