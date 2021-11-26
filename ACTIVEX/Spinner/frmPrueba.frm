VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   1560
   ClientTop       =   2475
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   7200
   Begin Spinner.uSpinner Spinner1 
      Height          =   420
      Left            =   1140
      TabIndex        =   1
      ToolTipText     =   "ejemplo "
      Top             =   660
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   741
      Max             =   9999
      Min             =   2000
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   4140
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1230
      Width           =   1035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Spinner1.Valor = 4500

End Sub

Private Sub Spinner1_Change()
Text1 = Spinner1.Valor
End Sub

Private Sub Spinner1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Text1.SetFocus
End If
End Sub

