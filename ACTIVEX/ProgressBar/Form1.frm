VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   2130
   ClientTop       =   2640
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   7200
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   390
      Left            =   3405
      TabIndex        =   1
      Top             =   2235
      Width           =   2340
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "Command1"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   975
      Width           =   2835
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a  As clsProgressBar
Private Sub Command1_Click()
Set a = New clsProgressBar

a.ShowForm Form1
a.CaptionSyle = eCap_CaptionPercent
DoEvents
a.Max = 5000
For i = 1 To a.Max
    a.Progress i, "Ejemplo", "Ejemplo de subtitulo", "Ejemplo", vbBlue
Next
a.CloseForm Me

End Sub

Private Sub Command2_Click()

 Load Form2

'SetParent Form2.hWnd, Me.hWnd


End Sub

Private Sub Form_Load()
Set a = New clsProgressBar
'MDIForm1.Enabled = False
'Dim lResult As Long
'lResult = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
lngOrigParenthWnd = SetWindowWord(Me.hwnd, -8, MDIForm1.hwnd)


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim lngResult&
lngResult = SetWindowWord(Me.hwnd, -8, lngOrigParenthWnd)
'MDIForm1.Enabled = True
End Sub
