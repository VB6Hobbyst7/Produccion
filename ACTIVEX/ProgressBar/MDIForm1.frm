VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5325
   ClientLeft      =   1575
   ClientTop       =   2325
   ClientWidth     =   7200
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuejemplo 
      Caption         =   "ejemplo"
   End
   Begin VB.Menu adsfasdf 
      Caption         =   "asdfasdf"
      Begin VB.Menu asdffff 
         Caption         =   "asdfasdf"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Activate()

'frmChild.Show

End Sub

Private Sub mnuejemplo_Click()
Form1.Show
'SetParent Form1.hwnd, Me.hwnd
'Me.Enabled = False
End Sub
