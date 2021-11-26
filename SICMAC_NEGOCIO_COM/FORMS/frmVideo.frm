VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmVideo 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   Caption         =   "LogOn"
   ClientHeight    =   5895
   ClientLeft      =   1560
   ClientTop       =   1785
   ClientWidth     =   7860
   Icon            =   "frmVideo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8085
      Top             =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   5955
      Left            =   0
      ScaleHeight     =   5895
      ScaleWidth      =   7800
      TabIndex        =   0
      Top             =   15
      Width           =   7860
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SHWVideo 
         Height          =   5850
         Left            =   15
         TabIndex        =   1
         Top             =   0
         Width           =   7800
         _cx             =   58201048
         _cy             =   58201048
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   0   'False
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         AllowScriptAccess=   "always"
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
      End
   End
End
Attribute VB_Name = "frmVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    SHWVideo.Loop = False
    SHWVideo.Movie = App.Path & "\avis\avilog.swf"
End Sub

Private Sub SHWVideo_GotFocus()
    Me.SetFocus
End Sub

Private Sub Timer1_Timer()
    If Not SHWVideo.IsPlaying Then
        Unload Me
    End If
End Sub
