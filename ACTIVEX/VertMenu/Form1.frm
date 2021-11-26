VERSION 5.00
Object = "{5F774E03-DB36-4DFC-AAC4-D35DC9379F2F}#1.0#0"; "VertMenu.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   1050
   ClientTop       =   1770
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7200
   Begin VB.CommandButton Command2 
      Caption         =   "false"
      Height          =   555
      Left            =   3975
      TabIndex        =   2
      Top             =   1245
      Width           =   2595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "true"
      Height          =   510
      Left            =   4020
      TabIndex        =   1
      Top             =   285
      Width           =   2595
   End
   Begin VertMenu.VerticalMenu VerticalMenu1 
      Height          =   4815
      Left            =   270
      TabIndex        =   0
      Top             =   195
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   8493
      MenusMax        =   4
      MenuCaption1    =   "Menu1"
      MenuItemIcon11  =   "Form1.frx":0000
      MenuCaption2    =   "Menu2"
      MenuItemIcon21  =   "Form1.frx":0452
      MenuCaption3    =   "Menu3"
      MenuItemIcon31  =   "Form1.frx":08A4
      MenuCaption4    =   "Menu4"
      MenuItemIcon41  =   "Form1.frx":0CF6
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
VerticalMenu1.Enabled = True
End Sub

Private Sub Command2_Click()
VerticalMenu1.Enabled = False
End Sub

Private Sub VerticalMenu1_ValidaMenuItem(pnItem As Long, Cancel As Boolean)
If pnItem = 1 Or pnItem = 2 Then
    Cancel = False
End If
End Sub
