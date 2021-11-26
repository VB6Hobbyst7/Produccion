VERSION 5.00
Begin VB.Form frmImput 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "frmImput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Sicmact.EditMoney EditMoney1 
      Height          =   315
      Left            =   1410
      TabIndex        =   1
      Top             =   540
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   1725
      TabIndex        =   0
      Top             =   990
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   75
      Width           =   4875
   End
End
Attribute VB_Name = "frmImput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCaption As String
Dim lsMensaje As String
Dim lnMonto As Currency

Public Function Ini(psCaption As String, psMensaje As String) As Currency
    lsCaption = psCaption
    lsMensaje = psMensaje
    lnMonto = 0
    Me.Show 1
    Ini = lnMonto
End Function

Private Sub cmdAceptar_Click()
    lnMonto = Me.EditMoney1.value
    Unload Me
End Sub

Private Sub EditMoney1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAceptar.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    Me.EditMoney1.SetFocus
End Sub

Private Sub Form_Load()
    Me.Caption = lsCaption
    Me.Label1.Caption = lsMensaje
    Me.EditMoney1.value = 0
End Sub
