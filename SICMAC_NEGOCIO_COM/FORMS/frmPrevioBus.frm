VERSION 5.00
Begin VB.Form frmPrevioBus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar"
   ClientHeight    =   1485
   ClientLeft      =   2430
   ClientTop       =   2085
   ClientWidth     =   5835
   Icon            =   "frmPrevioBus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkOpc2 
      Caption         =   "Coi&ncidir mayúsculas y minúsculas"
      Height          =   195
      Left            =   210
      TabIndex        =   7
      Top             =   1140
      Width           =   2895
   End
   Begin VB.CheckBox chkOpc1 
      Caption         =   "Sólo palabras &completas"
      Height          =   225
      Left            =   210
      TabIndex        =   6
      Top             =   870
      Width           =   2355
   End
   Begin VB.ComboBox cboDireccion 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   495
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   4260
      TabIndex        =   3
      Top             =   795
      Width           =   1380
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "B&uscar"
      Height          =   360
      Left            =   4260
      TabIndex        =   2
      Top             =   240
      Width           =   1395
   End
   Begin VB.ComboBox cboBuscar 
      Height          =   315
      Left            =   900
      TabIndex        =   1
      Top             =   150
      Width           =   3180
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "&Dirección :"
      Height          =   225
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   555
      Width           =   795
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "&Buscar :"
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   195
      Width           =   675
   End
End
Attribute VB_Name = "frmPrevioBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdBuscar.SetFocus
End Sub

Private Sub cmdBuscar_Click()
'Dim X As Integer
'Dim pCboAgr As Boolean
'pCboAgr = False
'For X = 0 To cboBuscar.ListCount + 1
'    If cboBuscar.List(X) = cboBuscar.Text Then pCboAgr = True
'Next X
'If pCboAgr = False Then cboBuscar.AddItem cboBuscar.Text
If Len(Trim(cboBuscar.Text)) = 0 Then
    MsgBox "Ingrese cadena a Buscar ...", vbInformation, " Aviso "
Else
    frmPrevioBus.Visible = False
End If
End Sub

Private Sub cmdCancelar_Click()
frmPrevioBus.Visible = False
cmdCancelar.Enabled = False
End Sub

Private Sub Form_Load()
cboDireccion.AddItem "Todo"
cboDireccion.AddItem "Abajo"
cboDireccion.ListIndex = 0
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub
