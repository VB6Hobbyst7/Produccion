VERSION 5.00
Begin VB.Form FrmSector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion de Sector Economico"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Sector"
      Height          =   4380
      Left            =   0
      TabIndex        =   4
      Top             =   510
      Width           =   5040
      Begin VB.ListBox LstSector 
         Height          =   3885
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   270
         Width           =   4890
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3510
      TabIndex        =   3
      Top             =   4980
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5040
      Begin VB.OptionButton OptSector 
         Caption         =   "&Todos"
         Height          =   210
         Index           =   0
         Left            =   900
         TabIndex        =   2
         Top             =   210
         Width           =   1485
      End
      Begin VB.OptionButton OptSector 
         Caption         =   "&Ninguno"
         Height          =   210
         Index           =   1
         Left            =   2970
         TabIndex        =   1
         Top             =   210
         Value           =   -1  'True
         Width           =   1485
      End
   End
End
Attribute VB_Name = "FrmSector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CargaAnalistas()
End Sub

Private Sub CmdAceptar_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
Dim R As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
Dim sAnalistas As String
Dim ObjCons As DConstante

Set ObjCons = New DConstante
    On Error GoTo ERRORCargaAnalistas
    Set R = ObjCons.GetSector
    LstSector.Clear
    Do While Not R.EOF
        LstSector.AddItem R!cConsDescripcion & Space(100) & R!nConsValor
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
      Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    Exit Sub
ERRORCargaAnalistas:
    MsgBox Err.Description, vbCritical, "Aviso"

  
    
End Sub


Private Sub OptSector_Click(Index As Integer)
Dim bCheck As Boolean
Dim i As Integer
    If Index = 0 Then
        bCheck = True
    Else
        bCheck = False
    End If
   If LstSector.ListCount <= 0 Then
        Exit Sub
    End If
    For i = 0 To LstSector.ListCount - 1
        LstSector.Selected(i) = bCheck
    Next i

End Sub
