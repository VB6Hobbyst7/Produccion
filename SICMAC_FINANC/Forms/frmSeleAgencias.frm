VERSION 5.00
Begin VB.Form frmSeleAgencias 
   Caption         =   "Selección de Agencias"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4905
   Icon            =   "frmSeleAgencias.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTodo 
      Caption         =   "Todas las Agencias"
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1725
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3720
      TabIndex        =   2
      Top             =   4140
      Width           =   1065
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   2610
      TabIndex        =   1
      Top             =   4140
      Width           =   1065
   End
   Begin VB.Frame fraAge 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   4665
      Begin VB.ListBox lstAge 
         Height          =   3210
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   600
         Width           =   4425
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A G E N C I A S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   120
         TabIndex        =   4
         Top             =   210
         Width           =   4425
      End
   End
End
Attribute VB_Name = "frmSeleAgencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs   As ADODB.Recordset
Dim lbConsolida As Boolean
Dim lOk As Boolean

Public Sub Inicio(pbConsolida As Boolean)
lbConsolida = pbConsolida
Me.Show 1
End Sub

Private Sub chkTodo_Click()
Dim k As Integer
For k = 0 To lstAge.ListCount - 1
   If chkTodo.value = vbcheked Then
      lstAge.Selected(k) = True
   Else
      lstAge.Selected(k) = False
   End If
Next
End Sub

Private Sub cmdAceptar_Click()
lOk = True
Me.Hide
End Sub

Private Sub cmdCancelar_Click()
lOk = False
Me.Hide
End Sub

Private Sub Form_Load()
CentraForm Me
Dim oAge As New DActualizaDatosArea
Set rs = oAge.GetAgencias()
Do While Not rs.EOF
   lstAge.AddItem rs!Codigo & " " & rs!Descripcion
   rs.MoveNext
Loop
If lbConsolida Then
   lstAge.AddItem "CONSOLIDADO"
End If
Set oAge = Nothing
RSClose rs
End Sub

Public Property Get pbOk() As Boolean
pbOk = lOk
End Property

Public Property Let pbOk(ByVal vNewValue As Boolean)
lOk = vNewValue
End Property

Private Sub lstAge_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdAceptar_Click
End If
End Sub
