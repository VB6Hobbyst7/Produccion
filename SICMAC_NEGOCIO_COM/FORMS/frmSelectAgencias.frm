VERSION 5.00
Begin VB.Form frmSelectAgencias 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4500
   ClientLeft      =   2925
   ClientTop       =   2235
   ClientWidth     =   4215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   915
      TabIndex        =   7
      Top             =   390
      Width           =   3120
      Begin VB.OptionButton OptAnalista 
         Caption         =   "&Todos"
         Height          =   210
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   195
         Width           =   915
      End
      Begin VB.OptionButton OptAnalista 
         Caption         =   "&Ninguno"
         Height          =   210
         Index           =   1
         Left            =   1620
         TabIndex        =   4
         Top             =   195
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   75
      ScaleHeight     =   300
      ScaleWidth      =   4095
      TabIndex        =   6
      Top             =   75
      Width           =   4095
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4080
      Left            =   75
      ScaleHeight     =   4080
      ScaleWidth      =   735
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdAceptar 
      Cancel          =   -1  'True
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
      Height          =   405
      Left            =   1965
      TabIndex        =   2
      Top             =   4005
      Width           =   1140
   End
   Begin VB.Frame fraContenedor 
      Caption         =   "Haga un Click en la Agencia a escoger "
      Height          =   2895
      Index           =   3
      Left            =   915
      TabIndex        =   0
      Top             =   975
      Width           =   3165
      Begin VB.ListBox List1 
         Height          =   2535
         ItemData        =   "frmSelectAgencias.frx":0000
         Left            =   105
         List            =   "frmSelectAgencias.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   255
         Width           =   2925
      End
   End
End
Attribute VB_Name = "frmSelectAgencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNomFrm As Form
Dim x As Integer

Private Sub CmdAceptar_Click()
    Me.Hide
End Sub

Public Function RecupAgencias() As String
Dim i As Integer
    RecupAgencias = "("
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
            RecupAgencias = RecupAgencias & "'" & Left(List1.List(i), 2) & "',"
        End If
    Next i
    RecupAgencias = Mid(RecupAgencias, 1, Len(RecupAgencias) - 1) & ")"
End Function

Private Sub Form_Load()
    'RotateText 90, Pic1, "Times New Roman", 18, 150, 2550, " AGENCIAS "
    'RotateText 0, Pic2, "Times New Roman", 10, 740, 10, " A      I N T E R C O N E C T A R "
    
    CentraForm Me
    OptAnalista(1).value = True
    Dim loCargaAg As COMDColocPig.DCOMColPFunciones
    Dim lrAgenc As ADODB.Recordset
    Set loCargaAg = New COMDColocPig.DCOMColPFunciones
        Set lrAgenc = loCargaAg.dObtieneAgencias(True)
    Set loCargaAg = Nothing
    If lrAgenc Is Nothing Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        Me.List1.Clear
        With lrAgenc
            Do While Not .EOF
                List1.AddItem !cAgeCod & " " & Trim(!cAgeDescripcion)
                If !cAgeCod = gsCodAge Then
                    List1.Selected(List1.ListCount - 1) = True
                End If
                .MoveNext
            Loop
        End With
    End If
End Sub

Public Sub Inicio(pNomFrm As Form)
    Set vNomFrm = pNomFrm
End Sub

Private Sub OptAnalista_Click(Index As Integer)
Dim bCheck As Boolean
Dim i As Integer
    If Index = 0 Then
        bCheck = True
    Else
        bCheck = False
    End If
    If List1.ListCount <= 0 Then
        Exit Sub
    End If
    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = bCheck
    Next i
End Sub
