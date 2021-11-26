VERSION 5.00
Begin VB.Form frmAdmCredListApoderados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Apoderados"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "frmAdmCredListApoderados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   300
      Left            =   3360
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame frmCapAperturas 
      Caption         =   " Usuarios "
      ForeColor       =   &H00FF0000&
      Height          =   2680
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4215
      Begin VB.ListBox lstUsuarios 
         Height          =   2310
         Left            =   150
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   240
         Width           =   3915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Nivel "
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.Label lblNivel 
         Caption         =   "lblNivel"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Label lblRequeridos 
      Caption         =   "Requeridos: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   3525
      Width           =   1335
   End
   Begin VB.Label lblSelecionados 
      Caption         =   "Seleccionados: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3525
      Width           =   1575
   End
End
Attribute VB_Name = "frmAdmCredListApoderados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmAdmCredListApoderados
'** Descripción : Formulario que lista los apoderados según el nivel
'** Creación    : RECO, 20150421 - ERS010-2015
'**********************************************************************************************
Option Explicit
Dim lnCanReq As Integer
Dim lsDatos() As Variant
Public Function Inicio(ByVal psTitulo As String, ByVal pnCanRequerido As Integer, ByVal prsDatos As ADODB.Recordset) As Variant
    Dim i As Integer
    lblNivel.Caption = Mid(psTitulo, 1, Len(psTitulo) - 6)
    lblRequeridos.Caption = lblRequeridos.Caption & pnCanRequerido
    lnCanReq = pnCanRequerido
    For i = 1 To prsDatos.RecordCount
        lstUsuarios.AddItem (prsDatos!cPersNombre & Space(50) & prsDatos!cUser)
        prsDatos.MoveNext
    Next
    Me.Show 1
    Inicio = lsDatos
End Function

Private Sub CmdAceptar_Click()
    Dim x As Integer
    Dim i As Integer
    ReDim Preserve lsDatos(2, 0 To CanSelec)
    For x = 0 To lstUsuarios.ListCount - 1
        If lstUsuarios.Selected(x) = True Then
          lstUsuarios.ListIndex = x
          lsDatos(1, i) = Right(lstUsuarios.Text, 4)
          lsDatos(2, i) = Trim(Mid(lstUsuarios.Text, 1, Len(lstUsuarios.Text) - 4))
          i = i + 1
        End If
    Next
    Unload Me
End Sub

Private Sub lstUsuarios_Click()
    lblSelecionados.Caption = "Seleccionados: " & CanSelec
    If CanSelec > lnCanReq Then
        MsgBox "No se puede seleccionar más elementos de los requeridos.", vbInformation, "Aviso"
        lstUsuarios.Selected(lstUsuarios.ListIndex) = False
    End If
End Sub

Private Function CanSelec() As Integer
    Dim x As Integer
    For x = 0 To lstUsuarios.ListCount - 1
        If lstUsuarios.Selected(x) = True Then
          CanSelec = CanSelec + 1
        End If
    Next
End Function

Private Sub lstUsuarios_DblClick()
    lblSelecionados.Caption = "Seleccionados: " & CanSelec
    If CanSelec > lnCanReq Then
        MsgBox "No se puede seleccionar más elementos de los requeridos.", vbInformation, "Aviso"
        lstUsuarios.Selected(lstUsuarios.ListIndex) = False
    End If
End Sub
