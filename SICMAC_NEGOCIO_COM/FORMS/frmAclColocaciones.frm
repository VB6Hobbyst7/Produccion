VERSION 5.00
Begin VB.Form frmAclColocaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Genera Tablas ACL de Colocaciones"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmAclColocaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4650
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkCarta 
      Caption         =   "Carta Fianza"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CheckBox ChkCreditos 
      Caption         =   "Creditos"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Pasa los datos de Creditos a archivos DBF, para ser usados en el programa de Auditoria ACL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmAclColocaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim loAcl  As ACL
Attribute loAcl.VB_VarHelpID = -1
Dim Progress As clsProgressBar

Private Sub cmdGenerar_Click()
Dim loAcl As New ACL
If Me.ChkCarta.value = 1 Or Me.ChkCreditos.value = 1 Then
    If Me.ChkCarta.value = 1 Then
        loAcl.Coloc_ACL_CartaFianza
    End If
    If Me.ChkCreditos.value = 1 Then
        loAcl.Coloc_ACL_Creditos
    End If
Else
    MsgBox "Elija una de las opciones", vbInformation, "AVISO"
End If
End Sub

Private Sub Form_Load()
Set Progress = New clsProgressBar
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub loAcl_CloseProgress()
Progress.CloseForm Me
End Sub

Private Sub loAcl_Progress(pnValor As Long, pnTotal As Long)
Progress.Max = pnTotal
Progress.Progress pnValor, "Generando Reporte", "Procesando ..."
End Sub

Private Sub loAcl_ShowProgress()
Progress.ShowForm Me
End Sub
