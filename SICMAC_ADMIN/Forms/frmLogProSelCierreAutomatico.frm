VERSION 5.00
Begin VB.Form frmLogProSelCierreAutomatico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proceso de Seleccion"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   Icon            =   "frmLogProSelCierreAutomatico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar Etapas"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00EDFBFC&
      Caption         =   "Cierre Automatico de las Etapas de Procesos de Seleccion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   660
      Left            =   210
      TabIndex        =   2
      Top             =   240
      Width           =   4635
   End
End
Attribute VB_Name = "frmLogProSelCierreAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCerrar_Click()
    If CierraEtapaAutomatico Then _
        MsgBox "Cierre Completo de Etapas", vbInformation, "Aviso"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmLogProSelCierreAutomatico = Nothing
End Sub
