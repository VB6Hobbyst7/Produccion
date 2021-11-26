VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRHUsuarios 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   Icon            =   "frmRHUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwU 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "frmRHUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsUsuario As String

Public Function Ini() As String
    Me.Show 1
    Ini = lsUsuario
End Function

Private Sub cmdAceptar_Click()
    lsUsuario = Me.lvwU.SelectedItem
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
    lsUsuario = ""
End Sub

Private Sub Form_Load()
    Dim llAux As ListItem
    Dim lnI As Integer
    Dim lnTope As Integer
    
    GetUsers gcPDC
        
    lnTope = UBound(Users)
    lvwU.HideColumnHeaders = False
    lvwU.ColumnHeaders.Add , , "Usuario", 1000
    lvwU.ColumnHeaders.Add , , "Nombre", 4500
    lvwU.View = lvwReport
    
    For lnI = 1 To lnTope
        If Len(Users(lnI)) <= 4 Then
            Set llAux = lvwU.ListItems.Add(, , Users(lnI))
            llAux.SubItems(1) = UCase(NombreCompleto(Users(lnI), gcDominio))
        End If
    Next lnI
End Sub

Private Sub lvwU_DblClick()
    cmdAceptar_Click
End Sub

Private Sub lvwU_KeyPress(KeyAscii As Integer)
    cmdAceptar_Click
End Sub
