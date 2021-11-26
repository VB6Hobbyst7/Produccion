VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGarantExtorno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno de Garantia Real"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdExtornar 
      Caption         =   "&Extornar"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   320
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Extorno Garantia"
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton OptExt 
         Caption         =   "Usuario"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptExt 
         Caption         =   "Todas"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4048
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmGarantExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Nrog As String
Dim NroMov As String
Private Sub cmdBuscar_Click()
Nrog = ""
NroMov = ""
msh.Clear
msh.ClearStructure
msh.Rows = 2
If OptExt(0).value Then
    CargaExtornos (0)
Else
    If Len(txtUser) = 4 Then
        Call CargaExtornos(1, UCase(txtUser))
    Else
        MsgBox "Nombre de Usuario Incompleto", vbInformation, "AVISO"
    End If
End If
Marco
End Sub
Sub CargaExtornos(ByVal opt As Integer, Optional sUser As String)
Dim NG As COMDCredito.DCOMGarantia
Dim rs As ADODB.Recordset
Dim sFecha As String
msh.Clear
msh.ClearStructure

Set NG = New COMDCredito.DCOMGarantia
Set rs = New ADODB.Recordset
sFecha = Mid(gdFecSis, 7, 4) & Mid(gdFecSis, 4, 2) & Mid(gdFecSis, 1, 2)
CmdExtornar.Visible = False
Set rs = NG.RecuperaExtornoGarantReal(sFecha, opt, sUser)
If rs.EOF And rs.BOF Then
Else
    Set msh.DataSource = rs
    
End If
Marco
End Sub
Sub Marco()
With msh
    .ColWidth(0) = 1300
    .ColWidth(1) = 1600
    .ColWidth(2) = 3000
End With
End Sub
Private Sub cmdExtornar_Click()
Dim OptBt As Integer
Dim Dg As COMDCredito.DCOMGarantia
Set Dg = New COMDCredito.DCOMGarantia
OptBt = MsgBox("Esta Seguro de Extornar la operacion", vbQuestion + vbYesNo, "AVISO")

If vbYes = OptBt Then
    Call Dg.ExtornaGarantiaReal(Nrog, NroMov)
    CmdExtornar.Visible = False
    Nrog = ""
    NroMov = ""
    msh.Rows = 2
    If OptExt(0).value Then
        CargaExtornos (0)
    Else
        Call CargaExtornos(1, UCase(txtUser))
    End If
Else
End If
CmdExtornar.Visible = False
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Nrog = ""
NroMov = ""
Marco
End Sub

Private Sub msh_Click()
If msh.Row <> 0 Then
    
    Nrog = msh.TextMatrix(msh.Row, 0)
    NroMov = msh.TextMatrix(msh.Row, 2)
    CmdExtornar.Visible = True
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End If
End Sub

Private Sub OptExt_Click(Index As Integer)

Select Case Index
    Case 0: txtUser.Visible = False
    Case 1: txtUser.Visible = True
End Select
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    If Len(txtUser) = 4 Then
        Call CargaExtornos(1, UCase(txtUser))
    Else
        MsgBox "Nombre de Usuario Incompleto", vbInformation, "AVISO"
    End If

 Else
    KeyAscii = Letras(KeyAscii)
 End If
 
End Sub
Public Function Letras(intTecla As Integer) As Integer
    Letras = Asc(UCase(Chr(SoloLetras(intTecla))))
End Function

Public Function SoloLetras(intTecla As Integer) As Integer
Dim cValidar  As String
    cValidar = "0123456789+:;'<>?_=+[]{}|!@#$%^&()*/ ·¿¨Çº-.,"
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) <> 0 Then
            intTecla = 0
            Beep
        End If
    End If
    SoloLetras = intTecla
End Function
