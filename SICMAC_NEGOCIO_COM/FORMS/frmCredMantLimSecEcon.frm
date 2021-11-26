VERSION 5.00
Begin VB.Form frmCredMantLimSecEcon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Límites por Sector"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   Icon            =   "frmCredMantLimSecEcon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   5640
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   360
      Left            =   4440
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Frame gbControles 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   80
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtLimite 
         Height          =   315
         Left            =   5640
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblSectorDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   5640
      TabIndex        =   0
      Top             =   5640
      Width           =   1050
   End
   Begin SICMACT.FlexEdit feSector 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   9340
      Cols0           =   4
      HighLight       =   1
      EncabezadosNombres=   "-cSectorCod-Sector-Limite (%)"
      EncabezadosAnchos=   "300-0-5000-900"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X"
      ListaControles  =   "0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-R"
      FormatosEdit    =   "0-1-1-2"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   3
      ColWidth0       =   300
      RowHeight0      =   300
   End
End
Attribute VB_Name = "frmCredMantLimSecEcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredMantLimSecEcon
'** Descripción : Formulario para administrar los límites de créditos por Sector económico
'**               creado segun TI-ERS029-2013
'** Creación : JUEZ, 20140530 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim oDPerGen As COMDPersona.DCOMPersGeneral
Dim rs As ADODB.Recordset
Dim fsSectorCod As String

Private Sub cmdCancelar_Click()
fsSectorCod = ""
lblSectorDesc.Caption = ""
txtLimite.Text = ""
CargarDatos
HabilitaControles False
End Sub

Private Sub CmdGrabar_Click()
If Trim(txtLimite.Text) = "" Then
    MsgBox "Debe ingresar el limite de crédito", vbInformation, "Aviso"
    Exit Sub
End If
Set oDPerGen = New COMDPersona.DCOMPersGeneral
Call oDPerGen.ActualizarLimiteSector(fsSectorCod, CDbl(txtLimite.Text))
Set oDPerGen = Nothing
MsgBox "Límite actualizado", vbInformation, "Aviso"
cmdCancelar_Click
End Sub

Public Sub inicia(ByVal pnTipo As Integer)
CargarDatos
If pnTipo = 1 Then
    cmdEditar.Visible = True
Else
    cmdEditar.Visible = False
End If
Me.Show 1
End Sub

Private Sub CargarDatos()
Dim lnFila As Integer

Set oDPerGen = New COMDPersona.DCOMPersGeneral
Set rs = oDPerGen.GetSectorEconomico
Set oDPerGen = Nothing
    Call LimpiaFlex(feSector)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feSector.AdicionaFila
            lnFila = feSector.row
            feSector.TextMatrix(lnFila, 1) = rs!cSectorCod
            feSector.TextMatrix(lnFila, 2) = rs!cSectorDesc
            feSector.TextMatrix(lnFila, 3) = Format(rs!nLimite, "#0.00")
            rs.MoveNext
        Loop
        feSector.TopRow = 1
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub CmdEditar_Click()
fsSectorCod = feSector.TextMatrix(feSector.row, 1)
lblSectorDesc.Caption = feSector.TextMatrix(feSector.row, 2)
txtLimite.Text = feSector.TextMatrix(feSector.row, 3)
HabilitaControles True
txtLimite.SetFocus
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub HabilitaControles(ByVal pbHabilita As Boolean)
gbControles.Visible = pbHabilita
cmdGrabar.Visible = pbHabilita
cmdCancelar.Visible = pbHabilita
cmdEditar.Visible = Not pbHabilita
End Sub

Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
    'permite que solo sean ingresados los numeros, el ENTER y el RETROCESO
    If InStr("0123456789.", Chr(KeyAscii)) = 0 Then
        SoloNumeros = 0
    Else
        SoloNumeros = KeyAscii
    End If
    ' teclas especiales permitidas
    If KeyAscii = 8 Then SoloNumeros = KeyAscii ' borrado atras
    If KeyAscii = 13 Then SoloNumeros = KeyAscii 'Enter
End Function

Private Sub txtLimite_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
    txtLimite.Text = Format(txtLimite.Text, "#0.00")
End If
End Sub

Private Sub txtLimite_KeyUp(KeyCode As Integer, Shift As Integer)
Dim nCantDecimales As Integer
    If val(txtLimite.Text) <> 0 Then
        If CInt(txtLimite.Text) > 100 Then
            MsgBox "El límite no debe ser superior a 100.00", vbInformation, "Aviso"
            txtLimite.Text = "100.00"
        End If
        If InStr(1, txtLimite.Text, ".") <> 0 Then
            nCantDecimales = Len(Mid(txtLimite.Text, InStr(1, txtLimite.Text, ".") + 1, Len(txtLimite.Text)))
            If nCantDecimales > 2 Then
                MsgBox "El límite solo permite 2 decimales", vbInformation, "Aviso"
                txtLimite.Text = Mid(txtLimite.Text, 1, Len(txtLimite.Text) - IIf(nCantDecimales > 3, 2, 1))
            End If
        End If
    Else
        txtLimite.Text = "0.00"
    End If
End Sub
