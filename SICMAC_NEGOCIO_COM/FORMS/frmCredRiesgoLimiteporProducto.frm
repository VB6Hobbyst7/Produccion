VERSION 5.00
Begin VB.Form frmCredRiesgoLimiteporProducto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuracion de Limites por Producto"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frmCredRiesgoLimiteporProducto.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Frame gbControles 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtLimite 
         Height          =   315
         Left            =   3960
         MaxLength       =   6
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblTpCreditoDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   360
      Left            =   2760
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   3960
      TabIndex        =   1
      Top             =   4560
      Width           =   1050
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1050
   End
   Begin SICMACT.FlexEdit feTpProducto 
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   7011
      Cols0           =   4
      HighLight       =   1
      EncabezadosNombres=   "-nCodTpCredito-Tipo de Producto-Limite (%)"
      EncabezadosAnchos=   "300-0-3500-900"
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
      EncabezadosAlineacion=   "C-R-L-R"
      FormatosEdit    =   "0-3-1-2"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   3
      ColWidth0       =   300
      RowHeight0      =   300
   End
End
Attribute VB_Name = "frmCredRiesgoLimiteporProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oDPerGen As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Dim fsSectorCod As String

Private Sub cmdCancelar_Click()
fsSectorCod = ""
lblTpCreditoDesc.Caption = ""
txtLimite.Text = ""
CargarDatos
HabilitaControles False
End Sub

Private Sub cmdGrabar_Click()
Dim nCant As Integer

nCant = Len(txtLimite.Text)
 
If (txtLimite.Text) > 100 Then
    MsgBox "Favor de Verificar el limite de Crédito", vbInformation, "Aviso"
Else
    If IsNumeric(txtLimite.Text) Then
        If Trim(txtLimite.Text) = "" Then
            MsgBox "Debe ingresar el limite de crédito", vbInformation, "Aviso"
            Exit Sub
        End If
        Set oDPerGen = New COMDCredito.DCOMCredito
        Call oDPerGen.UpdateLimiteTpCredito(fsSectorCod, CDbl(txtLimite.Text))
        Set oDPerGen = Nothing
        MsgBox "Límite actualizado", vbInformation, "Aviso"
        cmdCancelar_Click
    Else
        MsgBox "Favor de Verificar el limite de Crédito", vbInformation, "Aviso"
        txtLimite.Text = ""
        txtLimite.SetFocus
    End If
End If
End Sub

Public Sub Inicia(ByVal pnTipo As Integer)
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

Set oDPerGen = New COMDCredito.DCOMCredito
Set rs = oDPerGen.CargaDatosLimiteTpCredito

Set oDPerGen = Nothing
    Call LimpiaFlex(feTpProducto)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feTpProducto.AdicionaFila
            lnFila = feTpProducto.row
            feTpProducto.TextMatrix(lnFila, 1) = rs!nTpCredito
            feTpProducto.TextMatrix(lnFila, 2) = rs!cTpCredito
            feTpProducto.TextMatrix(lnFila, 3) = Format(rs!nLimite, "#0.00")
            rs.MoveNext
        Loop
        feTpProducto.TopRow = 1
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub CmdEditar_Click()
fsSectorCod = feTpProducto.TextMatrix(feTpProducto.row, 1)
lblTpCreditoDesc.Caption = feTpProducto.TextMatrix(feTpProducto.row, 2)
txtLimite.Text = feTpProducto.TextMatrix(feTpProducto.row, 3)
HabilitaControles True
txtLimite.SetFocus
End Sub

Private Sub cmdSalir_Click()
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

Private Sub Form_Load()
    Call CentraForm(Me)
End Sub

Private Sub txtLimite_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
    txtLimite.Text = Format(txtLimite.Text, "#0.00")
End If
End Sub

Private Sub txtLimite_KeyUp(KeyCode As Integer, Shift As Integer)
Dim nCantDecimales As Integer

If IsNumeric(txtLimite.Text) Then
    If (txtLimite.Text <> "") Then
        If CCur(txtLimite.Text) > 100 Then
            MsgBox "El límite no debe ser superior a 100.00", vbInformation, "Aviso"
            txtLimite.Text = "0.00"
        End If
        If InStr(1, txtLimite.Text, ".") <> 0 Then
            nCantDecimales = Len(Mid(txtLimite.Text, InStr(1, txtLimite.Text, ".") + 1, Len(txtLimite.Text)))
            If nCantDecimales > 2 Then
                MsgBox "El límite solo permite 2 decimales", vbInformation, "Aviso"
                txtLimite.Text = Mid(txtLimite.Text, 1, Len(txtLimite.Text) - IIf(nCantDecimales > 3, 2, 1))
            End If
        End If
    Else
       txtLimite.Text = ""
    End If
Else
 txtLimite.Text = ""
End If

End Sub

