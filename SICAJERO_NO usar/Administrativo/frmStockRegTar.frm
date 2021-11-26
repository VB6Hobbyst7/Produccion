VERSION 5.00
Begin VB.Form frmStockRegTar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Stock Actual de Usuario"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7845
   Icon            =   "frmStockRegTar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   120
      TabIndex        =   6
      Top             =   -15
      Width           =   7695
      Begin VB.Label lblUsuario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3075
         TabIndex        =   10
         Top             =   270
         Width           =   4320
      End
      Begin VB.Label Label3 
         Height          =   300
         Left            =   6765
         TabIndex        =   11
         Top             =   345
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario :"
         Height          =   285
         Left            =   2325
         TabIndex        =   9
         Top             =   285
         Width           =   765
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/01/2008"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   735
         TabIndex        =   8
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha :"
         Height          =   285
         Left            =   60
         TabIndex        =   7
         Top             =   285
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   120
      TabIndex        =   4
      Top             =   1530
      Width           =   7695
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   435
         Left            =   6255
         TabIndex        =   5
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stock Actual"
      Height          =   810
      Left            =   105
      TabIndex        =   0
      Top             =   720
      Width           =   7710
      Begin VB.TextBox txtCantidad 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   930
         TabIndex        =   2
         Text            =   "0"
         Top             =   225
         Width           =   1425
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   420
         Left            =   6180
         TabIndex        =   1
         Top             =   210
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   135
         TabIndex        =   3
         Top             =   255
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmStockRegTar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub CmdGrabar_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim sResp As String

    If Not IsNumeric(txtCantidad.Text) Then
        MsgBox "Cantidad Invalida", vbInformation, "Aviso"
        Exit Sub
    End If

    If CInt(txtCantidad.Text) = 0 Then
        MsgBox "Cantidad debe ser Mayor a Cero", vbInformation, "Aviso"
        Exit Sub
    End If

    Set Cmd = New ADODB.Command
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCboUsu", adVarChar, adParamInput, 100, Mid(lblUsuario.Caption, 1, 4))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCantidad", adInteger, adParamInput, , CInt(txtCantidad.Text))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFecha", adDBDate, adParamInput, , CDate(lblFecha.Caption))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCantidadRpta", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@bEstado", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCodAge", adChar, adParamInput, 4, gsCodAge)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RegistrarStockActual"
    Cmd.Execute

    MsgBox "Datos Registrados Correctamente", vbInformation, "Aviso"
    Call CargaDatos
    txtCantidad = Label3.Caption
    
    'CerrarConexion
    oConec.CierraConexion
    
    Set Cmd = Nothing
    Set Prm = Nothing

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    
    lblUsuario.Caption = gsCodUser & Space(1) & gsNomUser
    Call CargarFechaSistema
    Call VerificaDatos
    If CInt(Label3.Caption) > 0 Then
        MsgBox "Ya existe un registro para esta fecha", vbInformation, "Stock Boveda General"
    End If
    txtCantidad = Label3.Caption
End Sub

Private Sub VerificaDatos()
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim sResp As String

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFecha", adDBDate, adParamInput, , CDate(lblFecha.Caption))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCodUsu", adVarChar, adParamInput, 100, Mid(Me.lblUsuario.Caption, 1, 4))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCantidadRpta", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_VerificarRegActual"
    Cmd.Execute
    Label3.Caption = Cmd.Parameters(2).Value
    
    Set Cmd = Nothing
    'CerrarConexion
    oConec.CierraConexion
    
End Sub


Private Sub CargaFechaSistema()
 
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim R As ADODB.Recordset
    Set Prm = New ADODB.Parameter
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaFechaSistema"
    
    Set R = Cmd.Execute
    lblFecha.Caption = Format(CDate(Format(R!FechaSistema, "dd/mm/yyyy")), "dd/mm/yyyy")
    R.Close
    Set R = Nothing
    oConec.CierraConexion
End Sub


Private Sub CargaDatos()
 
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim sResp As String

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFecha", adDBDate, adParamInput, , CDate(lblFecha.Caption))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCboUsu", adVarChar, adParamInput, 10, Mid(lblUsuario.Caption, 1, 4))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCantidadRpta", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_VerificarRegActual"
    Cmd.Execute
    Label3.Caption = Cmd.Parameters(2)
    
    'If Cmd.Parameters(2).Value > 0 Then
     '   MsgBox "Ya existe un registro para esta fecha", vbInformation, "Stock Boveda General"
    'End If
    'txtCantidad = Cmd.Parameters(2).Value
    oConec.CierraConexion
        
End Sub

Private Sub CargarFechaSistema()
 
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim R As ADODB.Recordset
    Set Prm = New ADODB.Parameter
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaFechaSistema"
    
    Set R = Cmd.Execute
    lblFecha.Caption = Format(CDate(Format(R!FechaSistema, "dd/mm/yyyy")), "dd/mm/yyyy")
    R.Close
    Set R = Nothing
    oConec.CierraConexion
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48) Or (KeyAscii > 57) Then
        KeyAscii = 0
        txtCantidad.SetFocus
    End If
End Sub

