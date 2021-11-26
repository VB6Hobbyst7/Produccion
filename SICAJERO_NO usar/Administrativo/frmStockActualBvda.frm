VERSION 5.00
Begin VB.Form frmStockActualBvda 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registros Actual de Boveda"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5355
   Icon            =   "frmStockActualBvda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   30
      TabIndex        =   9
      Top             =   1680
      Width           =   5205
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   435
         Left            =   3870
         TabIndex        =   10
         Top             =   165
         Width           =   1140
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stock Actual"
      Height          =   810
      Left            =   60
      TabIndex        =   5
      Top             =   840
      Width           =   5190
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   420
         Left            =   3870
         TabIndex        =   7
         Top             =   180
         Width           =   1125
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   930
         TabIndex        =   6
         Text            =   "0"
         Top             =   225
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   135
         TabIndex        =   8
         Top             =   255
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   105
      TabIndex        =   0
      Top             =   45
      Width           =   5160
      Begin VB.Label lab 
         Height          =   195
         Left            =   2355
         TabIndex        =   11
         Top             =   420
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblSaldoAnt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B O V E D A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3435
         TabIndex        =   4
         Top             =   210
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha :"
         Height          =   285
         Left            =   150
         TabIndex        =   3
         Top             =   225
         Width           =   600
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/01/2008"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   780
         TabIndex        =   2
         Top             =   210
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario :"
         Height          =   285
         Left            =   2685
         TabIndex        =   1
         Top             =   210
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmStockActualBvda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

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

Private Sub GrabarStockActualBvda()
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter

    Set Cmd = New ADODB.Command
    
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
    Cmd.CommandText = "ATM_RegistrarStockActualBvda"
    Cmd.Execute

    MsgBox "Datos Registrados Correctamente", vbInformation, "Aviso"
    
    'CerrarConexion
    oConec.CierraConexion
    'Set Cmd = Nothing
    'Set Prm = Nothing
    
End Sub

Private Sub CmdGrabar_Click()
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter

    Set Cmd = New ADODB.Command
    
    'Primero voy a verificar si hay valor grabado para la fecha
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFecha", adDBDate, adParamInput, , CDate(lblFecha.Caption))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCantidadRpta", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
   Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCodAge", adInteger, adParamInput, , CInt(gsCodAge))
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_VerificarRegActualBvda"
    Cmd.Execute
    

    
    If Cmd.Parameters(1).Value = 0 Then
        Call GrabarStockActualBvda
        Exit Sub
    End If
    If MsgBox("Existe(n) " & Cmd.Parameters(1).Value & " unidades para esta fecha" & ", Desea Continuar ?", vbInformation + vbYesNo, "Registro Actual de Boveda") = vbNo Then
        Exit Sub
    End If
    Call GrabarStockActualBvda
    
    oConec.CierraConexion
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    Call CargaFechaSistema
    Call VerificaDatos
    If CInt(lab.Caption) > 0 Then
        MsgBox "Ya existe un registro para esta fecha", vbInformation, "Stock Boveda General"
    End If
    txtCantidad = lab.Caption
End Sub

Private Sub VerificaDatos()
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim sResp As String

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFecha", adDBDate, adParamInput, , CDate(lblFecha.Caption))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCantidadRpta", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCodAge", adInteger, adParamInput, , CInt(gsCodAge))
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_VerificarRegActualBvda"
    Cmd.Execute
    lab.Caption = Cmd.Parameters(1).Value
    oConec.CierraConexion
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub
