VERSION 5.00
Begin VB.Form frmStockRegBovGen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro Stock Boveda General"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "frmStockRegBovGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   45
      TabIndex        =   6
      Top             =   15
      Width           =   5895
      Begin VB.Label Label4 
         Caption         =   "Fecha de Ultimo Registro :"
         Height          =   285
         Left            =   2190
         TabIndex        =   10
         Top             =   270
         Width           =   1920
      End
      Begin VB.Label LblFecUltReg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/01/2008"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4140
         TabIndex        =   9
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/01/2008"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   630
         TabIndex        =   8
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha :"
         Height          =   285
         Left            =   60
         TabIndex        =   7
         Top             =   255
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   45
      TabIndex        =   4
      Top             =   1575
      Width           =   5850
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   435
         Left            =   4245
         TabIndex        =   5
         Top             =   225
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stock Actual"
      Height          =   810
      Left            =   45
      TabIndex        =   0
      Top             =   750
      Width           =   5880
      Begin VB.TextBox txtCantidad 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   930
         TabIndex        =   2
         Text            =   "0"
         Top             =   285
         Width           =   1425
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   420
         Left            =   4290
         TabIndex        =   1
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   135
         TabIndex        =   3
         Top             =   315
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmStockRegBovGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub CmdGrabar_Click()
    Call grabarDatos
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Set oConec = New DConecta
    Call CargaDatos
    Call VerificaDatos
End Sub

Private Sub VerificaDatos()
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim sResp As String
    Dim sDate As String
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFecha", adDBDate, adParamInput, , CDate(lblFecha.Caption))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCantidadRpta", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFecUltReg", adDate, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nMovENELDia", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_VerificarRegActualBvdaGral"
    Cmd.Execute
    
    If Cmd.Parameters(3).Value > 0 Then
        MsgBox "No se puede Regsitrar porque existen Movimientos de Boveda General el dia de hoy", vbInformation, "Aviso"
        Set Cmd = Nothing
        Set Prm = Nothing
        Me.CmdGrabar.Enabled = False
        Me.txtCantidad.Enabled = False
        Exit Sub
    
    End If
    
    sDate = IIf(IsNull(Cmd.Parameters(1).Value), "", Format(Cmd.Parameters(2).Value, "dd/mm/yyyy"))
    
    LblFecUltReg.Caption = sDate
    
    If Cmd.Parameters(1).Value > 0 Then
        MsgBox "Ya existe un registro para esta fecha", vbInformation, "Stock Boveda General"
    End If
    txtCantidad = Cmd.Parameters(1).Value
    
    oConec.CierraConexion
    Set Cmd = Nothing
    Set Prm = Nothing
    
End Sub
Private Sub CargaDatos()
Dim Cmd As Command
Dim Prm As ADODB.Parameter
Dim R As ADODB.Recordset

    Set Cmd = New Command
    Set Prm = New ADODB.Parameter
    
    Set Prm = New ADODB.Parameter
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaFechaSistema"
    
    Set R = Cmd.Execute
    Me.lblFecha.Caption = Format(CDate(Format(R!FechaSistema, "dd/mm/yyyy")), "dd/mm/yyyy")
    
       
    R.Close
    Set Cmd = Nothing
    
    Set R = Nothing
    
    'CerrarConexion
    oConec.CierraConexion
    
End Sub

Private Sub grabarDatos()
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    
    If Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "Cantidad Incorrecta", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFecha", adDate, adParamInput, 16, lblFecha.Caption)
    Cmd.Parameters.Append Prm
        
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@Cantidad", adInteger, adParamInput, , txtCantidad.Text)
    Cmd.Parameters.Append Prm
     
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RegistrarStockBvdaGnral"
    Cmd.Execute
    
    
    MsgBox "Stock de Boveda Gral. registrado con Exito", vbInformation, "Stock Boveda General"
    
    'CerrarConexion
    oConec.CierraConexion
    Set Cmd = Nothing
    Set Prm = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub
