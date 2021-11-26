VERSION 5.00
Begin VB.Form frmMantNroRetXDia 
   Caption         =   "Nro.Retiros por Dia"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5550
   Icon            =   "frmMantLimiteRetiros.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Nro. Limite Retiros"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nro. Retiros:"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMantNroRetXDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub CargaDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
    
    Set R = New ADODB.Recordset
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nAccion", adInteger, adParamInput, 8, 1)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@Valor", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@ValorS", adCurrency, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@ValorD", adCurrency, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    
    Cmd.CommandText = "ATM_RecuperaLimitesOperativos"
    
    R.CursorType = adOpenStatic
    R.LockType = adLockReadOnly
    Set R = Cmd.Execute
    txtCantidad = Cmd.Parameters(1).Value
    
    oConec.CierraConexion
    
End Sub


Private Sub CmdGrabar_Click()
    Call guardarDatos
    MsgBox "Se actualizo con exito", vbInformation, "Numero de Retiros por Dia"
End Sub

Private Sub guardarDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset

    Set R = New ADODB.Recordset
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nAccion", adInteger, adParamInput, 8, 31)
    Cmd.Parameters.Append Prm
    
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@MontoRDS", adCurrency, adParamInput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@MontoRDD", adCurrency, adParamInput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@Valor", adCurrency, adParamInput, , txtCantidad.Text)
    Cmd.Parameters.Append Prm
    
                   
    Cmd.CommandText = "ATM_ActualizarMontosLimOperativos"
    
    R.CursorType = adOpenStatic
    R.LockType = adLockReadOnly
    Set R = Cmd.Execute
    
    oConec.CierraConexion
    
    Call CargaDatos
        
End Sub


Private Sub Form_Load()
    Set oConec = New DConecta
    Call CargaDatos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub
