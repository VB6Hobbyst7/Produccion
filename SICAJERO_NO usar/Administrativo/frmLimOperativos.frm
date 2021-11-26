VERSION 5.00
Begin VB.Form frmMantNroOpeLibres 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operaciones Libres"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   Icon            =   "frmLimOperativos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Nro. Limite Retiros"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   2775
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3960
         TabIndex        =   1
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nro. Retiros:"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMantNroOpeLibres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub CmdGrabar_Click()
    Call guardarDatos
    MsgBox "Se actualizo con exito", vbInformation, "Numero de Operaciones Libres"
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    Call CargaDatos
End Sub

Private Sub guardarDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset

    Set R = New ADODB.Recordset
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva 'AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nAccion", adInteger, adParamInput, 8, 21)
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

Private Sub CargaDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset

    Set R = New ADODB.Recordset
        
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nAccion", adInteger, adParamInput, , 2)
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

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub
