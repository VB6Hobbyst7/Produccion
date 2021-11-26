VERSION 5.00
Begin VB.Form frmMantMontoMaxRetXDia 
   Caption         =   "Monto Max. Retiro por Dia"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5565
   Icon            =   "frmMantMontoMaxRetXDia.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "Monto Max. Retiro por Dia"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.Frame Frame6 
         Caption         =   "Soles"
         Height          =   855
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5055
         Begin VB.TextBox txtMontoRDS 
            Height          =   285
            Left            =   1560
            TabIndex        =   7
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton cmdGrabarRDS 
            Caption         =   "Grabar"
            Height          =   375
            Left            =   3840
            TabIndex        =   6
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Monto:"
            Height          =   255
            Left            =   480
            TabIndex        =   8
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Dolares"
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   5055
         Begin VB.CommandButton cmdGrabarRDD 
            Caption         =   "Grabar"
            Height          =   375
            Left            =   3840
            TabIndex        =   3
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtMontoRDD 
            Height          =   285
            Left            =   1560
            TabIndex        =   2
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Monto:"
            Height          =   255
            Left            =   480
            TabIndex        =   4
            Top             =   360
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "frmMantMontoMaxRetXDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub CargaDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset

    Set R = New ADODB.Recordset
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nAccion", adInteger, adParamInput, 8, 3)
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
    txtMontoRDS = Cmd.Parameters(2).Value
    txtMontoRDD = Cmd.Parameters(3).Value
    
    oConec.CierraConexion
        
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGrabarRDD_Click()
    Call guardarDatosMontoMaxRetXDia(12, 0, txtMontoRDD.Text)
    MsgBox "Se actualizo con exito", vbInformation, "Monto Maximo De Retirod por Dia"
End Sub

Private Sub cmdGrabarRDS_Click()
    Call guardarDatosMontoMaxRetXDia(11, txtMontoRDS.Text, 0)
    MsgBox "Se actualizo con exito", vbInformation, "Monto Maximo De Retirod por Dia"
    
End Sub

Private Sub guardarDatosMontoMaxRetXDia(ByVal accion As Integer, ByVal monto1 As Currency, ByVal monto2 As Currency)
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset

    Set R = New ADODB.Recordset
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nAccion", adInteger, adParamInput, 8, accion)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@MontoRDS", adCurrency, adParamInput, , monto1)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@MontoRDD", adCurrency, adParamInput, , monto2)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@Valor", adCurrency, adParamInput, , 0)
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
