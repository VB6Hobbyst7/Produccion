VERSION 5.00
Begin VB.Form frmStockDevTarjeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolucion de Tarjeta"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   Icon            =   "frmStockDevTarjeta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   3345
      Width           =   5895
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   120
         Width           =   1185
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   120
         Width           =   1185
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5925
      Begin VB.Frame Frame4 
         Caption         =   "Glosa"
         Height          =   2190
         Left            =   75
         TabIndex        =   4
         Top             =   1050
         Width           =   5760
         Begin VB.TextBox txtGlosa 
            Height          =   1890
            Left            =   120
            TabIndex        =   5
            Top             =   225
            Width           =   5565
         End
      End
      Begin VB.Frame Frame3 
         Height          =   780
         Left            =   75
         TabIndex        =   1
         Top             =   195
         Width           =   5745
         Begin VB.TextBox txtCantidad 
            Height          =   315
            Left            =   930
            TabIndex        =   2
            Top             =   285
            Width           =   1425
         End
         Begin VB.Label Label3 
            Caption         =   "Cantidad :"
            Height          =   255
            Left            =   135
            TabIndex        =   3
            Top             =   315
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "frmStockDevTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdGrabar_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim sResp As String
Dim loConec As New DConecta

    Set Cmd = New ADODB.Command
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCboUsuDes", adVarChar, adParamInput, 100, gsCodUser)
    Cmd.Parameters.Append Prm
       
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCantidad", adBigInt, adParamInput, 8, CInt(txtCantidad.Text))
    Cmd.Parameters.Append Prm
        
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cObservacion", adVarChar, adParamInput, 100, txtGlosa.Text)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCodAge", adChar, adParamInput, 4, gsCodAge)
    Cmd.Parameters.Append Prm
       
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RegistraDevTarjeta"
    Cmd.Execute
    
    'CerrarConexion
    loConec.CierraConexion
    MsgBox "Devolución Realizada Correctamente", vbInformation, "Aviso"
    
    Set loConec = Nothing
    Set Cmd = Nothing
    Set Prm = Nothing
    
    Frame1.Enabled = False
    Me.CmdGrabar.Enabled = False
    
End Sub

Private Sub cmdNuevo_Click()
    txtCantidad.Text = "0"
    txtGlosa.Text = " "
    Frame1.Enabled = True
    CmdGrabar.Enabled = True
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
CargaDatos
End Sub
Private Sub CargaDatos()
 
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim R As ADODB.Recordset
    Set Prm = New ADODB.Parameter
    
        
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@pnCodAge", adInteger, adParamInput, , CInt(gsCodAge))
'    Cmd.Parameters.Append Prm
'
'    Cmd.ActiveConnection = AbrirConexion
'    Cmd.CommandType = adCmdStoredProc
'    Cmd.CommandText = "ATM_RecuperaUsuarios"
'
'    Set R = Cmd.Execute
'    'CboCtas.Clear
'    Do While Not R.EOF
'         CboUsuDes.AddItem R!Codigo & "-" & R!Nombre
'       R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48) Or (KeyAscii > 57) Then
        KeyAscii = 0
        txtCantidad.SetFocus
    End If
End Sub


