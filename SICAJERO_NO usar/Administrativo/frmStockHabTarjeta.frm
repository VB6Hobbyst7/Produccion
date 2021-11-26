VERSION 5.00
Begin VB.Form frmStockHabTarjeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Habilitacion de Tarjeta"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   Icon            =   "frmStockHabTarjeta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   5895
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1185
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         Top             =   120
         Width           =   1185
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   120
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5895
      Begin VB.Frame Frame5 
         Caption         =   "Glosa"
         Height          =   1215
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   5655
         Begin VB.TextBox txtGlosa 
            Height          =   900
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   5355
         End
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   5655
         Begin VB.TextBox txtCantidad 
            Height          =   315
            Left            =   960
            TabIndex        =   1
            Text            =   "20"
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label3 
            Caption         =   "Cantidad :"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5655
         Begin VB.ComboBox CboUsuDes 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   4035
         End
         Begin VB.Label Label2 
            Caption         =   "Usuario Destino :"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1230
         End
      End
   End
End
Attribute VB_Name = "frmStockHabTarjeta"
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

    Set Cmd = New ADODB.Command
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCboUsuDes", adVarChar, adParamInput, 100, Mid(CboUsuDes.Text, 1, InStr(CboUsuDes.Text, "-") - 1))
    Cmd.Parameters.Append Prm
       
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCantidad", adBigInt, adParamInput, , CInt((txtCantidad.Text)))
    Cmd.Parameters.Append Prm
        
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cObservacion", adVarChar, adParamInput, 100, txtGlosa.Text)
    Cmd.Parameters.Append Prm

    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RegistraHabTarjeta"
    
    Cmd.Execute
    
    Set Cmd = Nothing
    Set Prm = Nothing
    
    'CerrarConexion
    oConec.CierraConexion
    
    MsgBox "Operacion Registrada Con Exito", vbInformation, "Aviso"
    cmdGrabar.Enabled = False
    Frame1.Enabled = False
End Sub

Private Sub cmdNuevo_Click()
    txtCantidad.Text = "0"
    txtGlosa.Text = " "
    cmdGrabar.Enabled = True
    Frame1.Enabled = True
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    Call CargaDatos
End Sub
Private Sub CargaDatos()
 
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim R As ADODB.Recordset
    Set Prm = New ADODB.Parameter
    
      
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCodAge", adInteger, adParamInput, , CInt(gsCodAge))
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaUsuarios"
    
    Set R = Cmd.Execute
    'CboCtas.Clear
    Do While Not R.EOF
         CboUsuDes.AddItem R!Codigo & "-" & R!Nombre
       R.MoveNext
    Loop
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

