VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmConfOperaConcilia 
   Caption         =   "Conciliar Operaciones"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5325
   Icon            =   "frmConfOperaConcilia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   15
      TabIndex        =   1
      Top             =   1395
      Width           =   5160
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   345
         Left            =   3660
         TabIndex        =   7
         Top             =   225
         Width           =   1005
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "Procesar"
         Height          =   345
         Left            =   495
         TabIndex        =   6
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   5130
      Begin MSMask.MaskEdBox lblFechaLog 
         Height          =   300
         Left            =   1440
         TabIndex        =   9
         Top             =   915
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha LOG:"
         Height          =   165
         Left            =   420
         TabIndex        =   8
         Top             =   930
         Width           =   915
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3360
         TabIndex        =   5
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label lblUsuario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1065
         TabIndex        =   4
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha:"
         Height          =   165
         Left            =   2730
         TabIndex        =   3
         Top             =   300
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario:"
         Height          =   165
         Left            =   390
         TabIndex        =   2
         Top             =   285
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmConfOperaConcilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdProcesar_Click()
    Dim Cmd As Command
    Dim Prm As Parameter
    Dim R As Recordset
    Dim sSQL As String
    
    Set Cmd = New Command
    Set Prm = New ADODB.Parameter
    
    Set Prm = New ADODB.Parameter
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    
    If lblFechaLog.Text = "__/__/____" Then
        MsgBox "Debe ingresar la Fecha del Archivo de Conciliacion", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If verificaConciliacion = 1 Then
        If MsgBox("El procesos de Confirmacion ya fue realizado, Desea realizarlo otra vez?", vbInformation + vbYesNo, "Proceso Confirmacion") = vbNo Then
            Exit Sub
        End If
    End If
    
    sSQL = "ConciliaOFFHostDia '" & Me.lblUsuario.Caption & "','" & Format(CDate(Me.lblFecha), "YYYY-MM-DD") & "','" & Format(CDate(Me.lblFechaLog), "YYYY-MM-DD") & "'"
    
    oConec.AbreConexion
    'C.Execute sSQL
    oConec.Ejecutar sSQL
    
    Call registraConciliacion
    'select
    oConec.CierraConexion
    MsgBox "Proceso de Confirmacion concluido con exito."
End Sub

Private Function verificaConciliacion() As Integer
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    
    Set Cmd = New ADODB.Command
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFechaSistema", adDate, adParamInput, 10, CDate(lblFecha.Caption))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nRpta", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_VerificaConciliacion"
    Cmd.Execute
    
    verificaConciliacion = Cmd.Parameters(1).Value
    
    oConec.CierraConexion
    Set Cmd = Nothing
    Set Prm = Nothing
    
End Function

Private Sub registraConciliacion()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim sResp As String
    
    Set Cmd = New ADODB.Command
        
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cUsuario", adVarChar, adParamInput, 50, gsCodUser)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFechaProceso", adDate, adParamInput, 10, CDate(lblFechaLog.Text))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFechaSistema", adDate, adParamInput, 10, CDate(lblFecha.Caption))
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RegistraConciliacion"
    Cmd.Execute
    
                    
    Set Cmd = Nothing
    Set Prm = Nothing
    
    'CerrarConexion
    oConec.CierraConexion
    
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    Call CargaFecha
    Call CargarUsuario
    'Call CargaFechaLog
    
    
End Sub

Private Sub CargarUsuario()
    lblUsuario.Caption = gsCodUser
End Sub

Private Sub CargaFechaLog()

    Dim Cmd As Command
    Dim Prm As Parameter
    Dim R As Recordset
    
    Set Cmd = New Command
    Set Prm = New ADODB.Parameter
    
    Set Prm = New ADODB.Parameter
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_CargaFechaLog"
    
    Set R = Cmd.Execute
    'Me.lblFechaLog.Caption = Format(CDate(Format(R!FechaLog, "dd/mm/yyyy")), "dd/mm/yyyy")
    R.Close
    Set Cmd = Nothing
    
    Set R = Nothing
    
    'CerrarConexion
    oConec.CierraConexion
End Sub


Private Sub CargaFecha()

    Dim Cmd As Command
    Dim Prm As Parameter
    Dim R As Recordset
    
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

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub
