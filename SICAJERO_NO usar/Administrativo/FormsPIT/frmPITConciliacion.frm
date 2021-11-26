VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPITConciliacion 
   Caption         =   "Conciliación de Operaciones InterCajas: Confrimación de Operaciones"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConciliar 
      Caption         =   "&Procesar"
      Height          =   345
      Left            =   3960
      TabIndex        =   8
      Top             =   1680
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5160
      TabIndex        =   7
      Top             =   1680
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   5970
      Begin MSMask.MaskEdBox lblFechaLog 
         Height          =   300
         Left            =   4080
         TabIndex        =   6
         Top             =   840
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
         Left            =   2880
         TabIndex        =   5
         Top             =   840
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
         Left            =   4080
         TabIndex        =   4
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
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha:"
         Height          =   165
         Left            =   3240
         TabIndex        =   2
         Top             =   300
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario:"
         Height          =   165
         Left            =   390
         TabIndex        =   1
         Top             =   285
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmPITConciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta
Dim nConciliacionId As Long

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdConciliar_Click()
Dim lRs As ADODB.Recordset
Dim loConec As DConecta
Dim loOIC As dPITFunciones
    
    
    If lblFechaLog.Text = "__/__/____" Then
        MsgBox "Debe ingresar la Fecha del Archivo de Conciliacion", vbInformation, "Aviso"
        Exit Sub
    End If
    
    
    Set loConec = New DConecta
    loConec.AbreConexion
        
    Set loOIC = New dPITFunciones
    Set lRs = loOIC.recuperaConciliacionPorFechaProc(CDate(lblFecha.Caption))
    
    If Not (lRs.EOF And lRs.BOF) Then
        MsgBox "El proceso de Conciliación ya fue realizado, no es posible realizarlo nuevamente", vbInformation + vbCritical, "Aviso"
        Exit Sub
    End If
        
    nConciliacionId = loOIC.nRegistraConciliacion(gdFecSis, CDate(lblFechaLog.Text), gsCodUser)
          
    
    MsgBox "El Proceso de Conciliación de Operaciones Intercajas ha Concluido con Exito."
    
    loConec.CierraConexion
    Set loOIC = Nothing
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
    lblFecha.Caption = Format(gdFecSis, "DD/MM/YYYY")
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

