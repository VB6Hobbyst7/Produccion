VERSION 5.00
Begin VB.Form frmRepTarjEmiPorAge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Tarjetas Emitidas por Agencia"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmRepTarjEmiPorAge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   720
      Left            =   45
      TabIndex        =   3
      Top             =   1020
      Width           =   5145
      Begin VB.CommandButton CmdGenerar 
         Caption         =   "Generar"
         Height          =   450
         Left            =   105
         TabIndex        =   5
         Top             =   180
         Width           =   1410
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   450
         Left            =   3630
         TabIndex        =   4
         Top             =   165
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1020
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CheckBox ChkTodas 
         Caption         =   "Todas las Agencias"
         Height          =   285
         Left            =   1140
         TabIndex        =   6
         Top             =   615
         Width           =   3855
      End
      Begin VB.ComboBox CboAgencia 
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   3930
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia :"
         Height          =   330
         Left            =   135
         TabIndex        =   2
         Top             =   255
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmRepTarjEmiPorAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdGenerar_Click()
    
    If Me.CboAgencia.ListIndex = -1 And Me.ChkTodas.Value = 0 Then
        MsgBox "Selecciones una Agencia", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Call ReporteTarjetasEmitidaPorAgencia(IIf(Me.ChkTodas.Value = 1, 999, Mid(Me.CboAgencia.Text, 1, 2)), IIf(Me.ChkTodas.Value = 1, "TODAS", Me.CboAgencia.Text))
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
Dim R As ADODB.Recordset
Dim sSQL As String
Dim loConec As New DConecta

    'sSql = "Select cAgeCod, cAgeDescripcion from agosto05..Agencias Order by cAgeCod"
    sSQL = "ATM_DevuelveAgencias "
    
    Me.CboAgencia.Clear
    'AbrirConexion
    loConec.AbreConexion
    Set R = New ADODB.Recordset
    R.Open sSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not R.EOF
        CboAgencia.AddItem (R!cAgeCod & Space(1) & R!cAgeDescripcion)
        
        R.MoveNext
    Loop
    R.Close
    'CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    Set R = Nothing
End Sub

