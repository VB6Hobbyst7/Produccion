VERSION 5.00
Begin VB.Form frmMantTarComRep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tarifario de Comisio por Reposicion"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4125
   Icon            =   "frmMantTarComRep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   60
      TabIndex        =   1
      Top             =   1155
      Width           =   4020
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   495
         Left            =   2775
         TabIndex        =   7
         Top             =   210
         Width           =   1155
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   495
         Left            =   90
         TabIndex        =   6
         Top             =   195
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   4005
      Begin VB.TextBox txtMonDol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         Height          =   300
         Left            =   2385
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   645
         Width           =   1215
      End
      Begin VB.TextBox txtMonSol 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2385
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "De Cuenta en Dolares Cobrar :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   690
         Width           =   2220
      End
      Begin VB.Label Label1 
         Caption         =   "De Cuenta en Soles  Cobrar :"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   2220
      End
   End
End
Attribute VB_Name = "frmMantTarComRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub CmdGrabar_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
            
    If Not IsNumeric(Me.txtMonDol.Text) Or Not IsNumeric(Me.txtMonSol.Text) Then
        MsgBox "Uno de los Montos es Incorrecto", vbInformation, "Aviso"
        Exit Sub
        
    End If
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
           
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMontoSol", adDouble, adParamInput, , CDbl(Me.txtMonSol.Text))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMontoDol", adDouble, adParamInput, , CDbl(Me.txtMonDol.Text))
    Cmd.Parameters.Append Prm
    
    
    Cmd.CommandText = "ATM_RegistraTarifComRep"
    
    
    Cmd.Execute
    
    Set Prm = Nothing
    Set Cmd = Nothing
    
    'CerrarConexion
    oConec.CierraConexion
    
    MsgBox "Datos Grabados", vbInformation, "Aviso"
        
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As New ADODB.Recordset

    Set oConec = New DConecta
    Set R = New ADODB.Recordset
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
           
    Cmd.CommandText = "ATM_RecuperaTarifComRep"
    
    R.CursorType = adOpenStatic
    R.LockType = adLockReadOnly
    Set R = Cmd.Execute
    
    Do While Not R.EOF
        If R!nMoneda = 1 Then
            Me.txtMonSol.Text = Format(R!nValor, "#,0.00")
        Else
            Me.txtMonDol.Text = Format(R!nValor, "#,0.00")
        
        End If
        
        R.MoveNext
    Loop
    
    R.Close
    
    Set R = Nothing
    Set Cmd = Nothing
    Set Prm = Nothing
    'CerrarConexion
    oConec.CierraConexion
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub
