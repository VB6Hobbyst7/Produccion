VERSION 5.00
Begin VB.Form frmMantTipoCuenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Tipo de Cuenta"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   Icon            =   "frmMantTipoCuenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   6135
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipos de Cuenta"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.ListBox LstTipoCta 
         Height          =   2760
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   240
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmMantTipoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub CargaDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R  As ADODB.Recordset

    Set R = New ADODB.Recordset
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Cmd.CommandText = "ATM_RecuperaTiposCuenta"
    
    R.CursorType = adOpenStatic
    R.LockType = adLockReadOnly
    Set R = Cmd.Execute
    LstTipoCta.Clear
    Do While Not R.EOF
        LstTipoCta.AddItem Left(R!nConsValor & Space(5), 5) & Left(R!cConsDescripcion & Space(25), 25)
        If R!nTipoCta = 1 Then
            LstTipoCta.Selected(LstTipoCta.ListCount - 1) = True
        Else
            LstTipoCta.Selected(LstTipoCta.ListCount - 1) = False
        End If
        R.MoveNext
    Loop
    
    R.Close
    Set R = Nothing
    oConec.CierraConexion
    
End Sub

Private Sub CmdGrabar_Click()
Dim i As Integer

    For i = 0 To LstTipoCta.ListCount - 1
        Dim Cmd As New Command
        Dim Prm As New ADODB.Parameter
           
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@pnTipoCta", adInteger, adParamInput, , CInt(Mid(Me.LstTipoCta.List(i), 1, 5)))
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nValor", adInteger, adParamInput, , IIf(LstTipoCta.Selected(i), 1, 0))
        Cmd.Parameters.Append Prm
            
        oConec.AbreConexion
        Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
        Cmd.CommandType = adCmdStoredProc
        Cmd.CommandText = "ATM_RegistraTipoCuenta"
        
        
        Call Cmd.Execute
        Set Prm = Nothing
        Set Cmd = Nothing
        
          
        'CerrarConexion
        oConec.CierraConexion
        
    Next i
    
    MsgBox "Datos Grabados"

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    CargaDatos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub
