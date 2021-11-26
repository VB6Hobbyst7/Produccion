VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectCuenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Cuenta de Ahorros"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   Icon            =   "frmSelectCuenta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   30
      TabIndex        =   2
      Top             =   2205
      Width           =   5880
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   360
         Left            =   4485
         TabIndex        =   4
         Top             =   195
         Width           =   1320
      End
      Begin VB.CommandButton CmdSeleccionar 
         Caption         =   "Seleccionar"
         Height          =   360
         Left            =   90
         TabIndex        =   3
         Top             =   195
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   5835
      Begin MSComctlLib.ListView LstCuentas 
         Height          =   1815
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuenta"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Saldo"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSelectCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private psPerscod As String
Private psCtaSelec As String
Private nMoneda As Integer


Private Sub CmdSalir_Click()
psCtaSelec = ""
Unload Me

End Sub

Private Sub CargaDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim L As ListItem
Dim R As ADODB.Recordset
Dim loConec As New DConecta

        Set R = New ADODB.Recordset
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@psPersCod", adVarChar, adParamInput, 20, psPerscod)
        Cmd.Parameters.Append Prm
        
             
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@psMoneda", adVarChar, adParamInput, 2, Trim(Str(nMoneda)))
        Cmd.Parameters.Append Prm
        
        loConec.AbreConexion
        Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
        Cmd.CommandType = adCmdStoredProc
        
        Cmd.CommandText = "ATM_RecuperaCtasTitularPorPersona"

        R.CursorType = adOpenStatic
        R.LockType = adLockReadOnly
        Set R = Cmd.Execute
        Me.LstCuentas.ListItems.Clear
        
        Do While Not R.EOF
            Set L = Me.LstCuentas.ListItems.Add(, , R!cCtaCod)
            Call L.ListSubItems.Add(, , Format(R!nSaldo, "#,0.00"))
            
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Set Cmd = Nothing
        Set Prm = Nothing
        
        loConec.CierraConexion
        Set loConec = Nothing
        
End Sub

Public Function Seleccionar(pnMoneda As Integer, ByVal pPerscod As String) As String

    psPerscod = pPerscod
    nMoneda = pnMoneda
    Me.Show 1
    Seleccionar = psCtaSelec
    
End Function

Private Sub CmdSeleccionar_Click()
If Me.LstCuentas.ListItems.Count > 0 Then
    psCtaSelec = Me.LstCuentas.SelectedItem.Text
    
Else
    psCtaSelec = ""
End If
    Unload Me
    
End Sub

Private Sub Form_Load()
    Call CargaDatos
    
End Sub

