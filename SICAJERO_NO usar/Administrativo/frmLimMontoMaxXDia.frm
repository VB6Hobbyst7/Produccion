VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLimMontoMaxXDia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Limite - Monto Maximo de Retiro Por Dia"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10320
   Icon            =   "frmLimMontoMaxXDia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   2550
      Width           =   10275
      Begin VB.TextBox txtHoraF 
         Height          =   285
         Left            =   4455
         MaxLength       =   2
         TabIndex        =   5
         Top             =   495
         Width           =   735
      End
      Begin VB.TextBox txtHoraI 
         Height          =   285
         Left            =   2685
         MaxLength       =   2
         TabIndex        =   4
         Top             =   495
         Width           =   735
      End
      Begin VB.TextBox txtMonto 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   90
         TabIndex        =   2
         Top             =   930
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   8910
         TabIndex        =   1
         Top             =   975
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Hra. Final:"
         Height          =   255
         Left            =   3660
         TabIndex        =   8
         Top             =   510
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Hra. Inic.:"
         Height          =   255
         Left            =   1950
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView LstLimites 
      Height          =   2535
      Left            =   15
      TabIndex        =   9
      Top             =   0
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripcion"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Hora Ini"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Hora Fin."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Moneda"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmLimMontoMaxXDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub CargaDatosListView()
Dim R As ADODB.Recordset
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim L As ListItem
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaLimMontosRetirosAcum"
    
    
    LstLimites.ListItems.Clear
    Set R = Cmd.Execute
        Do While Not R.EOF
         
         Set L = LstLimites.ListItems.Add(, , R!cItem)
         Call L.ListSubItems.Add(, , R!nDescLim)
         Call L.ListSubItems.Add(, , Format(R!nValor, "#,0.00"))
         Call L.ListSubItems.Add(, , R!nHoraI)
         Call L.ListSubItems.Add(, , R!nHoraF)
         Call L.ListSubItems.Add(, , R!nMoneda)
         
         R.MoveNext
    Loop
    R.Close
    
    'CerrarConexion
    oConec.CierraConexion
    Set Cmd = Nothing
    Set Prm = Nothing

End Sub

Private Sub CmdGrabar_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter

    If Not IsNumeric(txtMonto.Text) Or Not IsNumeric(txtHoraI.Text) Or Not IsNumeric(txtHoraF.Text) Then
        MsgBox "Uno de los datos es Incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnItem", adInteger, adParamInput, , CInt(LstLimites.SelectedItem.Text))
    Cmd.Parameters.Append Prm
    
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnValor", adDouble, adParamInput, , CDbl(txtMonto.Text))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnHoraIni", adInteger, adParamInput, , txtHoraI.Text)
    Cmd.Parameters.Append Prm
    
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnHoraFin", adInteger, adParamInput, , txtHoraF.Text)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RegistraLimMontosRetirosAcum"
    
    Cmd.Execute
    
    oConec.CierraConexion
    
    Call CargaDatosListView
    
    MsgBox "Datos Actualizados", vbInformation, "Aviso"
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Set oConec = New DConecta
    CargaDatosListView
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub LstLimites_Click()
    txtMonto.Text = Format(Me.LstLimites.SelectedItem.ListSubItems(2).Text, "#,0.00")
    txtHoraI.Text = Me.LstLimites.SelectedItem.ListSubItems(3).Text
    txtHoraF.Text = Me.LstLimites.SelectedItem.ListSubItems(4).Text
End Sub
