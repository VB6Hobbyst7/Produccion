VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMantMontoMaDeRetiro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Monto Maximo de Retiro"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
   Icon            =   "frmMantMaxRetiroxDiaS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   135
      TabIndex        =   7
      Top             =   2670
      Width           =   9540
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   8100
         TabIndex        =   13
         Top             =   945
         Width           =   1335
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   90
         TabIndex        =   6
         Top             =   930
         Width           =   1335
      End
      Begin VB.TextBox txtMonto 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtHoraI 
         Height          =   285
         Left            =   2685
         MaxLength       =   2
         TabIndex        =   2
         Top             =   495
         Width           =   735
      End
      Begin VB.TextBox txtMinutosI 
         Height          =   285
         Left            =   4620
         MaxLength       =   2
         TabIndex        =   3
         Top             =   510
         Width           =   735
      End
      Begin VB.TextBox txtHoraF 
         Height          =   285
         Left            =   6585
         MaxLength       =   2
         TabIndex        =   4
         Top             =   495
         Width           =   735
      End
      Begin VB.TextBox txtMinutosF 
         Height          =   285
         Left            =   8460
         MaxLength       =   2
         TabIndex        =   5
         Top             =   495
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label24 
         Caption         =   "Min. Inic.:"
         Height          =   255
         Left            =   3900
         TabIndex        =   11
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Hra. Inic.:"
         Height          =   255
         Left            =   1950
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Hra. Final:"
         Height          =   255
         Left            =   5790
         TabIndex        =   9
         Top             =   510
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Min. Final:"
         Height          =   255
         Left            =   7635
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView LstLimites 
      Height          =   2535
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   9615
      _ExtentX        =   16960
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Hora Ini"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Min Inic."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Hora Fin."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Min Final"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Moneda"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmMantMontoMaDeRetiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub cmdCancelar_Click()
    Unload Me
End Sub



Private Sub CmdGrabar_Click()
Dim R As ADODB.Recordset
Dim Cmd As New Command
Dim Prm As ADODB.Parameter

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("cItem", adInteger, adParamInput, 8, LstLimites.SelectedItem.Text)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nMonto", adCurrency, adParamInput, , txtMonto.Text)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("nHoraI", adInteger, adParamInput, 8, txtHoraI.Text)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("nMinutosI", adInteger, adParamInput, 8, txtMinutosI.Text)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("nHoraF", adInteger, adParamInput, 8, txtHoraF.Text)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("nMinutosF", adInteger, adParamInput, 8, txtMinutosF.Text)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ActualizarMontoMaximoRet"
    
    Set R = Cmd.Execute
    
    oConec.CierraConexion
    Call CargaDatosListView
    
    MsgBox "Datos Actualizados", vbInformation, "Aviso"
        
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    Call CargaDatosListView
End Sub

Private Sub CargaDatosListView()
Dim R As ADODB.Recordset
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim L As ListItem

    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaMontosRetiros"
    
    
    LstLimites.ListItems.Clear
    Set R = Cmd.Execute
        Do While Not R.EOF
         'lstAgeStock.AddItem Right("000" & Trim(Str(R!nCodAgeArea)), 3) & Space(3) & Left(R!cNomAgeArea & Space(25), 25) & Right(Space(5) & Trim(Str(R!nCantidad)), 5) & Right(Space(5) & Trim(Str(R!nMinimo)), 5)
         Set L = LstLimites.ListItems.Add(, , R!cItem)
         Call L.ListSubItems.Add(, , Format(R!nValor, "#,0.00"))
         Call L.ListSubItems.Add(, , R!nHoraI)
         Call L.ListSubItems.Add(, , R!nMinutosI)
         Call L.ListSubItems.Add(, , R!nHoraF)
         Call L.ListSubItems.Add(, , R!nMinutosF)
         Call L.ListSubItems.Add(, , R!nMoneda)
         'Call L.ListSubItems.Add(, , R!nMoneda)
         
         R.MoveNext
    Loop
    R.Close
    
    'CerrarConexion
    oConec.CierraConexion
    Set Cmd = Nothing
    Set Prm = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub LstLimites_Click()

    txtMonto.Text = Format(Me.LstLimites.SelectedItem.ListSubItems(1).Text, "#,0.00")
    txtHoraI.Text = Me.LstLimites.SelectedItem.ListSubItems(2).Text
    txtMinutosI.Text = Me.LstLimites.SelectedItem.ListSubItems(3).Text
    txtHoraF.Text = Me.LstLimites.SelectedItem.ListSubItems(4).Text
    txtMinutosF.Text = Me.LstLimites.SelectedItem.ListSubItems(5).Text
    
End Sub
