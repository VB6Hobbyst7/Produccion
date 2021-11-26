VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStockCtrlAgeActual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Actual de Agencias"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8940
   Icon            =   "frmStockCtrlAgeActual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Salir"
      Height          =   345
      Left            =   7470
      TabIndex        =   1
      Top             =   3480
      Width           =   1290
   End
   Begin MSComctlLib.ListView lstvAgencias 
      Height          =   3345
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   5900
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Agencia"
         Object.Width           =   4235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cantidad Registra"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cantidad Mas Movimientos"
         Object.Width           =   4939
      EndProperty
   End
End
Attribute VB_Name = "frmStockCtrlAgeActual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CargaDatosListView
End Sub

Private Sub CargaDatosListView()
Dim R As ADODB.Recordset
Dim Cmd As New Command
Dim L As ListItem
Dim loConec As New DConecta

        Me.lstvAgencias.ListItems.Clear
                
        loConec.AbreConexion
        Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
        Cmd.CommandType = adCmdStoredProc
        Cmd.CommandText = "ATM_RecuperaCantidadActualXAgencia"
        
        Set R = Cmd.Execute
        
        
        Do While Not R.EOF
            Set L = lstvAgencias.ListItems.Add(, , Str(R!nCodAgeArea))
            Call L.ListSubItems.Add(, , R!cNomAgeArea)
            Call L.ListSubItems.Add(, , Str(R!nCantidad))
            Call L.ListSubItems.Add(, , Str(R!nCantSaldo))
            
            R.MoveNext
        Loop
        R.Close
        
        'CerrarConexion
        loConec.CierraConexion
        Set loConec = Nothing
        Set Prm = Nothing
        Set Cmd = Nothing
End Sub

