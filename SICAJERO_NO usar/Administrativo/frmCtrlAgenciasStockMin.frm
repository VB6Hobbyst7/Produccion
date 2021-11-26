VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCtrlAgenciasStockMin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agencias Con Stock Minimo"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "frmCtrlAgenciasStockMin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   120
      TabIndex        =   2
      Top             =   2985
      Width           =   6195
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   360
         Left            =   4890
         TabIndex        =   3
         Top             =   225
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2940
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   6225
      Begin MSComctlLib.ListView lvwAgeCStockMin 
         Height          =   2580
         Left            =   90
         TabIndex        =   1
         Top             =   210
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   4551
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "         Agencia "
            Object.Width           =   2892
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "    Cantidad"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "   Stock Minimo"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCtrlAgenciasStockMin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CargaDatos
End Sub

Private Sub CargaDatos()
Dim R As ADODB.Recordset
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim L As ListItem
Dim loConec As New DConecta

    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaAgenciasCStockMin"
        
    lvwAgeCStockMin.ListItems.Clear
    Set R = Cmd.Execute
        Do While Not R.EOF
         'lstAgeStock.AddItem Right("000" & Trim(Str(R!nCodAgeArea)), 3) & Space(3) & Left(R!cNomAgeArea & Space(25), 25) & Right(Space(5) & Trim(Str(R!nCantidad)), 5) & Right(Space(5) & Trim(Str(R!nMinimo)), 5)
         Set L = lvwAgeCStockMin.ListItems.Add(, , R!Agencia)
         Call L.ListSubItems.Add(, , R!nCantidad)
         Call L.ListSubItems.Add(, , R!nMinimos)
         
         R.MoveNext
    Loop
    R.Close
    
    'CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    Set Cmd = Nothing
    Set Prm = Nothing
    
End Sub
