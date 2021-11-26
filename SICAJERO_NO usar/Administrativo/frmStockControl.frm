VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStockControlAge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Stock de Minimo"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "frmStockControl.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   750
      Left            =   30
      TabIndex        =   5
      Top             =   3690
      Width           =   6255
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   405
         Left            =   4785
         TabIndex        =   6
         Top             =   225
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3630
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   6255
      Begin VB.Frame Frame2 
         Height          =   750
         Left            =   60
         TabIndex        =   1
         Top             =   2760
         Width           =   6150
         Begin VB.TextBox txtStockMin 
            Height          =   285
            Left            =   1230
            TabIndex        =   4
            Text            =   "0"
            Top             =   270
            Width           =   660
         End
         Begin VB.CommandButton CmdActualiza 
            Caption         =   "Actualizar Stock Minimo"
            Height          =   375
            Left            =   4155
            TabIndex        =   2
            Top             =   225
            Width           =   1890
         End
         Begin VB.Label Label1 
            Caption         =   "Stock Minimo :"
            Height          =   270
            Left            =   105
            TabIndex        =   3
            Top             =   270
            Width           =   1125
         End
      End
      Begin MSComctlLib.ListView LstAgeStock2 
         Height          =   2580
         Left            =   105
         TabIndex        =   7
         Top             =   165
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   4551
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agencia"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Stock Minimo"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmStockControlAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta


Private Sub CargaDatos()
    Dim R As ADODB.Recordset
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim L As ListItem
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaStocks"
    
    
    LstAgeStock2.ListItems.Clear
    Set R = Cmd.Execute
        Do While Not R.EOF
         'lstAgeStock.AddItem Right("000" & Trim(Str(R!nCodAgeArea)), 3) & Space(3) & Left(R!cNomAgeArea & Space(25), 25) & Right(Space(5) & Trim(Str(R!nCantidad)), 5) & Right(Space(5) & Trim(Str(R!nMinimo)), 5)
         Set L = LstAgeStock2.ListItems.Add(, , R!nCodAgeArea)
         Call L.ListSubItems.Add(, , R!cNomAgeArea)
         Call L.ListSubItems.Add(, , R!nMinimo)
         
         R.MoveNext
    Loop
    R.Close
    
    'CerrarConexion
    oConec.CierraConexion
    Set Cmd = Nothing
    Set Prm = Nothing
    
End Sub

Private Sub CmdActualiza_Click()
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
       
    If Not IsNumeric(Me.txtStockMin.Text) Then
        MsgBox "Cantidad de Stock Minimo incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If
       
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCodAge", adInteger, adParamInput, , CInt(LstAgeStock2.SelectedItem.Text))
    Cmd.Parameters.Append Prm

        
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCantidad", adInteger, adParamInput, , CInt(Me.txtStockMin.Text))
    Cmd.Parameters.Append Prm
        
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ActualizaStockMinimo"
    Cmd.Execute
    
    'CerrarConexion
    oConec.CierraConexion
    Set Cmd = Nothing
    Set Prm = Nothing
    
    MsgBox "Datos Grabados Correctamente", vbInformation, "Aviso"
    CargaDatos
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    Call CargaDatos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub LstAgeStock2_Click()
    Me.txtStockMin.Text = LstAgeStock2.SelectedItem.ListSubItems(2).Text
End Sub
