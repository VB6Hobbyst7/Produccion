VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManLimRemEnTran 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Limite de Dias de Remesas en Transito"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "frmManLimRemEnTran.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   3240
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6015
      Begin MSComctlLib.ListView lvwAgecias 
         Height          =   2940
         Left            =   60
         TabIndex        =   6
         Top             =   195
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   5186
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
            Object.Width           =   2892
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agencia"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Dias"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   15
      TabIndex        =   1
      Top             =   3870
      Width           =   6060
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   405
         Left            =   4785
         TabIndex        =   4
         Top             =   210
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   15
      TabIndex        =   0
      Top             =   3225
      Width           =   6060
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "Actualizar"
         Height          =   405
         Left            =   2490
         TabIndex        =   7
         Top             =   150
         Width           =   1170
      End
      Begin VB.TextBox TxtNumDias 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1530
         TabIndex        =   3
         Text            =   "0"
         Top             =   195
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad de Dias :"
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   210
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmManLimRemEnTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub CmdActualizar_Click()
Dim Prm As ADODB.Parameter
Dim Cmd As New ADODB.Command
Dim R As New ADODB.Recordset

    If Not IsNumeric(Me.TxtNumDias.Text) Then
        MsgBox "Numero de Dias Incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If CInt(Me.TxtNumDias.Text) <= 0 Then
        MsgBox "Numero de Dias debe ser mayor a Cero", vbInformation, "Aviso"
        Exit Sub
    End If

            
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnDias", adInteger, adParamInput, , CInt(Me.TxtNumDias.Text))
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCodAge", adInteger, adParamInput, , CInt(Me.lvwAgecias.SelectedItem.Text))
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ActualizarDiasLimRemTran"
           
    Cmd.Execute
    
    oConec.CierraConexion

    Set Prm = Nothing
    Set Cmd = Nothing
         
    MsgBox "Datos Grabados Correctamente", vbInformation, "Aviso"
    
    
    Call CargaDatos

End Sub

Private Sub CmdSalir_Click()
        Unload Me
End Sub

Public Sub CargaDatos()
Dim Prm As ADODB.Parameter
Dim Cmd As New ADODB.Command
Dim R As New ADODB.Recordset
Dim L As ListItem
                       
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaDiasLimTran"
    
    Set R = Cmd.Execute
    lvwAgecias.ListItems.Clear
     
    Do While Not R.EOF
       Set L = lvwAgecias.ListItems.Add(, , R!nCodAgeArea)
       Call L.ListSubItems.Add(, , R!cNomAgeArea)
       Call L.ListSubItems.Add(, , R!nDiasLimTransito)
       
       R.MoveNext
    Loop
    
    
    R.Close
    
    oConec.CierraConexion
    Set R = Nothing
    Set Prm = Nothing
    Set Cmd = Nothing
         
         
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    Call CargaDatos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub lvwAgecias_Click()
    Me.TxtNumDias.Text = Me.lvwAgecias.SelectedItem.SubItems(2)
End Sub


