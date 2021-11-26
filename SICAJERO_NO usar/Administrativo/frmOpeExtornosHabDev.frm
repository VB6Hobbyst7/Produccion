VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpeExtornosHabDev 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extornar Habilitaciones y Devoluciones"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   Icon            =   "frmOpeExtornosHabDev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Busqueda:"
      Height          =   1965
      Left            =   15
      TabIndex        =   8
      Top             =   915
      Width           =   6525
      Begin VB.ComboBox cboOrigen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   435
         Width           =   5355
      End
      Begin VB.CommandButton cmdOperaciones 
         Caption         =   "C o n s u l t a r   O p e r a c i o n e s"
         Enabled         =   0   'False
         Height          =   375
         Left            =   945
         TabIndex        =   10
         Top             =   1395
         Width           =   4635
      End
      Begin VB.ComboBox cboDestino 
         Enabled         =   0   'False
         Height          =   315
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   5355
      End
      Begin VB.Label Label1 
         Caption         =   "Origen:"
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   435
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Destino:"
         Height          =   240
         Left            =   150
         TabIndex        =   12
         Top             =   840
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   15
      TabIndex        =   5
      Top             =   5985
      Width           =   6585
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "Extornar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   165
         TabIndex        =   7
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   4995
         TabIndex        =   6
         Top             =   195
         Width           =   1290
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3015
      Left            =   15
      TabIndex        =   3
      Top             =   2880
      Width           =   6525
      Begin MSComctlLib.ListView lvwOperaciones 
         Height          =   2715
         Left            =   105
         TabIndex        =   4
         Top             =   210
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   4789
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro.Trans"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Operacion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Criterio"
      Height          =   915
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   6540
      Begin VB.OptionButton Option1 
         Caption         =   "Habilitaciones"
         Height          =   225
         Left            =   915
         TabIndex        =   2
         Top             =   360
         Width           =   1410
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Devoluciones"
         Height          =   285
         Left            =   4395
         TabIndex        =   1
         Top             =   300
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmOpeExtornosHabDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta


Private Sub CargaDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
        
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCodAge", adInteger, adParamInput, , CInt(gsCodAge))
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaUsuarios"
    
    Set R = Cmd.Execute
    'CboCtas.Clear
    Do While Not R.EOF
        cboOrigen.AddItem R!Codigo & "-" & R!Nombre
         cboDestino.AddItem R!Codigo & "-" & R!Nombre
       R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    oConec.CierraConexion
    
End Sub

Private Sub cmdExtornar_Click()
    Call grabarDatos(3, lvwOperaciones.SelectedItem.Text)
    If Me.lvwOperaciones.ListItems.Count > 0 Then
        cmdExtornar.Enabled = True
    Else
        cmdExtornar.Enabled = False
    End If
End Sub

Private Sub grabarDatos(ByVal nAccion As Integer, ByVal nNumTran As Integer)
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim sResp As String
Dim Inicial As Long
Dim Final As Long
Dim cant As Long
    
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nAccion", adInteger, adParamInput, , nAccion)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCodAge", adInteger, adParamInput, , 0)
    Cmd.Parameters.Append Prm
        
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nNumTran", adInteger, adParamInput, , nNumTran)
    Cmd.Parameters.Append Prm
                
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ExtornarOperaciones"
    Cmd.Execute
    
    
    MsgBox "Los Extornos se registraron con Exito", vbInformation, "Extorno Operaciones"
    
    'CerrarConexion
    oConec.CierraConexion
    Set Cmd = Nothing
    Set Prm = Nothing
    '    End If
    
    
    Call CargaDatosListView
    

End Sub


Private Sub cmdOperaciones_Click()
    cmdExtornar.Enabled = True
    
    
    Call CargaDatosListView
                                  
    If Me.lvwOperaciones.ListItems.Count > 0 Then
        cmdExtornar.Enabled = True
    Else
        cmdExtornar.Enabled = False
    End If
End Sub

Private Sub CargaDatosListView()
Dim R As ADODB.Recordset
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim L As ListItem

    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
        
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCodusuOri", adVarChar, adParamInput, 20, Left(Me.cboOrigen.Text, 4))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCodusuDest", adVarChar, adParamInput, 20, Left(Me.cboDestino.Text, 4))
    Cmd.Parameters.Append Prm
    
        
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ConsultaOperacionesHabDev"
    
    lvwOperaciones.ListItems.Clear
    Set R = Cmd.Execute
        Do While Not R.EOF
         'lstAgeStock.AddItem Right("000" & Trim(Str(R!nCodAgeArea)), 3) & Space(3) & Left(R!cNomAgeArea & Space(25), 25) & Right(Space(5) & Trim(Str(R!nCantidad)), 5) & Right(Space(5) & Trim(Str(R!nMinimo)), 5)
         Set L = lvwOperaciones.ListItems.Add(, , R!Transaccion)
         Call L.ListSubItems.Add(, , R!Fecha)
         Call L.ListSubItems.Add(, , R!Operacion)
         Call L.ListSubItems.Add(, , R!Cantidad)
         
         R.MoveNext
    Loop
    R.Close
    
    'CerrarConexion
    oConec.CierraConexion
    Set Cmd = Nothing
    Set Prm = Nothing
    
End Sub



Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    Me.cmdOperaciones.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub Option1_Click()

    cboOrigen.Enabled = False
    cboOrigen.List(0) = "BOVE"
    cboOrigen.Text = cboOrigen.List(0)
    Call CargaDatos
    cboDestino.Enabled = True
    lvwOperaciones.ListItems.Clear
    Me.cmdExtornar.Enabled = False
End Sub

Private Sub Option2_Click()
    cboDestino.Enabled = False
    cboDestino.List(0) = "BOVE"
    cboDestino.Text = cboDestino.List(0)
    Call CargaDatos
    cboOrigen.Enabled = True
    cboOrigen.Text = cboOrigen.List(0)
    lvwOperaciones.ListItems.Clear
    Me.cmdExtornar.Enabled = False
End Sub
