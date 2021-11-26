VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmOPExtDevoluc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno de Devoluciones"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frmOPExtDevoluc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   15
      TabIndex        =   15
      Top             =   5235
      Width           =   6510
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "Extornar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   165
         TabIndex        =   17
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   4995
         TabIndex        =   16
         Top             =   195
         Width           =   1290
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3315
      Left            =   15
      TabIndex        =   13
      Top             =   1890
      Width           =   6525
      Begin MSComctlLib.ListView lvwOperaciones 
         Height          =   3045
         Left            =   90
         TabIndex        =   14
         Top             =   195
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   5371
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
      Height          =   675
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6540
      Begin VB.OptionButton Option1 
         Caption         =   "Remesas"
         Height          =   225
         Left            =   915
         TabIndex        =   2
         Top             =   360
         Width           =   1170
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Devoluciones"
         Height          =   285
         Left            =   4395
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Busqueda:"
      Enabled         =   0   'False
      Height          =   1800
      Left            =   15
      TabIndex        =   3
      Top             =   75
      Width           =   6525
      Begin VB.ComboBox cboDestino 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3870
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   435
         Width           =   2550
      End
      Begin VB.CommandButton cmdOperaciones 
         Caption         =   "C o n s u l t a r   O p e r a c i o n e s"
         Height          =   375
         Left            =   885
         TabIndex        =   5
         Top             =   1290
         Width           =   4635
      End
      Begin VB.ComboBox cboOrigen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   870
         TabIndex        =   4
         Top             =   435
         Width           =   2220
      End
      Begin MSMask.MaskEdBox txtfecIni 
         Height          =   300
         Left            =   1515
         TabIndex        =   7
         Top             =   870
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtFecFin 
         Height          =   300
         Left            =   4650
         TabIndex        =   8
         Top             =   900
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Destino:"
         Height          =   240
         Left            =   3150
         TabIndex        =   12
         Top             =   472
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Origen:"
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   472
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de Final :"
         Height          =   240
         Left            =   3390
         TabIndex        =   10
         Top             =   945
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Inicio :"
         Height          =   240
         Left            =   270
         TabIndex        =   9
         Top             =   915
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmOPExtDevoluc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub cboAgencias_Click()
    cmdOperaciones.Enabled = True
End Sub


Private Sub cmdExtornar_Click()
    If Not IsNumeric(lvwOperaciones.SelectedItem.Text) Then
        MsgBox "No Existe Movimiento Seleccionado", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Call grabarDatos(1, lvwOperaciones.SelectedItem.Text)
    
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
    
    
    Call CargaDatosListView(1, CInt(Mid(cboOrigen.Text, 1, InStr(cboOrigen.Text, "-") - 1)), CInt(Mid(cboDestino.Text, 1, InStr(cboDestino.Text, "-") - 1)))

End Sub

Private Sub cmdOperaciones_Click()
    cmdExtornar.Enabled = True
    
     If Not IsDate(Me.txtfecIni.Text) Or Not IsDate(Me.TxtFecFin.Text) Then
        MsgBox "Fecha Incorrecta", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If (Mid(cboOrigen.Text, 1, InStr(cboOrigen.Text, "-") - 1)) = (Mid(cboDestino.Text, 1, InStr(cboDestino.Text, "-") - 1)) Then
        MsgBox "Los valores de Origen y Destino deben ser diferentes"
        Exit Sub
    End If
    'MsgBox Mid(cboAgencias.Text, 1, InStr(cboAgencias.Text, "-") - 1)
    Call CargaDatosListView(1, CInt(Mid(cboOrigen.Text, 1, InStr(cboOrigen.Text, "-") - 1)), CInt(Mid(cboDestino.Text, 1, InStr(cboDestino.Text, "-") - 1)))
    
      If Me.lvwOperaciones.ListItems.Count > 0 Then
        cmdExtornar.Enabled = True
    Else
        cmdExtornar.Enabled = False
    End If
    
    
End Sub

Private Sub CargaDatosListView(ByVal nAccion As Integer, ByVal nCodAgeOri As Integer, ByVal nCodAgeDest As Integer)
Dim R As ADODB.Recordset
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim L As ListItem
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nAccion", adInteger, adParamInput, , nAccion)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCodAgeOri", adInteger, adParamInput, , nCodAgeOri)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCodAgeDest", adInteger, adParamInput, , nCodAgeDest)
    Cmd.Parameters.Append Prm
    
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pdFecIni", adDate, adParamInput, , CDate(Me.txtfecIni.Text))
    Cmd.Parameters.Append Prm
    
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pdFecFin", adDate, adParamInput, , CDate(Me.TxtFecFin.Text))
    Cmd.Parameters.Append Prm
        
        
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ConsultaOperaciones"
    
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
    Me.txtfecIni.Text = Format(Now, "dd/mm/yyyy")
    Me.TxtFecFin.Text = Format(Now, "dd/mm/yyyy")
    
    Call CargaDatos
    
    Frame1.Enabled = True
    cboDestino.Enabled = False
    cboOrigen.Enabled = False
    cboOrigen.List(1) = gsCodAge & "-" & gsNomAge
    cboOrigen.Text = cboOrigen.List(1)
    cboDestino.Text = cboDestino.List(0)
    
End Sub

Private Sub CargaDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset

    Set Prm = New ADODB.Parameter
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaAgencias"
    
    Set R = Cmd.Execute
    'CboCtas.Clear
    Do While Not R.EOF
         cboOrigen.AddItem R!Codigo & "-" & R!Agencia
         cboDestino.AddItem R!Codigo & "-" & R!Agencia
       R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    oConec.CierraConexion
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub Option1_Click()
    Frame1.Enabled = True
    cboOrigen.Text = cboOrigen.List(0)
    cboDestino.Text = cboDestino.List(1)
    cboDestino.Enabled = True
    cboOrigen.Enabled = False
    
End Sub

Private Sub Option2_Click()
    Frame1.Enabled = True
    cboDestino.Enabled = False
    cboOrigen.Enabled = False
    cboOrigen.List(1) = gsCodAge & "-" & gsNomAge
    cboOrigen.Text = cboOrigen.List(1)
    cboDestino.Text = cboDestino.List(0)
End Sub

