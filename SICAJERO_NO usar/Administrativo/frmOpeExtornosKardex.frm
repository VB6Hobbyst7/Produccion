VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmOpeExtornosKardex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extornar Ingresos y Salidas"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   Icon            =   "frmOpeExtornosKardex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Criterio"
      Height          =   915
      Left            =   15
      TabIndex        =   11
      Top             =   -30
      Width           =   6540
      Begin VB.OptionButton Option1 
         Caption         =   "Ingresos"
         Height          =   225
         Left            =   915
         TabIndex        =   13
         Top             =   360
         Width           =   1170
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Salidas"
         Height          =   285
         Left            =   4395
         TabIndex        =   12
         Top             =   300
         Width           =   1425
      End
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   15
      TabIndex        =   6
      Top             =   5970
      Width           =   6510
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "Extornar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   165
         TabIndex        =   8
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   4995
         TabIndex        =   7
         Top             =   195
         Width           =   1290
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3060
      Left            =   15
      TabIndex        =   4
      Top             =   2880
      Width           =   6525
      Begin MSComctlLib.ListView lvwOperaciones 
         Height          =   2730
         Left            =   120
         TabIndex        =   5
         Top             =   195
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   4815
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
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Busqueda:"
      Height          =   1860
      Left            =   15
      TabIndex        =   0
      Top             =   975
      Width           =   6525
      Begin MSMask.MaskEdBox txtfecIni 
         Height          =   300
         Left            =   1485
         TabIndex        =   14
         Top             =   870
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboDestino 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         TabIndex        =   10
         Top             =   285
         Width           =   2175
      End
      Begin VB.CommandButton cmdOperaciones 
         Caption         =   "C o n s u l t a r   O p e r a c i o n e s"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1020
         TabIndex        =   2
         Top             =   1335
         Width           =   4635
      End
      Begin VB.ComboBox cboOrigen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   945
         TabIndex        =   1
         Top             =   270
         Width           =   2250
      End
      Begin MSMask.MaskEdBox TxtFecFin 
         Height          =   300
         Left            =   4620
         TabIndex        =   16
         Top             =   900
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de Final :"
         Height          =   240
         Left            =   3360
         TabIndex        =   17
         Top             =   945
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Inicio :"
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Top             =   915
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "Destino:"
         Height          =   240
         Left            =   3465
         TabIndex        =   9
         Top             =   300
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Origen:"
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmOpeExtornosKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub cboAgencias_Click()
    cmdOperaciones.Enabled = True
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
        
    Call CargaDatosListView(2, CInt(Mid(cboOrigen.Text, 1, InStr(cboOrigen.Text, "-") - 1)), CInt(Mid(cboDestino.Text, 1, InStr(cboDestino.Text, "-") - 1)))
    
End Sub

Private Sub cboAgencias_Change()
    cmdOperaciones.Enabled = True
End Sub

Private Sub cmdExtornar_Click()
    If Not IsNumeric(lvwOperaciones.SelectedItem.Text) Then
        MsgBox "No Existe Movimiento Seleccionado", vbInformation, "Aviso"
        Exit Sub
    End If
    Call grabarDatos(2, lvwOperaciones.SelectedItem.Text)
End Sub

Private Sub cmdOperaciones_Click()
    cmdExtornar.Enabled = False
    
    If Not IsDate(Me.txtfecIni.Text) Or Not IsDate(Me.TxtFecFin.Text) Then
        MsgBox "Fecha Incorrecta", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Call CargaDatosListView(2, CInt(Mid(cboOrigen.Text, 1, InStr(cboOrigen.Text, "-") - 1)), CInt(Mid(cboDestino.Text, 1, InStr(cboDestino.Text, "-") - 1)))
    
    If Me.lvwOperaciones.ListItems.Count > 0 Then
        cmdExtornar.Enabled = True
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
    Call CargaDatos
    Me.txtfecIni.Text = Format(Now, "dd/mm/yyyy")
    Me.TxtFecFin.Text = Format(Now, "dd/mm/yyyy")
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
    cboOrigen.List(0) = "999-Agencia Externa"
    cboOrigen.Text = cboOrigen.List(0)
'    cboDestino.List(1) = gsCodAge & "-" & gsNomAge
'    cboDestino.Text = cboDestino.List(1)
    cboDestino.ListIndex = 0
    cmdOperaciones.Enabled = True
    
End Sub

Private Sub Option2_Click()
    cboOrigen.List(0) = "0-Boveda General"
     cboOrigen.ListIndex = 0

    
    cboDestino.List(1) = "999-Agencia Externa"
    cboDestino.Text = cboDestino.List(1)
    cmdOperaciones.Enabled = True
End Sub
