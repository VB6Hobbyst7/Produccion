VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmPersonaRolMantenimiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Persona - Rol "
   ClientHeight    =   4905
   ClientLeft      =   1485
   ClientTop       =   2145
   ClientWidth     =   8385
   Icon            =   "FrmPersonaRolMantenimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Txttemp 
      Height          =   315
      Left            =   7020
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "txtcodigo"
      Top             =   1350
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fraRegPersonaRol 
      Caption         =   "Registra Rol a Persona "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   195
      TabIndex        =   12
      Top             =   3510
      Visible         =   0   'False
      Width           =   6720
      Begin VB.ComboBox cbxEstado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmPersonaRolMantenimiento.frx":030A
         Left            =   4860
         List            =   "FrmPersonaRolMantenimiento.frx":0314
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   435
         Width           =   735
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5670
         TabIndex        =   7
         Top             =   720
         Width           =   1005
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5670
         TabIndex        =   6
         Top             =   360
         Width           =   1005
      End
      Begin VB.TextBox txtNombrePersona 
         Height          =   300
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "TXTNOMBRE"
         Top             =   435
         Width           =   3660
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         TabIndex        =   4
         Top             =   810
         Width           =   1020
      End
      Begin VB.Label Label3 
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   225
         TabIndex        =   17
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4980
         TabIndex        =   16
         Top             =   225
         Width           =   735
      End
      Begin VB.Label LblCodPers 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   195
         TabIndex        =   14
         Tag             =   "TXTCODIGO"
         Top             =   435
         Width           =   1140
      End
   End
   Begin VB.Frame fraRol 
      Caption         =   "Roles de Persona"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   780
      Left            =   135
      TabIndex        =   10
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox cboRolPersona 
         Height          =   315
         ItemData        =   "FrmPersonaRolMantenimiento.frx":031E
         Left            =   900
         List            =   "FrmPersonaRolMantenimiento.frx":0320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   5550
      End
      Begin VB.Label Label1 
         Caption         =   "Rol"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   375
         Width           =   675
      End
   End
   Begin MSComctlLib.ImageList imglst2 
      Left            =   7470
      Top             =   1170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPersonaRolMantenimiento.frx":0322
            Key             =   "grupos"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   8
      ToolTipText     =   "Salir al Menu Principal"
      Top             =   2520
      Width           =   1020
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Editar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7125
      TabIndex        =   2
      ToolTipText     =   "Modificar Datos de Usuario"
      Top             =   720
      Width           =   1110
   End
   Begin VB.CommandButton cmdagre 
      Caption         =   "A&gregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7110
      TabIndex        =   1
      ToolTipText     =   "Ingresar un Nuevo Usuario"
      Top             =   180
      Width           =   1140
   End
   Begin VB.Frame fraUsuario 
      Caption         =   "&Personas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2460
      Left            =   135
      TabIndex        =   9
      Top             =   975
      Width           =   6780
      Begin MSComctlLib.ListView LstPersona 
         Height          =   2040
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3598
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImgLst"
         SmallIcons      =   "ImgLst"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7409
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Estado"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImgLst 
      Left            =   7470
      Top             =   1770
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPersonaRolMantenimiento.frx":07E8
            Key             =   "usuario"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmPersonaRolMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Public Codi As String
Dim opcAgregaEdita As String * 1
Dim vCodigoRol As String * 3
Dim Codigos(20) As String
' Carga Lista Personas segun el rol seleccionado en el combo
Sub CargaPersonas(vCodigo As String)
Dim RegPers As New ADODB.Recordset
Dim SQL1 As String
Dim L As ListItem
'Dim CodNom As String
 SQL1 = "SELECT DISTINCT RolPerson.cCodPers, Persona.cNomPers, RolPerson.cEstado " _
       & "FROM RolPerson INNER JOIN " & gcCentralPers & "Persona Persona ON RolPerson.cCodPers = Persona.cCodPers  " _
       & "WHERE (RolPerson.cTipRol = '" & vCodigo & " ') "
         
 LstPersona.ListItems.Clear
 RegPers.Open SQL1, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
 If RSVacio(RegPers) Then
 Else
 Do While Not RegPers.EOF
    Set L = LstPersona.ListItems.Add(, , RegPers!cCodPers, "usuario", "usuario")
    L.SubItems(1) = Trim(RegPers!cNomPers)
    L.SubItems(2) = Trim(RegPers!cEstado)
    RegPers.MoveNext
 Loop
 End If
 RegPers.Close
 Set RegPers = Nothing
End Sub
' Muestra los datos de la persona seleccionada
Sub UbicaPersona(lsCodPersona As String)
Dim SQL1 As String
Dim RegPers As New ADODB.Recordset
 SQL1 = "SELECT RolPerson.cCodPers, Persona.cNomPers, RolPerson.cEstado " _
      & "FROM RolPerson INNER JOIN " & gcCentralPers & "Persona Persona ON RolPerson.cCodPers = Persona.cCodPers  " _
      & "WHERE (RolPerson.cCodPers = '" & lsCodPersona & " ' AND RolPerson.cTipRol = '" & vCodigoRol & "') "
 RegPers.Open SQL1, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
 
 LblCodPers = RegPers!cCodPers
 txtNombrePersona = Trim(RegPers!cNomPers)
 cbxEstado.Text = RegPers!cEstado
 RegPers.Close
 Set RegPers = Nothing
End Sub
Private Sub cboRolPersona_Click()
If cboRolPersona.ListIndex = -1 Then
    cboRolPersona.ListIndex = 0
End If
vCodigoRol = Codigos(cboRolPersona.ListIndex)
Call CargaPersonas(vCodigoRol)
End Sub

Private Sub cboRolPersona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdagre.SetFocus
End If
End Sub
Function ExisteRol() As Boolean
Dim SQL1 As String
Dim RegPers As New ADODB.Recordset
SQL1 = "SELECT * FROM RolPerson " _
    & "WHERE cCodPers = '" & LblCodPers & " ' AND cTipRol = '" & vCodigoRol & "'"
   RegPers.Open SQL1, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
   If RSVacio(RegPers) Then
      ExisteRol = False
   Else
      ExisteRol = True
   End If
End Function

Private Sub cbxEstado_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    CmdAceptar.SetFocus
 End If
End Sub
Private Sub CmdAceptar_Click()
Dim SQL1 As String
Dim Nombre As String
Dim Msg  As String
On Error GoTo ERROR
 If opcAgregaEdita = "A" Then ' Agregar
' Verificar si ya se encuentra registrado el rol
   If ExisteRol = False Then  ' No existe
      SQL1 = "INSERT INTO RolPerson " & _
             "(cCodPers,cTipRol, cEstado, dFecMod, cCodUsu) " & _
             " VALUES ('" & LblCodPers & "','" & vCodigoRol & "','" & cbxEstado.Text & "','" & _
             FechaHora(gdFecSis) & "','" & gsCodUser & "')"
        If MsgBox("Desea Guardar los Datos", vbInformation + vbYesNo, "Aviso") = vbYes Then
            dbCmact.BeginTrans
            dbCmact.Execute SQL1
            dbCmact.CommitTrans
        End If
        
    Else
       MsgBox "Persona ya se encuentra registrada en el rol", vbInformation, "Aviso"
    End If
   Me.cmdedit.Enabled = True
  Else  ' Editar
    SQL1 = "UPDATE RolPerson SET " _
    & " cEstado = '" & cbxEstado.Text & "'," _
    & " dFecMod = '" & Format(gdFecSis, "mm/dd/yyyy") & Space(1) & Format(Time, "hh:mm:ss") & "'," _
    & " cCodUsu = '" & gsCodUser & "' WHERE cCodPers = '" & LblCodPers & "'" _
    & " AND cTipRol = '" & vCodigoRol & "'"
    If MsgBox("Desea Guardar las Modificaciones", vbInformation + vbYesNo, "Aviso") = vbYes Then
       dbCmact.BeginTrans
       dbCmact.Execute SQL1
       If Err Then
         dbCmact.RollbackTrans
       Else
         dbCmact.CommitTrans
       End If
    End If
    Me.cmdagre.Enabled = True
 End If
    Call CargaPersonas(vCodigoRol)
    fraRegPersonaRol.Enabled = False
    fraRegPersonaRol.Visible = False
    fraRol.Enabled = True
    fraUsuario.Enabled = True
Me.Height = 3915
   Exit Sub
ERROR:
     MsgBox Err.Description
End Sub

Private Sub cmdagre_Click()
 opcAgregaEdita = "A"
 If cboRolPersona.ListIndex = -1 Then
     cboRolPersona.ListIndex = 0
 End If
 vCodigoRol = Codigos(cboRolPersona.ListIndex)
 fraRol.Enabled = False
 fraUsuario.Enabled = False
 fraRegPersonaRol.Enabled = True
 fraRegPersonaRol.Visible = True
 LblCodPers.Caption = ""
 txtNombrePersona.Text = ""
 LblCodPers.Enabled = True
 txtNombrePersona.Enabled = True
 CmdBuscar.Enabled = True
 CmdBuscar.SetFocus
 Me.cmdedit.Enabled = False
 Me.Height = 5340
End Sub

Private Sub cmdBuscar_Click()
Dim CadNom As String
 Txttemp.Text = ""
 Call frmBuscaCli.Inicia(Me, True)
 cbxEstado.SetFocus
End Sub

Private Sub cmdcancelar_Click()
 Call CargaPersonas(vCodigoRol)
 fraRegPersonaRol.Enabled = False
 fraRegPersonaRol.Visible = False
 fraRol.Enabled = True
 fraUsuario.Enabled = True
 Me.cmdagre.Enabled = True
 Me.cmdedit.Enabled = True
 Me.CmdAceptar.Enabled = False
 Me.Height = 3915
End Sub

Private Sub cmdedit_Click()
 opcAgregaEdita = "E"
If cboRolPersona.ListIndex = -1 Then
    cboRolPersona.ListIndex = 0
End If
 vCodigoRol = Codigos(cboRolPersona.ListIndex)
 If LstPersona.ListItems.Count <= 0 Then
   MsgBox "Selecione un Codigo", vbInformation, "Aviso"
   Exit Sub
 Else
   UbicaPersona (LstPersona.SelectedItem)
 End If
 Me.Height = 5340
 fraRol.Enabled = False
 fraUsuario.Enabled = False
 fraRegPersonaRol.Enabled = True
 fraRegPersonaRol.Visible = True
 LblCodPers.Enabled = False
 txtNombrePersona.Enabled = False
 CmdBuscar.Enabled = False
 cbxEstado.SetFocus
 Me.cmdagre.Enabled = False
End Sub

Private Sub cmdSalir_Click()
 Unload Me
End Sub

Function CargaCombo()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim i As Integer
SQL = "SELECT cNomTab,cCodtab,cValor FROM " & gcCentralCom & "tablacod WHERE ccodTab LIKE '71__'" _
      & "ORDER BY cNomTab"
rs.Open SQL, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
If RSVacio(rs) Then
  MsgBox "Comunicar a Sistemas para mantenimiento a Roles de Persona", vbInformation, "Aviso"
Else
   i = 0
   Do While Not rs.EOF
   Me.cboRolPersona.AddItem rs!cNomTab
   Codigos(i) = Trim(rs!cValor)
   i = i + 1
   rs.MoveNext
   Loop
End If
rs.Close
Set rs = Nothing
End Function
Private Sub Form_Activate()
    Me.cbxEstado.ListIndex = 0
End Sub
Private Sub Form_Load()
  CargaCombo
  Call CargaPersonas("001")
End Sub

Private Sub LblCodPers_Change()
If LblCodPers.Caption <> "" Then
   Me.CmdAceptar.Enabled = True
End If
End Sub

Private Sub LstPersona_DblClick()
    cmdedit_Click
End Sub
Private Sub LstPersona_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If LstPersona.ListItems.Count > 0 Then
           Me.cmdedit.SetFocus
        Else
            MsgBox "Lista de Usuarios Esta Vacia", vbInformation, "Aviso"
        End If
    End If
End Sub
