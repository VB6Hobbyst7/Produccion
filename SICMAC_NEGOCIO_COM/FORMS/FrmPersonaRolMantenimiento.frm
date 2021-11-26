VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPersonaRolMantenimiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Persona - Rol "
   ClientHeight    =   4905
   ClientLeft      =   1485
   ClientTop       =   2145
   ClientWidth     =   7965
   Icon            =   "FrmPersonaRolMantenimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   6360
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
         Left            =   5475
         List            =   "FrmPersonaRolMantenimiento.frx":0314
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
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
         Left            =   2400
         TabIndex        =   6
         Top             =   840
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
         Left            =   1320
         TabIndex        =   5
         Top             =   840
         Width           =   1005
      End
      Begin VB.TextBox txtNombrePersona 
         Height          =   300
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "TXTNOMBRE"
         Top             =   435
         Width           =   4020
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
         TabIndex        =   3
         Top             =   810
         Width           =   1020
      End
      Begin VB.Label LblCodPers 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Tag             =   "TXTCODIGO"
         Top             =   435
         Width           =   1260
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
         TabIndex        =   14
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
         Left            =   5400
         TabIndex        =   13
         Top             =   225
         Width           =   615
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
      TabIndex        =   9
      Top             =   120
      Width           =   6360
      Begin VB.ComboBox cboRolPersona 
         Height          =   315
         ItemData        =   "FrmPersonaRolMantenimiento.frx":033F
         Left            =   720
         List            =   "FrmPersonaRolMantenimiento.frx":0341
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   5415
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
         TabIndex        =   10
         Top             =   375
         Width           =   675
      End
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
      Left            =   6720
      TabIndex        =   7
      ToolTipText     =   "Salir al Menu Principal"
      Top             =   1200
      Width           =   1140
   End
   Begin VB.CommandButton cmdEditar 
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
      Left            =   6720
      TabIndex        =   2
      ToolTipText     =   "Modificar Datos de Usuario"
      Top             =   600
      Width           =   1140
   End
   Begin VB.CommandButton cmdAgregar 
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
      Left            =   6720
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
      Height          =   2535
      Left            =   135
      TabIndex        =   8
      Top             =   900
      Width           =   6360
      Begin MSComctlLib.ListView LstPersona 
         Height          =   2175
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3836
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
            Text            =   "Codigo"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Persona"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Estado"
            Object.Width           =   882
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmPersonaRolMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* Mantenimiento de Roles de Persona
'Archivo:  FrmPersonaRolMantenimiento.frm
'LAYG:
'Resumen:  Nos permite registrar los Roles que Cumple la Persona
Option Explicit

Dim fsOperacion As String * 1
Dim fnTipoRol As Integer

Private Sub cboRolPersona_Click()
If cboRolPersona.ListIndex = -1 Then
    cboRolPersona.ListIndex = 0
End If
fnTipoRol = Right(Trim(cboRolPersona.Text), 2)
Call CargaPersonas(fnTipoRol)
End Sub

Private Sub cbxEstado_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    CmdAceptar.SetFocus
 End If
End Sub
Private Sub CmdAceptar_Click()
Dim oPers As COMDPersona.DCOMRoles
Dim sMensaje As String

On Error GoTo Error
Set oPers = New COMDPersona.DCOMRoles
Call oPers.GrabarOperacionRolPersona(fsOperacion, LblCodPers, fnTipoRol, Trim(Right(cboRolPersona.Text, 2)), _
                                    Right(cbxEstado.Text, 1), sMensaje)
Set oPers = Nothing

If sMensaje <> "" Then
    MsgBox sMensaje, vbInformation, "Mensaje"
    Exit Sub
End If

If fsOperacion = "A" Then ' Agregar
    Me.cmdEditar.Enabled = True
    Me.Height = 3915
Else  ' Editar
    Me.cmdAgregar.Enabled = True
End If
    
Call CargaPersonas(fnTipoRol)
Call cboRolPersona_Click

fraRegPersonaRol.Enabled = False
fraRegPersonaRol.Visible = False
fraRol.Enabled = True
fraUsuario.Enabled = True
Me.Height = 3915
   Exit Sub
Error:
     MsgBox Err.Description
End Sub

Private Sub CmdAgregar_Click()
 fsOperacion = "A"
 If cboRolPersona.ListIndex = -1 Then
     cboRolPersona.ListIndex = 0
 End If
 fnTipoRol = Right(Trim(cboRolPersona.Text), 2)
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
 Me.cmdEditar.Enabled = False
 Me.Height = 5340
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String
Dim liFil As Integer
Dim ls As String
Dim loColPFunc As dColPFunciones
On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
Set loPers = frmBuscaPersona.Inicio

If Not loPers Is Nothing Then
    lsPersCod = loPers.sPersCod
    LblCodPers = loPers.sPersCod
    txtNombrePersona = loPers.sPersNombre
    cbxEstado.SetFocus
End If
Exit Sub

ControlError:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdCancelar_Click()
 Call CargaPersonas(fnTipoRol)
 fraRegPersonaRol.Enabled = False
 fraRegPersonaRol.Visible = False
 fraRol.Enabled = True
 fraUsuario.Enabled = True
 Me.cmdAgregar.Enabled = True
 Me.cmdEditar.Enabled = True
 Me.CmdAceptar.Enabled = False
 Me.Height = 3915
End Sub

Private Sub CmdEditar_Click()
 fsOperacion = "E"
If cboRolPersona.ListIndex = -1 Then
    cboRolPersona.ListIndex = 0
End If
 fnTipoRol = Right(Trim(cboRolPersona.Text), 2)
 If LstPersona.ListItems.Count <= 0 Then
   MsgBox "Selecione un Codigo", vbInformation, "Aviso"
   Exit Sub
 Else
   Me.LblCodPers = LstPersona.SelectedItem
   Me.txtNombrePersona = LstPersona.SelectedItem.ListSubItems(1)
   'Me.cbxEstado.Text = LstPersona.SelectedItem.ListSubItems(3)
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
 Me.cmdAgregar.Enabled = False
End Sub

Private Sub cmdsalir_Click()
 Unload Me
End Sub

Private Function CargaTipoRoles()
Dim lr As ADODB.Recordset
Dim i As Integer
Dim oPers As COMDPersona.DCOMRoles

On Error GoTo ERRORCargaTipoRoles
    Set oPers = New COMDPersona.DCOMRoles
    Set lr = oPers.CargaTipoRoles
    Set oPers = Nothing
    cboRolPersona.Clear
    If lr.BOF And lr.EOF Then
        MsgBox "Comunicar a Sistemas para mantenimiento a Roles de Persona", vbInformation, "Aviso"
    Else
        i = 0
        Do While Not lr.EOF
            Me.cboRolPersona.AddItem UCase(lr!cConsDescripcion) & Space(120) & lr!nConsValor
            lr.MoveNext
        Loop
    End If
    lr.Close
    Set lr = Nothing
    Exit Function
ERRORCargaTipoRoles:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

' Carga Lista Personas segun el rol seleccionado en el combo
Private Sub CargaPersonas(ByVal pnTipoRol As Integer)

Dim lRegPers As New ADODB.Recordset
Dim L As ListItem

Dim oPers As COMDPersona.DCOMRoles
 
 LstPersona.ListItems.Clear

    Set oPers = New COMDPersona.DCOMRoles
    Set lRegPers = oPers.CargaPersonas(pnTipoRol)
    Set oPers = Nothing

    If lRegPers.BOF And lRegPers.EOF Then
    Else
        Do While Not lRegPers.EOF
           Set L = LstPersona.ListItems.Add(, , lRegPers!cPersCod)
           L.SubItems(1) = Trim(lRegPers!cPersNombre)
           L.SubItems(2) = lRegPers!PersEstado
           lRegPers.MoveNext
        Loop
    End If
 lRegPers.Close
 Set lRegPers = Nothing

End Sub

Private Sub Form_Load()
  CargaTipoRoles
  'Call CargaPersonas(1)
  Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub LblCodPers_Change()
If LblCodPers.Caption <> "" Then
   Me.CmdAceptar.Enabled = True
End If
End Sub

Private Sub LstPersona_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If LstPersona.ListItems.Count > 0 Then
           Me.cmdEditar.SetFocus
        Else
            MsgBox "Lista de Usuarios Esta Vacia", vbInformation, "Aviso"
        End If
    End If
End Sub


'Private Function ExistePersonaRol(ByVal psPersCod As String, ByVal pnTipoRol As PersTipo) As Boolean
'Dim lsSQL As String
'Dim lrReg As ADODB.Recordset
'Dim loConecta As COMConecta.DCOMConecta
'
'lsSQL = "SELECT * FROM PersTpo " & _
'        "WHERE cPersCod = '" & psPersCod & " ' AND nPersTipo = '" & pnTipoRol & "' "
'
'    Set loConecta = New COMConecta.DCOMConecta
'    loConecta.AbreConexion
'    Set lrReg = loConecta.CargaRecordSet(lsSQL)
'    loConecta.CierraConexion
'    Set loConecta = Nothing
'    If lrReg.BOF And lrReg.EOF Then
'       ExistePersonaRol = False
'    Else
'       ExistePersonaRol = True
'   End If
'End Function
