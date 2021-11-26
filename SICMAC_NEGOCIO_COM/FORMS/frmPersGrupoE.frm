VERSION 5.00
Begin VB.Form frmPersGrupoE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Grupos Economicos"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7875
   Icon            =   "frmPersGrupoE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fracontrol 
      Height          =   585
      Left            =   75
      TabIndex        =   15
      Top             =   3000
      Width           =   7770
      Begin VB.CommandButton cmdexaminar 
         Caption         =   "E&xaminar"
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
         Left            =   3825
         TabIndex        =   23
         Top             =   150
         Width           =   1080
      End
      Begin VB.CommandButton cmdCancelar 
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
         Height          =   345
         Left            =   5700
         TabIndex        =   22
         Top             =   165
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eli&minar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1920
         TabIndex        =   21
         Top             =   165
         Width           =   915
      End
      Begin VB.CommandButton cmdImprimir 
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
         Height          =   375
         Left            =   2925
         Picture         =   "frmPersGrupoE.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Imprimir Solicitud"
         Top             =   150
         Visible         =   0   'False
         Width           =   435
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
         Height          =   345
         Left            =   6675
         TabIndex        =   19
         Top             =   165
         Width           =   975
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   18
         Top             =   165
         Width           =   930
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
         Height          =   345
         Left            =   1005
         TabIndex        =   17
         Top             =   165
         Width           =   900
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   75
         TabIndex        =   16
         Top             =   165
         Width           =   930
      End
   End
   Begin VB.Frame fraVinculado 
      Caption         =   "Datos Vinculado (Opcional)"
      Height          =   1740
      Left            =   75
      TabIndex        =   6
      Top             =   1200
      Width           =   7740
      Begin VB.TextBox txtCargoOtro 
         Height          =   345
         Left            =   1425
         TabIndex        =   25
         Top             =   1275
         Width           =   5790
      End
      Begin VB.TextBox txtPorcen 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1425
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   795
         Width           =   1365
      End
      Begin VB.CommandButton cmdBuscarVin 
         Caption         =   "..."
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
         Height          =   375
         Left            =   7200
         TabIndex        =   9
         Top             =   300
         Width           =   390
      End
      Begin VB.ComboBox cboVinculacion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4350
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   795
         Width           =   2790
      End
      Begin VB.Label Label3 
         Caption         =   "Cargo:"
         Height          =   315
         Left            =   225
         TabIndex        =   24
         Top             =   1275
         Width           =   840
      End
      Begin VB.Label Label5 
         Caption         =   "Porcentaje  (%):"
         Height          =   240
         Left            =   225
         TabIndex        =   13
         Top             =   825
         Width           =   1215
      End
      Begin VB.Label lblVinPersCod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   1440
         TabIndex        =   12
         Top             =   315
         Width           =   1350
      End
      Begin VB.Label lblVinPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   2805
         TabIndex        =   11
         Top             =   315
         Width           =   4350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vinculado  :"
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   330
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Vinculacion:"
         Height          =   240
         Left            =   3000
         TabIndex        =   8
         Top             =   825
         Width           =   1365
      End
   End
   Begin VB.ComboBox cboGrupoE 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   675
      Width           =   5715
   End
   Begin VB.CommandButton cmdBuscaCli 
      Caption         =   "..."
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
      Height          =   375
      Left            =   7275
      TabIndex        =   0
      Top             =   150
      Width           =   390
   End
   Begin VB.Label Label7 
      Caption         =   "Grupo Economico:"
      Height          =   240
      Left            =   75
      TabIndex        =   5
      Top             =   705
      Width           =   1365
   End
   Begin VB.Label LblCliPersCod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   1515
      TabIndex        =   3
      Top             =   165
      Width           =   1350
   End
   Begin VB.Label LblCliPersNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   2880
      TabIndex        =   2
      Top             =   165
      Width           =   4350
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Cliente                 :"
      Height          =   195
      Left            =   75
      TabIndex        =   1
      Top             =   180
      Width           =   1290
   End
End
Attribute VB_Name = "frmPersGrupoE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim nTipoOperacion As Integer '0 Nuevo...1 Modificar

Sub Limpiar_Controles()
    txtPorcen.Text = "0.00"
    LblCliPersCod.Caption = ""
    LblCliPersNombre.Caption = ""
    lblVinPersCod.Caption = ""
    lblVinPersNombre.Caption = ""
    cboGrupoE.ListIndex = -1
    cboVinculacion.ListIndex = -1
    txtCargoOtro.Text = ""
End Sub

Private Sub cboGrupoE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdBuscarVin.SetFocus
End Sub

Private Sub cboVinculacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCargoOtro.SetFocus
End Sub

Private Sub cmdBuscaCli_Click()
Dim oPers As COMDPersona.UCOMPersona
    Set oPers = frmBuscaPersona.Inicio
    If Not oPers Is Nothing Then
        LblCliPersCod.Caption = oPers.sPersCod
        LblCliPersNombre.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
    cboGrupoE.SetFocus
End Sub

Private Sub cmdBuscarVin_Click()
Dim oPers As COMDPersona.UCOMPersona
    Set oPers = frmBuscaPersona.Inicio
    If Not oPers Is Nothing Then
        lblVinPersCod.Caption = oPers.sPersCod
        lblVinPersNombre.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
    txtPorcen.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    Call Limpiar_Controles
    Call Habilita_Grabar(False)
    Call Habilita_Datos(False)
End Sub

Private Sub cmdEditar_Click()
nTipoOperacion = 1
Call Habilita_Grabar(True)
Call Habilita_Datos(True)
cmdgrabar.Enabled = True
cmdEditar.Enabled = False
cmdEliminar.Enabled = False
cmdcancelar.Enabled = True
cmdBuscaCli.Enabled = False 'No cambiar
cboGrupoE.Enabled = False 'No cambiar
End Sub

Private Sub cmdEliminar_Click()
Dim oGrupo As COMDPersona.DCOMGrupoE

If LblCliPersCod.Caption = "" Then Exit Sub

If MsgBox("Esta seguro que desea eliminar la relacion ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
    Set oGrupo = New COMDPersona.DCOMGrupoE
    Call oGrupo.EliminaPersGrupoE(Trim(LblCliPersCod.Caption), CInt(Trim(Right(cboGrupoE.Text, 20))))
    Set oGrupo = Nothing
    cmdEliminar.Enabled = False
    cmdEditar.Enabled = False
End If
End Sub

Private Sub cmdexaminar_Click()
Dim oGrupo As COMDPersona.DCOMGrupoE
Dim rs As ADODB.Recordset
Set oGrupo = New COMDPersona.DCOMGrupoE

frmPersGrupoEConsulta.Show 1

Set rs = oGrupo.CargaDatos(frmPersGrupoEConsulta.cPersCod, frmPersGrupoEConsulta.nGrupoEcon)

If Not rs.EOF Then
    LblCliPersCod.Caption = rs!cPersCodCli
    LblCliPersNombre.Caption = rs!cPersNombreCli
    cboGrupoE.ListIndex = IndiceListaCombo(cboGrupoE, Trim(Str(rs!nGrupoCod)), 2)
    If rs!cPersCodOtro <> "" Then
        lblVinPersCod.Caption = rs!cPersCodOtro
        lblVinPersNombre.Caption = rs!cPersNombreOtro
        txtPorcen.Text = Format(rs!nPorcenOtro, "#0.00")
        cboVinculacion.ListIndex = IndiceListaCombo(cboVinculacion, Trim(Str(rs!nTipoVinculacion)), 2)
        txtCargoOtro.Text = rs!cCargoOtro
    End If
End If
Set oGrupo = Nothing
cmdEditar.Enabled = True
End Sub

Private Sub cmdGrabar_Click()

Dim oGrupo As COMDPersona.DCOMGrupoE

If Valida_Datos = False Then Exit Sub

If MsgBox("Esta seguro de grabar los datos ??", vbQuestion + vbYesNo) = vbNo Then Exit Sub

Set oGrupo = New COMDPersona.DCOMGrupoE
If nTipoOperacion = 0 Then
    If lblVinPersCod.Caption = "" Then
        Call oGrupo.RegistraPersGrupoE(CInt(Trim(Right(cboGrupoE.Text, 20))), LblCliPersCod.Caption, 0, "", 0, "")
    Else
        Call oGrupo.RegistraPersGrupoE(CInt(Trim(Right(cboGrupoE.Text, 20))), LblCliPersCod.Caption, CInt(Trim(Right(cboVinculacion.Text, 20))), lblVinPersCod.Caption, CDbl(txtPorcen.Text), txtCargoOtro.Text)
    End If
Else
    Call oGrupo.ModificaPersGrupoE(CInt(Trim(Right(cboGrupoE.Text, 20))), LblCliPersCod.Caption, CInt(Trim(Right(cboVinculacion.Text, 20))), lblVinPersCod.Caption, CDbl(txtPorcen.Text), txtCargoOtro.Text)
End If
Set oGrupo = Nothing
    cmdgrabar.Enabled = False
    cmdEditar.Enabled = True
    cmdEliminar.Enabled = True
    cmdcancelar.Enabled = False
    Call Habilita_Grabar(False)
    Call Habilita_Datos(False)
End Sub

Private Sub cmdNuevo_Click()

nTipoOperacion = 0
Call Limpiar_Controles
Call Habilita_Grabar(True)
Call Habilita_Datos(True)
    cmdgrabar.Enabled = True
    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    cmdcancelar.Enabled = True
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim oCons As COMDConstantes.DCOMConstantes
Dim rs As ADODB.Recordset
Set oCons = New COMDConstantes.DCOMConstantes
Set rs = oCons.RecuperaConstantes(9069)
Call Llenar_Combo_con_Recordset(rs, cboGrupoE)
Set rs = oCons.RecuperaConstantes(9070)
Call Llenar_Combo_con_Recordset(rs, cboVinculacion)
Set oCons = Nothing

Call CentraForm(Me)
Call Habilita_Grabar(False)
Call Habilita_Datos(False)
cmdEditar.Enabled = False
End Sub

Sub Habilita_Grabar(ByVal pbHabilita As Boolean)
    cmdgrabar.Visible = pbHabilita
    cmdNuevo.Visible = Not pbHabilita
End Sub

Function Valida_Datos() As Boolean
Valida_Datos = True
If LblCliPersCod.Caption = "" Then
    MsgBox "Debe indicar el cliente", vbInformation, "Mensaje"
    cmdBuscaCli.SetFocus
    Valida_Datos = False
    Exit Function
End If
If cboGrupoE.ListIndex = -1 Then
    MsgBox "Debe indicar el grupo economico", vbInformation, "Mensaje"
    Valida_Datos = False
    cboGrupoE.SetFocus
    Exit Function
End If

If lblVinPersCod.Caption <> "" Then
    If cboVinculacion.ListIndex = -1 Then
        MsgBox "Debe indicar el tipo de vinculacion", vbInformation, "Mensaje"
        Valida_Datos = False
        cboVinculacion.SetFocus
    End If
End If
End Function

Sub Habilita_Datos(ByVal pbHabilita As Boolean)
    
    cmdBuscaCli.Enabled = pbHabilita
    cmdBuscarVin.Enabled = pbHabilita
    cboGrupoE.Enabled = pbHabilita
    cboVinculacion.Enabled = pbHabilita
    txtCargoOtro.Enabled = pbHabilita
    txtPorcen.Enabled = pbHabilita
End Sub

Private Sub txtCargoOtro_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtPorcen_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosDecimales(txtPorcen, KeyAscii)
     If KeyAscii = 13 Then cboVinculacion.SetFocus
End Sub

Private Sub txtPorcen_LostFocus()
If txtPorcen.Text = "" Then
    txtPorcen.Text = "0.00"
Else
    txtPorcen.Text = Format(txtPorcen.Text, "#0.00")
End If
End Sub
