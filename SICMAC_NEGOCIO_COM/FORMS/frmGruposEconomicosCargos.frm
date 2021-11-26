VERSION 5.00
Begin VB.Form frmGruposEconomicosCargos 
   Caption         =   "Registrar Vinculados - Cargos"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7980
   Icon            =   "frmGruposEconomicosCargos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grupo Economico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.Frame Frame2 
         Caption         =   "Acciones y Cargos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   0
         TabIndex        =   7
         Top             =   1440
         Width           =   7695
         Begin VB.TextBox txtAcciones 
            Height          =   285
            Left            =   1680
            TabIndex        =   10
            Top             =   1320
            Width           =   2175
         End
         Begin VB.ComboBox cboOtroCargo 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   840
            Width           =   4335
         End
         Begin VB.ComboBox cboCargo 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label lblTCargo 
            Caption         =   "Otro Cargo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Cargo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblAcciones 
            Caption         =   "Acciones %"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1320
            Width           =   1215
         End
      End
      Begin VB.Label lblPersonaVinculado 
         Caption         =   "[Mostrar]"
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label lblEmpresaVinculado 
         Caption         =   "[Mostrar]"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label lblGrupoEconomico 
         Caption         =   "[Mostrar]"
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "Vinculado           :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Empresa             :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo                :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmGruposEconomicosCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnGrupo As Integer
Dim lsPersCodEmpresa As String
Dim lsPersCodVinculado As String

Private Sub actualizar()
Dim oGrup As COMDpersona.DCOMGrupoE
    Dim rs As ADODB.Recordset
    Set oGrup = New COMDpersona.DCOMGrupoE
    Set rs = oGrup.ActualizarPersGrupoEconomicoCargos(lnGrupo, lsPersCodEmpresa, lsPersCodVinculado, Right(CboCargo.Text, 4), Right(cboOtroCargo.Text, 4), CInt(txtAcciones.Text))
    MsgBox "Datos se guardaron correctamente", vbApplicationModal
    Call cmdSalir_Click
End Sub
Private Sub cmdGrabar_Click()
    Call actualizar
End Sub

Private Sub cmdModificar_Click()
    Call actualizar
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Function MostrarDatos(ByVal nGrupo As Integer, ByVal sPersCodEmpresa As String, ByVal sPersCodVinculado As String) As Integer
    Dim nRetorno As Integer
    Dim oGrup As COMDpersona.DCOMGrupoE
    Dim rs As ADODB.Recordset
    Set oGrup = New COMDpersona.DCOMGrupoE
    Set rs = oGrup.ObtenerPersGrupoEconomico(nGrupo, sPersCodEmpresa, sPersCodVinculado)
    If rs.EOF Or rs.BOF Then
        nRetorno = 0
    Else
        Do Until rs.EOF
            lblGrupoEconomico.Caption = rs!cDesGrupoEconomico
            lblEmpresaVinculado.Caption = rs!cPersEmpresa
            lblPersonaVinculado.Caption = rs!cPersVinculado
            CboCargo.ListIndex = IndiceListaCombo(CboCargo, rs!nCargo)
            cboOtroCargo.ListIndex = IndiceListaCombo(cboOtroCargo, rs!nCargoOtro)
            txtAcciones.Text = rs!nPorcenOtro
            rs.MoveNext
        Loop
        nRetorno = 1
    End If
    Set oGrup = Nothing
    rs.Close
End Function
Public Sub Nuevo(ByVal pnGrupo As Integer, ByVal psPersCodEmpresa As String, ByVal psPersCodVinculado As String)
    lnGrupo = pnGrupo
    lsPersCodEmpresa = psPersCodEmpresa
    lsPersCodVinculado = psPersCodVinculado
    Call MostrarDatos(lnGrupo, lsPersCodEmpresa, lsPersCodVinculado)
    cmdGrabar.Enabled = True
    cmdModificar.Enabled = False
    Show 1
End Sub
Public Sub Modificar(ByVal pnGrupo As Integer, ByVal psPersCodEmpresa As String, ByVal psPersCodVinculado As String)
    lnGrupo = pnGrupo
    lsPersCodEmpresa = psPersCodEmpresa
    lsPersCodVinculado = psPersCodVinculado
    Call MostrarDatos(lnGrupo, lsPersCodEmpresa, lsPersCodVinculado)
    cmdGrabar.Enabled = False
    cmdModificar.Enabled = True
    Show 1
End Sub

Private Sub Form_Load()
    Dim oCons As COMDConstantes.DCOMConstantes
    Dim rs As ADODB.Recordset
    Set oCons = New COMDConstantes.DCOMConstantes
    Set rs = oCons.RecuperaConstantes(9981)
    Call Llenar_Combo_con_Recordset(rs, CboCargo)
    Set rs = Nothing
    Set oCons = Nothing
    Set oCons = New COMDConstantes.DCOMConstantes
    Set rs = oCons.RecuperaConstantes(9982)
    Call Llenar_Combo_con_Recordset(rs, cboOtroCargo)
    Set rs = Nothing
    Set oCons = Nothing
End Sub
