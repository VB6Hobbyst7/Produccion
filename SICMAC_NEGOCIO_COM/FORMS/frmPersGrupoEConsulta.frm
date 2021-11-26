VERSION 5.00
Begin VB.Form frmPersGrupoEConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Grupos Economicos"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6780
   Icon            =   "frmPersGrupoEConsulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtGrupoE 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Text            =   "TxtGrupoE"
      Top             =   600
      Width           =   3855
   End
   Begin SICMACT.FlexEdit FeGrupoE 
      Height          =   2265
      Left            =   225
      TabIndex        =   2
      Top             =   1200
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   3995
      Cols0           =   3
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-cPersCod-Cliente"
      EncabezadosAnchos=   "600-0-5600"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X"
      ListaControles  =   "0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L"
      FormatosEdit    =   "0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      ColWidth0       =   600
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.ComboBox cboGrupoE 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   5190
   End
   Begin VB.Label Label1 
      Caption         =   "Grupo Nuevo"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Grupo Economico:"
      Height          =   240
      Left            =   75
      TabIndex        =   1
      Top             =   180
      Width           =   1365
   End
End
Attribute VB_Name = "frmPersGrupoEConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nGrupoEcon As Integer
Public cPersCod As String

Private Sub cboGrupoE_Change()
Dim oGrupo As COMDPersona.DCOMGrupoE
Set oGrupo = New COMDPersona.DCOMGrupoE
Dim rs As ADODB.Recordset

Set rs = oGrupo.BuscarPersonasXGrupoEcon(CInt(Trim(Right(cboGrupoE.Text, 20))))
Set oGrupo = Nothing

With FeGrupoE
    .Clear
    .FormaCabecera
    .Rows = 2
    .rsFlex = rs
End With

End Sub

Private Sub cboGrupoE_Click()
    Call cboGrupoE_Change
End Sub

Private Sub cmdGrabar_Click()

End Sub

Private Sub FeGrupoE_DblClick()
    nGrupoEcon = CInt(Trim(Right(cboGrupoE.Text, 20)))
    cPersCod = Trim(FeGrupoE.TextMatrix(FeGrupoE.Row, 1))
    Unload Me
End Sub

Private Sub FeGrupoE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call FeGrupoE_DblClick
End Sub

Private Sub Form_Load()
CentraForm Me
    
Dim oCons As COMDConstantes.DCOMConstantes
Dim rs As ADODB.Recordset
Set oCons = New COMDConstantes.DCOMConstantes
Set rs = oCons.RecuperaConstantes(9069)
Set oCons = Nothing

Call Llenar_Combo_con_Recordset(rs, cboGrupoE)
cboGrupoE.Enabled = True
End Sub
