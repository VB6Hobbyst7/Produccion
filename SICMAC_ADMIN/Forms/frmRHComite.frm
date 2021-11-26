VERSION 5.00
Begin VB.Form frmRHComite 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   Icon            =   "frmRHComite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1125
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8745
      TabIndex        =   3
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame fraComite 
      Caption         =   "Comite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4260
      Left            =   45
      TabIndex        =   4
      Top             =   15
      Width           =   9720
      Begin Sicmact.FlexEdit FlexComite 
         Height          =   3540
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9540
         _ExtentX        =   16828
         _ExtentY        =   6244
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Nombre-Cargo"
         EncabezadosAnchos=   "350-1800-4000-3000"
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
         ColumnasAEditar =   "X-1-X-3"
         ListaControles  =   "0-1-0-3"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         RowHeight0      =   240
      End
      Begin VB.CommandButton cmdNuevoComite 
         Caption         =   "N&uevo"
         Height          =   375
         Left            =   7605
         TabIndex        =   6
         Top             =   3810
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminarComite 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   8640
         TabIndex        =   5
         Top             =   3810
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1125
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
End
Attribute VB_Name = "frmRHComite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbEditado As Boolean
Dim lsCodigo As String
Dim lsEvaCod As String
Dim rs As ADODB.Recordset
Dim rsIni As ADODB.Recordset
Dim lbIni As Boolean

Public Function Ini(psEvalCod As String, rsI As ADODB.Recordset) As ADODB.Recordset
    lsEvaCod = psEvalCod
    lbIni = True
    Set rsIni = rsI
    Show 1
    Set Ini = rs
End Function

Public Function IniMan(psEvalCod As String, rsI As ADODB.Recordset) As ADODB.Recordset
    lsEvaCod = psEvalCod
    lbIni = False
    Set rsIni = rsI
    Show 1
    Set IniMan = rs
End Function

Private Sub cmdCancelar_Click()
    Activa False
    'CargaComite
    lbEditado = False
End Sub

Private Sub cmdEditar_Click()
    Activa True
    lbEditado = True
End Sub

Private Sub cmdEliminarComite_Click()
    Me.FlexComite.EliminaFila Me.FlexComite.Row
End Sub

Private Sub cmdGrabar_Click()
    Dim sqlG As String
    Dim oCom As NActualizaProcesoSeleccion
    Set oCom = New NActualizaProcesoSeleccion
    
    oCom.AgregaComiteProSelec lsEvaCod, Me.FlexComite.GetRsNew, GetMovNro(gsCodUser, gsCodAge)
    
    Activa False
    CargaComite
    lbEditado = False
    Set oCom = Nothing
End Sub

Private Sub cmdNuevoComite_Click()
    Dim oPersona As UPersona
    If Me.FlexComite.TextMatrix(Me.FlexComite.Rows - 1, 0) = "#" Then
        FlexComite.Rows = 2
    End If
    
    If Me.FlexComite.TextMatrix(Me.FlexComite.Rows - 1, 0) = "" Then
        FlexComite.AdicionaFila 1
    Else
        FlexComite.AdicionaFila CLng(Me.FlexComite.TextMatrix(FlexComite.Rows - 1, 0)) + 1
    End If
    FlexComite.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsE As ADODB.Recordset
    Dim oCons As DConstantes
    Set oCons = New DConstantes
    Set rsE = New ADODB.Recordset
    Set rsE = oCons.GetConstante(gRHEvaluacionComite)
    Me.FlexComite.CargaCombo rsE
    Set oCons = Nothing
    If Not lbIni Then
        Activa False
    Else
        Set Me.FlexComite.Recordset = rsIni
        Me.fraComite.Enabled = True
        Me.cmdEditar.Enabled = False
        Me.cmdGrabar.Enabled = False
    End If
    
    CargaComite
End Sub

Private Sub Activa(pbValor As Boolean)
    Me.fraComite.Enabled = pbValor
    Me.cmdEditar.Visible = Not pbValor
    Me.cmdSalir.Enabled = Not pbValor
    If Not lbIni Then Me.cmdGrabar.Enabled = pbValor
    Me.cmdCancelar.Visible = pbValor
End Sub

Private Sub CargaComite()
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    'If lsEvaCod <> "" Then
    '    Set Me.FlexComite.Recordset = oEva.GetNomPersonasComite(lsEvaCod)
    'Else
        
    'End If
    Set oEva = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Me.FlexComite.GetRsNew
End Sub


