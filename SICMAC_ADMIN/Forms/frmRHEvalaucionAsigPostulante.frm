VERSION 5.00
Begin VB.Form frmRHEvaluacionAsigPostulante 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmRHEvalaucionAsigPostulante.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPostulantesEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1215
      TabIndex        =   10
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton cmdPostulantesCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1230
      TabIndex        =   7
      Top             =   4245
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   4245
      Width           =   1095
   End
   Begin VB.CommandButton cmdPostulantesImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4245
      Width           =   1095
   End
   Begin VB.ComboBox cmbEva 
      Height          =   315
      Left            =   1365
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   75
      Width           =   6720
   End
   Begin VB.Frame fraPostulantesSeleccion 
      Caption         =   "Personas Seleccion"
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
      Height          =   3780
      Left            =   30
      TabIndex        =   0
      Top             =   390
      Width           =   8115
      Begin VB.CommandButton cmdPostulantesEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   6945
         TabIndex        =   9
         Top             =   3345
         Width           =   1095
      End
      Begin VB.CommandButton cmdPostulantesNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   5790
         TabIndex        =   8
         Top             =   3345
         Width           =   1095
      End
      Begin Sicmact.FlexEdit FlexPostulantes 
         Height          =   3015
         Left            =   60
         TabIndex        =   5
         Top             =   270
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   5318
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Nombre"
         EncabezadosAnchos=   "500-1500-5500"
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
         ColumnasAEditar =   "X-1-X"
         ListaControles  =   "0-1-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         lbRsLoad        =   -1  'True
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbPuntero       =   -1  'True
         RowHeight0      =   240
      End
   End
   Begin VB.CommandButton cmdPostulantesGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   4260
      Width           =   1095
   End
   Begin VB.Label lblEva 
      Caption         =   "Evaluacion :"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   135
      Width           =   1065
   End
End
Attribute VB_Name = "frmRHEvaluacionAsigPostulante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnTipo As TipoOpe
Dim lnModo As RHProcesoSeleccionModal
Dim lbGanador As Boolean

Private Sub cmbEva_Click()
    CargaDatos
End Sub

Private Sub cmbEva_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.cmdPostulantesEditar.Enabled Then
            cmdPostulantesEditar.SetFocus
        End If
    End If
End Sub

Private Sub cmdPostulantesCancelar_Click()
    CargaDatos
    Activa False
End Sub

Private Sub cmdPostulantesEditar_Click()
    If Me.cmbEva.Text = "" Then Exit Sub
    Me.fraPostulantesSeleccion.Enabled = True
    Activa True
    If cmdPostulantesNuevo.Enabled Then
        Me.cmdPostulantesNuevo.SetFocus
    ElseIf Me.cmdPostulantesEliminar.Enabled Then
        Me.cmdPostulantesEliminar.SetFocus
    End If
End Sub

Private Sub cmdPostulantesEliminar_Click()
    Me.FlexPostulantes.EliminaFila CLng(FlexPostulantes.TextMatrix(FlexPostulantes.Row, 0))
End Sub

Private Sub cmdPostulantesGrabar_Click()
    Dim oCurDet As DActualizaProcesoSeleccion
    Dim I As Integer
    Set oCurDet = New DActualizaProcesoSeleccion
    
    If MsgBox("Desea Grabar la Información ??? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    oCurDet.EliminaPersonasProSelec Right(Me.cmbEva.Text, 6), Me.FlexPostulantes.GetRsNew
    For I = 1 To Me.FlexPostulantes.Rows - 1
        oCurDet.AgregaPersonaProSelec Right(Me.cmbEva.Text, 6), Me.FlexPostulantes.TextMatrix(I, 1), GetMovNro(gsCodUser, gsCodAge)
    Next I
    Set oCurDet = Nothing
    cmdPostulantesCancelar_Click
End Sub

Private Sub cmdEliminar_Click()

End Sub

Private Sub cmdPostulantesNuevo_Click()
    If Me.cmbEva.Text = "" Then Exit Sub
    
    If Me.FlexPostulantes.TextMatrix(FlexPostulantes.Rows - 1, 0) = "" Then
        FlexPostulantes.AdicionaFila 1
    Else
        FlexPostulantes.AdicionaFila CLng(Me.FlexPostulantes.TextMatrix(FlexPostulantes.Rows - 1, 0)) + 1
    End If
    Me.FlexPostulantes.SetFocus
End Sub

Private Sub cmdPostulantesImprimir_Click()
    Dim oPrevio As Previo.clsPrevio
    Dim lsCadena As String
    Set oPrevio = New Previo.clsPrevio
    Dim oEva As NActualizaProcesoSeleccion
    Set oEva = New NActualizaProcesoSeleccion
    
    lsCadena = oEva.GetReporteEvaPersonas(Right(Me.cmbEva.Text, 6), gsNomAge, gsEmpresa, gdFecSis)
    If lsCadena <> "" Then oPrevio.Show lsCadena, Caption, True, 66
    Set oPrevio = Nothing
    Set oEva = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Activa(pbValor As Boolean)
    Me.cmdPostulantesEditar.Visible = Not pbValor
    Me.cmdPostulantesGrabar.Enabled = pbValor
    Me.cmdPostulantesCancelar.Visible = pbValor
    Me.fraPostulantesSeleccion.Enabled = pbValor
    Me.cmbEva.Enabled = Not pbValor
    
    If lnTipo = gTipoOpeRegistro Then
        Me.cmdPostulantesEliminar.Enabled = False
        Me.cmdPostulantesImprimir.Enabled = False
    ElseIf lnTipo = gTipoOpeConsulta Then
        Me.cmdPostulantesEditar.Enabled = False
        Me.cmdPostulantesImprimir.Enabled = False
    ElseIf lnTipo = gTipoOpeReporte Then
        Me.cmdPostulantesEditar.Enabled = False
    ElseIf lnTipo = gTipoOpeMantenimiento Then
        Me.cmdPostulantesNuevo.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Dim rsEva As ADODB.Recordset
    Set rsEva = New ADODB.Recordset
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    Dim oCons As DConstantes
    
    Set rsEva = oEva.GetProcesosSeleccion(RHProcesoSeleccionEstado.gRHProcSelEstIniCiado)
    CargaCombo rsEva, Me.cmbEva, 200, 1, 0
    Activa False
End Sub

Private Sub CargaDatos()
    Dim oCurDet As DActualizaProcesoSeleccion
    Set oCurDet = New DActualizaProcesoSeleccion
    Dim rsEva As ADODB.Recordset
    
    If Me.cmbEva.Text = "" Then
        Me.cmdPostulantesNuevo.Enabled = False
        Me.cmdPostulantesEliminar.Enabled = False
        Exit Sub
    End If
    
    Set rsEva = New ADODB.Recordset
    Set rsEva = oCurDet.GetProcesosSeleccionDet(Right(Me.cmbEva.Text, 6))
    
    If Not (rsEva.BOF And rsEva.EOF) Then
        Me.FlexPostulantes.rsFlex = rsEva
    Else
        Me.FlexPostulantes.Clear
        Me.FlexPostulantes.Rows = 2
        Me.FlexPostulantes.FormaCabecera
    End If
    Set oCurDet = Nothing
End Sub

Public Sub Ini(pnTipo As TipoOpe, Optional pbGanador As Boolean = False)
    lnTipo = pnTipo
    lbGanador = pbGanador
    Me.Show 1
End Sub
