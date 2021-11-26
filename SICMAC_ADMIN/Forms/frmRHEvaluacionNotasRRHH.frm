VERSION 5.00
Begin VB.Form frmRHEvaluacionNotasRRHH 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   Icon            =   "frmRHEvaluacionNotasRRHH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExaEscritoCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   4140
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8505
      TabIndex        =   6
      Top             =   4125
      Width           =   1095
   End
   Begin VB.CommandButton cmdExaEscritoImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   2385
      TabIndex        =   5
      Top             =   4140
      Width           =   1095
   End
   Begin VB.Frame fraExaEscritoSeleccion 
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
      Height          =   3720
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   9585
      Begin Sicmact.FlexEdit FlexExaEscrito 
         Height          =   3405
         Left            =   60
         TabIndex        =   4
         Top             =   255
         Width           =   9450
         _ExtentX        =   14076
         _ExtentY        =   6006
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Codigo-Nombre-e"
         EncabezadosAnchos=   "500-1500-5500-1500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbRsLoad        =   -1  'True
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbPuntero       =   -1  'True
         RowHeight0      =   240
      End
   End
   Begin VB.CommandButton cmdExaEscritoGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   4140
      Width           =   1095
   End
   Begin VB.ComboBox cmbEva 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   5010
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   4140
      Width           =   1095
   End
   Begin VB.Label lblEva 
      Caption         =   "Evaluacion :"
      Height          =   195
      Left            =   75
      TabIndex        =   8
      Top             =   60
      Width           =   1065
   End
End
Attribute VB_Name = "frmRHEvaluacionNotasRRHH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnTipo As TipoOpe
Dim lnTipEva As RHTipoOpeEvaluacion
Dim lbRegistrar As Boolean

Private Sub cmbEva_Click()
    CargaDatos
End Sub

Private Sub cmbEva_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.cmdEditar.Enabled Then
            cmdEditar.SetFocus
        End If
    End If
End Sub

Private Sub cmdExaEscritoCancelar_Click()
    CargaDatos
    Activa False
End Sub

Private Sub cmdEditar_Click()
    Activa True
    Me.FlexExaEscrito.SetFocus
End Sub

Private Sub cmdExaEscritoGrabar_Click()
    Dim oCurDet As NActualizaProcesoSeleccion
    Dim I As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCurDet = New NActualizaProcesoSeleccion
    
    If MsgBox("Desea Grabar la Información ??? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Set rs = Me.FlexExaEscrito.GetRsNew
        
    oCurDet.ModificaPersonaProEval Right(Me.cmbEva.Text, 6), rs, CInt(lnTipEva), GetMovNro(gsCodUser, gsCodAge)
    
    Set oCurDet = Nothing
    cmdExaEscritoCancelar_Click
End Sub

Private Sub cmdExaEscritoImprimir_Click()
    Dim oEva As NActualizaProcesoSeleccion
    Dim oPrevio As Previo.clsPrevio
    Dim o
    Dim lsCadena As String
    
    Set oPrevio = New Previo.clsPrevio
    Set oEva = New NActualizaProcesoSeleccion
    
    lsCadena = oEva.GetReporteEvaPersonasNotas(Right(Me.cmbEva.Text, 6), gsNomAge, gsEmpresa, gdFecSis, CInt(lnTipEva))
    If lsCadena <> "" Then oPrevio.Show lsCadena, Caption, True, 66
    
    Set oPrevio = Nothing
    Set oEva = Nothing
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Activa(pbValor As Boolean)
    Me.cmdEditar.Visible = Not pbValor
    Me.cmdExaEscritoGrabar.Enabled = pbValor
    Me.cmdExaEscritoCancelar.Visible = pbValor
    Me.fraExaEscritoSeleccion.Enabled = pbValor
    Me.cmdSalir.Enabled = Not pbValor
    Me.cmbEva.Enabled = Not pbValor
    
    If lnTipo = gTipoOpeRegistro Then
        Me.cmdExaEscritoGrabar.Enabled = pbValor
        Me.cmdExaEscritoImprimir.Enabled = False
    ElseIf lnTipo = gTipoOpeMantenimiento Then
        Me.cmdExaEscritoGrabar.Enabled = pbValor
        Me.cmdExaEscritoImprimir.Enabled = pbValor
        Me.cmdExaEscritoImprimir.Enabled = False
    ElseIf lnTipo = gTipoOpeConsulta Then
        Me.cmdExaEscritoGrabar.Enabled = False
        Me.cmdExaEscritoImprimir.Enabled = False
        Me.cmdEditar.Enabled = False
    ElseIf lnTipo = gTipoOpeReporte Then
        Me.cmdExaEscritoGrabar.Enabled = False
        Me.cmdEditar.Enabled = False
    End If
    If lbRegistrar Then
        Me.fraExaEscritoSeleccion.Enabled = True
        Me.FlexExaEscrito.lbEditarFlex = False
    End If
    
End Sub

Private Sub FlexExaEscrito_KeyPress(KeyAscii As Integer)
    Exit Sub
    KeyAscii = 0
End Sub

Private Sub FlexExaEscrito_OnCellChange(pnRow As Long, pnCol As Long)
    If lnTipEva = 4 Then
        Me.FlexExaEscrito.TextMatrix(pnRow, FlexExaEscrito.Cols - 2) = Format((CCur(IIf(Me.FlexExaEscrito.TextMatrix(pnRow, 3) = "", "0", Me.FlexExaEscrito.TextMatrix(pnRow, 3))) + CCur(IIf(Me.FlexExaEscrito.TextMatrix(pnRow, 4) = "", "0", Me.FlexExaEscrito.TextMatrix(pnRow, 4))) + CCur(IIf(Me.FlexExaEscrito.TextMatrix(pnRow, 5) = "", "0", Me.FlexExaEscrito.TextMatrix(pnRow, 5))) + CCur(IIf(Me.FlexExaEscrito.TextMatrix(pnRow, 6) = "", "0", Me.FlexExaEscrito.TextMatrix(pnRow, 6)))) / 4, "#,##0.00")
    End If
End Sub

Private Sub Form_Load()
    Dim rsEva As ADODB.Recordset
    Set rsEva = New ADODB.Recordset
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    Dim oCons As DConstantes
    
    Set rsEva = oEva.GetProcesosEvaluacion(RHProcesoSeleccionEstado.gRHProcSelEstIniCiado)
    CargaCombo rsEva, Me.cmbEva, , 1, 0
    rsEva.Close
    Set rsEva = Nothing
    
    Activa False
    
    If lnTipEva = RHTipoOpeEvaConsolidado Then
        Me.FlexExaEscrito.EncabezadosAlineacion = "C-L-L-R-R-R-R-R-R"
        Me.FlexExaEscrito.EncabezadosAnchos = "500-1500-5500-1500-1500-1500-1500-1500-450"
        Me.FlexExaEscrito.EncabezadosNombres = "#-Codigo-Nombre-Escrito-Psicologico-Entrevista-Curriculum-Promedio-OK"
        Me.FlexExaEscrito.ColumnasAEditar = "X-X-X-3-4-5-6-X-8"
        Me.FlexExaEscrito.ListaControles = "0-0-0-0-0-0-0-0-4"
    Else
        Me.FlexExaEscrito.ColumnasAEditar = "X-X-X-3"
        Me.FlexExaEscrito.EncabezadosAlineacion = "C-L-L-R"
        Me.FlexExaEscrito.EncabezadosAnchos = "500-1500-5500-1500"
        
        If lnTipEva = RHTipoOpeEvaEscrito Then
            Me.FlexExaEscrito.EncabezadosNombres = "#-Codigo-Nombre-Escrito"
        ElseIf lnTipEva = RHTipoOpeEvaPsicologico Then
            Me.FlexExaEscrito.EncabezadosNombres = "#-Codigo-Nombre-Psicologico"
        ElseIf lnTipEva = RHTipoOpeEvaEntrevista Then
            Me.FlexExaEscrito.EncabezadosNombres = "#-Codigo-Nombre-Entrevista"
        ElseIf lnTipEva = RHTipoOpeEvaCurricular Then
            Me.FlexExaEscrito.EncabezadosNombres = "#-Codigo-Nombre-Curriculum"
        End If
    End If
End Sub

Private Sub CargaDatos()
    Dim oCurDet As DActualizaProcesoSeleccion
    Set oCurDet = New DActualizaProcesoSeleccion
    Dim rsEva As ADODB.Recordset
    Set rsEva = New ADODB.Recordset
    
    Set rsEva = oCurDet.GetProcesosEvaluacionDetExamen(Right(Me.cmbEva.Text, 6), CInt(lnTipEva))
    
    If Not (rsEva.BOF And rsEva.EOF) Then
        Me.FlexExaEscrito.rsFlex = rsEva
        If lnTipEva = 4 Then Exit Sub
        Me.FlexExaEscrito.EncabezadosAnchos = "500-1500-5500-1500-1"
    Else
        Me.FlexExaEscrito.Clear
        Me.FlexExaEscrito.Rows = 2
        If lnTipEva = 4 Then Exit Sub
        Me.FlexExaEscrito.EncabezadosAnchos = "500-1500-5500-1500-1"
        Me.FlexExaEscrito.FormaCabecera
    End If
    
    Set oCurDet = Nothing
End Sub

Public Sub Ini(pnTipo As TipoOpe, pnTipoOpeEvaluacion As RHTipoOpeEvaluacion, Optional LlamadaParaRegistro As Boolean = False)
    lnTipo = pnTipo
    lnTipEva = pnTipoOpeEvaluacion
    lbRegistrar = LlamadaParaRegistro
    Me.Show 1
End Sub

Private Sub FlexExaEscrito_RowColChange()
    If lnTipEva <> RHTipoOpeEvaConsolidado And lnTipo = gTipoOpeRegistro Then
        If Me.FlexExaEscrito.TextMatrix(FlexExaEscrito.Row, FlexExaEscrito.Cols - 1) = "0" Then
            Me.FlexExaEscrito.lbEditarFlex = False
        Else
            Me.FlexExaEscrito.lbEditarFlex = True
        End If
    End If
End Sub
