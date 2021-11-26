VERSION 5.00
Begin VB.Form frmRHEvaluacionNotas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   Icon            =   "frmRHEvaluacionNotas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExaCurrilarEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   4155
      Width           =   1095
   End
   Begin VB.CommandButton cmdExaCurrilarRegistrar 
      Caption         =   "&Registrar"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   4155
      Width           =   1095
   End
   Begin VB.ComboBox cmbEva 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   15
      Width           =   5010
   End
   Begin VB.Frame fraExaCurrilarSeleccion 
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
      TabIndex        =   6
      Top             =   375
      Width           =   9585
      Begin Sicmact.FlexEdit Flex 
         Height          =   3405
         Left            =   60
         TabIndex        =   7
         Top             =   240
         Width           =   9450
         _ExtentX        =   14076
         _ExtentY        =   6006
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Codigo-Nombre-e"
         EncabezadosAnchos=   "500-1500-5500-1500"
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
         ColumnasAEditar =   "X-X-X-3"
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
   Begin VB.CommandButton cmdExaCurrilarImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   4155
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8505
      TabIndex        =   9
      Top             =   4140
      Width           =   1095
   End
   Begin VB.CommandButton cmdExaCurrilarCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   4155
      Width           =   1095
   End
   Begin VB.CommandButton cmdExaCurrilarGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4155
      Width           =   1095
   End
   Begin VB.Label lblEva 
      Caption         =   "Evaluacion :"
      Height          =   195
      Left            =   75
      TabIndex        =   8
      Top             =   75
      Width           =   1065
   End
End
Attribute VB_Name = "frmRHEvaluacionNotas"
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
        MsgBox "hola"
    End If
End Sub

Private Sub cmdExaCurrilarCancelar_Click()
    CargaDatos
    Activa False
End Sub

Private Sub cmdExaCurrilarEditar_Click()
    Activa True
End Sub

Private Sub cmdExaCurrilarGrabar_Click()
    Dim oCurDet As NActualizaProcesoSeleccion
    Dim I As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCurDet = New NActualizaProcesoSeleccion
    
    If MsgBox("Desea Grabar la Información ??? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Set rs = Me.Flex.GetRsNew
        
    oCurDet.ModificaPersonaProSelec Right(Me.cmbEva.Text, 6), rs, CInt(lnTipEva), GetMovNro(gsCodUser, gsCodAge)
    
    Set oCurDet = Nothing
    cmdExaCurrilarCancelar_Click
End Sub

Private Sub cmdExaCurrilarImprimir_Click()
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

Private Sub cmdExaCurrilarNuevo_Click()
    Me.Flex.AdicionaFila
End Sub

Private Sub cmdExaCurrilarRegistrar_Click()
    If Flex.TextMatrix(Me.Flex.Row, 8) = "." Then
        frmRHEmpleado.IniRegistroEva Flex.TextMatrix(Me.Flex.Row, 1), Right(Me.cmbEva.Text, 6)
    Else
        MsgBox "No Puede Ingresar a una Perosna que no a sido Seleccionado.", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Activa(pbValor As Boolean)
    Me.cmdExaCurrilarEditar.Visible = Not pbValor
    Me.cmdExaCurrilarGrabar.Enabled = pbValor
    Me.cmdExaCurrilarCancelar.Visible = pbValor
    Me.fraExaCurrilarSeleccion.Enabled = pbValor
    Me.cmdSalir.Enabled = Not pbValor
    Me.cmbEva.Enabled = Not pbValor
    
    If lnTipo = gTipoOpeRegistro Then
        Me.cmdExaCurrilarGrabar.Enabled = pbValor
        Me.cmdExaCurrilarImprimir.Enabled = False
    ElseIf lnTipo = gTipoOpeMantenimiento Then
        Me.cmdExaCurrilarGrabar.Enabled = pbValor
        Me.cmdExaCurrilarImprimir.Enabled = pbValor
        Me.cmdExaCurrilarImprimir.Enabled = False
    ElseIf lnTipo = gTipoOpeConsulta Then
        Me.cmdExaCurrilarGrabar.Enabled = False
        Me.cmdExaCurrilarImprimir.Enabled = False
        Me.cmdExaCurrilarEditar.Enabled = False
    ElseIf lnTipo = gTipoOpeReporte Then
        Me.cmdExaCurrilarGrabar.Enabled = False
        Me.cmdExaCurrilarEditar.Enabled = False
    End If
    If lbRegistrar Then
        Me.fraExaCurrilarSeleccion.Enabled = True
        Me.Flex.lbEditarFlex = False
    End If
    
End Sub

Private Sub FLEX_KeyPress(KeyAscii As Integer)
    Exit Sub
    KeyAscii = 0
End Sub

Private Sub Flex_OnCellChange(pnRow As Long, pnCol As Long)
    If lnTipEva = 4 Then
        Me.Flex.TextMatrix(pnRow, Flex.Cols - 2) = Format((CCur(IIf(Me.Flex.TextMatrix(pnRow, 3) = "", "0", Me.Flex.TextMatrix(pnRow, 3))) + CCur(IIf(Me.Flex.TextMatrix(pnRow, 4) = "", "0", Me.Flex.TextMatrix(pnRow, 4))) + CCur(IIf(Me.Flex.TextMatrix(pnRow, 5) = "", "0", Me.Flex.TextMatrix(pnRow, 5))) + CCur(IIf(Me.Flex.TextMatrix(pnRow, 6) = "", "0", Me.Flex.TextMatrix(pnRow, 6)))) / 4, "#,##0.00")
    End If
End Sub


Private Sub Form_Load()
    Dim rsEva As ADODB.Recordset
    Set rsEva = New ADODB.Recordset
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    Dim oCons As DConstantes
'    Me.cmdXXXRegistrar.Visible = lbRegistrar
'
'    Set rsEva = oEva.GetProcesosSeleccion(RHProcesoSeleccionEstado.gRHProcSelEstIniCiado)
'    CargaCombo rsEva, Me.cmbEva, , 1, 0
'    rsEva.Close
'    Set rsEva = Nothing
'
'    Activa False
'
'    If lnTipEva = RHTipoOpeEvaConsolidado Then
'        Me.Flex.EncabezadosAlineacion = "C-L-L-R-R-R-R-R-R"
'        Me.Flex.EncabezadosAnchos = "500-1500-5500-1500-1500-1500-1500-1500-450"
'        Me.Flex.EncabezadosNombres = "#-Codigo-Nombre-Escrito-Psicologico-Entrevista-Curriculum-Promedio-OK"
'        Me.Flex.ColumnasAEditar = "X-X-X-3-4-5-6-X-8"
'        Me.Flex.ListaControles = "0-0-0-0-0-0-0-0-4"
'    Else
'        Me.Flex.ColumnasAEditar = "X-X-X-3"
'        Me.Flex.EncabezadosAlineacion = "C-L-L-R"
'        Me.Flex.EncabezadosAnchos = "500-1500-5500-1500"
'
'        If lnTipEva = RHTipoOpeEvaEscrito Then
'            Me.Flex.EncabezadosNombres = "#-Codigo-Nombre-Escrito"
'        ElseIf lnTipEva = RHTipoOpeEvaPsicologico Then
'            Me.Flex.EncabezadosNombres = "#-Codigo-Nombre-Psicologico"
'        ElseIf lnTipEva = RHTipoOpeEvaEntrevista Then
'            Me.Flex.EncabezadosNombres = "#-Codigo-Nombre-Entrevista"
'        ElseIf lnTipEva = RHTipoOpeEvaCurricular Then
'            Me.Flex.EncabezadosNombres = "#-Codigo-Nombre-Curriculum"
'        End If
'    End If
End Sub

Private Sub CargaDatos()
    Dim oCurDet As DActualizaProcesoSeleccion
    Set oCurDet = New DActualizaProcesoSeleccion
    Dim rsEva As ADODB.Recordset
    Set rsEva = New ADODB.Recordset
    
    Set rsEva = oCurDet.GetProcesosSeleccionDetExamen(Right(Me.cmbEva.Text, 6), CInt(lnTipEva))
    
    If Not (rsEva.BOF And rsEva.EOF) Then
        Me.Flex.rsFlex = rsEva
        If lnTipEva = 4 Then Exit Sub
        Me.Flex.EncabezadosAnchos = "500-1500-5500-1500-1"
    Else
        Me.Flex.Clear
        Me.Flex.Rows = 2
        If lnTipEva = 4 Then Exit Sub
        Me.Flex.EncabezadosAnchos = "500-1500-5500-1500-1"
        Me.Flex.FormaCabecera
    End If
    
    Set oCurDet = Nothing
End Sub

Public Sub Ini(pnTipo As TipoOpe, pnTipoOpeEvaluacion As RHTipoOpeEvaluacion, Optional LlamadaParaRegistro As Boolean = False)
    lnTipo = pnTipo
    lnTipEva = pnTipoOpeEvaluacion
    lbRegistrar = LlamadaParaRegistro
    Me.Show 1
End Sub

Private Sub flex_RowColChange()
    If lnTipEva <> RHTipoOpeEvaConsolidado And lnTipo = gTipoOpeRegistro Then
        If Me.Flex.TextMatrix(Flex.Row, Flex.Cols - 1) = "0" Then
            Me.Flex.lbEditarFlex = False
        Else
            Me.Flex.lbEditarFlex = True
        End If
    End If
End Sub
