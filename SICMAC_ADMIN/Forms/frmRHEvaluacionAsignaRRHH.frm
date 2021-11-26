VERSION 5.00
Begin VB.Form frmRHEvaluacionAsignaRRHH 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   Icon            =   "frmRHEvaluacionAsignaRRHH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   4170
      Width           =   1095
   End
   Begin VB.Frame fraPersopnaSeleccion 
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
      Left            =   0
      TabIndex        =   5
      Top             =   315
      Width           =   8115
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   5775
         TabIndex        =   7
         Top             =   3345
         Width           =   1095
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   6945
         TabIndex        =   6
         Top             =   3345
         Width           =   1095
      End
      Begin Sicmact.FlexEdit lFlexEdit 
         Height          =   3015
         Left            =   60
         TabIndex        =   8
         Top             =   270
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   5318
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Nombre"
         EncabezadosAnchos=   "500-1700-5500"
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
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   495
         RowHeight0      =   285
      End
   End
   Begin VB.ComboBox cmbEva 
      Height          =   315
      Left            =   1335
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   0
      Width           =   6720
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   2385
      TabIndex        =   3
      Top             =   4170
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7050
      TabIndex        =   2
      Top             =   4170
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   4170
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   4170
      Width           =   1095
   End
   Begin VB.Label lblEva 
      Caption         =   "Evaluacion :"
      Height          =   195
      Left            =   150
      TabIndex        =   10
      Top             =   60
      Width           =   1065
   End
End
Attribute VB_Name = "frmRHEvaluacionAsignaRRHH"
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
        If cmdEditar.Enabled Then
            cmdEditar.SetFocus
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    CargaDatos
    Activa False
End Sub

Private Sub cmdEditar_Click()
    If Me.cmbEva.Text = "" Then Exit Sub
    Me.fraPersopnaSeleccion.Enabled = True
    Activa True
    If cmdNuevo.Enabled Then
        Me.cmdNuevo.SetFocus
    Else
        If Me.cmdEliminar.Enabled Then
            Me.cmdEliminar.SetFocus
        End If
    End If
End Sub

Private Sub cmdEliminar_Click()
    Me.lFlexEdit.EliminaFila CLng(lFlexEdit.TextMatrix(lFlexEdit.Row, 0))
End Sub

Private Sub CmdGrabar_Click()
    Dim oCurDet As DActualizaProcesoSeleccion
    Dim I As Integer
    Set oCurDet = New DActualizaProcesoSeleccion
    
    If MsgBox("Desea Grabar la Información ??? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    oCurDet.AgregaPersonasEvaluacion Right(Me.cmbEva.Text, 6), lFlexEdit.GetRsNew, GetMovNro(gsCodUser, gsCodAge)
    
    Set oCurDet = Nothing
    cmdCancelar_Click
End Sub

Private Sub cmdNuevo_Click()
    If Me.cmbEva.Text = "" Then Exit Sub
    
    If Me.lFlexEdit.TextMatrix(lFlexEdit.Rows - 1, 0) = "" Then
        lFlexEdit.AdicionaFila 1
    Else
        lFlexEdit.AdicionaFila CLng(Me.lFlexEdit.TextMatrix(lFlexEdit.Rows - 1, 0)) + 1
    End If
    
    Me.lFlexEdit.SetFocus
End Sub

Private Sub cmdImprimir_Click()
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
    Me.cmdEditar.Visible = Not pbValor
    Me.cmdGrabar.Enabled = pbValor
    Me.cmdCancelar.Visible = pbValor
    Me.fraPersopnaSeleccion.Enabled = pbValor
    Me.cmbEva.Enabled = Not pbValor
    
    If lnTipo = gTipoOpeRegistro Then
        Me.cmdEliminar.Enabled = False
        Me.cmdImprimir.Enabled = False
    ElseIf lnTipo = gTipoOpeConsulta Then
        Me.cmdEditar.Enabled = False
        Me.cmdImprimir.Enabled = False
    ElseIf lnTipo = gTipoOpeReporte Then
        Me.cmdEditar.Enabled = False
    ElseIf lnTipo = gTipoOpeMantenimiento Then
        Me.cmdNuevo.Enabled = False
        Me.cmdImprimir.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Dim rsEva As ADODB.Recordset
    Set rsEva = New ADODB.Recordset
    Dim oEva As DActualizaProcesoSeleccion
    Set oEva = New DActualizaProcesoSeleccion
    Dim oCons As DConstantes
    
    Set rsEva = oEva.GetProcesosEvaluacion(RHProcesoSeleccionEstado.gRHProcSelEstIniCiado)
    CargaCombo rsEva, Me.cmbEva, 200, 1, 0
    Activa False
End Sub

Private Sub CargaDatos()
    Dim oCurDet As DActualizaProcesoSeleccion
    Dim oEva As DActualizaDatosArea
    Set oCurDet = New DActualizaProcesoSeleccion
    Set oEva = New DActualizaDatosArea
    Dim rsEva As ADODB.Recordset
    
    If Me.cmbEva.Text = "" Then
        Me.cmdNuevo.Enabled = False
        Me.cmdEliminar.Enabled = False
        Exit Sub
    End If
    
    Set rsEva = New ADODB.Recordset
    Set rsEva = oCurDet.GetProcesosEvaluacionDet(Right(Me.cmbEva.Text, 6))
    
    If Not (rsEva.BOF And rsEva.EOF) Then
        Me.lFlexEdit.rsFlex = rsEva
    Else
        Me.lFlexEdit.Clear
        Me.lFlexEdit.Rows = 2
        Me.lFlexEdit.FormaCabecera
    End If
    
    Set rsEva = oEva.GetAreasRRHHEvaluacion(Right(Me.cmbEva.Text, 6))
    If Not (rsEva.BOF And rsEva.EOF) Then Me.lFlexEdit.rsTextBuscar = rsEva
    
    Set oCurDet = Nothing
    Set rsEva = Nothing
End Sub

Public Sub Ini(pnTipo As TipoOpe, Optional pbGanador As Boolean = False)
    lnTipo = pnTipo
    lbGanador = pbGanador
    Me.Show 1
End Sub


