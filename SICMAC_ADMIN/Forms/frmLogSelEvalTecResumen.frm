VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogSelEvalTecResumen 
   Caption         =   "Resumen de Evaluacion tecnica"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11595
   Icon            =   "frmLogSelEvalTecResumen.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   11595
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10020
      TabIndex        =   10
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar"
      Height          =   375
      Left            =   8685
      TabIndex        =   9
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdPuntaje 
      Caption         =   "Aprobar"
      Height          =   375
      Left            =   7350
      TabIndex        =   12
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdActaBuenaPro 
      Caption         =   "Buena Pro"
      Height          =   375
      Left            =   6015
      TabIndex        =   13
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar Proceso"
      Height          =   375
      Left            =   4680
      TabIndex        =   20
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdempate 
      Caption         =   "Empate"
      Height          =   375
      Left            =   3360
      TabIndex        =   22
      Top             =   6360
      Width           =   1335
   End
   Begin VB.ComboBox cmbperiodo 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   180
      Width           =   1695
   End
   Begin VB.Frame s 
      Caption         =   "Proceso de  Seleccion"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   11415
      Begin VB.TextBox txtPuntMaximo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6480
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtPuntMinimo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txttipo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox txtdescripcion 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   8895
      End
      Begin Sicmact.TxtBuscar txtSeleccionA 
         Height          =   315
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
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
         TipoBusqueda    =   2
         sTitulo         =   ""
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Proceso"
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
         Height          =   195
         Index           =   7
         Left            =   7440
         TabIndex        =   19
         Top             =   120
         Width           =   1350
      End
      Begin VB.Label lblestado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
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
         Height          =   195
         Left            =   7560
         TabIndex        =   18
         Top             =   480
         Width           =   660
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C00000&
         BorderStyle     =   4  'Dash-Dot
         FillColor       =   &H8000000D&
         Height          =   495
         Left            =   7440
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblPuntMaximo 
         AutoSize        =   -1  'True
         Caption         =   "Punt.Maximo"
         Height          =   195
         Left            =   5400
         TabIndex        =   17
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblPuntMinimo 
         AutoSize        =   -1  'True
         Caption         =   "Punt. Minimo"
         Height          =   195
         Left            =   3480
         TabIndex        =   16
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label5 
         Caption         =   "Numero"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexEvaluacionTecnica 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      FixedRows       =   2
      FixedCols       =   0
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483624
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.OLE OLE1 
      Height          =   255
      Left            =   11160
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Periodo"
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
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   660
   End
End
Attribute VB_Name = "frmLogSelEvalTecResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim clsDGnral As DLogGeneral
Dim clsDGAdqui As DLogAdquisi
Dim ClsNAdqui As NActualizaProcesoSelecLog
Dim oCons As DConstantes
Dim saccion As String
Dim psTpoEval As String
Dim psFrmTpo As String
Dim psReqNro As Long

'Pa exportar
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Private Sub cmbperiodo_Click()
txtSeleccionA.Text = ""
txttipo.Text = ""
txtDescripcion.Text = ""
If psTpoEval = "T" Then
    flexEvaluacionTecnica.Clear
    flexEvaluacionTecnica.Rows = 3
    FormateaGrilla
ElseIf psTpoEval = "E" Then
    flexEvaluacionTecnica.Clear
    flexEvaluacionTecnica.Rows = 3
    FormateaGrilla_Eco
End If
Me.txtSeleccionA.rs = clsDGAdqui.LogSeleccionLista(cmbperiodo.Text)
End Sub


Private Sub cmdActaBuenaPro_Click()
Dim nestadoProc As Integer

If txtSeleccionA.Text = "" Then Exit Sub

nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(frmLogSelEvalTecResumen.txtSeleccionA.Text)
If nestadoProc <> SelEstProcesoCerrado Then
        MsgBox "El Procesos de Seleccion  " + frmLogSelEvalTecResumen.txtSeleccionA.Text + " aun no esta Cerrado", vbInformation, "Estado del proceso " + frmLogSelEvalTecResumen.txtSeleccionA.Text + " aun no esta cerrado"
        Exit Sub
End If

frmLogSelActaBuenaPro.Show vbModal
Exit Sub

End Sub

Private Sub cmdCerrar_Click()
Dim nResult As Integer
Dim npuntaje As Double
Dim scodBien As String
Dim sCodigo As String
Dim nValida As Integer
Dim nestadoProc As Integer
Dim nvalidaEmpate As Integer
'validar estado
'Validar Calificacion Tecnica este Ingresada

If txtSeleccionA.Text = "" Then Exit Sub
If txtSeleccionA.Text = "" Then Exit Sub
nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccionA.Text)
If nestadoProc = SelEstProcesoCerrado Then
        MsgBox "El Procesos de Seleccion  " + txtSeleccionA.Text + " ya esta Cerrado", vbInformation, "Estado del proceso " + txtSeleccionA.Text + " ya esta Cerrado"
        Exit Sub
End If
If nestadoProc <> SelEstProcesoFinEvaluacion Then
        MsgBox "El Procesos de Seleccion " + txtSeleccionA.Text + " No tiene el estado de Fin de Evaluacion", vbInformation, "Estado del proceso" + txtSeleccionA.Text + " Deberia estar en Fin de Evaluacion"
        Exit Sub
End If

nvalidaEmpate = clsDGAdqui.ValidaLogSelEmpate(Trim(txtSeleccionA.Text))

If nvalidaEmpate > 0 Then
   MsgBox "en este proceso existe empates ", vbInformation, "Existen empates"
   Exit Sub
End If


If MsgBox("Desea Dar Por Terminado este Proceso de Seleccion Nº " & txtSeleccionA.Text, vbQuestion + vbYesNo, "Se cambiara a estado de Cerrado ") = vbNo Then Exit Sub

        Screen.MousePointer = 11
        clsDGAdqui.ActualizaEstadoProcesoSeleccion txtSeleccionA.Text, TpoLogSelEstProceso.SelEstProcesoCerrado
        nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccionA.Text)
        lblestado.Caption = clsDGAdqui.CargaLogSelEstadoDesc(nestadoProc)
        MsgBox "Se Cerro el Proceso de Manera Satisfactoria", vbInformation, "Puntajes Grabados Correctamente"
        Screen.MousePointer = 0



End Sub

Private Sub cmdempate_Click()
Dim nvalidaEmpate As Integer
Dim nvalidaDesempate As Integer

If txtSeleccionA.Text = "" Then
   MsgBox "Seleccione un numero de Proceso ", vbInformation, "Seleccione Numero de Proceso"
   Exit Sub
End If
nvalidaEmpate = clsDGAdqui.ValidaLogSelEmpate(Trim(txtSeleccionA.Text))
nvalidaDesempate = clsDGAdqui.ValidaLogSelDesempate(Trim(txtSeleccionA.Text))
If nvalidaEmpate = 0 And nvalidaDesempate = 0 Then
   MsgBox "No existe empates en este Proceso", vbInformation, "No existe ni existio Empates "
   Exit Sub
End If
frmLogSelDesempate.Show 1
End Sub

Private Sub cmdExportar_Click()
Dim i As Integer
Dim n As Long
Dim lsArchivoN As String
Dim lbLibroOpen As Boolean
Dim lsCadAnt As String
Dim lnIni As Integer
Dim j As Integer
Dim sNombreArchivo As String
Dim nNumCols, nNumRows As Integer
Dim oConec As DConecta
Dim sHora As String
On Error Resume Next
'VERIFICAR NUMERO DE PROVEEDORES
nNumCols = (flexEvaluacionTecnica.Cols - 3)
nNumRows = flexEvaluacionTecnica.Rows - 1


sNombreArchivo = "LogSel-" + txtSeleccionA.Text + "-" + Format(Time, "hhmmss")
lsArchivoN = App.path & "\" + sNombreArchivo + ".xls"


OLE1.Class = "ExcelWorkSheet"
lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
If Not lbLibroOpen Then
   Err.Clear
   'Set objExcel = CreateObject("Excel.Application")
   If Err.Number Then
      MsgBox "Can't open Excel."
   End If
   Exit Sub
End If

Set xlHoja1 = xlLibro.Worksheets(1)
ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
Dim band  As Boolean
Dim letra As String
lnIni = 0
xlHoja1.Cells(2, 1).value = "PERIODO :" & cmbperiodo.Text
xlHoja1.Range("A2:B2").Merge
xlHoja1.Cells(3, 1).value = "PROCESO DE SELECCION : " & txtSeleccionA.Text
xlHoja1.Cells(4, 1).value = "TIPO :" & txttipo.Text
xlHoja1.Cells(5, 1).value = "DESCRIPCION PROCESO :" & txtDescripcion.Text
xlHoja1.Range("A5:K5").Merge
Dim g As Integer
For g = 7 To 7 + nNumRows
    CuadroExcel xlHoja1, 1, 7, 3 + nNumCols, g, False
Next
'CargaLogSelComiteRep
Set rs = clsDGAdqui.CargaLogSelComiteRep(txtSeleccionA.Text)
If rs.EOF = True Then
    MsgBox "No se ha definido miebros para el comite  ", vbInformation, "No Existen Miembros del Comite"
    Exit Sub
End If

If psTpoEval = "T" Then
    i = reporte_Tecnico()
ElseIf psTpoEval = "E" Then
    i = reporte_economico()
ElseIf psTpoEval = "R" Then
    i = reporte_comparativo()
End If

'--------------------------------------------------------------------------------------------------
 
'---------------------------------------------------------------------------------------------------
Imprime_comite i
OLE1.Class = "ExcelWorkSheet"
ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
OLE1.SourceDoc = lsArchivoN
OLE1.Verb = 1
OLE1.Action = 1
OLE1.DoVerb -1

End Sub

Private Sub cmdPuntaje_Click()
Dim nResult As Integer
Dim npuntaje As Double
Dim scodBien As String
Dim sCodigo As String
Dim nValida As Integer
Dim nestadoProc As Integer
'validar estado
'Validar Calificacion Tecnica este Ingresada
nValida = clsDGAdqui.ValidaLogSelCalificacionTecnica(txtSeleccionA.Text)

If Right(txtDescripcion.Text, 1) = 2 Then
Else
    If nValida = 0 Then
        MsgBox "Antes de Grabar Los Puntajes , debe Ingresar las Calificaciones tecnicas Respectivas", vbInformation, "Ingrese las Calificaciones Tecnicas"
        Exit Sub
    End If
End If
If txtSeleccionA.Text = "" Then Exit Sub
nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccionA.Text)
If nestadoProc = SelEstProcesoCerrado Or nestadoProc = SelEstProcesoCancelado Then
                 MsgBox "No se puede Modificar,El Procesos de Seleccion " + txtSeleccionA.Text + " Tiene un estado de Anulado o de Cerrado", vbInformation, "Estado del proceso" + txtSeleccionA.Text + " ya esta Cerrado o Cancelado"
                 Exit Sub
End If

If MsgBox("Desea grabar los puntajes de la Evaluacion Economica del Proceso de Seleccion Nº " & txtSeleccionA.Text, vbQuestion + vbYesNo, "Guardar Puntajes de Evaluacion Economica, Se cambiara a estado de Fin de Evaluacion ") = vbNo Then Exit Sub
'guardar Puntajes
Screen.MousePointer = 11
For i = 5 To flexEvaluacionTecnica.Cols - 1
        'obtiene Nombres
        'nombres = clsDGAdqui.CargaLogSelNombresProveedor(codigo)
         nResult = i Mod 2
         If nResult = 0 Then
         Else
             sCodigo = Right(flexEvaluacionTecnica.TextMatrix(0, i), 13)
             For f = 2 To flexEvaluacionTecnica.Rows - 2
                    npuntaje = flexEvaluacionTecnica.TextMatrix(f, i)
                    scodBien = flexEvaluacionTecnica.TextMatrix(f, 0)
                    'Guardar Puntaje
                    ClsNAdqui.ActualizaSelPuntajeEvaluacionEco txtSeleccionA.Text, sCodigo, scodBien, npuntaje, "2000"
             Next
                    clsDGAdqui.ActualizaEstadoProcesoSeleccion txtSeleccionA.Text, TpoLogSelEstProceso.SelEstProcesoFinEvaluacion
                    clsDGAdqui.ActualizaLogSelganador txtSeleccionA.Text
                    nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccionA.Text)
                    lblestado.Caption = clsDGAdqui.CargaLogSelEstadoDesc(nestadoProc)
                    
         End If
Next
MsgBox "Se Cambio el Estado de Manera Correcta", vbInformation, "Estado Cambiado de Manera Correcta"
Screen.MousePointer = 0
'Cambiar estado Fin de Evaluacion



End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub


Private Sub Form_Load()
Me.Height = 7275
Me.Width = 11640
Dim sAno As String
Set rs = New ADODB.Recordset
Set clsDGnral = New DLogGeneral
Set clsDGAdqui = New DLogAdquisi
Set ClsNAdqui = New NActualizaProcesoSelecLog
Set rs = clsDGnral.CargaPeriodo
Call CargaCombo(rs, cmbperiodo)
sAno = Year(gdFecSis)
ubicar_ano sAno, cmbperiodo
flexEvaluacionTecnica.BackColorBkg = -2147483643
If psTpoEval = "T" Then
    Me.Caption = "Evaluacion Tecnica"
    cmdActaBuenaPro.Visible = False
    FormateaGrilla
    cmdPuntaje.Visible = False
    cmdCerrar.Visible = False
    cmdempate.Visible = False
    
ElseIf psTpoEval = "E" Then
    cmdActaBuenaPro.Visible = False
    Me.Caption = "Evaluacion Economica"
    FormateaGrilla_Eco
    cmdPuntaje.Visible = True
    txtPuntMinimo.Visible = False
    txtPuntMaximo.Visible = False
    lblPuntMinimo.Visible = False
    lblPuntMaximo.Visible = False
    cmdCerrar.Visible = False
    cmdempate.Visible = False
ElseIf psTpoEval = "R" Then
    cmdPuntaje.Visible = False
    cmdActaBuenaPro.Visible = True
    FormateaGrilla_Cuadro
    Me.Caption = "Cuadro Comparativo de Cotizaciones"
    txtPuntMinimo.Visible = False
    txtPuntMaximo.Visible = False
    lblPuntMinimo.Visible = False
    lblPuntMaximo.Visible = False
    cmdCerrar.Visible = True
    cmdempate.Visible = True
End If
txtSeleccionA.Enabled = True
End Sub

Private Sub txtSelecciona_EmiteDatos()

If txtSeleccionA.Text = "" Then Exit Sub
'mostrar_criterios_procesos txtSeleccionA.Text
'Mostrar  Los criterios tecnicos del proceso con la configuracion de puntajes
    flexEvaluacionTecnica.Clear
    flexEvaluacionTecnica.Rows = 3
    mostrar_descripcion txtSeleccionA.Text
    
    
    If psTpoEval = "T" Then
            If Right(txtDescripcion.Text, 1) = 2 Then
                MsgBox "Este es un proceso directo y no necesita Evaluacion Tecnica", vbInformation, "Proceso Directo No Tiene Evaluacion Tecnica"
                Exit Sub
            End If
    
            If Validar_Resumen("T") = -1 Then Exit Sub
            Mostrar_Resumen_Tec txtSeleccionA.Text
    ElseIf psTpoEval = "E" Then
            If Validar_Resumen("E") = -1 Then Exit Sub
            Mostrar_Resumen_Eco txtSeleccionA.Text
    ElseIf psTpoEval = "R" Then
            If Validar_Resumen("R") = -1 Then Exit Sub
            Mostrar_Resumen_Cuadro txtSeleccionA.Text
    End If
End Sub
Sub mostrar_descripcion(nLogSelProceso As Long)
    Set rs = clsDGAdqui.CargaLogSelDescripcionProceso(nLogSelProceso)
    If rs.EOF = True Then
        txttipo.Text = ""
        txtDescripcion.Text = ""
        lblestado.Caption = ""
        Else
            txttipo.Text = UCase(rs!cTipo)
            txtDescripcion.Text = "COTIZACION Nº: " + rs!nLogSelNumeroCot + " - " + rs!cDescripcionProceso + " - TIPO PROCESO: " + rs!nLogSelDescProceso + Space(300) + Str(rs!nLogSelTipoProceso)
            If rs!nLogSelEstado = SelEstProcesoIniciado Then
                lblestado.Caption = "INICIADO"
            ElseIf rs!nLogSelEstado = SelEstProcesoEvaluacionTec Then
                lblestado.Caption = "EVALUACION TECNICA"
            ElseIf rs!nLogSelEstado = SelEstProcesoEvaluacionEco Then
                lblestado.Caption = "EVALUACION ECONOMICA"
            ElseIf rs!nLogSelEstado = SelEstProcesoFinEvaluacion Then
                lblestado.Caption = "FIN DE EVALUACION"
            ElseIf rs!nLogSelEstado = SelEstProcesoCerrado Then
                lblestado.Caption = "CERRADO"
            ElseIf rs!nLogSelEstado = SelEstProcesoCancelado Then
                lblestado.Caption = "CANCELADO"
            End If
    End If
End Sub
Sub Mostrar_Resumen_Cuadro(nLogSelProceso As Long)
    'validar
    Dim nValida As Integer
    nValida = clsDGAdqui.ValidaLogSelCalificacionTecnica(nLogSelProceso)
    If Right(txtDescripcion.Text, 1) = 2 Then
    Else
        If nValida = 0 Then
            FormateaGrilla
            MsgBox "El proceso de seleccion Nº " & txtSeleccionA.Text & " no Tiene Ingresada Calificacion Tecnica para los Proveedores  ", vbInformation, "Verifique las Evaluaciones Tecnicas "
            Exit Sub
    End If
    
    End If
    
    
    nValida = 0
    nValida = clsDGAdqui.ValidaLogSelCalificacionEconomica(nLogSelProceso)
    If nValida = 0 Then
        FormateaGrilla_Eco
            MsgBox "El proceso de seleccion Nº " & txtSeleccionA.Text & " no Tiene Ingresada las Cotizaciones de sus Proveedores ", vbInformation, "Verifique Las Cotizaciones de los Proveedores"
            Exit Sub
    End If
    
    Set rs = clsDGAdqui.CargaLogSelResumenCuadro(nLogSelProceso)
    If rs.EOF = True Then
        flexEvaluacionTecnica.Clear
        flexEvaluacionTecnica.Rows = 3
        Mostrar_Cabecera_Resumen_Cuadro
        Else
        Set flexEvaluacionTecnica.Recordset = rs
        Mostrar_Cabecera_Resumen_Cuadro
    End If
End Sub
Sub Mostrar_Resumen_Tec(nLogSelProceso As Long)
    'validar
    Dim nValida As Integer
    nValida = clsDGAdqui.ValidaLogSelCalificacionTecnica(nLogSelProceso)
    If nValida = 0 Then
        FormateaGrilla
        MsgBox "El proceso de seleccion Nº " & txtSeleccionA.Text & " no Tiene Ingresada Calificacion Tecnica para los Proveedores  ", vbInformation, "No Se Ingreso Calificacion Tecnica"
        Exit Sub
    End If
    Set rs = clsDGAdqui.CargaSeleccionPuntajes(nLogSelProceso)
    If rs.EOF = True Then
        txtPuntMinimo.Text = ""
        txtPuntMaximo.Text = ""
        Else
            txtPuntMinimo.Text = Format(rs!nPuntajeMinimo, "00.00")
            txtPuntMaximo.Text = Format(rs!nPuntajeMaximo, "00.00")
    End If
    
    Set rs = clsDGAdqui.CargaLogSelResumenEvaluacionTecnica(nLogSelProceso)
    If rs.EOF = True Then
        flexEvaluacionTecnica.Clear
        flexEvaluacionTecnica.Rows = 3
        Mostrar_Cabecera_Resumen
        Else
        Set flexEvaluacionTecnica.Recordset = rs
        Mostrar_Cabecera_Resumen
    End If
    
    
End Sub
Sub Mostrar_Resumen_Eco(nLogSelProceso As Long)
    'validar
    Dim nValida As Integer
    nValida = clsDGAdqui.ValidaLogSelCalificacionEconomica(nLogSelProceso)
    If nValida = 0 Then
        FormateaGrilla_Eco
            MsgBox "El proceso de seleccion Nº " & txtSeleccionA.Text & " no Tiene Ingresada las Cotizaciones de sus Proveedores ", vbInformation, "No se Ingreso Ninguna Cotizacion de los Proveedores"
            Exit Sub
    End If
    Set rs = clsDGAdqui.CargaLogSelResumenEvaluacionEconomica(nLogSelProceso)
    If rs.EOF = True Then
        flexEvaluacionTecnica.Clear
        flexEvaluacionTecnica.Rows = 3
        Mostrar_Cabecera_Resumen_Eco
        Else
        Set flexEvaluacionTecnica.Recordset = rs
        Mostrar_Cabecera_Resumen_Eco
    End If
End Sub
Sub Mostrar_Cabecera_Resumen_Cuadro()
'Obtener cabecera
Dim codigo  As String
Dim nombres As String
Dim result As Integer
Dim nRow As String
Dim nSuma As Double
flexEvaluacionTecnica.MergeCells = flexMergeRestrictRows
flexEvaluacionTecnica.MergeRow(0) = True
flexEvaluacionTecnica.ColWidth(0) = 1000
flexEvaluacionTecnica.ColWidth(1) = 3400
flexEvaluacionTecnica.TextMatrix(0, 0) = "Codigo"
flexEvaluacionTecnica.TextMatrix(1, 0) = "Codigo"
flexEvaluacionTecnica.TextMatrix(0, 1) = "Descripcion Bien"
flexEvaluacionTecnica.TextMatrix(1, 1) = "Descripcion Bien"
flexEvaluacionTecnica.TextMatrix(0, 2) = "Unidad"
flexEvaluacionTecnica.TextMatrix(1, 2) = "Unidad"
For i = 3 To flexEvaluacionTecnica.Cols - 1
    codigo = Right(flexEvaluacionTecnica.TextMatrix(0, i), 13)
    'obtiene Nombres
    nombres = clsDGAdqui.CargaLogSelNombresProveedor(codigo)
    flexEvaluacionTecnica.TextMatrix(0, i) = Trim(nombres)
    flexEvaluacionTecnica.ColWidth(i) = 1350
    
Next
For k = 3 To flexEvaluacionTecnica.Cols - 1 Step 4
        flexEvaluacionTecnica.TextMatrix(1, k + 0) = "Punt.Economico"
        flexEvaluacionTecnica.TextMatrix(1, k + 1) = "Punt.Tecnico"
        flexEvaluacionTecnica.TextMatrix(1, k + 2) = "Punt.Total"
        flexEvaluacionTecnica.TextMatrix(1, k + 3) = "Ind.Ganador"
        For X = 2 To flexEvaluacionTecnica.Rows - 1
            flexEvaluacionTecnica.Col = k + 2
            flexEvaluacionTecnica.Row = X
            'flexEvaluacionTecnica.CellForeColor = &HFFFF00       'HFFFFFF
            flexEvaluacionTecnica.CellBackColor = &H80000018
        Next
Next
flexEvaluacionTecnica.AddItem ""
nRow = flexEvaluacionTecnica.Rows
If flexEvaluacionTecnica.TextMatrix(2, 3) = "" Then Exit Sub
For i = 3 To flexEvaluacionTecnica.Cols - 1
        
        
            For f = 2 To nRow - 1
                
                Select Case flexEvaluacionTecnica.TextMatrix(f, i)
                Case "SI"
                    nSuma = 999
                    flexEvaluacionTecnica.Col = i
                    flexEvaluacionTecnica.Row = f
                    flexEvaluacionTecnica.CellBackColor = RGB(100, 200, 300)
                Case "NO"
                    nSuma = 999
                Case Else
                nSuma = nSuma + IIf(flexEvaluacionTecnica.TextMatrix(f, i) = "", 0, flexEvaluacionTecnica.TextMatrix(f, i))
                
                End Select
                
            Next
            If nSuma <> 999 Then
            'FlexEvaluacionTecnica.TextMatrix(nRow, i) = Format(nSuma, "##.00")
            End If
                
            If nSuma = 0 And nSuma <> 999 Then
                flexEvaluacionTecnica.TextMatrix(f - 1, i) = "DESCALIFICADO"
            End If
                nSuma = 0

Next
            'For i = 0 To FlexEvaluacionTecnica.Cols - 1
            '    FlexEvaluacionTecnica.Col = i
            '    FlexEvaluacionTecnica.Row = nRow
            '    'flexEvaluacionTecnica.CellForeColor = &HFFFF00       'HFFFFFF
            '    FlexEvaluacionTecnica.CellBackColor = &H80000018
            'Next
End Sub


Sub Mostrar_Cabecera_Resumen()
'Obtener cabecera
Dim codigo  As String
Dim nombres As String
Dim result As Integer
Dim nRow As String
Dim nSuma As Double
flexEvaluacionTecnica.MergeCells = flexMergeRestrictRows
flexEvaluacionTecnica.MergeRow(0) = True
flexEvaluacionTecnica.ColWidth(0) = 3400
flexEvaluacionTecnica.TextMatrix(0, 0) = "Criterio Tecnico"
flexEvaluacionTecnica.TextMatrix(1, 0) = "Criterio Tecnico"

flexEvaluacionTecnica.TextMatrix(0, 1) = "Puntaje Maximo"
flexEvaluacionTecnica.TextMatrix(1, 1) = "Puntaje Maximo"

For i = 2 To flexEvaluacionTecnica.Cols - 1
    codigo = Right(flexEvaluacionTecnica.TextMatrix(0, i), 13)
    'obtiene Nombres
    nombres = clsDGAdqui.CargaLogSelNombresProveedor(codigo)
    flexEvaluacionTecnica.TextMatrix(0, i) = Trim(nombres)
    
    flexEvaluacionTecnica.ColWidth(i) = 1350
    result = i Mod 2
    If result <> 0 Then
    flexEvaluacionTecnica.TextMatrix(1, i) = "Observacion"
    Else
    flexEvaluacionTecnica.TextMatrix(1, i) = "Puntaje"
    End If
Next
flexEvaluacionTecnica.AddItem "TOTAL"
nRow = flexEvaluacionTecnica.Rows - 1

For i = 0 To flexEvaluacionTecnica.Cols - 1
            flexEvaluacionTecnica.Col = i
            flexEvaluacionTecnica.Row = nRow
            'flexEvaluacionTecnica.CellForeColor = &HFFFF00       'HFFFFFF
            flexEvaluacionTecnica.CellBackColor = RGB(100, 200, 300)
Next

For i = 1 To flexEvaluacionTecnica.Cols - 1
     result = i Mod 2
        If result <> 0 Then
        Else
        For f = 2 To nRow - 1
        nSuma = nSuma + IIf(flexEvaluacionTecnica.TextMatrix(f, i) = "", 0, flexEvaluacionTecnica.TextMatrix(f, i))

        Next
        flexEvaluacionTecnica.TextMatrix(nRow, i) = Format(nSuma, "##.#0")
        If nSuma < Val(txtPuntMinimo.Text) Then
            flexEvaluacionTecnica.TextMatrix(nRow, i + 1) = "DESCALIFICADO"
            
        End If
        nSuma = 0
    End If
Next

End Sub
Sub Mostrar_Cabecera_Resumen_Eco()
'Obtener cabecera
    Dim codigo  As String
    Dim nombres As String
    Dim result As Integer
    Dim nRow As String
    Dim nSuma As Double
    flexEvaluacionTecnica.MergeCells = flexMergeRestrictRows
    flexEvaluacionTecnica.MergeRow(0) = True
    flexEvaluacionTecnica.ColWidth(0) = 1000
    flexEvaluacionTecnica.ColWidth(1) = 3400
    flexEvaluacionTecnica.TextMatrix(0, 0) = "Codigo"
    flexEvaluacionTecnica.TextMatrix(1, 0) = "Codigo"
    flexEvaluacionTecnica.TextMatrix(0, 1) = "Descripcion Bien"
    flexEvaluacionTecnica.TextMatrix(1, 1) = "Descripcion Bien"
    flexEvaluacionTecnica.TextMatrix(0, 2) = "Unidad"
    flexEvaluacionTecnica.TextMatrix(1, 2) = "Unidad"
    flexEvaluacionTecnica.TextMatrix(0, 3) = "Precio Referencial"
    flexEvaluacionTecnica.TextMatrix(1, 3) = "Precio Referencial"
    
    flexEvaluacionTecnica.MergeRow(0) = True
    flexEvaluacionTecnica.MergeCol(0) = True
    flexEvaluacionTecnica.MergeCol(1) = True
    flexEvaluacionTecnica.MergeCol(2) = True
    flexEvaluacionTecnica.MergeCol(3) = True
    For i = 4 To flexEvaluacionTecnica.Cols - 1
        codigo = Right(flexEvaluacionTecnica.TextMatrix(0, i), 13)
        'obtiene Nombres
        nombres = clsDGAdqui.CargaLogSelNombresProveedor(codigo)
        flexEvaluacionTecnica.TextMatrix(0, i) = Trim(nombres) + Space(100) + codigo
        flexEvaluacionTecnica.ColWidth(i) = 1350
        result = i Mod 2
            If result = 0 Then
                flexEvaluacionTecnica.TextMatrix(1, i) = "Precio Cotizacion"
            Else
                flexEvaluacionTecnica.TextMatrix(1, i) = "Puntaje"
            End If
    Next
    flexEvaluacionTecnica.AddItem "TOTAL"
    nRow = flexEvaluacionTecnica.Rows - 1
    For i = 3 To flexEvaluacionTecnica.Cols - 1
        For f = 2 To nRow - 1
        nSuma = nSuma + flexEvaluacionTecnica.TextMatrix(f, i)
        Next
        flexEvaluacionTecnica.TextMatrix(nRow, i) = Format(nSuma, "##.00")
        If nSuma = 0 Then
            flexEvaluacionTecnica.TextMatrix(nRow, i) = "DESCALIFICADO"
        End If
        
        nSuma = 0
    Next
        For i = 0 To flexEvaluacionTecnica.Cols - 1
         flexEvaluacionTecnica.Col = i
         flexEvaluacionTecnica.Row = nRow
         'flexEvaluacionTecnica.CellForeColor = &HFFFF00       'HFFFFFF
         flexEvaluacionTecnica.CellBackColor = RGB(100, 200, 300)
        Next
End Sub


Sub FormateaGrilla()
    flexEvaluacionTecnica.Clear
    flexEvaluacionTecnica.Cols = 4
    flexEvaluacionTecnica.MergeCells = flexMergeRestrictRows
    flexEvaluacionTecnica.ColWidth(0) = 3400
    flexEvaluacionTecnica.TextMatrix(0, 0) = "Criterio Tecnico"
    flexEvaluacionTecnica.TextMatrix(1, 0) = "Criterio Tecnico"
    flexEvaluacionTecnica.TextMatrix(0, 1) = "Criterio Tecnico"
    flexEvaluacionTecnica.TextMatrix(1, 1) = "Criterio Tecnico"
    
    flexEvaluacionTecnica.TextMatrix(0, 2) = "Proveedor"
    flexEvaluacionTecnica.TextMatrix(0, 3) = "Proveedor"
    flexEvaluacionTecnica.TextMatrix(1, 2) = "Puntaje"
    flexEvaluacionTecnica.TextMatrix(1, 3) = "Observacion"
    flexEvaluacionTecnica.MergeRow(0) = True
    flexEvaluacionTecnica.MergeCol(0) = True
    flexEvaluacionTecnica.MergeCol(1) = True
End Sub
Sub FormateaGrilla_Eco()
    flexEvaluacionTecnica.Clear
    flexEvaluacionTecnica.Cols = 6
    flexEvaluacionTecnica.MergeCells = flexMergeRestrictRows
    flexEvaluacionTecnica.ColWidth(0) = 1000
    flexEvaluacionTecnica.ColWidth(1) = 3400
    flexEvaluacionTecnica.TextMatrix(0, 0) = "Codigo"
    flexEvaluacionTecnica.TextMatrix(1, 0) = "Codigo"
    flexEvaluacionTecnica.TextMatrix(0, 1) = "Descripcion Bien"
    flexEvaluacionTecnica.TextMatrix(1, 1) = "Descripcion Bien"
    flexEvaluacionTecnica.TextMatrix(0, 2) = "Unidad"
    flexEvaluacionTecnica.TextMatrix(1, 2) = "Unidad"
    flexEvaluacionTecnica.TextMatrix(0, 3) = "Precio Ref."
    flexEvaluacionTecnica.TextMatrix(1, 3) = "Precio Ref."
    flexEvaluacionTecnica.TextMatrix(0, 4) = "Proveedor"
    flexEvaluacionTecnica.TextMatrix(0, 5) = "Proveedor"
    flexEvaluacionTecnica.TextMatrix(1, 4) = "Precio Cotizacion"
    flexEvaluacionTecnica.TextMatrix(1, 5) = "Puntaje"
    flexEvaluacionTecnica.MergeRow(0) = True
    flexEvaluacionTecnica.MergeCol(0) = True
    flexEvaluacionTecnica.MergeCol(1) = True
    flexEvaluacionTecnica.MergeCol(2) = True
End Sub
Sub FormateaGrilla_Cuadro()
    flexEvaluacionTecnica.Clear
    flexEvaluacionTecnica.Cols = 6
    flexEvaluacionTecnica.MergeCells = flexMergeRestrictRows
    flexEvaluacionTecnica.ColWidth(0) = 1000
    flexEvaluacionTecnica.ColWidth(1) = 3400
    flexEvaluacionTecnica.TextMatrix(0, 0) = "Codigo"
    flexEvaluacionTecnica.TextMatrix(1, 0) = "Codigo"
    flexEvaluacionTecnica.TextMatrix(0, 1) = "Descripcion Bien"
    flexEvaluacionTecnica.TextMatrix(1, 1) = "Descripcion Bien"
    flexEvaluacionTecnica.TextMatrix(0, 2) = "Unidad"
    flexEvaluacionTecnica.TextMatrix(1, 2) = "Unidad"
    flexEvaluacionTecnica.TextMatrix(0, 3) = "Proveedor"
    flexEvaluacionTecnica.TextMatrix(0, 4) = "Proveedor"
    flexEvaluacionTecnica.TextMatrix(0, 5) = "Proveedor"
    flexEvaluacionTecnica.TextMatrix(1, 3) = "Punt.Economico"
    flexEvaluacionTecnica.TextMatrix(1, 4) = "Punt.Tecnico"
    flexEvaluacionTecnica.TextMatrix(1, 5) = "Punt.Total"
    flexEvaluacionTecnica.MergeRow(0) = True
    flexEvaluacionTecnica.MergeCol(0) = True
    flexEvaluacionTecnica.MergeCol(1) = True
    flexEvaluacionTecnica.MergeCol(2) = True
End Sub

Public Sub Inicio(ByVal psTipoEval As String, ByVal psFormTpo As String, Optional ByVal psSeleccionNro As Long = 0)
psTpoEval = psTipoEval
psFrmTpo = psFormTpo
psReqNro = psSeleccionNro
Me.Show
End Sub

Function Validar_Resumen(Tipo As String) As Integer
Dim nNumProvedores As Integer
Dim nNumProvedoresEvalTec As Integer
Dim nNumProvedoresEvalCot As Integer
nNumProvedores = clsDGAdqui.CuentaLogSelProveedor(txtSeleccionA.Text)

Select Case Tipo
Case "T"
     nNumProvedoresEvalTec = clsDGAdqui.CuentaLogSelProveedorEvalTec(txtSeleccionA.Text)
     'Todos Los proveedores tengan Evalaucion tecnica
     If nNumProvedores - nNumProvedoresEvalTec = 0 Then
        'Ok
         Else
          If Right(txtDescripcion.Text, 1) = 2 Then
                
                Else
                MsgBox "Existen " + Str(nNumProvedores - nNumProvedoresEvalTec) + " Proveedores Que No Tienen Evaluacion Tecnica ", vbInformation, "Pendiente Evaluacion Tecnica de Proveedor"
                Validar_Resumen = -1
                Exit Function
          End If
      End If
Case "E"
     nNumProvedoresEvalTec = clsDGAdqui.CuentaLogSelProveedorEvalTec(txtSeleccionA.Text)
     'Todos Los proveedores tengan Evalaucion tecnica
     If nNumProvedores - nNumProvedoresEvalTec = 0 Then
        'Ok
         Else
         
         If Right(txtDescripcion.Text, 1) = 2 Then
            Else
                 MsgBox "Existen " + Str(nNumProvedores - nNumProvedoresEvalTec) + " Proveedores Que No Tienen Evaluacion Tecnica ", vbInformation, "Pendiente Evaluacion Tecnica de Proveedor"
                 Validar_Resumen = -1
                 Exit Function
         End If
      End If

      nNumProvedoresEvalCot = clsDGAdqui.CuentaLogSelProveedorEvalEco(txtSeleccionA.Text)
      If nNumProvedores - nNumProvedoresEvalCot = 0 Then
        'Ok
        Else
         MsgBox "Existen " + Str(nNumProvedores - nNumProvedoresEvalCot) + " Proveedores Que No Tienen Su Cotizacion Ingresada ", vbInformation, "Pendiente Ingreso de Cotizacion de Proveedor"
         Validar_Resumen = -1
         Exit Function
      End If
Case "R"
       'Verifica  estado
       nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccionA.Text)
       
       If nestadoProc <> SelEstProcesoFinEvaluacion Then
          If nestadoProc = SelEstProcesoCerrado Then
          
          Else
                MsgBox "El Proceso de Seleccion Nº " + txtSeleccionA.Text + " No Tiene el estado Fin de Evaluacion ", vbInformation, "No Tiene el estado de Fin de Evaluacion"
                Validar_Resumen = -1
                Exit Function
          End If
       End If
       nNumProvedoresEvalTec = clsDGAdqui.CuentaLogSelProveedorEvalTec(txtSeleccionA.Text)
     'Todos Los proveedores tengan Evalaucion tecnica
     If nNumProvedores - nNumProvedoresEvalTec = 0 Then
        'Ok
         Else
         
         If Right(txtDescripcion.Text, 1) = 2 Then
         Else
             MsgBox "Existen " + Str(nNumProvedores - nNumProvedoresEvalTec) + " Proveedores Que No Tienen Evaluacion Tecnica ", vbInformation, "Pendiente Evaluacion Tecnica de Proveedor"
             Validar_Resumen = -1
             Exit Function
         End If
         
      End If
    'Todos Los Proveedores Tienen Cotizacion Ingesada Todos
      nNumProvedoresEvalCot = clsDGAdqui.CuentaLogSelProveedorEvalEco(txtSeleccionA.Text)
      If nNumProvedores - nNumProvedoresEvalCot = 0 Then
        'Ok
        Else
         MsgBox "Existen " + Str(nNumProvedores - nNumProvedoresEvalCot) + " Proveedores Que No Tienen Su Cotizacion Ingresada ", vbInformation, "Pendiente Ingreso de Cotizacion de Proveedor"
        Validar_Resumen = -1
        Exit Function
      End If
       

End Select


End Function





Private Sub CuadroExcel(plHoja1 As Excel.Worksheet, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, Optional lbLineasVert As Boolean = False)
Dim i, j As Integer

For i = X1 To X2
    plHoja1.Range(plHoja1.Cells(Y1, i), plHoja1.Cells(Y1, i)).Borders(xlEdgeTop).LineStyle = xlContinuous
    plHoja1.Range(plHoja1.Cells(Y2, i), plHoja1.Cells(Y2, i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Next i
If lbLineasVert = False Then
    For i = X1 To X2
        For j = Y1 To Y2
            plHoja1.Range(plHoja1.Cells(j, i), plHoja1.Cells(j, i)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Next j
    Next i
End If
If lbLineasVert Then
    For j = Y1 To Y2
        plHoja1.Range(plHoja1.Cells(j, X1), plHoja1.Cells(j, X1)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Next j
End If

For j = Y1 To Y2
    plHoja1.Range(plHoja1.Cells(j, X2), plHoja1.Cells(j, X2)).Borders(xlEdgeRight).LineStyle = xlContinuous
Next j
End Sub
Sub Imprime_comite(b As Integer)
 Dim nCambio As Boolean
    i = b
    n = 3
    Do While Not rs.EOF
        If n > 11 Or nCambio = True Then
                If nCambio = False Then
                    n = 3
                End If
                nCambio = True
                xlHoja1.Cells(i + 17, n).value = "__________________________________"
                xlHoja1.Cells(i + 18, n).value = "                  " + rs!cDescripcion
                xlHoja1.Cells(i + 19, n).value = rs!cPersNombre
                xlHoja1.Cells(i + 20, n).value = "                  DNI  : " + rs!cPersIDnro
            
        ElseIf nCambio = False Then
                xlHoja1.Cells(i + 11, n).value = "__________________________________"
                xlHoja1.Cells(i + 12, n).value = "                  " + rs!cDescripcion
                xlHoja1.Cells(i + 13, n).value = rs!cPersNombre
                xlHoja1.Cells(i + 14, n).value = "                  DNI  : " + rs!cPersIDnro
    End If
                rs.MoveNext
                n = n + 4
    Loop
End Sub

Function reporte_comparativo() As Integer

lsCadAnt = ""
Dim Abc(22) As String
    For n = 0 To flexEvaluacionTecnica.Cols - 1
        flexEvaluacionTecnica.Col = n
        lnIni = 0
        For i = 0 To flexEvaluacionTecnica.Rows - 1
            If i = 1 And (n < 3) Then
            Else
                   flexEvaluacionTecnica.Row = i
                    If i = 0 And (n > 2) Then
                        If lsCadAnt = flexEvaluacionTecnica.TextMatrix(i, n) Then
                        Else
                            xlHoja1.Cells(i + 7, n + 1).value = flexEvaluacionTecnica.Text
                            lsCadAnt = xlHoja1.Cells(i + 7, n + 1).value
                        End If
                    Else
                        xlHoja1.Cells(i + 7, n + 1).value = flexEvaluacionTecnica.Text
                    End If
            End If
        Next
    Next

    xlHoja1.Range("A" & 7 & ":A" & 8 & "").Merge
    xlHoja1.Range("A" & 7 & ":A" & 8 & "").VerticalAlignment = xlCenter
    xlHoja1.Range("A" & 7 & ":A" & 8 & "").VerticalAlignment = xlCenter
    xlHoja1.Range("B" & 7 & ":B" & 8 & "").Merge
    xlHoja1.Range("B" & 7 & ":B" & 8 & "").VerticalAlignment = xlCenter
    xlHoja1.Range("B" & 7 & ":B" & 8 & "").VerticalAlignment = xlCenter
    xlHoja1.Range("C" & 7 & ":C" & 8 & "").Merge
    xlHoja1.Range("C" & 7 & ":C" & 8 & "").HorizontalAlignment = xlCenter
    xlHoja1.Range("C" & 7 & ":C" & 8 & "").VerticalAlignment = xlCenter
 
    Abc(0) = "D7": Abc(1) = "G7": Abc(2) = "H7"
    Abc(3) = "K7": Abc(4) = "L7": Abc(5) = "O7": Abc(6) = "P7": Abc(7) = "S7"
    Abc(8) = "T7": Abc(9) = "W7": Abc(10) = "X7": Abc(11) = "AA7": Abc(12) = "AB7": Abc(13) = "AE7": Abc(14) = "AF7"
    Abc(15) = "AI7": Abc(16) = "AJ7": Abc(17) = "AM7": Abc(18) = "AN7": Abc(19) = "AQ7": Abc(20) = "AR7"
    Abc(21) = "AU7"
    xlHoja1.Range("A1").ColumnWidth = 12
    xlHoja1.Range("B1").ColumnWidth = 40
    X = 0
    
Do While X <= (flexEvaluacionTecnica.Cols - 3 - 1) / 2
   With xlHoja1.Range("" & Abc(X) & "" & ":" & "" & Abc(X + 1) & "")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = True
    End With
    X = X + 2
Loop
reporte_comparativo = i

End Function

Function reporte_economico() As Integer
lsCadAnt = ""
Dim Abc(22) As String
    For n = 0 To flexEvaluacionTecnica.Cols - 1
        flexEvaluacionTecnica.Col = n
        lnIni = 0
        For i = 0 To flexEvaluacionTecnica.Rows - 1
            
            If n >= 4 And i = 0 Then
                flexEvaluacionTecnica.TextMatrix(i, n) = Left(flexEvaluacionTecnica.TextMatrix(i, n), 30)
            End If
            
            If i = 1 And (n < 4) Then
            Else
                   flexEvaluacionTecnica.Row = i
                    If i = 0 And (n > 2) Then
                        If lsCadAnt = flexEvaluacionTecnica.TextMatrix(i, n) Then
                        Else
                            xlHoja1.Cells(i + 7, n + 1).value = flexEvaluacionTecnica.Text
                            lsCadAnt = xlHoja1.Cells(i + 7, n + 1).value
                        End If
                    Else
                        xlHoja1.Cells(i + 7, n + 1).value = flexEvaluacionTecnica.Text
                    End If
            End If
        Next
    Next
    xlHoja1.Range("A" & 7 & ":A" & 8 & "").Merge
    xlHoja1.Range("A" & 7 & ":A" & 8 & "").VerticalAlignment = xlCenter
    xlHoja1.Range("A" & 7 & ":A" & 8 & "").VerticalAlignment = xlCenter
    xlHoja1.Range("B" & 7 & ":B" & 8 & "").Merge
    xlHoja1.Range("B" & 7 & ":B" & 8 & "").VerticalAlignment = xlCenter
    xlHoja1.Range("B" & 7 & ":B" & 8 & "").VerticalAlignment = xlCenter
    xlHoja1.Range("C" & 7 & ":C" & 8 & "").Merge
    xlHoja1.Range("C" & 7 & ":C" & 8 & "").HorizontalAlignment = xlCenter
    xlHoja1.Range("C" & 7 & ":C" & 8 & "").VerticalAlignment = xlCenter
    xlHoja1.Range("D" & 7 & ":D" & 8 & "").Merge
    xlHoja1.Range("D" & 7 & ":D" & 8 & "").HorizontalAlignment = xlCenter
    xlHoja1.Range("D" & 7 & ":D" & 8 & "").VerticalAlignment = xlCenter
    
   
'"GH"
'"I:J"
'"K:L"
'"M:N"
'"O:P"
'"Q:R"
'"S:T"
'"U:V"
    Abc(0) = "E7:F7": Abc(1) = "G7:H7": Abc(2) = "I7:J7"
    Abc(3) = "K7:L7": Abc(4) = "M7:N7": Abc(5) = "O7:P7": Abc(6) = "Q7:R7"
    Abc(7) = "S7:T7": Abc(8) = "U7:V7"
    
    xlHoja1.Range("A1").ColumnWidth = 12
    xlHoja1.Range("B1").ColumnWidth = 40
    X = 0
    
    'Do While X < 10
    'With xlHoja1.Range("" & Abc(X) & "" & ":" & "" & Abc(X + 1) & "")
    'With xlHoja1.Range(Abc(X))
    '    .Merge
    '    .HorizontalAlignment = xlCenter
    '    .VerticalAlignment = xlBottom
    '    .WrapText = False
    '    .Orientation = 0
    '    .AddIndent = False
    '  .ShrinkToFit = False
    '    .MergeCells = True
    'End With
    'X = X + 1
    'Loop
    reporte_economico = i
    End Function

Function reporte_Tecnico() As Integer
lsCadAnt = ""
Dim Abc(22) As String
    For n = 0 To flexEvaluacionTecnica.Cols - 1
        flexEvaluacionTecnica.Col = n
        lnIni = 0
        For i = 0 To flexEvaluacionTecnica.Rows - 1
            
            If n >= 2 And i = 0 Then
                flexEvaluacionTecnica.TextMatrix(i, n) = Left(flexEvaluacionTecnica.TextMatrix(i, n), 30)
            End If
            
            If i = 1 And (n < 2) Then
            Else
                   flexEvaluacionTecnica.Row = i
                    If i = 0 And (n > 1) Then
                        If lsCadAnt = flexEvaluacionTecnica.TextMatrix(i, n) Then
                        Else
                            xlHoja1.Cells(i + 7, n + 1).value = flexEvaluacionTecnica.Text
                            lsCadAnt = xlHoja1.Cells(i + 7, n + 1).value
                        End If
                    Else
                        xlHoja1.Cells(i + 7, n + 1).value = flexEvaluacionTecnica.Text
                    End If
            End If
        Next
    Next
    xlHoja1.Range("A" & 7 & ":A" & 8 & "").Merge
    xlHoja1.Range("A" & 7 & ":A" & 8 & "").VerticalAlignment = xlCenter
    xlHoja1.Range("A" & 7 & ":A" & 8 & "").VerticalAlignment = xlCenter
    xlHoja1.Range("B" & 7 & ":B" & 8 & "").Merge
    xlHoja1.Range("B" & 7 & ":B" & 8 & "").VerticalAlignment = xlCenter
    xlHoja1.Range("B" & 7 & ":B" & 8 & "").VerticalAlignment = xlCenter
    
    
   
    
    xlHoja1.Range("A1").ColumnWidth = 40
    
    reporte_Tecnico = i
    End Function



