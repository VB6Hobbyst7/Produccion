VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogSelCancelacionProceso 
   Caption         =   "Cancelacion de Proceso de Seleccion"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   Icon            =   "frmLogSelCancelacionProceso.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7245
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5160
      TabIndex        =   21
      Top             =   6810
      Width           =   1455
   End
   Begin VB.CommandButton cmdAnulacion 
      Caption         =   "Anular"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   6810
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Comite"
      TabPicture(0)   =   "frmLogSelCancelacionProceso.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FlexComite"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Proveedores"
      TabPicture(1)   =   "frmLogSelCancelacionProceso.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FlexProvedores"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Criterios Tecnicos"
      TabPicture(2)   =   "frmLogSelCancelacionProceso.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FlexCriterios"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Listado de Bienes"
      TabPicture(3)   =   "frmLogSelCancelacionProceso.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fgeBienesConfig"
      Tab(3).ControlCount=   1
      Begin Sicmact.FlexEdit FlexCriterios 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   16
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   6376
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Codigo-Descripcion-Ultima Actualizacion"
         EncabezadosAnchos=   "550-1200-3500-3000"
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
         ColumnasAEditar =   "X-1-X-X"
         ListaControles  =   "0-1-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-C"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         Enabled         =   0   'False
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   555
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit fgeBienesConfig 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6800
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   "Item-Código-Descripción-Unidad-ValorUnidad-Cantidad-Precio Ref-Sub Total-Participa"
         EncabezadosAnchos=   "450-1200-3500-700-0-1000-1000-1000-1000"
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
         ColumnasAEditar =   "X-1-X-X-X-5-6-X-8"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-0-0-0-0-0-4"
         EncabezadosAlineacion=   "R-L-L-L-R-R-C-C-C"
         FormatosEdit    =   "0-0-0-0-3-3-2-2-0"
         CantEntero      =   10
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         Enabled         =   0   'False
         lbFlexDuplicados=   0   'False
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit FlexProvedores 
         Height          =   3855
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   6800
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Codigo-Nombres-Ultima Actualizacion"
         EncabezadosAnchos=   "550-1500-3500-3000"
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
         ColumnasAEditar =   "X-1-X-X"
         ListaControles  =   "0-1-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         Enabled         =   0   'False
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbPuntero       =   -1  'True
         ColWidth0       =   555
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit FlexComite 
         Height          =   3720
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6562
         Cols0           =   4
         HighLight       =   1
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Nombre-Cargo"
         EncabezadosAnchos=   "350-1800-4000-3000"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-3"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   240
         CellBackColor   =   -2147483624
      End
   End
   Begin VB.ComboBox cmbMotAnulacion 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8790
      TabIndex        =   10
      Top             =   6810
      Width           =   1455
   End
   Begin VB.CommandButton CmdAnular 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   6810
      Width           =   1455
   End
   Begin VB.Frame s 
      Caption         =   "Proceso de  Seleccion"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   10095
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
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   8775
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
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   5895
      End
      Begin Sicmact.TxtBuscar txtSeleccionA 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
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
      Begin VB.Shape Shape2 
         BackColor       =   &H00C00000&
         BorderStyle     =   4  'Dash-Dot
         FillColor       =   &H8000000D&
         Height          =   400
         Left            =   7320
         Top             =   480
         Width           =   2655
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
         Left            =   7440
         TabIndex        =   14
         Top             =   600
         Width           =   660
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
         Left            =   7320
         TabIndex        =   13
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label Label5 
         Caption         =   "Numero"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   315
      End
   End
   Begin VB.ComboBox cmbperiodo 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblmotanulacion 
      AutoSize        =   -1  'True
      Caption         =   "Mot Anulacion"
      Height          =   195
      Left            =   1560
      TabIndex        =   12
      Top             =   6900
      Width           =   1020
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
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmLogSelCancelacionProceso"
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
txtdescripcion.Text = ""

Me.txtSeleccionA.rs = clsDGAdqui.LogSeleccionLista(cmbperiodo.Text)
End Sub


Private Sub cmdexportar_Click()
Dim i As Long
Dim n As Long
Dim lsArchivoN As String
Dim lbLibroOpen As Boolean
Dim lsCadAnt As String
Dim lnIni As Integer
Dim J As Integer
On Error Resume Next
lsArchivoN = App.path & "\seleccion.xls"
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

xlHoja1.Cells(2, 1).value = "Periodo"
xlHoja1.Cells(2, 2).value = cmbperiodo.Text
xlHoja1.Cells(3, 1).value = "Proceso de Seleccion"
xlHoja1.Cells(3, 2).value = txtSeleccionA.Text
xlHoja1.Cells(4, 1).value = "Tipo"
xlHoja1.Cells(4, 2).value = txttipo.Text
xlHoja1.Cells(5, 1).value = "Descripcion Proceso "
xlHoja1.Cells(5, 2).value = txtdescripcion.Text
For n = 0 To flexEvaluacionTecnica.Cols - 1
    flexEvaluacionTecnica.Col = n
    lnIni = 0
    For i = 0 To flexEvaluacionTecnica.Rows - 1
            flexEvaluacionTecnica.Row = i
            xlHoja1.Cells(i + 7, n + 1).value = flexEvaluacionTecnica.Text
    Next
Next
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
Dim nvalida As Integer
'validar estado
'Validar Calificacion Tecnica este Ingresada
nvalida = clsDGAdqui.ValidaLogSelCalificacionTecnica(txtSeleccionA.Text)
If nvalida = 0 Then
   MsgBox "Antes de Grabar Los Puntajes , debe Ingresar las Calificaciones tecnicas Respectivas", vbInformation, "Ingrese las Calificaciones Tecnicas"
   Exit Sub
End If
If txtSeleccionA.Text = "" Then Exit Sub
nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccion.Text)
If nestadoProc = SelEstProcesoCerrado Or nestadoProc = SelEstProcesoCancelado Then
                 MsgBox "No se puede Modificar,El Procesos de Seleccion " + txtSeleccion.Text + " Tiene un estado diferente al de  Iniciado,Evaluacion Tecnica , Evaluacion Economica o Fin de Evaluacion", vbInformation, "Estado del proceso" + txtSeleccion.Text + " ya esta Cerrdao o Cancelado"
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
                    clsDGAdqui.ActualizaEstadoProcesoSeleccion txtSeleccionA.Text, TpoLogSelEstProceso.SelEstProcesoFinEvaluacion
                    nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccionA.Text)
                    If nestadoProc = SelEstProcesoIniciado Then
                        lblestado.Caption = "INICIADO"
                    ElseIf nestadoProc = SelEstProcesoEvaluacionTec Then
                        lblestado.Caption = "EVALUACION TECNICA"
                    ElseIf nestadoProc = SelEstProcesoEvaluacionEco Then
                        lblestado.Caption = "EVALUACION ECONOMICA"
                    ElseIf nestadoProc = SelEstProcesoFinEvaluacion Then
                        lblestado.Caption = "FIN DE EVALUACION"
                    ElseIf nestadoProc = SelEstProcesoCerrado Then
                        lblestado.Caption = "CERRADO"
                    ElseIf nestadoProc = SelEstProcesoCancelado Then
                        lblestado.Caption = "ANULADO"
                    End If
                    
                    
             Next
         End If
Next
MsgBox "Se Grabaron los Puntajes de Manera Correcta", vbInformation, "Puntajes Grabados Correctamente"
Screen.MousePointer = 0
'Cambiar estado Fin de Evaluacion



End Sub



Private Sub cmdAnulacion_Click()
If txtSeleccionA.Text = "" Then Exit Sub
cmdAnulacion.Enabled = False
lblmotanulacion.Visible = True
cmbMotAnulacion.Visible = True
cmbMotAnulacion.Enabled = True
cmdCancelar.Visible = True
CmdAnular.Visible = True
End Sub

Private Sub CmdAnular_Click()
        Dim nestadoProc  As Integer
        Dim nMovCancel As Integer
        If txtSeleccionA.Text = "" Then
           MsgBox "Antes debe  Seleccionar un numero de Proceso", vbInformation, "Seleccione Numero de Seleccion"
           txtSeleccionA.SetFocus
           Exit Sub
        End If
        If cmbMotAnulacion.Text = "" Then
            MsgBox "Antes debe  Seleccionar un Motivo de Cancelacion", vbInformation, "Seleccione un Motivo de Cancelacion"
            txtSeleccionA.SetFocus
            Exit Sub
        End If
        
        
        'Validar
        
        nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccionA.Text)
        
        If nestadoProc = SelEstProcesoCancelado Or nestadoProc = SelEstProcesoCerrado Then
            MsgBox "No Se Puede anular este Proceso este Tiene un estado de Cerrado", vbInformation, "No Se Puede Anular  el Proceso se encuentra Cerrado "
            Exit Sub
        End If
        If nestadoProc = SelEstProcesoCancelado Or nestadoProc = SelEstProcesoCerrado Then
            MsgBox "No Se Puede anular este Proceso este Tiene un estado Cerrado", vbInformation, "No Se Puede Anular el Porceso ya esta Cerrado"
            Exit Sub
        End If
        sactualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                                        cmdAnulacion.Enabled = False
                    
                    
        
        
        If MsgBox("Desea Anular el proceso de Seleccion " + txtSeleccionA.Text + " este se encuentra con estado de " + lblestado.Caption, vbQuestion + vbYesNo, "Anular Proceso de Seleccion") = vbYes Then
                    clsDGAdqui.ActualizaEstadoProcesoSeleccion txtSeleccionA.Text, TpoLogSelEstProceso.SelEstProcesoCancelado
                    'Insertar Motivo Anulacion
                    clsDGAdqui.InsertaMotAnulacion txtSeleccionA.Text, Right(cmbMotAnulacion.Text, 1), sactualiza
                    nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(txtSeleccionA.Text)
                    
                    cmdAnulacion.Enabled = False
                    lblmotanulacion.Visible = True
                    cmbMotAnulacion.Visible = True
                    cmdCancelar.Visible = False
                    CmdAnular.Visible = False
                    
                    If nestadoProc = SelEstProcesoIniciado Then
                        lblestado.Caption = "INICIADO"
                    ElseIf nestadoProc = SelEstProcesoEvaluacionTec Then
                        lblestado.Caption = "EVALUACION TECNICA"
                    ElseIf nestadoProc = SelEstProcesoEvaluacionEco Then
                        lblestado.Caption = "EVALUACION ECONOMICA"
                    ElseIf nestadoProc = SelEstProcesoFinEvaluacion Then
                        lblestado.Caption = "FIN DE EVALUACION"
                    ElseIf nestadoProc = SelEstProcesoCerrado Then
                        lblestado.Caption = "CERRADO"
                    ElseIf nestadoProc = SelEstProcesoCancelado Then
                        lblestado.Caption = "ANULADO"
                        'Obtener el Motivo por el que fue Anulado ç
                        nMovCancel = clsDGAdqui.CargaLogSelMotCancel(txtSeleccionA.Text)
                        ubicar nMovCancel, cmbMotAnulacion
                        lblmotanulacion.Enabled = False
                        cmbMotAnulacion.Enabled = False
                    End If
         Exit Sub
        End If

End Sub

Private Sub cmdCancelar_Click()
cmdAnulacion.Enabled = True
lblmotanulacion.Visible = False
cmbMotAnulacion.Visible = False
cmdCancelar.Visible = False
CmdAnular.Visible = False

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Load()
Me.Height = 7680
Me.Width = 11640
Dim sAno As String
Set rs = New ADODB.Recordset
Set clsDGnral = New DLogGeneral
Set clsDGAdqui = New DLogAdquisi
Set ClsNAdqui = New NActualizaProcesoSelecLog
'Carga Periodo
Set rs = clsDGnral.CargaPeriodo
Call CargaCombo(rs, cmbperiodo)
Set rs = clsDGAdqui.CargaMotRechazos
Call CargaCombo(rs, cmbMotAnulacion)
sAno = Year(gdFecSis)
ubicar_ano sAno, cmbperiodo
fgeBienesConfig.ListaControles = "0-1-0-0-0-0-0-0-0"
fgeBienesConfig.EncabezadosAnchos = "450 - 1400 - 3500 - 700 - 0 - 1000 - 1000 - 1200 - 0"

cmdAnulacion.Enabled = True
lblmotanulacion.Visible = False
cmbMotAnulacion.Visible = False
cmdCancelar.Visible = False
CmdAnular.Visible = False

End Sub


Private Sub txtSeleccionA_EmiteDatos()

If txtSeleccionA.Text = "" Then Exit Sub
'mostrar_criterios_procesos txtSeleccionA.Text
'Mostrar  Los criterios tecnicos del proceso con la configuracion de puntajes
mostrar_descripcion txtSeleccionA.Text
mostrar_detalles txtSeleccionA.Text
'Mostrar descripcion estado



    
End Sub
Sub mostrar_descripcion(nLogSelProceso As Long)
    Dim nMovCancel As Integer
    Set rs = clsDGAdqui.CargaLogSelDescripcionProceso(nLogSelProceso)
    If rs.EOF = True Then
        txttipo.Text = ""
        txtdescripcion.Text = ""
        lblestado.Caption = ""
        Else
            txttipo.Text = UCase(rs!cTipo)
            txtdescripcion.Text = "COTIZACION Nº: " + rs!nLogSelNumeroCot + " - " + rs!cDescripcionProceso + " - TIPO PROCESO: " + rs!nLogSelDescProceso + Space(200) + Str(rs!nLogSelTipoProceso)
            cmdAnulacion.Enabled = True
            lblmotanulacion.Visible = False
            cmbMotAnulacion.Visible = False
            cmdCancelar.Visible = True
            CmdAnular.Visible = True
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
                lblestado.Caption = "ANULADO"
                 'Obtener el Motivo por el que fue Anulado ç
                nMovCancel = clsDGAdqui.CargaLogSelMotCancel(txtSeleccionA.Text)
                ubicar nMovCancel, cmbMotAnulacion
                cmdAnulacion.Enabled = False
                lblmotanulacion.Visible = True
                cmbMotAnulacion.Visible = True
                lblmotanulacion.Enabled = False
                cmbMotAnulacion.Enabled = False
                cmdCancelar.Visible = False
                CmdAnular.Visible = False
                
            End If
    End If
End Sub
Sub Mostrar_Resumen_Cuadro(nLogSelProceso As Long)
    'validar
    Dim nvalida As Integer
    nvalida = clsDGAdqui.ValidaLogSelCalificacionTecnica(nLogSelProceso)
    If nvalida = 0 Then
        FormateaGrilla
        MsgBox "El proceso de seleccion Nº " & txtSeleccionA.Text & " no Tiene Ingresada Calificacion Tecnica para los Proveedores  ", vbInformation, "Verifique las Evaluaciones Tecnicas "
        Exit Sub
    End If
    nvalida = 0
    nvalida = clsDGAdqui.ValidaLogSelCalificacionEconomica(nLogSelProceso)
    If nvalida = 0 Then
        FormateaGrilla_Eco
            MsgBox "El proceso de seleccion Nº " & txtSeleccionA.Text & " no Tiene Ingresada las Cotizaciones de sus Proveedores ", vbInformation, "Verifique Las Cotizaciones de los Proveedores"
            Exit Sub
    End If
    
    Set rs = clsDGAdqui.CargaLogSelResumenCuadro(nLogSelProceso)
    If rs.EOF = True Then
        flexEvaluacionTecnica.Clear
        flexEvaluacionTecnica.Rows = 3
        'Mostrar_Cabecera_Resumen_Cuadro
        Else
        Set flexEvaluacionTecnica.Recordset = rs
        'Mostrar_Cabecera_Resumen_Cuadro
    End If
End Sub
Sub Mostrar_Resumen_Tec(nLogSelProceso As Long)
    'validar
    Dim nvalida As Integer
    nvalida = clsDGAdqui.ValidaLogSelCalificacionTecnica(nLogSelProceso)
    If nvalida = 0 Then
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
        'Mostrar_Cabecera_Resumen
        Else
        Set flexEvaluacionTecnica.Recordset = rs
        'Mostrar_Cabecera_Resumen
    End If
    
    
End Sub
Sub Mostrar_Resumen_Eco(nLogSelProceso As Long)
    'validar
    Dim nvalida As Integer
    nvalida = clsDGAdqui.ValidaLogSelCalificacionEconomica(nLogSelProceso)
    If nvalida = 0 Then
        FormateaGrilla_Eco
            MsgBox "El proceso de seleccion Nº " & txtSeleccionA.Text & " no Tiene Ingresada las Cotizaciones de sus Proveedores ", vbInformation, "No se Ingreso Ninguna Cotizacion de los Proveedores"
            Exit Sub
    End If
    Set rs = clsDGAdqui.CargaLogSelResumenEvaluacionEconomica(nLogSelProceso)
    If rs.EOF = True Then
        flexEvaluacionTecnica.Clear
        flexEvaluacionTecnica.Rows = 3
        'Mostrar_Cabecera_Resumen_Eco
        Else
        Set flexEvaluacionTecnica.Recordset = rs
        'Mostrar_Cabecera_Resumen_Eco
    End If
End Sub




Sub FormateaGrilla()
    flexEvaluacionTecnica.Clear
    flexEvaluacionTecnica.Cols = 3
    flexEvaluacionTecnica.MergeCells = flexMergeRestrictRows
    flexEvaluacionTecnica.ColWidth(0) = 3400
    flexEvaluacionTecnica.TextMatrix(0, 0) = "Criterio Tecnico"
    flexEvaluacionTecnica.TextMatrix(1, 0) = "Criterio Tecnico"
    flexEvaluacionTecnica.TextMatrix(0, 1) = "Proveedor"
    flexEvaluacionTecnica.TextMatrix(0, 2) = "Proveedor"
    flexEvaluacionTecnica.TextMatrix(1, 1) = "Puntaje"
    flexEvaluacionTecnica.TextMatrix(1, 2) = "Observacion"
    flexEvaluacionTecnica.MergeRow(0) = True
    flexEvaluacionTecnica.MergeCol(0) = True
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
    flexEvaluacionTecnica.TextMatrix(0, 3) = "Precio Referencial"
    flexEvaluacionTecnica.TextMatrix(1, 3) = "Precio Referencial"
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

Sub mostrar_detalles(nLogSelProceso As Long)
'Comite
Set rs_sel = clsDGAdqui.CargaSeleccionComite(nLogSelProceso)
If rs_sel.EOF = True Then
   FlexComite.Rows = 2
   FlexComite.Clear
   FlexComite.FormaCabecera
   Else
   
   Set FlexComite.Recordset = rs_sel
End If
'Proveedores
Set rs = clsDGAdqui.CargaLogSelProveedores(nLogSelProceso)
If rs.EOF = True Then
    FlexProvedores.Rows = 2
    FlexProvedores.Clear
    FlexProvedores.FormaCabecera
    Else
    Set FlexProvedores.Recordset = rs
End If
'Detalle
Set rs = clsDGAdqui.CargaSelDetalle(nLogSelProceso, 1)
    If Not rs.EOF = True Then
        Set fgeBienesConfig.Recordset = rs
            fgeBienesConfig.AdicionaFila
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 1) = " --------------------------- "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 2) = " --------------------- TOTAL ------------- "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 3) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 5) = " --------------- " 'fgeBienesConfig.SumaRow(5)
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 6) = " ------------ "
            fgeBienesConfig.TextMatrix(fgeBienesConfig.Rows - 1, 7) = Format(fgeBienesConfig.SumaRow(7), "########.00")
        Else
        fgeBienesConfig.Clear
        fgeBienesConfig.FormaCabecera
        fgeBienesConfig.Rows = 2
    End If
'Criterios Tecnicos
Set rs = clsDGAdqui.CargaLogSelCriteriosProceso(nLogSelProceso, 1)
If rs.EOF = True Then
    FlexCriterios.Rows = 2
    FlexCriterios.Clear
    FlexCriterios.FormaCabecera
    Else
    Set FlexCriterios.Recordset = rs
End If
End Sub
Sub ubicar(codigo As Integer, combo As ComboBox)
Dim i As Integer
For i = 0 To combo.ListCount
If Right(combo.List(i), 1) = codigo Then
    combo.ListIndex = i
    Exit For
    End If
Next
End Sub


