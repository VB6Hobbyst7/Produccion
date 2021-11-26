VERSION 5.00
Begin VB.Form frmRHPrePlanilla 
   Caption         =   "Pre Planilla"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12690
   Icon            =   "frmRHPrePlanilla.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   12690
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   6060
      Width           =   1455
   End
   Begin VB.CommandButton cmdexporta 
      Caption         =   "Exportar >>>"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   6060
      Width           =   1455
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   8190
      TabIndex        =   4
      Top             =   6045
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   9675
      TabIndex        =   6
      Top             =   6045
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   11130
      TabIndex        =   5
      Top             =   6045
      Width           =   1455
   End
   Begin Sicmact.FlexEdit FlexPrePla 
      Height          =   5535
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   12615
      _ExtentX        =   22040
      _ExtentY        =   8070
      Cols0           =   4
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Item-Cod Pers-Cod RH-Nombres"
      EncabezadosAnchos=   "500-1200-1200-1200"
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
      ColumnasAEditar =   "X-X-X-X"
      ListaControles  =   "0-0-0-0"
      EncabezadosAlineacion=   "C-C-L-L"
      FormatosEdit    =   "0-0-0-0"
      CantEntero      =   10
      AvanceCeldas    =   1
      TextArray0      =   "Item"
      lbEditarFlex    =   -1  'True
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
      CellBackColor   =   -2147483624
   End
   Begin Sicmact.TxtBuscar TxtPlanilla 
      Height          =   300
      Left            =   1230
      TabIndex        =   0
      Top             =   120
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   529
      Appearance      =   0
      BackColor       =   -2147483624
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      sTitulo         =   ""
   End
   Begin VB.OLE OleExcel 
      Class           =   "Excel.Sheet.8"
      Height          =   255
      Left            =   3480
      OleObjectBlob   =   "frmRHPrePlanilla.frx":030A
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblPlaCod 
      Caption         =   "Cod.Planilla:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   165
      Width           =   990
   End
   Begin VB.Label lblPlanilla 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2685
      TabIndex        =   1
      Top             =   120
      Width           =   4830
   End
End
Attribute VB_Name = "frmRHPrePlanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_pre As New ADODB.Recordset
Dim oPla As DActualizaDatosConPlanilla
Dim oPlan As NActualizaDatosConPlanilla
Dim Progress As clsProgressBar

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet


Private Sub CmdCancelar_Click()
Dim nCols As Integer
Dim sCodigo  As String

If TxtPlanilla.Text = "" Then Exit Sub
If MsgBox("¿ Desea Cancelar lo Editado y Recuperar la Configuracion Original ", vbQuestion + vbYesNo, " Cancelar Edicion Efectuada ") = vbYes Then
        MousePointer = vbHourglass
        FlexPrePla.Cols = 12
        FlexPrePla.ColumnasAEditar = "X-X-X-X-5-6-7-8-9-10-11-12"
        FlexPrePla.EncabezadosAnchos = "500-1300-800-2700-1000-1000-1000-1000-1000-1000-1000-1000"
        FlexPrePla.FormatosEdit = "0-0-0-0-2-2-2-2-2-2-2-2"
        FlexPrePla.EncabezadosAlineacion = "C-C-L-L-R-R-R-R-R-R-R-R"
        Set FlexPrePla.Recordset = oPla.GetPrePlanilla(TxtPlanilla.Text)
        'Obtiene Nombres
        For i = 4 To FlexPrePla.Cols - 1
            sCodigo = Right(Trim(FlexPrePla.TextMatrix(0, i)), 3)
            'obtiene Nombres
            nombres = oPla.GetRHNombreConcepto(sCodigo)
            FlexPrePla.TextMatrix(0, i) = Trim(nombres)
            FlexPrePla.ColWidth(i) = 1600
        Next
        MousePointer = vbArrow
End If
End Sub

Private Sub cmdexporta_Click()
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    If Me.FlexPrePla.TextMatrix(1, 1) = "" Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Date), "yyyy") & Format(Time, "hhmmss") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       GeneraReportePla FlexPrePla, xlHoja1
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
       OleExcel.Appearance = 0
       OleExcel.Width = 500
    End If
    MousePointer = 0

End Sub

Private Sub cmdGrabar_Click()

Set oPlan = New NActualizaDatosConPlanilla
Dim f As Long
Dim i As Long
Dim sCodConcepto As String
Dim sCodPerson As String
If MsgBox("¿ Estás seguro de Grabar ", vbQuestion + vbYesNo, "Grabar Cambios Efectuados en los Montos de los Conceptos ") = vbYes Then
      oPlaEvento_ShowProgress
        For f = 1 To FlexPrePla.Rows - 2
            sCodPerson = Trim(FlexPrePla.TextMatrix(f, 1))
            For i = 4 To FlexPrePla.Cols - 1
            sCodConcepto = Right(Trim(FlexPrePla.TextMatrix(0, i)), 3)
            oPlan.ActualizaConceptosPla TxtPlanilla.Text, sCodPerson, sCodConcepto, IIf(FlexPrePla.TextMatrix(f, i) = "", 0, FlexPrePla.TextMatrix(f, i)), "1"
            Next
            oPlaEvento_Progress f, FlexPrePla.Rows - 1
        Next
            oPlaEvento_CloseProgress
End If
End Sub

Private Sub CmdLimpiar_Click()
Dim i As Long
Dim J As Long
If MsgBox("¿ Estás seguro de Limpiar Los Conceptos", vbQuestion + vbYesNo, "No se Grabaran los Datos si no pulsa Grabar") = vbYes Then
    oPlaEvento_ShowProgress
    For i = 1 To FlexPrePla.Rows - 1
        For J = 4 To FlexPrePla.Cols - 1
            FlexPrePla.TextMatrix(i, J) = 0
        Next
        oPlaEvento_Progress i, FlexPrePla.Rows - 1
    Next
    oPlaEvento_CloseProgress
End If

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub


Private Sub Form_Load()
Me.Width = 12780
Me.Height = 6915
Set oPla = New DActualizaDatosConPlanilla
Set Progress = New clsProgressBar
TxtPlanilla.rs = oPla.GetPlanillas(, True)
FlexPrePla.Enabled = True
FlexPrePla.BackColorBkg = 16777215
Set rs_pre = New ADODB.Recordset

End Sub

Private Sub TxtPlanilla_EmiteDatos()
Dim nCols As Integer
Dim sCodigo  As String

Dim i As Integer
Dim J As Integer
Dim lnAcumulador As Currency


Me.lblPlanilla.Caption = TxtPlanilla.psDescripcion
If TxtPlanilla.Text = "" Then Exit Sub

MousePointer = vbHourglass
FlexPrePla.Cols = 12
FlexPrePla.ColumnasAEditar = "X-X-X-X-5-6-7-8-9-10-11-12"
FlexPrePla.EncabezadosAnchos = "500-1300-800-2700-1000-1000-1000-1000-1000-1000-1000-1000"
FlexPrePla.FormatosEdit = "0-0-0-0-2-2-2-2-2-2-2-2"
FlexPrePla.EncabezadosAlineacion = "C-C-L-L-R-R-R-R-R-R-R-R"
Set FlexPrePla.Recordset = oPla.GetPrePlanilla(TxtPlanilla.Text)
'Obtiene Nombres
For i = 4 To FlexPrePla.Cols - 1
    sCodigo = Right(Trim(FlexPrePla.TextMatrix(0, i)), 3)
    'obtiene Nombres
    nombres = oPla.GetRHNombreConcepto(sCodigo)
    FlexPrePla.TextMatrix(0, i) = Trim(nombres)
    FlexPrePla.ColWidth(i) = 1600
Next
'poner


If FlexPrePla.TextMatrix(FlexPrePla.Rows - 1, 2) <> "" Then
        FlexPrePla.Rows = FlexPrePla.Rows + 1
        FlexPrePla.TextMatrix(FlexPrePla.Rows - 1, 3) = "TOTAL"
        
        For J = 4 To Me.FlexPrePla.Cols - 1
            lnAcumulador = 0
            If Left(FlexPrePla.TextMatrix(0, J), 2) <> "U_" And Left(FlexPrePla.TextMatrix(0, J), 1) <> "_" Then
                
                For i = 1 To Me.FlexPrePla.Rows - 2
                    If FlexPrePla.TextMatrix(i, J) <> "" Then
                        lnAcumulador = lnAcumulador + CCur(FlexPrePla.TextMatrix(i, J))
                    End If
                Next i
                FlexPrePla.TextMatrix(FlexPrePla.Rows - 1, J) = Format(lnAcumulador, "#,##.00")
                FlexPrePla.Row = FlexPrePla.Rows - 1
                FlexPrePla.Col = J
                FlexPrePla.CellBackColor = &HA0C000
                'FlexPrePla.CellFontBold = True
                lnAcumulador = lnAcumulador + CCur(FlexPrePla.TextMatrix(i, J))
            End If
        Next J
    End If
FlexPrePla.TextMatrix(FlexPrePla.Rows - 1, 1) = FlexPrePla.Rows - 2
FlexPrePla.ColWidth(1) = 0
FlexPrePla.ColWidth(2) = 0


MousePointer = vbArrow
End Sub

Private Sub oPlaEvento_ShowProgress()
    Progress.ShowForm Me
End Sub

Private Sub oPlaEvento_Progress(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Actualizando Montos de Conceptos"
End Sub
Private Sub oPlaEvento_Progress2(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Eliminando Consolidado"
End Sub

Private Sub oPlaEvento_CloseProgress()
    Progress.CloseForm Me
End Sub


Public Sub GeneraReportePla(pflex As FlexEdit, pxlHoja1 As Excel.Worksheet, Optional pnColFiltroVacia As Integer = 0)
    Dim i As Integer
    Dim K As Integer
    Dim J As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    
    For i = 0 To pflex.Rows - 1
        If pnColFiltroVacia = 0 Then
            For J = 0 To pflex.Cols - 1
                pxlHoja1.Cells(i + 1, J + 1) = pflex.TextMatrix(i, J)
            Next J
        Else
            If pflex.TextMatrix(i, pnColFiltroVacia) <> "" Then
                For J = 0 To pflex.Cols - 1
                    pxlHoja1.Cells(i + 1, J + 1) = pflex.TextMatrix(i, J)
                Next J
            End If
        End If
    Next i
    
End Sub
