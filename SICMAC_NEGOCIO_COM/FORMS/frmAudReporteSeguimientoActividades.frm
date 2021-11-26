VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAudReporteSeguimientoActividades 
   Caption         =   "AUDITORIA: SEGUIMIENTO DE ACTIVIDADES"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13215
   Icon            =   "frmAudReporteSeguimientoActividades.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   6330
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin SICMACT.FlexEdit grdDatos 
      Height          =   4095
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   7223
      Cols0           =   11
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-CODIGO-ACTIVIDAD-TIPO-ORIGEN-Nº PROC.-OBJETIVO GENERAL-OBJETIVO ESPECIFICO-USUARIO-FEC ASIG.-% AVENCE"
      EncabezadosAnchos=   "450-1200-2600-2500-2600-800-3500-2500-1500-1200-1200"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-L-C-L-L-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   450
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametros"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   315
         Left            =   3720
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   3720
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar Reporte"
         Height          =   315
         Left            =   3720
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cboUsuarios 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker dpkHasta 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   74907649
         CurrentDate     =   41523
      End
      Begin MSComCtl2.DTPicker dpkDesde 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   74907649
         CurrentDate     =   41523
      End
      Begin VB.Label Label3 
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAudReporteSeguimientoActividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sArregloCodPersona() As String

Public Sub CargarColaboradoresUAI()
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
    
    Dim lrDatos As ADODB.Recordset
    Set lrDatos = New ADODB.Recordset
    Set lrDatos = objCOMNAuditoria.ObtenerColaboradoresUAI
    
    Call CargarComboBox(lrDatos, cboUsuarios)
End Sub

Public Sub CargarComboBox(ByVal lrDatos As ADODB.Recordset, ByVal cboControl As ComboBox)
    Dim nContador As Integer
    
    cboUsuarios.AddItem "" & "-- TODOS --"
    cboUsuarios.ItemData(cboUsuarios.NewIndex) = "" & nContador
    ReDim Preserve sArregloCodPersona(nContador + 1)
    sArregloCodPersona(nContador) = ""
    nContador = nContador + 1
    
    Do Until lrDatos.EOF
        cboUsuarios.AddItem "" & lrDatos!cPersNombre
        cboUsuarios.ItemData(cboUsuarios.NewIndex) = "" & nContador
        'cboUsuarios.ItemData(cboUsuarios.NewIndex) = "" & lrDatos!cPersCod
        'cArregloCodPersona(nContador) = lrDatos!cUser
        ReDim Preserve sArregloCodPersona(nContador + 1)
        sArregloCodPersona(nContador) = lrDatos!cPersCod
        lrDatos.MoveNext
        nContador = nContador + 1
    Loop
    
    Set lrDatos = Nothing
    
    cboUsuarios.ListIndex = 0
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdExportar_Click()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsFile As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim i As Integer: Dim IniTablas As Integer
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    
    
    lsNomHoja = "Hoja1"
    lsFile = "Reporte_Seguimiento_Actividad"
    
    lsArchivo = "\spooler\" & "Reporte_Seguimiento_Actividad" & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
     
    IniTablas = 4
    For i = 1 To grdDatos.Rows - 1
        xlHoja1.Cells(IniTablas + i, 2) = grdDatos.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 3) = grdDatos.TextMatrix(i, 2)
        xlHoja1.Cells(IniTablas + i, 4) = grdDatos.TextMatrix(i, 3)
        xlHoja1.Cells(IniTablas + i, 5) = grdDatos.TextMatrix(i, 4)
        xlHoja1.Cells(IniTablas + i, 6) = grdDatos.TextMatrix(i, 5)
        xlHoja1.Cells(IniTablas + i, 7) = grdDatos.TextMatrix(i, 6)
        xlHoja1.Cells(IniTablas + i, 8) = grdDatos.TextMatrix(i, 7)
        xlHoja1.Cells(IniTablas + i, 9) = grdDatos.TextMatrix(i, 8)
        xlHoja1.Cells(IniTablas + i, 10) = grdDatos.TextMatrix(i, 9)
        xlHoja1.Cells(IniTablas + i, 11) = grdDatos.TextMatrix(i, 10)
    Next i
    
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(i + 2, 10)).Borders.LineStyle = 1
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.path & lsArchivo
    psArchivoAGrabarC = App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
End Sub

Private Sub cmdGenerar_Click()
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
    grdDatos.Clear
    grdDatos.FormaCabecera
    grdDatos.rsFlex = objCOMNAuditoria.ReporteSeguimientoActividades(dpkDesde.value, dpkHasta.value, _
                                                                    sArregloCodPersona(cboUsuarios.ItemData(cboUsuarios.ListIndex)))
End Sub

Private Sub Form_Load()
    CargarColaboradoresUAI
    dpkDesde.value = gdFecSis
    dpkHasta.value = gdFecSis
End Sub
