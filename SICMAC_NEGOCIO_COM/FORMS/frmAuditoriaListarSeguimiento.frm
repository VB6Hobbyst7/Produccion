VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAuditoriaListarSeguimiento 
   Caption         =   "Listar Seguimiento"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   Icon            =   "frmAuditoriaListarSeguimiento.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   4575
      Left            =   8160
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
      Begin VB.CommandButton cmdEstadistico 
         Caption         =   "Estad"
         Height          =   350
         Left            =   120
         TabIndex        =   15
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actual."
         Height          =   350
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   7935
      Begin MSDataGridLib.DataGrid dgBuscar 
         Height          =   4215
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   7435
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "iSesionDirId"
            Caption         =   "Id"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "SesionN"
            Caption         =   "Sesion Nro"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "vFecha"
            Caption         =   "Fecha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "tAsunto"
            Caption         =   "Asunto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "vSituacion"
            Caption         =   "Situacion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "vAcuerdo"
            Caption         =   "Acuerdo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            Size            =   182
            BeginProperty Column00 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2099.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label lblMensaje 
         Caption         =   "NO EXISTEN DATOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton btnBuscar 
         Caption         =   "Buscar"
         Height          =   350
         Left            =   6960
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   4800
         TabIndex        =   6
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   67698689
         CurrentDate     =   40234
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   67698689
         CurrentDate     =   40234
      End
      Begin VB.ComboBox cmbSituacion 
         Height          =   315
         ItemData        =   "frmAuditoriaListarSeguimiento.frx":030A
         Left            =   1200
         List            =   "frmAuditoriaListarSeguimiento.frx":031A
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Situacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   240
      Left            =   0
      SizeMode        =   1  'Stretch
      TabIndex        =   14
      Top             =   6000
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "frmAuditoriaListarSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objCOMNAuditoria As COMNAuditoria.NCOMSeguimiento
Dim lsmensaje As String

Dim oInventario As COMNAuditoria.NCOMSeguimiento
Dim oArea As DActualizaDatosArea
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Public Sub BuscarDatos()
    Dim rs As ADODB.Recordset
    Dim contador As Integer
    Set oInventario = New COMNAuditoria.NCOMSeguimiento
    lsmensaje = ""
    Set rs = oInventario.BuscarAcuerdoDir(DTPicker1.value, DTPicker2.value, lsmensaje)
        If lsmensaje = "" Then
            lblMensaje.Visible = False
            dgBuscar.Visible = True
            Set dgBuscar.DataSource = rs
            dgBuscar.Refresh
            Screen.MousePointer = 0
            dgBuscar.SetFocus
        Else
            Set dgBuscar.DataSource = Nothing
            dgBuscar.Refresh
            lblMensaje.Visible = True
            dgBuscar.Visible = False
        End If
        Set rs = Nothing
        Set objCOMNAuditoria = Nothing
End Sub

Private Sub btnBuscar_Click()
    BuscarDatos
End Sub

Private Sub cmdActualizar_Click()
    If lsmensaje = "" Then
        gSesionDirId = dgBuscar.Columns(0).Text
        gNroAcuerdo = dgBuscar.Columns(5).Text
        frmAuditoriaRegistrarAcuerdo.Show
    End If
End Sub

Private Sub cmdEstadistico_Click()
    Dim rsDatos As ADODB.Recordset
    Set rsDatos = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    Set objCOMNAuditoria = New COMNAuditoria.NCOMSeguimiento
    
    Set rsDatos = objCOMNAuditoria.ObtenerReporteEstadistico
    
        If rsDatos Is Nothing Then
            MsgBox "No existen datos.", vbInformation, "Aviso"
            Exit Sub
        End If
    
    lsArchivoN = App.path & "\Spooler\" & "ReporteEstadistico" & Format(CDate(gdFecSis), "yyyymmdd") & Format(Time, "hhmmss") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       ReporteCabeceraExcelEstadistico xlHoja1
       GeneraReporteEstadistico rsDatos
       OleExcel.Class = "ExcelWorkSheet"
       gFunGeneral.ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
    MousePointer = 0
End Sub

Private Sub cmdExcel_Click()
    Dim rsDatos As ADODB.Recordset
    Set rsDatos = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    Set objCOMNAuditoria = New COMNAuditoria.NCOMSeguimiento
    
    Set rsDatos = objCOMNAuditoria.ObtenerReporteAcuerdoDir(DTPicker1.value, DTPicker2.value, Mid(cmbSituacion.Text, 1, 1))
    
        If rsDatos Is Nothing Then
            MsgBox "No existen datos.", vbInformation, "Aviso"
            Exit Sub
        End If
    
    lsArchivoN = App.path & "\Spooler\" & "ReporteSeguimiento" & Format(CDate(gdFecSis), "yyyymmdd") & Format(Time, "hhmmss") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       ReporteCabeceraExcel xlHoja1
       GeneraReporte rsDatos
       OleExcel.Class = "ExcelWorkSheet"
       gFunGeneral.ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
    MousePointer = 0
End Sub

Public Function ReporteCabeceraExcelEstadistico(Optional xlHoja1 As Excel.Worksheet) As String
    xlHoja1.PageSetup.LeftMargin = 1.5
    xlHoja1.PageSetup.RightMargin = 0
    xlHoja1.PageSetup.BottomMargin = 1
    xlHoja1.PageSetup.TopMargin = 1
    xlHoja1.PageSetup.Zoom = 70
    
    xlHoja1.Cells(1, 1) = "CMAC-MAYNAS S.A."
    
    xlHoja1.Cells(2, 1) = "Órgano de Control Institucional"
    
    xlHoja1.Cells(3, 3) = "RESUMEN ESTADISTICO DE ACUERDOS DE DIRECTORIO"
    
    
    xlHoja1.Cells(4, 3) = "DE LA CMACMA-MAYNAS S.A."
    xlHoja1.Cells(5, 4) = "AL:" & " " & DTPicker2.value
    
    xlHoja1.Cells(1, 6) = "Anexo Nº 4"
    
    xlHoja1.Cells(9, 1) = "FECHA"
    xlHoja1.Cells(9, 2) = "SESION"
    xlHoja1.Cells(9, 3) = "EN PROCESO"
    xlHoja1.Cells(9, 4) = "PENDIENTES"
    xlHoja1.Cells(9, 5) = "IMPLEMENTADOS"
    xlHoja1.Cells(9, 6) = "TOTAL"
           
       
    xlHoja1.Range("A3:G6").Font.Bold = True
    xlHoja1.Range("A8:H9").Font.Bold = True
    
    xlHoja1.Range("A8:F9").BorderAround xlContinuous, xlMedium
    xlHoja1.Range("A8:F9").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range("A8:F9").Borders(xlInsideVertical).LineStyle = xlContinuous
    
    'Titulo CMAC
    xlHoja1.Range("A1:A1").Font.Bold = True
    
    xlHoja1.Range("F1:F1").Font.Bold = True
    
    xlHoja1.Range("A2:B2").MergeCells = True
    xlHoja1.Range("A2:B2").Font.Bold = True
    
    'Border y Merge del Titulo
    xlHoja1.Range("B3:E3").MergeCells = True
    xlHoja1.Range("B3:E3").Font.Bold = True
    
    xlHoja1.Range("B4:E4").MergeCells = True
    xlHoja1.Range("B4:E4").Font.Bold = True
    
    xlHoja1.Range("C5:D5").MergeCells = True
    xlHoja1.Range("D5:D5").Font.Bold = True
     
    xlHoja1.Range("A4:A4").ColumnWidth = 15
    xlHoja1.Range("B4:B4").ColumnWidth = 30
    xlHoja1.Range("C4:C4").ColumnWidth = 12
    xlHoja1.Range("D4:D4").ColumnWidth = 12
    xlHoja1.Range("E4:E4").ColumnWidth = 15
    xlHoja1.Range("F4:F4").ColumnWidth = 15
   
    xlHoja1.Range("B4:B4").HorizontalAlignment = xlLeft
    xlHoja1.Range("B5:B5").HorizontalAlignment = xlLeft
    xlHoja1.Range("B6:B6").HorizontalAlignment = xlLeft
    xlHoja1.Range("A8:H9").HorizontalAlignment = xlCenter
    xlHoja1.Range("A8:H9").VerticalAlignment = xlCenter
    
    'Centrar Titulo
    xlHoja1.Range("B3:E3").HorizontalAlignment = xlCenter
    xlHoja1.Range("B3:E3").VerticalAlignment = xlCenter

    xlHoja1.Range("B4:E4").HorizontalAlignment = xlCenter
    xlHoja1.Range("B4:E4").VerticalAlignment = xlCenter

    xlHoja1.Range("C5:D5").HorizontalAlignment = xlCenter
    xlHoja1.Range("C5:D5").VerticalAlignment = xlCenter
    
    xlHoja1.Range("F1:F1").HorizontalAlignment = xlRight
    
    'Tamaño de letra
    'Titulo CMAC
    xlHoja1.Range("A1:B1").Font.Size = 6
    xlHoja1.Range("A2:B2").Font.Size = 6
    
    xlHoja1.Range("F1:F1").Font.Size = 6
    
    'Titulo
    xlHoja1.Range("B3:E3").Font.Size = 8
    xlHoja1.Range("B4:E4").Font.Size = 8
    xlHoja1.Range("C5:D5").Font.Size = 8

    'Encabeza
    xlHoja1.Range("A9:E9").Font.Size = 7
    
    'Resto
    xlHoja1.Range("A9:F100").Font.Size = 7
    
End Function

Public Function ReporteCabeceraExcel(Optional xlHoja1 As Excel.Worksheet) As String
    xlHoja1.PageSetup.LeftMargin = 1.5
    xlHoja1.PageSetup.RightMargin = 0
    xlHoja1.PageSetup.BottomMargin = 1
    xlHoja1.PageSetup.TopMargin = 1
    xlHoja1.PageSetup.Zoom = 70
    
    xlHoja1.Cells(1, 1) = "CMAC-MAYNAS S.A."
    
    xlHoja1.Cells(2, 1) = "Órgano de Control Institucional"
    
    xlHoja1.Cells(3, 3) = "PRINCIPALES ACUERDOS DE DIRECTORIO"
    
    
    xlHoja1.Cells(4, 3) = "EN PROCESO DE IMPLEMENTACION"
    xlHoja1.Cells(5, 4) = "AL:" & " " & DTPicker2.value
    
    Dim var As String
    If cmbSituacion.Text = "En Proceso" Then
        var = "Anexo Nº " & "1"
    End If
    If cmbSituacion.Text = "Pendiente" Then
        var = "Anexo Nº " & "2"
    End If
    If cmbSituacion.Text = "Superada" Then
        var = "Anexo Nº " & "3"
    End If
    
    xlHoja1.Cells(1, 5) = var
    
    xlHoja1.Cells(9, 1) = "Sesion Nº"
    xlHoja1.Cells(9, 2) = "Asunto"
    xlHoja1.Cells(9, 3) = "Acuerdo Nº"
    xlHoja1.Cells(9, 4) = "Detalle"
    xlHoja1.Cells(9, 5) = "Situacion a la Fecha"
           
       
    xlHoja1.Range("A3:G6").Font.Bold = True
    xlHoja1.Range("A8:H9").Font.Bold = True
    
    xlHoja1.Range("A8:E9").BorderAround xlContinuous, xlMedium
    xlHoja1.Range("A8:E9").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range("A8:E9").Borders(xlInsideVertical).LineStyle = xlContinuous
    
    'Titulo CMAC
    xlHoja1.Range("A1:A1").Font.Bold = True
    
    xlHoja1.Range("E1:E1").Font.Bold = True
    
    xlHoja1.Range("A2:B2").MergeCells = True
    xlHoja1.Range("A2:B2").Font.Bold = True
    
    'Border y Merge del Titulo
    xlHoja1.Range("B3:E3").MergeCells = True
    xlHoja1.Range("B3:E3").Font.Bold = True
    
    xlHoja1.Range("B4:E4").MergeCells = True
    xlHoja1.Range("B4:E4").Font.Bold = True
    
    xlHoja1.Range("C5:D5").MergeCells = True
    xlHoja1.Range("D5:D5").Font.Bold = True
     
    xlHoja1.Range("A4:A4").ColumnWidth = 15
    xlHoja1.Range("B4:B4").ColumnWidth = 30
    xlHoja1.Range("C4:C4").ColumnWidth = 10
    xlHoja1.Range("D4:D4").ColumnWidth = 30
    xlHoja1.Range("E4:E4").ColumnWidth = 30
   
    xlHoja1.Range("B4:B4").HorizontalAlignment = xlLeft
    xlHoja1.Range("B5:B5").HorizontalAlignment = xlLeft
    xlHoja1.Range("B6:B6").HorizontalAlignment = xlLeft
    xlHoja1.Range("A8:H9").HorizontalAlignment = xlCenter
    xlHoja1.Range("A8:H9").VerticalAlignment = xlCenter
    
    'Centrar Titulo
    xlHoja1.Range("B3:E3").HorizontalAlignment = xlCenter
    xlHoja1.Range("B3:E3").VerticalAlignment = xlCenter

    xlHoja1.Range("B4:E4").HorizontalAlignment = xlCenter
    xlHoja1.Range("B4:E4").VerticalAlignment = xlCenter

    xlHoja1.Range("C5:D5").HorizontalAlignment = xlCenter
    xlHoja1.Range("C5:D5").VerticalAlignment = xlCenter
    
    xlHoja1.Range("E1:E1").HorizontalAlignment = xlRight
    
    'Tamaño de letra
    'Titulo CMAC
    xlHoja1.Range("A1:B1").Font.Size = 6
    xlHoja1.Range("A2:B2").Font.Size = 6
    
    xlHoja1.Range("E1:E1").Font.Size = 6
    
    'Titulo
    xlHoja1.Range("B3:E3").Font.Size = 8
    xlHoja1.Range("B4:E4").Font.Size = 8
    xlHoja1.Range("C5:D5").Font.Size = 8

    'Encabeza
    xlHoja1.Range("A9:E9").Font.Size = 7
    
    'Resto
    xlHoja1.Range("A9:E100").Font.Size = 7
    
End Function

Private Sub GeneraReporte(prRs As ADODB.Recordset)
    Dim i As Integer
    Dim J As Integer
        
    i = 8
    While Not prRs.EOF
        i = i + 1
                
        For J = 0 To prRs.Fields.Count - 1
                xlHoja1.Cells(i + 1, J + 1) = prRs.Fields(J)

            
            
        Next J

        
        prRs.MoveNext
  
    Wend
End Sub

Private Sub GeneraReporteEstadistico(prRs As ADODB.Recordset)
    Dim i As Integer
    Dim J As Integer
    
    Dim TotalE As Integer
    Dim TotalP As Integer
    Dim TotalI As Integer
    Dim TotalT As Integer
    
    i = 8
    While Not prRs.EOF
        i = i + 1
        
        Dim val As Integer
        
        For J = 0 To prRs.Fields.Count - 1
            If J = 0 Or J = 1 Then
                xlHoja1.Cells(i + 1, J + 1) = prRs.Fields(J)
            Else
                xlHoja1.Cells(i + 1, J + 1) = prRs.Fields(J)
                val = val + prRs.Fields(J)
                
                If J = 4 Then
                    TotalI = TotalI + prRs.Fields(J)
                End If
                
                If J = 3 Then
                    TotalP = TotalP + prRs.Fields(J)
                End If
                
                 If J = 2 Then
                    TotalE = TotalE + prRs.Fields(J)
                End If
                
            End If
            
            
        Next J
        
         xlHoja1.Cells(i + 1, J + 1) = val
         TotalT = TotalT + val
        
        prRs.MoveNext
        
        val = 0
        
    Wend
    
    xlHoja1.Cells(i + prRs.Fields.Count - 2, 6) = TotalT
    xlHoja1.Cells(i + prRs.Fields.Count - 2, 5) = TotalI
    xlHoja1.Cells(i + prRs.Fields.Count - 2, 4) = TotalP
    xlHoja1.Cells(i + prRs.Fields.Count - 2, 3) = TotalE
End Sub

Private Sub Form_Load()
    cmbSituacion.SelText = "Todos"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gSesionDirId = 0
End Sub
