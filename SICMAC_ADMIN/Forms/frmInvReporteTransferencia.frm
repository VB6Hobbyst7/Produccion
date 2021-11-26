VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInvReporteTransferencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPORTE DE LAS TRANSFERENCIAS"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvReporteTransferencia.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de Búsqueda"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   59572225
         CurrentDate     =   39883
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   59572225
         CurrentDate     =   39883
      End
      Begin VB.OLE OleExcel 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         Enabled         =   0   'False
         Height          =   240
         Left            =   0
         SizeMode        =   1  'Stretch
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "De:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmInvReporteTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oInventario As NInvTransferencia
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Private Sub Form_Load()
    DTPicker1.value = Date
    DTPicker2.value = Date
End Sub

Private Sub Command1_Click()
    Dim rsDatos As ADODB.Recordset
    Set rsDatos = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    Set oInventario = New NInvTransferencia
           
    Set rsDatos = oInventario.ObtenerReporteTransferencia(Mid(DTPicker1.value, 7, 4) & Mid(DTPicker1.value, 4, 2) & Mid(DTPicker1.value, 1, 2), Mid(DTPicker2.value, 7, 4) & Mid(DTPicker2.value, 4, 2) & Mid(DTPicker2.value, 1, 2))
    
    If rsDatos Is Nothing Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lsArchivoN = App.path & "\Spooler\" & "ReporteActivoFijo" & Format(CDate(gdFecSis), "yyyymmdd") & Format(Time, "hhmmss") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       ReporteTransferenciaCabeceraExcel xlHoja1
       GeneraReporte rsDatos
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
    MousePointer = 0
End Sub

Private Sub GeneraReporte(prRs As ADODB.Recordset)
    Dim i As Integer
    Dim j As Integer
    Dim lNegativo As Boolean
        
    i = 8
    While Not prRs.EOF
        i = i + 1
        For j = 0 To prRs.Fields.Count - 1
            xlHoja1.Cells(i + 1, j + 1) = prRs.Fields(j)
        Next j
        prRs.MoveNext
    Wend
    
'   'Border's Tabla
    xlHoja1.Range("A10:E" & prRs.RecordCount + 9).BorderAround xlContinuous, xlMedium
    
    If prRs.RecordCount <> "1" Then
        xlHoja1.Range("A10:E" & prRs.RecordCount + 9).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Else
        xlHoja1.Range("A10:E" & prRs.RecordCount + 10).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End If
    
    xlHoja1.Range("A10:E" & prRs.RecordCount + 9).Borders(xlInsideVertical).LineStyle = xlContinuous

End Sub

Public Function ReporteTransferenciaCabeceraExcel(Optional xlHoja1 As Excel.Worksheet) As String
    xlHoja1.PageSetup.LeftMargin = 1.5
    xlHoja1.PageSetup.RightMargin = 0
    xlHoja1.PageSetup.BottomMargin = 1
    xlHoja1.PageSetup.TopMargin = 1
    xlHoja1.PageSetup.Zoom = 70
    
    xlHoja1.Cells(2, 2) = "REPORTE DE TRANSFERENCIAS" & " " & "DEL" & " " & DTPicker1.value & " " & "AL" & " " & DTPicker2.value
    
    xlHoja1.Cells(4, 1) = "DENOMINACIÓN:"
    xlHoja1.Cells(5, 1) = "FECHA:"
    
    xlHoja1.Cells(4, 2) = "CMAC MAYNAS S.A."
    xlHoja1.Cells(5, 2) = Date
    
    xlHoja1.Range("B4").HorizontalAlignment = xlLeft
    xlHoja1.Range("B5").HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(8, 1) = "DESCRIPCION DEL ACTIVO FIJO"
    xlHoja1.Cells(8, 2) = "TIPO TRANSFERENCIA"
    
    xlHoja1.Cells(8, 3) = "FECHA DE TRANSFERENCIA"
    
    xlHoja1.Cells(8, 4) = "ORIGEN"
    xlHoja1.Cells(8, 5) = "DESTINO"
   
    xlHoja1.Range("A2:E5").Font.Bold = True
    xlHoja1.Range("A8:E9").Font.Bold = True

    xlHoja1.Range("A2:E2").Font.Size = 12
    xlHoja1.Range("A4:E4").Font.Size = 9
    xlHoja1.Range("A4:E9").Font.Size = 9

    xlHoja1.Range("A8:E9").BorderAround xlContinuous, xlMedium
    xlHoja1.Range("A8:E9").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range("A8:E9").Borders(xlInsideVertical).LineStyle = xlContinuous

    xlHoja1.Range("A8:A9").MergeCells = True
    xlHoja1.Range("B8:B9").MergeCells = True
    xlHoja1.Range("C8:C9").MergeCells = True
    xlHoja1.Range("D8:D9").MergeCells = True
    xlHoja1.Range("E8:E9").MergeCells = True

    xlHoja1.Range("B2:D2").MergeCells = True
    xlHoja1.Range("B2:D2").HorizontalAlignment = xlCenter
    xlHoja1.Range("B2:D2").VerticalAlignment = xlCenter

    xlHoja1.Range("A4:A4").ColumnWidth = 50
    xlHoja1.Range("B4:B4").ColumnWidth = 25
    xlHoja1.Range("C4:C4").ColumnWidth = 30
    xlHoja1.Range("D4:D4").ColumnWidth = 35
    xlHoja1.Range("D4:H4").ColumnWidth = 35

'    xlHoja1.Range("B4:B4").HorizontalAlignment = xlLeft
'    xlHoja1.Range("B5:B5").HorizontalAlignment = xlLeft
'    xlHoja1.Range("B6:B6").HorizontalAlignment = xlLeft
    xlHoja1.Range("A8:E9").HorizontalAlignment = xlCenter
    xlHoja1.Range("A8:E9").VerticalAlignment = xlCenter
    
End Function
