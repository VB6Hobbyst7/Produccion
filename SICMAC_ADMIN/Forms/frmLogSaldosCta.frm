VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogSaldosCta 
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   Icon            =   "frmLogSaldosCta.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7470
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7815
      TabIndex        =   6
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar >>>"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdver 
      Caption         =   "Generar"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFSALDO 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11245
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      BackColorUnpopulated=   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker dtFecha 
      Height          =   330
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   582
      _Version        =   393216
      Format          =   58195969
      CurrentDate     =   38503
   End
   Begin VB.OLE OLE1 
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Saldo"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "frmLogSaldosCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Para Exportar a Excel
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet





Private Sub cmdExportar_Click()
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    If Me.MSHFSALDO.TextMatrix(1, 1) = "" Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Date), "yyyy") & Format(Time, "hhmmss") & ".xls"
    OLE1.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       GeneraReporteSaldo MSHFSALDO, xlHoja1
       OLE1.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OLE1.SourceDoc = lsArchivoN
       OLE1.Verb = 1
       OLE1.Action = 1
       OLE1.DoVerb -1
       OLE1.Appearance = 0
       OLE1.Width = 500
    End If
    MousePointer = 0
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdver_Click()
 Dim oALmacen As DLogAlmacen
 Set oALmacen = New DLogAlmacen
        
 Set MSHFSALDO.Recordset = oALmacen.GetLogAlmacenSaldo(CDate(Me.dtFecha))
 
 
 MSHFSALDO.ColWidth(0) = 1200
 MSHFSALDO.ColWidth(1) = 2800
 MSHFSALDO.ColWidth(2) = 1500
 MSHFSALDO.ColWidth(3) = 1000
 MSHFSALDO.ColWidth(4) = 1000
 MSHFSALDO.ColWidth(5) = 0
 
     
End Sub

Public Sub GeneraReporteSaldo(pflex As MSHFlexGrid, pxlHoja1 As Excel.Worksheet, Optional pnColFiltroVacia As Integer = 0)
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
            For J = 0 To pflex.Cols - 2
                pxlHoja1.Cells(i + 1, J + 1) = pflex.TextMatrix(i, J)
            Next J
        Else
            If pflex.TextMatrix(i, pnColFiltroVacia) <> "" Then
                For J = 0 To pflex.Cols - 2
                    pxlHoja1.Cells(i + 1, J + 1) = pflex.TextMatrix(i, J)
                Next J
            End If
        End If
    Next i
    
End Sub




