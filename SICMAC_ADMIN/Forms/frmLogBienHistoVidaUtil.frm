VERSION 5.00
Begin VB.Form frmLogBienHistoVidaUtil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activo Fijo: Histórico de Vida Útil"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7470
   Icon            =   "frmLogBienHistoVidaUtil.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   70
      TabIndex        =   2
      Top             =   2640
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   6360
      TabIndex        =   1
      Top             =   2640
      Width           =   1050
   End
   Begin Sicmact.FlexEdit feHisto 
      Height          =   2535
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4471
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "#-Fecha-Usuario-Tiempo (mes)-Motivo-cSerie"
      EncabezadosAnchos=   "350-1200-1000-1500-3000-0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-C-C-C-L-L"
      FormatosEdit    =   "0-0-0-0-0-0"
      CantEntero      =   9
      TextArray0      =   "#"
      SelectionMode   =   1
      lbEditarFlex    =   -1  'True
      TipoBusqueda    =   0
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
   End
End
Attribute VB_Name = "frmLogBienHistoVidaUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'** Nombre : frmLogBienHistoVidaUtil
'** Descripción : Ajuste de Vida Útil Bienes creado segun ERS059-2013
'** Creación : EJVG, 20130621 09:00:00 AM
'***************************************************************************
Option Explicit

Private Sub Form_Load()
    CentraForm Me
End Sub
Public Sub Inicio(ByVal pnMovNro As Long)
    Dim oBien As New DBien
    Dim rs As New ADODB.Recordset
    Dim fila As Long
    
    Set rs = oBien.RecuperaHistorialVidaUtil(pnMovNro)
    LimpiaFlex feHisto
    Do While Not rs.EOF
        feHisto.AdicionaFila
        fila = feHisto.row
        feHisto.TextMatrix(fila, 1) = Format(rs!dfecha, "dd/mm/yyyy")
        feHisto.TextMatrix(fila, 2) = rs!cUsuario
        feHisto.TextMatrix(fila, 3) = rs!nBSPerDeprecia
        feHisto.TextMatrix(fila, 4) = rs!cMotivo 'NAGL 20191226 Según RFC1910190001
        feHisto.TextMatrix(fila, 5) = rs!cSerie 'NAGL 20191226 Según RFC1910190001
        rs.MoveNext
    Loop
    Set rs = Nothing
    Set oBien = Nothing
    Show 1
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdExportar_Click()
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim lnFila As Long, lnColumna As Long, lnColumnaMax As Long
    Dim I As Long, j As Long
    Dim lsArchivo As String
    
On Error GoTo ErrExportar
    
    If FlexVacio(feHisto) Then
        MsgBox "No hay información para exportar a formato Excel", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    lsArchivo = "\spooler\RptAjusteVidaUtilBien" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    Set xlsLibro = xlsAplicacion.Workbooks.Add

    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "Reporte Ajuste Vida Util"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    
    lnFila = 2
    
    For I = 0 To feHisto.Rows - 1
        lnColumna = 2
        For j = 0 To feHisto.Cols - 1
            If feHisto.ColWidth(j) > 0 Or j = 5 Then
                xlsHoja.Cells(lnFila, lnColumna) = "'" & feHisto.TextMatrix(I, j)
                lnColumna = lnColumna + 1
                lnColumnaMax = lnColumna
            End If
        Next
        lnFila = lnFila + 1
    Next

    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).Interior.Color = RGB(191, 191, 191)
    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).Font.Bold = True
    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(lnFila - 1, lnColumnaMax - 1)).Borders.Weight = xlThin

    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(lnFila - 1, lnColumnaMax - 1)).EntireColumn.AutoFit
    
    MsgBox "Se ha exportado satisfactoriamente la información", vbInformation, "Aviso"
    
    xlsHoja.SaveAs App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    
    Screen.MousePointer = 0
    
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlsHoja = Nothing
    Exit Sub
ErrExportar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub 'NAGL 20191226 Según RFC1910190001
