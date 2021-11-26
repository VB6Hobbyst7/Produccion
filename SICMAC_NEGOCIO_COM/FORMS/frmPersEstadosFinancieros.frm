VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPersEstadosFinancieros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estados Financieros"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6510
   Icon            =   "frmPersEstadosFinancieros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEF 
      Caption         =   "Estados Financieros"
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
      Height          =   3375
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   6375
      Begin VB.Frame frameComparaEEFF 
         Caption         =   "Generar Comparativo Anual:"
         Height          =   645
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Width           =   6135
         Begin VB.CommandButton cmdComparar 
            Caption         =   "Generar"
            Height          =   350
            Left            =   3840
            TabIndex        =   18
            ToolTipText     =   "Ver Comparativo Anual EERR / Balance General"
            Top             =   240
            Width           =   2175
         End
         Begin MSMask.MaskEdBox txtFechaComparaFin 
            Height          =   315
            Left            =   2160
            TabIndex        =   19
            Top             =   270
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFechaComparaIni 
            Height          =   315
            Left            =   480
            TabIndex        =   20
            Top             =   270
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            Caption         =   "Del:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "Al:"
            Height          =   255
            Left            =   1980
            TabIndex        =   21
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   5160
         TabIndex        =   15
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Height          =   375
         Left            =   5160
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   5160
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin SICMACT.FlexEdit feEstadoFinanciero 
         Height          =   2175
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3836
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Fecha EF-Fecha Reg-Usuario-Estado-nCodEstadoFinanciero-nAuditado"
         EncabezadosAnchos=   "500-1200-1900-1000-0-0-0"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5080
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   5080
      Width           =   975
   End
   Begin SICMACT.TxtBuscar TxtBCodPers 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   570
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   503
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
      TipoBusqueda    =   3
      sTitulo         =   ""
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6360
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Datos del Titular"
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
      Left            =   255
      TabIndex        =   7
      Top             =   240
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nombre :"
      Height          =   195
      Left            =   165
      TabIndex        =   6
      Top             =   915
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código :"
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   585
      Width           =   585
   End
   Begin VB.Label lblTrib 
      AutoSize        =   -1  'True
      Caption         =   "Doc. Juridico :"
      Height          =   195
      Left            =   3885
      TabIndex        =   4
      Top             =   1275
      Width           =   1020
   End
   Begin VB.Label lblNat 
      AutoSize        =   -1  'True
      Caption         =   "Doc. Natural :"
      Height          =   195
      Left            =   165
      TabIndex        =   3
      Top             =   1245
      Width           =   990
   End
   Begin VB.Label lblcodigo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   5070
      TabIndex        =   2
      Top             =   1230
      Width           =   1335
   End
   Begin VB.Label lblcodigo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   1230
   End
   Begin VB.Label lblcodigo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1200
      TabIndex        =   0
      Top             =   870
      Width           =   5175
   End
End
Attribute VB_Name = "frmPersEstadosFinancieros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************
'** DESARROLLADO POR: FRHU
'** REQUERIMIENTO: ERS013-2015
'** FECHA DESARROLLO: 20150326
'******************************
Option Explicit
Dim sPersCod As String
Dim nTipoAccion As Integer  '1: Nuevo, 2:Editar, 3: Consultar



'->***** LUCV20171015->*****
Private Sub txtFechaComparaIni_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        If Not IsDate(txtFechaComparaIni.Text) Then
            MsgBox "Verifique Dia,Mes,Año de la fecha inicial, Fecha Incorrecta", vbInformation, "Aviso"
            txtFechaComparaIni.SetFocus
        Else
            txtFechaComparaFin.SetFocus
        End If
   End If
End Sub
Private Sub txtFechaComparaFin_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        If Not IsDate(txtFechaComparaIni.Text) Then
            MsgBox "Verifique Dia,Mes,Año de la fecha inicial, Fecha Incorrecta", vbInformation, "Aviso"
            txtFechaComparaFin.SetFocus
        Else
            cmdComparar.SetFocus
        End If
   End If
End Sub
Private Sub cmdComparar_Click()
    Dim lsArchivo As String
    Dim lbLibroOpen As Boolean
    Dim lsPersCod As String
    Dim lsNombre As String
    Dim CargaDatos As Boolean
    Dim lsFechaInicial As String
    Dim lsFechaFinal As String
    Dim rsBalanceActivo As ADODB.Recordset
    Dim rsBalancePasivoPat As ADODB.Recordset
    
    Dim rsEERR As ADODB.Recordset
    Dim oNFormatosEval As COMNCredito.NCOMFormatosEval
    Set oNFormatosEval = New COMNCredito.NCOMFormatosEval
    
    If Trim(TxtBCodPers.Text) = "" Then Exit Sub
    lsPersCod = Trim(TxtBCodPers.Text)
    lsFechaInicial = Format(txtFechaComparaIni.Text, "yyyymmdd")
    lsFechaFinal = Format(txtFechaComparaFin.Text, "yyyymmdd")
    lsNombre = lblcodigo(1).Caption
    
   If ValidaDatos Then
        CargaDatos = oNFormatosEval.CargaDatosComparacionAnual(lsPersCod, _
                                                                lsFechaInicial, _
                                                                lsFechaFinal, _
                                                                rsBalanceActivo, _
                                                                rsBalancePasivoPat, _
                                                                rsEERR)
        If CargaDatos Then
                Call GeneraExcelComparativo(lsNombre, lsFechaInicial, lsFechaFinal, rsBalanceActivo, rsBalancePasivoPat, rsEERR)
               rsBalanceActivo.Close
               Set rsBalanceActivo = Nothing
               
               rsBalancePasivoPat.Close
               Set rsBalancePasivoPat = Nothing
               rsEERR.Close
               Set rsEERR = Nothing
        Else
            MsgBox "Hubo un problema en la carga de la generación del reporte", vbInformation, "Aviso"
        End If
    Else
        MsgBox "Verifique el ingreso de datos", vbInformation, "Aviso"
    End If
End Sub
Public Sub GeneraExcelComparativo(ByVal lsNombre As String, ByVal dFechaInicial As String, ByVal dFechaFinal As String, ByVal prsBalanceActivo As ADODB.Recordset, ByVal prsBalancePasivoPat As ADODB.Recordset, ByVal prsEERR As ADODB.Recordset)
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As New Excel.Workbook
    Dim xlsHoja As New Excel.Worksheet
    Dim lsArchivo As String
    'Ruta Excell
    lsArchivo = "\spooler\RptComparativoAnualEEFF_" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    
    'Hoja Estados de Resultados
    Set xlsLibro = xlsAplicacion.Workbooks.Add
    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "ComparativoEERR"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 10
    xlsHoja.PageSetup.Orientation = xlLandscape
    xlsHoja.PageSetup.CenterHorizontally = True
    xlsHoja.PageSetup.Zoom = 60
    Call GeneraHojaEstadosResultadosRpt(lsNombre, prsEERR, dFechaInicial, dFechaFinal, xlsHoja)
    
    'Hoja Balance General
    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "ComparativoAnualBalanceGeneral"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    xlsHoja.PageSetup.Orientation = xlLandscape
    xlsHoja.PageSetup.CenterHorizontally = True
    xlsHoja.PageSetup.Zoom = 60
    Call GeneraHojaBalanceGeneralRpt(lsNombre, prsBalanceActivo, prsBalancePasivoPat, dFechaInicial, dFechaFinal, xlsHoja)

    'proteger Libro
    xlsAplicacion.ActiveWorkbook.Protect ("" & UCase(gsCodUser) & "" & Format(gdFecSis, "YYYYMMDD") & "")
    xlsAplicacion.Worksheets("ComparativoAnualBalanceGeneral").Protect ("" & UCase(gsCodUser) & "" & Format(gdFecSis, "YYYYMMDD") & "")
    xlsAplicacion.Worksheets("ComparativoEERR").Protect ("" & UCase(gsCodUser) & "" & Format(gdFecSis, "YYYYMMDD") & "")
    
    MsgBox "Se ha generado satisfactoriamente el reporte comparativo anual", vbInformation, "Aviso"
    xlsHoja.SaveAs App.Path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True

    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlsHoja = Nothing
    Exit Sub
End Sub

Public Sub GeneraHojaBalanceGeneralRpt(ByVal psNombre As String, ByVal prsBalanceActivo As ADODB.Recordset, ByVal prsBalancePasivoPat As ADODB.Recordset, ByVal psFechaInicial As String, ByVal psFechaFinal As String, ByRef xlsHoja As Worksheet)
   Dim i As Integer
   Dim lnFila As Integer
    'Configuracion de cabecera
    xlsHoja.Cells(1, 1) = psNombre
    xlsHoja.Range(xlsHoja.Cells(1, 1), xlsHoja.Cells(1, 14)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(1, 1), xlsHoja.Cells(1, 14)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(1, 1), xlsHoja.Cells(1, 14)).Font.Bold = True
    xlsHoja.Range(xlsHoja.Cells(1, 1), xlsHoja.Cells(1, 14)).Cells.Interior.Color = RGB(238, 248, 255)
    
    xlsHoja.Cells(2, 1) = "BALANCE GENERAL"
    xlsHoja.Range(xlsHoja.Cells(2, 1), xlsHoja.Cells(2, 14)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(2, 1), xlsHoja.Cells(2, 14)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(2, 1), xlsHoja.Cells(2, 14)).Font.Bold = True
    
    xlsHoja.Cells(3, 1) = "Comparativos Anuales (Expresado en soles) del " & Left(psFechaInicial, 4) & " al " & Left(psFechaFinal, 4) & " "
    xlsHoja.Range(xlsHoja.Cells(3, 1), xlsHoja.Cells(3, 14)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(3, 1), xlsHoja.Cells(3, 14)).HorizontalAlignment = xlCenter
    
    'Cabecera de Activos
    xlsHoja.Cells(4, 1) = "ACTIVOS"
    xlsHoja.Range(xlsHoja.Cells(4, 1), xlsHoja.Cells(4, 1)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(4, 1), xlsHoja.Cells(4, 1)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 1), xlsHoja.Cells(4, 1)).Font.Bold = True
    
    xlsHoja.Cells(4, 2) = "Total " & Left(psFechaInicial, 4) & ""
    xlsHoja.Range(xlsHoja.Cells(4, 2), xlsHoja.Cells(4, 2)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 2), xlsHoja.Cells(4, 2)).Font.Bold = True
    
    xlsHoja.Cells(4, 3) = "%"
    xlsHoja.Range(xlsHoja.Cells(4, 3), xlsHoja.Cells(4, 3)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 3), xlsHoja.Cells(4, 3)).Font.Bold = True
    
    xlsHoja.Cells(4, 4) = "Total " & Left(psFechaFinal, 4) & ""
    xlsHoja.Range(xlsHoja.Cells(4, 4), xlsHoja.Cells(4, 4)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 4), xlsHoja.Cells(4, 4)).Font.Bold = True
    
    xlsHoja.Cells(4, 5) = "%"
    xlsHoja.Range(xlsHoja.Cells(4, 5), xlsHoja.Cells(4, 5)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 5), xlsHoja.Cells(4, 5)).Font.Bold = True
    
    xlsHoja.Cells(4, 6) = "Var. Abs"
    xlsHoja.Range(xlsHoja.Cells(4, 6), xlsHoja.Cells(4, 6)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 6), xlsHoja.Cells(4, 6)).Font.Bold = True
    
    xlsHoja.Cells(4, 7) = "Var. %"
    xlsHoja.Range(xlsHoja.Cells(4, 7), xlsHoja.Cells(4, 7)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 7), xlsHoja.Cells(4, 7)).Font.Bold = True
    
    'Cabecera de Pasivos y Patrimonio
    xlsHoja.Cells(4, 8) = "PASIVOS Y PATRIMONIO"
    xlsHoja.Range(xlsHoja.Cells(4, 8), xlsHoja.Cells(4, 8)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(4, 8), xlsHoja.Cells(4, 8)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 8), xlsHoja.Cells(4, 8)).Font.Bold = True
    
    xlsHoja.Cells(4, 9) = "Total " & Left(psFechaInicial, 4) & ""
    xlsHoja.Range(xlsHoja.Cells(4, 9), xlsHoja.Cells(4, 9)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 9), xlsHoja.Cells(4, 9)).Font.Bold = True
    
    xlsHoja.Cells(4, 10) = "%"
    xlsHoja.Range(xlsHoja.Cells(4, 10), xlsHoja.Cells(4, 10)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 10), xlsHoja.Cells(4, 10)).Font.Bold = True
    
    xlsHoja.Cells(4, 11) = "Total " & Left(psFechaFinal, 4) & ""
    xlsHoja.Range(xlsHoja.Cells(4, 11), xlsHoja.Cells(4, 11)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 11), xlsHoja.Cells(4, 11)).Font.Bold = True
    
    xlsHoja.Cells(4, 12) = "%"
    xlsHoja.Range(xlsHoja.Cells(4, 12), xlsHoja.Cells(4, 12)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 12), xlsHoja.Cells(4, 12)).Font.Bold = True
    
    xlsHoja.Cells(4, 13) = "Var. Abs"
    xlsHoja.Range(xlsHoja.Cells(4, 13), xlsHoja.Cells(4, 13)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 13), xlsHoja.Cells(4, 13)).Font.Bold = True
    
    xlsHoja.Cells(4, 14) = "Var. %"
    xlsHoja.Range(xlsHoja.Cells(4, 14), xlsHoja.Cells(4, 14)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 14), xlsHoja.Cells(4, 14)).Font.Bold = True
    
    lnFila = 4
    xlsHoja.Range(xlsHoja.Cells(lnFila, 1), xlsHoja.Cells(lnFila, 14)).Cells.Interior.Color = RGB(238, 248, 255)
    
    'Lines verticales
    xlsHoja.Range(xlsHoja.Cells(4, 7), xlsHoja.Cells(40, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlsHoja.Range(xlsHoja.Cells(4, 3), xlsHoja.Cells(40, 4)).Borders(xlInsideVertical).LineStyle = xlDot
    xlsHoja.Range(xlsHoja.Cells(4, 5), xlsHoja.Cells(40, 6)).Borders(xlInsideVertical).LineStyle = xlDot
    xlsHoja.Range(xlsHoja.Cells(4, 10), xlsHoja.Cells(40, 11)).Borders(xlInsideVertical).LineStyle = xlDot
    xlsHoja.Range(xlsHoja.Cells(4, 12), xlsHoja.Cells(40, 13)).Borders(xlInsideVertical).LineStyle = xlDot
    
    'Lineas Horizontales
    xlsHoja.Range(xlsHoja.Cells(3, 1), xlsHoja.Cells(6, 14)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlsHoja.Range(xlsHoja.Cells(25, 1), xlsHoja.Cells(27, 7)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlsHoja.Range(xlsHoja.Cells(40, 1), xlsHoja.Cells(41, 14)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlsHoja.Range(xlsHoja.Cells(38, 9), xlsHoja.Cells(40, 14)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    'Dobles
    xlsHoja.Range(xlsHoja.Cells(24, 2), xlsHoja.Cells(25, 7)).Borders(xlInsideHorizontal).LineStyle = xlDouble
    xlsHoja.Range(xlsHoja.Cells(36, 2), xlsHoja.Cells(37, 7)).Borders(xlInsideHorizontal).LineStyle = xlDouble
    
    xlsHoja.Range(xlsHoja.Cells(15, 9), xlsHoja.Cells(16, 14)).Borders(xlInsideHorizontal).LineStyle = xlDouble
    xlsHoja.Range(xlsHoja.Cells(27, 9), xlsHoja.Cells(28, 14)).Borders(xlInsideHorizontal).LineStyle = xlDouble
    xlsHoja.Range(xlsHoja.Cells(37, 9), xlsHoja.Cells(38, 14)).Borders(xlInsideHorizontal).LineStyle = xlDouble
    
    'Firma del Contador
    xlsHoja.Cells(48, 1) = "Firma del Contador de la empresa"
    xlsHoja.Range(xlsHoja.Cells(48, 1), xlsHoja.Cells(48, 4)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(48, 1), xlsHoja.Cells(48, 4)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(48, 1), xlsHoja.Cells(48, 4)).Font.Bold = True
    xlsHoja.Range(xlsHoja.Cells(47, 1), xlsHoja.Cells(48, 4)).Borders(xlInsideHorizontal).LineStyle = xlDot
        
    lnFila = 5
    'Conceptos
    If Not ((prsBalanceActivo.EOF And prsBalanceActivo.BOF) And (prsBalancePasivoPat.EOF And prsBalancePasivoPat.BOF)) Then
        For i = 1 To prsBalanceActivo.RecordCount 'Para Activos
                xlsHoja.Cells(lnFila, 1) = prsBalanceActivo!Concepto
                xlsHoja.Range(xlsHoja.Cells(lnFila, 1), xlsHoja.Cells(lnFila, 1)).ColumnWidth = 50
                xlsHoja.Range(xlsHoja.Cells(lnFila, 3), xlsHoja.Cells(lnFila, 3)).ColumnWidth = 6
                xlsHoja.Range(xlsHoja.Cells(lnFila, 5), xlsHoja.Cells(lnFila, 5)).ColumnWidth = 6
                xlsHoja.Range(xlsHoja.Cells(lnFila, 6), xlsHoja.Cells(lnFila, 6)).ColumnWidth = 12
                xlsHoja.Range(xlsHoja.Cells(lnFila, 7), xlsHoja.Cells(lnFila, 7)).ColumnWidth = 6
                
                xlsHoja.Range(xlsHoja.Cells(lnFila, 10), xlsHoja.Cells(lnFila, 10)).ColumnWidth = 6
                xlsHoja.Range(xlsHoja.Cells(lnFila, 12), xlsHoja.Cells(lnFila, 12)).ColumnWidth = 6
                xlsHoja.Range(xlsHoja.Cells(lnFila, 13), xlsHoja.Cells(lnFila, 13)).ColumnWidth = 12
                xlsHoja.Range(xlsHoja.Cells(lnFila, 14), xlsHoja.Cells(lnFila, 14)).ColumnWidth = 6
                
            If (prsBalanceActivo!nConsValor = -1) Or (prsBalanceActivo!nConsValor = -2) Then
                xlsHoja.Cells(lnFila, 2) = ""
                xlsHoja.Cells(lnFila, 3) = ""
                xlsHoja.Range(xlsHoja.Cells(lnFila, 1), xlsHoja.Cells(lnFila, 7)).Cells.Interior.Color = RGB(220, 220, 220)
                xlsHoja.Range(xlsHoja.Cells(lnFila, 1), xlsHoja.Cells(lnFila, 7)).Font.Bold = True
                xlsHoja.Range(xlsHoja.Cells(lnFila, 1), xlsHoja.Cells(lnFila, 7)).Font.Underline = True
            ElseIf (prsBalanceActivo!nConsValor = 100) Or (prsBalanceActivo!nConsValor = 200) Or (prsBalanceActivo!nConsValor = 1000) Then
                xlsHoja.Range(xlsHoja.Cells(lnFila, 1), xlsHoja.Cells(lnFila, 7)).Font.Bold = True
                xlsHoja.Range(xlsHoja.Cells(lnFila, 1), xlsHoja.Cells(lnFila, 7)).HorizontalAlignment = xlRight
                xlsHoja.Range(xlsHoja.Cells(lnFila, 1), xlsHoja.Cells(lnFila, 7)).Cells.Interior.Color = RGB(238, 248, 255)
                
                xlsHoja.Cells(lnFila, 2) = Format(prsBalanceActivo!nMontoIni, "#,##0.00")
                xlsHoja.Cells(lnFila, 3) = Format(prsBalanceActivo!nPorcentajeIni, "#0.00")
                xlsHoja.Range(xlsHoja.Cells(lnFila, 2), xlsHoja.Cells(lnFila, 7)).NumberFormat = "#,##0.00"
                xlsHoja.Cells(lnFila, 4) = Format(prsBalanceActivo!nMontoFin, "#,##0.00")
                xlsHoja.Cells(lnFila, 5) = Format(prsBalanceActivo!nPorcentajeFin, "#0.00")
                
                xlsHoja.Cells(lnFila, 6) = Format(prsBalanceActivo!nVarMonto, "#,##0.00")
                xlsHoja.Cells(lnFila, 7) = Format(prsBalanceActivo!nVarPorcentaje, "#0.00")
            Else
                xlsHoja.Cells(lnFila, 2) = Format(prsBalanceActivo!nMontoIni, "#,##0.00")
                xlsHoja.Cells(lnFila, 3) = Format(prsBalanceActivo!nPorcentajeIni, "#0.00")
                xlsHoja.Range(xlsHoja.Cells(lnFila, 2), xlsHoja.Cells(lnFila, 7)).NumberFormat = "#,##0.00"
                xlsHoja.Cells(lnFila, 4) = Format(prsBalanceActivo!nMontoFin, "#,##0.00")
                xlsHoja.Cells(lnFila, 5) = Format(prsBalanceActivo!nPorcentajeFin, "#0.00")
                
                xlsHoja.Cells(lnFila, 6) = Format(prsBalanceActivo!nVarMonto, "#0.00")
                xlsHoja.Cells(lnFila, 7) = Format(prsBalanceActivo!nVarPorcentaje, "#0.00")
            End If
            lnFila = lnFila + 1
            prsBalanceActivo.MoveNext
        Next i
        
        lnFila = 5
        For i = 1 To prsBalancePasivoPat.RecordCount 'Para Pasivos
                xlsHoja.Cells(lnFila, 8) = prsBalancePasivoPat!Concepto
                xlsHoja.Range(xlsHoja.Cells(lnFila, 8), xlsHoja.Cells(lnFila, 8)).ColumnWidth = 50
            If (prsBalancePasivoPat!nConsValor = -1) Or (prsBalancePasivoPat!nConsValor = -2) Or (prsBalancePasivoPat!nConsValor = -3) Then
                xlsHoja.Cells(lnFila, 9) = ""
                xlsHoja.Cells(lnFila, 10) = ""
                xlsHoja.Range(xlsHoja.Cells(lnFila, 8), xlsHoja.Cells(lnFila, 14)).Font.Bold = True
                xlsHoja.Range(xlsHoja.Cells(lnFila, 8), xlsHoja.Cells(lnFila, 14)).Font.Underline = True
                xlsHoja.Range(xlsHoja.Cells(lnFila, 8), xlsHoja.Cells(lnFila, 14)).Cells.Interior.Color = RGB(220, 220, 220)
            ElseIf (prsBalancePasivoPat!nConsValor = 100) Or (prsBalancePasivoPat!nConsValor = 200) Or (prsBalancePasivoPat!nConsValor = 300) Or (prsBalancePasivoPat!nConsValor = 1000) Or (prsBalancePasivoPat!nConsValor = 1002) Then
                xlsHoja.Range(xlsHoja.Cells(lnFila, 8), xlsHoja.Cells(lnFila, 14)).Font.Bold = True
                xlsHoja.Range(xlsHoja.Cells(lnFila, 8), xlsHoja.Cells(lnFila, 14)).HorizontalAlignment = xlRight
                xlsHoja.Range(xlsHoja.Cells(lnFila, 8), xlsHoja.Cells(lnFila, 14)).Cells.Interior.Color = RGB(238, 248, 255)
                
                xlsHoja.Cells(lnFila, 9) = Format(prsBalancePasivoPat!nMontoIni, "#,##0.00")
                xlsHoja.Cells(lnFila, 10) = Format(prsBalancePasivoPat!nPorcentajeIni, "#0.00")
                xlsHoja.Range(xlsHoja.Cells(lnFila, 9), xlsHoja.Cells(lnFila, 14)).NumberFormat = "#,##0.00"
                xlsHoja.Cells(lnFila, 11) = Format(prsBalancePasivoPat!nMontoFin, "#,##0.00")
                xlsHoja.Cells(lnFila, 12) = Format(prsBalancePasivoPat!nPorcentajeFin, "#0.00")
                
                xlsHoja.Cells(lnFila, 13) = Format(prsBalancePasivoPat!nVarMonto, "#,##0.00")
                xlsHoja.Cells(lnFila, 14) = Format(prsBalancePasivoPat!nVarPorcentaje, "#0.00")
            Else
                xlsHoja.Cells(lnFila, 9) = Format(prsBalancePasivoPat!nMontoIni, "#,##0.00")
                xlsHoja.Cells(lnFila, 10) = Format(prsBalancePasivoPat!nPorcentajeIni, "#0.00")
                xlsHoja.Range(xlsHoja.Cells(lnFila, 9), xlsHoja.Cells(lnFila, 14)).NumberFormat = "#,##0.00"
                xlsHoja.Cells(lnFila, 11) = Format(prsBalancePasivoPat!nMontoFin, "#,##0.00")
                xlsHoja.Cells(lnFila, 12) = Format(prsBalancePasivoPat!nPorcentajeFin, "#0.00")
                
                xlsHoja.Cells(lnFila, 13) = Format(prsBalancePasivoPat!nVarMonto, "#,##0.00")
                xlsHoja.Cells(lnFila, 14) = Format(prsBalancePasivoPat!nVarPorcentaje, "#0.00")
            End If
            lnFila = lnFila + 1
            prsBalancePasivoPat.MoveNext
        Next i

    Else
        MsgBox "Hubo un Error Comuníquese con el Área de TI", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'Autoconfiguracion
    xlsHoja.Cells.Select
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    xlsHoja.Cells.EntireColumn.AutoFit
    
    
End Sub

Public Sub GeneraHojaEstadosResultadosRpt(ByVal psNombre As String, ByVal prsEERR As ADODB.Recordset, ByVal psFechaInicial As String, ByVal psFechaFinal As String, ByRef xlsHoja As Worksheet)
   Dim i As Integer
   Dim lnFila As Integer
    'Configuracion de cabecera
    xlsHoja.Cells(1, 1) = psNombre
    xlsHoja.Range(xlsHoja.Cells(1, 1), xlsHoja.Cells(1, 7)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(1, 1), xlsHoja.Cells(1, 7)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(1, 1), xlsHoja.Cells(1, 7)).Font.Bold = True
    xlsHoja.Range(xlsHoja.Cells(1, 1), xlsHoja.Cells(1, 7)).Cells.Interior.Color = RGB(238, 248, 255)
    
    xlsHoja.Cells(2, 1) = "ESTADOS DE GANACIAS Y PERDIDAS"
    xlsHoja.Range(xlsHoja.Cells(2, 1), xlsHoja.Cells(2, 7)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(2, 1), xlsHoja.Cells(2, 7)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(2, 1), xlsHoja.Cells(2, 7)).Font.Bold = True
    
    xlsHoja.Cells(3, 1) = "Comparativos Anuales (Expresado en soles) del " & Left(psFechaInicial, 4) & " al " & Left(psFechaFinal, 4) & " "
    xlsHoja.Range(xlsHoja.Cells(3, 1), xlsHoja.Cells(3, 7)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(3, 1), xlsHoja.Cells(3, 7)).HorizontalAlignment = xlCenter
    
    'Conceptos
    xlsHoja.Cells(4, 1) = ""
    xlsHoja.Range(xlsHoja.Cells(4, 1), xlsHoja.Cells(4, 1)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(4, 1), xlsHoja.Cells(4, 1)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 1), xlsHoja.Cells(4, 1)).Font.Bold = True
    
    xlsHoja.Cells(4, 2) = "Total " & Left(psFechaInicial, 4) & ""
    xlsHoja.Range(xlsHoja.Cells(4, 2), xlsHoja.Cells(4, 2)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 2), xlsHoja.Cells(4, 2)).Font.Bold = True
    
    xlsHoja.Cells(4, 3) = "%"
    xlsHoja.Range(xlsHoja.Cells(4, 3), xlsHoja.Cells(4, 3)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 3), xlsHoja.Cells(4, 3)).Font.Bold = True
    
    xlsHoja.Cells(4, 4) = "Total " & Left(psFechaFinal, 4) & ""
    xlsHoja.Range(xlsHoja.Cells(4, 4), xlsHoja.Cells(4, 4)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 4), xlsHoja.Cells(4, 4)).Font.Bold = True
    
    xlsHoja.Cells(4, 5) = "%"
    xlsHoja.Range(xlsHoja.Cells(4, 5), xlsHoja.Cells(4, 5)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 5), xlsHoja.Cells(4, 5)).Font.Bold = True
    
    xlsHoja.Cells(4, 6) = "Var. Abs"
    xlsHoja.Range(xlsHoja.Cells(4, 6), xlsHoja.Cells(4, 6)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 6), xlsHoja.Cells(4, 6)).Font.Bold = True
    
    xlsHoja.Cells(4, 7) = "Var. %"
    xlsHoja.Range(xlsHoja.Cells(4, 7), xlsHoja.Cells(4, 7)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(4, 7), xlsHoja.Cells(4, 7)).Font.Bold = True
    
    lnFila = 4
    xlsHoja.Range(xlsHoja.Cells(lnFila, 1), xlsHoja.Cells(lnFila, 7)).Cells.Interior.Color = RGB(238, 248, 255)
    
    'Lines verticales
    xlsHoja.Range(xlsHoja.Cells(4, 7), xlsHoja.Cells(21, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlsHoja.Range(xlsHoja.Cells(4, 3), xlsHoja.Cells(21, 4)).Borders(xlInsideVertical).LineStyle = xlDot
    xlsHoja.Range(xlsHoja.Cells(4, 5), xlsHoja.Cells(21, 6)).Borders(xlInsideVertical).LineStyle = xlDot
    'Lineas Horizontales
    xlsHoja.Range(xlsHoja.Cells(3, 1), xlsHoja.Cells(5, 7)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlsHoja.Range(xlsHoja.Cells(3, 1), xlsHoja.Cells(5, 7)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlsHoja.Range(xlsHoja.Cells(9, 1), xlsHoja.Cells(11, 7)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlsHoja.Range(xlsHoja.Cells(12, 1), xlsHoja.Cells(14, 7)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlsHoja.Range(xlsHoja.Cells(18, 1), xlsHoja.Cells(20, 7)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlsHoja.Range(xlsHoja.Cells(20, 1), xlsHoja.Cells(22, 7)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    'Dobles
    xlsHoja.Range(xlsHoja.Cells(20, 2), xlsHoja.Cells(21, 7)).Borders(xlInsideHorizontal).LineStyle = xlDouble
    lnFila = 5
    'Conceptos
    If Not ((prsEERR.EOF And prsEERR.BOF)) Then
        For i = 1 To prsEERR.RecordCount
                xlsHoja.Cells(lnFila, 1) = prsEERR!Concepto
                xlsHoja.Range(xlsHoja.Cells(lnFila, 1), xlsHoja.Cells(lnFila, 1)).ColumnWidth = 50
                xlsHoja.Range(xlsHoja.Cells(lnFila, 2), xlsHoja.Cells(lnFila, 2)).ColumnWidth = 20
                xlsHoja.Range(xlsHoja.Cells(lnFila, 3), xlsHoja.Cells(lnFila, 3)).ColumnWidth = 10
                xlsHoja.Range(xlsHoja.Cells(lnFila, 4), xlsHoja.Cells(lnFila, 4)).ColumnWidth = 20
                xlsHoja.Range(xlsHoja.Cells(lnFila, 5), xlsHoja.Cells(lnFila, 5)).ColumnWidth = 10
                xlsHoja.Range(xlsHoja.Cells(lnFila, 6), xlsHoja.Cells(lnFila, 6)).ColumnWidth = 18
                xlsHoja.Range(xlsHoja.Cells(lnFila, 7), xlsHoja.Cells(lnFila, 7)).ColumnWidth = 10
                
            If (prsEERR!nConsValor = 10) Or (prsEERR!nConsValor = 20) Or (prsEERR!nConsValor = 30) Or (prsEERR!nConsValor = 50) Or (prsEERR!nConsValor = 60) Then
                xlsHoja.Range(xlsHoja.Cells(lnFila, 1), xlsHoja.Cells(lnFila, 7)).Font.Bold = True
                xlsHoja.Range(xlsHoja.Cells(lnFila, 1), xlsHoja.Cells(lnFila, 7)).HorizontalAlignment = xlRight
                xlsHoja.Range(xlsHoja.Cells(lnFila, 2), xlsHoja.Cells(lnFila, 7)).Cells.Interior.Color = RGB(220, 220, 220)
                
                xlsHoja.Cells(lnFila, 2) = Format(prsEERR!nMontoIni, "#,##0.00")
                xlsHoja.Cells(lnFila, 3) = Format(prsEERR!nPorcentajeIni, "#0.00")
                xlsHoja.Range(xlsHoja.Cells(lnFila, 2), xlsHoja.Cells(lnFila, 7)).NumberFormat = "#,##0.00"
                xlsHoja.Cells(lnFila, 4) = Format(prsEERR!nMontoFin, "#,##0.00")
                xlsHoja.Cells(lnFila, 5) = Format(prsEERR!nPorcentajeFin, "#0.00")
                
                xlsHoja.Cells(lnFila, 6) = Format(prsEERR!nVarMonto, "#,##0.00")
                xlsHoja.Cells(lnFila, 7) = Format(prsEERR!nVarPorcentaje, "#0.00")
            Else
                xlsHoja.Cells(lnFila, 2) = Format(prsEERR!nMontoIni, "#,##0.00")
                xlsHoja.Cells(lnFila, 3) = Format(prsEERR!nPorcentajeIni, "#0.00")
                xlsHoja.Range(xlsHoja.Cells(lnFila, 2), xlsHoja.Cells(lnFila, 7)).NumberFormat = "#,##0.00"
                xlsHoja.Cells(lnFila, 4) = Format(prsEERR!nMontoFin, "#,##0.00")
                xlsHoja.Cells(lnFila, 5) = Format(prsEERR!nPorcentajeFin, "#0.00")
                
                xlsHoja.Cells(lnFila, 6) = Format(prsEERR!nVarMonto, "#0.00")
                xlsHoja.Cells(lnFila, 7) = Format(prsEERR!nVarPorcentaje, "#0.00")
            End If
            lnFila = lnFila + 1
            prsEERR.MoveNext
        Next i
    Else
        MsgBox "Hubo un Error Comuníquese con el Área de TI", vbInformation, "Aviso"
        Exit Sub
    End If
    'Autoconfiguracion
    xlsHoja.Cells.Select
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    xlsHoja.Cells.EntireColumn.AutoFit
    
End Sub
Public Function ValidaDatos() As Boolean
    Dim nIndice As Integer
    Dim i As Integer
    ValidaDatos = False
    Dim lsMensajeIfi As String
    
    'Valida si existe cliente
    If (TxtBCodPers.Text = "") Then
        MsgBox "Nombre del cliente no válido", vbInformation, "Aviso"
        TxtBCodPers.SetFocus
    End If
    
    'Valida parámetros de fecha Ini. y Fin.
    If Right(Format(txtFechaComparaIni.Text, "yyyymmdd"), 4) <> "1231" And Right(Format(txtFechaComparaIni.Text, "yyyymmdd"), 4) <> "0630" Or _
       (txtFechaComparaIni.Text = "__/__/____" Or Not IsDate(Trim(txtFechaComparaIni.Text))) Then
        
        MsgBox "Debe Ingresar fecha semestrales para la comparación" & Chr(10) & " Formato: DD/MM/YYYYY", vbInformation, "Aviso"
        txtFechaComparaIni.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    If Right(Format(txtFechaComparaFin.Text, "yyyymmdd"), 4) <> "1231" And Right(Format(txtFechaComparaFin.Text, "yyyymmdd"), 4) <> "0630" Or _
       (txtFechaComparaFin.Text = "__/__/____" Or Not IsDate(Trim(txtFechaComparaFin.Text))) Then
        MsgBox "Debe Ingresar fecha semestrales para la comparación" & Chr(10) & " Formato: DD/MM/YYYYY", vbInformation, "Aviso"
        txtFechaComparaFin.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    'Valida FechaInicial haya sido registrada
      If Not ValidaExisteFechaEEFF(Format(txtFechaComparaIni.Text, "yyyymmdd")) Then
        MsgBox "Debe ingresar fecha que tenga EEFF registrado", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
      End If
    'Valida FechaFinal haya sido registrada
      If Not ValidaExisteFechaEEFF(Format(txtFechaComparaFin.Text, "yyyymmdd")) Then
        MsgBox "Debe Ingresar fecha que tenga EEFF registrado", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
      End If
    'Valida si los EEFF se encuetran actualizados
    If Not ValidaActualizacionEEFF(sPersCod, Format(txtFechaComparaIni.Text, "yyyymmdd"), Format(txtFechaComparaFin.Text, "yyyymmdd"), lsMensajeIfi) Then
        MsgBox "Los siguientes estados financieros son migrados y necesitan ser actualizados o editados:  " & Chr(10) & " " & lsMensajeIfi & " ", vbInformation, "Aviso"
        Exit Function
    End If
    ValidaDatos = True
End Function
Private Function ValidaExisteFechaEEFF(ByVal psFechaComparar As String) As Boolean
    Dim i As Integer
    ValidaExisteFechaEEFF = False
    For i = 1 To feEstadoFinanciero.rows - 1
        If Trim(Format(feEstadoFinanciero.TextMatrix(i, 1), "yyyymmdd")) = psFechaComparar Then
            ValidaExisteFechaEEFF = True
        End If
    Next
End Function
Private Function ValidaActualizacionEEFF(ByVal psPersCod As String, ByVal psFechaInicial As String, ByVal psFechaFinal As String, Optional ByRef psMensajeIfi As String) As Boolean
    Dim oDFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDFormatosEval = New COMDCredito.DCOMFormatosEval
    Dim rsListaFechas As New ADODB.Recordset
    Dim lsCodIfiMsj As String
        
    Set rsListaFechas = oDFormatosEval.VerificaActualizacionEEFF(psPersCod, psFechaInicial, psFechaFinal)
    psMensajeIfi = ""
    lsCodIfiMsj = ""
    ValidaActualizacionEEFF = True
    
    Do While Not rsListaFechas.EOF
        ValidaActualizacionEEFF = False
            lsCodIfiMsj = (rsListaFechas!dEstFinanFecha) & "  " & lsCodIfiMsj & ""
        rsListaFechas.MoveNext
    Loop

    psMensajeIfi = lsCodIfiMsj
    rsListaFechas.Close
    Set rsListaFechas = Nothing
End Function


'<-*****LUCV20171015 *****
Private Sub Form_Load()
    fraEF.Enabled = False
End Sub
Private Sub TxtBCodPers_EmiteDatos()
    Dim oDPersonaS As COMDPersona.DCOMPersonas
    Dim oRS As ADODB.Recordset
    
    If Trim(TxtBCodPers.Text) = "" Then Exit Sub
    
    fraEF.Enabled = True
    sPersCod = Trim(TxtBCodPers.Text)
    Set oDPersonaS = New COMDPersona.DCOMPersonas
    Set oRS = oDPersonaS.BuscaCliente(sPersCod, BusquedaCodigo)
    Set oDPersonaS = Nothing
    
    If Not oRS.EOF And Not oRS.BOF Then
        lblcodigo(1).Caption = oRS!cPersNombre
        lblcodigo(2).Caption = oRS!cPersIDnroDNI
        lblcodigo(3).Caption = oRS!cPersIDnroRUC
    End If
    Set oRS = Nothing
    Call CargarFlex
End Sub
Private Sub cmdNuevo_Click()
    Call frmPersEstadosFinancierosDetalle.Inicio(0, sPersCod, EFFilaNueva, 0, "")
    Call CargarFlex
End Sub
Private Sub CmdEditar_Click()
    Call MostrarEstadoFinancieroDetalle(EFFilaModificada)
    Call CargarFlex
End Sub
Private Sub cmdConsultar_Click()
    Call MostrarEstadoFinancieroDetalle(EFFilaConsulta)
End Sub
Private Sub MostrarEstadoFinancieroDetalle(ByVal pnTipoAccion As TEFCambios)
    Dim Fila As Integer, nAuditado As Integer, nCodEstFinan As Long
    Dim sFechaEF As String
    Fila = feEstadoFinanciero.row
    If feEstadoFinanciero.TextMatrix(1, 0) = "" Then
        MsgBox ("Debe seleccionar una fila")
        Exit Sub
    End If
    sFechaEF = Trim(feEstadoFinanciero.TextMatrix(Fila, 1))
    nCodEstFinan = CLng(feEstadoFinanciero.TextMatrix(Fila, 5))
    nAuditado = CLng(feEstadoFinanciero.TextMatrix(Fila, 6))
    Select Case pnTipoAccion
        Case 2
            Call frmPersEstadosFinancierosDetalle.Inicio(nCodEstFinan, sPersCod, EFFilaModificada, nAuditado, sFechaEF)
        Case 3
            Call frmPersEstadosFinancierosDetalle.Inicio(nCodEstFinan, sPersCod, EFFilaConsulta, nAuditado, sFechaEF)
    End Select
End Sub
Private Sub cmdEliminar_Click()
    Dim oDPersonaS As COMDPersona.DCOMPersonas
    Dim Fila As Integer, nCodEstFinan As Long
    Dim sFechaEF As String
    Dim lcMovNroEF As String 'EAAS20171103
    Dim oDFormatosEval As COMDCredito.DCOMFormatosEval 'EAAS20171103
    Set oDFormatosEval = New COMDCredito.DCOMFormatosEval 'EAAS20171103
    Fila = feEstadoFinanciero.row
    If feEstadoFinanciero.TextMatrix(1, 0) = "" Then
        MsgBox ("Debe seleccionar una fila")
        Exit Sub
    End If
    nCodEstFinan = CLng(feEstadoFinanciero.TextMatrix(Fila, 5))
    If MsgBox("Desea Eliminar el Estado Financiero Seleccionado?", vbYesNo) = vbYes Then
        Set oDPersonaS = New COMDPersona.DCOMPersonas
        oDPersonaS.EliminarEstFinan (nCodEstFinan)
        Set oDPersonaS = Nothing
        lcMovNroEF = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        Call oDFormatosEval.GrabarEstadoFinancieroEdicion(gsCodPersUser, nCodEstFinan, gdFecSis, nTipoAccion, "", lcMovNroEF, GetMaquinaUsuario)
    End If
    Call CargarFlex
End Sub
Private Sub cmdCancelar_Click()
    TxtBCodPers.Text = ""
    sPersCod = ""
    lblcodigo(1).Caption = ""
    lblcodigo(2).Caption = ""
    lblcodigo(3).Caption = ""
    txtFechaComparaIni.Text = "__/__/____" 'LUCV20171015, Agregó según ERS051-2017
    txtFechaComparaFin.Text = "__/__/____" 'LUCV20171015, Agregó según ERS051-2017
    Call FormateaFlex(feEstadoFinanciero)
End Sub
Private Sub CargarFlex()
    Dim oDPersonaS As COMDPersona.DCOMPersonas
    Dim oRS As ADODB.Recordset
    Dim Fila As Integer
    
    Set oDPersonaS = New COMDPersona.DCOMPersonas
    Set oRS = oDPersonaS.RecuperarDatosPersonaEstadoFinanciero(sPersCod)
    Set oDPersonaS = Nothing
    
    Fila = 0
    Call FormateaFlex(feEstadoFinanciero)
    Do While Not oRS.EOF
        Fila = Fila + 1
        feEstadoFinanciero.AdicionaFila
        feEstadoFinanciero.TextMatrix(Fila, 1) = oRS!dEstFinanFecha
        feEstadoFinanciero.TextMatrix(Fila, 2) = oRS!dEstFinanFechaReg
        feEstadoFinanciero.TextMatrix(Fila, 3) = oRS!cEstFinanUser
        feEstadoFinanciero.TextMatrix(Fila, 4) = oRS!nEstFinanEstado
        feEstadoFinanciero.TextMatrix(Fila, 5) = oRS!nCodEstFinan
        feEstadoFinanciero.TextMatrix(Fila, 6) = oRS!nEstFinanAuditado
        oRS.MoveNext
    Loop
    Set oRS = Nothing
    txtFechaComparaIni.Text = "__/__/____" 'LUCV20171015, Agregó según ERS051-2017
    txtFechaComparaFin.Text = "__/__/____" 'LUCV20171015, Agregó según ERS051-2017
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

