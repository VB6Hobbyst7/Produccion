VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFondoSeguroDep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FONDO SEGURO DEPOSITO  [FSD] ========"
   ClientHeight    =   4560
   ClientLeft      =   3000
   ClientTop       =   1470
   ClientWidth     =   4545
   Icon            =   "frmFondoSeguroDep.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.Animation Logo 
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   393216
      FullWidth       =   41
      FullHeight      =   41
   End
   Begin VB.CommandButton cmdExonerados 
      Caption         =   "&Clientes Exonerados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   180
      TabIndex        =   11
      Top             =   900
      Width           =   4230
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   180
      TabIndex        =   9
      Top             =   3690
      Width           =   4260
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2145
      Left            =   75
      TabIndex        =   5
      Top             =   1380
      Width           =   4350
      Begin VB.CommandButton cmdGeneraAnexo1 
         Caption         =   "Generar &Anexo"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   105
         TabIndex        =   10
         Top             =   1620
         Width           =   4125
      End
      Begin VB.CommandButton cmdReporteTipos 
         Caption         =   "Genera Hojas de Trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   105
         TabIndex        =   8
         Top             =   1170
         Width           =   4125
      End
      Begin VB.CommandButton cmdEstadistica1 
         Caption         =   "Generar &Estadística"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   105
         TabIndex        =   7
         Top             =   720
         Width           =   4125
      End
      Begin VB.CommandButton cmdTiposClientes 
         Caption         =   "Identificar &Tipos de Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   105
         TabIndex        =   6
         Top             =   270
         Width           =   4125
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fechas"
      Enabled         =   0   'False
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
      Height          =   765
      Left            =   840
      TabIndex        =   2
      Top             =   30
      Width           =   3570
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   600
         TabIndex        =   0
         Top             =   292
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecha2 
         Height          =   315
         Left            =   2205
         TabIndex        =   1
         Top             =   285
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "al"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1935
         TabIndex        =   4
         Top             =   345
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "del"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   345
         Width           =   270
      End
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   285
      Left            =   2085
      TabIndex        =   12
      Top             =   4245
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   4185
      Visible         =   0   'False
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFondoSeguroDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rs      As New ADODB.Recordset
Dim sSql    As String
Dim lSalir  As Boolean
Dim sCtaCod As String
Dim lbLibroOpen As Boolean

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim lsMoneda As String
Dim nFSD As Double
Dim nTipAct As Currency
Dim nTipNew As Currency
Dim lAjusteTC As Boolean

'Variables de Anexo
Dim lnAnxImporte_GasS As Currency, lnAnxImporte_GasD As Currency
Dim lnAnxImporte_AhAS As Currency, lnAnxImporte_AhAD As Currency, lnAnxImporte_PFiS As Currency, lnAnxImporte_PFiD As Currency
Dim lnAnxImporte_CTSS As Currency, lnAnxImporte_CTSD As Currency, lnAnxImporte_AhIS As Currency, lnAnxImporte_AhID As Currency

Dim lnContAhoS As Currency, lnContAhoD As Currency
Dim lnContPFS  As Currency, lnContPFD  As Currency

Dim ldFechaDel As Date
Dim ldFechaAl  As Date
Dim oBarra As clsProgressBar
Dim sservidorconsolidada As String

Public Sub Inicio(pdFechaDel As Date, pdFechaAl As Date)
ldFechaDel = pdFechaDel
ldFechaAl = pdFechaAl
Me.Show , frmReportes
End Sub

Private Sub FormatoHoja()
Dim nRow As Integer
Dim N    As Integer
Dim sCol As String
xlAplicacion.Range("A1:K1").EntireColumn.Font.FontStyle = "Arial"
xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.CenterHorizontally = True
xlHoja1.PageSetup.Zoom = 75
xlHoja1.PageSetup.TopMargin = 2
xlHoja1.Range("A2:A2").RowHeight = 17
xlHoja1.Range("A1:A1").ColumnWidth = 12
xlHoja1.Range("A1:A1").EntireColumn.Font.Size = 9
xlHoja1.Range("A1:A1").EntireColumn.ColumnWidth = 20
xlHoja1.Range("A1:A1").EntireColumn.Font.Bold = True
xlHoja1.Cells(1, 1) = "INSTITUCION : " & gsNomCmac
xlHoja1.Range("A1:A1").HorizontalAlignment = xlHAlignLeft

xlHoja1.Cells(3, 1) = "FSD-" & Format(txtFecha2, "mm-yyyy")
xlHoja1.Cells(3, 2) = "Fecha"
xlHoja1.Cells(5, 1) = "CUENTAS DE BALANCE"


nRow = CDate(txtFecha2) - CDate(txtFecha) + 2
xlHoja1.Range("B3:" & ExcelColumnaString(nRow + 2) & "4").Font.Bold = True

For N = 2 To nRow
   xlHoja1.Cells(4, N) = N - 1
   sCol = ExcelColumnaString(N)
   xlHoja1.Range(sCol & "1:" & sCol & "1").ColumnWidth = 13
   xlHoja1.Range(sCol & "1:" & sCol & "1").EntireColumn.NumberFormat = "#,##0.00;-#,##0.00"
   xlHoja1.Range(sCol & "1:" & sCol & "1").EntireColumn.Font.Size = 9
   xlHoja1.Range(sCol & "4:" & sCol & "4").NumberFormat = "#,##0;-#,##0"
   xlHoja1.Range(sCol & "4:" & sCol & "4").HorizontalAlignment = xlHAlignCenter
Next
xlHoja1.Range("B3:" & ExcelColumnaString(nRow) & "3").Merge
xlHoja1.Range("B3:" & ExcelColumnaString(nRow) & "3").HorizontalAlignment = xlHAlignCenter

xlHoja1.Cells(3, nRow + 2) = "PROMEDIO"
sCol = ExcelColumnaString(nRow + 2)
xlHoja1.Range(sCol & "1:" & sCol & "1").ColumnWidth = 13
xlHoja1.Range(sCol & "1:" & sCol & "1").EntireColumn.Font.Size = 9
xlHoja1.Range(sCol & "1:" & sCol & "1").EntireColumn.NumberFormat = "#,##0.00;-#,##0.00"
xlHoja1.Range(sCol & "4:" & sCol & "4").HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("A3:" & sCol & "4").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
xlHoja1.Range("A3:" & sCol & "4").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("B3:" & sCol & "4").Borders(xlInsideHorizontal).LineStyle = xlContinuous

End Sub


 
'Private Sub cmdEstadistica_Click()
'Dim oFSD As New NAnx_FSD
'Dim oCam As New nTipoCambio
'If Not PermiteModificarAsiento(Format(txtFecha2, gsFormatoMovFecha), False) Then
'    MsgBox "No se pueden generar Esdísticas de un mes ya Cerrado", vbInformation, "¡Aviso!"
'    Exit Sub
'End If
'Set oBarra = New clsProgressBar
'oBarra.CaptionSyle = eCap_CaptionPercent
'oBarra.Max = 1
'oBarra.ShowForm Me
'oBarra.Progress 0, "FONDO SEGURO DE DEPOSITOS", "Generando Estadísticas...", , vbBlue
'nTipNew = oCam.EmiteTipoCambio(CDate(txtFecha2) + 1, TCFijoMes)
'Set oCam = Nothing
'
'nFSD = GetMontoFSD()
'If Not oFSD.FSD_GeneraEstadisticas(txtFecha, 0, nFSD, nTipNew, 1) Then
'    cmdGeneraAnexo.Enabled = False
'End If
'Set oFSD = Nothing
'oBarra.Progress 1, "FONDO SEGURO DE DEPOSITOS", "Generando Estadísticas...", , vbBlue
'oBarra.CloseForm Me
'Set oBarra = Nothing
'MsgBox "Estadística generada satisfactoriamente...", vbInformation, "¡Aviso!"
'End Sub

Private Sub cmdEstadistica1_Click()
Dim oFSD As New NAnx_FSD
Dim oCam As New nTipoCambio
If Not PermiteModificarAsiento(Format(txtFecha2, gsFormatoMovFecha), False) Then
    MsgBox "No se pueden generar Esdísticas de un mes ya Cerrado", vbInformation, "¡Aviso!"
    Exit Sub
End If
'Set oBarra = New clsProgressBar
'oBarra.CaptionSyle = eCap_CaptionPercent
'oBarra.Max = 1
'oBarra.ShowForm Me
'oBarra.Progress 0, "FONDO SEGURO DE DEPOSITOS", "Generando Estadísticas...", , vbBlue
nTipNew = oCam.EmiteTipoCambio(CDate(txtFecha2) + 1, TCFijoMes)
Set oCam = Nothing

nFSD = GetMontoFSD()
If Not oFSD.FSD_GeneraEstadisticas(txtFecha, 0, nFSD, nTipNew, 2) Then
    cmdEstadistica1.Enabled = False
End If
Set oFSD = Nothing
'oBarra.Progress 1, "FONDO SEGURO DE DEPOSITOS", "Generando Estadísticas...", , vbBlue
'oBarra.CloseForm Me
'Set oBarra = Nothing
MsgBox "Estadística generada satisfactoriamente...", vbInformation, "¡Aviso!"

End Sub
'
'Private Sub cmdGeneraAnexo_Click()
'Dim lsHoja As String
'Dim lsArchivo As String
'Dim K As Integer
'MousePointer = 11
'Set oBarra = New clsProgressBar
'oBarra.Max = 5
'oBarra.CaptionSyle = eCap_CaptionPercent
'oBarra.ShowForm Me
'oBarra.Progress 0, "FONDO SEGURO DE DEPOSITOS", "Generando Anexo...", , vbBlue
'
'lsArchivo = App.path & "\SPOOLER\Anx17_FSD_" & Year(txtFecha) & ".xls"
'lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
'If Not lbLibroOpen Then
'    MousePointer = 0
'   Exit Sub
'End If
'AbreConexion
'
'lsHoja = "Anx_" & Format(txtFecha2, "mmyyyy")
'ExcelAddHoja lsHoja, xlLibro, xlHoja1
'
'oBarra.Progress 1, "FONDO SEGURO DE DEPOSITOS", "Generando Anexo...", , vbBlue
'xlHoja1.Range("A2:A4").HorizontalAlignment = xlCenter
'xlHoja1.Range("A1").Font.Italic = True
'xlHoja1.Cells(1, 1) = "Superintendencia de Banca y Seguros"
'xlHoja1.Cells(1, 6) = "Anexo No.17-A"
'xlHoja1.Cells(2, 1) = "CONTROL DE IMPOSICIONES CUBIERTAS POR EL"
'xlHoja1.Range("A2:F2").MergeCells = True
'
'xlHoja1.Cells(3, 1) = "FONDO SEGURO DE DEPOSITOS"
'xlHoja1.Range("A3:F3").MergeCells = True
'xlHoja1.Range("A1:F3").Font.Bold = True
'
'xlHoja1.Cells(4, 1) = "PROMEDIO DE SALDOS DIARIOS DEL MES " & Month(txtFecha) & " DEL " & Year(txtFecha)
'xlHoja1.Range("A4:F4").MergeCells = True
'
'
'xlHoja1.Cells(5, 1) = "Empresa: " & gsNomCmac
'xlHoja1.Cells(6, 1) = "TASA DE PRIMA ANUAL: 0.95%"
'
'xlHoja1.Cells(6, 4) = "CATEGORIA"
'xlHoja1.Cells(6, 5) = "SUBCATEGORIA (1)"
'xlHoja1.Cells(6, 6) = "CLASIFICADORA DE RIESGO"
'xlHoja1.Range("D6:F6").HorizontalAlignment = xlCenter
'xlHoja1.Range("D6:F6").VerticalAlignment = xlCenter
'xlHoja1.Range("D6:F6").WrapText = True
'xlHoja1.Range("C6:F6").ColumnWidth = 14
'xlHoja1.Range("D6:F8").Font.Size = 9
'xlHoja1.Range("D6:F8").BorderAround xlContinuous, xlThin
'xlHoja1.Range("D6:F8").Borders(xlInsideHorizontal).LineStyle = xlContinuous
'xlHoja1.Range("D6:F8").Borders(xlInsideVertical).LineStyle = xlContinuous
'
'oBarra.Progress 2, "FONDO SEGURO DE DEPOSITOS", "Generando Anexo...", , vbBlue
'
'xlHoja1.Cells(10, 1) = "CODIGO CTA"
'xlHoja1.Cells(10, 2) = "DENOMINACION"
'xlHoja1.Cells(10, 3) = "MONEDA NACIONAL     (a)"
'xlHoja1.Cells(10, 4) = "MONEDA EXTRANJERA (b) en " & gcMN
'xlHoja1.Cells(10, 5) = "TOTAL              (c) = (a) + (b)"
'xlHoja1.Cells(10, 6) = "%"
'xlHoja1.Range("A10:F10").HorizontalAlignment = xlCenter
'xlHoja1.Range("A10:F10").VerticalAlignment = xlCenter
'xlHoja1.Range("A10:F10").WrapText = True
'xlHoja1.Range("A10").ColumnWidth = 22
'xlHoja1.Range("B10").ColumnWidth = 48
'xlHoja1.Range("A10:F10").BorderAround xlContinuous, xlThin
'xlHoja1.Range("A10:F10").Borders(xlInsideVertical).LineStyle = xlContinuous
'
'xlHoja1.Cells(12, 1) = "A. DEPOSITOS NOMINATIVOS DE PERSONAS Y PERSONAS JURIDICAS PRIVADAS SIN FINES DE LUCRO Y DEPOSITOS"
'xlHoja1.Cells(13, 1) = "   A LA VISTA DE LAS DEMAS PERSONAS JURIDICAS (EXCEPTO DEL SISTEMA FINANCIERO MIEMBROS DEL FONDO)"
'xlHoja1.Range("A12:A13").Font.Bold = True
'
'oBarra.Progress 3, "FONDO SEGURO DE DEPOSITOS", "Generando Anexo...", , vbBlue
'
'xlHoja1.Cells(15, 1) = "'2101.01"
'xlHoja1.Cells(15, 2) = "Depósitos en Cuentas Corrientes"
'
'xlHoja1.Cells(16, 1) = "'2101.02"
'xlHoja1.Cells(16, 2) = "Cuentas Corrientes sin movimiento"
'
'xlHoja1.Cells(17, 1) = "2101.09.02.01"
'xlHoja1.Cells(17, 2) = "Certificados de Depósito No negociables vencidos"
'
'xlHoja1.Cells(18, 1) = "2101.09.03.01"
'xlHoja1.Cells(18, 2) = "Depósitos del Público Vencidos"
'
'xlHoja1.Cells(19, 1) = "2101.13.01"
'xlHoja1.Cells(19, 2) = "Retenciones Judiciales a Disposición"
'
'xlHoja1.Cells(20, 1) = "2101.12.01"
'xlHoja1.Cells(20, 2) = "Depósitos Judiciales y Administrativos"
'
'xlHoja1.Cells(21, 1) = "2108.01+2108.02.01+2108.03.01+2108.07.01"
'xlHoja1.Cells(21, 2) = "Gastos por Pagar de Obligaciones con el Público (5)"
'xlHoja1.Cells(21, 3) = lnAnxImporte_GasS
'xlHoja1.Cells(21, 4) = lnAnxImporte_GasD
'
'xlHoja1.Cells(22, 1) = "2102.01.01"
'xlHoja1.Cells(22, 2) = "Depósitos de Ahorro Activos"
'xlHoja1.Cells(22, 3) = lnAnxImporte_AhAS
'xlHoja1.Cells(22, 4) = lnAnxImporte_AhAD
'
'xlHoja1.Cells(23, 1) = "2103.01.02.01"
'xlHoja1.Cells(23, 2) = "Certificados de Depósito No Negociables"
'
'xlHoja1.Cells(24, 1) = "2103.03.01"
'xlHoja1.Cells(24, 2) = "Cuentas a Plazo"
'xlHoja1.Cells(24, 3) = lnAnxImporte_PFiS
'xlHoja1.Cells(24, 4) = lnAnxImporte_PFiD
'
'xlHoja1.Cells(25, 1) = "2103.09.01"
'xlHoja1.Cells(25, 2) = "Otras Obligaciones por Cuentas a Plazo"
'
'xlHoja1.Cells(26, 1) = "2107.04.01.01+2107.04.02.01+2107.04.09.01"
'xlHoja1.Cells(26, 2) = "Depósitos en Garantía"
'
'xlHoja1.Cells(27, 1) = "2103.04.01"
'xlHoja1.Cells(27, 2) = "Depósitos con  Planes Progresivos"
'
'xlHoja1.Cells(28, 1) = "'2103.05"
'xlHoja1.Cells(28, 2) = "Depósitos CTS"
'xlHoja1.Cells(28, 3) = lnAnxImporte_CTSS
'xlHoja1.Cells(28, 4) = lnAnxImporte_CTSD
'
'xlHoja1.Cells(29, 1) = "2107.01.01+2107.02.01+2107.03.01+2107.09.01"
'xlHoja1.Cells(29, 2) = "Obligaciones con el Público Restringidas (3)"
'
'xlHoja1.Cells(30, 1) = "2102.02.01"
'xlHoja1.Cells(30, 2) = "Depósitos de Ahorro Inactivos"
'xlHoja1.Cells(30, 3) = lnAnxImporte_AhIS
'xlHoja1.Cells(30, 4) = lnAnxImporte_AhID
'xlHoja1.Range("A14:F30").BorderAround xlContinuous, xlThin
'xlHoja1.Range("A14:F30").Borders(xlInsideVertical).LineStyle = xlContinuous
'
'oBarra.Progress 4, "FONDO SEGURO DE DEPOSITOS", "Generando Anexo...", , vbBlue
'xlHoja1.Cells(32, 1) = "B. DEPOSITOS E INSTRUMENTOS FINANCIEROS, DE PERSONAS NATURALES, ASOCIACIONES Y OTRAS PERSONAS"
'xlHoja1.Cells(33, 1) = "   JURIDICAS SON FINES DE LUCRO AMPARADOS POR LE LEGISLACION DEROGADA"
'xlHoja1.Range("A32:A33").Font.Bold = True
'
'xlHoja1.Cells(35, 1) = "2108+2808"
'xlHoja1.Cells(35, 2) = "Otras (5)"
'xlHoja1.Cells(36, 1) = "2103.01.01"
'xlHoja1.Cells(36, 2) = "Certificados de Depósitos Negociables(6)"
'xlHoja1.Cells(37, 1) = "'28"
'xlHoja1.Cells(37, 2) = "Valores en Circulación (7)"
'xlHoja1.Range("A35:F37").BorderAround xlContinuous, xlThin
'xlHoja1.Range("A35:F37").Borders(xlInsideVertical).LineStyle = xlContinuous
'
'For K = 15 To 30
'    If xlHoja1.Cells(K, 3) <> "" Then
'        xlHoja1.Range("E" & K).Formula = "=C" & K & "+D" & K
'        xlHoja1.Range("F" & K).Formula = "=ROUND(E" & K & "/ " & GetSaldoContable(xlHoja1.Cells(K, 1), txtFecha) & ",4 )"
'        xlHoja1.Range("F" & K).Style = "Percent"
'        xlHoja1.Range("F" & K).NumberFormat = "0.00%"
'    End If
'Next
'
'xlHoja1.Cells(39, 1) = "TOTAL SUJETO A COBERTURA"
'xlHoja1.Range("A39").Font.Bold = True
'xlHoja1.Range("C39").Formula = "=C21+C22+C24+C28+C30"
'xlHoja1.Range("D39").Formula = "=D21+D22+D24+D28+D30"
'xlHoja1.Range("E39").Formula = "=E21+E22+E24+E28+E30"
'xlHoja1.Range("A39:F39").BorderAround xlContinuous, xlThin
'xlHoja1.Range("A39:F39").Borders(xlInsideVertical).LineStyle = xlContinuous
'xlHoja1.Range("A39:F39").Font.Bold = True
'xlHoja1.Range("C15:E39").NumberFormat = "##,##0.00#"
'
'xlHoja1.Cells(41, 1) = "Tipo de Cambio : " & nTipNew
'xlHoja1.Range("A41").Font.Bold = True
'ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
'oBarra.Progress 5, "FONDO SEGURO DE DEPOSITOS", "Generando Anexo...", , vbBlue
'oBarra.CloseForm Me
'
'CargaArchivo lsArchivo, App.path & "\SPOOLER"
'CierraConexion
'MousePointer = 0
'
'MsgBox "Anexo FSD generado satisfactoriamente...", vbInformation, "¡Aviso!"
'End Sub

 
 
 

Private Sub GeneraFSD_Contab(psOpeCod As String, psMoneda As String)
Dim lbExisteHoja As Boolean
Dim lsArchivo    As String
Dim lsHoja As String
Dim nRow   As Integer
Dim nCol   As Integer
Dim sCol   As String    ' Columna en Letras
Dim sColT  As String    ' Columna de Totales
Dim rsCta  As New ADODB.Recordset
Dim oFSD As New NAnx_FSD

'On Error GoTo ErrGenera

lsMoneda = psMoneda

If lsMoneda = "1" Then
   gsSimbolo = gcMN
   lsArchivo = App.path & "\SPOOLER\FSD_MN_" & Format(txtFecha2, "mmyyyy") & ".xls"
   nTipAct = 1
   nTipNew = 1
Else
   gsSimbolo = gcME
   lsArchivo = App.path & "\SPOOLER\FSD_ME_" & Format(txtFecha2, "mmyyyy") & ".xls"
   'Calculamos el Tipo de Cambio Fijo
   Dim oCam As New nTipoCambio
   nTipAct = oCam.EmiteTipoCambio(txtFecha2, TCFijoMes)
   nTipNew = oCam.EmiteTipoCambio(CDate(txtFecha2) + 1, TCFijoMes)
   Set oCam = Nothing
   
   Dim oBal    As New NBalanceCont
   Dim nTipBal As Currency
   nTipBal = oBal.GetTipCambioBalance(Format(txtFecha2, gsFormatoMovFecha))
   lAjusteTC = (nTipBal = nTipNew)
   If nTipNew = 0 Then
        nTipNew = nTipAct
   End If
   Set oBal = Nothing
End If

lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
If Not lbLibroOpen Then
   Exit Sub
End If
Dim oOpe As New DOperacion
Set rsCta = oOpe.CargaOpeCta(psOpeCod)
Set oOpe = Nothing
If rsCta.EOF Then
   MsgBox "No se asignó Cuenta Contable a Operación"
   Exit Sub
End If
oBarra.Max = rsCta.RecordCount
oBarra.Progress 0, "FONDO SEGURO DE DEPOSITOS", , "Obteniendo datos...", vbBlue

Do While Not rsCta.EOF
   oBarra.Progress rsCta.Bookmark, "FONDO SEGURO DE DEPOSITOS", , "Procesando Cuenta [ " & rsCta!cCtaContCod & " ]", vbBlue
   If Len(rsCta!cCtaContCod) > 2 Then
      sCtaCod = Left(rsCta!cCtaContCod, 2) & lsMoneda & Mid(rsCta!cCtaContCod, 4, 22)
   Else
      sCtaCod = rsCta!cCtaContCod
   End If
   lsHoja = "Cta_" & sCtaCod
   ExcelAddHoja lsHoja, xlLibro, xlHoja1
   
   Set rs = oFSD.GetFSD_SaldosCont(sCtaCod, txtFecha, txtFecha2)
   If rs.EOF Then
      rs.Close: Set rs = Nothing
      Err.Raise 50001, "frmFondoSeguroDeposito", "No existen datos para generar el FSD de la Cuenta " & sCtaCod
   End If
   nCol = 5
   Dim sCtaDetalle As String
   Dim nRowLlena   As Integer
   Dim nSaldoIni   As Currency
   Dim N As Integer
   'Cabecera de Hoja
   FormatoHoja
   Do While Not rs.EOF
      sCtaDetalle = rs!cCtaContCod
      nSaldoIni = rs!nCtaSaldoImporte
      rs.MoveNext
      If Not rs.EOF Then
         If nSaldoIni <> 0 Or rs!cCtaContCod = sCtaDetalle Then
            nCol = nCol + 1
            nRow = CDate(txtFecha2) - CDate(txtFecha) + 4
            sCol = ExcelColumnaString(nRow)
            xlHoja1.Cells(nCol, 1) = "'" & sCtaDetalle
         End If
         If rs!cCtaContCod = sCtaDetalle Then
            Do While rs!cCtaContCod = sCtaDetalle
               nRow = rs!dCtaSaldofecha - CDate(txtFecha) + 2
               xlHoja1.Cells(nCol, nRow) = rs!nCtaSaldoImporte
               If nRow > 2 And xlHoja1.Cells(nCol, nRow - 1) = "" Then
                  nRowLlena = nRow - 1
                  Do While nRowLlena >= 2
                     If nRowLlena = 2 And xlHoja1.Cells(nCol, nRowLlena) = "" Then
                        If nSaldoIni <> 0 Then
                           For N = nRowLlena To nRow - 1
                              xlHoja1.Cells(nCol, N) = nSaldoIni
                           Next
                        End If
                        Exit Do
                     Else
                        If xlHoja1.Cells(nCol, nRowLlena) <> "" Then
                           For N = nRowLlena + 1 To nRow - 1
                              xlHoja1.Cells(nCol, N) = xlHoja1.Cells(nCol, nRowLlena)
                           Next
                           Exit Do
                        Else
                           nRowLlena = nRowLlena - 1
                        End If
                     End If
                  Loop
               End If
               rs.MoveNext
               If rs.EOF Then
                  nRowLlena = CDate(txtFecha2) - CDate(txtFecha) + 2
                  If nRow < nRowLlena Then
                     For N = nRow + 1 To nRowLlena
                        xlHoja1.Cells(nCol, N) = xlHoja1.Cells(nCol, nRow)
                     Next
                  End If
                  Exit Do
               End If
               If rs!cCtaContCod <> sCtaDetalle Then
                  nRowLlena = CDate(txtFecha2) - CDate(txtFecha) + 2
                  If nRow < nRowLlena Then
                     For N = nRow + 1 To nRowLlena
                        xlHoja1.Cells(nCol, N) = xlHoja1.Cells(nCol, nRow)
                     Next
                  End If
               End If
            Loop
            nRow = CDate(txtFecha2) - CDate(txtFecha) + 4
            sCol = ExcelColumnaString(nRow)
            xlHoja1.Range(sCol & nCol).Formula = "=AVERAGE(B" & nCol & ":" & ExcelColumnaString(nRow - 2) & nCol & ")"
         Else   'No tiene movimiento en el Mes
            If nSaldoIni <> 0 Then
               nRowLlena = CDate(txtFecha2) - CDate(txtFecha) + 2
               For N = 5 To nRowLlena
                  xlHoja1.Cells(nCol, N) = nSaldoIni
               Next
               nRow = CDate(txtFecha2) - CDate(txtFecha) + 4
               sCol = ExcelColumnaString(nRow)
               xlHoja1.Range(sCol & nCol).Formula = "=AVERAGE(B" & nCol & ":" & ExcelColumnaString(nRow - 2) & nCol & ")"
            End If
         End If
      Else
         If nSaldoIni <> 0 Then
            nCol = nCol + 1
            xlHoja1.Cells(nCol, 1) = "'" & sCtaDetalle
            rs.MovePrevious
            If rs!dCtaSaldofecha > CDate(txtFecha) Then
                nRow = rs!dCtaSaldofecha - CDate(txtFecha) + 2
            Else
                nRow = 2
            End If
            rs.MoveNext
            For N = nRow To nRowLlena
               xlHoja1.Cells(nCol, N) = nSaldoIni
            Next
            nRow = CDate(txtFecha2) - CDate(txtFecha) + 4
            sCol = ExcelColumnaString(nRow)
            xlHoja1.Range(sCol & nCol).Formula = "=AVERAGE(B" & nCol & ":" & ExcelColumnaString(nRow - 2) & nCol & ")"
         End If
      End If
   Loop
   RSClose rs

   xlHoja1.Range("A5:A" & nCol).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
   nCol = nCol + 1
   nRowLlena = CDate(txtFecha2) - CDate(txtFecha) + 4
   sColT = ExcelColumnaString(nRowLlena)
   xlHoja1.Cells(nCol, 1) = "TOTAL"
   xlHoja1.Range("A" & nCol & ":" & sColT & nCol).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
   xlHoja1.Range("A" & nCol & ":" & sColT & nCol).Borders(xlInsideVertical).LineStyle = xlContinuous
   'Calculo de Totales por Fila
   xlHoja1.Range("A" & nCol & ":" & sColT & nCol).Font.Bold = True
   For N = 2 To nRowLlena - 2
      sCol = ExcelColumnaString(N)
      xlHoja1.Range(sCol & nCol).Formula = "=SUM(" & sCol & "5" & ":" & sCol & nCol - 1 & ")"
   Next
   nRow = nRowLlena
   xlHoja1.Range(sColT & "5:" & sColT & nCol).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
   xlHoja1.Range(sColT & nCol).Formula = "=AVERAGE(B" & nCol & ":" & sCol & nCol & ")"
   xlHoja1.Range(sColT & "5" & ":" & sColT & nCol).Font.Bold = True
   
   If gsSimbolo = gcME Then
      nCol = nCol + 1
      xlHoja1.Cells(nCol, 1) = "ACTUALIZACION"
      xlHoja1.Range("A" & nCol & ":" & sColT & nCol).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
      For N = nRowLlena - 2 To nRowLlena - 2
         If lAjusteTC Then  'Ya se realiza Ajuste de TC de fin de Mes
            xlHoja1.Cells(nCol - 1, N) = Round(CCur(xlHoja1.Cells(nCol - 1, N)) * nTipAct / nTipNew, 2)
         End If
      Next
      sCol = sColT
      For N = 2 To nRowLlena - 2
         sColT = ExcelColumnaString(N)
         xlHoja1.Range(sColT & nCol).Formula = "=ROUND(" & sColT & nCol - 1 & "* " & nTipNew & "/" & nTipAct & ",2)"
      Next
      xlHoja1.Cells(nCol, nRowLlena) = Round(CCur(xlHoja1.Cells(nCol - 1, nRowLlena)) * nTipNew / nTipAct, 2)
      xlHoja1.Range("B" & nCol & ":" & sCol & nCol).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
      xlHoja1.Range(sCol & nCol).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
   End If
   If Left(sCtaCod, 4) = "2118" Then
        lnAnxImporte_GasS = xlHoja1.Cells(nCol, nRowLlena)
   End If
   If Left(sCtaCod, 4) = "2128" Then
        lnAnxImporte_GasD = xlHoja1.Cells(nCol, nRowLlena)
   End If
   
   Select Case Left(sCtaCod, 6)
      Case "211201", "212201": GeneraFSDExcel nFSD, "232", 0, Mid(sCtaCod, 3, 1), nCol
      Case "211202", "212202": GeneraFSDExcel nFSD, "232", 1, Mid(sCtaCod, 3, 1), nCol
      Case "211303", "212303": GeneraFSDExcel nFSD, "233", 0, Mid(sCtaCod, 3, 1), nCol
      Case "211305", "212305": GeneraFSDExcel nFSD, "234", 0, Mid(sCtaCod, 3, 1), nCol
   End Select
   rsCta.MoveNext
Loop
RSClose rs
RSClose rsCta

   ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
   CargaArchivo lsArchivo, App.path & "\SPOOLER"
   lbLibroOpen = False
Exit Sub
ErrGenera:
   MsgBox TextErr(Err.Description), vbInformation, "Aviso"
   ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
   lbLibroOpen = False
End Sub

Public Sub GeneraFSDExcel(ByVal nFSD As Double, psProd As String, pnActivo As Integer, pnMoneda As Integer, pnCol As Integer)
    Dim rsProd As ADODB.Recordset
    Dim rsCta As ADODB.Recordset
    Dim dbAux As ADODB.Connection
    Dim ldIni As Date, dFechaDia As Date
    Dim I As Integer, J As Integer, K As Integer, N As Integer
    Dim nDias As Integer, nNumTit As Integer
    Dim posx As Integer, posy As Integer
    Dim sProd As String
    Dim nSaldo As Double, nSaldoDia As Double
    Dim sArchivo As String, sCuenta As String
    Dim lbPrimerDia As Boolean
    Dim dFecha As Date
         
    Dim sCol As String, nColIni As Integer, nColFin As Integer, sColFin As String

'On Error GoTo ErrExcel

    posx = pnCol + 1
    dFecha = txtFecha
    nColFin = CDate(Me.txtFecha2) - CDate(Me.txtFecha) + 3
    sColFin = ExcelColumnaString(nColFin)
    
    ldIni = CDate("01" & "/" & Month(dFecha) & "/" & Year(dFecha))
    
        Dim oFSD As New NAnx_FSD
        Set rsProd = oFSD.GetFSD_Estadisticas(txtFecha, pnMoneda, pnActivo, psProd)
        sProd = ""
        If Not rsProd.EOF Then
            xlHoja1.Cells(posx, 1) = "Extracto de Cuenta"
        End If
        Do While Not rsProd.EOF
            sProd = Mid(rsProd!cCodCta, 3, 3)
            ldIni = CDate("01" & "/" & Month(dFecha) & "/" & Year(dFecha))
            sCuenta = rsProd("cCodCta")
            nNumTit = rsProd("nNumTit")

            Set rsCta = oFSD.GetFSD_SaldoCapCta(ldIni, sCuenta, nTipAct)
            
            posy = 2
            xlHoja1.Cells(posx, 1) = "'" & sCuenta
            nSaldo = oFSD.GetSaldoIniFSD(ldIni, sCuenta, nTipAct)
            lbPrimerDia = False
            Do While Not rsCta.EOF
               dFechaDia = rsCta("dFecTran")
               nDias = DateDiff("d", ldIni, dFechaDia)
               
               If Day(dFechaDia) = 1 Then lbPrimerDia = True
               If nDias > 0 Then
                   If Day(ldIni) = 1 And Not lbPrimerDia Then
                       If nSaldo > 0 Then xlHoja1.Cells(posx, posy) = Format$(nSaldo / nNumTit, "#,##0.00")
                       posy = posy + 1
                   End If
                   For J = Day(ldIni) + 1 To nDias + Day(ldIni) - 1
                       If nSaldo > 0 Then xlHoja1.Cells(posx, posy) = Format$(nSaldo / nNumTit, "#,##0.00")
                       posy = posy + 1
                   Next J
               End If
               nSaldoDia = rsCta("nSaldCnt")
               If nSaldoDia > 0 Then xlHoja1.Cells(posx, posy) = Format$(nSaldoDia / nNumTit, "#,##0.00")
               posy = posy + 1
               ldIni = dFechaDia
               nSaldo = nSaldoDia
               rsCta.MoveNext
            Loop
            xlHoja1.Range(sColFin & posx).Formula = "=AVERAGE(B" & posx & ":" & ExcelColumnaString(posy - 1) & posx & ")"
            rsCta.Close
            rsProd.MoveNext
            posx = posx + 1
        Loop
      
      If posx = pnCol + 1 Then
      Else
        nColIni = pnCol + 1
        xlHoja1.Cells(posx, 1) = "TOTAL CUENTAS"
        posy = 2 + CDate(txtFecha2) - CDate(txtFecha)
        For N = 2 To posy
           sCol = ExcelColumnaString(N)
           xlHoja1.Range(sCol & posx).Formula = "=SUM(" & sCol & nColIni & ":" & sCol & posx - 1 & ")"
        Next
        xlHoja1.Range(sColFin & posx).Formula = "=AVERAGE(B" & posx & ":" & ExcelColumnaString(posy) & posx & ")"
        xlHoja1.Range("A" & posx & ":" & sColFin & posx).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
        
        'Actualizacion de ME
        If gsSimbolo = gcME Then
           posx = posx + 1
           xlHoja1.Cells(posx, 1) = "ACT." & nTipNew
           For N = 2 To nColFin
              sCol = ExcelColumnaString(N)
              xlHoja1.Range(sCol & posx).Formula = "=ROUND(" & sCol & posx - 1 & "* " & nTipNew & "/" & nTipAct & ",2)"
           Next
           xlHoja1.Range("A" & posx & ":" & sColFin & posx).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
        End If
        xlHoja1.Range("A" & nColIni & ":" & sColFin & posx).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
        
        posx = posx + 2
        xlHoja1.Cells(posx, 1) = "Saldos Contables"
        xlHoja1.Cells(posx + 1, 1) = "Estados de Cuentas"
        If pnActivo = 0 Then    'Solo Activos
            xlHoja1.Cells(posx + 2, 1) = "Funcionarios"
            xlHoja1.Cells(posx + 4, 1) = "Nro. Personas"
            xlHoja1.Cells(posx + 5, 1) = "Monto de FSD"
            xlHoja1.Cells(posx + 6, 1) = "Total FSD"
            xlHoja1.Cells(posx + 7, 1) = "Monto a Asegurar"
        End If
        
        For N = 2 To nColFin - 1
           sCol = ExcelColumnaString(N)
           'Saldos Contables
           xlHoja1.Range(sCol & posx).Formula = "=" & xlHoja1.Range(sCol & nColIni - 1 & ":" & sCol & nColIni - 1).Address
           'Estados de Cuenta
           xlHoja1.Range(sCol & posx + 1).Formula = "=" & xlHoja1.Range(sCol & posx - 2 & ":" & sCol & posx - 2).Address
           If pnActivo = 0 Then
                'Funcionarios
                xlHoja1.Range(sCol & posx + 2).Formula = "=" & oFSD.GetSaldoFuncionarios(N + CDate(txtFecha) - 2, psProd, pnMoneda, nTipNew)
                If Not psProd = "234" Then
                    'Nro Personas
                    Select Case psProd
                         Case "232" And pnMoneda = 1: xlHoja1.Cells(posx + 4, N) = lnContAhoS
                         Case "232" And pnMoneda = 2: xlHoja1.Cells(posx + 4, N) = lnContAhoD
                         Case "233" And pnMoneda = 1: xlHoja1.Cells(posx + 4, N) = lnContPFS
                         Case "233" And pnMoneda = 2: xlHoja1.Cells(posx + 4, N) = lnContPFD
                    End Select
                    'Monto FSD
                    xlHoja1.Cells(posx + 5, N) = nFSD
                    'Total FSD
                    xlHoja1.Range(sCol & posx + 6).Formula = "=" & xlHoja1.Range(sCol & posx + 4).Address & "*" & xlHoja1.Range(sCol & posx + 5).Address
                End If
                'Monto a Asegurar
                xlHoja1.Range(sCol & posx + 7).Formula = "=" & xlHoja1.Range(sCol & posx).Address & "-" & xlHoja1.Range(sCol & posx + 1).Address & "-" & xlHoja1.Range(sCol & posx + 2).Address & "+" & xlHoja1.Range(sCol & posx + 6).Address
            Else
                'Monto a Asegurar
                xlHoja1.Range(sCol & posx + 2).Formula = "=" & xlHoja1.Range(sCol & posx).Address & "-" & xlHoja1.Range(sCol & posx + 1).Address
            End If
        Next
        If pnActivo = 0 Then
            xlHoja1.Range(sColFin & posx + 7).Formula = "=AVERAGE(B" & posx + 7 & ":" & ExcelColumnaString(posy) & posx + 7 & ")"
            Select Case psProd
                Case "232" And pnMoneda = 1: lnAnxImporte_AhAS = xlHoja1.Range(sColFin & posx + 7).value
                Case "232" And pnMoneda = 2: lnAnxImporte_AhAD = xlHoja1.Range(sColFin & posx + 7).value
                Case "233" And pnMoneda = 1: lnAnxImporte_PFiS = xlHoja1.Range(sColFin & posx + 7).value
                Case "233" And pnMoneda = 2: lnAnxImporte_PFiD = xlHoja1.Range(sColFin & posx + 7).value
                Case "234" And pnMoneda = 1: lnAnxImporte_CTSS = xlHoja1.Range(sColFin & posx + 7).value
                Case "234" And pnMoneda = 2: lnAnxImporte_CTSD = xlHoja1.Range(sColFin & posx + 7).value
            End Select
            xlHoja1.Range("A" & posx & ":" & sColFin & posx + 7).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
            xlHoja1.Range("A" & posx & ":" & sColFin & posx + 7).Font.Bold = True
        Else
            xlHoja1.Range(sColFin & posx + 2).Formula = "=AVERAGE(B" & posx + 2 & ":" & ExcelColumnaString(posy) & posx + 2 & ")"
            xlHoja1.Range("A" & posx & ":" & sColFin & posx + 2).Font.Bold = True

            If pnMoneda = 1 Then
                lnAnxImporte_AhIS = xlHoja1.Range(sColFin & posx + 2).value
            Else
                lnAnxImporte_AhID = xlHoja1.Range(sColFin & posx + 2).value
            End If
            xlHoja1.Range("A" & posx & ":" & sColFin & posx + 2).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
        End If
      End If
      
    RSClose rsProd
    RSClose rsCta
    Exit Sub
ErrExcel:
    MsgBox TextErr(Err.Description), vbExclamation, "¡Aviso!"
    RSClose rsProd
    RSClose rsCta
End Sub

Private Function GetMontoFSD() As Double
Dim oGen As New nCapDefinicion
GetMontoFSD = oGen.GetCapParametro(gMonFSD)
Set oGen = Nothing
End Function

 
Private Sub cmdGeneraAnexo1_Click()
Dim lsHoja As String
Dim lsArchivo As String
Dim K As Integer
MousePointer = 11
Set oBarra = New clsProgressBar
oBarra.Max = 5
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.ShowForm Me
oBarra.Progress 0, "FONDO SEGURO DE DEPOSITOS", "Generando Anexo...", , vbBlue

lsArchivo = App.path & "\SPOOLER\Anx17N_FSD_" & Year(txtFecha) & ".xls"
lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
If Not lbLibroOpen Then
    MousePointer = 0
   Exit Sub
End If

lsHoja = "Anx_" & Format(txtFecha2, "mmyyyy")
ExcelAddHoja lsHoja, xlLibro, xlHoja1

oBarra.Progress 1, "FONDO SEGURO DE DEPOSITOS", "Generando Anexo...", , vbBlue
xlHoja1.Range("A2:A4").HorizontalAlignment = xlCenter
xlHoja1.Range("A2:A4").VerticalAlignment = xlCenter
xlHoja1.Range("A1").Font.Italic = True
xlHoja1.Cells(1, 1) = "Superintendencia de Banca y Seguros"
xlHoja1.Cells(1, 6) = "Anexo No.17-A"
xlHoja1.Cells(2, 1) = "CONTROL DE IMPOSICIONES CUBIERTAS POR EL FONDO SEGURO DE DEPOSITOS"
xlHoja1.Range("A2:F2").MergeCells = True
xlHoja1.Range("A1:F2").Font.Bold = True

xlHoja1.Cells(4, 1) = "PROMEDIO DE SALDOS DIARIOS DEL MES " & Month(txtFecha) & " DEL " & Year(txtFecha)
xlHoja1.Range("A4:F4").MergeCells = True


xlHoja1.Cells(5, 1) = "Empresa: " & gsNomCmac
xlHoja1.Cells(6, 1) = "TASA DE PRIMA ANUAL: 0.95%"

xlHoja1.Cells(6, 4) = "CATEGORIA"
xlHoja1.Cells(7, 4) = "C+"
xlHoja1.Cells(6, 5) = "SUBCATEGORIA (1)"
xlHoja1.Cells(6, 6) = "CLASIFICADORA DE RIESGO"
xlHoja1.Range("D6:I6").HorizontalAlignment = xlCenter
xlHoja1.Range("D6:I6").VerticalAlignment = xlCenter
xlHoja1.Range("D6:I6").WrapText = True
xlHoja1.Range("C6:I6").ColumnWidth = 14
xlHoja1.Range("D6:F8").BorderAround xlContinuous, xlThin
xlHoja1.Range("D6:F8").Borders(xlInsideHorizontal).LineStyle = xlContinuous
xlHoja1.Range("D6:F8").Borders(xlInsideVertical).LineStyle = xlContinuous

oBarra.Progress 2, "FONDO SEGURO DE DEPOSITOS", "Generando Anexo...", , vbBlue

xlHoja1.Cells(9, 4) = "-En Nuevos soles y porcentajes-"

xlHoja1.Cells(11, 1) = "A. DEPÓSITOS NOMINATIVOS DE PERSONAS Y PERSONAS JURIDICAS PRIVADAS SIN FINES DE LUCRO Y DEPOSITOS A LA VISTA DE LAS DEMAS PERSONAS JURÍDICAS"
xlHoja1.Cells(12, 1) = "   (EXCEPTO DEL SISTEMA FINANCIERO MIEMBROS DEL FONDO)"
xlHoja1.Range("A10:A11").Font.Bold = True

xlHoja1.Cells(13, 1) = "CODIGO CTA"
xlHoja1.Cells(13, 2) = "DENOMINACION"
xlHoja1.Cells(13, 3) = "MONEDA NACIONAL     (a)"
xlHoja1.Cells(13, 4) = "MONEDA EXTRANJERA (b) en " & gcMN
xlHoja1.Cells(13, 5) = "TOTAL                  (c) = (a) + (b)"
xlHoja1.Cells(13, 6) = "Nº CUENTAS MONEDA NACIONAL (d)"
xlHoja1.Cells(13, 7) = "Nº CUENTAS MONEDA EXTRANJERA (e)"
xlHoja1.Cells(13, 8) = "TOTAL Nº CUENTAS           (f) = (d) + (e)"
xlHoja1.Cells(13, 9) = "Nº CUENTAS ABIERTAS DURANTE EL MES (2)"
xlHoja1.Range("A13:I13").HorizontalAlignment = xlCenter
xlHoja1.Range("A13:I13").VerticalAlignment = xlCenter
xlHoja1.Range("A13:I13").WrapText = True
xlHoja1.Range("A13").ColumnWidth = 22
xlHoja1.Range("B13").ColumnWidth = 48
xlHoja1.Range("A13:I13").BorderAround xlContinuous, xlThin
xlHoja1.Range("A13:I13").Borders(xlInsideVertical).LineStyle = xlContinuous


oBarra.Progress 3, "FONDO SEGURO DE DEPOSITOS", "Generando Anexo...", , vbBlue

xlHoja1.Cells(15, 1) = "'2101.01"
xlHoja1.Cells(15, 2) = "Depósitos en Cuentas Corrientes"

xlHoja1.Cells(16, 1) = "'2101.02"
xlHoja1.Cells(16, 2) = "Cuentas Corrientes sin movimiento"

xlHoja1.Cells(17, 1) = "2101.09.02.01"
xlHoja1.Cells(17, 2) = "Certificados de Depósito No negociables vencidos"

xlHoja1.Cells(18, 1) = "2101.09.03.01"
xlHoja1.Cells(18, 2) = "Otros Depósitos del Público Vencidos"

xlHoja1.Cells(19, 1) = "2101.13.01"
xlHoja1.Cells(19, 2) = "Retenciones Judiciales a Disposición"

Set rs = GetDatosResumenFSD(gCapAhorros, Me.txtFecha, nFSD, True, False, , , , True)
If Not rs.EOF Then
    xlHoja1.Cells(19, 3) = rs!nMontoMN
    xlHoja1.Cells(19, 4) = rs!nMontoME
    xlHoja1.Cells(19, 5) = rs!nMontoMN + rs!nMontoME
    xlHoja1.Cells(19, 6) = rs!nCuentasMN
    xlHoja1.Cells(19, 7) = rs!nCuentasME
    xlHoja1.Cells(19, 8) = rs!nCuentasMN + rs!nCuentasME
    xlHoja1.Cells(19, 9) = rs!nCuentasApe
End If


xlHoja1.Cells(20, 1) = "2101.12.01"
xlHoja1.Cells(20, 2) = "Depósitos Judiciales y Administrativos"

xlHoja1.Cells(21, 1) = "2108.01+2108.02.01+2108.03.01+2108.07.01"
xlHoja1.Cells(21, 2) = "Gastos por Pagar de Obligaciones con el Público (3)"

'Intereses
'QUITADO A PEDIDO DE AGUSTO BACA
'Set rs = GetDatosResumenFSD(0, Me.txtFecha, nFSD, False, False)
'If Not rs.EOF Then
'    xlHoja1.Cells(21, 3) = rs!nMontoMN
'    xlHoja1.Cells(21, 4) = rs!nMontoME
'    xlHoja1.Cells(21, 5) = rs!nMontoMN + rs!nMontoME
'    xlHoja1.Cells(21, 6) = rs!nCuentasMN
'    xlHoja1.Cells(21, 7) = rs!nCuentasME
'    xlHoja1.Cells(21, 8) = rs!nCuentasMN + rs!nCuentasME
'    xlHoja1.Cells(21, 9) = 0
'End If

xlHoja1.Cells(22, 1) = "2102.01.01"
xlHoja1.Cells(22, 2) = "Depósitos de Ahorro Activos"
'Ahorros Activos
Set rs = GetDatosResumenFSD(gCapAhorros, Me.txtFecha, nFSD, True, False)
If Not rs.EOF Then
    xlHoja1.Cells(22, 3) = rs!nMontoMN
    xlHoja1.Cells(22, 4) = rs!nMontoME
    xlHoja1.Cells(22, 5) = rs!nMontoMN + rs!nMontoME
    xlHoja1.Cells(22, 6) = rs!nCuentasMN
    xlHoja1.Cells(22, 7) = rs!nCuentasME
    xlHoja1.Cells(22, 8) = rs!nCuentasMN + rs!nCuentasME
    xlHoja1.Cells(22, 9) = rs!nCuentasApe
End If

xlHoja1.Cells(23, 1) = "2103.01.02.01"
xlHoja1.Cells(23, 2) = "Certificados de Depósito No Negociables"

xlHoja1.Cells(24, 1) = "2103.03.01"
xlHoja1.Cells(24, 2) = "Cuentas a Plazo"
'Plazo Fijo
Set rs = GetDatosResumenFSD(gCapPlazoFijo, Me.txtFecha, nFSD, True, False)
If Not rs.EOF Then
    xlHoja1.Cells(24, 3) = rs!nMontoMN
    xlHoja1.Cells(24, 4) = rs!nMontoME
    xlHoja1.Cells(24, 5) = rs!nMontoMN + rs!nMontoME
    xlHoja1.Cells(24, 6) = rs!nCuentasMN
    xlHoja1.Cells(24, 7) = rs!nCuentasME
    xlHoja1.Cells(24, 8) = rs!nCuentasMN + rs!nCuentasME
    xlHoja1.Cells(24, 9) = rs!nCuentasApe
End If

xlHoja1.Cells(25, 1) = "2103.09.01"
xlHoja1.Cells(25, 2) = "Otras Obligaciones por Cuentas a Plazo"

xlHoja1.Cells(26, 1) = "2107.04.01.01+2107.04.02.01+2107.04.09.01"
xlHoja1.Cells(26, 2) = "Depósitos en Garantía"

xlHoja1.Cells(27, 1) = "2103.04.01"
xlHoja1.Cells(27, 2) = "Depósitos con  Planes Progresivos"

xlHoja1.Cells(28, 1) = "'2103.05"
xlHoja1.Cells(28, 2) = "Depósitos CTS"
'CTS
Set rs = GetDatosResumenFSD(gCapCTS, Me.txtFecha, nFSD, True, False)
If Not rs.EOF Then
    xlHoja1.Cells(28, 3) = rs!nMontoMN
    xlHoja1.Cells(28, 4) = rs!nMontoME
    xlHoja1.Cells(28, 5) = rs!nMontoMN + rs!nMontoME
    xlHoja1.Cells(28, 6) = rs!nCuentasMN
    xlHoja1.Cells(28, 7) = rs!nCuentasME
    xlHoja1.Cells(28, 8) = rs!nCuentasMN + rs!nCuentasME
    xlHoja1.Cells(28, 9) = rs!nCuentasApe
End If

xlHoja1.Cells(29, 1) = "2107.01.01+2107.02.01+2107.03.01+2107.09.01"
xlHoja1.Cells(29, 2) = "Obligaciones con el Público Restringidas (3)"

xlHoja1.Cells(30, 1) = "2102.02.01"
xlHoja1.Cells(30, 2) = "Depósitos de Ahorro Inactivos"
'Ahorros Inactivos
Set rs = GetDatosResumenFSD(gCapAhorros, Me.txtFecha, nFSD, True, False, 1)
If Not rs.EOF Then
    xlHoja1.Cells(30, 3) = rs!nMontoMN
    xlHoja1.Cells(30, 4) = rs!nMontoME
    xlHoja1.Cells(30, 5) = rs!nMontoMN + rs!nMontoME
    xlHoja1.Cells(30, 6) = rs!nCuentasMN
    xlHoja1.Cells(30, 7) = rs!nCuentasME
    xlHoja1.Cells(30, 8) = rs!nCuentasMN + rs!nCuentasME
    xlHoja1.Cells(30, 9) = rs!nCuentasApe
End If
xlHoja1.Range("A14:I30").BorderAround xlContinuous, xlThin
xlHoja1.Range("A14:I30").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A14:A30").WrapText = True

xlHoja1.Cells(32, 1) = "TOTAL A"
xlHoja1.Range("C32:C32").Formula = "=SUM(" & xlHoja1.Range("C14:C30").Address(False, False) & ")"
xlHoja1.Range("C32:C32").Copy
xlHoja1.Range("D32:I32").PasteSpecial xlPasteFormulas
xlHoja1.Range("A32:I32").BorderAround xlContinuous, xlThin
xlHoja1.Range("B32:I32").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A32:I32").Font.Bold = True


oBarra.Progress 4, "FONDO SEGURO DE DEPOSITOS", "Generando Anexo...", , vbBlue
xlHoja1.Cells(34, 1) = "B. DEPOSITOS E INSTRUMENTOS FINANCIEROS, DE PERSONAS NATURALES, ASOCIACIONES Y OTRAS PERSONAS JURIDICAS SON FINES DE LUCRO AMPARADOS POR LE LEGISLACION "
xlHoja1.Cells(35, 1) = "   DEROGADA"
xlHoja1.Range("A34:A35").Font.Bold = True
xlHoja1.Cells(36, 1) = "CÓDIGO CTA."
xlHoja1.Cells(36, 2) = "DENOMINACION"
xlHoja1.Cells(36, 3) = "MONEDA NACIONAL     (a)"
xlHoja1.Cells(36, 4) = "MONEDA EXTRANJERA en " & gcMN & " (b) "
xlHoja1.Cells(36, 5) = "TOTAL                  (c) = (a) + (b)"
xlHoja1.Cells(36, 6) = "Nº DE INSTRUMENTOS MONEDA NACIONAL (d)"
xlHoja1.Cells(36, 7) = "Nº DE INSTRUMENTOS MONEDA EXTRANJERA (e)"
xlHoja1.Cells(36, 8) = "TOTAL Nº INSTRUMENTOS (f) = (d) + (e)"
xlHoja1.Range("A36:H36").HorizontalAlignment = xlCenter
xlHoja1.Range("A36:H36").VerticalAlignment = xlCenter
xlHoja1.Range("A36:H36").WrapText = True
xlHoja1.Range("A36:H36").BorderAround xlContinuous, xlThin
xlHoja1.Range("B36:H36").Borders(xlInsideVertical).LineStyle = xlContinuous

xlHoja1.Cells(37, 1) = "2108+2808"
xlHoja1.Cells(37, 2) = "Otras (4)"
xlHoja1.Cells(38, 1) = "2103.01.01"
xlHoja1.Cells(38, 2) = "Certificados de Depósitos Negociables"
xlHoja1.Cells(39, 1) = "'28"
xlHoja1.Cells(39, 2) = "Valores en Circulación (5)"
xlHoja1.Range("A37:H39").BorderAround xlContinuous, xlThin
xlHoja1.Range("A37:H39").Borders(xlInsideVertical).LineStyle = xlContinuous

xlHoja1.Cells(41, 1) = "TOTAL B"
xlHoja1.Range("C41:C41").Formula = "=SUM(" & xlHoja1.Range("C37:C39").Address(False, False) & ")"
xlHoja1.Range("C41:C41").Copy
xlHoja1.Range("D41:H41").PasteSpecial xlPasteFormulas
xlHoja1.Range("A41:H41").BorderAround xlContinuous, xlThin
xlHoja1.Range("B41:H41").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A41:H41").Font.Bold = True

xlHoja1.Cells(43, 1) = "TOTAL A + B"
xlHoja1.Range("C43:C43").Formula = "=" & xlHoja1.Range("C32:C32").Address(False, False) & "+" & xlHoja1.Range("C41:C41").Address(False, False)
xlHoja1.Range("C43:C43").Copy
xlHoja1.Range("D43:I43").PasteSpecial xlPasteFormulas
xlHoja1.Range("A43:I43").BorderAround xlContinuous, xlThin
xlHoja1.Range("B43:I43").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A43:I43").Font.Bold = True

'-- Inicio C1
xlHoja1.Cells(45, 1) = "C. MONTO SUJETO A COBERTURA"
xlHoja1.Cells(46, 1) = "DENOMINACION (6)"
xlHoja1.Cells(46, 3) = "MONEDA NACIONAL     (a)"
xlHoja1.Cells(46, 4) = "MONEDA EXTRANJERA (b) en " & gcMN
xlHoja1.Cells(46, 5) = "TOTAL                  (c) = (a) + (b)"
xlHoja1.Cells(46, 6) = "Nº CUENTAS E INSTRUMENTOS MONEDA NACIONAL (d)"
xlHoja1.Cells(46, 7) = "Nº CUENTAS E INSTRUMENTOS MONEDA EXTRANJERA (e)"
xlHoja1.Cells(46, 8) = "TOTAL Nº CUENTAS E INSTRUMENTOS   (f) = (d) + (e)"
xlHoja1.Cells(46, 9) = "%                          (c) / (A + B)"
xlHoja1.Range("A46:I46").HorizontalAlignment = xlCenter
xlHoja1.Range("A46:I46").VerticalAlignment = xlCenter
xlHoja1.Range("A46:I46").WrapText = True
xlHoja1.Range("A46:B46").Merge True
xlHoja1.Range("A46:I46").BorderAround xlContinuous, xlThin
xlHoja1.Range("A46:I46").Borders(xlInsideVertical).LineStyle = xlContinuous

xlHoja1.Cells(47, 1) = "CLIENTES CON INFORMACION TIPO 1"
xlHoja1.Cells(48, 1) = "Depósitos en Cuentas Corrientes"
xlHoja1.Cells(49, 1) = "Cuentas Corrientes sin movimiento"
xlHoja1.Cells(50, 1) = "Certificados de Depósito No negociables vencidos"
xlHoja1.Cells(51, 1) = "Otros Depósitos del Público Vencidos"
xlHoja1.Cells(52, 1) = "Retenciones Judiciales a Disposición"
xlHoja1.Cells(53, 1) = "Depósitos Judiciales y Administrativos"
xlHoja1.Cells(54, 1) = "Depósitos de Ahorro Activos"
'Coberturados Ahorros Activos C1
Set rs = GetDatosResumenFSD(gCapAhorros, Me.txtFecha, nFSD, True, True, , False, 1)
If Not rs.EOF Then
    xlHoja1.Cells(54, 3) = rs!nMontoMN
    xlHoja1.Cells(54, 4) = rs!nMontoME
    xlHoja1.Cells(54, 5) = rs!nMontoMN + rs!nMontoME
    xlHoja1.Cells(54, 6) = rs!nCuentasMN
    xlHoja1.Cells(54, 7) = rs!nCuentasME
    xlHoja1.Cells(54, 8) = rs!nCuentasMN + rs!nCuentasME
    If xlHoja1.Cells(22, 5) <> 0 Then
        xlHoja1.Cells(54, 9) = xlHoja1.Cells(54, 5) / xlHoja1.Cells(22, 5)
    End If
    xlHoja1.Range(xlHoja1.Cells(54, 9), xlHoja1.Cells(54, 9)).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(54, 5), xlHoja1.Cells(54, 5)).Address() & "/" & xlHoja1.Range(xlHoja1.Cells(22, 5), xlHoja1.Cells(22, 5)).Address()
End If

xlHoja1.Cells(55, 1) = "Certificados de Depósito No negociables "
xlHoja1.Cells(56, 1) = "Cuentas a Plazo"
'Coberturados Plazo Fijo C1
Set rs = GetDatosResumenFSD(gCapPlazoFijo, Me.txtFecha, nFSD, True, True, , False, 1)
If Not rs.EOF Then
    xlHoja1.Cells(56, 3) = rs!nMontoMN
    xlHoja1.Cells(56, 4) = rs!nMontoME
    xlHoja1.Cells(56, 5) = rs!nMontoMN + rs!nMontoME
    xlHoja1.Cells(56, 6) = rs!nCuentasMN
    xlHoja1.Cells(56, 7) = rs!nCuentasME
    xlHoja1.Cells(56, 8) = rs!nCuentasMN + rs!nCuentasME
    xlHoja1.Range(xlHoja1.Cells(56, 9), xlHoja1.Cells(56, 9)).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(56, 5), xlHoja1.Cells(56, 5)).Address() & "/" & xlHoja1.Range(xlHoja1.Cells(24, 5), xlHoja1.Cells(24, 5)).Address()
End If

xlHoja1.Cells(57, 1) = "Otras Obligaciones por Cuentas a Plazo"
xlHoja1.Cells(58, 1) = "Depósitos en Garantía"
xlHoja1.Cells(59, 1) = "Depósitos para Planes Progresivos"
xlHoja1.Cells(60, 1) = "Depósitos CTS"
'Coberturados CTS C1
Set rs = GetDatosResumenFSD(gCapCTS, Me.txtFecha, nFSD, True, True, , False, 1)
If Not rs.EOF Then
    xlHoja1.Cells(60, 3) = rs!nMontoMN
    xlHoja1.Cells(60, 4) = rs!nMontoME
    xlHoja1.Cells(60, 5) = rs!nMontoMN + rs!nMontoME
    xlHoja1.Cells(60, 6) = rs!nCuentasMN
    xlHoja1.Cells(60, 7) = rs!nCuentasME
    xlHoja1.Cells(60, 8) = rs!nCuentasMN + rs!nCuentasME
    xlHoja1.Range(xlHoja1.Cells(60, 9), xlHoja1.Cells(60, 9)).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(60, 5), xlHoja1.Cells(60, 5)).Address() & "/" & xlHoja1.Range(xlHoja1.Cells(28, 5), xlHoja1.Cells(28, 5)).Address()
End If

xlHoja1.Cells(61, 1) = "Obligaciones con el Público Restringidas"
xlHoja1.Cells(62, 1) = "Depósitos de Ahorro Inactivas"
'Coberturados Ahorros Inactivos C1
Set rs = GetDatosResumenFSD(gCapAhorros, Me.txtFecha, nFSD, True, True, 1, False, 1)
If Not rs.EOF Then
    xlHoja1.Cells(62, 3) = rs!nMontoMN
    xlHoja1.Cells(62, 4) = rs!nMontoME
    xlHoja1.Cells(62, 5) = rs!nMontoMN + rs!nMontoME
    xlHoja1.Cells(62, 6) = rs!nCuentasMN
    xlHoja1.Cells(62, 7) = rs!nCuentasME
    xlHoja1.Cells(62, 8) = rs!nCuentasMN + rs!nCuentasME
    xlHoja1.Range(xlHoja1.Cells(62, 9), xlHoja1.Cells(62, 9)).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(62, 5), xlHoja1.Cells(62, 5)).Address() & "/" & xlHoja1.Range(xlHoja1.Cells(30, 5), xlHoja1.Cells(30, 5)).Address()
End If
xlHoja1.Cells(63, 1) = "Certificados de Depósito Negociables (7)"
xlHoja1.Cells(64, 1) = "Valores en Circulación (7)"
xlHoja1.Range("A47:I64").BorderAround xlContinuous, xlThin
xlHoja1.Range("B47:I64").Borders(xlInsideVertical).LineStyle = xlContinuous

xlHoja1.Cells(65, 1) = "TOTAL (C1)"
xlHoja1.Range("C65:C65").Formula = "=SUM(" & xlHoja1.Range("C47:C64").Address(False, False) & ")"
xlHoja1.Range("C65:C65").Copy
xlHoja1.Range("D65:H65").PasteSpecial xlPasteFormulas
xlHoja1.Range("A65:H65").BorderAround xlContinuous, xlThin
xlHoja1.Range("B65:H65").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A65:H65").Font.Bold = True

xlHoja1.Range(xlHoja1.Cells(65, 9), xlHoja1.Cells(65, 9)).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(65, 5), xlHoja1.Cells(65, 5)).Address() & "/" & xlHoja1.Range(xlHoja1.Cells(43, 5), xlHoja1.Cells(43, 5)).Address()


'-- Fin C1

'-- Inicio C2
xlHoja1.Cells(66, 1) = "CLIENTES CON INFORMACION TIPO 2"
xlHoja1.Cells(67, 1) = "Depósitos en Cuentas Corrientes"
xlHoja1.Cells(68, 1) = "Cuentas Corrientes sin movimiento"
xlHoja1.Cells(69, 1) = "Certificados de Depósito No negociables vencidos"
xlHoja1.Cells(70, 1) = "Otros Depósitos del Público Vencidos"
xlHoja1.Cells(71, 1) = "Retenciones Judiciales a Disposición"
xlHoja1.Cells(72, 1) = "Depósitos Judiciales y Administrativos"
xlHoja1.Cells(73, 1) = "Depósitos de Ahorro Activos"
'Coberturados Ahorros Activos C2
Set rs = GetDatosResumenFSD(gCapAhorros, Me.txtFecha, nFSD, True, True, , False, 2)
If Not rs.EOF Then
    xlHoja1.Cells(73, 3) = rs!nMontoMN
    xlHoja1.Cells(73, 4) = rs!nMontoME
    xlHoja1.Cells(73, 5) = rs!nMontoMN + rs!nMontoME
    xlHoja1.Cells(73, 6) = rs!nCuentasMN
    xlHoja1.Cells(73, 7) = rs!nCuentasME
    xlHoja1.Cells(73, 8) = rs!nCuentasMN + rs!nCuentasME
    xlHoja1.Range(xlHoja1.Cells(73, 9), xlHoja1.Cells(73, 9)).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(73, 5), xlHoja1.Cells(73, 5)).Address() & "/" & xlHoja1.Range(xlHoja1.Cells(22, 5), xlHoja1.Cells(22, 5)).Address()
End If

xlHoja1.Cells(74, 1) = "Certificados de Depósito No negociables "
xlHoja1.Cells(75, 1) = "Cuentas a Plazo"
'Coberturados Plazo Fijo C2
Set rs = GetDatosResumenFSD(gCapPlazoFijo, Me.txtFecha, nFSD, True, True, , False, 2)
If Not rs.EOF Then
    xlHoja1.Cells(75, 3) = rs!nMontoMN
    xlHoja1.Cells(75, 4) = rs!nMontoME
    xlHoja1.Cells(75, 5) = rs!nMontoMN + rs!nMontoME
    xlHoja1.Cells(75, 6) = rs!nCuentasMN
    xlHoja1.Cells(75, 7) = rs!nCuentasME
    xlHoja1.Cells(75, 8) = rs!nCuentasMN + rs!nCuentasME
    xlHoja1.Range(xlHoja1.Cells(75, 9), xlHoja1.Cells(75, 9)).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(75, 5), xlHoja1.Cells(75, 5)).Address() & "/" & xlHoja1.Range(xlHoja1.Cells(24, 5), xlHoja1.Cells(24, 5)).Address()
End If

xlHoja1.Cells(76, 1) = "Otras Obligaciones por Cuentas a Plazo"
xlHoja1.Cells(77, 1) = "Depósitos en Garantía"
xlHoja1.Cells(78, 1) = "Depósitos para Planes Progresivos"
xlHoja1.Cells(79, 1) = "Depósitos CTS"
'Coberturados CTS C2
Set rs = GetDatosResumenFSD(gCapCTS, Me.txtFecha, nFSD, True, True, , False, 2)
If Not rs.EOF Then
    xlHoja1.Cells(79, 3) = rs!nMontoMN
    xlHoja1.Cells(79, 4) = rs!nMontoME
    xlHoja1.Cells(79, 5) = rs!nMontoMN + rs!nMontoME
    xlHoja1.Cells(79, 6) = rs!nCuentasMN
    xlHoja1.Cells(79, 7) = rs!nCuentasME
    xlHoja1.Cells(79, 8) = rs!nCuentasMN + rs!nCuentasME
    xlHoja1.Range(xlHoja1.Cells(79, 9), xlHoja1.Cells(79, 9)).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(79, 5), xlHoja1.Cells(79, 5)).Address() & "/" & xlHoja1.Range(xlHoja1.Cells(28, 5), xlHoja1.Cells(28, 5)).Address()
End If

xlHoja1.Cells(80, 1) = "Obligaciones con el Público Restringidas"
xlHoja1.Cells(81, 1) = "Depósitos de Ahorro Inactivas"
'Coberturados Ahorros Inactivos C2
Set rs = GetDatosResumenFSD(gCapAhorros, Me.txtFecha, nFSD, True, True, 1, False, 2)
If Not rs.EOF Then
    xlHoja1.Cells(81, 3) = rs!nMontoMN
    xlHoja1.Cells(81, 4) = rs!nMontoME
    xlHoja1.Cells(81, 5) = rs!nMontoMN + rs!nMontoME
    xlHoja1.Cells(81, 6) = rs!nCuentasMN
    xlHoja1.Cells(81, 7) = rs!nCuentasME
    xlHoja1.Cells(81, 8) = rs!nCuentasMN + rs!nCuentasME
    If xlHoja1.Cells(30, 5) Then
        xlHoja1.Cells(81, 9) = xlHoja1.Cells(81, 5) / xlHoja1.Cells(30, 5)
    End If
    xlHoja1.Range(xlHoja1.Cells(81, 9), xlHoja1.Cells(81, 9)).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(81, 5), xlHoja1.Cells(81, 5)).Address() & "/" & xlHoja1.Range(xlHoja1.Cells(30, 5), xlHoja1.Cells(30, 5)).Address()
End If
xlHoja1.Cells(82, 1) = "Certificados de Depósito Negociables (7)"
xlHoja1.Cells(83, 1) = "Valores en Circulación (7)"
xlHoja1.Range("A67:I83").BorderAround xlContinuous, xlThin
xlHoja1.Range("B67:I83").Borders(xlInsideVertical).LineStyle = xlContinuous

xlHoja1.Cells(84, 1) = "TOTAL (C2)"
xlHoja1.Range("C84:C84").Formula = "=SUM(" & xlHoja1.Range("C67:C83").Address(False, False) & ")"
xlHoja1.Range("C84:C84").Copy
xlHoja1.Range("D84:H84").PasteSpecial xlPasteFormulas
xlHoja1.Range("A84:H84").BorderAround xlContinuous, xlThin
xlHoja1.Range("B84:H84").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A84:H84").Font.Bold = True

xlHoja1.Range(xlHoja1.Cells(84, 9), xlHoja1.Cells(84, 9)).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(84, 5), xlHoja1.Cells(84, 5)).Address() & "/" & xlHoja1.Range(xlHoja1.Cells(43, 5), xlHoja1.Cells(43, 5)).Address()

'--Fin C2

'C1 + C2
xlHoja1.Cells(86, 1) = "TOTAL C (C1 + C2)"
xlHoja1.Range("C86:C86").Formula = "=" & xlHoja1.Range("C65:C65").Address(False, False) & "+" & xlHoja1.Range("C84:C84").Address(False, False)
xlHoja1.Range("C86:C86").Copy
xlHoja1.Range("D86:I86").PasteSpecial xlPasteFormulas
xlHoja1.Range("A86:I86").BorderAround xlContinuous, xlThin
xlHoja1.Range("B86:I86").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A86:I86").Font.Bold = True

'D. NO SUJETO A COBERTURA
xlHoja1.Cells(88, 1) = "D. MONTO NO SUJETO A COBERTURA"
xlHoja1.Cells(89, 3) = "MONEDA NACIONAL     (a)"
xlHoja1.Cells(89, 4) = "MONEDA EXTRANJERA (b) en " & gcMN
xlHoja1.Cells(89, 5) = "TOTAL                  (c) = (a) + (b)"
xlHoja1.Cells(89, 6) = "%                          (c) / (A + B)"
xlHoja1.Range("A89:F89").HorizontalAlignment = xlCenter
xlHoja1.Range("A89:F89").VerticalAlignment = xlCenter
xlHoja1.Range("A89:F89").WrapText = True
xlHoja1.Range("A89:B89").Merge True
xlHoja1.Range("A89:B90").Font.Bold = True
xlHoja1.Range("A89:F89").BorderAround xlContinuous, xlThin
xlHoja1.Range("A90:F90").BorderAround xlContinuous, xlThin
xlHoja1.Range("B89:F90").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Cells(90, 1) = "TOTAL D"
xlHoja1.Range("C90:C90").Formula = "=" & xlHoja1.Range("C43:C43").Address(False, False) & "-" & xlHoja1.Range("C86:C86").Address(False, False)
xlHoja1.Range("D90:D90").Formula = "=" & xlHoja1.Range("D43:D43").Address(False, False) & "-" & xlHoja1.Range("D86:D86").Address(False, False)
xlHoja1.Range("E90:E90").Formula = "=" & xlHoja1.Range("C90:C90").Address(False, False) & "+" & xlHoja1.Range("D90:D90").Address(False, False)
xlHoja1.Range("F90:F90").Formula = "=" & xlHoja1.Range("E90:E90").Address(False, False) & "/" & xlHoja1.Range("E43:E43").Address(False, False)

xlHoja1.Range("C6:I90").Font.Size = 9
xlHoja1.Range("C14:E90").NumberFormat = "#,##0.00"
xlHoja1.Range("F14:I90").NumberFormat = "#,##0"
xlHoja1.Range("I47:I90").NumberFormat = "#,##0.00%"
xlHoja1.Range("F90:F90").NumberFormat = "#,##0.00%"

xlHoja1.Cells(21, 3) = 0
xlHoja1.Cells(21, 4) = 0
xlHoja1.Cells(21, 5) = 0
xlHoja1.Cells(21, 6) = 0
xlHoja1.Cells(21, 7) = 0
xlHoja1.Cells(21, 8) = 0
xlHoja1.Cells(21, 9) = 0


'xddd
Dim rsGaranS As New ADODB.Recordset
Dim rsGaranD As New ADODB.Recordset

Set rsGaranS = GetDatosGarantia(CDate(txtFecha2), gnTipCambio, 1, False)
Set rsGaranD = GetDatosGarantia(CDate(txtFecha2), gnTipCambio, 2, False)

xlHoja1.Cells(26, 3) = rsGaranS!Saldo
xlHoja1.Cells(26, 4) = Round(rsGaranD!Saldo * gnTipCambio, 2)
xlHoja1.Cells(26, 5) = rsGaranS!Saldo + Round(rsGaranD!Saldo * gnTipCambio, 2)
xlHoja1.Cells(26, 6) = rsGaranS!Numero
xlHoja1.Cells(26, 7) = rsGaranD!Numero
xlHoja1.Cells(26, 8) = rsGaranS!Numero + rsGaranD!Numero
xlHoja1.Cells(26, 9) = rsGaranS!NroAper + rsGaranD!NroAper

xlHoja1.Cells(29, 3) = 0
xlHoja1.Cells(29, 4) = 0
xlHoja1.Cells(29, 5) = 0
xlHoja1.Cells(29, 6) = 0
xlHoja1.Cells(29, 7) = 0
xlHoja1.Cells(29, 8) = 0
xlHoja1.Cells(29, 9) = 0

Set rsGaranS = GetDatosGarantia(CDate(txtFecha2), gnTipCambio, 1, True)
Set rsGaranD = GetDatosGarantia(CDate(txtFecha2), gnTipCambio, 2, True)

xlHoja1.Cells(58, 3) = rsGaranS!Saldo
xlHoja1.Cells(58, 4) = Round(rsGaranD!Saldo * gnTipCambio, 2)
xlHoja1.Cells(58, 5) = rsGaranS!Saldo + Round(rsGaranD!Saldo * gnTipCambio, 2)
xlHoja1.Cells(58, 6) = rsGaranS!Numero
xlHoja1.Cells(58, 7) = rsGaranD!Numero
xlHoja1.Cells(58, 8) = rsGaranS!Numero + rsGaranD!Numero
xlHoja1.Cells(58, 9) = "=$E$58/$E$26"


xlHoja1.Cells(92, 1) = "Tipo de Cambio:"
xlHoja1.Cells(92, 2) = Format(gnTipCambio, "0.000")


oBarra.Progress 5, "FONDO SEGURO DE DEPOSITOS", "Generando Anexo...", , vbBlue
oBarra.CloseForm Me
ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
MsgBox "Anexo FSD generado satisfactoriamente...", vbInformation, "¡Aviso!"

CargaArchivo lsArchivo, App.path & "\SPOOLER"
CierraConexion
MousePointer = 0

End Sub

Private Sub cmdReporteTipos_Click()
Dim nFSD As Double
Dim rs1 As New ADODB.Recordset
Dim lsArchivo As String
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim E As Integer
Dim oCam As New nTipoCambio
Dim tProd As String
Dim tHoja As String
Dim nFin As Integer

prg.Visible = True
Status.Visible = True

nTipNew = oCam.EmiteTipoCambio(CDate(IIf(Val(Mid(txtFecha, 4, 2)) < 12, IIf(Val(Mid(txtFecha, 4, 2)) + 1 < 10, "0" & Trim(Str(Val(Mid(txtFecha, 4, 2)) + 1)), Trim(Str(Val(Mid(txtFecha, 4, 2)) + 1))) & "/" & Trim(Str(Mid(txtFecha, 7, 4))), "01/" & Trim(Str(Val(Mid(txtFecha, 7, 4)) + 1)))), TCFijoMes)
Set oCam = Nothing


nFSD = GetMontoFSD()

For E = 1 To 2 '1 Sin Exonerados 2 Con Exonerados
    If E = 1 Then 'Sin Exonerados es lo normal
        nFin = 2 'Los 2 tipos  de informacion
    Else
        nFin = 1 'Solo se recorrera una vez por moneda
    End If
        
    For I = 1 To nFin 'Tipo ' Tipo de Informacion 1 o 2 Nota: Los 2 estan en la tabla FSDPersonaClasif con cPersTipo=2
                            ' Los Que son Con Exonerados (2) no se sacaran por tipo de informacion
        For J = 1 To 2 'Moneda
            If E = 1 Then 'Sin Exonerados
                lsArchivo = App.path & "\Spooler\" & "RptFSD" & "_SinExo_" & "Tipo" & I & "_" & IIf(J = 1, "MN", "ME") & "_" & Mid(txtFecha, 4, 2) & Mid(txtFecha, 7, 4) & ".XLS"
            Else 'Incluir Exonerados
                lsArchivo = App.path & "\Spooler\" & "RptFSD" & "_IncExo_" & IIf(J = 1, "MN", "ME") & "_" & Mid(txtFecha, 4, 2) & Mid(txtFecha, 7, 4) & ".XLS"
            End If
            ExcelBegin lsArchivo, xlAplicacion, xlLibro, True
            For K = 1 To 4
                If K = 1 Then
                    tProd = ""
                    tHoja = "Interes"
                    Genera_ReporteTipos tProd, 0, J, nFSD, I, E, xlLibro, tHoja
                ElseIf K = 2 Then
                    tProd = "234"
                    tHoja = "CTS"
                    Genera_ReporteTipos tProd, 0, J, nFSD, I, E, xlLibro, tHoja
                ElseIf K = 3 Then
                    tProd = "233"
                    tHoja = "Plazo Fijo"
                    Genera_ReporteTipos tProd, 0, J, nFSD, I, E, xlLibro, tHoja
                ElseIf K = 4 Then
                    tProd = "232"
                    tHoja = "Ahorro Inactivo"
                    Genera_ReporteTipos tProd, 1, J, nFSD, I, E, xlLibro, tHoja
                    tHoja = "Ahorro Activo"
                    Genera_ReporteTipos tProd, 0, J, nFSD, I, E, xlLibro, tHoja
                End If
                    
            Next
            ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
        Next
    Next
Next
      
''Tipo1 y 2
''Soles y Dolares
''Ahorros Inactiva y Activa
''
'CargaArchivo "ReporteFSD_MN_" & Mid(txtFecha, 4, 2) & Mid(txtFecha, 7, 4) & ".XLS", App.path & "\Spooler"
''CargaArchivo "ReporteFSD_ME_" & Mid(txtFecha, 4, 2) & Mid(txtFecha, 7, 4) & ".XLS", App.path & "\Spooler"

prg.Visible = False
Status.Visible = False

MsgBox "Archivos generados con exito", vbInformation, "Aviso"
End Sub
Private Sub Genera_ReporteTipos(psProd As String, pnActivo As Integer, pnMoneda As Integer, pnFSD As Double, pnTipoInformacion As Integer, pnExonerados As Integer, xlLibro As Excel.Workbook, lsHoja As String)
Dim oAnx As NAnx_FSD
Dim oFSD As New NAnx_FSD
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim nColFin As Integer
Dim nCol As Integer
Dim nFil As Long
Dim I As Integer
Dim nPos1 As Integer
Dim nTemporal As String
Dim nMontoDis As Currency
Dim nMontoCob As Currency

Dim nTemp As String
Dim nColTemp As Integer
Dim nTemp1 As String

Dim lsArchivo    As String
Dim lbExcel As Boolean
Dim nExiste As Integer 'nExiste = 1 si existe algun registro
                       'nExiste = 2 si no existe algun registro

    
    Status.Panels(1).Text = "Calculando " & lsHoja & "..."
    prg.Min = 0
    
    nTemp = ""
    nTemp1 = ""
    nPos1 = 0
    nCol = 0
    nColFin = 0
    nFil = 0
    nTemporal = ""
    nExiste = 2
 
    ExcelAddHoja lsHoja, xlLibro, xlHoja1, True

    xlHoja1.Cells(1, 1) = "Institución: " & gsNomCmac
    xlHoja1.Range("A1:H1").MergeCells = True
    xlHoja1.Cells(2, 1) = "CLIENTES CON INFORMACION TIPO " & pnTipoInformacion
    xlHoja1.Range("A2:H2").MergeCells = True
    xlHoja1.Cells(3, 1) = "FSD-" & Mid(txtFecha, 4, 2) & "-" & Mid(txtFecha, 7, 4)
    xlHoja1.Range("A3:H3").MergeCells = True
    nColFin = Val(Mid(DateAdd("d", -1, "01/" & IIf(Val(Mid(txtFecha, 4, 2)) < 12, IIf(Val(Mid(txtFecha, 4, 2)) + 1 < 10, "0" & Trim(Str(Val(Mid(txtFecha, 4, 2)) + 1)), Trim(Str(Val(Mid(txtFecha, 4, 2)) + 1))) & "/" & Trim(Str(Mid(txtFecha, 7, 4))), "01/" & Trim(Str(Val(Mid(txtFecha, 7, 4)) + 1)))), 1, 2))
    nColFin = nColFin + 1
    
    nFil = 4
    For I = 1 To nColFin - 1
        xlHoja1.Cells(nFil, I + 1) = I
    Next
    
    xlHoja1.Cells(nFil, nColFin + 1) = "Prom. Calc."
    xlHoja1.Cells(nFil, nColFin + 2) = "Prom.  B.D."
    xlHoja1.Cells(nFil, nColFin + 3) = "Coberturado"
    
    If nColFin + 97 >= 123 Then
        nTemp = "A" & UCase(Chr(nColFin + 97 - 26))
        nTemp1 = "A" & UCase(Chr(nColFin + 97 - 26 - 1))
    Else
        nTemp = UCase(Chr(nColFin + 97))
        nTemp1 = UCase(Chr(nColFin + 97 - 1))
    End If
     
    xlHoja1.Range("A" & nFil, nTemp & (Trim(Str(nFil)))).Font.Bold = True
    
    ExcelCuadro xlHoja1, 2, Val(nFil), nColFin + 1, Val(nFil)
    
    'Primera Parte
    
    nFil = nFil + 1
    xlHoja1.Cells(nFil, 1) = "CUENTAS DE PERSONAS NATURALES Y PERSONAS JURIDICAS SIN FINES DE LUCRO"
    xlHoja1.Range("A" & nFil, "H" & nFil).MergeCells = True
    xlHoja1.Range("A" & nFil, "A" & nFil).Font.Bold = True
    
    
    Set rs1 = oFSD.GetFSD_EstadisticasNew(txtFecha.Text, pnMoneda, pnActivo, psProd, pnTipoInformacion, pnExonerados, IIf(Len(Trim(psProd)) = 0, 2, 1))
    prg.Max = IIf(rs1.RecordCount = 0, 1, rs1.RecordCount)
    prg.value = 0
    
    If rs1.BOF Then
        nExiste = 2
        nPos1 = 0
    Else
        nExiste = 1
        nPos1 = nFil + 1
    End If
    
    Do While Not rs1.EOF
    

        prg.value = rs1.Bookmark
        '''
    
        If rs1!cCodCta <> nTemporal Then
            If Len(Trim(nTemporal)) > 0 Then
               'Calcula el Promedio
                xlHoja1.Range(nTemp & nFil & ":" & nTemp & nFil).Formula = "=AVERAGE(" & "B" & nFil & ":" & nTemp1 & nFil & ")"
              
                'Imprimo las dos columnas ultimas Promedio y Capital Promedio o Interes Promedio
                xlHoja1.Cells(nFil, nColFin + 2) = nMontoDis
                xlHoja1.Cells(nFil, nColFin + 3) = nMontoCob
            End If
                
            nFil = nFil + 1
            xlHoja1.Cells(nFil, 1) = rs1!cCodCta
            nTemporal = rs1!cCodCta
            nMontoDis = rs1!nMontoDis
            nMontoCob = rs1!nMontoCob
        End If
         
        ' IMPRIME SALDO EN LA FILA nFil y Columna dia de la fecha +1
        
        xlHoja1.Cells(nFil, Val(Mid(rs1!cFecha, 7, 2)) + 1) = rs1!nMonto
        
        'Me.Caption = lsHoja & " [Parte A]: " & Format(rs1.AbsolutePosition / 30, "0") & " / " & Format(rs1.RecordCount / 30, "0")
        
        rs1.MoveNext
        
    Loop
    
    If nExiste = 1 Then
    
        If Len(Trim(nTemporal)) > 0 Then
            xlHoja1.Range(nTemp & nFil & ":" & nTemp & nFil).Formula = "=AVERAGE(" & "B" & nFil & ":" & nTemp1 & nFil & ")"
            xlHoja1.Cells(nFil, nColFin + 2) = nMontoDis
            xlHoja1.Cells(nFil, nColFin + 3) = nMontoCob
        End If
        
        rs1.Close
        Set rs1 = Nothing
        
        ExcelCuadro xlHoja1, 1, CCur(nPos1), nColFin + 1, CCur(nFil)
        
        nFil = nFil + 1
        
        xlHoja1.Cells(nFil, 1) = "TOTAL CUENTAS"
        
        nTemp = ""
        For I = 2 To nColFin + 3
            nColTemp = I + 96
            nTemp = ""
            If nColTemp >= 123 Then
                'siguiente Letra a Z
                nTemp = "A" & UCase(Chr(nColTemp - 26))
            Else
                nTemp = UCase(Chr(nColTemp))
            End If
            xlHoja1.Range(nTemp & Trim(Str(nFil)), nTemp & Trim(Str(nFil))).Formula = "=SUM(" & nTemp & Trim(Str(nPos1)) & ":" & nTemp & Trim(Str(nFil - 1)) & ")"
        Next
        xlHoja1.Range("A" & nFil, nTemp & (Trim(Str(nFil)))).Font.Bold = True
        ExcelCuadro xlHoja1, 1, Val(nFil), nColFin + 1, Val(nFil)
    Else
        nFil = nFil + 1
        xlHoja1.Cells(nFil, 1) = "TOTAL CUENTAS"
        ExcelCuadro xlHoja1, 1, Val(nFil), nColFin + 1, Val(nFil)
    End If
        
    xlHoja1.Cells(nFil + 2, 1) = "Monto FSD:"
    xlHoja1.Cells(nFil + 2, 2) = pnFSD
    
    
    xlHoja1.PageSetup.Zoom = 80
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.Cells.Select
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 8
    xlHoja1.Cells.NumberFormat = "#,###,##0.00"
    xlHoja1.Range("B4:" & nTemp & "4").NumberFormat = "0"
    xlHoja1.Cells.EntireColumn.AutoFit
   
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdTiposClientes_Click()
    frmFSDTiposClientes.Show 1
End Sub

Private Sub cmdExonerados_Click()
    frmFSDClientesExonerados.Show 1
End Sub

Private Sub Form_Load()
CentraForm Me
frmReportes.Enabled = False
txtFecha = ldFechaDel
txtFecha2 = ldFechaAl
lSalir = False
Me.cmdGeneraAnexo1.Enabled = True
If Dir(App.path & "\videos\LogoA.avi") <> "" Then
    Logo.AutoPlay = True
    Logo.Open App.path & "\videos\LogoA.avi"
End If
Dim oConst As NConstSistemas
Set oConst = New NConstSistemas

sservidorconsolidada = oConst.LeeConstSistema(gConstSistServCentralRiesgos)
Set oConst = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmReportes.Enabled = True
End Sub

Private Sub txtFecha_GotFocus()
txtFecha.SelStart = 0
txtFecha.SelLength = Len(txtFecha)
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFecha) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Aviso"
      Exit Sub
   End If
   txtFecha2.SetFocus
End If
End Sub

Private Sub txtFecha_Validate(Cancel As Boolean)
If ValidaFecha(txtFecha) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "Aviso"
   Cancel = True
End If
End Sub

Private Sub txtFecha2_GotFocus()
txtFecha2.SelStart = 0
txtFecha2.SelLength = Len(txtFecha2)
End Sub

Private Sub txtFecha2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFecha2) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Aviso"
      Exit Sub
   End If
   If CDate(txtFecha2) < CDate(txtFecha) Then
      MsgBox "Fecha no puede ser menor a fecha Inicial"
      Exit Sub
   End If
   'cmdGenerar.SetFocus
End If
End Sub

Private Sub txtFecha2_Validate(Cancel As Boolean)
If ValidaFecha(txtFecha2) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "Aviso"
   Cancel = True
   Exit Sub
End If
If CDate(txtFecha2) < CDate(txtFecha) Then
   MsgBox "Fecha no puede ser menor a fecha Inicial"
   Cancel = True
End If
End Sub

Private Function GetSaldoContable(psCta As String, pdFecha As Date) As Currency
Dim lnSaldoBala As Currency
Dim lsCta As String
Dim oBala As New NBalanceCont
lnSaldoBala = 0
psCta = Replace(psCta, ".", "")
Do While psCta <> ""
    If InStr(psCta, "+") > 0 Then
        lsCta = Left(psCta, InStr(psCta, "+") - 1)
        psCta = Mid(psCta, InStr(psCta, "+") + 1)
    Else
        lsCta = psCta
        psCta = ""
    End If
    lnSaldoBala = lnSaldoBala + oBala.getImporteBalanceMes(lsCta, 2, 0, Month(pdFecha), Year(pdFecha))
Loop
GetSaldoContable = lnSaldoBala

End Function


Private Function GetDatosResumenFSD(psProd As Producto, pdFecha As Date, pnMontoFSD As Double, pbCapital As Boolean, pbCoberturado As Boolean, Optional pnInactiva As Integer = 0, Optional pbGetAperturas As Boolean = True, Optional pnClienteTipo As Integer = 0, Optional pbRetencionesJudiciales As Boolean = False) As ADODB.Recordset
Dim lsSql  As String
Dim lsAper As String
Dim oCon   As New DConecta
Dim lsFiltroClte As String
Dim sServidorFSD As String
Dim sBaseFSD As String
Dim oConst As New NConstSistemas
sServidorFSD = oConst.LeeConstSistema(gConstSistServFSD)
sBaseFSD = sServidorFSD 'Mid(sServidorFSD, InStr(sServidorFSD, "]") + 2)
Set oConst = Nothing

If gbBitCentral Then
    
    
    If pnClienteTipo > 0 Then
        lsFiltroClte = " and cPersCod "
        If pnClienteTipo = 1 Then
           lsFiltroClte = lsFiltroClte & " NOT "
        End If
        lsFiltroClte = lsFiltroClte & " IN (Select cPersCod FROM " & sBaseFSD & "FSDPersonaClasif) "
    End If
    Select Case psProd
        Case Producto.gCapAhorros
            lsAper = "Select Sum(nNumAper) from CapEstadMovimiento where datediff(m,dEstad,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and nProducto = " & gCapAhorros
        Case Producto.gCapPlazoFijo
            lsAper = "Select Sum(nNumAper) from CapEstadMovimiento where datediff(m,dEstad,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and nProducto = " & gCapPlazoFijo
        Case Producto.gCapCTS
            lsAper = "Select Sum(nNumAper) from CapEstadMovimiento where datediff(m,dEstad,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and nProducto = " & gCapCTS
    End Select
    oCon.AbreConexion
    'If Not oCon.AbreConexionRemota(gsCodAge, True, False, "03") Then
    If Not oCon.AbreConexion Then
        Exit Function
    End If
    
    If Not pbRetencionesJudiciales Then
        If pbCoberturado Then
            lsSql = "Select ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 1 THEN nSdoCob + nIntCob ELSE 0 END) ,0) nMontoMN, " _
                  & "       ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 2 THEN nSdoCob + nIntCob ELSE 0 END) ,0) nMontoME, " _
                  & "       ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 1 and nSdoCob + nIntCob > 0 THEN 1 ELSE 0 END) ,0) nCuentasMN, " _
                  & "       ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 2 and nSdoCob + nIntCob > 0 THEN 1 ELSE 0 END) ,0) nCuentasME " _
                  & IIf(pbGetAperturas, ", nCuentasApe = ISNULL( ( " & lsAper & " ) ,0) ", "") _
                  & "FROM (SELECT cCtaCod cCtaCod, Round(SUM(nSdoCob),2) nSdoCob, Round(SUM(nIntCob),2) nIntCob " _
                  & "      FROM " & sBaseFSD & "FSDPersonaDet " _
                  & "      WHERE cCtaCod Not In (" _
                  & " SELECT Distinct CA.cCtaCod FROM DBConsolidada..ProductoBloqueosConsol PB" _
                  & " Inner Join DBCMACMaynas..Captaciones CA On PB.cCtaCod = CA.cCtaCod" _
                  & " Inner Join ( " _
                  & "   SELECT cCtaCod, nEstCtaAC nPrdEstado FROM DBConsolidada..AhorroCConsol " _
                  & "       Union SELECT cCtaCod, nEstCtaPF nPrdEstado FROM DBConsolidada..PlazofijoConsol " _
                  & "       Union SELECT cCtaCod, nEstCtaCTS nPrdEstado FROM DBConsolidada..CTSConsol " _
                  & " ) PR On PB.cCtaCod = PR.cCtaCod And nPrdEstado Not In (1300,1400) " _
                  & " WHERE cMovNroDbl Is Null And nBlqMotivo In (1,3,23) And nPersoneria Not In (3,4,5,6) And dCierre = '" & Format(DateAdd("d", -1, DateAdd("m", 1, pdFecha)), gsFormatoFecha) & "') And SubString(cCtaCod,6,3) = '" & psProd & "' and bInactiva = " & pnInactiva & " and not cPersCod IN (SELECT cPersCod FROM " & sBaseFSD & "FSDPersonaExonerados ) "
        Else
            If pbCapital Then
                lsSql = "Select  ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 1 THEN nSdoDis ELSE 0 END),0) nMontoMN, " _
                      & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 2 THEN ROUND(nSdoDis * " & gnTipCambio & ",2) ELSE 0 END),0) nMontoME, " _
                      & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 1 and nSdoDis > 0 THEN 1 ELSE 0 END),0) nCuentasMN, " _
                      & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 2 and nSdoDis > 0 THEN 1 ELSE 0 END),0) nCuentasME " _
                      & IIf(pbGetAperturas, ", nCuentasApe = ISNULL( ( " & lsAper & " ) ,0) ", "") _
                      & " FROM (SELECT cCtaCod cCtaCod, Round(SUM(nSdoDis),2) nSdoDis, SUM(Round(nIntMes,2)) nIntMes  " _
                      & "      FROM " & sBaseFSD & "FSDPersonaDet " _
                      & "      WHERE cCtaCod Not In (" _
                      & " SELECT Distinct CA.cCtaCod FROM DBConsolidada..ProductoBloqueosConsol PB" _
                      & " Inner Join DBCMACMaynas..Captaciones CA On PB.cCtaCod = CA.cCtaCod" _
                      & " Inner Join ( " _
                      & "   SELECT cCtaCod, nEstCtaAC nPrdEstado FROM DBConsolidada..AhorroCConsol " _
                      & "       Union SELECT cCtaCod, nEstCtaPF nPrdEstado FROM DBConsolidada..PlazofijoConsol " _
                      & "       Union SELECT cCtaCod, nEstCtaCTS nPrdEstado FROM DBConsolidada..CTSConsol " _
                      & " ) PR On PB.cCtaCod = PR.cCtaCod And nPrdEstado Not In (1300,1400) " _
                      & " WHERE cMovNroDbl Is Null And nBlqMotivo In (1,3,23) And nPersoneria Not In (3,4,5,6) And dCierre = '" & Format(DateAdd("d", -1, DateAdd("m", 1, pdFecha)), gsFormatoFecha) & "') And SubString(cCtaCod,6,3) = '" & psProd & "' and bInactiva = " & pnInactiva & " and not cPersCod IN (SELECT cPersCod FROM " & sBaseFSD & "FSDPersonaExonerados ) "
            Else
                lsSql = "Select  ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 1 THEN nIntMes ELSE 0 END),0) nMontoMN, " _
                      & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 2 THEN ROUND(nIntMes * " & gnTipCambio & ",2) ELSE 0 END),0) nMontoME, " _
                      & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 1 and nIntMes > 0 THEN 1 ELSE 0 END),0) nCuentasMN, " _
                      & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 2 and nIntMes > 0 THEN 1 ELSE 0 END),0) nCuentasME " _
                      & "FROM (SELECT cCtaCod cCtaCod, Round(SUM(nSdoDis),2) nSdoDis, SUM(Round(nIntMes,2)) nIntMes " _
                      & " FROM " & sBaseFSD & "FSDPersonaDet WHERE cCtaCod Not In (" _
                      & " SELECT Distinct CA.cCtaCod FROM DBConsolidada..ProductoBloqueosConsol PB" _
                      & " Inner Join DBCMACMaynas..Captaciones CA On PB.cCtaCod = CA.cCtaCod" _
                      & " Inner Join ( " _
                      & "   SELECT cCtaCod, nEstCtaAC nPrdEstado FROM DBConsolidada..AhorroCConsol " _
                      & "       Union SELECT cCtaCod, nEstCtaPF nPrdEstado FROM DBConsolidada..PlazofijoConsol " _
                      & "       Union SELECT cCtaCod, nEstCtaCTS nPrdEstado FROM DBConsolidada..CTSConsol " _
                      & " ) PR On PB.cCtaCod = PR.cCtaCod And nPrdEstado Not In (1300,1400) " _
                      & " WHERE cMovNroDbl Is Null And nBlqMotivo In (1,23) And nPersoneria Not In (3,4,5,6) And dCierre = '" & Format(DateAdd("d", -1, DateAdd("m", 1, pdFecha)), gsFormatoFecha) & "') And not cPersCod IN (SELECT cPersCod FROM " & sBaseFSD & "FSDPersonaExonerados ) "
            End If
        End If
        lsSql = lsSql & lsFiltroClte & " GROUP BY cCtaCod ) aa "
    Else
        lsSql = "Select  ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 1 THEN nSdoDis ELSE 0 END),0) nMontoMN, " _
              & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 2 THEN ROUND(nSdoDis * " & gnTipCambio & ",2) ELSE 0 END),0) nMontoME, " _
              & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 1 and nSdoDis > 0 THEN 1 ELSE 0 END),0) nCuentasMN, " _
              & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,9,1) = 2 and nSdoDis > 0 THEN 1 ELSE 0 END),0) nCuentasME " _
              & IIf(pbGetAperturas, ", nCuentasApe = 0  ", "") _
              & "FROM (SELECT cCtaCod cCtaCod, Round(SUM(nSdoDis),2) nSdoDis, SUM(Round(nIntMes,2)) nIntMes  " _
              & "      FROM " & sBaseFSD & "FSDPersonaDet " _
              & "      WHERE cCtaCod In (" _
              & " SELECT Distinct CA.cCtaCod FROM DBConsolidada..ProductoBloqueosConsol PB" _
              & " Inner Join DBCMACMaynas..Captaciones CA On PB.cCtaCod = CA.cCtaCod" _
              & " Inner Join ( " _
              & "   SELECT cCtaCod, nEstCtaAC nPrdEstado FROM DBConsolidada..AhorroCConsol " _
              & "       Union SELECT cCtaCod, nEstCtaPF nPrdEstado FROM DBConsolidada..PlazofijoConsol " _
              & "       Union SELECT cCtaCod, nEstCtaCTS nPrdEstado FROM DBConsolidada..CTSConsol " _
              & " ) PR On PB.cCtaCod = PR.cCtaCod And nPrdEstado Not In (1300,1400) " _
              & " WHERE cMovNroDbl Is Null And nBlqMotivo In (1,23) And nPersoneria Not In (3,4,5,6) And dCierre = '" & Format(DateAdd("d", -1, DateAdd("m", 1, pdFecha)), gsFormatoFecha) & "')  And Not cPersCod IN (SELECT cPersCod FROM " & sBaseFSD & "FSDPersonaExonerados ) "
        lsSql = lsSql & lsFiltroClte & " GROUP BY cCtaCod ) aa "
    End If
Else
    If pnClienteTipo > 0 Then
        lsFiltroClte = " and cPersCod "
        If pnClienteTipo = 1 Then
           lsFiltroClte = lsFiltroClte & " NOT "
        End If
        lsFiltroClte = lsFiltroClte & " IN (Select cPersCod FROM " & sBaseFSD & "FSDPersonaClasif) "
    End If
    Select Case psProd
        Case Producto.gCapAhorros
            lsAper = "Select Sum(nNumAperAC) from EstadDiaAcConsol where datediff(m,dEstadAc,'" & Format(pdFecha, gsFormatoFecha) & "') = 0"
        Case Producto.gCapPlazoFijo
            lsAper = "Select Sum(nNumAperPF) from EstadDiaPFConsol where datediff(m,dEstadPF,'" & Format(pdFecha, gsFormatoFecha) & "') = 0"
        Case Producto.gCapCTS
            lsAper = "Select Sum(nNumAperCTS) from EstadDiaCTS where datediff(m,dEstadCTS,'" & Format(pdFecha, gsFormatoFecha) & "') = 0"
    End Select
    oCon.AbreConexion
    If Not oCon.AbreConexion Then 'Remota(gsCodAge, True, False, "03") Then '
        Exit Function
    End If
    If pbCoberturado Then
        lsSql = "Select ISNULL( SUM(CASE WHEN SubString(cCtaCod,6,1) = 1 THEN nSdoCob + nIntCob ELSE 0 END) ,0) nMontoMN, " _
              & "       ISNULL( SUM(CASE WHEN SubString(cCtaCod,6,1) = 2 THEN nSdoCob + nIntCob ELSE 0 END) ,0) nMontoME, " _
              & "       ISNULL( SUM(CASE WHEN SubString(cCtaCod,6,1) = 1 and nSdoCob + nIntCob > 0 THEN 1 ELSE 0 END) ,0) nCuentasMN, " _
              & "       ISNULL( SUM(CASE WHEN SubString(cCtaCod,6,1) = 2 and nSdoCob + nIntCob > 0 THEN 1 ELSE 0 END) ,0) nCuentasME " _
              & IIf(pbGetAperturas, ", nCuentasApe = ISNULL( ( " & lsAper & " ) ,0) ", "") _
              & "FROM (SELECT cCtaCod, Round(SUM(nSdoCob),2) nSdoCob, SUM(Round(nIntCob,2)) nIntCob " _
              & "      FROM " & sBaseFSD & "FSDPersonaDet " _
              & "      WHERE SubString(cCtaCod,3,3) = '" & psProd & "' and bInactiva = " & pnInactiva & " and not cPersCod IN (SELECT cPersCod FROM " & sBaseFSD & "FSDPersonaExonerados ) "
    Else
        If pbCapital Then
            lsSql = "Select  ISNULL( SUM(CASE WHEN SubString(cCtaCod,6,1) = 1 THEN nSdoDis ELSE 0 END),0) nMontoMN, " _
                  & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,6,1) = 2 THEN ROUND(nSdoDis * " & gnTipCambio & ",2) ELSE 0 END),0) nMontoME, " _
                  & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,6,1) = 1 and nSdoDis > 0 THEN 1 ELSE 0 END),0) nCuentasMN, " _
                  & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,6,1) = 2 and nSdoDis > 0 THEN 1 ELSE 0 END),0) nCuentasME " _
                  & IIf(pbGetAperturas, ", nCuentasApe = ISNULL( ( " & lsAper & " ) ,0) ", "") _
                  & "FROM (SELECT cCtaCod, Round(SUM(nSdoDis),2) nSdoDis, SUM(Round(nIntMes,2)) nIntMes " _
                  & "    FROM " & sBaseFSD & "FSDPersonaDet " _
                  & "    WHERE SubString(cCtaCod,3,3) = '" & psProd & "' and bInactiva = " & pnInactiva & " and not cPersCod IN (SELECT cPersCod FROM " & sBaseFSD & "FSDPersonaExonerados ) " _
                  & " "
        Else
            lsSql = "Select  ISNULL( SUM(CASE WHEN SubString(cCtaCod,6,1) = 1 THEN nIntMes ELSE 0 END),0) nMontoMN, " _
                  & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,6,1) = 2 THEN ROUND(nIntMes * " & gnTipCambio & ",2) ELSE 0 END),0) nMontoME, " _
                  & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,6,1) = 1 and nIntMes > 0 THEN 1 ELSE 0 END),0) nCuentasMN, " _
                  & "        ISNULL( SUM(CASE WHEN SubString(cCtaCod,6,1) = 2 and nIntMes > 0 THEN 1 ELSE 0 END),0) nCuentasME " _
                  & "FROM (SELECT cCtaCod, Round(SUM(nSdoDis),2) nSdoDis, SUM(Round(nIntMes,2)) nIntMes " _
                  & "      FROM " & sBaseFSD & "FSDPersonaDet WHERE not cPersCod IN (SELECT cPersCod FROM " & sBaseFSD & "FSDPersonaExonerados ) " _
                  & " "
        End If
    End If
    lsSql = lsSql & lsFiltroClte & " GROUP BY cCtaCod ) aa "
End If

Set GetDatosResumenFSD = oCon.CargaRecordSet(lsSql)

End Function


Private Function GetDatosGarantia(pdFecha As Date, pnTipoCambio As Currency, pnMoneda As String, pbSinTipo2NiExonerados As Boolean) As ADODB.Recordset
    Dim lsSql  As String
    Dim lsAper As String
    Dim oCon   As New DConecta
    Dim lsFiltroClte As String
    Dim sServidorFSD As String
    Dim sBaseFSD As String
    Dim oConst As New NConstSistemas
    sServidorFSD = oConst.LeeConstSistema(gConstSistServFSD)
    sBaseFSD = sServidorFSD 'Mid(sServidorFSD, InStr(sServidorFSD, "]") + 2)
    Set oConst = Nothing
    
    If gbBitCentral Then
        oCon.AbreConexion
        If Not pbSinTipo2NiExonerados Then
            lsSql = " SELECT Sum(nSaldCnt) Saldo, Count(*) Numero, IsNull(Sum(Case Datediff(Month,dApertura,'" & Format(pdFecha, gsFormatoFecha) & "') When 0 Then 1 Else 0 End),0) NroAper " _
                  & " From (SELECT PC.cCtaCod, nSaldCnt, dApertura " _
                  & "       FROM DBConsolidada..ProductoBloqueosConsol PC" _
                  & "           Inner Join Captaciones b on PC.cCtaCod = b.cCtaCod" _
                  & "           Inner Join CapsaldosDiarios c on PC.cCtaCod = c.cCtaCod and dfecha >= '" & Format(pdFecha, gsFormatoFecha) & "' and dfecha < '" & Format(pdFecha + 1, gsFormatoFecha) & "'" _
                  & "       WHERE cMovNroDbl IS NULL AND nBlqMotivo = 3 And PC.cCtaCod Like '_____233" & pnMoneda & "%' " _
                  & "       Group By PC.cCtaCod, nSaldCnt, dApertura) As A"
        
            Set GetDatosGarantia = oCon.CargaRecordSet(lsSql)
        Else
            lsSql = " Select cCtaCod Into ##AAA  from DBConsolidada..ProductoPersonaConsol A " _
                  & " Inner Join (Select cPersCod from " & sBaseFSD & "FSDPersonaClasif " _
                  & " Union All  Select cPersCod from " & sBaseFSD & "FSDPersonaExonerados) B" _
                  & " On A.cPersCod = B.cPersCod Where cCtaCod Like '_____233" & pnMoneda & "%'"
            oCon.Ejecutar lsSql
            
            lsSql = " Create Index Aa on ##AAA (cCtaCod)"
            oCon.Ejecutar lsSql
            
            lsSql = " SELECT Sum(nSaldCnt) Saldo, Count(*) Numero, IsNull(Sum(Case Datediff(Month,dApertura,'" & Format(pdFecha, gsFormatoFecha) & "') When 0 Then 1 Else 0 End),0) NroAper " _
                  & " From (SELECT PC.cCtaCod, nSaldCnt, dApertura " _
                  & "       FROM DBConsolidada..ProductoBloqueosConsol PC" _
                  & "           Inner Join Captaciones b on PC.cCtaCod = b.cCtaCod " _
                  & "           Inner Join CapsaldosDiarios c on PC.cCtaCod = c.cCtaCod and dfecha >= '" & Format(pdFecha, gsFormatoFecha) & "' and dfecha < '" & Format(pdFecha + 1, gsFormatoFecha) & "'" _
                  & "       WHERE cMovNroDbl IS NULL AND nBlqMotivo = 3 And PC.cCtaCod Like '_____233" & pnMoneda & "%' And b.cCtaCod Not In " _
                  & "           (Select cCtaCod from ##AAA) " _
                  & "       Group By PC.cCtaCod, nSaldCnt, dApertura) As A"
        
            Set GetDatosGarantia = oCon.CargaRecordSet(lsSql)
            
            lsSql = " Drop Table ##AAA"
            oCon.Ejecutar lsSql
        End If
        
    End If
End Function

