VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCredMetasAnalistas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Metas de Analistas"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19095
   Icon            =   "frmCredMetasAnalistas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   19095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17640
      TabIndex        =   6
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16200
      TabIndex        =   5
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   " Cargar Datos "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   18855
      Begin VB.TextBox txtCarga 
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   8295
      End
      Begin VB.CommandButton cmdFormato 
         Caption         =   "Formato"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10320
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtAnio 
         Height          =   315
         Left            =   14640
         TabIndex        =   3
         Top             =   270
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   16920
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   270
         Width           =   1815
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8500
         TabIndex        =   0
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Mes :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   16320
         TabIndex        =   9
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Año :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14040
         TabIndex        =   8
         Top             =   330
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SICMACT.FlexEdit FEMetas 
      Height          =   4665
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   18855
      _ExtentX        =   33258
      _ExtentY        =   8229
      Cols0           =   14
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmCredMetasAnalistas.frx":030A
      EncabezadosAnchos=   "300-0-2000-0-900-3500-2600-1400-1300-1200-1200-1400-1200-1500"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   65535
      BackColorControl=   65535
      BackColorControl=   65535
      EncabezadosAlineacion=   "C-C-L-C-C-L-L-R-R-R-R-R-R-R"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483635
   End
   Begin ComctlLib.ProgressBar pgbExcel 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5955
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblProcess 
      Caption         =   "Procesando..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5760
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmCredMetasAnalistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCredMetasAnalistas
'** Descripción : Formulario para ingresar las metas mensuales para los analistas
'** Creación : JUEZ, 20160407 09:00:00 AM
'*****************************************************************************************************

Option Explicit

'Private WithEvents oNCred As COMNCredito.NCOMCredito
Dim oNCred As COMNCredito.NCOMCredito

Private Sub Form_Load()
    CargarMeses
    txtAnio.Text = Year(gdFecSis)
    cboMes.ListIndex = Month(gdFecSis) - 1
End Sub

Private Sub CargarMeses()
Dim oDConst As COMDConstantes.DCOMConstantes
Dim rs As ADODB.Recordset

    Set oDConst = New COMDConstantes.DCOMConstantes
        Set rs = oDConst.RecuperaConstantes(gMeses)
    Set oDConst = Nothing
    
    cboMes.Clear
    While Not rs.EOF
        cboMes.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
        rs.MoveNext
    Wend
End Sub

Private Sub cmdLoad_Click()
    CommonDialog1.Filter = "Archivos de Excel (*.xls)|*.xls| Archivos de Excel (*.xlsx)|*.xlsx| Todos los Archivo (*.*)|*.*"
    CommonDialog1.ShowOpen
    txtCarga.Text = CommonDialog1.FileName
End Sub

Private Sub cmdCargar_Click()
Dim oGen As COMDConstSistema.DCOMGeneral
Dim oAge As COMDConstantes.DCOMAgencias
Dim rs As ADODB.Recordset
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim vMatLista As Variant
Dim cNombreHoja As String
Dim i As Long, n As Long, j As Long
Dim pbExisteHoja As Boolean

Dim lsArchivo As String
Dim lbLibroOpen As Boolean
Dim lsHoja As String

Dim lbCargaFlex As Boolean

Dim lsAgeCod As String
Dim lsPersCod As String

LimpiaFlex FEMetas

Set xlApp = New Excel.Application
If Trim(txtCarga.Text) = "" Then
    MsgBox "Debe indicar la ruta del Archivo Excel", vbInformation, "Mensaje"
    Exit Sub
Else
    pgbExcel.value = 0
    pgbExcel.Min = 0
    Set xlLibro = xlApp.Workbooks.Open(txtCarga.Text, True, True, , "")
    cNombreHoja = "Metas"
    For Each xlHoja In xlLibro.Worksheets
        If xlHoja.Name = cNombreHoja Then
            pbExisteHoja = True
            Exit For
        End If
    Next
    If pbExisteHoja = False Then
        MsgBox "No existe ninguna hoja con nombre 'Metas'", vbInformation, "Aviso"
        Exit Sub
    End If
    'validar nombre de hoja
    Set xlHoja = xlApp.Worksheets(cNombreHoja)
    vMatLista = xlHoja.Range("A1:K65536").value
    xlLibro.Close SaveChanges:=False
    xlApp.Quit
    Set xlHoja = Nothing
    Set xlLibro = Nothing
    Set xlApp = Nothing
    
    For i = 2 To UBound(vMatLista)
        If Trim(CStr(vMatLista(i, 1))) = "" And Trim(CStr(vMatLista(i, 2))) = "" Then Exit For
        n = n + 1
    Next i
    If n = 0 Then
        MsgBox "No hay datos para la carga", vbInformation, "Aviso"
        Exit Sub
    End If
    
    pgbExcel.Max = n
    
    If vMatLista(1, 1) <> "Agencia" Or vMatLista(1, 2) <> "Usuario" Or vMatLista(1, 3) <> "Apellidos y Nombres" Or vMatLista(1, 4) <> "Cargo" Or _
       vMatLista(1, 5) <> "Meta Saldo de Cartera Cierre" Or vMatLista(1, 6) <> "Meta Número de Clientes Cierre" Or vMatLista(1, 7) <> "Meta Número de Operaciones Cierre" Or vMatLista(1, 8) <> "Meta CA" Or _
       vMatLista(1, 9) <> "Saldo a Bajar CA" Or vMatLista(1, 10) <> "Meta CAR" Or vMatLista(1, 11) <> "Saldo a Bajar CAR" Then
        MsgBox "Archivo no tiene estructura correcta, se debe respetar la estructura del Formato", vbExclamation, "Mensaje"
        Exit Sub
    End If
    
    lsArchivo = App.Path & "\SPOOLER\ErroresMetasAnalista_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".xls"
    lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivo, xlApp, xlLibro, False)
    
    lsHoja = "Metas"
    gFunGeneral.ExcelAddHoja lsHoja, xlLibro, xlHoja
    GeneraDatosExcel xlHoja
    j = 1
    
    Set oGen = New COMDConstSistema.DCOMGeneral
    Set oAge = New COMDConstantes.DCOMAgencias
    
    For i = 1 To n
        lbCargaFlex = True
        If Trim(CStr(vMatLista(i + 1, 1))) = "" Then Exit For
        
        If vMatLista(i + 1, 1) = "" Or vMatLista(i + 1, 2) = "" Or vMatLista(i + 1, 3) = "" Or vMatLista(i + 1, 4) = "" Or _
           vMatLista(i + 1, 5) = "" Or vMatLista(i + 1, 6) = "" Or vMatLista(i + 1, 7) = "" Or vMatLista(i + 1, 8) = "" Or _
           vMatLista(i + 1, 9) = "" Or vMatLista(i + 1, 10) = "" Or vMatLista(i + 1, 11) = "" Then
            CargaListaEnExcelError "Datos Incompletos", xlHoja, vMatLista, j, i
            j = j + 1
        ElseIf Not IsNumeric(vMatLista(i + 1, 5)) Or Not IsNumeric(vMatLista(i + 1, 6)) Or Not IsNumeric(vMatLista(i + 1, 7)) Or Not IsNumeric(vMatLista(i + 1, 8)) Or _
               Not IsNumeric(vMatLista(i + 1, 9)) Or Not IsNumeric(vMatLista(i + 1, 10)) Or Not IsNumeric(vMatLista(i + 1, 11)) Then
            CargaListaEnExcelError "Los valores de las metas no están ingresadas correctamente", xlHoja, vMatLista, j, i
            j = j + 1
        Else
            lsAgeCod = oAge.ObtieneCodigoAgencia(Trim(CStr(vMatLista(i + 1, 1))))
            If lsAgeCod = "" Then
                CargaListaEnExcelError "No se reconoce la agencia", xlHoja, vMatLista, j, i
                lbCargaFlex = False
                j = j + 1
            End If
            
            Set rs = oGen.GetDataUser(vMatLista(i + 1, 2))
            If Not rs.EOF And Not rs.BOF Then
                lsPersCod = rs!cPersCod
            Else
                CargaListaEnExcelError "Usuario no está registrado en el Sistema", xlHoja, vMatLista, j, i
                lbCargaFlex = False
                j = j + 1
            End If
            
            If lbCargaFlex Then
                FEMetas.AdicionaFila
                FEMetas.TextMatrix(FEMetas.row, 1) = lsAgeCod
                FEMetas.TextMatrix(FEMetas.row, 2) = vMatLista(i + 1, 1)
                FEMetas.TextMatrix(FEMetas.row, 3) = lsPersCod
                FEMetas.TextMatrix(FEMetas.row, 4) = vMatLista(i + 1, 2)
                FEMetas.TextMatrix(FEMetas.row, 5) = vMatLista(i + 1, 3)
                FEMetas.TextMatrix(FEMetas.row, 6) = vMatLista(i + 1, 4)
                FEMetas.TextMatrix(FEMetas.row, 7) = Format(vMatLista(i + 1, 5), "#,##0.00")
                FEMetas.TextMatrix(FEMetas.row, 8) = Format(vMatLista(i + 1, 6), "#,##0")
                FEMetas.TextMatrix(FEMetas.row, 9) = Format(vMatLista(i + 1, 7), "#,##0")
                FEMetas.TextMatrix(FEMetas.row, 10) = Round(CDbl(vMatLista(i + 1, 8)) * 100, 2)
                FEMetas.TextMatrix(FEMetas.row, 11) = Format(vMatLista(i + 1, 9), "#,##0.00")
                FEMetas.TextMatrix(FEMetas.row, 12) = Round(CDbl(vMatLista(i + 1, 10)) * 100, 2)
                FEMetas.TextMatrix(FEMetas.row, 13) = Format(vMatLista(i + 1, 11), "#,##0.00")
            End If
        End If
        pgbExcel.value = pgbExcel.value + 1
    Next i
    
    Set oGen = Nothing
    Set oAge = Nothing
    
    gFunGeneral.ExcelEnd lsArchivo, xlApp, xlLibro, xlHoja
    If j > 1 Then
        gFunGeneral.CargaArchivo lsArchivo, App.Path & "\SPOOLER\"
    Else
        Dim oFileSys As FileSystemObject
        Set oFileSys = New FileSystemObject
        oFileSys.DeleteFile lsArchivo, True
        Set oFileSys = Nothing
    End If
    
    cmdGrabar.SetFocus
    pgbExcel.value = 0
    pgbExcel.Min = 0
End If
End Sub

Private Sub CargaListaEnExcelError(ByVal psMensaje As String, ByVal pxlHoja As Excel.Worksheet, ByVal pvMatLista As Variant, ByVal pnFilaExcel As Long, ByVal pnFilaMat As Long)
    pxlHoja.Cells(pnFilaExcel + 1, 1) = pvMatLista(pnFilaMat + 1, 1)
    pxlHoja.Cells(pnFilaExcel + 1, 2) = pvMatLista(pnFilaMat + 1, 2)
    pxlHoja.Cells(pnFilaExcel + 1, 3) = pvMatLista(pnFilaMat + 1, 3)
    pxlHoja.Cells(pnFilaExcel + 1, 4) = pvMatLista(pnFilaMat + 1, 4)
    pxlHoja.Cells(pnFilaExcel + 1, 5) = pvMatLista(pnFilaMat + 1, 5)
    pxlHoja.Cells(pnFilaExcel + 1, 6) = pvMatLista(pnFilaMat + 1, 6)
    pxlHoja.Cells(pnFilaExcel + 1, 7) = pvMatLista(pnFilaMat + 1, 7)
    pxlHoja.Cells(pnFilaExcel + 1, 8) = pvMatLista(pnFilaMat + 1, 8)
    pxlHoja.Cells(pnFilaExcel + 1, 9) = pvMatLista(pnFilaMat + 1, 9)
    pxlHoja.Cells(pnFilaExcel + 1, 10) = pvMatLista(pnFilaMat + 1, 10)
    pxlHoja.Cells(pnFilaExcel + 1, 11) = pvMatLista(pnFilaMat + 1, 11)
    pxlHoja.Cells(pnFilaExcel + 1, 12) = psMensaje
    pxlHoja.Range("A" & pnFilaExcel + 1 & ":L" & pnFilaExcel + 1).Interior.Color = RGB(255, 255, 0)
    pxlHoja.Range("L" & pnFilaExcel + 1 & ":L" & pnFilaExcel + 1).Font.Bold = True
End Sub

Private Sub cmdFormato_Click()
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim lbLibroOpen As Boolean
Dim lsArchivo As String
Dim lsHoja As String
Dim xlHoja1 As Excel.Worksheet
Dim nLin As Long
Dim nItem As Long
Dim sColumna As String
    pgbExcel.value = 0
    pgbExcel.Min = 0
    pgbExcel.Max = 3
    lsArchivo = App.Path & "\SPOOLER\FormatoMetasAnalista_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".xls"
    lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If Not lbLibroOpen Then
        Exit Sub
    End If
    nLin = 1
    lsHoja = "Metas"
    gFunGeneral.ExcelAddHoja lsHoja, xlLibro, xlHoja1
    
    pgbExcel.value = 1
    
    GeneraDatosExcel xlHoja1
    
    pgbExcel.value = 3
    
    gFunGeneral.ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
    gFunGeneral.CargaArchivo lsArchivo, App.Path & "\SPOOLER\"
    pgbExcel.value = 0
    pgbExcel.Min = 0
End Sub

Private Sub GeneraDatosExcel(ByVal pxlHoja As Excel.Worksheet)
Dim nLin As Long
    
    pxlHoja.Range("A1:Y1").EntireColumn.Font.FontStyle = "Calibri"
    pxlHoja.Range("A1:Y1").EntireColumn.Font.Size = "10"
    pxlHoja.PageSetup.CenterHorizontally = True
    pxlHoja.PageSetup.Zoom = 75
    pxlHoja.PageSetup.TopMargin = 2
    
    pxlHoja.Range("A1:A1").RowHeight = 33
    pxlHoja.Range("A1:A1").ColumnWidth = 20
    pxlHoja.Range("B1:B1").ColumnWidth = 10
    pxlHoja.Range("C1:C1").ColumnWidth = 40
    pxlHoja.Range("D1:D1").ColumnWidth = 25
    pxlHoja.Range("E1:E1").ColumnWidth = 15
    pxlHoja.Range("F1:F1").ColumnWidth = 15
    pxlHoja.Range("G1:G1").ColumnWidth = 18
    pxlHoja.Range("H1:H1").ColumnWidth = 11
    pxlHoja.Range("I1:I1").ColumnWidth = 13
    pxlHoja.Range("J1:J1").ColumnWidth = 11
    pxlHoja.Range("K1:K1").ColumnWidth = 13
    pxlHoja.Range("L1:L1").ColumnWidth = 40
    
    pxlHoja.Cells(1, 1) = "Agencia"
    pxlHoja.Cells(1, 2) = "Usuario"
    pxlHoja.Cells(1, 3) = "Apellidos y Nombres"
    pxlHoja.Cells(1, 4) = "Cargo"
    pxlHoja.Cells(1, 5) = "Meta Saldo de Cartera Cierre"
    pxlHoja.Cells(1, 6) = "Meta Número de Clientes Cierre"
    pxlHoja.Cells(1, 7) = "Meta Número de Operaciones Cierre"
    pxlHoja.Cells(1, 8) = "Meta CA"
    pxlHoja.Cells(1, 9) = "Saldo a Bajar CA"
    pxlHoja.Cells(1, 10) = "Meta CAR"
    pxlHoja.Cells(1, 11) = "Saldo a Bajar CAR"
    
    pxlHoja.Range("A1:K1").Font.Bold = True
    pxlHoja.Range("A1:K1").HorizontalAlignment = xlHAlignCenter
    pxlHoja.Range("A1:K1").VerticalAlignment = xlHAlignCenter
    pxlHoja.Range("A1:K1").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    pxlHoja.Range("A1:K1").Borders(xlInsideVertical).LineStyle = xlContinuous
    pxlHoja.Range("A1:K1").Borders(xlInsideVertical).Color = vbBlack
    pxlHoja.Range("A1:K1").WrapText = True
    pxlHoja.Range("E2:E65536").NumberFormat = "#,##0.00"
    pxlHoja.Range("F2:G65536").NumberFormat = "#,##0"
    pxlHoja.Range("H2:H65536").NumberFormat = "0.00%"
    pxlHoja.Range("I2:I65536").NumberFormat = "#,##0.00"
    pxlHoja.Range("J2:J65536").NumberFormat = "0.00%"
    pxlHoja.Range("K2:K65536").NumberFormat = "#,##0.00"
    pxlHoja.Range("A1:K1").Interior.Color = RGB(255, 50, 50)
    pxlHoja.Range("A1:K1").Font.Color = RGB(255, 255, 255)
    
    With pxlHoja.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
    
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
    End With
End Sub

Private Sub cmdGrabar_Click()
Dim rsMetas As ADODB.Recordset
Dim lbRegistra As Boolean
    
    'If CInt(CStr(txtAnio.Text) + Trim(Right(Me.cboMes.Text, 2))) < CInt(CStr(Year(gdFecSis)) + CStr(Month(gdFecSis))) Then
    If CDbl(CStr(txtAnio.Text) + Trim(Right(Me.cboMes.Text, 2))) < CDbl(CStr(Year(gdFecSis)) + CStr(Month(gdFecSis))) Then 'JOEP 20161004
        MsgBox "No puede ingresar metas de anteriores metas", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set rsMetas = IIf(FEMetas.Rows - 1 > 0, FEMetas.GetRsNew(), Nothing)
    
    If rsMetas Is Nothing Then
        MsgBox "Debe cargar las metas a registrar", vbInformation, "Aviso"
        Exit Sub
    End If
    If rsMetas.RecordCount <= 0 Then
        MsgBox "Debe cargar las metas a registrar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Se van a registrar las metas de los analistas, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    pgbExcel.value = 0
    pgbExcel.Min = 0
    pgbExcel.Max = rsMetas.RecordCount
    
    Set oNCred = New COMNCredito.NCOMCredito
        lblProcess.Visible = True
        pgbExcel.value = val(rsMetas.RecordCount) / 2
        lbRegistra = oNCred.GrabarMetasAnalistas(rsMetas, CInt(txtAnio.Text), CInt(Trim(Right(Me.cboMes.Text, 2))), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
        pgbExcel.value = rsMetas.RecordCount
        lblProcess.Visible = False
    Set oNCred = Nothing
    
    If lbRegistra Then
        MsgBox "Se han registrado las metas", vbInformation, "Aviso"
        cmdCancelar_Click
    End If
    
    pgbExcel.value = 0
End Sub

'Private Sub oNCred_ValorProgressBar(pnValor As Long)
'    pgbExcel.value = pnValor
'End Sub

Private Sub cmdCancelar_Click()
    txtCarga.Text = ""
    txtAnio.Text = Year(gdFecSis)
    cboMes.ListIndex = Month(gdFecSis) - 1
    LimpiaFlex FEMetas
End Sub
