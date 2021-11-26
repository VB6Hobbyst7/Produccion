VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPersPREDA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuarios PREDA"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6900
   Icon            =   "frmPersPREDA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar pgbExcel 
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   5760
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
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
      Left            =   3000
      TabIndex        =   7
      Top             =   5040
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdFormato 
      Caption         =   "Formato"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "Cargar"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   " Carga de Nuevos Usuarios "
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
      TabIndex        =   2
      Top             =   4800
      Width           =   6615
      Begin VB.CommandButton cmdAutorizar 
         Caption         =   "Grabar Autorización"
         Height          =   375
         Left            =   4680
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label txtCarga 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
   End
   Begin SICMACT.FlexEdit fePREDA 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7435
      Cols0           =   4
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-DNI-Nombre-Autorizado"
      EncabezadosAnchos=   "600-1000-3600-1000"
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
      ColumnasAEditar =   "X-X-X-3"
      ListaControles  =   "0-0-0-4"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "R-R-L-C"
      FormatosEdit    =   "0-0-0-1"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   600
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Usuarios PREDA registrados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmPersPREDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmPersPREDA
'** Descripción : Formulario administrar usuarios PREDA según TI-ERS099-2013
'** Creación : JUEZ, 20130718 09:00:00 AM
'**********************************************************************************************
Option Explicit
Dim oPREDA As COMDPersona.DCOMPersonas
Dim rs As ADODB.Recordset
Dim sFilename As String

Public Sub Inicio()
    CargaDatos
    Me.Show 1
End Sub

Public Sub CargaDatos()
    Dim lnFila As Long
    Set oPREDA = New COMDPersona.DCOMPersonas
    Set rs = oPREDA.RecuperaDatosPREDA
    Call LimpiaFlex(fePREDA)
        
    Do While Not rs.EOF
        fePREDA.AdicionaFila
        lnFila = fePREDA.row
        fePREDA.TextMatrix(lnFila, 1) = rs!cPersIDnro
        fePREDA.TextMatrix(lnFila, 2) = rs!cPersNombre
        fePREDA.TextMatrix(lnFila, 3) = IIf(rs!bAutorizado, "1", "") 'ORCR-20140913*********
        rs.MoveNext
    Loop
    
    'Evaluar Cargo de
    fePREDA.lbEditarFlex = True
    
    fePREDA.TopRow = 1
End Sub

Private Sub cmdAutorizar_Click() 'ORCR-20140913*********
    Dim i As Integer
    Set oPREDA = New COMDPersona.DCOMPersonas
    
    For i = 1 To fePREDA.Rows - 1
        Call oPREDA.ActualizarPersonaPREDAAutorizado(fePREDA.TextMatrix(i, 1), fePREDA.TextMatrix(i, 3))
    Next i
    
    CargaDatos
    
    MsgBox "Operacion Completada", vbInformation, "Grabar Autoriazación"
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCargar_Click()
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim varMatriz As Variant
Dim cNombreHoja As String
Dim i As Long, n As Long
Dim nUsuCarga As Integer, nUsuExiste As Integer, nUsuNoCarga As Integer
Dim pbExisteHoja As Boolean

Set xlApp = New Excel.Application
If Trim(txtCarga.Caption) = "" Then
    MsgBox "Debe indicar la ruta del Archivo Excel", vbInformation, "Mensaje"
    Exit Sub
Else
    pgbExcel.value = 0
    pgbExcel.Min = 0
    Set xlLibro = xlApp.Workbooks.Open(txtCarga.Caption, True, True, , "")
    cNombreHoja = "Formato"
    For Each xlHoja In xlLibro.Worksheets
        If xlHoja.Name = cNombreHoja Then
            pbExisteHoja = True
            Exit For
        End If
    Next
    If pbExisteHoja = False Then
        MsgBox "No existe ninguna hoja con nombre 'Formato'", vbInformation, "Aviso"
        Exit Sub
    End If
    'validar nombre de hoja
    Set xlHoja = xlApp.Worksheets(cNombreHoja)
    varMatriz = xlHoja.Range("A1:B65536").value
    xlLibro.Close SaveChanges:=False
    xlApp.Quit
    Set xlHoja = Nothing
    Set xlLibro = Nothing
    Set xlApp = Nothing
    
    For i = 2 To UBound(varMatriz)
        If CStr(varMatriz(i, 1)) = "" And varMatriz(i, 2) = "" Then Exit For
        n = n + 1
    Next i
    If n = 0 Then
        MsgBox "No hay datos para la carga", vbInformation, "Aviso"
        Exit Sub
    End If
    
    pgbExcel.Max = n
    
    Set oPREDA = New COMDPersona.DCOMPersonas
    
    If varMatriz(1, 1) <> "DNI" Or varMatriz(1, 2) <> "NOMBRES" Then
        MsgBox "Archivo No tiene Estructura Correcta, la cabecera DNI debe estar en la fila A:1 y los Nombres en la fila B:1", vbCritical, "Mensaje"
        Exit Sub
    End If
    
    For i = 1 To n
        If CStr(varMatriz(i + 1, 1)) = "" And varMatriz(i + 1, 2) = "" Then Exit For
        If varMatriz(i + 1, 1) = "" Or varMatriz(i + 1, 2) = "" Then
            nUsuNoCarga = nUsuNoCarga + 1
        ElseIf varMatriz(i + 1, 1) <> "" And varMatriz(i + 1, 2) <> "" Then
            If oPREDA.VerificarPersonaPREDA(varMatriz(i + 1, 1), 2) Then
                nUsuExiste = nUsuExiste + 1
            Else
                Call oPREDA.InsertaPersonaPREDA(varMatriz(i + 1, 1), UCase(varMatriz(i + 1, 2)))
                nUsuCarga = nUsuCarga + 1
            End If
        End If
        pgbExcel.value = pgbExcel.value + 1
    Next i
    Set oPREDA = Nothing
    txtCarga.Caption = ""
    MsgBox "Carga realizada con éxito:" & Chr(10) & "* " & nUsuCarga & " usuarios cargados" & Chr(10) & "* " & nUsuExiste & " usuarios ya existentes" & Chr(10) & "* " & nUsuNoCarga & " registros no pudieron cargarse por no contar con los datos completos", vbInformation, "Carga Finalizada"
    pgbExcel.value = 0
    pgbExcel.Min = 0
    CargaDatos
End If
End Sub

Private Sub cmdExportar_Click()
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim lbLibroOpen As Boolean
Dim lsArchivo As String
Dim lsHoja As String
Dim xlHoja1 As Excel.Worksheet
Dim xlHoja2 As Excel.Worksheet
Dim nLin As Long
Dim nItem As Long
Dim sColumna As String

    pgbExcel.Min = 0
    pgbExcel.value = 0

    lsArchivo = App.Path & "\SPOOLER\BaseDeDatosPREDA_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".xls"
    lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If Not lbLibroOpen Then
        Exit Sub
    End If
    nLin = 1
    lsHoja = "Formato"
    gFunGeneral.ExcelAddHoja lsHoja, xlLibro, xlHoja1
    
    xlHoja1.Range("A1:Y1").EntireColumn.Font.FontStyle = "Arial"
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.TopMargin = 2
    
    xlHoja1.Range("A1:A1").RowHeight = 18
    xlHoja1.Range("A1:A1").ColumnWidth = 10
    xlHoja1.Range("B1:B1").ColumnWidth = 50
    
    xlHoja1.Cells(nLin, 1) = "DNI"
    xlHoja1.Cells(nLin, 2) = "NOMBRES"
    
    xlHoja1.Range("A" & nLin & ":B" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":B" & nLin).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A" & nLin & ":B" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("A" & nLin & ":B" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A" & nLin & ":B" & nLin).Borders(xlInsideVertical).Color = vbBlack
    xlHoja1.Range("A" & nLin & ":B" & nLin).Interior.Color = RGB(255, 50, 50)
    xlHoja1.Range("A" & nLin & ":B" & nLin).Font.Color = RGB(255, 255, 255)
    
    With xlHoja1.PageSetup
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
    
    nItem = 1
    nLin = nLin + 1
    pgbExcel.Max = fePREDA.Rows - 1
    For nItem = 1 To fePREDA.Rows - 1
        xlHoja1.Range("A" & nLin & ":B" & nLin).HorizontalAlignment = xlHAlignLeft
        xlHoja1.Cells(nLin, 1) = "'" & fePREDA.TextMatrix(nItem, 1)
        xlHoja1.Cells(nLin, 2) = fePREDA.TextMatrix(nItem, 2)
        pgbExcel.value = pgbExcel.value + 1
        nLin = nLin + 1
    Next nItem
    
    gFunGeneral.ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
    gFunGeneral.CargaArchivo lsArchivo, App.Path & "\SPOOLER\"
    pgbExcel.Min = 0
    pgbExcel.value = 0
End Sub

Private Sub cmdFormato_Click()
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim lbLibroOpen As Boolean
Dim lsArchivo As String
Dim lsHoja As String
Dim xlHoja1 As Excel.Worksheet
Dim xlHoja2 As Excel.Worksheet
Dim nLin As Long
Dim nItem As Long
Dim sColumna As String
    pgbExcel.value = 0
    pgbExcel.Min = 0
    pgbExcel.Max = 3
    lsArchivo = App.Path & "\SPOOLER\FormatoPREDA_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".xls"
    lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If Not lbLibroOpen Then
        Exit Sub
    End If
    nLin = 1
    lsHoja = "Formato"
    gFunGeneral.ExcelAddHoja lsHoja, xlLibro, xlHoja1
    
    pgbExcel.value = 1
    
    xlHoja1.Range("A1:Y1").EntireColumn.Font.FontStyle = "Arial"
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.TopMargin = 2
    
    xlHoja1.Range("A1:A1").RowHeight = 18
    xlHoja1.Range("A1:A1").ColumnWidth = 10
    xlHoja1.Range("B1:B1").ColumnWidth = 40
    
    xlHoja1.Cells(nLin, 1) = "DNI"
    xlHoja1.Cells(nLin, 2) = "NOMBRES"
    
    pgbExcel.value = 2
    
    xlHoja1.Range("A" & nLin & ":B" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":B" & nLin).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A" & nLin & ":B" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("A" & nLin & ":B" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A" & nLin & ":B" & nLin).Borders(xlInsideVertical).Color = vbBlack
    xlHoja1.Range("A" & nLin & ":B" & nLin).Interior.Color = RGB(255, 50, 50)
    xlHoja1.Range("A" & nLin & ":B" & nLin).Font.Color = RGB(255, 255, 255)
    
    With xlHoja1.PageSetup
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
    
    pgbExcel.value = 3
    
    gFunGeneral.ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
    gFunGeneral.CargaArchivo lsArchivo, App.Path & "\SPOOLER\"
    pgbExcel.value = 0
    pgbExcel.Min = 0
End Sub

Private Sub cmdLoad_Click()
    CommonDialog1.Filter = "Archivos de Excel (*.xls)|*.xls|Todos los Archivo (*.*)|*.*"
    CommonDialog1.ShowOpen
    txtCarga.Caption = CommonDialog1.FileName
End Sub
