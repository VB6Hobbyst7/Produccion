VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredBPPBonoPromotores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Promotores - Bonificación Mensual"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13230
   Icon            =   "frmCredBPPBonoPromotores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   13230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   12000
      TabIndex        =   4
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detalle Por Promotor"
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
      TabIndex        =   2
      Top             =   3480
      Width           =   12975
      Begin SICMACT.FlexEdit feDetalle 
         Height          =   2895
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   5106
         Cols0           =   12
         HighLight       =   1
         EncabezadosNombres=   "#-Nº Crédito-Cliente-Telefonos-Dirección-Tipo Cred.-Prod. Cred.-Moneda-Monto S/.-TEM-Analista-Comisión"
         EncabezadosAnchos=   "500-1800-3000-1000-2500-1800-1800-1200-1500-1000-1000-1500"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-L-L-C-R-C-C-R"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-2-2-0-2"
         CantEntero      =   15
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resumen Por Promotor"
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
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   12975
      Begin SICMACT.FlexEdit fePromotores 
         Height          =   1935
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   3413
         Cols0           =   7
         HighLight       =   1
         EncabezadosNombres=   "#-Promotor-Nº Créditos-Total Saldo S/.-Bonif. Bruta-Bonif. Neta-Cod"
         EncabezadosAnchos=   "500-3500-1100-1500-1500-1500-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-R-R-R-L"
         FormatosEdit    =   "0-0-0-2-2-2-0"
         CantEntero      =   15
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro de Búsqueda"
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
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      Begin VB.ComboBox cmbMeses 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   390
         Width           =   1815
      End
      Begin VB.ComboBox cmbAgencias 
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   390
         Width           =   2175
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Left            =   7200
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin Spinner.uSpinner uspAnio 
         Height          =   315
         Left            =   2880
         TabIndex        =   7
         Top             =   390
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Max             =   9999
         Min             =   1900
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes - Año :"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   450
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agencia :"
         Height          =   195
         Left            =   4080
         TabIndex        =   9
         Top             =   450
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmCredBPPBonoPromotores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************************************
''** Nombre      : frmCredBPPBonoPromotores
''** Descripción : Formulario para la generacion del bono promotores
''** Creación    : WIOR, 20140620 10:00:00 AM
''*****************************************************************************************************
'Option Explicit
'Private fsMesAnio As String
'Private fsAge As String
'Private fsCodPers As String
'Private Sub cmdCerrar_Click()
'Unload Me
'End Sub
'Private Sub CargaCombos()
'CargaComboAgencias cmbAgencias
'CargaComboMeses cmbMeses
'End Sub
'
'Private Sub cmdExportar_Click()
'If Trim(fePromotores.TextMatrix(1, 1)) <> "" Then
'    Dim nTipo As Integer
'    frmCredBPPBonoPromotoresSel.Show 1
'    nTipo = frmCredBPPBonoPromotoresSel.Tipo
'    If nTipo <> 0 Then
'        If Trim(fsCodPers) = "" And nTipo = 2 Then
'            MsgBox "Aun no ha seleccionado a ningun Promotor.", vbInformation, "Aviso"
'        Else
'            Call GenerarExcel(nTipo, fsCodPers)
'        End If
'    End If
'Else
'    MsgBox "Aun no se ha generado el Bono.", vbInformation, "Aviso"
'End If
'End Sub
'
'Private Sub cmdMostrar_Click()
'If Valida Then
'    Dim oBPP As COMNCredito.NCOMBPPR
'    Dim rsBPP As ADODB.Recordset
'    Dim i As Integer
'    fsMesAnio = uspAnio.valor & Trim(Right(cmbMeses.Text, 2))
'    fsAge = Trim(Right(cmbAgencias.Text, 2))
'    fsCodPers = ""
'    Set oBPP = New COMNCredito.NCOMBPPR
'    Set rsBPP = oBPP.ObtenerPromotoresMes(fsMesAnio, fsAge)
'
'
'    LimpiaFlex fePromotores
'    LimpiaFlex feDetalle
'    If Not (rsBPP.EOF And rsBPP.BOF) Then
'        For i = 1 To rsBPP.RecordCount
'            fePromotores.AdicionaFila
'            fePromotores.TextMatrix(i, 1) = Trim(rsBPP!Promotor)
'            fePromotores.TextMatrix(i, 2) = Trim(rsBPP!NCreditos)
'            fePromotores.TextMatrix(i, 3) = Format(rsBPP!TotalSaldo, "###," & String(15, "#") & "#0.00")
'            fePromotores.TextMatrix(i, 4) = Format(rsBPP!BoniBruta, "###," & String(15, "#") & "#0.00")
'            fePromotores.TextMatrix(i, 5) = Format(rsBPP!BoniNeta, "###," & String(15, "#") & "#0.00")
'            fePromotores.TextMatrix(i, 6) = Trim(rsBPP!cPromotor)
'            rsBPP.MoveNext
'        Next i
'        fePromotores.TopRow = 1
'    Else
'        MsgBox "No hay Datos", vbInformation, "Aviso"
'    End If
'End If
'End Sub
'
'Private Sub fePromotores_Click()
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
'Dim i As Integer
'Dim sCod As String
'fsCodPers = Trim(fePromotores.TextMatrix(fePromotores.row, 6))
'
'If fsCodPers <> "" Then
'    Set oBPP = New COMNCredito.NCOMBPPR
'    Set rsBPP = oBPP.ObtenerPromotoresMesDet(fsMesAnio, fsAge, fsCodPers)
'
'    LimpiaFlex feDetalle
'    If Not (rsBPP.EOF And rsBPP.BOF) Then
'        For i = 1 To rsBPP.RecordCount
'            feDetalle.AdicionaFila
'            feDetalle.TextMatrix(i, 1) = rsBPP!cCtaCod
'            feDetalle.TextMatrix(i, 2) = rsBPP!Cliente
'            feDetalle.TextMatrix(i, 3) = rsBPP!Telefonos
'            feDetalle.TextMatrix(i, 4) = rsBPP!Direccion
'            feDetalle.TextMatrix(i, 5) = rsBPP!TipoCred
'            feDetalle.TextMatrix(i, 6) = rsBPP!tipoprod
'            feDetalle.TextMatrix(i, 7) = rsBPP!Moneda
'            feDetalle.TextMatrix(i, 8) = Format(rsBPP!MontoSoles, "###," & String(15, "#") & "#0.00")
'            feDetalle.TextMatrix(i, 9) = Format(rsBPP!nInteres, "###," & String(15, "#") & "#0.000")
'            feDetalle.TextMatrix(i, 10) = rsBPP!UserAnalista
'            feDetalle.TextMatrix(i, 11) = Format(rsBPP!comision, "###," & String(15, "#") & "#0.00")
'            rsBPP.MoveNext
'        Next i
'        feDetalle.TopRow = 1
'    Else
'        MsgBox "No hay Datos", vbInformation, "Aviso"
'    End If
'End If
'End Sub
'
'Private Sub Form_Load()
'CargaCombos
'uspAnio.valor = Year(gdFecSis)
'End Sub
'
'Private Function Valida() As Boolean
'Valida = True
'
'If Trim(cmbMeses.Text) = "" Then
'    MsgBox "Seleccione el Mes", vbInformation, "Aviso"
'    cmbMeses.SetFocus
'    Valida = False
'    Exit Function
'End If
'
'If Trim(uspAnio.valor) = "" Or Trim(uspAnio.valor) = "0" Then
'    MsgBox "Ingrese el Año", vbInformation, "Aviso"
'    uspAnio.SetFocus
'    Valida = False
'    Exit Function
'End If
'
'If Trim(cmbAgencias.Text) = "" Then
'    MsgBox "Seleccione la Agencia", vbInformation, "Aviso"
'    cmbAgencias.SetFocus
'    Valida = False
'    Exit Function
'End If
'
'End Function
'
'Private Sub GenerarExcel(ByVal pnTipo As Integer, Optional ByVal psCod As String = "")
'Dim fs As Scripting.FileSystemObject
'Dim xlsAplicacion As Excel.Application
'Dim lsArchivo As String
'Dim lsFile As String
'Dim lsNomHoja As String
'Dim xlsLibro As Excel.Workbook
'Dim xlHoja1 As Excel.Worksheet
'Dim lbExisteHoja As Boolean
'Dim psArchivoAGrabarC As String
'Dim lnExcel As Long
'Dim sFormatoConta As String
'Dim sFormatoConta2 As String
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
'Dim i As Integer
'
'    On Error GoTo ErrorGeneraExcelFormato
'
'    Set fs = New Scripting.FileSystemObject
'    Set xlsAplicacion = New Excel.Application
'
'    lsNomHoja = "Promotores"
'
'    Select Case pnTipo
'        Case 1: lsFile = "FormatoBPPPromotorResumen"
'        Case 2: lsFile = "FormatoBPPPromotorDetalleSel"
'        Case 3: lsFile = "FormatoBPPPromotorDetalle"
'    End Select
'
'
'    lsArchivo = "\spooler\" & Replace(lsFile, "FormatoBPP", "") & Format(DateAdd("D", -1, DateAdd("M", 1, CDate("01/" + Right(fsMesAnio, 2) + "/" + Left(fsMesAnio, 4)))), "yyyymmdd") & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
'    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
'        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
'    Else
'        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
'        Exit Sub
'    End If
'
'    'Activar Hoja
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    lnExcel = 6
'
'
'    sFormatoConta = "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ *  - ??_ ;_ @_ "
'    sFormatoConta2 = "_ * #,##0.000_ ;_ * -#,##0.000_ ;_ *  - ??_ ;_ @_ "
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'    Select Case pnTipo
'        Case 1: Set rsBPP = oBPP.ObtenerPromotoresMes(fsMesAnio, fsAge)
'        Case 2: Set rsBPP = oBPP.ObtenerPromotoresMesDet(fsMesAnio, fsAge, psCod)
'        Case 3: Set rsBPP = oBPP.ObtenerPromotoresMesDet(fsMesAnio, fsAge)
'    End Select
'
'    If Not (rsBPP.EOF And rsBPP.BOF) Then
'        If pnTipo = 1 Or pnTipo = 3 Then
'            xlHoja1.Cells(lnExcel - 3, 2) = Trim(rsBPP!Agencia)
'        Else
'            xlHoja1.Cells(lnExcel - 3, 2) = Trim(rsBPP!Promotor)
'        End If
'
'        For i = 0 To rsBPP.RecordCount - 1
'            xlHoja1.Cells(lnExcel + i, 2) = i + 1
'            Select Case pnTipo
'                Case 1:
'                        xlHoja1.Cells(lnExcel + i, 3) = Trim(rsBPP!Promotor)
'                        xlHoja1.Cells(lnExcel + i, 4) = Trim(rsBPP!NCreditos)
'                        xlHoja1.Cells(lnExcel + i, 5).NumberFormat = sFormatoConta
'                        xlHoja1.Cells(lnExcel + i, 5) = CDbl(rsBPP!TotalSaldo)
'                        xlHoja1.Cells(lnExcel + i, 6).NumberFormat = sFormatoConta
'                        xlHoja1.Cells(lnExcel + i, 6) = CDbl(rsBPP!BoniBruta)
'                        xlHoja1.Cells(lnExcel + i, 7).NumberFormat = sFormatoConta
'                        xlHoja1.Cells(lnExcel + i, 7) = CDbl(rsBPP!BoniNeta)
'                        xlHoja1.Range(xlHoja1.Cells(lnExcel + i, 2), xlHoja1.Cells(lnExcel + i, 7)).Borders.LineStyle = 1
'                Case 2:
'                        xlHoja1.Cells(lnExcel + i, 3).NumberFormat = "@"
'                        xlHoja1.Cells(lnExcel + i, 3) = Trim(rsBPP!cCtaCod)
'                        xlHoja1.Cells(lnExcel + i, 4).NumberFormat = "@"
'                        xlHoja1.Cells(lnExcel + i, 4) = Trim(rsBPP!dVigencia)
'                        xlHoja1.Cells(lnExcel + i, 5) = Trim(rsBPP!Cliente)
'                        xlHoja1.Cells(lnExcel + i, 6) = Trim(rsBPP!Condicion)
'                        xlHoja1.Cells(lnExcel + i, 7).NumberFormat = "@"
'                        xlHoja1.Cells(lnExcel + i, 7) = Trim(rsBPP!Telefonos)
'                        xlHoja1.Cells(lnExcel + i, 8) = Trim(rsBPP!Direccion)
'                        xlHoja1.Cells(lnExcel + i, 9) = Trim(rsBPP!TipoCred)
'                        xlHoja1.Cells(lnExcel + i, 10) = Trim(rsBPP!tipoprod)
'                        xlHoja1.Cells(lnExcel + i, 11) = Trim(rsBPP!Moneda)
'                        xlHoja1.Cells(lnExcel + i, 12).NumberFormat = sFormatoConta
'                        xlHoja1.Cells(lnExcel + i, 12) = CDbl(rsBPP!MontoSoles)
'                        xlHoja1.Cells(lnExcel + i, 13).NumberFormat = sFormatoConta2
'                        xlHoja1.Cells(lnExcel + i, 13) = CDbl(rsBPP!nInteres)
'                        xlHoja1.Cells(lnExcel + i, 14) = Trim(rsBPP!UserAnalista)
'                        xlHoja1.Cells(lnExcel + i, 15).NumberFormat = "0.00%"
'                        xlHoja1.Cells(lnExcel + i, 15) = CDbl(rsBPP!Porcentaje)
'                        xlHoja1.Cells(lnExcel + i, 16).NumberFormat = sFormatoConta
'                        xlHoja1.Cells(lnExcel + i, 16) = CDbl(rsBPP!comision)
'                        xlHoja1.Range(xlHoja1.Cells(lnExcel + i, 2), xlHoja1.Cells(lnExcel + i, 16)).Borders.LineStyle = 1
'                Case 3:
'                        xlHoja1.Cells(lnExcel + i, 3) = Trim(rsBPP!Promotor)
'                        xlHoja1.Cells(lnExcel + i, 4).NumberFormat = "@"
'                        xlHoja1.Cells(lnExcel + i, 4) = Trim(rsBPP!cCtaCod)
'                        xlHoja1.Cells(lnExcel + i, 5).NumberFormat = "@"
'                        xlHoja1.Cells(lnExcel + i, 5) = Trim(rsBPP!dVigencia)
'                        xlHoja1.Cells(lnExcel + i, 6) = Trim(rsBPP!Cliente)
'                        xlHoja1.Cells(lnExcel + i, 7) = Trim(rsBPP!Condicion)
'                        xlHoja1.Cells(lnExcel + i, 8).NumberFormat = "@"
'                        xlHoja1.Cells(lnExcel + i, 8) = Trim(rsBPP!Telefonos)
'                        xlHoja1.Cells(lnExcel + i, 9) = Trim(rsBPP!Direccion)
'                        xlHoja1.Cells(lnExcel + i, 10) = Trim(rsBPP!TipoCred)
'                        xlHoja1.Cells(lnExcel + i, 11) = Trim(rsBPP!tipoprod)
'                        xlHoja1.Cells(lnExcel + i, 12) = Trim(rsBPP!Moneda)
'                        xlHoja1.Cells(lnExcel + i, 13).NumberFormat = sFormatoConta
'                        xlHoja1.Cells(lnExcel + i, 13) = CDbl(rsBPP!MontoSoles)
'                        xlHoja1.Cells(lnExcel + i, 14).NumberFormat = sFormatoConta2
'                        xlHoja1.Cells(lnExcel + i, 14) = CDbl(rsBPP!nInteres)
'                        xlHoja1.Cells(lnExcel + i, 15) = Trim(rsBPP!UserAnalista)
'                        xlHoja1.Cells(lnExcel + i, 16).NumberFormat = "0.00%"
'                        xlHoja1.Cells(lnExcel + i, 16) = CDbl(rsBPP!Porcentaje)
'                        xlHoja1.Cells(lnExcel + i, 17).NumberFormat = sFormatoConta
'                        xlHoja1.Cells(lnExcel + i, 17) = CDbl(rsBPP!comision)
'                        xlHoja1.Range(xlHoja1.Cells(lnExcel + i, 2), xlHoja1.Cells(lnExcel + i, 17)).Borders.LineStyle = 1
'            End Select
'            rsBPP.MoveNext
'        Next i
'
'        Select Case pnTipo
'            Case 1:
'                    xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 5).NumberFormat = sFormatoConta
'                    xlHoja1.Range("E" & (lnExcel + rsBPP.RecordCount)).Formula = "=SUM(E6:E" & (lnExcel + rsBPP.RecordCount - 1) & ")"
'                    xlHoja1.Range(xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 5), xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 5)).Borders.LineStyle = 1
'
'                    xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 7).NumberFormat = sFormatoConta
'                    xlHoja1.Range("G" & (lnExcel + rsBPP.RecordCount)).Formula = "=SUM(G6:G" & (lnExcel + rsBPP.RecordCount - 1) & ")"
'                    xlHoja1.Range(xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 7), xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 7)).Borders.LineStyle = 1
'            Case 2:
'                    xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 12).NumberFormat = sFormatoConta
'                    xlHoja1.Range("L" & (lnExcel + rsBPP.RecordCount)).Formula = "=SUM(L6:L" & (lnExcel + rsBPP.RecordCount - 1) & ")"
'                    xlHoja1.Range(xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 12), xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 12)).Borders.LineStyle = 1
'
'                    xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 16).NumberFormat = sFormatoConta
'                    xlHoja1.Range("P" & (lnExcel + rsBPP.RecordCount)).Formula = "=SUM(P6:P" & (lnExcel + rsBPP.RecordCount - 1) & ")"
'                    xlHoja1.Range(xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 16), xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 16)).Borders.LineStyle = 1
'            Case 3:
'                    xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 13).NumberFormat = sFormatoConta
'                    xlHoja1.Range("M" & (lnExcel + rsBPP.RecordCount)).Formula = "=SUM(M6:M" & (lnExcel + rsBPP.RecordCount - 1) & ")"
'                    xlHoja1.Range(xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 13), xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 13)).Borders.LineStyle = 1
'
'                    xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 17).NumberFormat = sFormatoConta
'                    xlHoja1.Range("Q" & (lnExcel + rsBPP.RecordCount)).Formula = "=SUM(Q6:Q" & (lnExcel + rsBPP.RecordCount - 1) & ")"
'                    xlHoja1.Range(xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 17), xlHoja1.Cells(lnExcel + rsBPP.RecordCount, 17)).Borders.LineStyle = 1
'        End Select
'    End If
'    Set rsBPP = Nothing
'    Set oBPP = Nothing
'    xlHoja1.Range(xlHoja1.Cells(lnExcel + i, 2), xlHoja1.Cells(lnExcel + i, 17)).EntireColumn.AutoFit
'
'    xlHoja1.SaveAs App.path & lsArchivo
'    psArchivoAGrabarC = App.path & lsArchivo
'    xlsAplicacion.Visible = True
'    xlsAplicacion.Windows(1).Visible = True
'    Set xlsAplicacion = Nothing
'    Set xlsLibro = Nothing
'    Set xlHoja1 = Nothing
'
'    MsgBox "Fromato Generado Satisfactoriamente en la ruta: " & psArchivoAGrabarC, vbInformation, "Aviso"
'
'    Exit Sub
'ErrorGeneraExcelFormato:
'    MsgBox err.Description, vbCritical, "Error a Generar El Formato Excel"
'
'End Sub
