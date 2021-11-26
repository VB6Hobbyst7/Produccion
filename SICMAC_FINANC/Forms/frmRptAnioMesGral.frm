VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRptAnioMesGral 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5355
   Icon            =   "frmRptAnioMesGral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraRep 
      Height          =   1365
      Left            =   50
      TabIndex        =   4
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "&Generar"
         Height          =   345
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   345
         Left            =   3960
         TabIndex        =   3
         Top             =   600
         Width           =   1155
      End
      Begin VB.Frame fraMes 
         Caption         =   "Periodo"
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
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   3735
         Begin VB.ComboBox cboMes 
            Height          =   315
            ItemData        =   "frmRptAnioMesGral.frx":030A
            Left            =   2040
            List            =   "frmRptAnioMesGral.frx":0332
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   300
            Width           =   1455
         End
         Begin VB.TextBox txtAnio 
            Alignment       =   1  'Right Justify
            Height          =   280
            Left            =   600
            MaxLength       =   4
            TabIndex        =   0
            Top             =   300
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mes :"
            Height          =   195
            Left            =   1560
            TabIndex        =   7
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Año :"
            Height          =   195
            Left            =   180
            TabIndex        =   6
            Top             =   360
            Width           =   375
         End
      End
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "frmRptAnioMesGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fsOpeCod As String

Public Sub Ini(ByVal psOpeCod As String, ByVal psOpeDesc As String)
    fsOpeCod = psOpeCod
    Me.Caption = psOpeDesc
    Me.Show 1
End Sub
Private Function obtenerFechaFinMes(ByVal pnMes As Integer, ByVal pnAnio As Integer) As Date
    Dim sFecha  As Date
    sFecha = CDate("01/" & Format(pnMes, "00") & "/" & pnAnio)
    sFecha = DateAdd("m", 1, sFecha)
    sFecha = sFecha - 1
    obtenerFechaFinMes = sFecha
End Function

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub

Private Sub cmdGenerar_Click()
    If Not validaGenerar Then Exit Sub
    Select Case fsOpeCod
        Case gOpeRptTripleJump
            generarRptTripleJump
        Case gOpeRptGtoFinanProd
            generarRptGastoFinancieroxProducto
        Case gOpeRptInfoEstadColocBCRP 'EJVG20121112
            generarRptInfoEstadColocBCRP
    End Select
    PB1.Visible = False
End Sub
Private Sub generarRptGastoFinancieroxProducto()
    Dim oCtaCont As DbalanceCont
    Dim clsTC As COMDConstSistema.NCOMTipoCambio
    Dim obj_Excel As Object, Libro As Object, Hoja As Object
    Dim celda As Excel.Range
    Dim fs As New Scripting.FileSystemObject
        
    Dim lsPath As String
    Dim ldFechaFinMesActual As Date, ldFechaMesSgte As Date
    Dim i As Integer, lnContCta As Integer, lnIndicador As Integer, lnCuenta As Integer
    Dim lbAbierto As Boolean
    
    lnIndicador = 7
    lsPath = App.path & "\FormatoCarta\RptGastoFinancieroxProducto.xls"
    
    If Not validarDatos Then
        Exit Sub
    End If
    If Len(Dir(lsPath)) = 0 Then
        MsgBox "No se Pudo Encontrar el Archivo:" & lsPath, vbCritical
        Exit Sub
    End If
    
    'verifica si existe el archivo
    If fs.FileExists(lsPath) Then
        lbAbierto = True
        Do While lbAbierto
            If ArchivoEstaAbierto(lsPath) Then
                lbAbierto = True
                If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(lsPath) + " para continuar", vbRetryCancel, "Aviso") = vbCancel Then
                    Exit Sub
                End If
            Else
                lbAbierto = False
            End If
        Loop
    End If

    On Error GoTo error_GastoFinancieroxProducto
    
    Set obj_Excel = CreateObject("Excel.Application")
    obj_Excel.DisplayAlerts = False
    Set Libro = obj_Excel.Workbooks.Open(lsPath)
    Set Hoja = Libro.ActiveSheet
    
    Set oCtaCont = New DbalanceCont
    ldFechaFinMesActual = obtenerFechaFinMes(1, Val(txtAnio.Text) - 1)
    lnCuenta = 12 + (cboMes.ListIndex + 1)
    
    PB1.Min = 0
    PB1.Max = lnCuenta
    PB1.value = 0
    PB1.Visible = True
    Me.MousePointer = vbHourglass
    
    For i = 1 To lnCuenta
        lnContCta = i + 26
        'Ctas Contables
        Set celda = obj_Excel.Range("B" & lnContCta)
        celda.value = ldFechaFinMesActual
        Set celda = obj_Excel.Range("C" & lnContCta)
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("410102", ldFechaFinMesActual, "0", "1")
        Set celda = obj_Excel.Range("D" & lnContCta)
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("410103", ldFechaFinMesActual, "0", "1")
        Set celda = obj_Excel.Range("E" & lnContCta)
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("410302", ldFechaFinMesActual, "0", "1")
        Set celda = obj_Excel.Range("F" & lnContCta)
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("410303", ldFechaFinMesActual, "0", "1")
        Set celda = obj_Excel.Range("G" & lnContCta)
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("4104", ldFechaFinMesActual, "0", "1")
        Set celda = obj_Excel.Range("H" & lnContCta)
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("410107", ldFechaFinMesActual, "0", "1")
        Set celda = obj_Excel.Range("I" & lnContCta)
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2102", ldFechaFinMesActual, "0", "1")
        Set celda = obj_Excel.Range("J" & lnContCta)
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2302", ldFechaFinMesActual, "0", "1")
        Set celda = obj_Excel.Range("K" & lnContCta)
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2103", ldFechaFinMesActual, "0", "1")
        Set celda = obj_Excel.Range("L" & lnContCta)
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2303", ldFechaFinMesActual, "0", "1")
        Set celda = obj_Excel.Range("M" & lnContCta)
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2107", ldFechaFinMesActual, "0", "1")
        Set celda = obj_Excel.Range("N" & lnContCta)
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("24", ldFechaFinMesActual, "0", "1")
        Set celda = obj_Excel.Range("O" & lnContCta)
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("26", ldFechaFinMesActual, "0", "1")
        'Indicadores
        If i > lnCuenta - 13 Then
            Set celda = obj_Excel.Range("B" & lnIndicador)
            celda.value = ldFechaFinMesActual
            Set celda = obj_Excel.Range("C" & lnIndicador)
            celda.Formula = IIf(Month(ldFechaFinMesActual) = 1, "=+((C" & lnContCta & "+E" & lnContCta & ")*12)/(I" & lnContCta & "+J" & lnContCta & ")", "=+(((C" & lnContCta & "+E" & lnContCta & ")-(C" & (lnContCta - 1) & "-E" & (lnContCta - 1) & "))*12)/(((I" & lnContCta & "+J" & lnContCta & ")+(I" & (lnContCta - 1) & "+J" & (lnContCta - 1) & "))/2)")
            Set celda = obj_Excel.Range("D" & lnIndicador)
            celda.Formula = IIf(Month(ldFechaFinMesActual) = 1, "=+(D" & lnContCta & "+F" & lnContCta & "+H" & lnContCta & ")*12/(K" & lnContCta & "+L" & lnContCta & "+M" & lnContCta & ")", "=+((D" & lnContCta & "+F" & lnContCta & "+H" & lnContCta & "-D" & (lnContCta - 1) & "-F" & (lnContCta - 1) & "-H" & (lnContCta - 1) & ")*12)/(((K" & lnContCta & "+L" & lnContCta & "+M" & lnContCta & ")+(K" & (lnContCta - 1) & "+L" & (lnContCta - 1) & "+M" & (lnContCta - 1) & "))/2)")
            Set celda = obj_Excel.Range("E" & lnIndicador)
            celda.Formula = IIf(Month(ldFechaFinMesActual) = 1, "=+(C" & lnContCta & "+D" & lnContCta & "+E" & lnContCta & "+F" & lnContCta & "+H" & lnContCta & ")*12/(I" & lnContCta & "+J" & lnContCta & "+K" & lnContCta & "+L" & lnContCta & "+M" & lnContCta & ")", "=+((C" & lnContCta & "+D" & lnContCta & "+E" & lnContCta & "+F" & lnContCta & "+H" & lnContCta & "-D" & (lnContCta - 1) & "-C" & (lnContCta - 1) & "-E" & (lnContCta - 1) & "-F" & (lnContCta - 1) & "-H" & (lnContCta - 1) & ")*12)/((I" & lnContCta & "+J" & lnContCta & "+K" & lnContCta & "+L" & lnContCta & "+M" & lnContCta & "+I" & (lnContCta - 1) & "+J" & (lnContCta - 1) & "+K" & (lnContCta - 1) & "+L" & (lnContCta - 1) & "+M" & (lnContCta - 1) & ")/2)")
            Set celda = obj_Excel.Range("F" & lnIndicador)
            celda.Formula = IIf(Month(ldFechaFinMesActual) = 1, "=+(G" & lnContCta & ")*12/(N" & lnContCta & "+O" & lnContCta & ")", "=+((G" & lnContCta & "-G" & (lnContCta - 1) & ")*12)/((N" & lnContCta & "+O" & lnContCta & "+N" & (lnContCta - 1) & "+O" & (lnContCta - 1) & ")/2)")
            Set celda = obj_Excel.Range("G" & lnIndicador)
            celda.Formula = IIf(Month(ldFechaFinMesActual) = 1, "=+((C" & lnContCta & "+D" & lnContCta & "+E" & lnContCta & "+F" & lnContCta & "+G" & lnContCta & "+H" & lnContCta & ")*12)/(I" & lnContCta & "+J" & lnContCta & "+K" & lnContCta & "+L" & lnContCta & "+M" & lnContCta & "+N" & lnContCta & "+O" & lnContCta & ")", "=+((C" & lnContCta & "+D" & lnContCta & "+E" & lnContCta & "+F" & lnContCta & "+G" & lnContCta & "+H" & lnContCta & "-C" & (lnContCta - 1) & "-D" & (lnContCta - 1) & "-E" & (lnContCta - 1) & "-F" & (lnContCta - 1) & "-G" & (lnContCta - 1) & "-H" & (lnContCta - 1) & ")*12)/((I" & lnContCta & "+J" & lnContCta & "+K" & lnContCta & "+L" & lnContCta & "+M" & lnContCta & "+N" & lnContCta & "+O" & lnContCta & "+I" & (lnContCta - 1) & "+J" & (lnContCta - 1) & "+K" & (lnContCta - 1) & "+L" & (lnContCta - 1) & "+M" & (lnContCta - 1) & "+N" & (lnContCta - 1) & "+O" & (lnContCta - 1) & ")/2)")
            lnIndicador = lnIndicador + 1
        End If
        
        ldFechaMesSgte = DateAdd("M", 1, ldFechaFinMesActual)
        ldFechaFinMesActual = obtenerFechaFinMes(Month(ldFechaMesSgte), Year(ldFechaMesSgte))
        
        PB1.value = i
    Next
        
    Dim lsArchivo As String
    lsArchivo = App.path & "\Spooler\RptGastoFinancieroxProducto_" & Format(gdFecSis, "yyyymmdd") & Format(Now, "hhmmss") & ".xls"
    Hoja.SaveAs lsArchivo
    Libro.Close
    obj_Excel.Quit

    Set Hoja = Nothing
    Set Libro = Nothing
    Set obj_Excel = Nothing
    
    Dim m_excel As New Excel.Application
    m_excel.Workbooks.Open (lsArchivo)
    m_excel.Visible = True
    PB1.Visible = False
    Me.MousePointer = vbDefault
Exit Sub
error_GastoFinancieroxProducto:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
    Set Libro = Nothing
    Set obj_Excel = Nothing
    Set Hoja = Nothing
    PB1.Visible = False
    Me.MousePointer = vbDefault
End Sub
Private Sub generarRptTripleJump()
    Dim oCtaCont As New DbalanceCont
    Dim rsCtaCont As New ADODB.Recordset
    Dim oContImp As NContImprimir
    Dim clsTC As COMDConstSistema.NCOMTipoCambio
    Dim obj_Excel As Object, Libro As Object, Hoja As Object
    Dim celda As Excel.Range
    Dim lnCPDeposPlazoFijo, lnLPDeposPlazoFijo, lnTipoCambio As Currency
    Dim ldFechaFinMes As Date
    Dim rsPlazoFijo As Recordset
    Dim i As Integer
    
    Dim fs As New Scripting.FileSystemObject
    Dim lsPathTripleJump As String
    
    lsPathTripleJump = App.path & "\Spooler\Triple_Jump_" + Format(gdFecSis, "yyyymmdd") + ".xls"
    If Not validarDatos Then
        Exit Sub
    End If

    On Error GoTo error_TripleJump
    
    Set oContImp = New NContImprimir
    ldFechaFinMes = obtenerFechaFinMes(Val(cboMes.ListIndex + 1), Val(txtAnio.Text))
    
    If oContImp.validarGeneracionRepTripleJump(ldFechaFinMes) = False Then
        MsgBox "No se puede generar el Reporte Triple Jump a esta Fecha", vbInformation, "Aviso"
        Exit Sub
    End If

    If fs.FileExists(lsPathTripleJump) Then
        Dim lbAbierto As Boolean
        lbAbierto = True
        Do While lbAbierto
            If ArchivoEstaAbierto(lsPathTripleJump) Then
                lbAbierto = True
                If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(lsPathTripleJump) + " para continuar", vbRetryCancel, "Aviso") = vbCancel Then
                    Exit Sub
                End If
            Else
                lbAbierto = False
            End If
        Loop
    End If

    lsPathTripleJump = App.path & "\FormatoCarta\Triple_Jump.xls"
    If Len(Dir(lsPathTripleJump)) = 0 Then
        MsgBox "No se Pudo Encontrar el Archivo:" & lsPathTripleJump, vbCritical
        Exit Sub
    End If
    
    PB1.Min = 0
    PB1.Max = 18
    PB1.value = 0
    PB1.Visible = True
    
    Set obj_Excel = CreateObject("Excel.Application")
    obj_Excel.DisplayAlerts = False
    Set Libro = obj_Excel.Workbooks.Open(lsPathTripleJump)
    Set Hoja = Libro.ActiveSheet
    
    Me.MousePointer = vbHourglass
    lnTipoCambio = TipoCambioCierre(txtAnio, cboMes.ListIndex + 1, True)
    
    Set celda = obj_Excel.Range("C5")
    celda.value = ldFechaFinMes
    
    PB1.value = 1
    
    Set celda = obj_Excel.Range("C8")
    celda.value = Abs(oCtaCont.ObtenerCtaContBalanceMensual("11", ldFechaFinMes, "0", "1"))
    Set celda = obj_Excel.Range("C10")
    celda.Formula = "=" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140102060201", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140103060901", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140103139901", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("1401090607050301", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140112060201", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140113060201", ldFechaFinMes, "0", "1")) _
                   & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140402060201", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140403060901", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140412060201", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140413060201", ldFechaFinMes, "0", "1")) _
                   & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140502060201", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140502190601", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140503060901", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140503139901", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140503190601", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140512060201", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140512190601", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140513060201", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140513190601", ldFechaFinMes, "0", "1")) _
                   & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140602060201", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140603060901", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140612060201", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140613060201", ldFechaFinMes, "0", "1"))
    Set celda = obj_Excel.Range("C11")
    celda.Formula = "=" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140102060202", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140103060902", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140104060102", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("1401090607050302", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("1401090607051002", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("1401090607051102", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140112060202", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140113060202", ldFechaFinMes, "0", "1")) _
                   & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140402060202", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140403060902", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140412060202", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140413060202", ldFechaFinMes, "0", "1")) _
                   & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140502060202", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140502190602", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140503060902", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140503190602", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140504060102", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140512060202", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140513060202", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140513190602", ldFechaFinMes, "0", "1")) _
                   & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140602060202", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140603060902", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140604060102", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("140613060202", ldFechaFinMes, "0", "1"))
    Set celda = obj_Excel.Range("C12")
    
    PB1.value = 2
    
    celda.value = Abs(oCtaCont.ObtenerCtaContBalanceMensual("1409", ldFechaFinMes, "0", "1"))
    Set celda = obj_Excel.Range("C13")
    celda.value = Abs(oCtaCont.ObtenerCtaContBalanceMensual("18", ldFechaFinMes, "0", "1"))
    Set celda = obj_Excel.Range("C15")
    celda.Formula = "=" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("1904", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("1704", ldFechaFinMes, "0", "1"))
    Set celda = obj_Excel.Range("C14")
    celda.Formula = "=" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("1", ldFechaFinMes, "0", "1")) & "- (C8+C9-C12+C13)-C15"
    
    PB1.value = 3
    
    Dim dBalance As DCapMovimientos
    Dim rsAhorro As Recordset
    Dim lnCPAhorro, lnLPAhorro As Currency
    
    Set dBalance = New DCapMovimientos
    Set rsAhorro = New Recordset
    Set rsAhorro = dBalance.ReporteResumenAhorro(ldFechaFinMes, lnTipoCambio)
    
    Do While Not rsAhorro.EOF
        If rsAhorro!cMoneda = 1 And rsAhorro!nPlazo = 1 Then
            lnCPAhorro = lnCPAhorro + CCur(rsAhorro!nSaldoPF)
        End If
        If rsAhorro!cMoneda = 2 And rsAhorro!nPlazo = 1 Then
            lnCPAhorro = lnCPAhorro + CCur(rsAhorro!nSaldoPF)
        End If
        rsAhorro.MoveNext
    Loop
    
    Dim lnCP_DPF As Currency
    Dim lnLP_DPF As Currency
    Dim lnUtilidadAcum As Currency
   
    Call oContImp.obtenerCortoLargoDepositoPlazoFijo(ldFechaFinMes, lnTipoCambio, lnCP_DPF, lnLP_DPF)
            
    PB1.value = 4
            
    Set celda = obj_Excel.Range("C17")
    celda.Formula = "=" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("2102", ldFechaFinMes, "0", "1")) _
                        & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("2302", ldFechaFinMes, "0", "1")) _
                        & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("210305", ldFechaFinMes, "0", "1")) _
                        & "+" & lnCP_DPF
    
    PB1.value = 5
    
    Set celda = obj_Excel.Range("C18")
    celda.value = Abs(oCtaCont.ObtenerCtaContBalanceMensual("24", ldFechaFinMes, "0", "1"))
    Set celda = obj_Excel.Range("C19")
    celda.Formula = "=" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("2101", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("2104", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("2106", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("25", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("27", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("2908", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("2308", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("2901", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("2902", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("2909", ldFechaFinMes, "0", "1"))
    Set celda = obj_Excel.Range("C20")
    celda.Formula = lnLP_DPF
    
    PB1.value = 6
    
    Set celda = obj_Excel.Range("C22")
    celda.Formula = "=" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("2602020102", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("2607010102", ldFechaFinMes, "0", "1"))
    
    PB1.value = 7
    
    Set celda = obj_Excel.Range("C21")
    celda.Formula = "=" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("26", ldFechaFinMes, "0", "1")) & "- C22"
    Set celda = obj_Excel.Range("C23")
    celda.Formula = "=" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("2108", ldFechaFinMes, "0", "1")) & "-" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("2903", ldFechaFinMes, "0", "1"))
    Set celda = obj_Excel.Range("C25")
    celda.Formula = "=" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("31", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("32", ldFechaFinMes, "0", "1"))
    
    PB1.value = 8
    
    lnUtilidadAcum = Abs(oCtaCont.ObtenerCtaContBalanceMensual("5", ldFechaFinMes, "1", "2")) + Abs(oCtaCont.ObtenerCtaContBalanceMensual("62", ldFechaFinMes, "1", "2")) + Abs(oCtaCont.ObtenerCtaContBalanceMensual("64", ldFechaFinMes, "1", "2")) - (Abs(oCtaCont.ObtenerCtaContBalanceMensual("4", ldFechaFinMes, "1", "2")) + Abs(oCtaCont.ObtenerCtaContBalanceMensual("63", ldFechaFinMes, "1", "2")) + Abs(oCtaCont.ObtenerCtaContBalanceMensual("65", ldFechaFinMes, "1", "2"))) + Abs(oCtaCont.ObtenerCtaContBalanceMensual("69", ldFechaFinMes, "1", "2")) + oCtaCont.ObtenerCtaContBalanceMensual("67", ldFechaFinMes, "1", "2") * -1 + oCtaCont.ObtenerCtaContBalanceMensual("68", ldFechaFinMes, "1", "2") * -1
    Set celda = obj_Excel.Range("C26")
    celda.Formula = "=" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("33", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("38", ldFechaFinMes, "0", "1")) & "+" & lnUtilidadAcum
    
    PB1.value = 9
    
    Set celda = obj_Excel.Range("C31")
    celda.value = Abs(oCtaCont.ObtenerCtaContBalanceMensual("51", ldFechaFinMes, "0", "1"))
    Set celda = obj_Excel.Range("C32")
    celda.value = Abs(oCtaCont.ObtenerCtaContBalanceMensual("52", ldFechaFinMes, "0", "1"))
    
    PB1.value = 10
    
    Set celda = obj_Excel.Range("C33")
    celda.Formula = "=" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("41", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("42", ldFechaFinMes, "0", "1"))
    Set celda = obj_Excel.Range("C35")
    celda.value = Abs(oCtaCont.ObtenerCtaContBalanceMensual("4302", ldFechaFinMes, "0", "1")) - Abs(oCtaCont.ObtenerCtaContBalanceMensual("640104", ldFechaFinMes, "0", "1"))
    
    PB1.value = 11
    
    Set celda = obj_Excel.Range("C36")
    celda.value = Abs(oCtaCont.ObtenerCtaContBalanceMensual("45", ldFechaFinMes, "0", "1"))
    Set celda = obj_Excel.Range("C39")
    celda.value = Abs(oCtaCont.ObtenerCtaContBalanceMensual("62,64", ldFechaFinMes, "0", "1")) - Abs(oCtaCont.ObtenerCtaContBalanceMensual("640104", ldFechaFinMes, "0", "1"))
    
    PB1.value = 12
    
    Set celda = obj_Excel.Range("C40")
    celda.Formula = "=" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("43", ldFechaFinMes, "0", "1")) & "-" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("4302", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("44", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("63", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("65", ldFechaFinMes, "0", "1"))
    Set celda = obj_Excel.Range("C42")
    celda.Formula = "=" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("67", ldFechaFinMes, "0", "1")) & "+" & Abs(oCtaCont.ObtenerCtaContBalanceMensual("68", ldFechaFinMes, "0", "1"))
    
    PB1.value = 13
    
    Set celda = obj_Excel.Range("F8")
    celda.Formula = oContImp.obtenerTotalMargenFormaAYB(3, ldFechaFinMes, "222")
    Set celda = obj_Excel.Range("F9") 'Activos Nuevos
    celda.value = oContImp.obtenerTotalImporteActivoNuevo(ldFechaFinMes)
    Set celda = obj_Excel.Range("F10")
    celda.Formula = oContImp.obtenerTotalMargenFormaAYB(3, ldFechaFinMes, "207")
    
    PB1.value = 14
    
    Dim rsCartera As Recordset
    Set rsCartera = New Recordset
    
    Set rsCartera = oContImp.obtenerCarteraxRangoTripleJump(ldFechaFinMes)
    
    If rsCartera.RecordCount <> 6 Then 'Tienen que ser 6 registros de carteras
        MsgBox "Comuniquese con Sistemas para revisar el Reporte Triple Jump", vbInformation
        Exit Sub
    End If
    
    For i = 1 To 6
        Select Case i
            Case 1
                Set celda = obj_Excel.Range("F15")
            Case 2
                Set celda = obj_Excel.Range("F16")
            Case 3
                Set celda = obj_Excel.Range("F17")
            Case 4
                Set celda = obj_Excel.Range("F18")
            Case 5
                Set celda = obj_Excel.Range("F19")
            Case 6
                Set celda = obj_Excel.Range("F20")
        End Select
        celda.value = rsCartera!nSaldo
        rsCartera.MoveNext
    Next
    
    PB1.value = 15
    
    Set rsCartera = oContImp.obtenerCarteraCastigadaxRangoAño(ldFechaFinMes)
    Set celda = obj_Excel.Range("F21")
    celda.value = rsCartera!nCAPITALMN
    
    Set rsCartera = oContImp.obtenerEstadisticaOperacionTripleJump(ldFechaFinMes)
    For i = 1 To 9
        Select Case i
            Case 1
                Set celda = obj_Excel.Range("F24")
            Case 2
                Set celda = obj_Excel.Range("F25")
            Case 3
                Set celda = obj_Excel.Range("F26")
            Case 4
                Set celda = obj_Excel.Range("F27")
            Case 5
                Set celda = obj_Excel.Range("F28")
            Case 6
                Set celda = obj_Excel.Range("F29")
            Case 7
                Set celda = obj_Excel.Range("F30")
            Case 8
                Set celda = obj_Excel.Range("F31")
            Case 9
                Set celda = obj_Excel.Range("F32")
        End Select
        celda.value = rsCartera!nValor
        rsCartera.MoveNext
    Next
    
    PB1.value = 16
    
    Set celda = obj_Excel.Range("F35") 'Back to Back
    celda.Formula = 0
    Set celda = obj_Excel.Range("F36")
    celda.value = Abs(oCtaCont.ObtenerCtaContBalanceMensual("1", ldFechaFinMes, "2", "1"))
    Set celda = obj_Excel.Range("F37")
    celda.value = Abs(oCtaCont.ObtenerCtaContBalanceMensual("2", ldFechaFinMes, "2", "1"))
    
    PB1.value = 17
    
    'verifica si existe el archivo
    lsPathTripleJump = App.path & "\Spooler\Triple_Jump_" + Format(gdFecSis, "yyyymmdd") + ".xls"
    If fs.FileExists(lsPathTripleJump) Then
        If ArchivoEstaAbierto(lsPathTripleJump) Then
            MsgBox "Debe Cerrar el Archivo:" + fs.GetFileName(lsPathTripleJump)
        End If
'            Exit Sub
        'Set Libro = obj_Excel.Workbooks.Add
        fs.DeleteFile lsPathTripleJump, True
    End If
    'guarda el archivo
    Hoja.SaveAs lsPathTripleJump
    obj_Excel.Visible = True
    Libro.Close
    obj_Excel.Quit

    Set Hoja = Nothing
    Set Libro = Nothing
    Set obj_Excel = Nothing
    Me.MousePointer = vbDefault

    Dim m_excel As New Excel.Application
    m_excel.Workbooks.Open (lsPathTripleJump)
    m_excel.Visible = True
    PB1.value = 18
    PB1.Visible = False
Exit Sub
error_TripleJump:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
    Libro.Close
    Set Libro = Nothing
    Set obj_Excel = Nothing
    Set Hoja = Nothing
    PB1.Visible = False
    Me.MousePointer = vbDefault
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Function validarDatos() As Boolean
    validarDatos = True
    If txtAnio.Text = "" Or Len(txtAnio.Text) < 4 Then
        Me.MousePointer = vbDefault
        MsgBox "Ingrese un Año Válido", vbInformation
        txtAnio.SetFocus
        validarDatos = False
        Exit Function
    End If
    If cboMes.ListIndex = -1 Then
        Me.MousePointer = vbDefault
        MsgBox "Seleccione el Mes", vbInformation
        cboMes.SetFocus
        validarDatos = False
        Exit Function
    End If
End Function
Private Function ArchivoEstaAbierto(ByVal Ruta As String) As Boolean
On Error GoTo HayErrores
Dim f As Integer
   f = FreeFile
   Open Ruta For Append As f
   Close f
   ArchivoEstaAbierto = False
   Exit Function
HayErrores:
   If Err.Number = 70 Then
      ArchivoEstaAbierto = True
   Else
      Err.Raise Err.Number
   End If
End Function

Private Sub Command1_Click()
    'FormatoPlanillaLiquidacion.xls
    Dim fs As New Scripting.FileSystemObject
    Dim lsPathTripleJump As String
    Dim obj_Excel As Object, Libro As Object, Hoja As Object
    Dim celda As Excel.Range
    
    lsPathTripleJump = App.path & "\Spooler\Planilla_Liquidacion_" + Format(gdFecSis, "yyyymmdd") + ".xls"

    If fs.FileExists(lsPathTripleJump) Then
        If ArchivoEstaAbierto(lsPathTripleJump) Then
            If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(lsPathTripleJump) + " para continuar", vbRetryCancel) = vbCancel Then
               Me.MousePointer = vbDefault
               Exit Sub
            End If
            Me.MousePointer = vbHourglass
        End If
        fs.DeleteFile lsPathTripleJump, True
    End If

    lsPathTripleJump = App.path & "\FormatoCarta\FormatoPlanillaLiquidacionDet.xls"
    If Len(Dir(lsPathTripleJump)) = 0 Then
        MsgBox "No se Pudo Encontrar el Archivo:" & lsPathTripleJump, vbCritical
        Me.MousePointer = vbDefault
        PB1.Visible = False
        Exit Sub
    End If
   
    Set obj_Excel = CreateObject("Excel.Application")
    obj_Excel.DisplayAlerts = False
    Set Libro = obj_Excel.Workbooks.Open(lsPathTripleJump)
    Set Hoja = Libro.ActiveSheet
    
    Set celda = obj_Excel.Range("C5")
    'celda.value = ldFechaFinMes
    
'    Set Libro = obj_Excel.Workbooks.Add
'    Set Hoja = Libro.ActiveSheet
    
    'verifica si existe el archivo
    lsPathTripleJump = App.path & "\Spooler\Triple_Jump_" + Format(gdFecSis, "yyyymmdd") + ".xls"
    If fs.FileExists(lsPathTripleJump) Then
        If ArchivoEstaAbierto(lsPathTripleJump) Then
            MsgBox "Debe Cerrar el Archivo:" + fs.GetFileName(lsPathTripleJump)
    End If
    'Exit Sub
        'Set Libro = obj_Excel.Workbooks.Add
        fs.DeleteFile lsPathTripleJump, True
    End If
    'guarda el archivo
    Hoja.SaveAs lsPathTripleJump
    obj_Excel.Visible = True
    Libro.Close
    obj_Excel.Quit

    Set Hoja = Nothing
    Set Libro = Nothing
    Set obj_Excel = Nothing
    Me.MousePointer = vbDefault

    Dim m_excel As New Excel.Application
    m_excel.Workbooks.Open (lsPathTripleJump)
    m_excel.Visible = True
    PB1.value = 30
    PB1.Visible = False
End Sub
Private Sub Form_Load()
    CentraForm Me
End Sub
Private Sub txtAnio_Change()
    If Len(txtAnio.Text) = 4 Then
        cboMes.SetFocus
    End If
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub
'EJVG20121112 ***
Private Sub generarRptInfoEstadColocBCRP()
    Dim oAgencia As New DAgencia
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet, xlsHoja1 As Excel.Worksheet
    Dim rsAgencia As New ADODB.Recordset
    Dim ldFecha As Date
    Dim lsArchivo As String
    
    On Error GoTo ErrRptInfoEstadBCRP

    lsArchivo = "\spooler\RptInfoEstadColocBCRP" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    ldFecha = obtenerFechaFinMes(Val(cboMes.ListIndex + 1), Val(txtAnio.Text))

    Set xlsLibro = xlsAplicacion.Workbooks.Add
    Set rsAgencia = oAgencia.RecuperaAgenciasRptInfoEstadColocBCRP()
    
    PB1.Min = 0
    PB1.Max = rsAgencia.RecordCount + 2
    PB1.value = 0
    PB1.Visible = True
    
    'HOJAS AGENCIAS ***
    If Not RSVacio(rsAgencia) Then
        rsAgencia.Sort = "cAgeCod desc"
        Do While Not rsAgencia.EOF
            Set xlsHoja = xlsLibro.Worksheets.Add
            xlsHoja.Name = "CRED-" & rsAgencia!cAgeDescripcion
            Call generaHojaExcelAgencia_RptInfoEstadColocBCRP(xlsHoja, ldFecha, rsAgencia!cAgeCod)
            PB1.value = PB1.value + 1
            rsAgencia.MoveNext
        Loop
    Else
        MsgBox "No hay data para generar la información", vbInformation, "Aviso"
        PB1.Visible = False
        Exit Sub
    End If
    rsAgencia.MoveFirst
    rsAgencia.Sort = "cAgeCod asc"
    'HOJA TOTAL DE AGENCIAS ***
    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "CREDITOS TOTAL-AGENCIAS"
    Call generaHojaExcelTotalAgencias_RptInfoEstadColocBCRP(xlsHoja, ldFecha, rsAgencia)
    PB1.value = PB1.value + 1
    rsAgencia.MoveFirst
    'HOJA RESUMEN ***
    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "RESUMEN"
    Call generaHojaExcelResumen_RptInfoEstadColocBCRP(xlsHoja, ldFecha, rsAgencia)
    
    For Each xlsHoja1 In xlsLibro.Worksheets
        If UCase(xlsHoja1.Name) = "HOJA1" Or UCase(xlsHoja1.Name) = "HOJA2" Or UCase(xlsHoja1.Name) = "HOJA3" Then
            xlsHoja1.Delete
        End If
    Next

    PB1.value = PB1.value + 1
    MsgBox "Se ha generado el Reporte de Información Estadistica de Colocaciones BCRP", vbInformation, "Aviso"
    
    PB1.Visible = False
    xlsHoja.SaveAs App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    
    Set oAgencia = Nothing
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlsHoja = Nothing
    Exit Sub
ErrRptInfoEstadBCRP:
    MsgBox Err.Description, vbCritical, "Aviso"
    Exit Sub
End Sub
Private Sub generaHojaExcelAgencia_RptInfoEstadColocBCRP(ByRef xlsHoja As Worksheet, ByVal pdFecha As Date, ByVal psAgeCod As String)
    Dim oAgencia As New DAgencia
    Dim oBalance As New DbalanceCont
    Dim rsInfoEstadColoc As New ADODB.Recordset
    Dim lnMesNombre As String
    Dim lnLineaActual As Integer, lnLineaAnterior As Integer
    
    lnMesNombre = dameNombreMes(Month(pdFecha), True)
    
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    xlsHoja.Range("A:A").ColumnWidth = 4
    xlsHoja.Range("B:B").ColumnWidth = 20
    xlsHoja.Range("C:C").ColumnWidth = 60
    xlsHoja.Range("D:D").ColumnWidth = 20
    xlsHoja.Range("E:E").ColumnWidth = 20
    xlsHoja.Range("F:F").ColumnWidth = 20
    xlsHoja.Range("B13").RowHeight = 3
    
    xlsHoja.Range("B1") = "BANCO CENTRAL DE RESERVA DEL PERÚ"
    xlsHoja.Range("B2") = "SEDE REGIONAL LORETO"
    xlsHoja.Range("E4") = "AGENCIA:"
    xlsHoja.Range("F4") = psAgeCod & "-" & UCase(Trim(oAgencia.GetAgencias(psAgeCod)))
    xlsHoja.Range("B6") = "INSTITUCIÓN FINANCIERA: CAJA MUNICIPAL DE AHORROS Y CREDITOS MAYNAS S.A."
    xlsHoja.Range("B7") = "SALDOS AL MES DE " & lnMesNombre & " " & Year(pdFecha)
    '''xlsHoja.Range("B9") = "Llenar en miles de nuevos soles" 'marg ers-044-2016
    xlsHoja.Range("B9") = "Llenar en miles de " & LCase(gcPEN_PLURAL) 'marg ers-044-2016
    xlsHoja.Range("B10") = "CUENTA"
    xlsHoja.Range("C10") = "DESCRIPCION"
    xlsHoja.Range("D10") = "MONEDA NACIONAL"
    xlsHoja.Range("E10") = "MONEDA EXTRANJERA 1/"
    xlsHoja.Range("F10") = "TOTAL M/N Y M/E"
    xlsHoja.Range("B11") = "MN/ ME"
    xlsHoja.Range("B12") = "14"
    xlsHoja.Range("C12") = "CREDITOS"
    
    xlsHoja.Range("B9:F9").MergeCells = True
    
    xlsHoja.Range("F4").Interior.Color = RGB(255, 255, 0)
    xlsHoja.Range("B10:F10").Interior.Color = RGB(199, 199, 199)
    xlsHoja.Range("B6:B7").Font.Color = RGB(0, 0, 128)
    
    xlsHoja.Range("A1:F14").Font.Bold = True
    xlsHoja.Range("B11:B12").HorizontalAlignment = xlLeft
    xlsHoja.Range("B9:F10").HorizontalAlignment = xlCenter
    
    xlsHoja.Range("B10:F10").Borders.Weight = xlThin

    lnLineaActual = 13
    lnLineaAnterior = 11

    Set rsInfoEstadColoc = oBalance.RecuperaInfoEstadColocBCRP(Year(pdFecha), Month(pdFecha), psAgeCod)
    If Not RSVacio(rsInfoEstadColoc) Then
        Do While Not rsInfoEstadColoc.EOF
            If Len(rsInfoEstadColoc!cCtaContCod) = 4 Then
                xlsHoja.Range("B" & lnLineaActual).RowHeight = 3
                lnLineaActual = lnLineaActual + 1
                xlsHoja.Range("B" & lnLineaActual & ":F" & lnLineaActual) = Left(rsInfoEstadColoc!cCtaContCod, 2) & "0" & Mid(rsInfoEstadColoc!cCtaContCod, 4, 1) & "/" & Left(rsInfoEstadColoc!cCtaContCod, 2) & "2" & Mid(rsInfoEstadColoc!cCtaContCod, 4, 1)
                xlsHoja.Range("B" & lnLineaActual & ":F" & lnLineaActual).Font.Bold = True
            Else
                xlsHoja.Range("B" & lnLineaActual & ":F" & lnLineaActual) = Left(rsInfoEstadColoc!cCtaContCod, 2) & "0" & Mid(rsInfoEstadColoc!cCtaContCod, 4, 1) & "." & Mid(rsInfoEstadColoc!cCtaContCod, 5, 2) & "/" & Left(rsInfoEstadColoc!cCtaContCod, 2) & "2" & Mid(rsInfoEstadColoc!cCtaContCod, 4, 1) & "." & Mid(rsInfoEstadColoc!cCtaContCod, 5, 2)
            End If
            xlsHoja.Range("C" & lnLineaActual) = rsInfoEstadColoc!cCtaContDesc
            'xlsHoja.Range("D" & lnLineaActual) = rsInfoEstadColoc!nTotSoles
            xlsHoja.Range("D" & lnLineaActual) = "=" & rsInfoEstadColoc!nTotSoles & "/1000"  'EJVG20130225
            'xlsHoja.Range("E" & lnLineaActual) = rsInfoEstadColoc!nTotDolares
            xlsHoja.Range("E" & lnLineaActual) = "=" & rsInfoEstadColoc!nTotDolares & "/1000"  'EJVG20130225
            xlsHoja.Range("F" & lnLineaActual) = "= D" & lnLineaActual & "+E" & lnLineaActual
            lnLineaActual = lnLineaActual + 1
            rsInfoEstadColoc.MoveNext
        Loop
    Else
        lnLineaActual = lnLineaActual + 1
    End If
    
    xlsHoja.Range("B" & lnLineaAnterior & ":C" & lnLineaActual - 1).HorizontalAlignment = xlLeft
    xlsHoja.Range("D" & lnLineaAnterior & ":F" & lnLineaActual - 1).Font.Color = RGB(0, 0, 128)
    xlsHoja.Range("D" & lnLineaAnterior & ":F" & lnLineaActual - 1).HorizontalAlignment = xlRight
    
    xlsHoja.Range("B" & lnLineaAnterior & ":F" & lnLineaActual - 1).Borders(xlEdgeTop).Weight = xlThin
    xlsHoja.Range("B" & lnLineaAnterior & ":F" & lnLineaActual - 1).Borders(xlEdgeBottom).Weight = xlThin
    xlsHoja.Range("B" & lnLineaAnterior & ":F" & lnLineaActual - 1).Borders(xlEdgeLeft).Weight = xlThin
    xlsHoja.Range("B" & lnLineaAnterior & ":F" & lnLineaActual - 1).Borders(xlEdgeRight).Weight = xlThin
    xlsHoja.Range("B" & lnLineaAnterior & ":F" & lnLineaActual - 1).Borders(xlInsideVertical).Weight = xlThin
    xlsHoja.Range("D" & lnLineaAnterior & ":F" & lnLineaActual - 1).NumberFormat = "#,##0"
    
    xlsHoja.Range("B" & lnLineaActual) = "1/ Equivalente en moneda nacional"

    Set rsInfoEstadColoc = Nothing
    Set oAgencia = Nothing
    Set oBalance = Nothing
End Sub
Private Sub generaHojaExcelResumen_RptInfoEstadColocBCRP(ByRef xlsHoja As Worksheet, ByVal pdFecha As Date, ByVal rsAgencia As ADODB.Recordset)
    Dim oBalance As New DbalanceCont
    Dim rsInfoEstadColoc As New ADODB.Recordset
    Dim MatCtas As Variant
    Dim lnLineaActual As Integer, lnColumnaActual As Integer
    Dim tam As Long, iMat As Integer
    Dim bExisteCta As Boolean
    Dim lnMonto As Double

    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    xlsHoja.Range("A:A").ColumnWidth = 4
    
    xlsHoja.Cells(2, 2) = "INSTITUCIÓN FINANCIERA: CAJA MUNICIPAL DE AHORRO Y CREDITOS MAYNAS S.A."
    xlsHoja.Cells(3, 2) = "SALDOS AL MES DE " & dameNombreMes(Month(pdFecha), True) & " " & Year(pdFecha)
    
    'Para obtener distinct de cuentas a 4 digitos
    ReDim MatCtas(1 To 2, 0 To 0)
    Do While Not rsAgencia.EOF
        Set rsInfoEstadColoc = oBalance.RecuperaInfoEstadColocBCRP(Year(pdFecha), Month(pdFecha), rsAgencia!cAgeCod)
        Do While Not rsInfoEstadColoc.EOF
            If Len(rsInfoEstadColoc!cCtaContCod) = 4 Then
                bExisteCta = False
                For iMat = 1 To UBound(MatCtas, 2)
                    If MatCtas(1, iMat) = rsInfoEstadColoc!cCtaContCod Then
                        bExisteCta = True
                        Exit For
                    End If
                Next
                If Not bExisteCta Then
                    tam = UBound(MatCtas, 2) + 1
                    ReDim Preserve MatCtas(1 To 2, 0 To tam)
                    MatCtas(1, tam) = rsInfoEstadColoc!cCtaContCod
                    MatCtas(2, tam) = rsInfoEstadColoc!cCtaContDesc
                End If
            End If
            rsInfoEstadColoc.MoveNext
        Loop
        rsAgencia.MoveNext
    Loop

    xlsHoja.Cells(6, 2) = "Cuenta"
    xlsHoja.Cells(6, 3) = "Descripción"
    xlsHoja.Cells(5, 4) = "Agencias"
    
    xlsHoja.Range(xlsHoja.Cells(5, 2), xlsHoja.Cells(6, 2)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(5, 3), xlsHoja.Cells(6, 3)).MergeCells = True

    lnLineaActual = 7
    lnColumnaActual = 4
    For iMat = 1 To UBound(MatCtas, 2)
        lnColumnaActual = 4
        xlsHoja.Cells(lnLineaActual, 2) = Left(MatCtas(1, iMat), 2) & "0" & Mid(MatCtas(1, iMat), 4, 1) & "/" & Left(MatCtas(1, iMat), 2) & "2" & Mid(MatCtas(1, iMat), 4, 1)
        xlsHoja.Cells(lnLineaActual, 3) = MatCtas(2, iMat)
        rsAgencia.MoveFirst
        Do While Not rsAgencia.EOF
            xlsHoja.Cells(6, lnColumnaActual) = rsAgencia!cAgeDescripcion
            lnMonto = 0
            Set rsInfoEstadColoc = oBalance.RecuperaInfoEstadColocBCRP(Year(pdFecha), Month(pdFecha), rsAgencia!cAgeCod)
            Do While Not rsInfoEstadColoc.EOF
                If rsInfoEstadColoc!cCtaContCod = MatCtas(1, iMat) Then
                    lnMonto = rsInfoEstadColoc!nTotSoles + rsInfoEstadColoc!nTotDolares
                End If
                rsInfoEstadColoc.MoveNext
            Loop
            'xlsHoja.Cells(lnLineaActual, lnColumnaActual) = lnMonto
            xlsHoja.Cells(lnLineaActual, lnColumnaActual) = "=" & lnMonto & "/1000" 'EJVG20130225
            lnColumnaActual = lnColumnaActual + 1
            rsAgencia.MoveNext
        Loop
        lnLineaActual = lnLineaActual + 1
    Next

    xlsHoja.Range("A1:AZ6").Font.Bold = True
    xlsHoja.Range(xlsHoja.Cells(5, 4), xlsHoja.Cells(5, IIf(lnColumnaActual = 4, 5, lnColumnaActual) - 1)).MergeCells = True
    xlsHoja.Range("C:C").ColumnWidth = 35
    xlsHoja.Range("D:AZ").ColumnWidth = 18
    xlsHoja.Range("B5:AZ6").HorizontalAlignment = xlCenter
    xlsHoja.Range("B5:AZ6").VerticalAlignment = xlCenter
    
    xlsHoja.Range(xlsHoja.Cells(5, 2), xlsHoja.Cells(lnLineaActual - 1, IIf(lnColumnaActual = 4, 5, lnColumnaActual) - 1)).Borders.Weight = xlThin
    xlsHoja.Range(xlsHoja.Cells(5, 2), xlsHoja.Cells(6, IIf(lnColumnaActual = 4, 4, lnColumnaActual) - 1)).Interior.Color = RGB(199, 199, 199)
    xlsHoja.Range("A:AZ").NumberFormat = "#,##0"
    
    Set oBalance = Nothing
    Set rsInfoEstadColoc = Nothing
    Set MatCtas = Nothing
End Sub
Private Sub generaHojaExcelTotalAgencias_RptInfoEstadColocBCRP(ByRef xlsHoja As Worksheet, ByVal pdFecha As Date, ByVal rsAgencia As ADODB.Recordset)
    Dim oBalance As New DbalanceCont
    Dim rsInfoEstadColoc As New ADODB.Recordset
    Dim MatCtas As Variant, MatFinal As Variant
    Dim lnLineaActual As Integer, lnLineaAnterior As Integer, lnColumnaActual As Integer
    Dim tam As Long, iMat As Integer
    Dim bExisteCta As Boolean
    Dim lnMonto As Double
    Dim lnMesNombre As String
    Dim lnCtaCod As Long
    Dim lsCtaDesc As String
    Dim lnMontoSoles As Double, lnMontoDolares As Double
    Dim i As Integer, j As Integer

    lnMesNombre = dameNombreMes(Month(pdFecha), True)
    
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    xlsHoja.Range("A:A").ColumnWidth = 4
    xlsHoja.Range("B:B").ColumnWidth = 20
    xlsHoja.Range("C:C").ColumnWidth = 60
    xlsHoja.Range("D:D").ColumnWidth = 20
    xlsHoja.Range("E:E").ColumnWidth = 20
    xlsHoja.Range("F:F").ColumnWidth = 20
    xlsHoja.Range("B13").RowHeight = 3
    
    xlsHoja.Range("B1") = "BANCO CENTRAL DE RESERVA DEL PERÚ"
    xlsHoja.Range("B2") = "SEDE REGIONAL LORETO"
    xlsHoja.Range("F4") = "BCRP-LORETO"
    xlsHoja.Range("B6") = "INSTITUCIÓN FINANCIERA: CAJA MUNICIPAL DE AHORROS Y CREDITOS MAYNAS S.A."
    xlsHoja.Range("B7") = "SALDOS AL MES DE " & lnMesNombre & " " & Year(pdFecha)
    '''xlsHoja.Range("B9") = "Llenar en miles de nuevos soles" 'marg ers-044-2016
    xlsHoja.Range("B9") = "Llenar en miles de " & LCase(gcPEN_PLURAL) 'marg ers-044-2016
    xlsHoja.Range("B10") = "CUENTA"
    xlsHoja.Range("C10") = "DESCRIPCION"
    xlsHoja.Range("D10") = "MONEDA NACIONAL"
    xlsHoja.Range("E10") = "MONEDA EXTRANJERA 1/"
    xlsHoja.Range("F10") = "TOTAL M/N Y M/E"
    xlsHoja.Range("B11") = "MN/ ME"
    xlsHoja.Range("B12") = "14"
    xlsHoja.Range("C12") = "CREDITOS"
    
    xlsHoja.Range("B9:F9").MergeCells = True
    
    xlsHoja.Range("F4").Interior.Color = RGB(255, 255, 0)
    xlsHoja.Range("B10:F10").Interior.Color = RGB(199, 199, 199)
    xlsHoja.Range("B6:B7").Font.Color = RGB(0, 0, 128)
    
    xlsHoja.Range("A1:F14").Font.Bold = True
    xlsHoja.Range("B11:B12").HorizontalAlignment = xlLeft
    xlsHoja.Range("B9:F10").HorizontalAlignment = xlCenter
    
    xlsHoja.Range("B10:F10").Borders.Weight = xlThin

    'Para obtener distinct de todos las cuentas
    ReDim MatCtas(1 To 4, 0 To 0)
    Do While Not rsAgencia.EOF
        Set rsInfoEstadColoc = oBalance.RecuperaInfoEstadColocBCRP(Year(pdFecha), Month(pdFecha), rsAgencia!cAgeCod)
        Do While Not rsInfoEstadColoc.EOF
            bExisteCta = False
            For iMat = 1 To UBound(MatCtas, 2)
                If MatCtas(1, iMat) = rsInfoEstadColoc!cCtaContCod & IIf(Len(rsInfoEstadColoc!cCtaContCod) = 4, "00", "") Then
                    bExisteCta = True
                    Exit For
                End If
            Next
            If Not bExisteCta Then
                tam = UBound(MatCtas, 2) + 1
                ReDim Preserve MatCtas(1 To 4, 0 To tam)
                MatCtas(1, tam) = rsInfoEstadColoc!cCtaContCod & IIf(Len(rsInfoEstadColoc!cCtaContCod) = 4, "00", "")
                MatCtas(2, tam) = rsInfoEstadColoc!cCtaContDesc
            Else
                tam = iMat
            End If
            MatCtas(3, tam) = CDbl(MatCtas(3, tam)) + rsInfoEstadColoc!nTotSoles
            MatCtas(4, tam) = CDbl(MatCtas(4, tam)) + rsInfoEstadColoc!nTotDolares
            
            rsInfoEstadColoc.MoveNext
        Loop
        rsAgencia.MoveNext
    Loop
    'Ordenamos la Matriz
    For i = 1 To UBound(MatCtas, 2)
        For j = 1 To UBound(MatCtas, 2)
            If CLng(MatCtas(1, i)) < CLng(MatCtas(1, j)) Then
                lnCtaCod = MatCtas(1, i)
                lsCtaDesc = MatCtas(2, i)
                lnMontoSoles = MatCtas(3, i)
                lnMontoDolares = MatCtas(4, i)
                MatCtas(1, i) = MatCtas(1, j)
                MatCtas(2, i) = MatCtas(2, j)
                MatCtas(3, i) = MatCtas(3, j)
                MatCtas(4, i) = MatCtas(4, j)
                MatCtas(1, j) = lnCtaCod
                MatCtas(2, j) = lsCtaDesc
                MatCtas(3, j) = lnMontoSoles
                MatCtas(4, j) = lnMontoDolares
            End If
        Next
    Next
    'Regresamos a 4 digitos los que terminan en "00"
    For i = 1 To UBound(MatCtas, 2)
        If Right(MatCtas(1, i), 2) = "00" Then
            MatCtas(1, i) = Left(MatCtas(1, i), 4)
        End If
    Next
    
    lnLineaActual = 13
    lnLineaAnterior = 11

    For i = 1 To UBound(MatCtas, 2)
        If Len(MatCtas(1, i)) = 4 Then
            xlsHoja.Range("B" & lnLineaActual).RowHeight = 3
            lnLineaActual = lnLineaActual + 1
            xlsHoja.Range("B" & lnLineaActual & ":F" & lnLineaActual) = Left(MatCtas(1, i), 2) & "0" & Mid(MatCtas(1, i), 4, 1) & "/" & Left(MatCtas(1, i), 2) & "2" & Mid(MatCtas(1, i), 4, 1)
            xlsHoja.Range("B" & lnLineaActual & ":F" & lnLineaActual).Font.Bold = True
        Else
            xlsHoja.Range("B" & lnLineaActual & ":F" & lnLineaActual) = Left(MatCtas(1, i), 2) & "0" & Mid(MatCtas(1, i), 4, 1) & "." & Mid(MatCtas(1, i), 5, 2) & "/" & Left(MatCtas(1, i), 2) & "2" & Mid(MatCtas(1, i), 4, 1) & "." & Mid(MatCtas(1, i), 5, 2)
        End If
        xlsHoja.Range("C" & lnLineaActual) = MatCtas(2, i)
        'xlsHoja.Range("D" & lnLineaActual) = MatCtas(3, i)
        xlsHoja.Range("D" & lnLineaActual) = "=" & MatCtas(3, i) & "/1000" 'EJVG20130225
        'xlsHoja.Range("E" & lnLineaActual) = MatCtas(4, i)
        xlsHoja.Range("E" & lnLineaActual) = "=" & MatCtas(4, i) & "/1000" 'EJVG20130225
        xlsHoja.Range("F" & lnLineaActual) = "= D" & lnLineaActual & "+E" & lnLineaActual
        lnLineaActual = lnLineaActual + 1
    Next
    
    xlsHoja.Range("B" & lnLineaAnterior & ":C" & lnLineaActual - 1).HorizontalAlignment = xlLeft
    xlsHoja.Range("D" & lnLineaAnterior & ":F" & lnLineaActual - 1).Font.Color = RGB(0, 0, 128)
    xlsHoja.Range("D" & lnLineaAnterior & ":F" & lnLineaActual - 1).HorizontalAlignment = xlRight
    
    xlsHoja.Range("B" & lnLineaAnterior & ":F" & lnLineaActual - 1).Borders(xlEdgeTop).Weight = xlThin
    xlsHoja.Range("B" & lnLineaAnterior & ":F" & lnLineaActual - 1).Borders(xlEdgeBottom).Weight = xlThin
    xlsHoja.Range("B" & lnLineaAnterior & ":F" & lnLineaActual - 1).Borders(xlEdgeLeft).Weight = xlThin
    xlsHoja.Range("B" & lnLineaAnterior & ":F" & lnLineaActual - 1).Borders(xlEdgeRight).Weight = xlThin
    xlsHoja.Range("B" & lnLineaAnterior & ":F" & lnLineaActual - 1).Borders(xlInsideVertical).Weight = xlThin
    xlsHoja.Range("D" & lnLineaAnterior & ":F" & lnLineaActual - 1).NumberFormat = "#,##0"
    
    xlsHoja.Range("B" & lnLineaActual) = "1/ Equivalente en moneda nacional"

    Set rsInfoEstadColoc = Nothing
    Set oBalance = Nothing
    Set MatCtas = Nothing
End Sub
Private Function validaGenerar() As Boolean
    validaGenerar = True
    If Val(txtAnio.Text) <= 1900 Then
        MsgBox "Ud. debe especificar el año", vbInformation, "Aviso"
        txtAnio.SetFocus
        validaGenerar = False
        Exit Function
    End If
    If cboMes.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el mes", vbInformation, "Aviso"
        cboMes.SetFocus
        validaGenerar = False
        Exit Function
    End If
End Function
'END EJVG *******
