VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmNIIFBaseFormulasEEFFRep 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5205
   Icon            =   "frmBaseFormulasEEFFRep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Agencia"
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
      Left            =   80
      TabIndex        =   22
      Top             =   3000
      Width           =   5055
      Begin VB.CheckBox ckAgencia 
         Caption         =   "Agencia"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Frame frTipoReporte 
      Caption         =   "Tipo de Reporte"
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
      Left            =   80
      TabIndex        =   18
      Top             =   2160
      Width           =   5055
      Begin VB.OptionButton OptProyectado 
         Caption         =   "Proyectado"
         Height          =   375
         Left            =   3000
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton OptConsolidado 
         Caption         =   "Consolidado"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton OptMensual 
         Caption         =   "Mensual"
         Height          =   375
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CheckBox chkPeriodoCompara 
      Caption         =   "Periodo a Comparar"
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
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Frame fraPerComparar 
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
      Height          =   675
      Left            =   80
      TabIndex        =   13
      Top             =   720
      Width           =   5055
      Begin VB.TextBox txtAnioCompara 
         Alignment       =   1  'Right Justify
         Height          =   280
         Left            =   720
         MaxLength       =   4
         TabIndex        =   15
         Top             =   240
         Width           =   780
      End
      Begin VB.ComboBox cboMesCompara 
         Height          =   315
         ItemData        =   "frmBaseFormulasEEFFRep.frx":030A
         Left            =   2760
         List            =   "frmBaseFormulasEEFFRep.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   2280
         TabIndex        =   16
         Top             =   270
         Width           =   390
      End
   End
   Begin VB.Frame fraPerEvaluar 
      Caption         =   "Periodo a Evaluar"
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
      Height          =   675
      Left            =   80
      TabIndex        =   8
      Top             =   0
      Width           =   5055
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmBaseFormulasEEFFRep.frx":030E
         Left            =   2760
         List            =   "frmBaseFormulasEEFFRep.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         Height          =   280
         Left            =   720
         MaxLength       =   4
         TabIndex        =   9
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   2280
         TabIndex        =   12
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   270
         Width           =   375
      End
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
      Height          =   360
      Left            =   2640
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   3885
      Width           =   1575
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   6
      ToolTipText     =   "Generar Reporte Nota Estado"
      Top             =   3885
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Moneda"
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
      Height          =   675
      Left            =   80
      TabIndex        =   3
      Top             =   1440
      Width           =   5055
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmBaseFormulasEEFFRep.frx":0312
         Left            =   1080
         List            =   "frmBaseFormulasEEFFRep.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   270
         Width           =   675
      End
   End
   Begin MSComctlLib.ProgressBar BarraProgreso 
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   4320
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar EstadoBarra 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4290
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmNIIFBaseFormulasEEFFRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fsOpeCod As String
Private Type TCtaCont
    CuentaContable As String
    Saldo As Currency
    bSaldoA As Boolean
    bSaldoD As Boolean
End Type

Private Sub cboMes_LostFocus()
    cboMesCompara.ListIndex = cboMes.ListIndex
End Sub

Private Sub cmdGenerar_Click()
    Dim ldFechaEvalua As Date
    Dim ldFechaCompara As Date
    Dim lnMoneda As Integer
    If validaGenerar = False Then Exit Sub
    
    On Error GoTo ErrGenerar
    Screen.MousePointer = 11
    ldFechaEvalua = obtenerFechaFinMes(cboMes.ListIndex + 1, txtAnio.Text)
    ldFechaCompara = obtenerFechaFinMes(cboMesCompara.ListIndex + 1, txtAnioCompara.Text)
    lnMoneda = CInt(Trim(Right(cboMoneda.Text, 2)))
    If fsOpeCod = gContRepEstadoSitFinanEEFF1 And OptConsolidado.value = True And ckAgencia.value = 0 Then
        Call generarReporteEESSFF(fsOpeCod, ldFechaEvalua, IIf(chkPeriodoCompara.value = vbChecked, True, False), ldFechaCompara, lnMoneda, "ESTADO DE SITUACIÓN FINANCIERA CONSOLIDADO COMPARATIVO", "CUADRO   Nº  1", True, False, False, False)
    ElseIf fsOpeCod = gContRepEstadoSitFinanEEFF2 And OptConsolidado.value = True And ckAgencia.value = 0 Then
        Call generarReporteEESSFF(fsOpeCod, ldFechaEvalua, IIf(chkPeriodoCompara.value = vbChecked, True, False), ldFechaCompara, lnMoneda, "ESTADO DE RESULTADOS CONSOLIDADO", "CUADRO   Nº  2", True, False, False, False)
    ElseIf fsOpeCod = gContRepEstadoSitFinanEEFF2 And OptMensual.value = True And ckAgencia.value = 0 Then
        Call generarReporteEESSFF(fsOpeCod, ldFechaEvalua, IIf(chkPeriodoCompara.value = vbChecked, True, False), ldFechaCompara, lnMoneda, "RESULTADOS  MENSUALES  (" & CStr(Year(ldFechaEvalua)) & ")", "CUADRO   Nº  3", False, True, False, False)
    ElseIf fsOpeCod = gContRepEstadoSitFinanEEFF3 And OptMensual.value = True And ckAgencia.value = 0 Then
        Call generarReporteEESSFF(fsOpeCod, ldFechaEvalua, IIf(chkPeriodoCompara.value = vbChecked, True, False), ldFechaCompara, lnMoneda, "PRINCIPALES INDICADORES FINANCIEROS", "CUADRO   Nº  4", False, True, False, False)
    ElseIf fsOpeCod = gContRepEstadoSitFinanEEFF1 And OptProyectado.value = True And ckAgencia.value = 0 Then
        Call generarReporteEESSFF(fsOpeCod, ldFechaEvalua, IIf(chkPeriodoCompara.value = vbChecked, True, False), ldFechaCompara, lnMoneda, "ESTADO DE SITUACIÓN FINANCIERA CONSOLIDADO PROYECTADO VS EJECUTADO", "CUADRO   Nº  5", False, False, True, False)
    ElseIf fsOpeCod = gContRepEstadoSitFinanEEFF2 And OptProyectado.value = True And ckAgencia.value = 0 Then
        Call generarReporteEESSFF(fsOpeCod, ldFechaEvalua, IIf(chkPeriodoCompara.value = vbChecked, True, False), ldFechaCompara, lnMoneda, "ESTADO DE RESULTADOS CONSOLIDADO PROYECTADO VS EJECUTADO", "CUADRO   Nº  6", False, False, True, False)
    ElseIf fsOpeCod = gContRepEstadoSitFinanEEFF2 And ckAgencia.value = 1 Then
        Call generarReporteEESSFF(fsOpeCod, ldFechaEvalua, IIf(chkPeriodoCompara.value = vbChecked, True, False), ldFechaCompara, lnMoneda, "ESTADO DE RESULTADOS", "CUADRO Nº 7", True, False, False, True)
    ElseIf fsOpeCod = gContRepEstadoSitFinanEEFF1 And ckAgencia.value = 1 Then
        Call generarReporteEESSFF(fsOpeCod, ldFechaEvalua, IIf(chkPeriodoCompara.value = vbChecked, True, False), ldFechaCompara, lnMoneda, "ESTADO DE SITUACIÓN FINANCIERA POR  AGENCIA", "CUADRO Nº 8", False, False, True, True)
    Else
        MsgBox "No existe configurado el reporte según las opciones seleccionadas, verifique.", vbInformation, "Aviso"
    End If
 Screen.MousePointer = 0
    Exit Sub
ErrGenerar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub generarReporteEESSFF(ByVal psOpeCod As String, ByVal pdFechaEvalua As Date, ByVal pbComparaPeriodo As Boolean, pdFechaCompara As Date, ByVal pnMoneda As Integer, ByVal psReporteDesc As String, ByVal psNroReporte As String, ByVal pbConsolidado As Boolean, ByVal pbMensual As Boolean, ByVal pbProyectado As Boolean, ByVal pbAgencia As Boolean)
    Dim ix As Integer
    Dim ixAc As Integer
    Dim nPos As Integer
    Dim rsNotas As New ADODB.Recordset
    Dim oRep As New DRepFormula
    Dim oRepA As DRepFormula
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim fs As New Scripting.FileSystemObject
    Dim nPorcenPrincipal As Integer
    Dim rsNotasDet As New ADODB.Recordset
    Dim iCab As Long, iDet As Long
    Dim lsPath As String, lsArchivo As String
    Dim lbAbierto As Boolean
    Dim lnFilaActual As Integer, lnColumnaActual As Integer
    Dim lnNivelMax As Integer, lnUltimaColumna As Integer
    Dim lsNombreMesEvalua As String, lsNombreMesCompara As String
    Dim lnAnioEvalua As Integer, lnAnioCompara As Integer
    Dim lsFormula1Evalua As String, lsFormula1Compara As String
    Dim lsFormula2Evalua As String, lsFormula2Compara As String
    Dim lsFormula3Evalua As String, lsFormula3Compara As String
    Dim lsFormula4Evalua As String, lsFormula4Compara As String
    Dim lsFormula5Evalua As String, lsFormula5Compara As String
    Dim lnMontoFormula1Evalua As Currency, lnMontoFormula1Compara As Currency
    Dim lnMontoFormula2Evalua As Currency, lnMontoFormula2Compara As Currency
    Dim lnMontoFormula3Evalua As Currency, lnMontoFormula3Compara As Currency
    Dim lnMontoFormula4Evalua As Currency, lnMontoFormula4Compara As Currency
    Dim lnMontoFormula5Evalua As Currency, lnMontoFormula5Compara As Currency
    Dim lnMontoConsolidadoAnu As Currency
    Dim lnMontoConsolidadoMen As Currency
    Dim lnMontoConsolidadoAnuAc As Currency
    Dim nCantidadTrabajadores As Integer
    Dim nReNeAc As Currency
    Dim nReNeAcAA As Currency
    Dim nPromedio12Meses1 As Currency
    Dim nPromedio12Meses2 As Currency
    Dim nPromedio12Meses3 As Currency
    Dim lsFormula1 As String
    Dim lsFormula2 As String
    Dim lsFormula3 As String
    Dim lsDepartamento As String
    Dim nSumaFiCol As Currency
    Dim oBala As DbalanceCont
    Dim oRS As ADODB.Recordset
    Dim ldFechaProceso As Date
    Dim oBalInser As DbalanceCont
    Dim oRsInsert As ADODB.Recordset
    Dim lnMontoResultadoMes As Currency
    Dim nSumaDistCreditosDirectos As Currency
    Dim lnSumaActivoMes1 As Currency
    Dim lnSumaActivoMes2 As Currency
    Dim oMov As New DMov
    Dim sMovNro As String
    sMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If pbConsolidado = True And (psOpeCod = gContRepEstadoSitFinanEEFF1 Or psOpeCod = gContRepEstadoSitFinanEEFF2) And pbAgencia = False Then
        lsPath = App.path & "\FormatoCarta\EEFFC_EESSFF.xls"
        lsArchivo = "EEFFC_EESSFF_" & UCase(gsCodUser) & "_" & Format(gdFecSis, "yyyymmdd") & Format(Now, "hhmmss") & "_" & Format(pdFechaEvalua, "yyyymmdd") & ".xls"
    ElseIf pbMensual = True And (psOpeCod = gContRepEstadoSitFinanEEFF2) And pbAgencia = False Then
        lsPath = App.path & "\FormatoCarta\EEFFC_EESSFF_C3.xls"
        lsArchivo = "EEFFC_EESSFF_" & UCase(gsCodUser) & "_" & Format(gdFecSis, "yyyymmdd") & Format(Now, "hhmmss") & "_" & Format(pdFechaEvalua, "yyyymmdd") & ".xls"
    ElseIf pbMensual = True And (psOpeCod = gContRepEstadoSitFinanEEFF3) And pbAgencia = False Then
        lsPath = App.path & "\FormatoCarta\EEFFC_EESSFF_C4.xls"
        lsArchivo = "EEFFC_EESSFF_" & UCase(gsCodUser) & "_" & Format(gdFecSis, "yyyymmdd") & Format(Now, "hhmmss") & "_" & Format(pdFechaEvalua, "yyyymmdd") & ".xls"
    ElseIf pbProyectado = True And (psOpeCod = gContRepEstadoSitFinanEEFF1 Or psOpeCod = gContRepEstadoSitFinanEEFF2) And pbAgencia = False Then
        lsPath = App.path & "\FormatoCarta\EEFFC_EESSFF_C5.xls"
        lsArchivo = "EEFFC_EESSFF_" & UCase(gsCodUser) & "_" & Format(gdFecSis, "yyyymmdd") & Format(Now, "hhmmss") & "_" & Format(pdFechaEvalua, "yyyymmdd") & ".xls"
    ElseIf (psOpeCod = gContRepEstadoSitFinanEEFF1 Or psOpeCod = gContRepEstadoSitFinanEEFF2) And pbAgencia = True Then
        lsPath = App.path & "\FormatoCarta\EEFFC_EESSFF_C6.xls"
        lsArchivo = "EEFFC_EESSFF_" & UCase(gsCodUser) & "_" & Format(gdFecSis, "yyyymmdd") & Format(Now, "hhmmss") & "_" & Format(pdFechaEvalua, "yyyymmdd") & ".xls"
    End If
    'valida formato carta
    If Len(Dir(lsPath)) = 0 Then
        MsgBox "No se pudo encontrar el archivo: " & lsPath & "," & Chr(10) & "comuniquese con el Dpto. de TI", vbCritical, "Aviso"
        Exit Sub
    End If
    'verifica formato carta abierto
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
    
    nSumaDistCreditosDirectos = oRep.ObtenerDistCreditosDirectos(pdFechaCompara, Trim(Right(cboMoneda.Text, 30)))
    
    Set oRep = Nothing
    
    'Set rsNotas = oRep.ObtenerEEFF(psOpeCod)
    Set rsNotas = oRep.ObtenerEEFF(psOpeCod, Year(pdFechaEvalua), Month(pdFechaEvalua)) 'EJVG20140912
    
    If Not RSVacio(rsNotas) Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(lsPath)
        Set xlsHoja = xlsLibro.ActiveSheet
        
        BarraProgreso.value = 0
        BarraProgreso.Min = 0
        BarraProgreso.Max = rsNotas.RecordCount
        BarraProgreso.value = 0
        EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
        
        lsNombreMesEvalua = dameNombreMes(Month(pdFechaEvalua), True)
        lsNombreMesCompara = dameNombreMes(Month(pdFechaCompara), True)
        lnAnioEvalua = Year(pdFechaEvalua)
        lnAnioCompara = Year(pdFechaCompara)
        
        lnNivelMax = 1 'oRep.ObtenerUltimoNivelConfig(psOpeCod)
        lnUltimaColumna = lnNivelMax + IIf(pbComparaPeriodo, 10, 5) '5 por las formulas evalua y 10 con todo las comparativas
        
        lnFilaActual = 1
        lnColumnaActual = 10
        
        If pbConsolidado = True And (psOpeCod = gContRepEstadoSitFinanEEFF1 Or psOpeCod = gContRepEstadoSitFinanEEFF2) And pbAgencia = False Then
            xlsHoja.Range("A:AZ").NumberFormat = "#,##0"
            xlsHoja.Cells(1, 10) = psNroReporte
            xlsHoja.Cells(3, 2) = "ANALISIS VERTICAL Y HORIZONTAL (" & lsNombreMesEvalua & ") (" & lnAnioEvalua & ") - (" & lsNombreMesCompara & ") ( " & lnAnioCompara & ")"
            xlsHoja.Cells(5, 2) = psReporteDesc
            xlsHoja.Cells(6, 2) = "(" & Trim(Left(cboMoneda.Text, 30)) & ")"
            xlsHoja.Cells(9, 6) = Format(pdFechaEvalua, "DD/MM/YYYY")
            xlsHoja.Cells(9, 8) = Format(pdFechaCompara, "DD/MM/YYYY")
            For iCab = 1 To rsNotas.RecordCount
                xlsHoja.Cells(lnColumnaActual, 3) = rsNotas!cConceptoDesc
                xlsHoja.Cells(lnColumnaActual, 5) = rsNotas!cFormulaCons
                xlsHoja.Cells(lnColumnaActual, 6) = ObtenerResultadoFormula(pdFechaEvalua, IIf(DateDiff("d", "2012/12/31", pdFechaEvalua) <= 0, rsNotas!cFormulaAgen, rsNotas!cFormulaCons), Trim(Right(cboMoneda.Text, 5)))
                If rsNotas!nCorreInt = 1 And psOpeCod = gContRepEstadoSitFinanEEFF1 Then
                    lnSumaActivoMes1 = CCur(xlsHoja.Cells(lnColumnaActual, 6))
                End If
                If Trim(rsNotas!nNivelCod) = "1" Then
                    nPorcenPrincipal = lnColumnaActual
                    If rsNotas!nCorreInt > 60 And psOpeCod = gContRepEstadoSitFinanEEFF1 Then
                        'xlsHoja.Cells(lnColumnaActual, 7) = xlsHoja.Cells(lnColumnaActual, 6) / lnSumaActivoMes1
                        If lnSumaActivoMes1 = 0 Then
                            xlsHoja.Cells(lnColumnaActual, 7) = 0
                        Else
                            xlsHoja.Cells(lnColumnaActual, 7) = xlsHoja.Cells(lnColumnaActual, 6) / lnSumaActivoMes1
                        End If
                    Else
                        'xlsHoja.Cells(lnColumnaActual, 7) = xlsHoja.Cells(lnColumnaActual, 6) / xlsHoja.Cells(lnColumnaActual, 6)
                        If xlsHoja.Cells(lnColumnaActual, 6) = 0 Then 'EJVG20140818
                            xlsHoja.Cells(lnColumnaActual, 7) = 0
                        Else
                            xlsHoja.Cells(lnColumnaActual, 7) = xlsHoja.Cells(lnColumnaActual, 6) / xlsHoja.Cells(lnColumnaActual, 6)
                        End If
                    End If
                    xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 2), xlsHoja.Cells(lnColumnaActual, 11)).Interior.Color = RGB(0, 204, 255)
                    xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 2), xlsHoja.Cells(lnColumnaActual, 11)).Borders.LineStyle = 1
    
                Else
                    If Trim(rsNotas!nNivelCod) = "2" Then
                        xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 2), xlsHoja.Cells(lnColumnaActual, 11)).Interior.Color = RGB(255, 255, 153)
                        xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 2), xlsHoja.Cells(lnColumnaActual, 11)).Borders.LineStyle = 1
                    Else
                        xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 2), xlsHoja.Cells(lnColumnaActual, 11)).Borders.LineStyle = 1
                    End If
                     If rsNotas!nCorreInt > 60 And psOpeCod = gContRepEstadoSitFinanEEFF1 Then
                        'xlsHoja.Cells(lnColumnaActual, 7) = xlsHoja.Cells(lnColumnaActual, 6) / lnSumaActivoMes1
                        If lnSumaActivoMes1 = 0 Then 'EJVG20140818
                            xlsHoja.Cells(lnColumnaActual, 7) = 0
                        Else
                            xlsHoja.Cells(lnColumnaActual, 7) = xlsHoja.Cells(lnColumnaActual, 6) / lnSumaActivoMes1
                        End If
                     Else
                        'xlsHoja.Cells(lnColumnaActual, 7) = xlsHoja.Cells(lnColumnaActual, 6) / xlsHoja.Cells(nPorcenPrincipal, 6)
                        If xlsHoja.Cells(nPorcenPrincipal, 6) = 0 Then 'EJVG20140818
                            xlsHoja.Cells(lnColumnaActual, 7) = 0
                        Else
                            xlsHoja.Cells(lnColumnaActual, 7) = xlsHoja.Cells(lnColumnaActual, 6) / xlsHoja.Cells(nPorcenPrincipal, 6)
                        End If
                     End If
                End If
                xlsHoja.Cells(lnColumnaActual, 7).NumberFormat = "0.00%"
                xlsHoja.Cells(lnColumnaActual, 8) = ObtenerResultadoFormula(pdFechaCompara, IIf(DateDiff("d", "2012/12/31", pdFechaCompara) <= 0, rsNotas!cFormulaAgen, rsNotas!cFormulaCons), Trim(Right(cboMoneda.Text, 5)))
                If IIf(DateDiff("d", "2012/12/31", pdFechaCompara) <= 0, 0, 1) = 0 Then
                    If (rsNotas!nCorreInt = 28 Or rsNotas!nCorreInt = 29 Or rsNotas!nCorreInt = 49 Or rsNotas!nCorreInt = 55 Or rsNotas!nCorreInt = 66 Or rsNotas!nCorreInt = 67 Or rsNotas!nCorreInt = 68 Or rsNotas!nCorreInt = 39) And psOpeCod = gContRepEstadoSitFinanEEFF2 Then
                        If (rsNotas!nCorreInt = 28 Or rsNotas!nCorreInt = 29 Or rsNotas!nCorreInt = 49 Or rsNotas!nCorreInt = 55 Or rsNotas!nCorreInt = 66 Or rsNotas!nCorreInt = 39) Then
                            xlsHoja.Cells(lnColumnaActual, 8) = CCur(xlsHoja.Cells(lnColumnaActual, 8)) + nSumaDistCreditosDirectos
                        Else
                            xlsHoja.Cells(lnColumnaActual, 8) = CCur(xlsHoja.Cells(lnColumnaActual, 8)) - nSumaDistCreditosDirectos
                        End If
                    End If
                End If
                If rsNotas!nCorreInt = 1 And psOpeCod = gContRepEstadoSitFinanEEFF1 Then
                    lnSumaActivoMes2 = CCur(xlsHoja.Cells(lnColumnaActual, 8))
                End If
                If Trim(rsNotas!nNivelCod) = "1" Then
                    If rsNotas!nCorreInt > 60 And psOpeCod = gContRepEstadoSitFinanEEFF1 Then
                        xlsHoja.Cells(lnColumnaActual, 9) = xlsHoja.Cells(lnColumnaActual, 8) / lnSumaActivoMes2
                    Else
                        xlsHoja.Cells(lnColumnaActual, 9) = xlsHoja.Cells(lnColumnaActual, 8) / xlsHoja.Cells(lnColumnaActual, 8)
                    End If
                Else
                    If rsNotas!nCorreInt > 60 And psOpeCod = gContRepEstadoSitFinanEEFF1 Then
                        xlsHoja.Cells(lnColumnaActual, 9) = xlsHoja.Cells(lnColumnaActual, 8) / lnSumaActivoMes2
                    Else
                        xlsHoja.Cells(lnColumnaActual, 9) = xlsHoja.Cells(lnColumnaActual, 8) / xlsHoja.Cells(nPorcenPrincipal, 8)
                    End If
                End If
                xlsHoja.Cells(lnColumnaActual, 9).NumberFormat = "0.00%"
                xlsHoja.Cells(lnColumnaActual, 10) = xlsHoja.Cells(lnColumnaActual, 6) - xlsHoja.Cells(lnColumnaActual, 8)
                If Trim(rsNotas!nNivelCod) = "1" Then
                    xlsHoja.Cells(lnColumnaActual, 11) = xlsHoja.Cells(lnColumnaActual, 10) / xlsHoja.Cells(lnColumnaActual, 8)
                Else
                    'xlsHoja.Cells(lnColumnaActual, 11) = xlsHoja.Cells(lnColumnaActual, 10) / xlsHoja.Cells(nPorcenPrincipal, 8)
                    If xlsHoja.Cells(lnColumnaActual, 8) = 0 Then
                        xlsHoja.Cells(lnColumnaActual, 11) = 0
                    Else
                        xlsHoja.Cells(lnColumnaActual, 11) = xlsHoja.Cells(lnColumnaActual, 10) / xlsHoja.Cells(lnColumnaActual, 8)
                    End If
                End If
                xlsHoja.Cells(lnColumnaActual, 11).NumberFormat = "0.00%"
                lnColumnaActual = lnColumnaActual + 1
                BarraProgreso.value = iCab
                EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
                rsNotas.MoveNext
            Next iCab
        ElseIf pbMensual = True And (psOpeCod = gContRepEstadoSitFinanEEFF2) And pbAgencia = False Then
            
            xlsHoja.Cells(1, 29) = psNroReporte
            xlsHoja.Cells(3, 1) = psReporteDesc
            xlsHoja.Range(xlsHoja.Cells(7, 4), xlsHoja.Cells(7, 4)).Borders.LineStyle = 1
            xlsHoja.Range(xlsHoja.Cells(7, 4), xlsHoja.Cells(7, 4)).Interior.Color = RGB(0, 204, 255)
            xlsHoja.Cells(7, 4) = Format(pdFechaCompara, "mmm-yy")
            For ix = 1 To Month(pdFechaEvalua)
                nPos = ix  'IIf(ix = 1, 2, 2 * (ix)) - 1
                xlsHoja.Range(xlsHoja.Cells(7, 4 + nPos), xlsHoja.Cells(7, 5 + nPos)).Borders.LineStyle = 1
                xlsHoja.Range(xlsHoja.Cells(7, 4 + nPos), xlsHoja.Cells(7, 5 + nPos)).Interior.Color = RGB(0, 204, 255)
                xlsHoja.Cells(7, 4 + nPos) = UCase(DameMes(ix)) & "-" & Format(pdFechaEvalua, "yy")
'                xlsHoja.Cells(7, 5 + nPos) = "%"
                If ix = Month(pdFechaEvalua) Then
                    xlsHoja.Range(xlsHoja.Cells(7, 4 + nPos + 1), xlsHoja.Cells(7, 4 + nPos + 1)).Borders.LineStyle = 1
                    xlsHoja.Range(xlsHoja.Cells(7, 4 + nPos + 1), xlsHoja.Cells(7, 4 + nPos + 1)).Interior.Color = RGB(0, 204, 255)
                    xlsHoja.Cells(7, 4 + nPos + 1) = UCase(DameMes(ix)) & "-" & Format(pdFechaEvalua, "yy")
'                    xlsHoja.Cells(7, 5 + nPos + 2) = "%"
                    
                    xlsHoja.Range(xlsHoja.Cells(6, 4 + nPos + 1), xlsHoja.Cells(6, 4 + nPos + 1)).Borders.LineStyle = 1
                    xlsHoja.Range(xlsHoja.Cells(6, 4 + nPos + 1), xlsHoja.Cells(6, 4 + nPos + 1)).Interior.Color = RGB(255, 255, 153)
                    xlsHoja.Cells(6, 4 + nPos + 1) = "Acumulado"
                End If
            Next ix
            lnColumnaActual = 9
            For iCab = 1 To rsNotas.RecordCount
                lnMontoConsolidadoAnu = 0
                lnMontoConsolidadoAnuAc = 0
                lnMontoConsolidadoMen = 0
                xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 2), xlsHoja.Cells(lnColumnaActual, 4)).Borders.LineStyle = 1
                xlsHoja.Cells(lnColumnaActual, 2) = rsNotas!cConceptoDesc
                xlsHoja.Cells(lnColumnaActual, 3) = rsNotas!cFormulaCons
                xlsHoja.Cells(lnColumnaActual, 4) = ObtenerResultadoFormula(pdFechaCompara, IIf(DateDiff("d", "2012/12/31", pdFechaCompara) <= 0, rsNotas!cFormulaAgen, rsNotas!cFormulaCons), Trim(Right(cboMoneda.Text, 5)))
                For ix = 1 To Month(pdFechaEvalua)
                    nPos = ix 'IIf(ix = 1, 2, 2 * (ix)) - 1
                    xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 4 + nPos), xlsHoja.Cells(lnColumnaActual, 4 + nPos)).Borders.LineStyle = 1
                    xlsHoja.Cells(lnColumnaActual, 4 + nPos) = ObtenerResultadoFormula(obtenerFechaFinMes(ix, Year(pdFechaEvalua)), IIf(DateDiff("d", "2012/12/31", obtenerFechaFinMes(ix, Year(pdFechaEvalua))) <= 0, rsNotas!cFormulaAgen, rsNotas!cFormulaCons), Trim(Right(cboMoneda.Text, 5)))
                    lnMontoConsolidadoAnu = CCur(xlsHoja.Cells(lnColumnaActual, 4 + nPos))
                    xlsHoja.Cells(lnColumnaActual, 4 + nPos) = CCur(xlsHoja.Cells(lnColumnaActual, 4 + nPos)) - lnMontoConsolidadoAnuAc
                    lnMontoConsolidadoAnuAc = lnMontoConsolidadoAnu
                    xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 4 + nPos), xlsHoja.Cells(lnColumnaActual, 4 + nPos)).NumberFormat = "#,##0.00;-#,##0.00"
                'Fin
                    lnMontoConsolidadoMen = lnMontoConsolidadoMen + CCur(xlsHoja.Cells(lnColumnaActual, 4 + nPos))
                    If ix = Month(pdFechaEvalua) Then
                        xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 4 + nPos + 1), xlsHoja.Cells(7, 4 + nPos + 1)).Borders.LineStyle = 1
                        xlsHoja.Cells(lnColumnaActual, 4 + nPos + 1) = lnMontoConsolidadoMen
                        lnMontoConsolidadoMen = 0
                    End If
                Next ix
                
                lnColumnaActual = lnColumnaActual + 1
                rsNotas.MoveNext
                
                BarraProgreso.value = iCab
                EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
            Next iCab
        ElseIf pbMensual = True And (psOpeCod = gContRepEstadoSitFinanEEFF3) And pbAgencia = False Then
            Dim lnMes As Integer
            Dim lnYear As Integer
            
            Dim lnMesA As Integer
            Dim lnYearA As Integer
            
            Dim lnMesAc As Integer
            Dim lnYearAc As Integer
            Dim lnlogicoGrabar As Integer
            
            Set oBalInser = New DbalanceCont
            Set oRsInsert = New ADODB.Recordset
            lnlogicoGrabar = 0
            Set oRsInsert = oBalInser.ObtenerHistEEFFMensual(pdFechaEvalua, psOpeCod, Trim(Right(cboMoneda.Text, 5)))
            If Not (oRsInsert.BOF Or oRsInsert.EOF) Then
             If MsgBox("El reporte ya fue generado, Desea generarlo otra vez ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                lnlogicoGrabar = 0
             Else
                lnlogicoGrabar = 1
             End If
            Else
                lnlogicoGrabar = 1
            End If
            Set oBalInser = Nothing
            Set oRsInsert = Nothing
            lnColumnaActual = 7
            xlsHoja.Cells(3, 40) = psNroReporte
            xlsHoja.Cells(4, 2) = psReporteDesc
'            lnMes = Format(DateAdd("M", 1, DateAdd("Y", -2, pdFechaEvalua)), "MM")
'            lnYear = Format(DateAdd("M", 1, DateAdd("Y", -1, pdFechaEvalua)), "YYYY")
'            lnYear = IIf(lnMes = 1, lnYear, lnYear - 1)
            lnMes = Format(DateAdd("Y", -12, pdFechaEvalua), "MM")
            lnYear = Format(DateAdd("M", -12, pdFechaEvalua), "YYYY")
            'lnYear = IIf(lnMes = 1, lnYear, lnYear - 1)
            For ix = 1 To 13
                    xlsHoja.Cells(6, 2 + ix) = "'" & UCase(DameMes(lnMes)) & "-" & Right(CStr(lnYear), 2)
                    lnMes = lnMes + 1
                    lnYear = IIf(lnMes = "13", lnYear + 1, lnYear)
                    lnMes = IIf(lnMes = "13", 1, lnMes)
            Next ix
            For iCab = 1 To rsNotas.RecordCount
                
'                lnMes = Format(DateAdd("M", 1, DateAdd("Y", -1, pdFechaEvalua)), "MM")
'                lnYear = Format(DateAdd("M", 1, DateAdd("Y", -1, pdFechaEvalua)), "YYYY")
'                lnYear = IIf(lnMes = 1, lnYear, lnYear - 1)
                lnMes = Format(DateAdd("Y", -12, pdFechaEvalua), "MM")
                lnYear = Format(DateAdd("M", -12, pdFechaEvalua), "YYYY")
                
                xlsHoja.Cells(lnColumnaActual, 2) = rsNotas!cConceptoDesc
                xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 2), xlsHoja.Cells(lnColumnaActual, 2)).Borders.LineStyle = 1
                    For ix = 1 To 13
                        If Trim(rsNotas!cFormulaCons) <> "" Then
                            Set oBala = New DbalanceCont
                            nCantidadTrabajadores = oBala.ObtenerCantidadPersonal(IIf(Len(lnMes) = 1, "0" & CStr(lnMes), CStr(lnMes)), CStr(lnYear))
                            If rsNotas!nCorreInt = 16 Or rsNotas!nCorreInt = 18 Then
                                xlsHoja.Cells(lnColumnaActual, 2 + ix) = ObtenerResultadoFormula(obtenerFechaFinMes(lnMes, lnYear), IIf(DateDiff("d", "2012/12/31", obtenerFechaFinMes(lnMes, lnYear)) <= 0, rsNotas!cFormulaAgen, rsNotas!cFormulaCons), Trim(Right(cboMoneda.Text, 5))) / nCantidadTrabajadores
                            ElseIf rsNotas!nCorreInt = 17 Then
                                xlsHoja.Cells(lnColumnaActual, 2 + ix) = (ObtenerResultadoFormula(obtenerFechaFinMes(lnMes, lnYear), IIf(DateDiff("d", "2012/12/31", obtenerFechaFinMes(lnMes, lnYear)) <= 0, rsNotas!cFormulaAgen, rsNotas!cFormulaCons), Trim(Right(cboMoneda.Text, 5))) / lnMes) * 12
                            ElseIf rsNotas!nCorreInt = 24 Or rsNotas!nCorreInt = 25 Then
                                Set oRepA = New DRepFormula
                                If ix = 13 Then
                                    lnMesAc = IIf(lnMes + 1 = 13, 1, lnMes + 1)
                                    lnYearAc = IIf(lnMes = 12, lnYear, lnYear - 1)
                                    nPromedio12Meses1 = 0
                                    nPromedio12Meses2 = 0
                                    nPromedio12Meses3 = 0
                                    For ixAc = 1 To 12
                                        If rsNotas!nCorreInt = 24 Then
                                            lsFormula2 = IIf(DateDiff("d", "2012/12/31", obtenerFechaFinMes(lnMesAc, lnYearAc)) <= 0, oRepA.ObtenerEEFFxId("760108", 62, "A"), oRepA.ObtenerEEFFxId("760108", 62, ""))
                                            nPromedio12Meses2 = nPromedio12Meses2 + (ObtenerResultadoFormula(obtenerFechaFinMes(lnMesAc, lnYearAc), lsFormula2, Trim(Right(cboMoneda.Text, 5)))) / 12
                                        Else
                                            lsFormula3 = IIf(DateDiff("d", "2012/12/31", obtenerFechaFinMes(lnMesAc, lnYearAc)) <= 0, oRepA.ObtenerEEFFxId("760108", 1, "A"), oRepA.ObtenerEEFFxId("760108", 1, ""))
                                            nPromedio12Meses3 = nPromedio12Meses3 + (ObtenerResultadoFormula(obtenerFechaFinMes(lnMesAc, lnYearAc), lsFormula3, Trim(Right(cboMoneda.Text, 5)))) / 12
                                        End If
                                        lnMesAc = lnMesAc + 1
                                        lnYearAc = IIf(lnMesAc = "13", lnYearAc + 1, lnYearAc)
                                        lnMesAc = IIf(lnMesAc = "13", 1, lnMesAc)
                                    Next ixAc
                                    lnMesA = 12
                                    lnYearA = lnYear - 1
                                    ldFechaProceso = DateAdd("m", -12, obtenerFechaFinMes(lnMes, lnYear))
                                    ldFechaProceso = DateAdd("d", -Day(ldFechaProceso), ldFechaProceso)
                                    lsFormula1 = IIf(DateDiff("d", "2012/12/31", ldFechaProceso) <= 0, oRepA.ObtenerEEFFxId("760109", 71, "A"), oRepA.ObtenerEEFFxId("760109", 71, ""))
                                    nPromedio12Meses1 = (ObtenerResultadoFormula(ldFechaProceso, lsFormula1, Trim(Right(cboMoneda.Text, 5))))
                                    nReNeAc = ObtenerResultadoFormula(obtenerFechaFinMes(lnMesA, lnYearA), IIf(DateDiff("d", "2012/12/31", obtenerFechaFinMes(lnMesA, lnYearA)) <= 0, oRepA.ObtenerEEFFxId("760109", 71, "A"), oRepA.ObtenerEEFFxId("760109", 71, "")), Trim(Right(cboMoneda.Text, 5)))
                                    nReNeAcAA = ObtenerResultadoFormula(obtenerFechaFinMes(lnMes, lnYear - 1), IIf(DateDiff("d", "2012/12/31", obtenerFechaFinMes(lnMes, lnYear - 1)) <= 0, oRepA.ObtenerEEFFxId("760109", 71, "A"), oRepA.ObtenerEEFFxId("760109", 71, "")), Trim(Right(cboMoneda.Text, 5)))
                                    lnMontoResultadoMes = ObtenerResultadoFormula(obtenerFechaFinMes(lnMes, lnYear), IIf(DateDiff("d", "2012/12/31", obtenerFechaFinMes(lnMes, lnYear)) <= 0, oRepA.ObtenerEEFFxId("760109", 71, "A"), oRepA.ObtenerEEFFxId("760109", 71, "")), Trim(Right(cboMoneda.Text, 5)))
                                    If rsNotas!nCorreInt = 24 Then
                                        xlsHoja.Cells(lnColumnaActual, 2 + ix) = ((nReNeAc + lnMontoResultadoMes - nPromedio12Meses1) / nPromedio12Meses2) * 100
                                    Else
                                        xlsHoja.Cells(lnColumnaActual, 2 + ix) = ((nReNeAc + lnMontoResultadoMes - nPromedio12Meses1) / nPromedio12Meses3) * 100
                                        'xlsHoja.Cells(lnColumnaActual, 2 + ix) = ((nReNeAc + ObtenerResultadoFormula(obtenerFechaFinMes(lnMes, lnYear), IIf(DateDiff("d", "2012/12/31", obtenerFechaFinMes(lnMes, lnYear)) <= 0, oRepA.ObtenerEEFFxId("760109", 71, "A"), oRepA.ObtenerEEFFxId("760109", 71, "")), Trim(Right(cboMoneda.Text, 5))) - nReNeAcAA) / nPromedio12Meses) * 100
                                    End If
                                    Set oRepA = Nothing
                                Else
                                    Set oBalInser = New DbalanceCont
                                    Set oRsInsert = New ADODB.Recordset
                                    ldFechaProceso = obtenerFechaFinMes(lnMes, lnYear)
                                    xlsHoja.Cells(lnColumnaActual, 2 + ix) = oBalInser.ObtenerHistEEFFMensualxCorrelativo(ldFechaProceso, psOpeCod, Trim(Right(cboMoneda.Text, 5)), rsNotas!nCorreInt)
                                End If
                            Else
                                ldFechaProceso = obtenerFechaFinMes(lnMes, lnYear)
                                If DateDiff("d", pdFechaEvalua, ldFechaProceso) = 0 Or Not (rsNotas!nCorreInt = 22) Then
                                    xlsHoja.Cells(lnColumnaActual, 2 + ix) = ObtenerResultadoFormula(obtenerFechaFinMes(lnMes, lnYear), IIf(DateDiff("d", "2012/12/31", obtenerFechaFinMes(lnMes, lnYear)) <= 0, rsNotas!cFormulaAgen, rsNotas!cFormulaCons), Trim(Right(cboMoneda.Text, 5)))
                                Else
                                    Set oBalInser = New DbalanceCont
                                    Set oRsInsert = New ADODB.Recordset
                                    xlsHoja.Cells(lnColumnaActual, 2 + ix) = oBalInser.ObtenerHistEEFFMensualxCorrelativo(ldFechaProceso, psOpeCod, Trim(Right(cboMoneda.Text, 5)), rsNotas!nCorreInt)
                                End If
                            End If
                            Set oBala = Nothing
                            xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 2 + ix), xlsHoja.Cells(lnColumnaActual, 2 + ix)).NumberFormat = "#,##0.00;-#,##0.00"
                            
                            If lnlogicoGrabar = 1 Then
                            If Month(pdFechaEvalua) = lnMes And Year(pdFechaEvalua) = lnYear Then
                                 Set oBalInser = New DbalanceCont
                                 Call oBalInser.InsertarHistEEFFMensual(sMovNro, pdFechaEvalua, rsNotas!nCorreInt, CCur(xlsHoja.Cells(lnColumnaActual, 2 + ix)), psOpeCod, Trim(Right(cboMoneda.Text, 5)))
                                 Set oBalInser = Nothing
                            End If
                            End If
                        End If
                        'obtenerFechaFinMes(cboMesCompara.ListIndex + 1, txtAnioCompara.Text)
                        xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 2 + ix), xlsHoja.Cells(lnColumnaActual, 2 + ix)).Borders.LineStyle = 1
                        lnMes = lnMes + 1
                        lnYear = IIf(lnMes = "13", lnYear + 1, lnYear)
                        lnMes = IIf(lnMes = "13", 1, lnMes)
                    Next ix
                
                lnColumnaActual = lnColumnaActual + 1
                rsNotas.MoveNext
                
                BarraProgreso.value = iCab
                EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
            Next iCab
        ElseIf pbProyectado = True And (psOpeCod = gContRepEstadoSitFinanEEFF1 Or psOpeCod = gContRepEstadoSitFinanEEFF2) And pbAgencia = False Then
            lnColumnaActual = 8
            xlsHoja.Cells(1, 6) = psNroReporte
            xlsHoja.Cells(4, 1) = psReporteDesc
            xlsHoja.Cells(5, 1) = "(" & Trim(Left(cboMoneda.Text, 30)) & ")"
            xlsHoja.Cells(7, 4) = Format(pdFechaEvalua, "mmm-yy")
            For iCab = 1 To rsNotas.RecordCount
                xlsHoja.Cells(lnColumnaActual, 2) = rsNotas!cConceptoDesc
                xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 1), xlsHoja.Cells(lnColumnaActual, 7)).Borders.LineStyle = 1
                xlsHoja.Cells(lnColumnaActual, 6) = ObtenerResultadoFormula(pdFechaEvalua, IIf(DateDiff("d", "2012/12/31", pdFechaEvalua) <= 0, rsNotas!cFormulaAgen, rsNotas!cFormulaCons), Trim(Right(cboMoneda.Text, 5)))
                If Trim(rsNotas!nNivelCod) = "1" Then
                    nPorcenPrincipal = lnColumnaActual
                    xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 1), xlsHoja.Cells(lnColumnaActual, 7)).Interior.Color = RGB(0, 204, 255)
                    If xlsHoja.Cells(lnColumnaActual, 6) = 0# Then
                        xlsHoja.Cells(lnColumnaActual, 7) = 0
                    Else
                        xlsHoja.Cells(lnColumnaActual, 7).Formula = "=F" & lnColumnaActual & "/D" & lnColumnaActual ' = xlsHoja.Cells(lnColumnaActual, 6) / xlsHoja.Cells(lnColumnaActual, 6)
                    End If
                    
                Else
                    If xlsHoja.Cells(nPorcenPrincipal, 6) = 0# Then
                        xlsHoja.Cells(lnColumnaActual, 7) = 0
                    Else
                        xlsHoja.Cells(lnColumnaActual, 7).Formula = "=F" & lnColumnaActual & "/D" & lnColumnaActual '  xlsHoja.Cells(lnColumnaActual, 6) / xlsHoja.Cells(nPorcenPrincipal, 6)
                    End If
                    
                End If
                xlsHoja.Cells(lnColumnaActual, 4) = rsNotas!nProyeccion 'EJVG20140912
                xlsHoja.Cells(lnColumnaActual, 7).NumberFormat = "0.00%"
                lnColumnaActual = lnColumnaActual + 1
                rsNotas.MoveNext
                
                BarraProgreso.value = iCab
                EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
            Next iCab
        ElseIf (psOpeCod = gContRepEstadoSitFinanEEFF1 Or psOpeCod = gContRepEstadoSitFinanEEFF2) And pbAgencia = True Then
            Dim nPosInicial As Integer
            Dim nPosFinal As Integer
            Dim cCodDep As String
            Dim nPosPrincipal As Integer
            Dim nPosUltimaAgencia As Integer
            lnColumnaActual = 8
            lnFilaActual = 3
            nPosFinal = 3
            xlsHoja.Cells(1, 16) = psNroReporte
            xlsHoja.Cells(3, 1) = psReporteDesc
            xlsHoja.Cells(5, 1) = "(" & Trim(Left(cboMoneda.Text, 30)) & ")"
            xlsHoja.Cells(4, 1) = "AL " & Format(pdFechaEvalua, "DD") & " DE " & Trim(Left(cboMes.Text, 30)) & " DE " & Format(pdFechaEvalua, "YYYY")
            
            Set oBala = New DbalanceCont
            Set oRS = New ADODB.Recordset
            Set oRS = oBala.RecuperarDepartamAgencia
            If Not (oRS.BOF Or oRS.EOF) Then
                Do While Not oRS.EOF
                    xlsHoja.Cells(7, lnFilaActual) = oRS!cAgeDescripcion
                    xlsHoja.Range(xlsHoja.Cells(7, lnFilaActual), xlsHoja.Cells(7, lnFilaActual)).Borders.LineStyle = 1
                    If cCodDep <> oRS!cDepaCod And lnFilaActual > 3 Then
                        If nPosFinal <> lnFilaActual - 1 Then
                            xlsHoja.Range(Mid(Replace(xlsHoja.Cells(6, nPosFinal).AddressLocal, "$", ""), 1, 2) & ":" & Mid(Replace(xlsHoja.Cells(6, lnFilaActual - 1).AddressLocal, "$", ""), 1, 2)).Merge
                        End If
                        xlsHoja.Range(xlsHoja.Cells(6, nPosFinal), xlsHoja.Cells(6, lnFilaActual - 1)).Borders.LineStyle = 1
                        xlsHoja.Cells(6, nPosFinal) = lsDepartamento
                        nPosFinal = lnFilaActual
                    End If
                    lnFilaActual = lnFilaActual + 1
                    cCodDep = oRS!cDepaCod
                    lsDepartamento = oRS!cUbiGeoDescripcion
                    oRS.MoveNext
                Loop
                If lnFilaActual > 3 Then
                    If nPosFinal <> lnFilaActual - 1 Then
                       xlsHoja.Range(Mid(Replace(xlsHoja.Cells(6, nPosFinal).AddressLocal, "$", ""), 1, 2) & ":" & Mid(Replace(xlsHoja.Cells(6, lnFilaActual - 1).AddressLocal, "$", ""), 1, 2)).Merge
                    End If
                    xlsHoja.Cells(6, nPosFinal) = lsDepartamento
                    xlsHoja.Range(xlsHoja.Cells(6, nPosFinal), xlsHoja.Cells(6, lnFilaActual - 1)).Borders.LineStyle = 1
                End If
                xlsHoja.Range(xlsHoja.Cells(6, lnFilaActual), xlsHoja.Cells(7, lnFilaActual)).Borders.LineStyle = 1
                xlsHoja.Cells(6, lnFilaActual) = "Acumulado"
                xlsHoja.Cells(7, lnFilaActual) = Format(pdFechaEvalua, "YYYY/MM/DD")
                
            End If
            lnFilaActual = 3
            nPosFinal = 3
            oRS.MoveFirst
            For iCab = 1 To rsNotas.RecordCount
                xlsHoja.Cells(lnColumnaActual, 2) = rsNotas!cConceptoDesc
                xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, 2), xlsHoja.Cells(lnColumnaActual, 2)).Borders.LineStyle = 1
                oRS.MoveFirst
                lnFilaActual = 3
                nSumaFiCol = 0
                If Not (oRS.BOF Or oRS.EOF) Then
                    nPosUltimaAgencia = 3 + oRS.RecordCount
                    Do While Not oRS.EOF
                        If oRS!cAgeCod = "01" Then
                            nPosPrincipal = lnFilaActual
                        End If
                        If rsNotas!nCorreInt = 70 Then
                            xlsHoja.Cells(lnColumnaActual, lnFilaActual) = xlsHoja.Cells(lnColumnaActual - 1, lnFilaActual) * 0.3 * -1
                        Else
                            xlsHoja.Cells(lnColumnaActual, lnFilaActual) = ObtenerResultadoFormula(pdFechaEvalua, rsNotas!cFormulaAgen, Trim(Right(cboMoneda.Text, 5)), oRS!cAgeCod)
                        End If
                        xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, lnFilaActual), xlsHoja.Cells(lnColumnaActual, lnFilaActual)).Borders.LineStyle = 1
                        nSumaFiCol = nSumaFiCol + xlsHoja.Cells(lnColumnaActual, lnFilaActual)
                        xlsHoja.Cells(lnColumnaActual, lnFilaActual + 1) = nSumaFiCol
                        xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, lnFilaActual + 1), xlsHoja.Cells(lnColumnaActual, lnFilaActual + 1)).Borders.LineStyle = 1
                        xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, lnFilaActual), xlsHoja.Cells(lnColumnaActual, lnFilaActual + 1)).NumberFormat = "#,##0.00;-#,##0.00"
                        lnFilaActual = lnFilaActual + 1
                        oRS.MoveNext
                    Loop
                    lnColumnaActual = lnColumnaActual + 1
                    rsNotas.MoveNext
                End If
                BarraProgreso.value = iCab
                EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
            Next iCab
            'Llenar nuevamente la agencia principal
            
            BarraProgreso.value = 0
            BarraProgreso.Min = 0
            BarraProgreso.Max = rsNotas.RecordCount
            BarraProgreso.value = 0
            EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
            
            lnColumnaActual = 8
            rsNotas.MoveFirst
            For iCab = 1 To rsNotas.RecordCount
                nSumaFiCol = xlsHoja.Cells(lnColumnaActual, nPosUltimaAgencia) - xlsHoja.Cells(lnColumnaActual, nPosPrincipal)
                If rsNotas!nCorreInt = 70 Then
                    xlsHoja.Cells(lnColumnaActual, nPosPrincipal) = xlsHoja.Cells(lnColumnaActual - 1, nPosPrincipal) * 0.3 * -1
                Else
                    xlsHoja.Cells(lnColumnaActual, nPosPrincipal) = ObtenerResultadoFormula(pdFechaEvalua, rsNotas!cFormulaCons, Trim(Right(cboMoneda.Text, 5)), "") - nSumaFiCol
                End If
                xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, nPosPrincipal), xlsHoja.Cells(lnColumnaActual, nPosPrincipal)).Borders.LineStyle = 1
                nSumaFiCol = nSumaFiCol + xlsHoja.Cells(lnColumnaActual, nPosPrincipal)
                xlsHoja.Cells(lnColumnaActual, nPosUltimaAgencia) = nSumaFiCol
                xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, nPosPrincipal + 1), xlsHoja.Cells(lnColumnaActual, nPosPrincipal + 1)).Borders.LineStyle = 1
                xlsHoja.Range(xlsHoja.Cells(lnColumnaActual, nPosPrincipal), xlsHoja.Cells(lnColumnaActual, nPosPrincipal + 1)).NumberFormat = "#,##0.00;-#,##0.00"
                lnColumnaActual = lnColumnaActual + 1
                rsNotas.MoveNext
                BarraProgreso.value = iCab
                EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"

            Next iCab
            'Fin llenado
        End If
        xlsHoja.SaveAs App.path & "\Spooler\" & lsArchivo
        MsgBox "Reporte se generó satisfactoriamente en " & App.path & "\Spooler\" & lsArchivo, vbInformation, "Aviso"
        xlsAplicacion.Visible = True
        xlsAplicacion.Windows(1).Visible = True
        EstadoBarra.Panels(1) = "Proceso Terminado"
    Else
        MsgBox "No existe la configuración respectiva para generar el presente Reporte", vbInformation, "Aviso"
    End If
    
    Set rsNotas = Nothing
    Set rsNotasDet = Nothing
    Set oRep = Nothing
    Set xlsHoja = Nothing
    Set xlsLibro = Nothing
    Set xlsAplicacion = Nothing
End Sub
Private Function DameMes(ByVal pnMes As Integer)
    If pnMes = "1" Then
        DameMes = "ENE"
    ElseIf pnMes = "2" Then
        DameMes = "FEB"
    ElseIf pnMes = "3" Then
        DameMes = "MAR"
    ElseIf pnMes = "4" Then
        DameMes = "ABR"
    ElseIf pnMes = "5" Then
        DameMes = "MAY"
    ElseIf pnMes = "6" Then
        DameMes = "JUN"
    ElseIf pnMes = "7" Then
        DameMes = "JUL"
    ElseIf pnMes = "8" Then
        DameMes = "AGO"
    ElseIf pnMes = "9" Then
        DameMes = "SET"
    ElseIf pnMes = "10" Then
        DameMes = "OCT"
    ElseIf pnMes = "11" Then
        DameMes = "NOV"
    ElseIf pnMes = "12" Then
        DameMes = "DIC"
    End If
End Function
Private Sub Form_Load()
    CentraForm Me
    cargarMoneda
    cargarMes
    txtAnio.Text = Year(gdFecSis)
    txtAnioCompara.Text = Year(gdFecSis) - 1
    cboMes.ListIndex = IndiceListaCombo(cboMes, Month(gdFecSis))
    cboMesCompara.ListIndex = IndiceListaCombo(cboMesCompara, Month(gdFecSis))
End Sub
Public Sub Inicio(ByVal psOpeCod As String, ByVal psOpeDesc As String)
    fsOpeCod = psOpeCod
    Caption = UCase(psOpeDesc)
    OptConsolidado.value = True
    '.Caption = ""
    Show 1
End Sub
Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdGenerar.SetFocus
    End If
End Sub
Private Sub OptConsolidado_Click()
    OptConsolidado.value = True
    OptMensual.value = False
End Sub
Private Sub OptMensual_Click()
    OptConsolidado.value = False
    OptMensual.value = True
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cboMes.SetFocus
    End If
End Sub
Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkPeriodoCompara.SetFocus
    End If
End Sub

Private Sub txtAnio_LostFocus()
    If Val(txtAnio.Text) > 0 Then
        txtAnioCompara.Text = Val(txtAnio.Text) - 1
    End If
End Sub

Private Sub txtAnioCompara_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cboMesCompara.SetFocus
    End If
End Sub
Private Sub cboMesCompara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboMoneda.SetFocus
    End If
End Sub
Private Sub chkMuestraResultAnioAnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub
Private Sub chkPeriodoCompara_Click()
    If chkPeriodoCompara.value = vbChecked Then
        fraPerComparar.Enabled = True
    Else
        fraPerComparar.Enabled = False
    End If
End Sub
Private Sub chkPeriodoCompara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkPeriodoCompara.value = vbChecked Then
            txtAnioCompara.SetFocus
        Else
            cboMoneda.SetFocus
        End If
    End If
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub cargarMoneda()
    cboMoneda.AddItem "UNIFICADO" & Space(200) & "0"
    '''cboMoneda.AddItem "NUEVOS SOLES" & Space(200) & "1" 'MARG ERS044-2016
    cboMoneda.AddItem StrConv(gcPEN_PLURAL, vbUpperCase) & Space(200) & "1" 'MARG ERS044-2016
    cboMoneda.AddItem "DOLARES" & Space(200) & "2"
End Sub
Private Sub cargarMes()
    cboMes.AddItem "ENERO" & Space(200) & "1"
    cboMes.AddItem "FEBRERO" & Space(200) & "2"
    cboMes.AddItem "MARZO" & Space(200) & "3"
    cboMes.AddItem "ABRIL" & Space(200) & "4"
    cboMes.AddItem "MAYO" & Space(200) & "5"
    cboMes.AddItem "JUNIO" & Space(200) & "6"
    cboMes.AddItem "JULIO" & Space(200) & "7"
    cboMes.AddItem "AGOSTO" & Space(200) & "8"
    cboMes.AddItem "SEPTIEMBRE" & Space(200) & "9"
    cboMes.AddItem "OCTUBRE" & Space(200) & "10"
    cboMes.AddItem "NOVIEMBRE" & Space(200) & "11"
    cboMes.AddItem "DICIEMBRE" & Space(200) & "12"
    
    cboMesCompara.AddItem "ENERO" & Space(200) & "1"
    cboMesCompara.AddItem "FEBRERO" & Space(200) & "2"
    cboMesCompara.AddItem "MARZO" & Space(200) & "3"
    cboMesCompara.AddItem "ABRIL" & Space(200) & "4"
    cboMesCompara.AddItem "MAYO" & Space(200) & "5"
    cboMesCompara.AddItem "JUNIO" & Space(200) & "6"
    cboMesCompara.AddItem "JULIO" & Space(200) & "7"
    cboMesCompara.AddItem "AGOSTO" & Space(200) & "8"
    cboMesCompara.AddItem "SEPTIEMBRE" & Space(200) & "9"
    cboMesCompara.AddItem "OCTUBRE" & Space(200) & "10"
    cboMesCompara.AddItem "NOVIEMBRE" & Space(200) & "11"
    cboMesCompara.AddItem "DICIEMBRE" & Space(200) & "12"
End Sub
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

Private Function ObtenerResultadoFormula(ByVal pdFecha As Date, ByVal psFormula As String, ByVal pnMoneda As Integer, Optional psAgencia As String = "") As Currency
    Dim oBal As New DbalanceCont
    Dim oNBal As New NBalanceCont
    Dim oFormula As New NInterpreteFormula
    Dim lsFormula As String, lsTmp As String, lsTmp1 As String, lsCadFormula As String
    Dim MatDatos() As TCtaCont
    Dim i As Long, j As Long, nCtaCont As Long
    Dim sTempAD As String
    Dim nPosicion As Integer
    Dim signo As String
    Dim LsSigno As String
    lsFormula = Trim(psFormula)
    ReDim MatDatos(0)
    nCtaCont = 0
    lsTmp = ""
    lsFormula = Replace(lsFormula, "M", pnMoneda)
    sTempAD = ""
    For i = 1 To Len(lsFormula)
        If (Mid(Trim(lsFormula), i, 1) >= "0" And Mid(Trim(lsFormula), i, 1) <= "9") Then
            lsTmp = lsTmp + Mid(Trim(lsFormula), i, 1)
        Else
            If Len(lsTmp) > 0 Then
                nCtaCont = nCtaCont + 1
                ReDim Preserve MatDatos(nCtaCont)
                
                MatDatos(nCtaCont).CuentaContable = lsTmp
                
                If MatDatos(nCtaCont).CuentaContable = "100" Or MatDatos(nCtaCont).CuentaContable = "1000" Then
                    MatDatos(nCtaCont).Saldo = MatDatos(nCtaCont).CuentaContable
                Else
                    If Trim(psAgencia) = "" Then
                        MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual2(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True)
                    Else
                        MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensualxAgencia(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True, psAgencia)
                    End If
                End If
                
                If nCtaCont > 1 Then
                    If Mid(Trim(lsFormula), i, 1) = ")" Then
                        nPosicion = 0
                    Else
                        nPosicion = i
                    End If
                End If
                    If sTempAD = "" Then
                        If nCtaCont = 1 Then
                            If ((i - Len(Trim(lsTmp))) - 3) > 1 Then
                                sTempAD = Mid(Trim(lsFormula), (i - Len(Trim(lsTmp))) - 3, 2)
                            Else
                                sTempAD = ""
                            End If
                        Else
                            sTempAD = Mid(Trim(lsFormula), (i - Len(MatDatos(nCtaCont).CuentaContable)) - 3, 2)
                        End If
                    End If
                
                If sTempAD = "SA" Or sTempAD = "SD" Then
                    MatDatos(nCtaCont).CuentaContable = DepuraSaldoAD(MatDatos(nCtaCont).CuentaContable)
                    If sTempAD = "SA" Then
                        MatDatos(nCtaCont).bSaldoA = True
                        MatDatos(nCtaCont).bSaldoD = False
                    Else
                        MatDatos(nCtaCont).bSaldoA = False
                        MatDatos(nCtaCont).bSaldoD = True
                    End If
                    Else
                        MatDatos(nCtaCont).bSaldoA = False
                        MatDatos(nCtaCont).bSaldoD = False
                End If
            End If
            If nPosicion = 0 Then
               sTempAD = ""
            End If
            lsTmp = ""
        End If
    Next i
    If Len(lsTmp) > 0 Then
        nCtaCont = nCtaCont + 1
        ReDim Preserve MatDatos(nCtaCont)
        MatDatos(nCtaCont).CuentaContable = lsTmp
        'MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual2(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True)
        If MatDatos(nCtaCont).CuentaContable = "100" Or MatDatos(nCtaCont).CuentaContable = "1000" Then
            MatDatos(nCtaCont).Saldo = MatDatos(nCtaCont).CuentaContable
        Else
            If Trim(psAgencia) = "" Then
                MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual2(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True)
            Else
                MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensualxAgencia(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True, psAgencia)
            End If
        End If
    End If
    'Genero la formula en cadena
    lsTmp = ""
    lsCadFormula = ""
    Dim nEncontrado As Integer
    For i = 1 To Len(lsFormula)
        If (Mid(Trim(lsFormula), i, 1) >= "0" And Mid(Trim(lsFormula), i, 1) <= "9") Then
            lsTmp = lsTmp + Mid(Trim(lsFormula), i, 1)
        Else
            If Len(lsTmp) > 0 Then
                For j = 1 To nCtaCont
                    nEncontrado = 0
                    If MatDatos(j).CuentaContable = lsTmp Then
                            
                            If MatDatos(j).bSaldoA = True Or MatDatos(j).bSaldoD = True Then
                                MatDatos(j).Saldo = oNBal.CalculaSaldoBECuentaAD(MatDatos(j).CuentaContable, pnMoneda, MatDatos(j).bSaldoA, CStr(pnMoneda), Trim(psAgencia), Format(pdFecha, "YYYY"), Format(pdFecha, "MM"))
                                nEncontrado = 1
                            End If
                                If Left(Format(MatDatos(j).Saldo, "#0.00"), 1) = "-" And (Right(lsCadFormula, 1) = "-" Or Right(lsCadFormula, 1) = "+") Then
                                    
                                    If Right(Trim(lsCadFormula), 1) = "-" Or Right(Trim(lsCadFormula), 1) = "+" Then
                                        If Right(Trim(lsCadFormula), 1) = "-" And MatDatos(j).Saldo >= 0 Then
                                            LsSigno = "-"
                                        ElseIf Right(Trim(lsCadFormula), 1) = "-" And MatDatos(j).Saldo < 0 Then
                                            LsSigno = "+"
                                        ElseIf Right(Trim(lsCadFormula), 1) = "+" And MatDatos(j).Saldo >= 0 Then
                                            LsSigno = "+"
                                        ElseIf Right(Trim(lsCadFormula), 1) = "+" And MatDatos(j).Saldo < 0 Then
                                            LsSigno = "-"
                                        End If
                                    Else
                                        LsSigno = ""
                                    End If
                                    If LsSigno = "" Then
                                        lsCadFormula = Left(lsCadFormula, Len(lsCadFormula) - 1) & Format(MatDatos(j).Saldo, "#0.00")
                                    Else
                                        lsCadFormula = Left(lsCadFormula, Len(lsCadFormula) - 1) & LsSigno & Format(Abs(MatDatos(j).Saldo), "#0.00")
                                    End If
                                    nEncontrado = 1
                                Else
                                    lsCadFormula = lsCadFormula & Format(MatDatos(j).Saldo, "#0.00")
                                    nEncontrado = 1
                                End If
                        Exit For
                    End If
                Next j
            End If
            lsTmp = ""
            If nEncontrado = 1 Or (Mid(Trim(lsFormula), i, 1) = "S" Or Mid(Trim(lsFormula), i, 1) = "A" Or Mid(Trim(lsFormula), i, 1) = "D") Then
            lsCadFormula = lsCadFormula & Mid(Trim(lsFormula), i, 1)
            Else
            lsCadFormula = lsCadFormula & "" & Mid(Trim(lsFormula), i, 1)
            End If
        End If
    Next
    If Len(lsTmp) > 0 Then
        For j = 1 To nCtaCont
           If MatDatos(j).CuentaContable = lsTmp Then
               'lsCadFormula = lsCadFormula & Format(MatDatos(j).Saldo, "#0.00")
               If Left(Format(MatDatos(j).Saldo, "#0.00"), 1) = "-" And (Right(lsCadFormula, 1) = "-" Or Right(lsCadFormula, 1) = "+") Then
                    lsCadFormula = Left(lsCadFormula, Len(lsCadFormula) - 1) & Format(MatDatos(j).Saldo, "#0.00")
                    nEncontrado = 1
                Else
                    lsCadFormula = lsCadFormula & Format(MatDatos(j).Saldo, "#0.00")
                    nEncontrado = 1
                End If
               Exit For
           End If
        Next j
    End If
    lsCadFormula = Replace(Replace(lsCadFormula, "SA", ""), "SD", "")
    ObtenerResultadoFormula = oFormula.ExprANum(lsCadFormula)
    Set oBal = Nothing
    Set oFormula = Nothing
End Function

Private Function DepuraSaldoAD(ByVal sCta As String) As String
Dim i As Integer
Dim Cad As String
    Cad = ""
    For i = 1 To Len(sCta)
        If Mid(sCta, i, 1) >= "0" And Mid(sCta, i, 1) <= "9" Then
            Cad = Cad + Mid(sCta, i, 1)
        End If
    Next i
    DepuraSaldoAD = Cad
End Function
Private Function validaGenerar() As Boolean
    validaGenerar = False
    If Val(txtAnio.Text) <= 1900 Then
        MsgBox "Ud. debe ingresar un año válido", vbInformation, "Aviso"
        EnfocaControl txtAnio
        Exit Function
    End If
    If cboMes.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar un mes", vbInformation, "Aviso"
        EnfocaControl cboMes
        Exit Function
    End If
    If chkPeriodoCompara.value = 1 Then
        If Val(txtAnioCompara.Text) <= 1900 Then
            MsgBox "Ud. debe ingresar un año válido", vbInformation, "Aviso"
            EnfocaControl txtAnioCompara
            Exit Function
        End If
        If cboMesCompara.ListIndex = -1 Then
            MsgBox "Ud. debe seleccionar un mes", vbInformation, "Aviso"
            EnfocaControl cboMesCompara
            Exit Function
        End If
    End If
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la moneda", vbInformation, "Aviso"
        EnfocaControl cboMoneda
        Exit Function
    End If
    validaGenerar = True
End Function


