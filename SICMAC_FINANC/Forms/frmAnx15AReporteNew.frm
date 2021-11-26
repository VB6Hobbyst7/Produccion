VERSION 5.00
Begin VB.Form frmAnx15AReporteNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Tesorería y Liquidéz Anexo No 15A"
   ClientHeight    =   1575
   ClientLeft      =   4410
   ClientTop       =   3975
   ClientWidth     =   4350
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboMes 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame fraRango 
      Caption         =   "Balance Diario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtAnio 
         Height          =   330
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblGuion 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1900
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.TextBox txtFechaReporte15A 
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Text            =   "Fec15A"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "frmAnx15AReporteNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsArchivo As String
Dim lbExcel As Boolean
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim ldFecha As Date
Dim oBarra As clsProgressBar
Dim oCon As DConecta

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAnio.SetFocus
    End If
End Sub

Private Sub txtAnio_GotFocus()
    fEnfoque txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Function ValidaDatos(pdFecha As Date) As Boolean
    If (Month(pdFecha) > 12) Or (Month(pdFecha) <= 2) Then
        If txtAnio.Text = "" Or IIf(CInt(cboMes.ListIndex) = 0, txtAnio.Text > Year(pdFecha) Or txtAnio.Text <= Year(DateAdd("yyyy", -1, pdFecha)), txtAnio.Text >= Year(pdFecha) Or txtAnio.Text < Year(DateAdd("yyyy", -1, pdFecha))) Then
            MsgBox "Debe ingresar el año correspondiente", vbInformation, "Aviso"
            txtAnio.SetFocus
            Exit Function
        End If

        If (CInt(cboMes.ListIndex) + 1) >= Month(pdFecha) And (CInt(cboMes.ListIndex) + 1) < 11 Then
            MsgBox "El mes ingresado es incorrecto", vbInformation, "Aviso"
                cboMes.SetFocus
            Exit Function
        End If
        If Day(pdFecha) < 15 Then
            If (CInt(cboMes.ListIndex) + 1) = Month(DateAdd("m", -1, pdFecha)) Or (CInt(cboMes.ListIndex) + 1) < Month(DateAdd("m", -2, pdFecha)) Then
                If MsgBox("Desea actualizar las cuentas con el Balance del Mes Ingresado?", vbInformation + vbYesNo, "Atención") = vbNo Then
                    cboMes.SetFocus
                    Exit Function
                End If
            End If
        Else
            If (CInt(cboMes.ListIndex) + 1) < Month(DateAdd("m", -1, pdFecha)) Then
                If MsgBox("Desea actualizar las cuentas con el Balance del Mes Ingresado?", vbInformation + vbYesNo, "Atención") = vbNo Then
                    cboMes.SetFocus
                    Exit Function
                End If
            End If
        End If

    Else
        If txtAnio.Text = "" Or txtAnio.Text > Year(pdFecha) Or txtAnio.Text < Year(pdFecha) Then
            MsgBox "Debe ingresar el año correspondiente", vbInformation, "Aviso"
            txtAnio.SetFocus
            Exit Function
        End If

        If (CInt(cboMes.ListIndex) + 1) >= Month(pdFecha) Then
            MsgBox "El mes ingresado es incorrecto", vbInformation, "Aviso"
                cboMes.SetFocus
            Exit Function
        Else
            If Day(pdFecha) >= 15 Then
                If (CInt(cboMes.ListIndex) + 1) < Month(DateAdd("m", -1, pdFecha)) Then
                    If MsgBox("Desea actualizar las cuentas con el Balance del Mes Ingresado?", vbInformation + vbYesNo, "Atención") = vbNo Then
                        cboMes.SetFocus
                        Exit Function
                    End If
                    'MsgBox "El mes ingresado es incorrecto", vbInformation, "Aviso"
                    'cboMes.SetFocus
                    'Exit Function
                End If
            Else
                If (CInt(cboMes.ListIndex) + 1) < Month(DateAdd("m", -2, pdFecha)) Or (CInt(cboMes.ListIndex) + 1) = Month(DateAdd("m", -1, pdFecha)) Then
                    If MsgBox("Desea actualizar las cuentas con el Balance del Mes Ingresado?", vbInformation + vbYesNo, "Atención") = vbNo Then
                        cboMes.SetFocus
                        Exit Function
                    End If
                    'MsgBox "El mes ingresado es incorrecto", vbInformation, "Aviso"
                    'cboMes.SetFocus
                    'Exit Function
                End If
            End If
        End If
    End If
    ValidaDatos = True
End Function 'NAGL 20170426

Private Function ValRegValorizacionDiaria(pdFecha As Date) As Boolean
    Dim DAnxVal As New DAnexoRiesgos
    Dim psRegistro As String
    psRegistro = DAnxVal.ObtenerRegValorizacionDiaria(pdFecha)
    If psRegistro = "0" Then
        If MsgBox("La Valorización Diaria no ha sido ingresada, Desea continuar?", vbInformation + vbYesNo, "Atención") = vbNo Then
            cboMes.SetFocus
            Exit Function
        End If
    End If
    ValRegValorizacionDiaria = True
End Function

Public Sub ImprimeAnexo15A(psOpeCod As String, psMoneda As String, pdFecha As Date, Optional ByVal pnTipoCambio As Currency = 1, Optional ByVal pnValorEncajeMesAnterior As Currency = 0, Optional ByVal pnValorEncajeMesAnteriorME As Currency = 0, Optional ByVal pnValorToseMN As Currency = 0, Optional ByVal pnValorToseME As Currency = 0, Optional ByVal pnTotalFdoEncaje As Currency = 0, Optional ByVal pnTotalFdoEncajeME As Currency = 0, Optional ByVal PnPatrimonioEfectivo As Currency = 1, Optional ByVal psTasaSoles As String = "", Optional ByVal psTasaDolares As String = "", Optional ByVal pnToseAcumuladoMN As Currency = 0, Optional ByVal pnToseAcumuladoME As Currency = 0, Optional ByVal pnDepositoBCRAcumuladoMN As Currency = 0, Optional ByVal pnDepositoBCRAcumuladoME As Currency = 0, Optional ByVal pnEstadoEncajeAcumuladoMN As Currency = 0, Optional ByVal pnEstadoEncajeAcumuladoME As Currency = 0)
    Dim rs As New ADODB.Recordset
    Dim oGen As New DGeneral
    Dim psMesBalanceDiario As String
    Dim psAnioBalanceDiario As String
    Dim pdFechaFinMesAnt As String

    If Day(pdFecha) >= 15 Then
        pdFechaFinMesAnt = DateAdd("d", -Day(pdFecha), pdFecha)
    Else
        pdFechaFinMesAnt = DateAdd("d", -Day(DateAdd("m", -1, pdFecha)), DateAdd("m", -1, pdFecha))
    End If
    psMesBalanceDiario = Month(pdFechaFinMesAnt)
    psAnioBalanceDiario = Year(pdFechaFinMesAnt)
    
    Set rs = oGen.GetConstante(1010)
    While Not rs.EOF
        cboMes.AddItem rs.Fields(0) & space(50) & rs.Fields(1)
        rs.MoveNext
    Wend
    txtFechaReporte15A.Text = Format(pdFecha, "dd/MM/YYYY")
    txtAnio.Text = psAnioBalanceDiario
    cboMes.ListIndex = CInt(psMesBalanceDiario) - 1
    Me.Show 1
End Sub

Private Sub cmdGenerar_Click()
    Dim psMoneda As String, psMesBalanceDiario As String, psAnioBalanceDiario As String
    Dim pdFecha As Date
    Dim pdFechaSist As Date 'NAGL 20170904
    Dim lsPermitGenfromBitac As String 'NAGL 20170904
    Dim oDbalanceCont As New DbalanceCont  'NAGL 20170904

    psMoneda = "1"
    pdFecha = txtFechaReporte15A.Text
    psAnioBalanceDiario = txtAnio.Text
    pdFechaSist = Format(gdFecSis, "dd/mm/yyyy") 'NAGL 20170904

    If ValRegValorizacionDiaria(pdFecha) Then
        If (CInt(cboMes.ListIndex) + 1) <= 9 Then
            psMesBalanceDiario = "0" & CStr(CInt(cboMes.ListIndex) + 1)
        Else
            psMesBalanceDiario = CStr(CInt(cboMes.ListIndex) + 1)
        End If
        If ValidaDatos(pdFecha) Then 'Valida Datos con respecto a la fecha ingresada
            lsPermitGenfromBitac = oDbalanceCont.ObtenerPermiteGenerarfromBitacora15A_15B(pdFecha, pdFechaSist) 'NAGL 20170904
            '***NAGL 20170904
            If lsPermitGenfromBitac = "NO" Then
                Call GeneraEstadisticaDiaria(psMoneda, pdFecha, psMesBalanceDiario, psAnioBalanceDiario)
            Else
                Call GeneraEstadisticaDiariafromBitacora(psMoneda, pdFecha, pdFechaSist, psMesBalanceDiario, psAnioBalanceDiario)
            End If
            '***FIN NAGL 20170904
        End If
    End If
End Sub 'NAGL 20170425

'Private Sub GeneraEstadisticaDiaria(psOpeCod As String, psMoneda As String, pdFecha As Date, Optional ByVal pnTipoCambio As Currency = 1, Optional ByVal pnValorEncajeMesAnterior As Currency = 0, Optional ByVal pnValorEncajeMesAnteriorME As Currency = 0, Optional ByVal pnValorToseMN As Currency = 0, Optional ByVal pnValorToseME As Currency = 0, Optional ByVal pnTotalFdoEncaje As Currency = 0, Optional ByVal pnTotalFdoEncajeME As Currency = 0, Optional ByVal PnPatrimonioEfectivo As Currency = 1, Optional ByVal psTasaSoles As String = "", Optional ByVal psTasaDolares As String = "", Optional ByVal pnToseAcumuladoMN As Currency = 0, Optional ByVal pnToseAcumuladoME As Currency = 0, Optional ByVal pnDepositoBCRAcumuladoMN As Currency = 0, Optional ByVal pnDepositoBCRAcumuladoME As Currency = 0, Optional ByVal pnEstadoEncajeAcumuladoMN As Currency = 0, Optional ByVal pnEstadoEncajeAcumuladoME As Currency = 0)
Private Sub GeneraEstadisticaDiaria(psMoneda As String, pdFecha As Date, psMesBalanceDiario As String, psAnioBalanceDiario As String) 'NAGL 20170425
    Dim fs As New Scripting.FileSystemObject
    Dim lsMoneda As String
    Dim lsTotalActivos() As String
    Dim lsTotalPasivos() As String
    Dim lsTotalValores() As String 'NAGL20170407
    Dim lsTotalRatioLiquidez() As String
    Dim lbExisteHoja As Boolean
    Dim lsTotalesActivos() As String
    Dim lsTotalesPasivos() As String
    Dim i As Long
    Dim Y1 As Integer, Y2 As Integer
    Dim Yvalor1 As Integer, Yvalor2 As Integer  'NAGL20170407
    Dim lnFila As Integer
    Dim lnFilaFondosCaja As Integer
    Dim nMonto1 As Currency
    Dim nMonto2 As Currency
    Dim nMontoMN As Currency, nMontoME As Currency 'NAGL20170407
    Dim nTasaSubasta1 As Double, nTasaSubasta2 As Double 'NAGL20170407
    Dim oEst As New NEstadisticas
    Dim oDbalanceCont As DbalanceCont
    
    Dim nSaldoDiario1 As Currency
    Dim nSaldoDiario2 As Currency
    Dim nCajaFondo As Integer
    Dim oCambio As nTipoCambio
    Dim nTipoCambioAn As Currency
    Dim lnTipoCambioFC As Currency
    Dim lnTipoCambioFCMA As Currency

    Dim dFechaAnte As Date
    Dim nDia As Integer
    Dim lnPosInicial As Integer
    Dim ln2_1EncaExigible As Integer
    Dim ln3_1ObligSujetasEncajePos As Integer
    Dim ln3_1ObligSujetasEncajeMN As Currency
    Dim ln3_1ObligSujetasEncajeME As Currency
    Dim lnDiasToseBaseRef As Currency
    Dim lnToseBaseExigiBCRPMN As Currency
    Dim lnToseBaseExigiBCRPME As Currency
    Dim nTotalObligSugEncajMN As Currency
    Dim nTotalObligSugEncajME As Currency
    Dim nTotalTasaBaseEncajMN_DADiario As Currency
    Dim nTotalTasaBaseEncajME_DADiario As Currency
    Dim nlnAdeducirAhorroMN As Currency
    Dim nlnAdeducirAhorroME As Currency
    Dim lnTasaEncajeMN As Double
    Dim lnTasaEncajeME As Double
    Dim ix As Integer
    Dim lnToseRGMN As Currency
    Dim lnToseRGME As Currency
    Dim ldFechaPro As Date
    Dim nSaldoCajaDiarioMesAnteriorMN As Currency
    Dim nSaldoCajaDiarioMesAnteriorME As Currency

    Dim lnToTalTotalCajaFondosMN As Currency
    Dim lnToTalTotalCajaFondosME As Currency
    Dim lnToTalOME As Currency
    Dim lnToTalOMN As Currency
    Dim lnToTalCajaFondosMN As Currency
    Dim lnToTalCajaFondosME As Currency
    Dim lnTotalSaldoBCRPAnexoDiarioMN As Currency
    Dim lnTotalSaldoBCRPAnexoDiarioME As Currency
    Dim lnTotalObligacionesAlDiaMN As Currency
    Dim lnTotalObligacionesAlDiaME As Currency
    Dim nSubValor1 As Currency
    Dim nSubValor2 As Currency
    Dim pdFechaFinDeMes As Date, pdFechaBalanceDiario As Date 'NAGL 20170415
    Dim pdFechaFinDeMesMA As Date
    Dim nTotalTOSEMN As Currency
    Dim nTotalTOSEME As Currency
    Dim nTasaEncajeME As Double

    Dim nTasaEncajeMarginalME As Currency '***NAGL20170407
    Dim nItemEncaje As Integer
    Dim oValor As New DAnexoRiesgos
    Dim rsvalor As New ADODB.Recordset
    Dim oCtaIf As New NCajaCtaIF
    Dim rsDetalleCtas As New ADODB.Recordset 'NAGL20170407

    Dim nContar As Integer
    Dim nContar2 As Integer
    Dim X As Integer
    Dim nSaldoPFN As Double, nSaldoPFME As Double 'NAGL20170407
    Dim nSumarTotalMN As Double
    Dim nSumarMN As Double  '*********NAGL20170407
    Dim nTipoCambioBalance As Currency
    Dim nsDiv As Double
    Dim nsDiv2 As Double, nsDivME As Double
    Dim nPromedio As Double
    Dim nObligMNDiario1 As Currency, nObligMEDiario2 As Currency
    Dim oDAnexos As New DAnexoRiesgos
    'Dim nUltimoTOSE

    'INICIO VAPA 20170909
    Dim lnRatioLiquidezMN As Double
    Dim lnRatioLiquidezME As Double
    Dim lnRatioLAjusRecursosPrestadosMN As Double
    Dim lnRatioLAjusRecursosPrestadosME As Double
    Dim lnRatioInversionesLiquidasMN As Double
    Dim lnEncajeExigALMN As Double
    Dim lnEncajeExigALME As Double
    Dim lnTotalaMN As Double
    Dim lnTotalaME As Double
    'VAPA 20170909 END

    On Error GoTo GeneraExcelErr 'NAGL 20170425
    
    Set oBarra = New clsProgressBar
    Unload Me
    oBarra.ShowForm frmReportes
    oBarra.Max = 10
    oBarra.Progress 0, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    pdFechaFinDeMes = DateAdd("d", -Day(DateAdd("m", 1, pdFecha)), DateAdd("m", 1, pdFecha))
    pdFechaFinDeMesMA = DateAdd("d", -Day(pdFechaFinDeMes), pdFechaFinDeMes)
    nDia = Day(pdFecha)
    Set oDbalanceCont = New DbalanceCont
    
    lsArchivo = App.path & "\SPOOLER\" & "Anx15A_" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx" 'NAGL 20170415
    lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    ExcelAddHoja "Anx15A", xlLibro, xlHoja1
    
    Set oCambio = New nTipoCambio
    If CInt(psMesBalanceDiario) < 9 Then
        pdFechaBalanceDiario = CDate("01" & "/" & "0" & CStr(CInt(psMesBalanceDiario) + 1) & "/" & psAnioBalanceDiario)
    ElseIf CInt(psMesBalanceDiario) = 12 Then
        pdFechaBalanceDiario = CDate("01" & "/" & "01" & "/" & CStr(CInt(psAnioBalanceDiario) + 1))
    Else
        pdFechaBalanceDiario = CDate("01" & "/" & CStr(CInt(psMesBalanceDiario) + 1) & "/" & psAnioBalanceDiario)
    End If

    nTipoCambioBalance = Format(oCambio.EmiteTipoCambio(pdFechaBalanceDiario, TCFijoDia), "#,##0.0000")  'NAGL 20170425

    If Month(pdFecha) = Month(DateAdd("d", 1, pdFecha)) Then
        lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(pdFecha, TCFijoDia), "#,##0.0000")
    Else
        lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(DateAdd("d", -1, pdFecha), TCFijoDia), "#,##0.0000") 'SE CAMBIO DE DateAdd("d", 1, pdFecha) A DateAdd("d", -1, pdFecha)
    End If

    nTipoCambioAn = lnTipoCambioFC

    ldFecha = pdFecha
    lsMoneda = Mid(gsOpeCod, 3, 1)
    xlHoja1.PageSetup.Zoom = 100
    For i = 2 To 30
        If i <> 6 Then
            xlHoja1.Range(xlHoja1.Cells(i, 3), xlHoja1.Cells(i, 5)).Merge True
        End If
    Next

    ReDim lsTotalActivos(2)
    ReDim lsTotalPasivos(2)
    ReDim lsTotalValores(2) '****NAGL
    ReDim lsSumaTotalBCRP(2)   '*******NAGL
    ReDim lsTotalRatioLiquidez(2)

    xlHoja1.Range("A1:R500").Font.Size = 10 'NAGL 20190614 Cambio de 9 a 10
    xlHoja1.Range("A1:R500").Font.Name = "Arial Narrow" 'NAGL 20190614

    xlHoja1.Range("A1").ColumnWidth = 7
    xlHoja1.Range("B1").ColumnWidth = 30
    xlHoja1.Range("C1").ColumnWidth = 18 '36
    xlHoja1.Range("D1:E1").ColumnWidth = 17
    xlHoja1.Range("F1:G1").ColumnWidth = 14.29 '15
    xlHoja1.Range("H1").ColumnWidth = 13 '10

    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(6, 8)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(10, 8)).HorizontalAlignment = xlCenter

    xlHoja1.Range("B1:B50").HorizontalAlignment = xlLeft

    lnFila = 1
    xlHoja1.Cells(lnFila, 2) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "ANEXO Nº 15A"
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "REPORTE DE TESORERIA Y POSICION DE LIQUIDEZ"
    'lnFila = lnFila + 1
    'xlHoja1.Cells(lnFila, 3) = "(EN NUEVOS SOLES)"
    lnFila = lnFila + 2
    xlHoja1.Cells(lnFila, 2) = "EMPRESA: " & gsNomCmac
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).MergeCells = True

    xlHoja1.Cells(lnFila, 6) = "Fecha: " & Format(pdFecha, "dd mmmm yyyy")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).HorizontalAlignment = xlLeft

    lnFila = lnFila + 2

    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)

    '*******************I RATIOS DE LIQUIDEZ**************************

    xlHoja1.Cells(lnFila, 2) = "I RATIOS DE LIQUIDEZ"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Borders.LineStyle = xlContinuous
    lnFila = lnFila + 1
    'Comentado by NAGL 20190614
    'xlHoja1.Cells(lnFila, 6) = "MONEDA ": xlHoja1.Cells(lnFila, 7) = "MONEDA "
    'lnFila = lnFila + 1
    'xlHoja1.Cells(lnFila, 6) = "NACIONAL": xlHoja1.Cells(lnFila, 7) = "EXTRANJERA"

    '**Agregado by NAGL 20190614
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila + 1, 2)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila + 1, 2)).VerticalAlignment = xlJustify

    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila + 1, 5)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila + 1, 5)).VerticalAlignment = xlJustify

    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila + 1, 7)).Font.Bold = True
    xlHoja1.Cells(lnFila, 6) = "MONEDA NACIONAL"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila + 1, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila + 1, 6)).VerticalAlignment = xlJustify

    xlHoja1.Cells(lnFila, 7) = "MONEDA EXTRANJERA"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila + 1, 7)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila + 1, 7)).VerticalAlignment = xlJustify
    lnFila = lnFila + 1
    '*******************************

    ExcelCuadro xlHoja1, 2, lnFila - 2, 7, CCur(lnFila)    'ExcelCuadro 2, 8, 7, 10
    oBarra.Progress 1, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue
    lnFila = lnFila + 1 'FILA 10
    Y1 = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
    xlHoja1.Cells(lnFila, 3) = "Activos Líquidos"
    lnFila = lnFila + 1

    nCajaFondo = lnFila

    Dim nCajaFondosMN As Currency, nCajaFondosME As Currency

    Call oDbalanceCont.InsertaDetallaReporte15A(1, pdFecha, "1", 0, 0, "100", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(1, pdFecha, "2", 0, 0, "100", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(2, pdFecha, "1", 0, 0, "200", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(2, pdFecha, "2", 0, 0, "200", "A1")

    'Caja y fondos fijos
    'nSaldoDiario1 = oDbalanceCont.ObtenerCtaContSaldoDiario("1111", pdFecha) + oDbalanceCont.ObtenerCtaContSaldoDiario("111701", pdFecha)
    nCajaFondosMN = SaldoCajasObligExoneradas(pdFecha, 1) + oDAnexos.ObtieneSaldoEfectTransitoTotal(pdFecha, 1) 'NAGL 20181002 Agregó el Efectivo en Tránsito
    'nSaldoDiario2 = oDbalanceCont.ObtenerCtaContSaldoDiario("1121", pdFecha) + oDbalanceCont.ObtenerCtaContSaldoDiario("112701", pdFecha)
    nCajaFondosME = SaldoCajasObligExoneradas(pdFecha, 2) + oDAnexos.ObtieneSaldoEfectTransitoTotal(pdFecha, 2) 'NAGL 20181002 Agregó el Efectivo en Tránsito
    Call PintaFilasExcel(xlHoja1, "1101+1107.01", "Caja y Fondos Fijos", nCajaFondosMN, nCajaFondosME, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(3, pdFecha, "1", nCajaFondosMN, 1, "300", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(3, pdFecha, "2", nCajaFondosME, 1, "300", "A1")

    lsTotalActivos(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalActivos(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)

    oBarra.Progress 2, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    CargaValidacionCtaContIntDeveng15A xlHoja1.Application, pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn, psMesBalanceDiario, psAnioBalanceDiario  '***NAGL ERS 079-2017 20180123

    'Fondos disponibles en el BCRP
    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerDispobiblesenSFN("1090100822183", "1112", pdFecha, 1) + oValor.ObtenerSaldEstadistAnx15Ay15B("111802", "1", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn) 'NAGL ERS079-2017 20180128  'ObtenerCtaContSaldoBalanceDiario("111802", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
    nSaldoDiario2 = oDbalanceCont.ObtenerDispobiblesenSFN("1090100822183", "1122", pdFecha, 1) + Round(oValor.ObtenerSaldEstadistAnx15Ay15B("112802", "2", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn), 2) 'NAGL ERS079-2017 20180128  'Round(ObtenerCtaContSaldoBalanceDiario("112802", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario) / nTipoCambioBalance, 2)
    Call PintaFilasExcel(xlHoja1, "1102+1108.02", "Fondos disponibles en el BCRP", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(4, pdFecha, "1", nSaldoDiario1, 1, "425", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(4, pdFecha, "2", nSaldoDiario2, 1, "425", "A1")

    oBarra.Progress 3, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    'Fondos disponibles en empresas del sistema financiero nacional (2)
    lnFila = lnFila + 1
    nSaldoDiario1 = (oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondos(ldFecha, 1, "01,02,04", "") + oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondos(ldFecha, 1, "03", "1090100012521") - oDbalanceCont.ObtenerDispobiblesenSFN("1090100822183", "1112", pdFecha, 1)) + oValor.ObtenerSaldEstadistAnx15Ay15B("111803", "1", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn) - oValor.ObtieneRestringidosSFN(pdFecha, "1")  'NAGL ERS079-2017 20180128  'ObtenerCtaContSaldoBalanceDiario("111803", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
    nSaldoDiario2 = (oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondos(ldFecha, 2, "01,02,04", "") + oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondos(ldFecha, 2, "03", "1090100012521") - oDbalanceCont.ObtenerDispobiblesenSFN("1090100822183", "1122", pdFecha, 1)) + Round(oValor.ObtenerSaldEstadistAnx15Ay15B("112803", "2", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn), 2) - oValor.ObtieneRestringidosSFN(pdFecha, "2") 'NAGL ERS079-2017 20180128  'Round(ObtenerCtaContSaldoBalanceDiario("112803", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario) / nTipoCambioBalance, 2)
    Call PintaFilasExcel(xlHoja1, "1103+1108.03", "Fondos disponibles en empresas del sistema financiero nacional (2)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(5, pdFecha, "1", nSaldoDiario1, 1, "450", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(5, pdFecha, "2", nSaldoDiario2, 1, "450", "A1")

    'Fondos disponibles en bancos del exterior de primera categoría
    '***NAGL 20190615 Agregó ObtieneCtaSaldoDiarioAnx15A("CtaMas", "CtaMenos" ,pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)en todas las secciones
    'Según Anx02_ERS006-2019
    lnFila = lnFila + 1
    nSaldoDiario1 = ObtieneCtaSaldoDiarioAnx15A("111401,111804", "", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)
    nSaldoDiario2 = ObtieneCtaSaldoDiarioAnx15A("112401,112804", "", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)
    Call PintaFilasExcel(xlHoja1, "1104.01+1108.04(p)", "Fondos disponibles en bancos del exterior de primera categoría (3)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(6, pdFecha, "1", nSaldoDiario1, 0, "500", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(6, pdFecha, "2", nSaldoDiario2, 1, "500", "A1")

    'Fondos interbancarios netos activos (4)
    lnFila = lnFila + 1
    nSaldoDiario1 = ObtieneCtaSaldoDiarioAnx15A("121", "221", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)
    nSaldoDiario2 = ObtieneCtaSaldoDiarioAnx15A("122", "222", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)
    Call PintaFilasExcel(xlHoja1, "1200-2200", "Fondos interbancarios netos activos (4)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "1", nSaldoDiario1, 1, "600", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "2", nSaldoDiario2, 1, "600", "A1")

    lnFila = lnFila + 1
    Yvalor1 = lnFila
    'Valores representativos de deuda emitidos por el BCRP (5)
    nSaldoDiario1 = oDbalanceCont.ObtenerSumaValorRepresentativos(pdFecha, "C[BD]") '****NAGL
    nSaldoDiario2 = 0#
    Call PintaFilasExcel(xlHoja1, "1302.02.01+1304.02.01+1305.02.01", "Valores representativos de deuda emitidos por el BCRP (5)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(8, pdFecha, "1", nSaldoDiario1, 1, "725", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(8, pdFecha, "2", nSaldoDiario2, 1, "725", "A1")

    lsTotalValores(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)

    'Valores representativos de deuda emitidos por el Gobierno Central (6)
    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerSumaValorRepresentativos(pdFecha, "LT") '****NAGL
    nSaldoDiario2 = 0#
    Call PintaFilasExcel(xlHoja1, "1302.01.01.01+1304.01.01.01+1305.01.01.01", "Valores representativos de deuda emitidos por el Gobierno Central (6)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(9, pdFecha, "1", nSaldoDiario1, 1, "750", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(9, pdFecha, "2", nSaldoDiario2, 1, "750", "A1")

    lsTotalValores(1) = lsTotalValores(1) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)   '*********NAGL

    'Certificados de depósito negociables y certificados bancarios (7)
    lnFila = lnFila + 1
    nSaldoDiario1 = ObtieneCtaSaldoDiarioAnx15A("13120512,13140512,1319040512", "", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)
    nSaldoDiario2 = ObtieneCtaSaldoDiarioAnx15A("13220512,13240512,1329040512", "", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)
    Call PintaFilasExcel(xlHoja1, "1302.05.12+1302.05.19(p)+1304.05.12+1304.05.19(p) +1309.04.05.12+1309.04.05.19(p)", "Certificados de depósito negociables y certificados bancarios (7)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(10, pdFecha, "1", nSaldoDiario1, 1, "800", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(10, pdFecha, "2", nSaldoDiario2, 1, "800", "A1")

    'Valores representativos de deuda pública y de los sistemas financiero y de seguros del exterior (8)
    lnFila = lnFila + 1
    nSaldoDiario1 = 0# 'ObtieneCtaSaldoDiarioAnx15A("13140507", "", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)
    nSaldoDiario2 = 0# 'ObtieneCtaSaldoDiarioAnx15A("13240507", "", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)
    Call PintaFilasExcel(xlHoja1, "1302.01.01.02(p)+1304.01.01.02(p)+1302.05(p)+1302.06(p)+1304.05(p)+1304.06(p)+1305.01.01.02(p) +1309.04.01.01(p) +1309.04.05(p) +1309.04.06(p) +1309.05.01.01(p)", "Valores representativos de deuda pública y de los sistemas financiero y de seguros del exterior (8)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "1", nSaldoDiario1, 1, "900", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "2", nSaldoDiario2, 1, "900", "A1")

    '************Agregado by NAGL 20190613 Según Anexo02 - ERS006-2019*************
    'Bonos corporativos emitidos por empresas privadas del sector no financiero (8A)
    lnFila = lnFila + 1
    nSaldoDiario1 = 0#
    nSaldoDiario2 = 0#
    Call PintaFilasExcel(xlHoja1, "1302(p)+1304(p)+1305(p)", "Bonos corporativos emitidos por empresas privadas del sector no financiero (8A)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "1", nSaldoDiario1, 1, "910", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "2", nSaldoDiario2, 1, "910", "A1")

    'Valores representativos de deuda del BCRP y Gobierno Central recibidos en operaciones de reporte (8B)
    lnFila = lnFila + 1
    nSaldoDiario1 = 0#
    nSaldoDiario2 = 0#
    Call PintaFilasExcel(xlHoja1, "1507.11(p)", "Valores representativos de deuda del BCRP y Gobierno Central recibidos en operaciones de reporte (8B)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "1", nSaldoDiario1, 1, "920", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "2", nSaldoDiario2, 1, "920", "A1")

    'Valores representativos de deuda de Gobiernos del Exterior recibidos en operaciones de reporte (8B)
    lnFila = lnFila + 1
    nSaldoDiario1 = 0#
    nSaldoDiario2 = 0#
    Call PintaFilasExcel(xlHoja1, "1507.11(p)", "Valores representativos de deuda de Gobiernos del Exterior recibidos en operaciones de reporte (8B)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "1", nSaldoDiario1, 1, "930", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "2", nSaldoDiario2, 1, "930", "A1")

    'Bonos corporativos emitidos por empresas privadas del sector no financiero recibidos en operaciones de reporte (8B)
    lnFila = lnFila + 1
    nSaldoDiario1 = 0#
    nSaldoDiario2 = 0#
    Call PintaFilasExcel(xlHoja1, "1507.11(p)", "Bonos corporativos emitidos por empresas privadas del sector no financiero recibidos en operaciones de reporte (8B)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "1", nSaldoDiario1, 1, "940", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "2", nSaldoDiario2, 1, "940", "A1")
    '*************************END NAGL 20190613************************************

    lsTotalActivos(1) = lsTotalActivos(1) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalActivos(2) = lsTotalActivos(2) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)

    oBarra.Progress 4, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    '*************** TOTALES ACTIVOS DE LIQUIDEZ *************************
    lnFila = lnFila + 3 'NAGL 20190614 Cambio de 1 a 3
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "", "Total(a)", "Sum(" & lsTotalActivos(1) & ")", "Sum(" & lsTotalActivos(2) & ")", lnFila, True, True, True)

    lnTotalaMN = Format(xlHoja1.Cells(lnFila, 6), "#,##0.00;-#,##0.00") 'VAPA 20170911
    lnTotalaME = Format(Round(xlHoja1.Cells(lnFila, 7), 2), "#,##0.00;-#,##0.00") 'VAPA 20170911

    Call oDbalanceCont.InsertaDetallaReporte15A(12, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 1, "1000", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(12, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 1, "1000", "A1")

    ReDim lsTotalesActivos(2)
    ReDim lsTotalesPasivos(2)

    '****************PASIVOS DE CORTO PLAZO*****************
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).HorizontalAlignment = xlCenter
    Call PintaFilasExcel(xlHoja1, "", "Pasivos de Corto Plazo", "", "", lnFila, False, True, False)
    Call oDbalanceCont.InsertaDetallaReporte15A(13, pdFecha, "1", 0, 0, "1100", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(13, pdFecha, "2", 0, 0, "1100", "A1")

    lnFila = lnFila + 1
    lsTotalPasivos(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalPasivos(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)

    'Obligaciones a la vista (9)
    xlHoja1.Cells(lnFila, 2) = "2101 -2101.18 +2301(p) +2108.01 +2308.01(p)"
    xlHoja1.Cells(lnFila, 3) = "Obligaciones a la vista (9)"
    nSaldoDiario1 = ObtieneCtaSaldoDiarioAnx15A("2111,2311,211801,231801", "211118", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn) 'NAGL 20191015 Agregó 211118, para excluirlo de Obligaciones a la Vista 'ObtenerCtaContSaldoBalanceDiario("2111", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)'Comentado by NAGL 20190617
    nSaldoDiario2 = ObtieneCtaSaldoDiarioAnx15A("2121,2321,212801,232801", "212118", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn) 'NAGL 20191015 Agregó 212118, para excluirlo de Obligaciones a la Vista 'Round(ObtenerCtaContSaldoBalanceDiario("2121", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario) / nTipoCambioBalance, 2)'Comentado by NAGL 20190617
    Call PintaFilasExcel(xlHoja1, "2101+2301(p)+2108.01+2308.01(p)", "Obligaciones a la vista (9)", nSaldoDiario1, nSaldoDiario2, lnFila, False, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(14, pdFecha, "1", nSaldoDiario1, 1, "1210", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(14, pdFecha, "2", nSaldoDiario2, 1, "1210", "A1")

    'Obligaciones con instituciones recaudadoras de tributos (10)

    'nSubValor1 = ObtenerCtaContSaldoBalanceDiario("25170301", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
    'nSaldoDiario1 = IIf(nSubValor1 < 0, 0, nSubValor1)
    'nSubValor1 = ObtenerCtaContSaldoBalanceDiario("25170302", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
    'nSaldoDiario1 = nSaldoDiario1 + IIf(nSubValor1 < 0, 0, nSubValor1)
    'nSubValor1 = ObtenerCtaContSaldoBalanceDiario("25170303", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
    'nSaldoDiario1 = nSaldoDiario1 + nSubValor1
    'nSubValor1 = ObtenerCtaContSaldoBalanceDiario("251704", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
    'nSaldoDiario1 = nSaldoDiario1 + nSubValor1
    'nSubValor1 = ObtenerCtaContSaldoBalanceDiario("251705", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
    'nSaldoDiario1 = nSaldoDiario1 + nSubValor1

    'nSubValor2 = ObtenerCtaContSaldoBalanceDiario("25270301", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario)
    'nSaldoDiario2 = IIf(nSubValor2 < 0, 0, nSubValor2)
    'nSubValor2 = ObtenerCtaContSaldoBalanceDiario("25270302", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario)
    'nSaldoDiario2 = nSaldoDiario2 + IIf(nSubValor2 < 0, 0, nSubValor2)
    'nSubValor2 = ObtenerCtaContSaldoBalanceDiario("25270303", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario)
    'nSaldoDiario2 = nSaldoDiario2 + nSubValor2
    'nSubValor2 = ObtenerCtaContSaldoBalanceDiario("252704", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario)
    'nSaldoDiario2 = nSaldoDiario2 + nSubValor2
    'nSubValor2 = ObtenerCtaContSaldoBalanceDiario("252705", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario)
    'nSaldoDiario2 = Round((nSaldoDiario2 + nSubValor2 / nTipoCambioBalance), 2)
    'Comentado by NAGL 20190617
    lnFila = lnFila + 1
    nSaldoDiario1 = ObtieneCtaSaldoDiarioAnx15A("25170301,25170302,25170303,25170309,251704,251705,251706,251801", "", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)
    nSaldoDiario2 = ObtieneCtaSaldoDiarioAnx15A("25270301,25270302,25270303,25270309,252704,252705,252706,252801", "", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)
    Call PintaFilasExcel(xlHoja1, "2507.03+2507.04+2507.05+2507.06+2508(p)", "Obligaciones con instituciones recaudadoras de tributos (10)", nSaldoDiario1, nSaldoDiario2, lnFila, False, False, True) 'NAGL Cambio el último de False a True
    Call oDbalanceCont.InsertaDetallaReporte15A(15, pdFecha, "1", nSaldoDiario1, 1, "1225", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(15, pdFecha, "2", nSaldoDiario2, 1, "1225", "A1")

    'Cuentas por pagar por operaciones de reporte (34)
    lnFila = lnFila + 1
    Set rsvalor = oValor.AdeudadoReactiva(pdFecha, 1) 'JIPR20200824
    If Not (rsvalor.EOF And rsvalor.BOF) Then
        nSaldoDiario1 = rsvalor!nSaldo
        nSaldoDiario2 = 0
    Else
        nSaldoDiario1 = 0
        nSaldoDiario2 = 0
    End If

    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2504.11(p)", "Cuentas por pagar por operaciones de reporte (34)", nSaldoDiario1, nSaldoDiario2, lnFila, False, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(16, pdFecha, "1", nSaldoDiario1, 1, "1230", "A1") 'JIPR20200824  nSaldoDiario1
    Call oDbalanceCont.InsertaDetallaReporte15A(16, pdFecha, "2", nSaldoDiario2, 1, "1230", "A1") 'JIPR20200824  nSaldoDiario2
    'NAGL 20190613

    'Cuentas por pagar por ventas en corto (11)
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2504.12", "Cuentas por pagar por ventas en corto(11)", "", "", lnFila, False, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(16, pdFecha, "1", 0, 1, "1250", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(16, pdFecha, "2", 0, 1, "1250", "A1")

    'Fondos interbancarios netos pasivos (4)
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2200-1200", "Fondos interbancarios netos pasivos (4)", "", "", lnFila, False, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(17, pdFecha, "1", 0, 1, "1300", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(17, pdFecha, "2", 0, 1, "1300", "A1")

    'Obligaciones por cuentas de ahorro
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True

    '****Agregado by NAGL 20190617 Según ANX02_ERS006-2019*******
    nSaldoDiario1 = oDbalanceCont.SaldoAhoPlaFijCTS(Format(pdFecha, "yyyymmdd"), 1, "232") '2112 + 2312
    nSaldoDiario1 = nSaldoDiario1 - oValor.ObtenerSaldEstadistAnx15Ay15B("211701", "1", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)
    nSaldoDiario1 = nSaldoDiario1 + ObtieneCtaSaldoDiarioAnx15A("211802,231802", "", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)

    nSaldoDiario2 = oDbalanceCont.SaldoAhoPlaFijCTS(Format(pdFecha, "yyyymmdd"), 2, "232") '2122 + 2322
    nSaldoDiario2 = nSaldoDiario2 - oValor.ObtenerSaldEstadistAnx15Ay15B("212701", "2", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + ObtieneCtaSaldoDiarioAnx15A("212802,232802", "", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)
    '*************END NAGL 20190617*******************************
    Call PintaFilasExcel(xlHoja1, "2102+2302(p)+2108.02+2308.02(p)", "Obligaciones por cuentas de ahorro", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "1", nSaldoDiario1, 1, "1400", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "2", nSaldoDiario2, 1, "1400", "A1")

    'Obligaciones por cuentas a plazo (12)
    Set rsvalor = Nothing
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    'nSaldoDiario1 = oDbalanceCont.ObtenerPlazoFijoxRango(pdFecha, 1, 1, 360) - oDbalanceCont.ObtenerSaldoDiarioRestringido(pdFecha, "1", 360) + oDbalanceCont.ObtenerPlazoFijoxRango(pdFecha, 2, 1, 360) + oValor.ObtenerSaldEstadistAnx15Ay15B("211803", "1", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn) + oValor.ObtenerSaldEstadistAnx15Ay15B("231803", "1", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn) 'NAGL ERS079-2017 20180128
    'nSaldoDiario2 = oDbalanceCont.ObtenerPlazoFijoxRango(pdFecha, 1, 2, 360) - oDbalanceCont.ObtenerSaldoDiarioRestringido(pdFecha, "2", 360) + oDbalanceCont.ObtenerPlazoFijoxRango(pdFecha, 2, 2, 360) + Round(oValor.ObtenerSaldEstadistAnx15Ay15B("212803", "2", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn), 2) + Round(oValor.ObtenerSaldEstadistAnx15Ay15B("232803", "2", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn), 2) 'NAGL ERS079-2017 20180128
    'Comentado by NAGL 20190618
    '****Agregado by NAGL 20190617 Según ANX02_ERS006-2019*******
    Set rsvalor = oValor.ObtieneOtrasObligaciones15A(pdFecha, "1", "233", "Gen")
    nSaldoDiario1 = rsvalor!nSaldo
    Set rsvalor = Nothing
    Set rsvalor = oValor.ObtieneOtrasObligaciones15A(pdFecha, "2", "233", "Gen")
    nSaldoDiario2 = rsvalor!nSaldo
    '********************END NAGL 20190617************************
    Call PintaFilasExcel(xlHoja1, "2103(p)-2103.05(p)+2303(p)+2108.03(p)+2308.03(p)", "Obligaciones por cuentas a plazo (12)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(19, pdFecha, "1", nSaldoDiario1, 1, "1450", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(19, pdFecha, "2", nSaldoDiario2, 1, "1450", "A1")

    'Adeudos y obligaciones financieras del país (13)
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerSaldoAdeudadoMyP(pdFecha, "1", 1, 1) + oValor.ObtenerSaldEstadistAnx15Ay15B("241802", "1", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn) + oValor.ObtenerSaldEstadistAnx15Ay15B("241806", "1", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn) 'NAGL ERS079-2017 20180128
    nSaldoDiario1 = nSaldoDiario1 + oValor.ObtenerSaldEstadistAnx15Ay15B("241803", "1", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn) 'Agregado by NAGL 20190619 Según Anx02_ERS006-2019

    nSaldoDiario2 = oDbalanceCont.ObtenerSaldoAdeudadoMyP(pdFecha, "2", 1, 1) + Round(oValor.ObtenerSaldEstadistAnx15Ay15B("242802", "2", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn), 2) + Round(oValor.ObtenerSaldEstadistAnx15Ay15B("242806", "2", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn), 2) 'NAGL ERS079-2017 20180128
    nSaldoDiario2 = nSaldoDiario2 + oValor.ObtenerSaldEstadistAnx15Ay15B("242803", "2", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn) 'Agregado by NAGL 20190619 Según Anx02_ERS006-2019

    Call PintaFilasExcel(xlHoja1, "2401+2402+2403+2406 +2408.01+2408.02+2408.03 +2408.06+2409.01+2602(p)+2603(p)+2606(p)+2608.02(p)+2608.03(p)+2608.06(p)+2609.01", "Adeudos y obligaciones financieras del país (13)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(20, pdFecha, "1", nSaldoDiario1, 1, "1500", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(20, pdFecha, "2", nSaldoDiario2, 1, "1500", "A1")

    'Adeudos y obligaciones financieras del exterior (13)
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerSaldoAdeudadoMyP(pdFecha, "1", 1, 0) + oValor.ObtenerSaldEstadistAnx15Ay15B("241804", "1", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn) + oValor.ObtenerSaldEstadistAnx15Ay15B("241805", "1", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn) + oValor.ObtenerSaldEstadistAnx15Ay15B("241807", "1", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn) 'NAGL ERS079-2017 20180128
    nSaldoDiario2 = oDbalanceCont.ObtenerSaldoAdeudadoMyP(pdFecha, "2", 1, 0) + Round(oValor.ObtenerSaldEstadistAnx15Ay15B("242804", "2", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn), 2) + Round(oValor.ObtenerSaldEstadistAnx15Ay15B("242805", "2", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn), 2) + Round(oValor.ObtenerSaldEstadistAnx15Ay15B("242807", "2", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn), 2) 'NAGL ERS079-2017 20180128
    Call PintaFilasExcel(xlHoja1, "2404+2405+2407+2408.04+2408.05+2408.07+2409.02+2409.03+2604(p)+2605(p)+2607(p)+2608.04(p)+2608.05(p)+2608.07(p)+2609.02+2609.03", "Adeudos y obligaciones financieras del exterior (13)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "1", nSaldoDiario1, 1, "1510", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "2", nSaldoDiario2, 1, "1510", "A1")

    'Valores, títulos y obligaciones en circulación (14)
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2800(p)", "Valores, títulos y obligaciones en circulación (14)", "0", "0", lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(22, pdFecha, "1", 0, 1, "1520", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(22, pdFecha, "2", 0, 1, "1520", "A1")

    oBarra.Progress 5, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    lsTotalPasivos(1) = lsTotalPasivos(1) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalPasivos(2) = lsTotalPasivos(2) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)

    'Trasladado en la Parte Superior by NAGL 20190515
    'CargaValidacionCtaContIntDeveng15A xlHoja1.Application, pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn, psMesBalanceDiario, psAnioBalanceDiario  '***NAGL ERS 079-2017 20180123

    '******************** TOTALES DE PASIVOS DE CORTO PLAZO ***************************
    'Me.prgBarra.value = 50
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True 'NAGL 20190614

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).HorizontalAlignment = xlCenter
    Call PintaFilasExcel(xlHoja1, "", "Total(b)", "Sum(" & lsTotalPasivos(1) & ")", "Sum(" & lsTotalPasivos(2) & ")", lnFila, True, True, True)

    Call oDbalanceCont.InsertaDetallaReporte15A(23, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 0, "1600", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(23, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 0, "1600", "A1")

    lsTotalesPasivos(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalesPasivos(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)

    Y2 = lnFila
    xlHoja1.Range(xlHoja1.Cells(Y1, 2), xlHoja1.Cells(Y2, 7)).Borders.LineStyle = xlContinuous

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Merge True 'NAGL 20190614

    'Ratios de Liquidez[(a)/(b)]*100
    lnFila = lnFila + 1 'NAGL Cambio de 2 a 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "Ratios de Liquidez[(a)/(b)]*100", "", "(sum(" & lsTotalActivos(1) & ")/sum(" & lsTotalPasivos(1) & "))*100", "(sum(" & lsTotalActivos(2) & ")/sum(" & lsTotalPasivos(2) & "))*100", lnFila, True, True, True)

    lnRatioLiquidezMN = Format(xlHoja1.Cells(lnFila, 6), "#,##0.00;-#,##0.00") 'VAPA20170911
    lnRatioLiquidezMN = Format(Round(xlHoja1.Cells(lnFila, 6), 2), "#,##0.00;-#,##0.00") 'VAPA20170911
    lnRatioLiquidezME = Format(Round(xlHoja1.Cells(lnFila, 7), 2), "#,##0.00;-#,##0.00") 'VAPA20170911

    Call oDbalanceCont.InsertaDetallaReporte15A(24, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 1, "1700", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(24, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 1, "1700", "A1")
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    
    Set rsvalor = oValor.ObligacionBN(pdFecha) '***NAGL
    If Not (rsvalor.EOF And rsvalor.BOF) Then
        nObligMNDiario1 = rsvalor!mObligacionMN
        nObligMEDiario2 = rsvalor!mObligacionME
    Else
        nObligMNDiario1 = 0
        nObligMEDiario2 = 0
    End If

    lnFila = lnFila + 1
    'Activos líquidos ajustados por recursos prestados (c)(15)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "Activos líquidos ajustados por recursos prestados (c)(15)", "", "Sum(" & lsTotalActivos(1) & ") - " & nObligMNDiario1, "Sum(" & lsTotalActivos(2) & ")-" & nObligMEDiario2, lnFila, True, False, True) '********NAGL
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(25, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 1, "1710", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(25, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 1, "1710", "A1")

    'Pasivos de corto plazo ajustados por recursos prestados (d)(15)
    lnFila = lnFila + 1
    nSaldoDiario1 = ObtenerCtaContSaldoBalanceDiario("4513012903", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario) + ObtenerCtaContSaldoBalanceDiario("4513011002", pdFecha, "1", psMesBalanceDiario, psAnioBalanceDiario)
    nSaldoDiario2 = Round(ObtenerCtaContSaldoBalanceDiario("4523012903", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario) + ObtenerCtaContSaldoBalanceDiario("4523011002", pdFecha, "2", psMesBalanceDiario, psAnioBalanceDiario) / nTipoCambioBalance)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "Pasivos de corto plazo ajustados por recursos prestados (d)(15)", "", "Sum(" & lsTotalPasivos(1) & ") - " & nObligMNDiario1, "Sum(" & lsTotalPasivos(2) & ")-" & nObligMEDiario2, lnFila, True, False, True) '*********NAGL
    Call oDbalanceCont.InsertaDetallaReporte15A(26, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 1, "1720", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(26, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 1, "1720", "A1")
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    'Ratio de liquidez ajustado por recursos prestados [(c)/(d)]x100
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    nSaldoDiario1 = xlHoja1.Cells(lnFila - 2, 6) / xlHoja1.Cells(lnFila - 1, 6)
    nSaldoDiario2 = xlHoja1.Cells(lnFila - 2, 7) / xlHoja1.Cells(lnFila - 1, 7)

    lnRatioLAjusRecursosPrestadosMN = nSaldoDiario1 * 100 'VAPA 20170911
    lnRatioLAjusRecursosPrestadosME = nSaldoDiario2 * 100 'VAPA 20170911

    '***NAGL 20181121
    Call PintaFilasExcel(xlHoja1, "Ratio de liquidez ajustado por recursos prestados [(c)/(d)]x100", "", "(" & xlHoja1.Range(xlHoja1.Cells(lnFila - 2, 6), xlHoja1.Cells(lnFila - 2, 6)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 6), xlHoja1.Cells(lnFila - 1, 6)).Address(False, False) & ")" & "*" & "100", "(" & xlHoja1.Range(xlHoja1.Cells(lnFila - 2, 7), xlHoja1.Cells(lnFila - 2, 7)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 7), xlHoja1.Cells(lnFila - 1, 7)).Address(False, False) & ")" & "*" & "100", lnFila, True, True, True)
    '***END NAGL
    Call oDbalanceCont.InsertaDetallaReporte15A(27, pdFecha, "1", nSaldoDiario1 * 100, 1, "1730", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(27, pdFecha, "2", nSaldoDiario2 * 100, 1, "1730", "A1")
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    'Posiciones largas en forwards de monedas (e) (15)
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "Posiciones largas en forwards de monedas (e) (15)", "", "", "", lnFila, False, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(28, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 1, "1740", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(28, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 1, "1740", "A1")

    'Posiciones cortas en forwards de monedas (f) (15)
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "Posiciones cortas en forwards de monedas (f) (15)", "", "", "", lnFila, False, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(29, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 1, "1745", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(29, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 1, "1745", "A1")
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    'Ratio de liquidez ajustado por forwards de monedas [((a)+(e))/((b)+(f))]x100
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "Ratio de liquidez ajustado por forwards de monedas [((a)+(e))/((b)+(f))]x100", "", "(sum(" & lsTotalActivos(1) & ")/sum(" & lsTotalPasivos(1) & "))*100", "(sum(" & lsTotalActivos(2) & ")/sum(" & lsTotalPasivos(2) & "))*100", lnFila, True, True, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(30, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 1, "1750", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(30, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 1, "1750", "A1")
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    'Depósitos overnight BCRP (g)
    lnFila = lnFila + 1

    lsTotalRatioLiquidez(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalRatioLiquidez(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)

    nSaldoDiario1 = oDbalanceCont.ObtenerSaldoOverNight("1", pdFecha, "2")
    nSaldoDiario2 = oDbalanceCont.ObtenerSaldoOverNight("2", pdFecha, "2")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "Depósitos overnight BCRP (g)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(31, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 1, "1755", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(31, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 1, "1755", "A1")
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    'Depósitos plazo BCRP (h)
    lnFila = lnFila + 1
    'ALPA 20151106***************************
    ' nSaldoDiario1 = oDbalanceCont.ObtenerCtaSaldoDiario("131502", pdFecha) + oDbalanceCont.ObtenerCtaSaldoDiario("11120501", pdFecha) comentado ANPS15042021
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaSaldoDiario("11120501", pdFecha) 'modificado ANPS15042021
    nSaldoDiario2 = oDbalanceCont.ObtenerCtaSaldoDiario("132502", pdFecha) + oDbalanceCont.ObtenerCtaSaldoDiario("11220501", pdFecha) 'NAGL
    '****************************************
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "Depósitos plazo BCRP (h)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(32, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 0, "1760", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(32, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 0, "1760", "A1")
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    'Valores representativos de deuda emitidos por el BCRP y Gobierno Central (i)
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True '**********NAGL  lsTotalValores(1)
    Call PintaFilasExcel(xlHoja1, "Valores representativos de deuda emitidos por el BCRP y Gobierno Central (i)", "", "Sum(" & lsTotalValores(1) & ")", "0.00", lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(33, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 1, "1770", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(33, pdFecha, "2", 0, 1, "1770", "A1")
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lsTotalRatioLiquidez(1) = lsTotalRatioLiquidez(1) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalRatioLiquidez(2) = lsTotalRatioLiquidez(2) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)

    oBarra.Progress 6, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    'Valores representativos de deuda emitidos por Gobiernos del Exterior (j)
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "Valores representativos de deuda emitidos por Gobiernos del Exterior (j)", "", "", "", lnFila, False, False, False)
    Call oDbalanceCont.InsertaDetallaReporte15A(34, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 0, "1780", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(34, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 1, "1780", "A1")
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    'Ratio de inversiones liquidas  [((g)+(h)+(i)+(j))/(a)]x100
    'JIPR20200824
    ReDim lsTotalRatioLiquidezMN(1) As String
    ReDim lsTotalRatioLiquidezME(1) As String
    'lsTotalRatioLiquidezMN(1) = xlHoja1.Range(xlHoja1.Cells(48, 6), xlHoja1.Cells(48, 6)).Address(False, False) JIPR20201128 AJUSTE
    lsTotalRatioLiquidezMN(1) = xlHoja1.Range(xlHoja1.Cells(48, 6), xlHoja1.Cells(50, 6)).Address(False, False)
    lsTotalRatioLiquidezME(1) = xlHoja1.Range(xlHoja1.Cells(48, 7), xlHoja1.Cells(48, 7)).Address(False, False)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    'JIPR20200824
    Call PintaFilasExcel(xlHoja1, "Ratio de inversiones liquidas  [((g)+(h)+(i)+(j))/(a)]x100", "", "(sum(" & lsTotalRatioLiquidezMN(1) & ")/sum(" & lsTotalActivos(1) & "))*100", "(sum(" & lsTotalRatioLiquidezME(1) & ")/sum(" & lsTotalActivos(2) & "))*100", lnFila, True, True, True)
    'Call PintaFilasExcel(xlHoja1, "Ratio de inversiones liquidas  [((g)+(h)+(i)+(j))/(a)]x100", "", "(sum(" & lsTotalRatioLiquidez(1) & ")/sum(" & lsTotalActivos(1) & "))*100", "(sum(" & lsTotalRatioLiquidez(2) & ")/sum(" & lsTotalActivos(2) & "))*100", lnFila, True, True, True)
    'Call PintaFilasExcel(xlHoja1, "Ratio de inversiones liquidas  [((g)+(h)+(i)+(j))/(a)]x100", "", "((sum(" & lsTotalRatioLiquidez(1) & ")+" & xlHoja1.Range(xlHoja1.Cells(Yvalor1, 6), xlHoja1.Cells(Yvalor1, 6)).Address(False, False) & ")" & "/sum(" & lsTotalActivos(1) & "))*100", "(" & xlHoja1.Range(xlHoja1.Cells(lnFila - 4, 7), xlHoja1.Cells(lnFila - 4, 7)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(Yvalor1, 7), xlHoja1.Cells(Yvalor1, 7)).Address(False, False) & ")/sum(" & lsTotalActivos(2) & ")*100", lnFila, True, True, True) 'NAGL 20190622

    lnRatioInversionesLiquidasMN = Format(Round(xlHoja1.Cells(lnFila, 6), 2), "#,##0.00;-#,##0.00") 'VAPA20170911
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 0, "1790", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 0, "1790", "A1")
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    '*******************II OTRAS OPERACIONES**************************

    lnFila = lnFila + 3  'FILA 55
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Merge True
    Call PintaFilasExcel(xlHoja1, "II.  OTRAS OPERACIONES", "", "", "", lnFila, False, True, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Borders.LineStyle = xlContinuous
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)

    lnFila = lnFila + 1 '*********NAGL

    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 8)).Merge True

    'Uso de PintaFilasExcel2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Cells(lnFila, 5) = "Monto(16A)"
    Call PintaFilasExcel(xlHoja1, "", "Tasas de interés (16)", "", "Saldos(16B)", lnFila, False, True, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)

    'xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Borders.LineStyle = xlContinuous
    '****Agregado by NAGL 20190614
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila + 1, 2)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila + 1, 2)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Cells(lnFila, 3) = "Moneda ": xlHoja1.Cells(lnFila, 4) = "Moneda ": xlHoja1.Cells(lnFila, 5) = "Moneda ": xlHoja1.Cells(lnFila, 6) = "Moneda ": xlHoja1.Cells(lnFila, 7) = "Moneda ": xlHoja1.Cells(lnFila, 8) = "Moneda "
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "Nacional": xlHoja1.Cells(lnFila, 4) = "Extranjera": xlHoja1.Cells(lnFila, 5) = "Nacional": xlHoja1.Cells(lnFila, 6) = "Extranjera": xlHoja1.Cells(lnFila, 7) = "Nacional": xlHoja1.Cells(lnFila, 8) = "Extranjera"
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila - 1), 8, CCur(lnFila)
    '*****************END 20190614*******************'

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel2(xlHoja1, "1. Operaciones overnight(17)", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1 'CELDA PENDIENTE
    Yvalor2 = lnFila

    lnFila = lnFila + 1
    nSaldoDiario1 = 0 'oDbalanceCont.ObtenerSaldoOverNight("1", pdFecha, "1")
    nSaldoDiario2 = 0 'oDbalanceCont.ObtenerSaldoOverNight("2", pdFecha, "1")
    Call PintaFilasExcel2(xlHoja1, "       1.1.1 Empresas del sistema financiero", "", "", nSaldoDiario1, nSaldoDiario2, "", "", lnFila, False, False, True)
    lnFila = lnFila + 1

    'Para el Cálculo en Sección 1.1.2 Otras
    nSumarMN = 0
    nSaldoPFN = 0
    nSaldoPFME = 0
    nsDiv = 0
    nsDiv2 = 0
    nsDivME = 0
    
    Set rsvalor = oValor.MuestraBCRP(pdFecha, "C[BD]")
    nContar = 0
    If Not (rsvalor.EOF And rsvalor.BOF) Then
        For X = 0 To rsvalor.RecordCount - 1
            lnFila = lnFila + 1
            Call PintaFilasExcel2(xlHoja1, "                      - BCRP", Format(rsvalor!nTasaInteres, gsFormatoNumeroView), "", Format(rsvalor!nValorRazonable, gsFormatoNumeroView), "", "", "", lnFila, False, False, True)
            nSumarMN = rsvalor!nValorRazonable + nSumarMN 'total valor razonable CDBCRP
            nContar = nContar + 1
            rsvalor.MoveNext
        Next
    Else
        nSumarMN = 0
        nContar = 0
    End If
      
    Set rsDetalleCtas = oCtaIf.obtenerBCRInversionDPF(pdFecha)
    nContar2 = 0
    If Not (rsDetalleCtas.EOF And rsDetalleCtas.BOF) Then
        Do While Not rsDetalleCtas.EOF
            lnFila = lnFila + 1
            Call PintaFilasExcel2(xlHoja1, "                      - BCRP", Format(rsDetalleCtas!TEAMN, gsFormatoNumeroView), Format(rsDetalleCtas!TEAME, gsFormatoNumeroView), Format(rsDetalleCtas!nSaldoMN, gsFormatoNumeroView), Format(rsDetalleCtas!nSaldoME, gsFormatoNumeroView), "", "", lnFila, False, False, True)
            nSaldoPFN = rsDetalleCtas!nSaldoMN + nSaldoPFN  'total overnight BCRP MN
            nSaldoPFME = rsDetalleCtas!nSaldoME + nSaldoPFME 'total overnight BCRP ME
            nContar2 = nContar2 + 1
            rsDetalleCtas.MoveNext
        Loop
    Else
        nContar2 = 0
        nSaldoPFN = 0
        nSaldoPFME = 0
    End If
    '***NAGL***

    nSumarTotalMN = nSumarMN + nSaldoPFN 'total entre valor razon. y overnight MN
    
    Set rsvalor = oValor.MuestraBCRP(pdFecha, "C[BD]")
    If Not (rsvalor.EOF And rsvalor.BOF) Then
        For X = 0 To rsvalor.RecordCount - 1
            nsDiv = (rsvalor!nValorRazonable / nSumarTotalMN) * rsvalor!nTasaInteres + nsDiv
            rsvalor.MoveNext
        Next
    Else
        nsDiv = 0
    End If
    
    Set rsDetalleCtas = oCtaIf.obtenerBCRInversionDPF(pdFecha)
    If Not (rsDetalleCtas.EOF And rsDetalleCtas.BOF) Then
        Do While Not rsDetalleCtas.EOF
            If (nSaldoPFN = 0) Then
                nsDiv2 = 0
            Else
                nsDiv2 = (rsDetalleCtas!nSaldoMN / nSumarTotalMN) * rsDetalleCtas!TEAMN + nsDiv2
            End If
            If (nSaldoPFME = 0) Then
                nsDivME = 0
            Else
                nsDivME = (rsDetalleCtas!nSaldoME / nSaldoPFME) * rsDetalleCtas!TEAME + nsDivME
            End If
            rsDetalleCtas.MoveNext
        Loop
    Else
        nsDiv2 = 0
        nsDivME = 0
    End If

    nPromedio = nsDiv + nsDiv2

    lnFila = lnFila - nContar - nContar2 - 2
    Call PintaFilasExcel2(xlHoja1, "   1.1 Activas", Format(nPromedio, gsFormatoNumeroView), Format(nsDivME, gsFormatoNumeroView), Format(nSumarTotalMN, gsFormatoNumeroView), Format(nSaldoPFME, gsFormatoNumeroView), "", "", lnFila, False, False, True)

    lnFila = lnFila + 2
    Call PintaFilasExcel2(xlHoja1, "       1.1.2 Otras", Format(nPromedio, gsFormatoNumeroView), Format(nsDivME, gsFormatoNumeroView), Format(nSumarTotalMN, gsFormatoNumeroView), Format(nSaldoPFME, gsFormatoNumeroView), "", "", lnFila, False, False, True)

    lnFila = lnFila + nContar + nContar2 + 1  '******* NAGL ANTES :lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "   1.2 Pasivas", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       1.2.1 Empresas del sistema financiero", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       1.2.2 Otras", "", "", "", "", "", "", lnFila, False, False, False)
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)
    '****NAGL 20190614
    xlHoja1.Range(xlHoja1.Cells(Yvalor2, 7), xlHoja1.Cells(lnFila, 8)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(Yvalor2 - 1, 2), xlHoja1.Cells(lnFila, 2)).Interior.ColorIndex = 2
    xlHoja1.Range(xlHoja1.Cells(Yvalor2, 3), xlHoja1.Cells(lnFila, 6)).Interior.ColorIndex = 2
    '******************

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Call PintaFilasExcel2(xlHoja1, "2. Fondos interbancarios", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).Interior.Color = RGB(153, 153, 255)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "      2.1 Activos (Cuenta 1201)", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "      2.2 Pasivos (Cuenta 2201)", "", "", "", "", "", "", lnFila, False, False, False)
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 7), xlHoja1.Cells(lnFila, 8)).Interior.Color = RGB(153, 153, 255)

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Call PintaFilasExcel2(xlHoja1, "3. Obligaciones con el Banco de la Nación (18)", "", "", "", "", Format(nObligMNDiario1, gsFormatoNumeroView), Format(nObligMEDiario2, gsFormatoNumeroView), lnFila, False, False, False)
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Call PintaFilasExcel2(xlHoja1, "4. Operaciones de venta con compromiso de recompra y operaciones de compra y venta simultánea de valores (19)", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 15 '38.25
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "   4.1 Adquiriente", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       4.1.1 Con instrumentos de inversión del BCRP y del Tesoro Público", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 25.5
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify

    '*****************Agregado by NAGL 20190614********************'
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.1.1.1 Empresas del sistema financiero", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.1.1.2 Otras", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       4.1.2 Con otros ALAC", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.1.2.1 Empresas del sistema financiero", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.1.2.2 Otras", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       4.1.3 Con otros Instrumentos", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.1.3.1 Empresas del sistema financiero", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.1.3.2 Otras", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "   4.2 Enajenante", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       4.2.1 Con instrumentos de inversión del BCRP y del Tesoro Público", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 15 '25.5
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.2.1.1 Empresas del sistema financiero", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.2.1.2 Otras", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       4.2.2 Con otros ALAC", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.2.2.1 Empresas del sistema financiero", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.2.2.2 Otras", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       4.2.3 Con otros Instrumentos", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.2.3.1 Empresas del sistema financiero", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.2.3.2 Otras", "", "", "", "", "", "", lnFila, False, False, False)

    xlHoja1.Range(xlHoja1.Cells(lnPosInicial - 3, 2), xlHoja1.Cells(lnPosInicial - 3, 2)).Interior.ColorIndex = 2
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial - 2, 2), xlHoja1.Cells(lnFila, 6)).Interior.ColorIndex = 2
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2

    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)
    '**************************END NAGL**************************'

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Call PintaFilasExcel2(xlHoja1, "5. Transferencia temporal de valores (20)", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "      5.1 Con activos líquidos", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "      5.2 Con activos no líquidos", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Call PintaFilasExcel2(xlHoja1, "6. Créditos del BCRP con fines de regulación monetaria", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 25.5
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "7. Operaciones de reporte de monedas con el BCRP (21)", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 25.5
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Call PintaFilasExcel2(xlHoja1, "      7.1 Repo Regular", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "      7.2 Repo Expansión", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "      7.3 Repo Sustitución", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Call PintaFilasExcel2(xlHoja1, "8.  Swaps cambiarios con el BCRP (22)", "", "", "", "", "", "", lnFila, False, False, False)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.SaldosRepoGartAnexo("2", "1", "000")  'JIPR20210318 REPOGARTR
    Call PintaFilasExcel2(xlHoja1, "9.  Operaciones de reporte de cartera de créditos con el BCRP (22A)", 0, 0, 0, 0, Format(nSaldoDiario1, gsFormatoNumeroView), 0, lnFila, False, False, False)  'nSaldoDiario1 JIPR20210318
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)

  'ANPS20210818 REPOPROG
     CargaReprogramados xlHoja1.Application, lnFila, lnPosInicial
      
 lnFila = lnFila + 11  'ANPS20210818
    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Set rsvalor = oDbalanceCont.ObtenerMontosyTasasSubastaTesoroPublico(pdFecha)
    If Not rsvalor.BOF And Not rsvalor.EOF Then
        Do While Not rsvalor.EOF
            nTasaSubasta1 = nTasaSubasta1 + rsvalor!TasaParcialMN
            nTasaSubasta2 = nTasaSubasta2 + rsvalor!TasaParcialME
            nMontoMN = nMontoMN + rsvalor!MontoMN
            nMontoME = nMontoME + rsvalor!MontoME
            rsvalor.MoveNext
        Loop
    Else
        nTasaSubasta1 = 0
        nTasaSubasta2 = 0
        nMontoMN = 0
        nMontoME = 0
    End If

    nSaldoDiario1 = 0 'oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado("1", pdFecha, nTipoCambioAn) Inhabilitado by NAGL
    nSaldoDiario2 = 0 'oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado("2", pdFecha, nTipoCambioAn)
    Call PintaFilasExcel2(xlHoja1, "10. Subastas del Tesoro Publico(22B)", Format(nTasaSubasta1, gsFormatoNumeroView), Format(nTasaSubasta2, gsFormatoNumeroView), Format(nMontoMN, gsFormatoNumeroView), Format(nMontoME, gsFormatoNumeroView), Format(nSaldoDiario1, gsFormatoNumeroView), Format(nSaldoDiario2, gsFormatoNumeroView), lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)

   '******III. ENCAJE**********
    lnFila = lnFila + 3
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "III.  ENCAJE", "", "Moneda Nacional", "Moneda Extranjera", lnFila, False, True, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(36, pdFecha, "1", 0, 0, "1800", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(36, pdFecha, "2", 0, 0, "1800", "A1")
    lnFila = lnFila + 1
    lnPosInicial = lnFila
    ln3_1ObligSujetasEncajePos = lnPosInicial
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "1. Total de obligaciones sujetas a encaje - TOSE (23)", "", "sum(F" & (lnFila + 1) & ":F" & (lnFila + 5) & ")", "sum(G" & (lnFila + 1) & ":G" & (lnFila + 5) & ")", lnFila, True, False, True) 'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(37, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 0, "1900", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(37, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 0, "1900", "A1")
    nItemEncaje = lnFila

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.SaldoCtas(1, "761201", pdFecha, pdFechaFinDeMes, nTipoCambioAn, nTipoCambioAn) + oDbalanceCont.ObtenerPlazoFijoxRango(pdFecha, 1, 1, 30)
    nSaldoDiario2 = oDbalanceCont.SaldoCtas(1, "762201", pdFecha, pdFechaFinDeMes, nTipoCambioAn, nTipoCambioAn) + oDbalanceCont.ObtenerPlazoFijoxRango(pdFecha, 1, 2, 30) 'En obligaciones inmediateas se cambio de lnTipocambioFC A nTipoCambioAn
    Call PintaFilasExcel(xlHoja1, "1.1 Obligaciones inmediatas y a plazo hasta 30 días", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True)    'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 0, "2000", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 0, "2000", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerPlazoFijoxRango(pdFecha, 1, 1, 999999) + oDbalanceCont.SaldoAhoPlaFijCTS(Format(pdFecha, "yyyymmdd"), 1, "234") - oDbalanceCont.ObtenerPlazoFijoxRango(pdFecha, 1, 1, 30)
    nSaldoDiario2 = oDbalanceCont.ObtenerPlazoFijoxRango(pdFecha, 1, 2, 999999) + oDbalanceCont.SaldoAhoPlaFijCTS(Format(pdFecha, "yyyymmdd"), 2, "234") - oDbalanceCont.ObtenerPlazoFijoxRango(pdFecha, 1, 2, 30)
    Call PintaFilasExcel(xlHoja1, "1.2 Obligaciones a plazo mayor a 30 días", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True) 'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(39, pdFecha, "1", nSaldoDiario1, 0, "2100", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(39, pdFecha, "2", nSaldoDiario2, 0, "2100", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.SaldoAhoPlaFijCTS(Format(pdFecha, "yyyymmdd"), 1, "232") - (oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(pdFecha, "yyyymmdd"), 1, "232"))
    nSaldoDiario2 = oDbalanceCont.SaldoAhoPlaFijCTS(Format(pdFecha, "yyyymmdd"), 2, "232") - (oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(pdFecha, "yyyymmdd"), 2, "232"))
    Call PintaFilasExcel(xlHoja1, "1.3 Ahorros", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True) 'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(40, pdFecha, "1", nSaldoDiario1, 0, "2200", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(40, pdFecha, "2", nSaldoDiario2, 0, "2200", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = 0#
    nSaldoDiario2 = 0#
    Call PintaFilasExcel(xlHoja1, "1.4 Obligaciones en moneda nacional con rendimiento vinculado al tipo de cambio en moneda extranjera o a operaciones swap y similares", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True) 'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "1", nSaldoDiario1, 0, "2250", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(41, pdFecha, "2", nSaldoDiario2, 0, "2250", "A1")
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = 0#
    nSaldoDiario2 = 0#
    Call PintaFilasExcel(xlHoja1, "1.5 Otros", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True) 'anps
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 6), xlHoja1.Cells(lnFila, 7)).Font.Size = 9
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 6), xlHoja1.Cells(lnFila, 7)).Font.Name = "Calibri"
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 6), xlHoja1.Cells(lnFila, 7)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 7)).Interior.ColorIndex = 2

    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(42, pdFecha, "1", nSaldoDiario1, 0, "2280", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(42, pdFecha, "2", nSaldoDiario2, 0, "2280", "A1")

    ln3_1ObligSujetasEncajeMN = xlHoja1.Cells(ln3_1ObligSujetasEncajePos, 6)
    ln3_1ObligSujetasEncajeME = xlHoja1.Cells(ln3_1ObligSujetasEncajePos, 7)
    nTotalTOSEMN = xlHoja1.Cells(nItemEncaje, 6) 'TOSE MN
    nTotalTOSEME = xlHoja1.Cells(nItemEncaje, 7) 'TOSE ME

    nlnAdeducirAhorroMN = oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(pdFecha, "yyyymmdd"), 1, "232") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(pdFecha, "yyyymmdd"), 1, "233") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(pdFecha, "yyyymmdd"), 1, "234")
    nlnAdeducirAhorroME = oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(pdFecha, "yyyymmdd"), 2, "232") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(pdFecha, "yyyymmdd"), 2, "233") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(pdFecha, "yyyymmdd"), 2, "234")

    ln3_1ObligSujetasEncajeMN = ln3_1ObligSujetasEncajeMN - xlHoja1.Cells(lnFila, 6)
    ln3_1ObligSujetasEncajeME = ln3_1ObligSujetasEncajeME - xlHoja1.Cells(lnFila, 7)

    lnDiasToseBaseRef = oDbalanceCont.ObtenerParamEncDiarioxCodigo("10")
    lnToseBaseExigiBCRPME = oDbalanceCont.ObtenerParamEncDiarioxCodigo("06") / lnDiasToseBaseRef 'NAGL BASE DIARIO ME

    lnToseRGMN = ln3_1ObligSujetasEncajeMN - (lnToseBaseExigiBCRPMN / Day(pdFechaFinDeMes)) * 0
    lnToseRGME = ln3_1ObligSujetasEncajeME - (lnToseBaseExigiBCRPME / Day(pdFechaFinDeMes)) * 0

    oBarra.Progress 7, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    '******************************NAGL
    Dim nTotalTasaBaseEncajMN_DADiarioTotal As Currency, nTotalTasaBaseEncajME_DADiarioTotal As Currency
    Dim nTotalTotalTOSEMN As Currency, nTotalTotalTOSEME As Currency
    Dim nTasaEncajeMN As Double
    Dim nGranTotalTOSEMN As Currency, nGranTotalTOSEME As Currency
    Dim nEncajeExibleProm As Currency
    Dim lnTipoCambioProceso
    
    'JIPR20211005 MEJORAS 15A
    Dim nMontoBaseResolucion As Currency
    Dim nTasaEncajeMarginalMN As Currency
    Dim nTasaEncajeA As Currency
    Dim nEncajeMarginalB As Currency
    Dim nMontoI As Currency
    Dim nMontoII As Currency
    Dim nTasaMediaMinima As Currency
    
    Dim nMontoEncajexDia As Currency
    Dim nMontoFondoxDia As Currency
    Dim nEncajeFondoDifxDia As Currency
    
    Dim EncajeExigibleDiarioAcumulado As Currency
    Dim FondoEncajeDiarioAcumulado As Currency
    Dim ResultadoFondoExigibleDiarioAcumulado As Currency
    Dim CajaPromedioDiarioAcumuladoEncajeAnterior As Currency
    Dim CuentaCorrienteBCRPAcumulado As Currency
    
    'JIPR20211005 MEJORAS 15A
    
    nTotalTasaBaseEncajMN_DADiarioTotal = 0
    nTotalTasaBaseEncajME_DADiarioTotal = 0
    nGranTotalTOSEMN = 0
    nGranTotalTOSEME = 0

    nEncajeExibleProm = Round((oDbalanceCont.ObtenerParamEncDiarioxCodigo("08") / lnDiasToseBaseRef), 2)
    nTasaEncajeMN = (oDbalanceCont.ObtenerParamEncDiarioxCodigo("32") / 100)
    nTasaEncajeMarginalME = (oDbalanceCont.ObtenerParamEncDiarioxCodigo("03") / 100) 'Encaje Marginal ME
    nTasaEncajeME = Round(nEncajeExibleProm / lnToseBaseExigiBCRPME, 6)
    
    'JIPR20211005 MEJORAS 15A
    nMontoBaseResolucion = (oDbalanceCont.ObtenerParamEncDiarioxCodigo("36"))
    nTasaEncajeMarginalMN = (oDbalanceCont.ObtenerParamEncDiarioxCodigo("02") / 100)
    nTasaMediaMinima = (oDbalanceCont.ObtenerParamEncDiarioxCodigo("38") / 100)
    'JIPR20211005 MEJORAS 15A

    '**********Posición acumulada a la fecha TOSE y Tasa Basa Encaje
    ldFechaPro = DateAdd("d", -Day(pdFecha), pdFecha)

    For ix = 1 To Day(pdFecha)

        ldFechaPro = DateAdd("d", 1, ldFechaPro)

        nTotalTotalTOSEMN = oDbalanceCont.SaldoCtas(1, "761201", ldFechaPro, pdFechaFinDeMes, nTipoCambioAn, nTipoCambioAn) 'Obligaciones Inmediatas Antes lntipocambioFC DESPUES nTipoCambioAn
        nTotalTotalTOSEMN = nTotalTotalTOSEMN + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "232") - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "232") 'Ahorros
        nTotalTotalTOSEMN = nTotalTotalTOSEMN + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "233") + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "234")
        nTotalTotalTOSEMN = nTotalTotalTOSEMN - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "233") - oDbalanceCont.SaldoCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "233") 'Depositos a plazo fijo
        nGranTotalTOSEMN = nGranTotalTOSEMN + nTotalTotalTOSEMN

        nTotalTasaBaseEncajMN_DADiarioTotal = nTotalTasaBaseEncajMN_DADiarioTotal + (nTotalTotalTOSEMN * nTasaEncajeMN)

        nTotalTotalTOSEME = oDbalanceCont.SaldoCtas(1, "762201", ldFechaPro, pdFechaFinDeMes, nTipoCambioAn, nTipoCambioAn)
        nTotalTotalTOSEME = nTotalTotalTOSEME + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "232") - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "232") 'Ahorros
        nTotalTotalTOSEME = nTotalTotalTOSEME + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "233") + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "234")
        nTotalTotalTOSEME = nTotalTotalTOSEME - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "233") - oDbalanceCont.SaldoCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "233") 'Depositos a plazo fijo
        nGranTotalTOSEME = nGranTotalTOSEME + nTotalTotalTOSEME

        nTotalTasaBaseEncajME_DADiarioTotal = nTotalTasaBaseEncajME_DADiarioTotal + Round(IIf(nTotalTotalTOSEME > lnToseBaseExigiBCRPME, (lnToseBaseExigiBCRPME * nTasaEncajeME) + (nTotalTotalTOSEME - lnToseBaseExigiBCRPME) * nTasaEncajeMarginalME, nTotalTotalTOSEME * nTasaEncajeME), 2)

    Next ix

    'Encaje Exigible MN
    'nTotalTasaBaseEncajMN_DADiario = Round((nTotalTOSEMN * nTasaEncajeMN), 2)

    'JIPR20211005 MEJORAS 15A
    If nTotalTOSEMN > nMontoBaseResolucion Then
        nTasaEncajeA = nMontoBaseResolucion * nTasaEncajeMN
    Else
        nTasaEncajeA = nTotalTOSEMN * nTasaEncajeMN
    End If
    
    If nTotalTOSEMN > nMontoBaseResolucion Then
        nEncajeMarginalB = (nTotalTOSEMN - nMontoBaseResolucion) * nTasaEncajeMarginalMN
    Else
        nEncajeMarginalB = 0
    End If
    
    nMontoI = nTasaEncajeA + nEncajeMarginalB
    nMontoII = nTotalTOSEMN * nTasaMediaMinima
    
    If nMontoI > nMontoII Then
    nTotalTasaBaseEncajMN_DADiario = nMontoI
    Else
    nTotalTasaBaseEncajMN_DADiario = nMontoII
    End If
    
    'JIPR20211005 MEJORAS 15A

    'Encaje Exigible ME
    If nTotalTOSEME > lnToseBaseExigiBCRPME Then
        nTotalTasaBaseEncajME_DADiario = (lnToseBaseExigiBCRPME * nTasaEncajeME) + ((nTotalTOSEME - lnToseBaseExigiBCRPME) * nTasaEncajeMarginalME)
    Else
        nTotalTasaBaseEncajME_DADiario = (nTotalTOSEME * nTasaEncajeME)
    End If

    '******************************NAGL

    oBarra.Progress 8, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2. Posición de encaje", "", "", "", lnFila, False, False, False)
    Call oDbalanceCont.InsertaDetallaReporte15A(43, pdFecha, "1", 0, 0, "2300", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(43, pdFecha, "2", 0, 0, "2300", "A1")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Interior.ColorIndex = 2
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 6), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)

    lnFila = lnFila + 1
    ln2_1EncaExigible = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2.1  Encaje exigible", "", Round(nTotalTasaBaseEncajMN_DADiario, 2), Round(nTotalTasaBaseEncajME_DADiario, 2), lnFila, True, False, True)    'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(44, pdFecha, "1", nTotalTasaBaseEncajMN_DADiario, 0, "2400", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(44, pdFecha, "2", nTotalTasaBaseEncajME_DADiario, 0, "2400", "A1")

    lnEncajeExigALMN = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "2400") 'VAPA 20171215
    lnEncajeExigALME = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "2400") 'VAPA 20171215

    lnEncajeExigALMN = Round((lnEncajeExigALMN / lnTotalaMN * 100), 2) 'vapa 20171120
    lnEncajeExigALME = Round((lnEncajeExigALME / lnTotalaME * 100), 2) 'vapa 20171120

    InsertaLiquidezAlertaTemprana ldFecha, lnRatioLiquidezMN, lnRatioLiquidezME, lnRatioLAjusRecursosPrestadosMN, lnRatioLAjusRecursosPrestadosME, lnRatioInversionesLiquidasMN, lnEncajeExigALMN, lnEncajeExigALME 'VAPA 20171003

    lnFila = lnFila + 1
    ldFechaPro = DateAdd("d", -Day(pdFecha), pdFecha)
    ldFechaPro = DateAdd("d", -Day(ldFechaPro), ldFechaPro)
    nSaldoCajaDiarioMesAnteriorMN = 0
    nSaldoCajaDiarioMesAnteriorME = 0

    For ix = 1 To Day(pdFechaFinDeMesMA)
        ldFechaPro = DateAdd("d", 1, ldFechaPro)
        nSaldoCajaDiarioMesAnteriorMN = nSaldoCajaDiarioMesAnteriorMN + (oDbalanceCont.SaldoCtas(32, "761201", ldFechaPro, pdFechaFinDeMesMA, lnTipoCambioFCMA, lnTipoCambioFCMA) / Day(pdFechaFinDeMesMA))
    Next ix

    nSaldoDiario1 = nSaldoCajaDiarioMesAnteriorMN + oDbalanceCont.SaldoBCRPAnexoDiario(Format(pdFecha, "yyyymmdd"), 1)
    nSaldoDiario2 = oDbalanceCont.ObtenerCtaContSaldoDiario("1121", pdFecha) + oDbalanceCont.ObtenerCtaContSaldoDiario("112701", pdFecha) + oDbalanceCont.SaldoBCRPAnexoDiario(Format(pdFecha, "yyyymmdd"), 2)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2.2  Fondos de encaje (24)", "", "sum(F" & (lnFila + 2) & ":F" & (lnFila + 3) & ")", "+ G" & (lnFila + 1) & "+ G" & (lnFila + 3) & "", lnFila, True, False, True) 'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(45, pdFecha, "1", nSaldoDiario1, 0, "2500", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(45, pdFecha, "2", nSaldoDiario2, 0, "2500", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = nCajaFondosMN - oDbalanceCont.ObtenerParamEncDiarioxCodigo("04", "1", Format(pdFecha, "yyyymmdd"))
    nSaldoDiario2 = nCajaFondosME - oDbalanceCont.ObtenerParamEncDiarioxCodigo("04", "2", Format(pdFecha, "yyyymmdd"))
    'NAGL Agregó oDbalanceCont.ObtenerParamEncDiarioxCodigo(("04", "M", Format(pdFecha, "yyyymmdd")) Según TIC1810110004 20181015 Tanto en MN como en ME
    Call PintaFilasExcel(xlHoja1, "- Caja del día", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True)  'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(46, pdFecha, "1", nSaldoDiario1, 0, "2600", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(46, pdFecha, "2", nSaldoDiario2, 0, "2600", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = nSaldoCajaDiarioMesAnteriorMN
    nSaldoDiario2 = oDbalanceCont.ObtenerCtaContSaldoDiario("112701", pdFecha)
    Call PintaFilasExcel(xlHoja1, "- Caja promedio diario del período de encaje anterior", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True)    'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(47, pdFecha, "1", nSaldoDiario1, 0, "2650", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(47, pdFecha, "2", nSaldoDiario2, 0, "2650", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.SaldoBCRPAnexoDiario(Format(pdFecha, "yyyymmdd"), 1)
    nSaldoDiario2 = oDbalanceCont.SaldoBCRPAnexoDiario(Format(pdFecha, "yyyymmdd"), 2)
    Call PintaFilasExcel(xlHoja1, "- Cuenta corriente BCRP", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True) 'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(48, pdFecha, "1", nSaldoDiario1, 0, "2700", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(48, pdFecha, "2", nSaldoDiario2, 0, "2700", "A1")
    lnFila = lnFila + 1
    nSaldoDiario1 = Val(xlHoja1.Cells(lnFila - 4, 6)) - Val(xlHoja1.Cells(lnFila - 5, 6))
    nSaldoDiario2 = Val(xlHoja1.Cells(lnFila - 4, 7)) - Val(xlHoja1.Cells(lnFila - 5, 7))
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2.3  Resultados del día (fondos de encaje -  encaje exigible)", "", "F" & (lnFila - 4) & "-" & "F" & (lnFila - 5), "G" & (lnFila - 4) & "-" & "G" & (lnFila - 5), lnFila, True, False, True)
    Call oDbalanceCont.InsertaDetallaReporte15A(49, pdFecha, "1", nSaldoDiario1, 0, "2800", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(49, pdFecha, "2", nSaldoDiario2, 0, "2800", "A1")

    'Efectivos y Depositos de Encaje a la fecha de hoy
    lnToTalCajaFondosMN = nSaldoCajaDiarioMesAnteriorMN
    lnToTalTotalCajaFondosMN = 0
    lnToTalTotalCajaFondosME = 0
    lnTotalSaldoBCRPAnexoDiarioMN = 0
    lnTotalSaldoBCRPAnexoDiarioME = 0
    ldFechaPro = DateAdd("d", -Day(pdFecha), pdFecha)
    lnToTalOMN = 0
    lnToTalOME = 0
    For ix = 1 To Day(pdFecha)
        ldFechaPro = DateAdd("d", 1, ldFechaPro)
        lnToTalTotalCajaFondosMN = Round(lnToTalCajaFondosMN, 2) + lnToTalTotalCajaFondosMN
        'lnToTalTotalCajaFondosME = lnToTalTotalCajaFondosME + oDbalanceCont.SaldoCajasObligExoneradas(Format(ldFechaPro, "yyyymmdd"), 2)
        lnToTalTotalCajaFondosME = lnToTalTotalCajaFondosME + oDbalanceCont.ObtenerCtaContSaldoDiario("1121", ldFechaPro) + oDbalanceCont.ObtenerCtaContSaldoDiario("112701", ldFechaPro)
        lnTotalSaldoBCRPAnexoDiarioMN = lnTotalSaldoBCRPAnexoDiarioMN + oDbalanceCont.SaldoBCRPAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1)
        lnTotalSaldoBCRPAnexoDiarioME = lnTotalSaldoBCRPAnexoDiarioME + oDbalanceCont.SaldoBCRPAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2)
    Next ix
    lnToTalOMN = lnToTalTotalCajaFondosMN + lnTotalSaldoBCRPAnexoDiarioMN
    lnToTalOME = lnToTalTotalCajaFondosME + lnTotalSaldoBCRPAnexoDiarioME

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True

    'JIPR20211005 MEJORAS 15A
    
    nSaldoDiario1 = oDbalanceCont.ObtenerTotalEncajeAcumuladoFechAnx15A(pdFecha)
    'nSaldoDiario1 = lnToTalOMN - nTotalTasaBaseEncajMN_DADiarioTotal
    nSaldoDiario2 = lnToTalOME - nTotalTasaBaseEncajME_DADiarioTotal

    'JIPR20211005 MEJORAS 15A

    Call PintaFilasExcel(xlHoja1, "2.4  Posición de encaje acumulada del período a la fecha", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True) 'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(50, pdFecha, "1", nSaldoDiario1, 0, "2900", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(50, pdFecha, "2", nSaldoDiario2, 0, "2900", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    'nSaldoDiario1 = nGranTotalTOSEMN * 0.01 JIPR20200411
    'nSaldoDiario2 = nGranTotalTOSEME * 0.03 JIPR20200411
    nSaldoDiario1 = (nGranTotalTOSEMN * (oDbalanceCont.ObtenerParamEncDiarioxCodigo("34") / 100)) / Day(pdFecha) 'JIPR20211005 ENTRE LA CANTIDAD DE DIAS
    'nSaldoDiario1 = nGranTotalTOSEMN * (oDbalanceCont.ObtenerParamEncDiarioxCodigo("34") / 100) 'JIPR20200411 CORREO Cambios en el Anexo 15-B por disposición del BCRP
    nSaldoDiario2 = nGranTotalTOSEME * (oDbalanceCont.ObtenerParamEncDiarioxCodigo("35") / 100) 'JIPR20200411 CORREO Cambios en el Anexo 15-B por disposición del BCRP

    Call PintaFilasExcel(xlHoja1, "2.5  Posición acumulada del requerimiento mínimo en cuenta corriente BCRP a la fecha", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True)  'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(51, pdFecha, "1", nSaldoDiario1, 0, "2950", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(51, pdFecha, "2", nSaldoDiario2, 0, "2950", "A1")

    xlHoja1.Range(xlHoja1.Cells(lnPosInicial + 1, 2), xlHoja1.Cells(lnFila, 7)).Interior.ColorIndex = 2
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 6), xlHoja1.Cells(lnFila, 7)).Font.Size = 9
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 6), xlHoja1.Cells(lnFila, 7)).Font.Name = "Calibri"
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 6), xlHoja1.Cells(lnFila, 7)).Font.Bold = True
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 7, CCur(lnFila)

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "3. Cheques a deducir del total de obligaciones sujetas a encaje", "", "sum(F" & (lnFila + 1) & ":G" & (lnFila + 5) & ")", "sum(F" & (lnFila + 1) & ":G" & (lnFila + 5) & ")", lnFila, True, False, True) 'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(52, pdFecha, "1", xlHoja1.Cells(lnFila, 6), 0, "3000", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(52, pdFecha, "2", xlHoja1.Cells(lnFila, 7), 0, "3000", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = 0
    nSaldoDiario2 = 0
    Call PintaFilasExcel(xlHoja1, "3.1  A deducir de obligaciones a la vista y a plazo hasta 30 días", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True) 'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(53, pdFecha, "1", nSaldoDiario1, 0, "3100", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(53, pdFecha, "2", nSaldoDiario2, 0, "3100", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = 0
    nSaldoDiario2 = 0
    Call PintaFilasExcel(xlHoja1, "3.2  A deducir de obligaciones a plazo mayor de 30 días", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True) 'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(54, pdFecha, "1", nSaldoDiario1, 0, "3200", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(54, pdFecha, "2", nSaldoDiario2, 0, "3200", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = nlnAdeducirAhorroMN
    nSaldoDiario2 = nlnAdeducirAhorroME
    Call PintaFilasExcel(xlHoja1, "3.3  A deducir de ahorro", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True) 'anps
    Call oDbalanceCont.InsertaDetallaReporte15A(55, pdFecha, "1", nSaldoDiario1, 0, "3300", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(55, pdFecha, "2", nSaldoDiario2, 0, "3300", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = 0
    nSaldoDiario2 = 0
    Call PintaFilasExcel(xlHoja1, "3.4  A deducir de obligaciones en moneda nacional con rendimiento vinculado al tipo de cambio en moneda extranjera o a operaciones swap y similares", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True) 'anps
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 7)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(56, pdFecha, "1", nSaldoDiario1, 0, "3305", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(56, pdFecha, "2", nSaldoDiario2, 0, "3305", "A1")

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = 0
    nSaldoDiario2 = 0
    Call PintaFilasExcel(xlHoja1, "4. Obligaciones con entidades financieras del exterior (25)", "", Round(nSaldoDiario1, 2), Round(nSaldoDiario2, 2), lnFila, True, False, True) 'anps
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 7)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(57, pdFecha, "1", nSaldoDiario1, 0, "3310", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(57, pdFecha, "2", nSaldoDiario2, 0, "3310", "A1")

    ObtieneOtrasSeccionesAnexo15A xlHoja1.Application, pdFecha, nTipoCambioAn, lnFila  '***NAGL 20190615 Traslado de Otras Secciones del Anexo 15A en un método

    Set oEst = Nothing
    oBarra.Progress 10, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "Generación Terminada", "", vbBlue
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
    If lsArchivo <> "" Then
        CargaArchivo lsArchivo, App.path & "\SPOOLER\"
    End If
    Exit Sub
GeneraExcelErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    If lbExcel = True Then
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Sub
 'ANPS20210818 REPOPROG ********************************************************************************
Private Sub CargaReprogramados(ByVal xlAplicacion As Excel.Application, ByVal lnFila As Integer, ByVal lnPosInicial As Integer)

  Dim nSaldoDiario1 As Currency
    Dim nSaldoDiario2 As Currency
      Dim oDbalanceCont As DbalanceCont
    Set oDbalanceCont = New DbalanceCont
'JIPR20210318 REPOGARTR INICIO
    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.SaldosRepoGartAnexo("1", "1", "008")
    Call PintaFilasExcel2(xlHoja1, "REPOGARTR00898", 0, 0, 0, 0, Format(nSaldoDiario1, gsFormatoNumeroView), 0, lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)
    
    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.SaldosRepoGartAnexo("1", "1", "042")
    Call PintaFilasExcel2(xlHoja1, "REPOGARTR04298", 0, 0, 0, 0, Format(nSaldoDiario1, gsFormatoNumeroView), 0, lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)
      
    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.SaldosRepoGartAnexo("1", "1", "046")
    Call PintaFilasExcel2(xlHoja1, "REPOGARTR04695", 0, 0, 0, 0, Format(nSaldoDiario1, gsFormatoNumeroView), 0, lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)
    
    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.SaldosRepoGartAnexo("1", "1", "066")
    Call PintaFilasExcel2(xlHoja1, "REPOGARTR06695", 0, 0, 0, 0, Format(nSaldoDiario1, gsFormatoNumeroView), 0, lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)
    
    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.SaldosRepoGartAnexo("1", "1", "078")
    Call PintaFilasExcel2(xlHoja1, "REPOGARTR07898", 0, 0, 0, 0, Format(nSaldoDiario1, gsFormatoNumeroView), 0, lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)
        
    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.SaldosRepoGartAnexo("1", "1", "043")
    Call PintaFilasExcel2(xlHoja1, "REPOGARTE04398", 0, 0, 0, 0, Format(nSaldoDiario1, gsFormatoNumeroView), 0, lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)
       
    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.SaldosRepoGartAnexo("1", "1", "084")
    Call PintaFilasExcel2(xlHoja1, "REPOGARTE08498", 0, 0, 0, 0, Format(nSaldoDiario1, gsFormatoNumeroView), 0, lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)
        
    'JIPR20210318 REPOGARTR FIN
  
  lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.SaldosRepoGartAnexo("1", "1", "0277")
    Call PintaFilasExcel2(xlHoja1, "REPOPROG0277", 0, 0, 0, 0, Format(nSaldoDiario1, gsFormatoNumeroView), 0, lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)
    
     lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.SaldosRepoGartAnexo("1", "1", "0290")
    Call PintaFilasExcel2(xlHoja1, "REPOPROG0290", 0, 0, 0, 0, Format(nSaldoDiario1, gsFormatoNumeroView), 0, lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.SaldosRepoGartAnexo("1", "1", "0291")
    Call PintaFilasExcel2(xlHoja1, "REPOPROG0291", 0, 0, 0, 0, Format(nSaldoDiario1, gsFormatoNumeroView), 0, lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.SaldosRepoGartAnexo("1", "1", "0292")
    Call PintaFilasExcel2(xlHoja1, "REPOPROG0292", 0, 0, 0, 0, Format(nSaldoDiario1, gsFormatoNumeroView), 0, lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)
    
End Sub  'ANPS20210818 REPOPROG FIN************************************************************


'***********************************BEGIN NAGL 20170904*********************************************************************
Private Sub GeneraEstadisticaDiariafromBitacora(psMoneda As String, pdFecha As Date, pdFechaSist As Date, psMesBalanceDiario As String, psAnioBalanceDiario As String)
    Dim fs As New Scripting.FileSystemObject
    Dim lsMoneda As String
    Dim lsTotalActivos() As String
    Dim lsTotalPasivos() As String
    Dim lsTotalValores() As String
    Dim lsTotalRatioLiquidez() As String
    Dim lbExisteHoja As Boolean
    Dim lsTotalesActivos() As String
    Dim lsTotalesPasivos() As String
    Dim i As Long
    Dim Y1 As Integer, Y2 As Integer
    Dim Yvalor1 As Integer, Yvalor2 As Integer
    Dim lnFila As Integer
    Dim lnFilaFondosCaja As Integer
    Dim nMonto1 As Currency
    Dim nMonto2 As Currency
    Dim nMontoMN As Currency, nMontoME As Currency
    Dim nTasaSubasta1 As Double, nTasaSubasta2 As Double
    Dim oEst As New NEstadisticas
    Dim oDbalanceCont As DbalanceCont
    Dim nSaldoDiario1 As Currency
    Dim nSaldoDiario2 As Currency
    Dim nCajaFondo As Integer
    Dim oCambio As nTipoCambio
    Dim nTipoCambioAn As Currency
    Dim lnTipoCambioFC As Currency
    Dim lnTipoCambioFCMA As Currency

    Dim dFechaAnte As Date
    Dim nDia As Integer
    Dim lnPosInicial As Integer
    Dim ln2_1EncaExigible As Integer
    Dim ln3_1ObligSujetasEncajePos As Integer
    Dim ln3_1ObligSujetasEncajeMN As Currency
    Dim ln3_1ObligSujetasEncajeME As Currency
    Dim lnDiasToseBaseRef As Currency
    Dim lnToseBaseExigiBCRPMN As Currency
    Dim lnToseBaseExigiBCRPME As Currency
    Dim nTotalObligSugEncajMN As Currency
    Dim nTotalObligSugEncajME As Currency

    Dim nTotalTasaBaseEncajMN_DADiario As Currency
    Dim nTotalTasaBaseEncajME_DADiario As Currency
    Dim nlnAdeducirAhorroMN As Currency
    Dim nlnAdeducirAhorroME As Currency
    Dim lnTasaEncajeMN As Double
    Dim lnTasaEncajeME As Double

    Dim ix As Integer
    Dim lnToseRGMN As Currency
    Dim lnToseRGME As Currency
    Dim ldFechaPro As Date
    Dim nSaldoCajaDiarioMesAnteriorMN As Currency
    Dim nSaldoCajaDiarioMesAnteriorME As Currency

    Dim lnToTalTotalCajaFondosMN As Currency
    Dim lnToTalTotalCajaFondosME As Currency
    Dim lnToTalOME As Currency
    Dim lnToTalOMN As Currency
    Dim lnToTalCajaFondosMN As Currency
    Dim lnToTalCajaFondosME As Currency
    Dim lnTotalSaldoBCRPAnexoDiarioMN As Currency
    Dim lnTotalSaldoBCRPAnexoDiarioME As Currency
    Dim lnTotalObligacionesAlDiaMN As Currency
    Dim lnTotalObligacionesAlDiaME As Currency

    Dim nSubValor1 As Currency
    Dim nSubValor2 As Currency
    Dim pdFechaFinDeMes As Date, pdFechaBalanceDiario As Date
    Dim pdFechaFinDeMesMA As Date
    Dim nTotalTOSEMN As Currency
    Dim nTotalTOSEME As Currency
    Dim nTasaEncajeME As Double

    Dim nTasaEncajeMarginalME As Currency
    Dim nItemEncaje As Integer

    Dim oValor As New DAnexoRiesgos
    Dim rsvalor As New ADODB.Recordset
    Dim oCtaIf As New NCajaCtaIF
    Dim rsDetalleCtas As New ADODB.Recordset

    Dim nContar As Integer
    Dim nContar2 As Integer
    Dim X As Integer
    Dim nSaldoPFN As Double, nSaldoPFME As Double
    Dim nSumarTotalMN As Double
    Dim nSumarMN As Double
    Dim nTipoCambioBalance As Currency
    Dim nsDiv As Double
    Dim nsDiv2 As Double, nsDivME As Double
    Dim nPromedio As Double
    Dim nObligMNDiario1 As Currency, nObligMEDiario2 As Currency
    'Dim nUltimoTOSE
    'INICIO VAPA20170909
    Dim lnRatioLiquidezMN As Double
    Dim lnRatioLiquidezME As Double
    Dim lnRatioLAjusRecursosPrestadosMN As Double
    Dim lnRatioLAjusRecursosPrestadosME As Double
    Dim lnRatioInversionesLiquidasMN As Double
    Dim lnEncajeExigALMN As Double
    Dim lnEncajeExigALME As Double
    Dim lnTotalaMN As Double
    Dim lnTotalaME As Double
    'VAPA20170909 END

    On Error GoTo GeneraExcelErr
    
    Set oBarra = New clsProgressBar
    Unload Me
    oBarra.ShowForm frmReportes
    oBarra.Max = 10
    oBarra.Progress 0, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    pdFechaFinDeMes = DateAdd("d", -Day(DateAdd("m", 1, pdFecha)), DateAdd("m", 1, pdFecha))
    pdFechaFinDeMesMA = DateAdd("d", -Day(pdFechaFinDeMes), pdFechaFinDeMes)
    nDia = Day(pdFecha)
    Set oDbalanceCont = New DbalanceCont
    
    lsArchivo = App.path & "\SPOOLER\" & "Anx15A_" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx" 'NAGL 20170415
    lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    ExcelAddHoja "Anx15A", xlLibro, xlHoja1
    
    Set oCambio = New nTipoCambio
    If CInt(psMesBalanceDiario) < 9 Then
        pdFechaBalanceDiario = CDate("01" & "/" & "0" & CStr(CInt(psMesBalanceDiario) + 1) & "/" & psAnioBalanceDiario)
    ElseIf CInt(psMesBalanceDiario) = 12 Then
        pdFechaBalanceDiario = CDate("01" & "/" & "01" & "/" & CStr(CInt(psAnioBalanceDiario) + 1))
    Else
        pdFechaBalanceDiario = CDate("01" & "/" & CStr(CInt(psMesBalanceDiario) + 1) & "/" & psAnioBalanceDiario)
    End If

    nTipoCambioBalance = Format(oCambio.EmiteTipoCambio(pdFechaBalanceDiario, TCFijoDia), "#,##0.0000")  'NAGL 20170425

    If Month(pdFecha) = Month(DateAdd("d", 1, pdFecha)) Then
        lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(pdFecha, TCFijoDia), "#,##0.0000")
    Else
        lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(DateAdd("d", -1, pdFecha), TCFijoDia), "#,##0.0000") 'SE CAMBIO DE DateAdd("d", 1, pdFecha) A DateAdd("d", -1, pdFecha)
    End If

    nTipoCambioAn = lnTipoCambioFC

    ldFecha = pdFecha
    lsMoneda = Mid(gsOpeCod, 3, 1)
    xlHoja1.PageSetup.Zoom = 100
    For i = 2 To 30
        If i <> 6 Then
            xlHoja1.Range(xlHoja1.Cells(i, 3), xlHoja1.Cells(i, 5)).Merge True
        End If
    Next

    ReDim lsTotalActivos(2)
    ReDim lsTotalPasivos(2)
    ReDim lsTotalValores(2) '****NAGL
    ReDim lsSumaTotalBCRP(2)   '*******NAGL
    ReDim lsTotalRatioLiquidez(2)

    xlHoja1.Range("A1:R500").Font.Size = 10 'NAGL 20190614 Cambio de 9 a 10
    xlHoja1.Range("A1:R500").Font.Name = "Arial Narrow" 'NAGL 20190614

    xlHoja1.Range("A1").ColumnWidth = 7
    xlHoja1.Range("B1").ColumnWidth = 30
    xlHoja1.Range("C1").ColumnWidth = 18 '36
    xlHoja1.Range("D1:E1").ColumnWidth = 17
    xlHoja1.Range("F1:G1").ColumnWidth = 14.29 '15
    xlHoja1.Range("H1").ColumnWidth = 13 '10

    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(6, 8)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(10, 8)).HorizontalAlignment = xlCenter

    xlHoja1.Range("B1:B50").HorizontalAlignment = xlLeft

    lnFila = 1
    xlHoja1.Cells(lnFila, 2) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "ANEXO Nº 15A"
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "REPORTE DE TESORERIA Y POSICION DE LIQUIDEZ"
    'lnFila = lnFila + 1
    'xlHoja1.Cells(lnFila, 3) = "(EN NUEVOS SOLES)"
    lnFila = lnFila + 2
    xlHoja1.Cells(lnFila, 2) = "EMPRESA: " & gsNomCmac
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).MergeCells = True

    xlHoja1.Cells(lnFila, 6) = "Fecha: " & Format(pdFecha, "dd mmmm yyyy")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).HorizontalAlignment = xlLeft

    lnFila = lnFila + 2

    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)

    '*******************I RATIOS DE LIQUIDEZ**************************

    xlHoja1.Cells(lnFila, 2) = "I RATIOS DE LIQUIDEZ"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Borders.LineStyle = xlContinuous
    lnFila = lnFila + 1
    'Comentado by NAGL 20190614
    'xlHoja1.Cells(lnFila, 6) = "MONEDA ": xlHoja1.Cells(lnFila, 7) = "MONEDA "
    'lnFila = lnFila + 1
    'xlHoja1.Cells(lnFila, 6) = "NACIONAL": xlHoja1.Cells(lnFila, 7) = "EXTRANJERA"

    '**Agregado by NAGL 20190614
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila + 1, 2)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila + 1, 2)).VerticalAlignment = xlJustify

    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila + 1, 5)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila + 1, 5)).VerticalAlignment = xlJustify

    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila + 1, 7)).Font.Bold = True
    xlHoja1.Cells(lnFila, 6) = "MONEDA NACIONAL"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila + 1, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila + 1, 6)).VerticalAlignment = xlJustify

    xlHoja1.Cells(lnFila, 7) = "MONEDA EXTRANJERA"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila + 1, 7)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila + 1, 7)).VerticalAlignment = xlJustify
    lnFila = lnFila + 1
    '*******************************

    ExcelCuadro xlHoja1, 2, lnFila - 2, 7, CCur(lnFila)    'ExcelCuadro 2, 8, 7, 10

    oBarra.Progress 1, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    lnFila = lnFila + 1 'FILA 10
    Y1 = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
    xlHoja1.Cells(lnFila, 3) = "Activos Líquidos"
    lnFila = lnFila + 1

    nCajaFondo = lnFila

    Dim nCajaFondosMN As Currency, nCajaFondosME As Currency

    nCajaFondosMN = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "300")
    nCajaFondosME = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "300")
    Call PintaFilasExcel(xlHoja1, "1101+1107.01", "Caja y Fondos Fijos", nCajaFondosMN, nCajaFondosME, lnFila, True, False, True)

    lsTotalActivos(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalActivos(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)
    lnFila = lnFila + 1

    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "425")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "425")
    Call PintaFilasExcel(xlHoja1, "1102+1108.02", "Fondos disponibles en el BCRP", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    oBarra.Progress 2, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue
    lnFila = lnFila + 1

    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "450")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "450")
    Call PintaFilasExcel(xlHoja1, "1103+1108.03", "Fondos disponibles en empresas del sistema financiero nacional (2)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "500")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "500")
    Call PintaFilasExcel(xlHoja1, "1104.01+1108.04(p)", "Fondos disponibles en bancos del exterior de primera categoría (3)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "600")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "600")
    Call PintaFilasExcel(xlHoja1, "1200-2200", "Fondos interbancarios netos activos (4)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    Yvalor1 = lnFila
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "725")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "725")
    Call PintaFilasExcel(xlHoja1, "1302.02.01+1304.02.01+1305.02.01", "Valores representativos de deuda emitidos por el BCRP (5)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lsTotalValores(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "750")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "750")
    Call PintaFilasExcel(xlHoja1, "1302.01.01.01+1304.01.01.01+1305.01.01.01", "Valores representativos de deuda emitidos por el Gobierno Central (6)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lsTotalValores(1) = lsTotalValores(1) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)   '*********NAGL
    oBarra.Progress 3, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "800")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "800")
    Call PintaFilasExcel(xlHoja1, "1302.05.12+1302.05.19(p)+1304.05.12+1304.05.19(p) +1309.04.05.12+1309.04.05.19(p)", "Certificados de depósito negociables y certificados bancarios (7)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "900")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "900")
    Call PintaFilasExcel(xlHoja1, "1302.01.01.02(p)+1304.01.01.02(p)+1302.05(p)+1302.06(p)+1304.05(p)+1304.06(p)+1305.01.01.02(p) +1309.04.01.01(p) +1309.04.05(p) +1309.04.06(p) +1309.05.01.01(p)", "Valores representativos de deuda pública y de los sistemas financiero y de seguros del exterior (8)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    '************Agregado by NAGL 20190613 Según Anexo02 - ERS006-2019*************
    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "910")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "910")
    Call PintaFilasExcel(xlHoja1, "1302(p)+1304(p)+1305(p)", "Bonos corporativos emitidos por empresas privadas del sector no financiero (8A)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "920")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "920")
    Call PintaFilasExcel(xlHoja1, "1507.11(p)", "Valores representativos de deuda del BCRP y Gobierno Central recibidos en operaciones de reporte (8B)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "930")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "930")
    Call PintaFilasExcel(xlHoja1, "1507.11(p)", "Valores representativos de deuda de Gobiernos del Exterior recibidos en operaciones de reporte (8B)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "940")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "940")
    Call PintaFilasExcel(xlHoja1, "1507.11(p)", "Bonos corporativos emitidos por empresas privadas del sector no financiero recibidos en operaciones de reporte (8B)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    '*************************END NAGL 20190613************************************

    lsTotalActivos(1) = lsTotalActivos(1) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalActivos(2) = lsTotalActivos(2) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)

    oBarra.Progress 4, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue
    '*************** TOTALES ACTIVOS DE LIQUIDEZ *************************

    lnFila = lnFila + 3
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "", "Total(a)", "Sum(" & lsTotalActivos(1) & ")", "Sum(" & lsTotalActivos(2) & ")", lnFila, True, True, True)

    lnTotalaMN = Format(xlHoja1.Cells(lnFila, 6), "#,##0.00;-#,##0.00") 'VAPA20170911
    lnTotalaME = Format(Round(xlHoja1.Cells(lnFila, 7), 2), "#,##0.00;-#,##0.00") 'VAPA20170911

    ReDim lsTotalesActivos(2)
    ReDim lsTotalesPasivos(2)

    '****************PASIVOS DE CORTO PLAZO*****************
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).HorizontalAlignment = xlCenter
    Call PintaFilasExcel(xlHoja1, "", "Pasivos de Corto Plazo", "", "", lnFila, False, True, False)

    lnFila = lnFila + 1
    lsTotalPasivos(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalPasivos(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)

    xlHoja1.Cells(lnFila, 2) = "2101+2301(p)+2108.01+2308.01(p)"
    xlHoja1.Cells(lnFila, 3) = "Obligaciones a la vista (9)"

    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "1210")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "1210")
    Call PintaFilasExcel(xlHoja1, "2101+2301(p)+2108.01+2308.01(p)", "Obligaciones a la vista (9)", nSaldoDiario1, nSaldoDiario2, lnFila, False, False, True)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "1225")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "1225")
    Call PintaFilasExcel(xlHoja1, "2507.03+2507.04+2507.05+2507.06+2508(p)", "Obligaciones con instituciones recaudadoras de tributos (10)", nSaldoDiario1, nSaldoDiario2, lnFila, False, False, True) 'NAGL Cambio el último de False a True

    lnFila = lnFila + 1
    ' nSaldoDiario1 = MontoDesembolsoReactiva(pdFecha, 1)  'JIPR20200522
    ' nSaldoDiario1 = MontoDesembolsoReactiva(pdFecha, 2)  'JIPR20200522
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2504.11(p)", "Cuentas por pagar por operaciones de reporte (34)", 0, 0, lnFila, False, False, True)


    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2504.12", "Cuentas por pagar por ventas en corto(11)", "", "", lnFila, False, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2200-1200", "Fondos interbancarios netos pasivos (4)", "", "", lnFila, False, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "1400")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "1400")
    Call PintaFilasExcel(xlHoja1, "2102+2302(p)+2108.02+2308.02(p)", "Obligaciones por cuentas de ahorro", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "1450")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "1450")
    Call PintaFilasExcel(xlHoja1, "2103(p)-2103.05(p)+2303(p)+2108.03(p)+2308.03(p)", "Obligaciones por cuentas a plazo (12)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "1500")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "1500")
    Call PintaFilasExcel(xlHoja1, "2401+2402+2403+2406 +2408.01+2408.02+2408.03 +2408.06+2409.01+2602(p)+2603(p)+2606(p)+2608.02(p)+2608.03(p)+2608.06(p)+2609.01", "Adeudos y obligaciones financieras del país (13)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "1510")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "1510")
    Call PintaFilasExcel(xlHoja1, "2404+2405+2407+2408.04+2408.05+2408.07+2409.02+2409.03+2604(p)+2605(p)+2607(p)+2608.04(p)+2608.05(p)+2608.07(p)+2609.02+2609.03", "Adeudos y obligaciones financieras del exterior (13)", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2800(p)+2103.01.01(p)+2108.03(p)", "Valores, títulos y obligaciones en circulación (14)", "0", "0", lnFila, True, False, True)

    oBarra.Progress 5, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    lsTotalPasivos(1) = lsTotalPasivos(1) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalPasivos(2) = lsTotalPasivos(2) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)

    '******************** TOTALES DE PASIVOS DE CORTO PLAZO ***************************

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True 'NAGL 20190614

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).HorizontalAlignment = xlCenter
    Call PintaFilasExcel(xlHoja1, "", "Total(b)", "Sum(" & lsTotalPasivos(1) & ")", "Sum(" & lsTotalPasivos(2) & ")", lnFila, True, True, True)

    lsTotalesPasivos(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalesPasivos(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)

    Y2 = lnFila
    xlHoja1.Range(xlHoja1.Cells(Y1, 2), xlHoja1.Cells(Y2, 7)).Borders.LineStyle = xlContinuous

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Merge True 'NAGL 20190614

    'Ratios de Liquidez[(a)/(b)]*100
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "Ratios de Liquidez[(a)/(b)]*100", "", "(sum(" & lsTotalActivos(1) & ")/sum(" & lsTotalPasivos(1) & "))*100", "(sum(" & lsTotalActivos(2) & ")/sum(" & lsTotalPasivos(2) & "))*100", lnFila, True, True, True)

    lnRatioLiquidezMN = Format(xlHoja1.Cells(lnFila, 6), "#,##0.00;-#,##0.00") 'VAPA20170911
    lnRatioLiquidezMN = Format(Round(xlHoja1.Cells(lnFila, 6), 2), "#,##0.00;-#,##0.00") 'VAPA20170911
    lnRatioLiquidezME = Format(Round(xlHoja1.Cells(lnFila, 7), 2), "#,##0.00;-#,##0.00") 'VAPA20170911

    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
     
     Set rsvalor = oValor.ObligacionBN(pdFecha) '***NAGL
     If Not (rsvalor.EOF And rsvalor.BOF) Then
        nObligMNDiario1 = rsvalor!mObligacionMN
        nObligMEDiario2 = rsvalor!mObligacionME
    Else
        nObligMNDiario1 = 0
        nObligMEDiario2 = 0
    End If

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "Activos líquidos ajustados por recursos prestados (c)(15)", "", "Sum(" & lsTotalActivos(1) & ") - " & nObligMNDiario1, "Sum(" & lsTotalActivos(2) & ")-" & nObligMEDiario2, lnFila, True, False, True) '********NAGL

    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "1720")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "1720")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "Pasivos de corto plazo ajustados por recursos prestados (d)(15)", "", "Sum(" & lsTotalPasivos(1) & ") - " & nObligMNDiario1, "Sum(" & lsTotalPasivos(2) & ")-" & nObligMEDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    nSaldoDiario1 = xlHoja1.Cells(lnFila - 2, 6) / xlHoja1.Cells(lnFila - 1, 6)
    nSaldoDiario2 = xlHoja1.Cells(lnFila - 2, 7) / xlHoja1.Cells(lnFila - 1, 7)
    lnRatioLAjusRecursosPrestadosMN = nSaldoDiario1 * 100 'VAPA20170911
    lnRatioLAjusRecursosPrestadosME = nSaldoDiario2 * 100 'VAPA20170911

    '***NAGL 20181121
    Call PintaFilasExcel(xlHoja1, "Ratio de liquidez ajustado por recursos prestados [(c)/(d)]x100", "", "(" & xlHoja1.Range(xlHoja1.Cells(lnFila - 2, 6), xlHoja1.Cells(lnFila - 2, 6)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 6), xlHoja1.Cells(lnFila - 1, 6)).Address(False, False) & ")" & "*" & "100", "(" & xlHoja1.Range(xlHoja1.Cells(lnFila - 2, 7), xlHoja1.Cells(lnFila - 2, 7)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 7), xlHoja1.Cells(lnFila - 1, 7)).Address(False, False) & ")" & "*" & "100", lnFila, True, True, True)
    '***END NAGL

    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "Posiciones largas en forwards de monedas (e) (15)", "", "", "", lnFila, False, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "Posiciones cortas en forwards de monedas (f) (15)", "", "", "", lnFila, False, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "Ratio de liquidez ajustado por forwards de monedas [((a)+(e))/((b)+(f))]x100", "", "(sum(" & lsTotalActivos(1) & ")/sum(" & lsTotalPasivos(1) & "))*100", "(sum(" & lsTotalActivos(2) & ")/sum(" & lsTotalPasivos(2) & "))*100", lnFila, True, True, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    lsTotalRatioLiquidez(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalRatioLiquidez(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)

    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "1755")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "1755")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "Depósitos overnight BCRP (g)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "1760")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "1760")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "Depósitos plazo BCRP (h)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True '**********NAGL  lsTotalValores(1)
    Call PintaFilasExcel(xlHoja1, "Valores representativos de deuda emitidos por el BCRP y Gobierno Central (i)", "", "Sum(" & lsTotalValores(1) & ")", "0.00", lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lsTotalRatioLiquidez(1) = lsTotalRatioLiquidez(1) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalRatioLiquidez(2) = lsTotalRatioLiquidez(2) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)

    oBarra.Progress 6, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "Valores representativos de deuda emitidos por Gobiernos del Exterior (j)", "", "", "", lnFila, False, False, False)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    'Ratio de inversiones liquidas  [((g)+(h)+(i)+(j))/(a)]x100
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "Ratio de inversiones liquidas  [((g)+(h)+(i)+(j))/(a)]x100", "", "(sum(" & lsTotalRatioLiquidez(1) & ")/sum(" & lsTotalActivos(1) & "))*100", "(sum(" & lsTotalRatioLiquidez(2) & ")/sum(" & lsTotalActivos(2) & "))*100", lnFila, True, True, True)
    'Call PintaFilasExcel(xlHoja1, "Ratio de inversiones liquidas  [((g)+(h)+(i)+(j))/(a)]x100", "", "((sum(" & lsTotalRatioLiquidez(1) & ")+" & xlHoja1.Range(xlHoja1.Cells(Yvalor1, 6), xlHoja1.Cells(Yvalor1, 6)).Address(False, False) & ")" & "/sum(" & lsTotalActivos(1) & "))*100", "(" & xlHoja1.Range(xlHoja1.Cells(lnFila - 4, 7), xlHoja1.Cells(lnFila - 4, 7)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(Yvalor1, 7), xlHoja1.Cells(Yvalor1, 7)).Address(False, False) & ")/sum(" & lsTotalActivos(2) & ")*100", lnFila, True, True, True) 'NAGL 20190619

    lnRatioInversionesLiquidasMN = Format(Round(xlHoja1.Cells(lnFila, 6), 2), "#,##0.00;-#,##0.00") 'VAPA20170911
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 3  'FILA 55
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Merge True
    Call PintaFilasExcel(xlHoja1, "II.  OTRAS OPERACIONES", "", "", "", lnFila, False, True, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Borders.LineStyle = xlContinuous
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)

    lnFila = lnFila + 1

    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 8)).Merge True

    'Uso de PintaFilasExcel2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Cells(lnFila, 5) = "Monto(16A)"
    Call PintaFilasExcel(xlHoja1, "", "Tasas de interés (16)", "", "Saldos(16B)", lnFila, False, True, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)

     '****Agregado by NAGL 20190614
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila + 1, 2)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila + 1, 2)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Cells(lnFila, 3) = "Moneda ": xlHoja1.Cells(lnFila, 4) = "Moneda ": xlHoja1.Cells(lnFila, 5) = "Moneda ": xlHoja1.Cells(lnFila, 6) = "Moneda ": xlHoja1.Cells(lnFila, 7) = "Moneda ": xlHoja1.Cells(lnFila, 8) = "Moneda "
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "Nacional": xlHoja1.Cells(lnFila, 4) = "Extranjera": xlHoja1.Cells(lnFila, 5) = "Nacional": xlHoja1.Cells(lnFila, 6) = "Extranjera": xlHoja1.Cells(lnFila, 7) = "Nacional": xlHoja1.Cells(lnFila, 8) = "Extranjera"
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila - 1), 8, CCur(lnFila)
    '*****************END 20190614*******************'

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel2(xlHoja1, "1. Operaciones overnight(17)", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1 'CELDA PENDIENTE
    Yvalor2 = lnFila

    lnFila = lnFila + 1
    nSaldoDiario1 = 0
    nSaldoDiario2 = 0

    Call PintaFilasExcel2(xlHoja1, "       1.1.1 Empresas del sistema financiero", "", "", nSaldoDiario1, nSaldoDiario2, "", "", lnFila, False, False, True)
    lnFila = lnFila + 1

    'Para el Cálculo en Sección 1.1.2 Otras
    nSumarMN = 0
    nSaldoPFN = 0
    nSaldoPFME = 0
    nsDiv = 0
    nsDiv2 = 0
    nsDivME = 0

    Set rsvalor = oValor.MuestraBCRP(pdFecha, "C[BD]")
    nContar = 0

    Set rsvalor = oValor.MuestraBCRP(pdFecha, "C[BD]")
    nContar = 0
    If Not (rsvalor.EOF And rsvalor.BOF) Then
        For X = 0 To rsvalor.RecordCount - 1
            lnFila = lnFila + 1
            Call PintaFilasExcel2(xlHoja1, "                      - BCRP", Format(rsvalor!nTasaInteres, gsFormatoNumeroView), "", Format(rsvalor!nValorRazonable, gsFormatoNumeroView), "", "", "", lnFila, False, False, True)
            nSumarMN = rsvalor!nValorRazonable + nSumarMN 'total valor razonable CDBCRP
            nContar = nContar + 1
            rsvalor.MoveNext
        Next
    Else
        nSumarMN = 0
        nContar = 0
    End If
    
    Set rsDetalleCtas = oCtaIf.obtenerBCRInversionDPF(pdFecha)
    nContar2 = 0
    If Not (rsDetalleCtas.EOF And rsDetalleCtas.BOF) Then
        Do While Not rsDetalleCtas.EOF
            lnFila = lnFila + 1
            Call PintaFilasExcel2(xlHoja1, "                      - BCRP", Format(rsDetalleCtas!TEAMN, gsFormatoNumeroView), Format(rsDetalleCtas!TEAME, gsFormatoNumeroView), Format(rsDetalleCtas!nSaldoMN, gsFormatoNumeroView), Format(rsDetalleCtas!nSaldoME, gsFormatoNumeroView), "", "", lnFila, False, False, True)
            nSaldoPFN = rsDetalleCtas!nSaldoMN + nSaldoPFN  'total overnight BCRP MN
            nSaldoPFME = rsDetalleCtas!nSaldoME + nSaldoPFME 'total overnight BCRP ME
            nContar2 = nContar2 + 1
            rsDetalleCtas.MoveNext
        Loop
    Else
        nContar2 = 0
        nSaldoPFN = 0
        nSaldoPFME = 0
    End If
    '***NAGL***

    nSumarTotalMN = nSumarMN + nSaldoPFN 'total entre valor razon. y overnight MN
  
    Set rsvalor = oValor.MuestraBCRP(pdFecha, "C[BD]")
    If Not (rsvalor.EOF And rsvalor.BOF) Then
        For X = 0 To rsvalor.RecordCount - 1
            nsDiv = (rsvalor!nValorRazonable / nSumarTotalMN) * rsvalor!nTasaInteres + nsDiv
            rsvalor.MoveNext
        Next
    Else
        nsDiv = 0
    End If
    
    Set rsDetalleCtas = oCtaIf.obtenerBCRInversionDPF(pdFecha)
    
    If Not (rsDetalleCtas.EOF And rsDetalleCtas.BOF) Then
        Do While Not rsDetalleCtas.EOF
            If (nSaldoPFN = 0) Then
                nsDiv2 = 0
            Else
                nsDiv2 = (rsDetalleCtas!nSaldoMN / nSumarTotalMN) * rsDetalleCtas!TEAMN + nsDiv2
            End If
            If (nSaldoPFME = 0) Then
                nsDivME = 0
            Else
                nsDivME = (rsDetalleCtas!nSaldoME / nSaldoPFME) * rsDetalleCtas!TEAME + nsDivME
            End If
            rsDetalleCtas.MoveNext
        Loop
    Else
        nsDiv2 = 0
        nsDivME = 0
    End If

    nPromedio = nsDiv + nsDiv2

    lnFila = lnFila - nContar - nContar2 - 2
    Call PintaFilasExcel2(xlHoja1, "   1.1 Activas", Format(nPromedio, gsFormatoNumeroView), Format(nsDivME, gsFormatoNumeroView), Format(nSumarTotalMN, gsFormatoNumeroView), Format(nSaldoPFME, gsFormatoNumeroView), "", "", lnFila, False, False, True)

    lnFila = lnFila + 2
    Call PintaFilasExcel2(xlHoja1, "   1.1.2 Otras", Format(nPromedio, gsFormatoNumeroView), Format(nsDivME, gsFormatoNumeroView), Format(nSumarTotalMN, gsFormatoNumeroView), Format(nSaldoPFME, gsFormatoNumeroView), "", "", lnFila, False, False, True)

    lnFila = lnFila + nContar + nContar2 + 1  '******* NAGL ANTES :lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "   1.2 Pasivas", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       1.2.1 Empresas del sistema financiero", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       1.2.2 Otras", "", "", "", "", "", "", lnFila, False, False, False)
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)
    '****NAGL 20190614
    xlHoja1.Range(xlHoja1.Cells(Yvalor2, 7), xlHoja1.Cells(lnFila, 8)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(Yvalor2 - 1, 2), xlHoja1.Cells(lnFila, 2)).Interior.ColorIndex = 2
    xlHoja1.Range(xlHoja1.Cells(Yvalor2, 3), xlHoja1.Cells(lnFila, 6)).Interior.ColorIndex = 2
    '******************

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Call PintaFilasExcel2(xlHoja1, "2. Fondos interbancarios", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).Interior.Color = RGB(153, 153, 255)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "      2.1 Activos (Cuenta 1201)", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "      2.2 Pasivos (Cuenta 2201)", "", "", "", "", "", "", lnFila, False, False, False)
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 7), xlHoja1.Cells(lnFila, 8)).Interior.Color = RGB(153, 153, 255)

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Call PintaFilasExcel2(xlHoja1, "3. Obligaciones con el Banco de la Nación (18)", "", "", "", "", Format(nObligMNDiario1, gsFormatoNumeroView), Format(nObligMEDiario2, gsFormatoNumeroView), lnFila, False, False, False)
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Call PintaFilasExcel2(xlHoja1, "4. Operaciones de venta con compromiso de recompra y operaciones de compra y venta simultánea de valores (19)", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 38.25
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "   4.1 Adquiriente", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       4.1.1 Con instrumentos de inversión del BCRP y del Tesoro Público", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 25.5
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify

    '*****************Agregado by NAGL 20190614********************'
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.1.1.1 Empresas del sistema financiero", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.1.1.2 Otras", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       4.1.2 Con otros ALAC", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.1.2.1 Empresas del sistema financiero", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.1.2.2 Otras", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       4.1.3 Con otros Instrumentos", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.1.3.1 Empresas del sistema financiero", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.1.3.2 Otras", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "   4.2 Enajenante", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       4.2.1 Con instrumentos de inversión del BCRP y del Tesoro Público", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 25.5
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.2.1.1 Empresas del sistema financiero", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.2.1.2 Otras", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       4.2.2 Con otros ALAC", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.2.2.1 Empresas del sistema financiero", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.2.2.2 Otras", "", "", "", "", "", "", lnFila, False, False, False)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "       4.2.3 Con otros Instrumentos", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.2.3.1 Empresas del sistema financiero", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "         4.2.3.2 Otras", "", "", "", "", "", "", lnFila, False, False, False)

    xlHoja1.Range(xlHoja1.Cells(lnPosInicial - 3, 2), xlHoja1.Cells(lnPosInicial - 3, 2)).Interior.ColorIndex = 2
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial - 2, 2), xlHoja1.Cells(lnFila, 6)).Interior.ColorIndex = 2
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2

    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)
    '**************************END NAGL**************************'

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Call PintaFilasExcel2(xlHoja1, "5. Transferencia temporal de valores (20)", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "      5.1 Con activos líquidos", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "      5.2 Con activos no líquidos", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Call PintaFilasExcel2(xlHoja1, "6. Créditos del BCRP con fines de regulación monetaria", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 25.5
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)

    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "7. Operaciones de reporte de monedas con el BCRP (21)", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 25.5
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Call PintaFilasExcel2(xlHoja1, "      7.1 Repo Regular", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "      7.2 Repo Expansión", "", "", "", "", "", "", lnFila, False, False, False)
    lnFila = lnFila + 1
    Call PintaFilasExcel2(xlHoja1, "      7.3 Repo Sustitución", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    Call PintaFilasExcel2(xlHoja1, "8.  Swaps cambiarios con el BCRP (22)", "", "", "", "", "", "", lnFila, False, False, False)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)

    lnFila = lnFila + 1

    Call PintaFilasExcel2(xlHoja1, "9.  Operaciones de reporte de cartera de créditos con el BCRP (22A)", "", "", "", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlJustify
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)

    lnFila = lnFila + 1
    lnPosInicial = lnFila

    Set rsvalor = oDbalanceCont.ObtenerMontosyTasasSubastaTesoroPublico(pdFecha)
    If Not rsvalor.BOF And Not rsvalor.EOF Then
        Do While Not rsvalor.EOF
            nTasaSubasta1 = nTasaSubasta1 + rsvalor!TasaParcialMN
            nTasaSubasta2 = nTasaSubasta2 + rsvalor!TasaParcialME
            nMontoMN = nMontoMN + rsvalor!MontoMN
            nMontoME = nMontoME + rsvalor!MontoME
            rsvalor.MoveNext
        Loop
    Else
        nTasaSubasta1 = 0
        nTasaSubasta2 = 0
        nMontoMN = 0
        nMontoME = 0
    End If
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "3322")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "3322")
    Call PintaFilasExcel2(xlHoja1, "10. Subastas del Tesoro Publico(22B)", Format(nTasaSubasta1, gsFormatoNumeroView), Format(nTasaSubasta2, gsFormatoNumeroView), Format(nMontoMN, gsFormatoNumeroView), Format(nMontoME, gsFormatoNumeroView), Format(nSaldoDiario1, gsFormatoNumeroView), Format(nSaldoDiario2, gsFormatoNumeroView), lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 8, CCur(lnFila)

    '******ENCAJE**********
    lnFila = lnFila + 3
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "III.  ENCAJE", "", "Moneda Nacional", "Moneda Extranjera", lnFila, False, True, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    ln3_1ObligSujetasEncajePos = lnPosInicial
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "1. Total de obligaciones sujetas a encaje - TOSE (23)", "", "sum(F" & (lnFila + 1) & ":F" & (lnFila + 5) & ")", "sum(G" & (lnFila + 1) & ":G" & (lnFila + 5) & ")", lnFila, True, False, True)

    nItemEncaje = lnFila

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "2000")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "2000")
    Call PintaFilasExcel(xlHoja1, "1.1 Obligaciones inmediatas y a plazo hasta 30 días", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "2100")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "2100")
    Call PintaFilasExcel(xlHoja1, "1.2 Obligaciones a plazo mayor a 30 días", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "2200")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "2200")
    Call PintaFilasExcel(xlHoja1, "1.3 Ahorros", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = 0#
    nSaldoDiario2 = 0#
    Call PintaFilasExcel(xlHoja1, "1.4 Obligaciones en moneda nacional con rendimiento vinculado al tipo de cambio en moneda extranjera o a operaciones swap y similares", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = 0#
    nSaldoDiario2 = 0#
    Call PintaFilasExcel(xlHoja1, "1.5 Otros", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 6), xlHoja1.Cells(lnFila, 7)).Font.Size = 9
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 6), xlHoja1.Cells(lnFila, 7)).Font.Name = "Calibri"
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 6), xlHoja1.Cells(lnFila, 7)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 7)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 7, CCur(lnFila)

    oBarra.Progress 8, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue
    lnFila = lnFila + 1
    lnPosInicial = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2. Posición de encaje", "", "", "", lnFila, False, False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Interior.ColorIndex = 2
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 6), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)

    lnFila = lnFila + 1
    ln2_1EncaExigible = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "2400")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "2400")
    Call PintaFilasExcel(xlHoja1, "2.1  Encaje exigible", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnEncajeExigALMN = nSaldoDiario1 'vapa 20171120
    lnEncajeExigALME = nSaldoDiario2 'vapa 20171120

    lnEncajeExigALMN = Round((lnEncajeExigALMN / lnTotalaMN * 100), 2) 'vapa 20171120
    lnEncajeExigALME = Round((lnEncajeExigALME / lnTotalaME * 100), 2) 'vapa 20171120

    InsertaLiquidezAlertaTemprana ldFecha, lnRatioLiquidezMN, lnRatioLiquidezME, lnRatioLAjusRecursosPrestadosMN, lnRatioLAjusRecursosPrestadosME, lnRatioInversionesLiquidasMN, lnEncajeExigALMN, lnEncajeExigALME 'VAPA 20171003

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "2500")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "2500")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2.2  Fondos de encaje (24)", "", "sum(F" & (lnFila + 2) & ":F" & (lnFila + 3) & ")", "+ G" & (lnFila + 1) & "+ G" & (lnFila + 3) & "", lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "2600")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "2600")
    Call PintaFilasExcel(xlHoja1, "- Caja del día", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "2650")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "2650")
    Call PintaFilasExcel(xlHoja1, "- Caja promedio diario del período de encaje anterior", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "2700")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "2700")
    Call PintaFilasExcel(xlHoja1, "- Cuenta corriente BCRP", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    nSaldoDiario1 = Val(xlHoja1.Cells(lnFila - 4, 6)) - Val(xlHoja1.Cells(lnFila - 5, 6))
    nSaldoDiario2 = Val(xlHoja1.Cells(lnFila - 4, 7)) - Val(xlHoja1.Cells(lnFila - 5, 7))
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2.3  Resultados del día (fondos de encaje -  encaje exigible)", "", "F" & (lnFila - 4) & "-" & "F" & (lnFila - 5), "G" & (lnFila - 4) & "-" & "G" & (lnFila - 5), lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "2900")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "2900")
    Call PintaFilasExcel(xlHoja1, "2.4  Posición de encaje acumulada del período a la fecha", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "2950")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "2950")
    Call PintaFilasExcel(xlHoja1, "2.5  Posición acumulada del requerimiento mínimo en cuenta corriente BCRP a la fecha", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    xlHoja1.Range(xlHoja1.Cells(lnPosInicial + 1, 2), xlHoja1.Cells(lnFila, 7)).Interior.ColorIndex = 2
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 6), xlHoja1.Cells(lnFila, 7)).Font.Size = 9
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 6), xlHoja1.Cells(lnFila, 7)).Font.Name = "Calibri"
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 6), xlHoja1.Cells(lnFila, 7)).Font.Bold = True
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 7, CCur(lnFila)

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "3. Cheques a deducir del total de obligaciones sujetas a encaje", "", "sum(F" & (lnFila + 1) & ":G" & (lnFila + 5) & ")", "sum(F" & (lnFila + 1) & ":G" & (lnFila + 5) & ")", lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = 0
    nSaldoDiario2 = 0
    Call PintaFilasExcel(xlHoja1, "3.1  A deducir de obligaciones a la vista y a plazo hasta 30 días", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = 0
    nSaldoDiario2 = 0
    Call PintaFilasExcel(xlHoja1, "3.2  A deducir de obligaciones a plazo mayor de 30 días", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "3300")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "3300")
    Call PintaFilasExcel(xlHoja1, "3.3  A deducir de ahorro", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = 0
    nSaldoDiario2 = 0
    Call PintaFilasExcel(xlHoja1, "3.4  A deducir de obligaciones en moneda nacional con rendimiento vinculado al tipo de cambio en moneda extranjera o a operaciones swap y similares", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 7)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 7, CCur(lnFila)

    lnFila = lnFila + 1
    lnPosInicial = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = 0
    nSaldoDiario2 = 0
    Call PintaFilasExcel(xlHoja1, "4. Obligaciones con entidades financieras del exterior (25)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    xlHoja1.Range(xlHoja1.Cells(lnPosInicial, 2), xlHoja1.Cells(lnFila, 7)).Interior.ColorIndex = 2
    ExcelCuadro xlHoja1, 2, CCur(lnPosInicial), 7, CCur(lnFila)

    '******DEPOSITOS A GRANDES ACREEDORES********
    lnFila = lnFila + 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Dim d As ColorConstants
    Call PintaFilasExcel(xlHoja1, "IV.  SALDO DE DEPÓSITOS DE GRANDES ACREEDORES", "", "Moneda Nacional", "Moneda Extranjera", lnFila, False, True, False)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).HorizontalAlignment = xlCenter

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "3322")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "3322")
    Call PintaFilasExcel(xlHoja1, "1.     Estado(26)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "3324")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "3324")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2.     AFPs (27)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "3326")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "3326")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "3.     Fondos mutuos y fondos de inversión.", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "3328")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "3328")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "4.     Empresas del sistema de seguros (28)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "3329")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "3329")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "5.     Sociedad agente de bolsa (SAB)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "3330")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "3330")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "6.     Otros depositantes (29)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "", "", "", "", lnFila, False, True, False)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 2

    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "V.  SALDO DE DEPÓSITOS DE EMPRESAS DEL SISTEMA FINANCIERO", "", "Moneda Nacional", "Moneda Extranjera", lnFila, False, True, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "3335")
    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "3335")
    Call PintaFilasExcel(xlHoja1, "1. Sistema financiero nacional", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = 0
    nSaldoDiario2 = 0
    Call PintaFilasExcel(xlHoja1, "2. Sistema financiero del exterior", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "VI. NUMERO DE DÍAS DE REDESCUENTO EN LOS ÚLTIMOS 180 DÍAS (30)", "", "", "", lnFila, False, True, False)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 6, CCur(lnFila)

    oBarra.Progress 9, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "VII. PÉRDIDA NETA DE DERIVADOS PARA NEGOCIACIÓN (31)", "", "", "", lnFila, False, True, False)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 6, CCur(lnFila)

    lnFila = lnFila + 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "VIII.    POSICIONES DE CAMBIO", "", "Global (33)", "", lnFila, False, True, False)
    xlHoja1.Cells(lnFila, 4) = "Balance (32)"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 6)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 6, CCur(lnFila)

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True

    nSaldoDiario2 = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "3400")
    nSaldoDiario1 = nSaldoDiario2

    xlHoja1.Cells(lnFila, 2) = "Moneda Extranjera"
    xlHoja1.Cells(lnFila, 4) = nSaldoDiario1
    xlHoja1.Cells(lnFila, 6) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 6)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 6)).Font.Bold = False
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 6, CCur(lnFila)

    lnFila = lnFila + 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 5)).Font.Bold = False
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Cells(lnFila, 2) = ""
    xlHoja1.Cells(lnFila, 4) = "En Moneda Extranjera (USD)"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 5, CCur(lnFila)

    lnFila = lnFila + 1

    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 5)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True
    'xlHoja1.Cells(lnFila, 2) = "IX.    Posición Neta en Productos Financieros Derivados (Moneda Extranjera / PEN)" 'Comentado by NAGL 20180419
    xlHoja1.Cells(lnFila, 2) = "IX.    Posición Contable Neta en Productos Financieros Derivados (Moneda Extranjera / PEN)" 'NAGL 20180419 RFC1804130002
    xlHoja1.Cells(lnFila, 4) = ""
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).HorizontalAlignment = xlHAlignJustify
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlVAlignTop
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 5, CCur(lnFila)
    
    Set oEst = Nothing
    oBarra.Progress 10, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "Generación Terminada", "", vbBlue
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
    If lsArchivo <> "" Then
        CargaArchivo lsArchivo, App.path & "\SPOOLER\"
    End If
    Exit Sub
GeneraExcelErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    If lbExcel = True Then
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Sub '****************************************FIN NAGL 20170904******************************************'

Private Sub PintaFilasExcel(ByRef pxlHoja1 As Excel.Worksheet, ByVal psCodigo As String, ByVal psDescripcion As String, ByVal pnValor1 As String, ByVal pnValor2 As String, pnFilaX As Integer, ByVal pbFormulas As Boolean, ByVal pbBold As Boolean, ByVal pbNumberFormat As Boolean)
    pxlHoja1.Cells(pnFilaX, 2) = psCodigo
    pxlHoja1.Cells(pnFilaX, 3) = psDescripcion
    If pbFormulas Then
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 6), pxlHoja1.Cells(pnFilaX, 6)).Formula = "=" & pnValor1
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 7), pxlHoja1.Cells(pnFilaX, 7)).Formula = "=" & pnValor2
    Else
        xlHoja1.Cells(pnFilaX, 6) = pnValor1
        xlHoja1.Cells(pnFilaX, 7) = pnValor2
    End If
    If pbNumberFormat Then
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 6), pxlHoja1.Cells(pnFilaX, 6)).NumberFormat = "#,##0.00;-#,##0.00"
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 7), pxlHoja1.Cells(pnFilaX, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    End If
    pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 1), pxlHoja1.Cells(pnFilaX, 8)).Font.Bold = pbBold
End Sub

'VAPA20170926
Private Sub PintaFilasExcelParaInversion(ByRef pxlHoja1 As Excel.Worksheet, ByVal psCodigo As String, ByVal psDescripcion As String, ByVal pnValor1 As String, ByVal pnValor2 As String, pnFilaX As Integer, ByVal pbFormulas As Boolean, ByVal pbBold As Boolean, ByVal pbNumberFormat As Boolean)
    pxlHoja1.Cells(pnFilaX, 2) = psCodigo
    pxlHoja1.Cells(pnFilaX, 3) = psDescripcion
    If pbFormulas Then
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 6), pxlHoja1.Cells(pnFilaX, 6)).Formula = "=" & pnValor1
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 7), pxlHoja1.Cells(pnFilaX, 7)).Formula = "=" & pnValor2
    Else
        xlHoja1.Cells(pnFilaX, 6) = pnValor1
        xlHoja1.Cells(pnFilaX, 7) = pnValor2
    End If
    If pbNumberFormat Then
        'lnRatioInversionesLiquidasMN = pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 6), pxlHoja1.Cells(pnFilaX, 6)).NumberFormat = "#,##0.00;-#,##0.00"
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 6), pxlHoja1.Cells(pnFilaX, 6)).NumberFormat = "#,##0.00;-#,##0.00"
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 7), pxlHoja1.Cells(pnFilaX, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    End If
    pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 1), pxlHoja1.Cells(pnFilaX, 8)).Font.Bold = pbBold
End Sub
'VAPA END

Private Sub PintaFilasExcel2(ByRef pxlHoja1 As Excel.Worksheet, ByVal psDescripcion As String, ByVal pnValor1 As String, ByVal pnValor2 As String, ByVal pnValor3 As String, ByVal pnValor4 As String, ByVal pnValor5 As String, ByVal pnValor6 As String, pnFilaX As Integer, ByVal pbFormulas As Boolean, ByVal pbBold As Boolean, ByVal pbNumberFormat As Boolean)
    'pxlHoja1.Cells(pnFilaX, 2) = psCodigo
    pxlHoja1.Cells(pnFilaX, 2) = psDescripcion
    If pbFormulas Then
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 3), pxlHoja1.Cells(pnFilaX, 3)).Formula = "=" & pnValor1
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 4), pxlHoja1.Cells(pnFilaX, 4)).Formula = "=" & pnValor2
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 5), pxlHoja1.Cells(pnFilaX, 5)).Formula = "=" & pnValor3
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 6), pxlHoja1.Cells(pnFilaX, 6)).Formula = "=" & pnValor4
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 7), pxlHoja1.Cells(pnFilaX, 7)).Formula = "=" & pnValor5
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 8), pxlHoja1.Cells(pnFilaX, 8)).Formula = "=" & pnValor6
    Else
        xlHoja1.Cells(pnFilaX, 3) = pnValor1
        xlHoja1.Cells(pnFilaX, 4) = pnValor2
        xlHoja1.Cells(pnFilaX, 5) = pnValor3
        xlHoja1.Cells(pnFilaX, 6) = pnValor4
        xlHoja1.Cells(pnFilaX, 7) = pnValor5
        xlHoja1.Cells(pnFilaX, 8) = pnValor6
    End If
    If pbNumberFormat Then
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 3), pxlHoja1.Cells(pnFilaX, 3)).NumberFormat = "#,##0.00;-#,##0.00"
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 4), pxlHoja1.Cells(pnFilaX, 4)).NumberFormat = "#,##0.00;-#,##0.00"
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 5), pxlHoja1.Cells(pnFilaX, 5)).NumberFormat = "#,##0.00;-#,##0.00"
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 6), pxlHoja1.Cells(pnFilaX, 6)).NumberFormat = "#,##0.00;-#,##0.00"
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 7), pxlHoja1.Cells(pnFilaX, 7)).NumberFormat = "#,##0.00;-#,##0.00"
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 8), pxlHoja1.Cells(pnFilaX, 8)).NumberFormat = "#,##0.00;-#,##0.00"
    End If
    pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 1), pxlHoja1.Cells(pnFilaX, 8)).Font.Bold = pbBold
End Sub '********NAGL ERS 079-2016 20170407

Private Sub PintaFilasExcel3(ByRef pxlHoja1 As Excel.Worksheet, ByVal psDescripcion As String, ByVal pnValor1 As String, ByVal pnValor2 As String, ByVal pnValor3 As String, ByVal pnValor4 As String, ByVal pnValor5 As String, ByVal pnValor6 As String, ByVal pnValor7 As String, pnFilaX As Integer, ByVal pbFormulas As Boolean, ByVal pbBold As Boolean, ByVal pbNumberFormat As Boolean)
    pxlHoja1.Cells(pnFilaX, 2) = psDescripcion
    If pbFormulas Then
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 3), pxlHoja1.Cells(pnFilaX, 3)).Formula = "=" & pnValor1
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 4), pxlHoja1.Cells(pnFilaX, 4)).Formula = "=" & pnValor2
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 5), pxlHoja1.Cells(pnFilaX, 5)).Formula = "=" & pnValor3
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 6), pxlHoja1.Cells(pnFilaX, 6)).Formula = "=" & pnValor4
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 7), pxlHoja1.Cells(pnFilaX, 7)).Formula = "=" & pnValor5
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 8), pxlHoja1.Cells(pnFilaX, 8)).Formula = "=" & pnValor6
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 9), pxlHoja1.Cells(pnFilaX, 9)).Formula = "=" & pnValor7
    Else
        xlHoja1.Cells(pnFilaX, 3) = pnValor1
        xlHoja1.Cells(pnFilaX, 4) = pnValor2
        xlHoja1.Cells(pnFilaX, 5) = pnValor3
        xlHoja1.Cells(pnFilaX, 6) = pnValor4
        xlHoja1.Cells(pnFilaX, 7) = pnValor5
        xlHoja1.Cells(pnFilaX, 8) = pnValor6
        xlHoja1.Cells(pnFilaX, 9) = pnValor7
    End If
    If pbNumberFormat Then
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 3), pxlHoja1.Cells(pnFilaX, 3)).NumberFormat = "#,##0.00;-#,##0.00"
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 4), pxlHoja1.Cells(pnFilaX, 4)).NumberFormat = "#,##0.00;-#,##0.00"
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 5), pxlHoja1.Cells(pnFilaX, 5)).NumberFormat = "#,##0.00;-#,##0.00"
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 6), pxlHoja1.Cells(pnFilaX, 6)).NumberFormat = "#,##0.00;-#,##0.00"
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 7), pxlHoja1.Cells(pnFilaX, 7)).NumberFormat = "#,##0.00;-#,##0.00"
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 8), pxlHoja1.Cells(pnFilaX, 8)).NumberFormat = "#,##0.00;-#,##0.00"
        pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 9), pxlHoja1.Cells(pnFilaX, 9)).NumberFormat = "#,##0.00;-#,##0.00"
    End If
    pxlHoja1.Range(pxlHoja1.Cells(pnFilaX, 1), pxlHoja1.Cells(pnFilaX, 9)).Font.Bold = pbBold
End Sub '********NAGL 20190615

Private Function SaldoCajasObligExoneradas(ByVal pdFecha As String, ByVal pnMoneda As Moneda) As Currency
    Dim oNCaja As New NCajaCtaIF
    Dim oEnc As New nEncajeBCR
    Dim rsEncDiario As New ADODB.Recordset
    
    Set rsEncDiario = oEnc.ObtenerParamEncajeDiarioxCod("04", CStr(pnMoneda), Format(pdFecha, "yyyymmdd")) 'NAGL Agregó pnMoneda,pdFecha Según TIC1810110004 20181015
    SaldoCajasObligExoneradas = ObtieneEfectivoDiaxRptConcentracionGastos(pdFecha, pnMoneda) + IIf(pnMoneda = gMonedaNacional, rsEncDiario!nValor, 0)
    
    Set rsEncDiario = Nothing
    Set oEnc = Nothing
    Set oNCaja = Nothing
End Function

Public Function ObtenerCtaContSaldoBalanceDiario(psCtaContCod As String, pdFecDiaProceso As Date, psMoneda As String, psMesBalanceDiario As String, psAnioBalanceDiario As String) As Currency
    On Error GoTo ObtenerCtaContSaldoBalanceDiarioErr
    Dim oRS As ADODB.Recordset
    Dim oConec As DConecta
    Dim nSaldo As Currency
    Dim psSql As String
   Set oRS = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
    psSql = "exec stp_sel_ObtenerDatosBalanceDiarioReal '" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "', '" & psMesBalanceDiario & "', '" & psAnioBalanceDiario & "'"
   Set oRS = oConec.CargaRecordSet(psSql)
   If Not oRS.BOF And Not oRS.EOF Then
        Do While Not oRS.EOF
            nSaldo = oRS!nSaldoFinImporte
            oRS.MoveNext
        Loop
    Else
        nSaldo = 0
    End If
    oConec.CierraConexion
    ObtenerCtaContSaldoBalanceDiario = nSaldo
    Exit Function
ObtenerCtaContSaldoBalanceDiarioErr:
    Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceDiario Method")
End Function 'NAGL 20170425

'VAPA 20171109
Public Sub InsertaLiquidezAlertaTemprana(ByVal pdFecha As Date, ByVal pnRatioLiquidezMN As Double, ByVal pnRatioLiquidezME As Double, ByVal pnRatioLAjusRecursosPrestadosMN As Double, ByVal pnRatioLAjusRecursosPrestadosME As Double, ByVal pnRatioInversionesLiquidasMN As Double, pnEncajeMN As Double, pnEncajeME As Double)
    On Error GoTo InsertaLiquidezAlertaTempranaErr
    Dim oConec As DConecta
    Dim psSql As String
   Set oConec = New DConecta
   oConec.AbreConexion
    psSql = "exec stp_ins_ReporteLiquidezAlertaTemprana '" & Format(pdFecha, "YYYY/MM/DD") & "'," & pnRatioLiquidezMN & "," & pnRatioLiquidezME & "," & pnRatioLAjusRecursosPrestadosMN & "," & pnRatioLAjusRecursosPrestadosME & "," & pnRatioInversionesLiquidasMN & "," & pnEncajeMN & "," & pnEncajeME & ""
    oConec.Ejecutar (psSql)
    oConec.CierraConexion
    Exit Sub
InsertaLiquidezAlertaTempranaErr:
    Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceDiario Method")
End Sub
'VAPA END

Public Function GetCuentasIntDevengados() As ADODB.Recordset
    Dim psSql As String
    Dim oConec As New DConecta
    psSql = "Exec stp_sel_OtieneCuentasIntDevengados "
    oConec.AbreConexion
Set GetCuentasIntDevengados = oConec.CargaRecordSet(psSql)
oConec.CierraConexion
Set oConec = Nothing
End Function '****NAGL ERS079-2017 20180130

Public Function ObtieneCtaSaldoDiarioAnx15A(psCtaContCodMas As String, psCtaContCodMenos As String, pdFecha As Date, pdFechaBalanceDiario, pnTipoCambBalance As Currency, pnTipoCambioMes As Currency, Optional psTipo As String = "") As Currency
    Dim psSql As String
    Dim rs As New ADODB.Recordset
    Dim oConec As New DConecta
    Dim nSaldo As Currency
    oConec.AbreConexion
    psSql = "Exec stp_sel_ObtieneCtaSaldoDiarioAnx15A '" & psCtaContCodMas & "','" & psCtaContCodMenos & "','" & Format(pdFecha, "yyyymmdd") & "','" & Format(pdFechaBalanceDiario, "yyyymmdd") & "'," & pnTipoCambBalance & "," & pnTipoCambioMes & ", '" & psTipo & "'"
    Set rs = oConec.CargaRecordSet(psSql)
    If Not rs.BOF And Not rs.EOF Then
        nSaldo = IIf(IsNull(rs!nSaldo), 0, rs!nSaldo)
    Else
        nSaldo = 0
    End If
    oConec.CierraConexion
    ObtieneCtaSaldoDiarioAnx15A = nSaldo
Set oConec = Nothing
End Function '****NAGL Según Anexo02 ERS006-2019 20190615

Public Function ObtieneCtaSaldoDiarioAnx15A_Det(psCtaContCodMas As String, psCtaContCodMenos As String, pdFecha As Date, pdFechaBalanceDiario, pnTipoCambBalance As Currency, pnTipoCambioMes As Currency, Optional psTipo As String = "") As ADODB.Recordset
    Dim psSql As String
    Dim oConec As New DConecta
    psSql = "Exec stp_sel_ObtieneCtaSaldoDiarioAnx15A '" & psCtaContCodMas & "','" & psCtaContCodMenos & "','" & Format(pdFecha, "yyyymmdd") & "','" & Format(pdFechaBalanceDiario, "yyyymmdd") & "'," & pnTipoCambBalance & "," & pnTipoCambioMes & ", '" & psTipo & "'"
    oConec.AbreConexion
Set ObtieneCtaSaldoDiarioAnx15A_Det = oConec.CargaRecordSet(psSql)
oConec.CierraConexion
Set oConec = Nothing
End Function '****NAGL Según Anexo02 ERS006-2019 20190615

Private Sub CargaValidacionCtaContIntDeveng15A(ByVal xlAplicacion As Excel.Application, ByVal pdFecha As Date, ByVal pdFechaBalanceDiario As Date, ByVal nTipoCambioBalance As Currency, ByVal nTipoCambioAn As Currency, ByVal psMesBalanceDiario As String, ByVal psAnioBalanceDiario As String)
    Dim pdFechaBalReal As Date
    Dim DAnxRies As New DAnexoRiesgos
    Dim liFila As Integer, liFilaIni As Integer, Cant As Integer
    Dim cCtaCont As String
    Dim rsCta As New ADODB.Recordset
    Dim rs As New ADODB.Recordset

    pdFechaBalReal = DateAdd("D", -Day(pdFechaBalanceDiario), pdFechaBalanceDiario)
    xlHoja1.Cells(7, 9) = "VALIDACION DE INTERESES DEVENGADOS"
    xlHoja1.Range(xlHoja1.Cells(7, 9), xlHoja1.Cells(7, 11)).Merge True
xlHoja1.Range(xlHoja1.Cells(7, 9), xlHoja1.Cells(7, 11)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 9, 7, 11, 7

xlHoja1.Cells(8, 9) = "Cuenta Contable"
    xlHoja1.Cells(8, 10) = "Saldos Estadísticos"
    xlHoja1.Cells(8, 11) = "Balance al " & Format(pdFechaBalReal, "dd/mm/yyyy")
    ExcelCuadro xlHoja1, 9, 8, 11, 8

'******Agregado by NAGL 20190506 ERS006-2019**************
Set rs = DAnxRies.CargaSaldosEstadistAnx15Ay15B(pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn)
liFila = 9
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            xlHoja1.Cells(liFila, 9) = rs!cCtaContCod
            xlHoja1.Range(xlHoja1.Cells(liFila, 9), xlHoja1.Cells(liFila, 9)).Font.Bold = True
            xlHoja1.Cells(liFila, 10) = Format(rs!nSaldoImporte, "#,##0.00")
            xlHoja1.Cells(liFila, 11) = Round(ObtenerCtaContSaldoBalanceDiario(rs!cCtaContCod, pdFecha, Mid(rs!cCtaContCod, 3, 1), psMesBalanceDiario, psAnioBalanceDiario) / IIf(Mid(rs!cCtaContCod, 3, 1) = "1", 1, nTipoCambioBalance), 2)
            xlHoja1.Range(xlHoja1.Cells(liFila, 9), xlHoja1.Cells(liFila, 11)).Font.Name = "Arial Narrow"
            xlHoja1.Range(xlHoja1.Cells(liFila, 9), xlHoja1.Cells(liFila, 11)).Font.Size = 10
            xlHoja1.Range(xlHoja1.Cells(liFila, 9), xlHoja1.Cells(liFila, 11)).HorizontalAlignment = xlCenter
            ExcelCuadro xlHoja1, 9, liFila, 11, CCur(liFila)
    liFila = liFila + 1
            rs.MoveNext
        Loop
    End If 'Agregado by NAGL 20190416 ERS006-2019
    xlHoja1.Range(xlHoja1.Cells(9, 10), xlHoja1.Cells(liFila - 1, 11)).NumberFormat = "#,##0.00;-#,##0.00"

    xlHoja1.Range(xlHoja1.Cells(7, 9), xlHoja1.Cells(8, 11)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(8, 9), xlHoja1.Cells(26, 9)).Font.Bold = True

    xlHoja1.Range(xlHoja1.Cells(7, 9), xlHoja1.Cells(8, 11)).Interior.ColorIndex = 15

    xlHoja1.Range(xlHoja1.Cells(8, 9), xlHoja1.Cells(liFila - 1, 9)).EntireColumn.AutoFit
    xlHoja1.Range(xlHoja1.Cells(8, 10), xlHoja1.Cells(liFila - 1, 10)).EntireColumn.AutoFit
    xlHoja1.Range(xlHoja1.Cells(8, 11), xlHoja1.Cells(liFila - 1, 11)).EntireColumn.AutoFit
    '**************END NAGL ERS006-2019 ********************

    'ObtieneReporteConsolEstadisticoBCRPSFNADEUD
    liFila = 9
    liFilaIni = 8
    Cant = 0
Set rsCta = GetCuentasIntDevengados
Do While Not rsCta.EOF
        cCtaCont = rsCta!cCtaContCod
     Set rs = DAnxRies.ObtieneReporteConsolEstadisticoBCRPSFNADEUD(pdFecha, cCtaCont)
        If rs.RecordCount <> 0 Then
            liFila = liFila + Cant
            liFilaIni = liFila
        End If
        If cCtaCont = "11_802" Or cCtaCont = "11_803" Then
            If Not rs.BOF And Not rs.BOF Then
                Do While Not rs.EOF
                    If liFilaIni = liFila Then
                        xlHoja1.Cells(liFila - 2, 13) = "INVERSIONES DE CAJA MAYNAS EN BANCOS Y OTRAS EMPRESAS DEL SISTEMA FINANCIERO " & Mid(cCtaCont, 1, 2) + "0" + Mid(cCtaCont, 4, 6)
                        xlHoja1.Range(xlHoja1.Cells(liFila - 2, 13), xlHoja1.Cells(liFila - 2, 23)).Merge True
                   xlHoja1.Range(xlHoja1.Cells(liFila - 2, 13), xlHoja1.Cells(liFila - 2, 23)).HorizontalAlignment = xlCenter
                        ExcelCuadro xlHoja1, 13, liFila - 2, 23, CCur(liFila - 2)

                   xlHoja1.Cells(liFila - 1, 13) = "CtaContCod"
                        xlHoja1.Cells(liFila - 1, 14) = "Moneda"
                        xlHoja1.Cells(liFila - 1, 15) = "Institución Financiera"
                        xlHoja1.Cells(liFila - 1, 16) = "Depósito"
                        xlHoja1.Cells(liFila - 1, 17) = "TEA"
                        xlHoja1.Cells(liFila - 1, 18) = "Capital"
                        xlHoja1.Cells(liFila - 1, 19) = "Int.Deveng Acum"
                        xlHoja1.Cells(liFila - 1, 20) = "Apertura"
                        xlHoja1.Cells(liFila - 1, 21) = "Vencimiento"
                        xlHoja1.Cells(liFila - 1, 22) = "Interés Pactado"
                        xlHoja1.Cells(liFila - 1, 23) = "Estado"
                        ExcelCuadro xlHoja1, 13, liFila - 1, 23, CCur(liFila - 1)

                   xlHoja1.Range(xlHoja1.Cells(liFila - 2, 13), xlHoja1.Cells(liFila - 1, 23)).Interior.ColorIndex = 15
                        xlHoja1.Range(xlHoja1.Cells(liFila - 2, 13), xlHoja1.Cells(liFila - 1, 23)).Font.Bold = True
                        xlHoja1.Range(xlHoja1.Cells(liFila - 2, 13), xlHoja1.Cells(liFila - 1, 23)).HorizontalAlignment = xlCenter
                        xlHoja1.Range(xlHoja1.Cells(liFila - 1, 19), xlHoja1.Cells(liFila - 1, 19)).Interior.ColorIndex = 44

                    End If

                    xlHoja1.Cells(liFila, 13) = rs!cCtaContCod
                    xlHoja1.Range(xlHoja1.Cells(liFila, 13), xlHoja1.Cells(liFila, 13)).Font.Bold = True
                    xlHoja1.Range(xlHoja1.Cells(liFila, 13), xlHoja1.Cells(liFila, 13)).HorizontalAlignment = xlCenter
                    xlHoja1.Cells(liFila, 14) = rs!cMoneda
                    xlHoja1.Range(xlHoja1.Cells(liFila, 14), xlHoja1.Cells(liFila, 14)).HorizontalAlignment = xlCenter
                    xlHoja1.Cells(liFila, 15) = rs!cInstFinanc
                    xlHoja1.Range(xlHoja1.Cells(liFila, 15), xlHoja1.Cells(liFila, 15)).HorizontalAlignment = xlLeft
                    xlHoja1.Cells(liFila, 16) = rs!cDeposito
                    xlHoja1.Cells(liFila, 17) = rs!intvalor / 100
                    xlHoja1.Cells(liFila, 17).NumberFormat = "0.00%"
                    xlHoja1.Cells(liFila, 18) = Format(rs!nCapital, "#,##0.00")
                    xlHoja1.Cells(liFila, 19) = Format(rs!IntDevengadoAcum, "#,##0.00")
                    xlHoja1.Cells(liFila, 20) = Format(rs!dCtaIFAper, "mm/dd/yyyy")
                    xlHoja1.Cells(liFila, 21) = Format(rs!dCtaIFVenc, "mm/dd/yyyy")
                    xlHoja1.Cells(liFila, 22) = Format(rs!IntPactado, "#,##0.00")
                    xlHoja1.Cells(liFila, 23) = rs!EstadoActual

                    xlHoja1.Range(xlHoja1.Cells(liFila, 16), xlHoja1.Cells(liFila, 23)).HorizontalAlignment = xlCenter
                    ExcelCuadro xlHoja1, 13, liFila, 23, CCur(liFila)
                   liFila = liFila + 1
                    rs.MoveNext
                Loop
                Cant = 5
            End If
        Else
            If Not rs.BOF And Not rs.EOF Then
                Do While Not rs.EOF
                    If liFilaIni = liFila Then
                        xlHoja1.Cells(liFila - 2, 13) = "ADEUDADOS Y OBLIGACIONES FINANCIERAS (2408)"
                        xlHoja1.Range(xlHoja1.Cells(liFila - 2, 13), xlHoja1.Cells(liFila - 2, 23)).Merge True
                   xlHoja1.Range(xlHoja1.Cells(liFila - 2, 13), xlHoja1.Cells(liFila - 2, 23)).HorizontalAlignment = xlCenter
                        ExcelCuadro xlHoja1, 13, liFila - 2, 23, CCur(liFila - 2)

                   xlHoja1.Cells(liFila - 1, 13) = "CtaContCod"
                        xlHoja1.Cells(liFila - 1, 14) = "Moneda"
                        xlHoja1.Cells(liFila - 1, 15) = "Entidad Acreedora"
                        xlHoja1.Cells(liFila - 1, 16) = "Línea De Crédito"
                        xlHoja1.Cells(liFila - 1, 17) = "TEA"
                        xlHoja1.Cells(liFila - 1, 18) = "Capital"
                        xlHoja1.Cells(liFila - 1, 19) = "Int.Deveng Acum"
                        xlHoja1.Cells(liFila - 1, 20) = "Último Pago"
                        xlHoja1.Cells(liFila - 1, 21) = "Apertura"
                        xlHoja1.Cells(liFila - 1, 22) = "Vencimiento"
                        xlHoja1.Cells(liFila - 1, 23) = "Días a Calcu."
                        ExcelCuadro xlHoja1, 13, liFila - 1, 23, CCur(liFila - 1)

                   xlHoja1.Range(xlHoja1.Cells(liFila - 2, 13), xlHoja1.Cells(liFila - 1, 23)).Interior.ColorIndex = 10
                        xlHoja1.Range(xlHoja1.Cells(liFila - 2, 13), xlHoja1.Cells(liFila - 1, 23)).Font.Bold = True
                        xlHoja1.Range(xlHoja1.Cells(liFila - 2, 13), xlHoja1.Cells(liFila - 1, 23)).HorizontalAlignment = xlCenter
                        xlHoja1.Range(xlHoja1.Cells(liFila - 1, 19), xlHoja1.Cells(liFila - 1, 19)).Interior.ColorIndex = 44
                        xlHoja1.Range(xlHoja1.Cells(liFila - 2, 13), xlHoja1.Cells(liFila - 1, 23)).Font.Color = vbWhite
                        xlHoja1.Range(xlHoja1.Cells(liFila - 2, 19), xlHoja1.Cells(liFila - 1, 19)).Font.Color = vbBlack

                    End If

                    xlHoja1.Cells(liFila, 13) = rs!cCtaContCod
                    xlHoja1.Range(xlHoja1.Cells(liFila, 13), xlHoja1.Cells(liFila, 13)).Font.Bold = True
                    xlHoja1.Range(xlHoja1.Cells(liFila, 13), xlHoja1.Cells(liFila, 13)).HorizontalAlignment = xlCenter
                    xlHoja1.Cells(liFila, 14) = rs!cMoneda
                    xlHoja1.Range(xlHoja1.Cells(liFila, 14), xlHoja1.Cells(liFila, 14)).HorizontalAlignment = xlCenter
                    xlHoja1.Cells(liFila, 15) = rs!cPersNombre
                    xlHoja1.Range(xlHoja1.Cells(liFila, 15), xlHoja1.Cells(liFila, 15)).HorizontalAlignment = xlLeft
                    xlHoja1.Cells(liFila, 16) = rs!cDesLinCred
                    xlHoja1.Cells(liFila, 17) = rs!nCtaIFIntValor / 100
                    xlHoja1.Cells(liFila, 17).NumberFormat = "0.00%"
                    xlHoja1.Cells(liFila, 18) = Format(rs!nSaldoCap, "#,##0.00")
                    xlHoja1.Cells(liFila, 19) = Format(rs!InteresDevengAcumulado, "#,##0.00")
                    xlHoja1.Cells(liFila, 20) = Format(rs!dCuotaUltPago, "mm/dd/yyyy")
                    xlHoja1.Cells(liFila, 21) = Format(rs!dCtaIFAper, "mm/dd/yyyy")
                    xlHoja1.Cells(liFila, 22) = Format(rs!dVencimiento, "mm/dd/yyyy")
                    xlHoja1.Cells(liFila, 23) = Format(rs!nDiasUltPAgo, "#,##0.00")

                    xlHoja1.Range(xlHoja1.Cells(liFila, 16), xlHoja1.Cells(liFila, 23)).HorizontalAlignment = xlCenter
                    ExcelCuadro xlHoja1, 13, liFila, 23, CCur(liFila)
                   liFila = liFila + 1
                    rs.MoveNext
                Loop
                Cant = 5
            End If
        End If
        If rs.RecordCount <> 0 Then
            liFila = liFila - 1

            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 23)).Font.Name = "Arial Narrow"
            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 23)).Font.Size = 10

            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 13)).EntireColumn.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 14), xlHoja1.Cells(liFila, 14)).EntireColumn.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 15), xlHoja1.Cells(liFila, 15)).EntireColumn.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 16), xlHoja1.Cells(liFila, 16)).EntireColumn.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 17), xlHoja1.Cells(liFila, 17)).EntireColumn.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 18), xlHoja1.Cells(liFila, 18)).EntireColumn.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 19), xlHoja1.Cells(liFila, 19)).EntireColumn.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 20), xlHoja1.Cells(liFila, 20)).EntireColumn.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 21), xlHoja1.Cells(liFila, 21)).EntireColumn.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 22), xlHoja1.Cells(liFila, 22)).EntireColumn.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 23), xlHoja1.Cells(liFila, 23)).EntireColumn.AutoFit
        End If
        rsCta.MoveNext
    Loop
    xlHoja1.Range(xlHoja1.Cells(10, 10), xlHoja1.Cells(10, 10)).Font.Bold = False


'*********Sección Obligaciones Ahorro y Plazo Fijo Según Anx02_ERS006-2019*************'
'***Obligaciones por Cuentas de Ahorro
Set rs = Nothing
liFila = liFila + 3
    xlHoja1.Cells(liFila, 13) = "OBLIGACIONES POR CUENTAS DE AHORRO"
    xlHoja1.Range(xlHoja1.Cells(liFila, 13), xlHoja1.Cells(liFila, 15)).Merge True
ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
liFila = liFila + 1
    xlHoja1.Cells(liFila, 13) = "Cuenta"
    xlHoja1.Cells(liFila, 14) = "Saldo"
    xlHoja1.Cells(liFila, 15) = "Descripción"
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).Interior.ColorIndex = 9
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).Font.Color = vbWhite
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
liFila = liFila + 1
    liFilaIni = liFila

'MN
Set rs = DAnxRies.ObtieneOtrasObligaciones15A(pdFecha, "1", "232", "Det")
If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            xlHoja1.Cells(liFila, 13) = rs!cCtaCnt
            xlHoja1.Cells(liFila, 14) = Format(rs!nSaldo, "#,##0.00")
            xlHoja1.Cells(liFila, 15) = rs!cDescrip
            ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
        liFila = liFila + 1
            rs.MoveNext
        Loop
    End If
    xlHoja1.Cells(liFila, 13) = "211701"
    xlHoja1.Cells(liFila, 14) = Format(DAnxRies.ObtenerSaldEstadistAnx15Ay15B("211701", "1", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn), "#,##0.00") * -1
    xlHoja1.Cells(liFila, 15) = "Dep.Inmovilizados a Excluir (-)"
    ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
liFila = liFila + 1
Set rs = Nothing
Set rs = ObtieneCtaSaldoDiarioAnx15A_Det("211802,231802", "", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn, "Det")
If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            xlHoja1.Cells(liFila, 13) = rs!cCtaCnt
            xlHoja1.Cells(liFila, 14) = Format(rs!nSaldo, "#,##0.00")
            xlHoja1.Cells(liFila, 15) = rs!cDescrip
            ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
        liFila = liFila + 1
            rs.MoveNext
        Loop
    End If
'ME
Set rs = DAnxRies.ObtieneOtrasObligaciones15A(pdFecha, "2", "232", "Det")
If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            xlHoja1.Cells(liFila, 13) = rs!cCtaCnt
            xlHoja1.Cells(liFila, 14) = Format(rs!nSaldo, "#,##0.00")
            xlHoja1.Cells(liFila, 15) = rs!cDescrip
            ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
        liFila = liFila + 1
            rs.MoveNext
        Loop
    End If
    xlHoja1.Cells(liFila, 13) = "212701"
    xlHoja1.Cells(liFila, 14) = Format(DAnxRies.ObtenerSaldEstadistAnx15Ay15B("212701", "1", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn), "#,##0.00") * -1
    xlHoja1.Cells(liFila, 15) = "Dep.Inmovilizados a Excluir (-)"
    ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
liFila = liFila + 1
Set rs = Nothing
Set rs = ObtieneCtaSaldoDiarioAnx15A_Det("212802,232802", "", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn, "Det")
If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            xlHoja1.Cells(liFila, 13) = rs!cCtaCnt
            xlHoja1.Cells(liFila, 14) = Format(rs!nSaldo, "#,##0.00")
            xlHoja1.Cells(liFila, 15) = rs!cDescrip
            ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
        liFila = liFila + 1
            rs.MoveNext
        Loop
    End If
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 15)).Font.Size = 10
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 15)).Font.Name = "Arial Narrow"
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 13)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 13)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(liFilaIni, 14), xlHoja1.Cells(liFila, 14)).HorizontalAlignment = xlLeft

'***Obligaciones por Cuentas a Plazo
Set rs = Nothing
liFila = liFila + 2
    xlHoja1.Cells(liFila, 13) = "OBLIGACIONES POR CUENTAS A PLAZO"
    xlHoja1.Range(xlHoja1.Cells(liFila, 13), xlHoja1.Cells(liFila, 15)).Merge True
ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
liFila = liFila + 1
    xlHoja1.Cells(liFila, 13) = "Cuenta"
    xlHoja1.Cells(liFila, 14) = "Saldo"
    xlHoja1.Cells(liFila, 15) = "Descripción"
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).Interior.ColorIndex = 9
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).Font.Color = vbWhite
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
liFila = liFila + 1
    liFilaIni = liFila

'MN
Set rs = DAnxRies.ObtieneOtrasObligaciones15A(pdFecha, "1", "233", "Det")
If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            xlHoja1.Cells(liFila, 13) = rs!cCtaCnt
            xlHoja1.Cells(liFila, 14) = Format(rs!nSaldo, "#,##0.00")
            xlHoja1.Cells(liFila, 15) = rs!cDescrip
            ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
        If rs!cCtaCnt = "211704" Then
                xlHoja1.Range(xlHoja1.Cells(liFila, 13), xlHoja1.Cells(liFila, 15)).Interior.ColorIndex = 44
            End If
            liFila = liFila + 1
            rs.MoveNext
        Loop
    End If
'ME
Set rs = Nothing
Set rs = DAnxRies.ObtieneOtrasObligaciones15A(pdFecha, "2", "233", "Det")
If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            xlHoja1.Cells(liFila, 13) = rs!cCtaCnt
            xlHoja1.Cells(liFila, 14) = Format(rs!nSaldo, "#,##0.00")
            xlHoja1.Cells(liFila, 15) = rs!cDescrip
            ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
        If rs!cCtaCnt = "212704" Then
                xlHoja1.Range(xlHoja1.Cells(liFila, 13), xlHoja1.Cells(liFila, 15)).Interior.ColorIndex = 44
            End If
            liFila = liFila + 1
            rs.MoveNext
        Loop
    End If

    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 15)).Font.Size = 10
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 15)).Font.Name = "Arial Narrow"
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 13)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 13)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(liFilaIni, 14), xlHoja1.Cells(liFila, 14)).HorizontalAlignment = xlLeft

'***Obligaciones a la Vista
Set rs = Nothing
liFila = liFila + 2
    xlHoja1.Cells(liFila, 13) = "OBLIGACIONES A LA VISTA"
    xlHoja1.Range(xlHoja1.Cells(liFila, 13), xlHoja1.Cells(liFila, 15)).Merge True
ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
liFila = liFila + 1
    xlHoja1.Cells(liFila, 13) = "Cuenta"
    xlHoja1.Cells(liFila, 14) = "Saldo"
    xlHoja1.Cells(liFila, 15) = "Descripción"
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).Interior.ColorIndex = 9
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).Font.Color = vbWhite
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
liFila = liFila + 1
    liFilaIni = liFila
'MN
Set rs = Nothing
Set rs = ObtieneCtaSaldoDiarioAnx15A_Det("2111,2311,211801,231801", "211118", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn, "Det") 'NAGL 20191015 Agregó 211118, para excluirlo de Obligaciones a la Vista
If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            xlHoja1.Cells(liFila, 13) = rs!cCtaCnt
            xlHoja1.Cells(liFila, 14) = Format(rs!nSaldo, "#,##0.00")
            xlHoja1.Cells(liFila, 15) = rs!cDescrip
            ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
        liFila = liFila + 1
            rs.MoveNext
        Loop
    End If
'ME
Set rs = Nothing
Set rs = ObtieneCtaSaldoDiarioAnx15A_Det("2121,2321,212801,232801", "212118", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn, "Det") 'NAGL 20191015 Agregó 212118, para excluirlo de Obligaciones a la Vista
If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            xlHoja1.Cells(liFila, 13) = rs!cCtaCnt
            xlHoja1.Cells(liFila, 14) = Format(rs!nSaldo, "#,##0.00")
            xlHoja1.Cells(liFila, 15) = rs!cDescrip
            ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
        liFila = liFila + 1
            rs.MoveNext
        Loop
    End If
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 15)).Font.Size = 10
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 15)).Font.Name = "Arial Narrow"
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 13)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 13)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(liFilaIni, 14), xlHoja1.Cells(liFila, 14)).HorizontalAlignment = xlLeft

'***Obligaciones con instituciones recaudadoras de tributos
Set rs = Nothing
liFila = liFila + 2
    xlHoja1.Cells(liFila, 13) = "OBLIGACIONES CON INST.RECAUDADORAS DE TRIBUTOS"
    xlHoja1.Range(xlHoja1.Cells(liFila, 13), xlHoja1.Cells(liFila, 15)).Merge True
ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
liFila = liFila + 1
    xlHoja1.Cells(liFila, 13) = "Cuenta"
    xlHoja1.Cells(liFila, 14) = "Saldo"
    xlHoja1.Cells(liFila, 15) = "Descripción"
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).Interior.ColorIndex = 9
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).Font.Color = vbWhite
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 13), xlHoja1.Cells(liFila, 15)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
liFila = liFila + 1
    liFilaIni = liFila
'MN
Set rs = Nothing
Set rs = ObtieneCtaSaldoDiarioAnx15A_Det("25170301,25170302,25170303,25170309,251704,251705,251706,251801", "", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn, "Det")
If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            xlHoja1.Cells(liFila, 13) = rs!cCtaCnt
            xlHoja1.Cells(liFila, 14) = Format(rs!nSaldo, "#,##0.00")
            xlHoja1.Cells(liFila, 15) = rs!cDescrip
            ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
        liFila = liFila + 1
            rs.MoveNext
        Loop
    End If
'ME
Set rs = Nothing
Set rs = ObtieneCtaSaldoDiarioAnx15A_Det("25270301,25270302,25270303,25270309,252704,252705,252706,252801", "", pdFecha, pdFechaBalanceDiario, nTipoCambioBalance, nTipoCambioAn, "Det")
If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            xlHoja1.Cells(liFila, 13) = rs!cCtaCnt
            xlHoja1.Cells(liFila, 14) = Format(rs!nSaldo, "#,##0.00")
            xlHoja1.Cells(liFila, 15) = rs!cDescrip
            ExcelCuadro xlHoja1, 13, liFila, 15, CCur(liFila)
        liFila = liFila + 1
            rs.MoveNext
        Loop
    End If
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 15)).Font.Size = 10
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 15)).Font.Name = "Arial Narrow"
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 13)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 2, 13), xlHoja1.Cells(liFila, 13)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(liFilaIni, 14), xlHoja1.Cells(liFila, 14)).HorizontalAlignment = xlLeft

    '*************************END NAGL 20190617**************************************'
End Sub '***************NAGL ERS079-2017 20180123

Private Sub ObtieneOtrasSeccionesAnexo15A(ByVal xlAplicacion As Excel.Application, ByVal pdFecha As Date, ByVal nTipoCambioAn As Currency, ByVal lnFila As Integer)
    Dim oDbalanceCont As DbalanceCont
    Dim rsvalor As ADODB.Recordset
    Dim oValor As DAnexoRiesgos
    Dim nContar As Integer, X As Integer, nMoneda As Integer
    Dim nSaldoDiario1 As Currency, nSaldoDiario2 As Currency
    Set oDbalanceCont = New DbalanceCont
    Set rsvalor = New ADODB.Recordset
    Set oValor = New DAnexoRiesgos
    '******DEPOSITOS A GRANDES ACREEDORES
    lnFila = lnFila + 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Dim d As ColorConstants
    Call PintaFilasExcel(xlHoja1, "IV.  SALDO DE DEPÓSITOS DE GRANDES ACREEDORES", "", "Moneda Nacional", "Moneda Extranjera", lnFila, False, True, False)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).HorizontalAlignment = xlCenter
    Call oDbalanceCont.InsertaDetallaReporte15A(58, pdFecha, "1", 0, 0, "3320", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(58, pdFecha, "2", 0, 0, "3320", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado("1", pdFecha, nTipoCambioAn)
    nSaldoDiario2 = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado("2", pdFecha, nTipoCambioAn)
    Call PintaFilasExcel(xlHoja1, "1.     Estado(26)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(59, pdFecha, "1", nSaldoDiario1, 0, "3322", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(59, pdFecha, "2", nSaldoDiario2, 0, "3322", "A1")

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Otros("1", pdFecha, nTipoCambioAn, "AFP") 'NAGL 20190627
    nSaldoDiario2 = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Otros("2", pdFecha, nTipoCambioAn, "AFP") 'NAGL 20190627
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "2.     AFPs (27)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(60, pdFecha, "1", nSaldoDiario1, 0, "3324", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(60, pdFecha, "2", nSaldoDiario2, 0, "3324", "A1")

    lnFila = lnFila + 1
    nSaldoDiario1 = 0
    nSaldoDiario2 = 0
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "3.     Fondos mutuos y fondos de inversión.", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(61, pdFecha, "1", nSaldoDiario1, 0, "3326", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(61, pdFecha, "2", nSaldoDiario2, 0, "3326", "A1")

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Otros("1", pdFecha, nTipoCambioAn, "ESeg") 'NAGL 20190627
    nSaldoDiario2 = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Otros("2", pdFecha, nTipoCambioAn, "ESeg") 'NAGL 20190627
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "4.     Empresas del sistema de seguros (28)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(62, pdFecha, "1", nSaldoDiario1, 0, "3328", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(62, pdFecha, "2", nSaldoDiario2, 0, "3328", "A1")

    lnFila = lnFila + 1
    nSaldoDiario1 = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Otros("1", pdFecha, nTipoCambioAn, "SAB") 'NAGL 20190627
    nSaldoDiario2 = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Otros("2", pdFecha, nTipoCambioAn, "SAB") 'NAGL 20190627
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "5.     Sociedad agente de bolsa (SAB)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(63, pdFecha, "1", nSaldoDiario1, 0, "3329", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(63, pdFecha, "2", nSaldoDiario2, 0, "3329", "A1")

    lnFila = lnFila + 1
    Set rsvalor = oValor.ListaOtrosDepositantes(20, pdFecha, "1", nTipoCambioAn, "Gen") 'NAGL 20190511 ERS006-2019
    'nSaldoDiario1 = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptaciones(20, pdFecha, "1", nTipoCambioAn) 'Comentado by NAGL 20190510
    nSaldoDiario1 = Format(rsvalor!nSaldoMN, gsFormatoNumeroView) 'NAGL 20190511 ERS006-2019
    nSaldoDiario2 = Format(rsvalor!nSaldoME, gsFormatoNumeroView) 'NAGL 20190627 ERS006-2019 'oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptaciones(20, pdFecha, "2", nTipoCambioAn)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    Call PintaFilasExcel(xlHoja1, "6.     Otros depositantes (29)", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(64, pdFecha, "1", nSaldoDiario1, 0, "3330", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(64, pdFecha, "2", nSaldoDiario2, 0, "3330", "A1")

    lnFila = lnFila + 1

    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "", "", "", "", lnFila, False, True, False)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)

    lnFila = lnFila + 2

    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "V.  SALDO DE DEPÓSITOS DE EMPRESAS DEL SISTEMA FINANCIERO", "", "Moneda Nacional", "Moneda Extranjera", lnFila, False, True, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(65, pdFecha, "1", 0, 0, "3333", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(65, pdFecha, "2", 0, 0, "3333", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaSaldoDiario("2313", pdFecha) + oDbalanceCont.ObtenerCtaSaldoDiario("2312", pdFecha) + oDbalanceCont.ObtenerCtaSaldoDiario("2318", pdFecha)
    nSaldoDiario2 = Round(oDbalanceCont.ObtenerCtaSaldoDiario("2323", pdFecha), 2) + oDbalanceCont.ObtenerCtaSaldoDiario("2322", pdFecha) + oDbalanceCont.ObtenerCtaSaldoDiario("2328", pdFecha)
    Call PintaFilasExcel(xlHoja1, "1. Sistema financiero nacional", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(66, pdFecha, "1", nSaldoDiario1, 0, "3335", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(66, pdFecha, "2", nSaldoDiario2, 0, "3335", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    nSaldoDiario1 = 0
    nSaldoDiario2 = 0
    Call PintaFilasExcel(xlHoja1, "2. Sistema financiero del exterior", "", nSaldoDiario1, nSaldoDiario2, lnFila, True, False, True)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(67, pdFecha, "1", nSaldoDiario1, 0, "3338", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(67, pdFecha, "2", nSaldoDiario2, 0, "3338", "A1")

    lnFila = lnFila + 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "VI. NUMERO DE DÍAS DE REDESCUENTO EN LOS ÚLTIMOS 180 DÍAS (30)", "", "", "", lnFila, False, True, False)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 6, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(68, pdFecha, "1", 0, 0, "3340", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(68, pdFecha, "2", 0, 0, "3340", "A1")

    oBarra.Progress 9, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "VII. PÉRDIDA NETA DE DERIVADOS PARA NEGOCIACIÓN (31)", "", "", "", lnFila, False, True, False)
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 6, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(69, pdFecha, "1", 0, 0, "3345", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(69, pdFecha, "2", 0, 0, "3345", "A1")

    lnFila = lnFila + 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Interior.Color = RGB(153, 153, 255)
    Call PintaFilasExcel(xlHoja1, "VIII.    POSICIONES DE CAMBIO", "", "Global (33)", "", lnFila, False, True, False)
    xlHoja1.Cells(lnFila, 4) = "Balance (32)"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 6)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 6, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(70, pdFecha, "1", 0, 0, "3350", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(70, pdFecha, "2", 0, 0, "3350", "A1")

    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True
    'nSaldoDiario1 = Round((oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("1", pdFecha, "2") - oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("2", pdFecha, "2")) / nTipoCambioAn, 2)
    'nSaldoDiario1 = oDbalanceCont.GetSaldoMEPosCambiaria(Format(pdFecha, "YYYYMMDD"), "1") - oDbalanceCont.GetSaldoMEPosCambiaria(Format(pdFecha, "YYYYMMDD"), "2")
    nSaldoDiario1 = oDbalanceCont.GetSaldoMEPosCambiaria(Format(pdFecha, "YYYYMMDD"), "1", "BitRep") - oDbalanceCont.GetSaldoMEPosCambiaria(Format(pdFecha, "YYYYMMDD"), "2", "BitRep") 'NAGL 20170803
    nSaldoDiario2 = nSaldoDiario1
    Call oDbalanceCont.InsertaDetallaReporte15A(71, pdFecha, "1", 0, 0, "3400", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(71, pdFecha, "2", nSaldoDiario2, 0, "3400", "A1")

    xlHoja1.Cells(lnFila, 2) = "Moneda Extranjera"
    xlHoja1.Cells(lnFila, 4) = nSaldoDiario1
    xlHoja1.Cells(lnFila, 6) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 6)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 6)).Font.Bold = False

    ExcelCuadro xlHoja1, 2, CCur(lnFila), 6, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(72, pdFecha, "1", 0, 0, "3500", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(72, pdFecha, "2", nSaldoDiario2, 0, "3500", "A1")

    lnFila = lnFila + 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 5)).Font.Bold = False
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Cells(lnFila, 2) = ""
    xlHoja1.Cells(lnFila, 4) = "En Moneda Extranjera (USD)"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 5, CCur(lnFila)

    lnFila = lnFila + 1

    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 5)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True
    'xlHoja1.Cells(lnFila, 2) = "IX.    Posición Neta en Productos Financieros Derivados (Moneda Extranjera / PEN)" 'Comentado by NAGL 20180419
    xlHoja1.Cells(lnFila, 2) = "IX.    Posición Contable Neta en Productos Financieros Derivados (Moneda Extranjera / PEN)" 'NAGL 20180419 RFC1804130002
    xlHoja1.Cells(lnFila, 4) = ""
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).HorizontalAlignment = xlHAlignJustify
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).VerticalAlignment = xlVAlignTop
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).RowHeight = 27
    ExcelCuadro xlHoja1, 2, CCur(lnFila), 5, CCur(lnFila)
    Call oDbalanceCont.InsertaDetallaReporte15A(73, pdFecha, "1", 0, 0, "3600", "A1")
    Call oDbalanceCont.InsertaDetallaReporte15A(73, pdFecha, "2", 0, 0, "3600", "A1")
    lnFila = lnFila + 1

    'oEst.EliminaEstadAnexos pdFecha, "LIQUIDSOLES", "1"
    'oEst.EliminaEstadAnexos pdFecha, "LIQUIDDOLARES", "2"
    'oEst.InsertaEstadAnexos pdFecha, "LIQUIDSOLES", "1", Trim(Str(nMonto1))
    'oEst.InsertaEstadAnexos pdFecha, "LIQUIDDOLARES", "2", Trim(Str(nMonto2))

    lnFila = lnFila + 1
    Set rsvalor = oValor.ListaDepositosSFN(pdFecha, nTipoCambioAn)
    If Not (rsvalor.BOF And rsvalor.EOF) Then
        xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 15
        xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Merge True
        Call PintaFilasExcel(xlHoja1, "DEPÓSITOS DE EMPRESAS DEL SISTEMA FINANCIERO", "", "", "", lnFila, False, True, False)
        xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Borders.LineStyle = xlContinuous
        ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
        lnFila = lnFila + 1
        xlHoja1.Cells(lnFila, 2) = "Institución Publica"
        xlHoja1.Cells(lnFila, 3) = "Fecha Apertura"
        xlHoja1.Cells(lnFila, 4) = "Fecha Vencimiento"
        xlHoja1.Cells(lnFila, 5) = "TEA"
        xlHoja1.Cells(lnFila, 6) = "Moneda"
        xlHoja1.Cells(lnFila, 7) = "Monto"
        xlHoja1.Cells(lnFila, 8) = "Int.Deveng"
        'Call PintaFilasExcel2(xlHoja1, "Institución Publica", "Fecha Apertura", "Fecha Vencimiento", "TEA", "Moneda", "Monto", "", lnFila, False, False, True)
        xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 15
        ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)
            For X = 0 To rsvalor.RecordCount - 1
            lnFila = lnFila + 1
            Call PintaFilasExcel3(xlHoja1, rsvalor!Cliente, Format(rsvalor!dApertura, gsFormatoFecha), IIf(Format(rsvalor!dFecVencimiento, gsFormatoFecha) = "01/01/1900", "", Format(rsvalor!dFecVencimiento, gsFormatoFecha)), Format(rsvalor!TEA, gsFormatoNumeroView), Format(rsvalor!cMoneda, gsFormatoNumeroView), Format(rsvalor!Monto, gsFormatoNumeroView), Format(rsvalor!nIntDeveng, gsFormatoNumeroView), "", lnFila, False, False, False)
            xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlCenter
            ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)
                rsvalor.MoveNext
        Next
    Else
        lnFila = lnFila + 1
    End If
    lnFila = lnFila + 2
    'Set rsvalor = oValor.ListaDepositosESTADO(pdFecha, nTipoCambioAn)'Comentado by NAGL 20190627
    Set rsvalor = oDbalanceCont.ListarSaldoDepGrandesAcreedores("0", pdFecha, nTipoCambioAn, "", "Det")
    If Not (rsvalor.BOF And rsvalor.EOF) Then
        xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 15
        xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Merge True
        Call PintaFilasExcel(xlHoja1, "DEPÓSITOS DE GRANDES ACREEDORES", "", "", "", lnFila, False, True, False)
        xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Borders.LineStyle = xlContinuous
        ExcelCuadro xlHoja1, 2, CCur(lnFila), 7, CCur(lnFila)
        lnFila = lnFila + 1
        xlHoja1.Cells(lnFila, 2) = "Institución Publica"
        xlHoja1.Cells(lnFila, 3) = "Fecha Apertura"
        xlHoja1.Cells(lnFila, 4) = "Fecha Vencimiento"
        xlHoja1.Cells(lnFila, 5) = "TEA"
        xlHoja1.Cells(lnFila, 6) = "Moneda"
        xlHoja1.Cells(lnFila, 7) = "Monto"
        xlHoja1.Cells(lnFila, 8) = "Tipo Acreedor"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Interior.ColorIndex = 15
        ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)
        For X = 0 To rsvalor.RecordCount - 1
            lnFila = lnFila + 1
            Call PintaFilasExcel3(xlHoja1, rsvalor!Cliente, Format(rsvalor!dApertura, gsFormatoFecha), IIf(Format(rsvalor!dFecVencimiento, gsFormatoFecha) = "01/01/1900", "", Format(rsvalor!dFecVencimiento, gsFormatoFecha)), Format(rsvalor!TEA, gsFormatoNumeroView), Format(rsvalor!cMoneda, gsFormatoNumeroView), Format(rsvalor!Monto, gsFormatoNumeroView), rsvalor!cTipoDetPers, "", lnFila, False, False, False)
            xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 7)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlLeft 'NAGL 20190627
            ExcelCuadro xlHoja1, 2, CCur(lnFila), 8, CCur(lnFila)
            rsvalor.MoveNext
        Next
    Else
        lnFila = lnFila + 1
    End If
    Set rsvalor = Nothing
    lnFila = lnFila + 2
    nContar = 0
    For nMoneda = 1 To 2
        Set rsvalor = oValor.ListaOtrosDepositantes(20, pdFecha, CStr(nMoneda), nTipoCambioAn, "Det")
        If Not (rsvalor.BOF And rsvalor.EOF) Then
            xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 4)).Interior.ColorIndex = 15
            xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 4)).Merge True
            Call PintaFilasExcel(xlHoja1, "DEPOSITOS DE OTROS DEPOSITANTES - " & IIf(nMoneda = 1, "MN", "ME"), "", "", "", lnFila, False, True, False)
            xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 4)).Borders.LineStyle = xlContinuous
            ExcelCuadro xlHoja1, 2, CCur(lnFila), 4, CCur(lnFila)
            lnFila = lnFila + 1

            xlHoja1.Cells(lnFila, 2) = "Cliente"
            xlHoja1.Cells(lnFila, 3) = "Saldo"
            xlHoja1.Cells(lnFila, 4) = "Personería"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 4)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila, 4)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 4)).Interior.ColorIndex = 15
            ExcelCuadro xlHoja1, 2, CCur(lnFila), 4, CCur(lnFila)
            For X = 0 To rsvalor.RecordCount - 1
                lnFila = lnFila + 1
                nContar = nContar + 1
                Call PintaFilasExcel2(xlHoja1, rsvalor!Cliente, Format(rsvalor!Saldo, gsFormatoNumeroView), rsvalor!cPersoneria, "", "", "", "", lnFila, False, False, False)
                xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).HorizontalAlignment = xlLeft
                ExcelCuadro xlHoja1, 2, CCur(lnFila), 4, CCur(lnFila)
                rsvalor.MoveNext
            Next
        Else
            lnFila = lnFila + 1
        End If
        lnFila = lnFila + 2
        nContar = 0
        Set rsvalor = Nothing
    Next nMoneda
End Sub '****NAGL Según Anexo02 ERS006-2019 20190615



Private Sub Form_Load()
    CentraForm Me
End Sub


