VERSION 5.00
Begin VB.Form frmRepEncajeBCR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Encaje BCR"
   ClientHeight    =   615
   ClientLeft      =   1095
   ClientTop       =   2160
   ClientWidth     =   2535
   Icon            =   "frmRepEncajeBCR.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmRepEncajeBCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TCuentasCmac
    cDescrip As String
    cCta As String
End Type
Private Type TCreditos
    cDescrip As String
    cCta As String
End Type
Dim oBarra As clsProgressBar

'************** Nombre del Archivo *********************
Dim lsArchivo   As String
Dim lsMoneda    As String
Dim lbExcelOpen As Boolean
Dim lbMN        As Boolean
Dim lsCadenaMoneda As String
Dim lsLineas() As String
Dim lnTotalLineas As Long
Dim lsReportes() As String
Dim Totales() As Currency
Dim Filas As Long
'Variables de Excel
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim xlHojaP As Excel.Worksheet

Dim Col As Long
Dim Columna As Long
Dim ColumnaFinal As Long

Dim TotalColumnas As Long
'Anexo1
Dim lsAnexo1() As String
Dim lsTotalesAnexo1() As String
Dim FilaEncaje As Integer

'Anexo 2
Dim lsAnexo2Col() As String
Dim lsAnexo2() As String
Dim lsTotalesAnexo2() As String
Dim TotalCol As Long
Dim SubCol As Long
Dim ColFinal As Long
Dim CuentasCmac() As TCuentasCmac
Dim Creditos() As TCreditos
Dim lnTCF  As Currency
Dim lnTCFF As Currency
Dim nCantE As Integer
Dim MNumEnc() As String
Dim nIntNEnc As Integer
'******* UTILIZADO PARA EL CALCULO DEL ENCAJE EXIJIBLE EN DOLARES **************
Private Type EncajeBasico
    lnDias As Integer
    lnTose As Currency
    lnToseBase As Currency
    lnExceso As Currency
    lnEncBasico As Currency
    lnTotalEncaje As Currency
    lnEncajeMin As Currency
    lnEncajeExig As Currency
    lnEncajeMarginal As Currency
End Type
Dim Encaje() As EncajeBasico

Public Sub ImprimeEncajeBCR(psOpeCod As String, pdFecIni As Date, pdFecFin As Date, ByVal pnlnTCF As Double, ByVal plnTCFF As Double)
 'On Error GoTo ImprimeEncajeBCRErr
lsMoneda = Mid(psOpeCod, 3, 1)
lsArchivo = App.path & "\SPOOLER\ENCAJEBCR_" & Left(Format(pdFecIni, gsFormatoMovFecha), 6) & IIf(lsMoneda = "1", "MN", "ME") & ".XLS"

If lsMoneda = "2" Then
Dim oTC As nTipoCambio
   'Set oTC = New NTipoCambio
   'lnTCF = oTC.EmiteTipoCambio(pdFecFin, TCFijoMes)
   'lnTCFF = oTC.EmiteTipoCambio(pdFecFin + 1, TCFijoMes)
   'Set oTC = Nothing
   lnTCF = pnlnTCF
   lnTCFF = plnTCFF
End If
If Not ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False) Then
   Exit Sub
End If
lbExcelOpen = True
   Select Case psOpeCod
      Case RepCGEncBCRObligacion, RepCGEncBCRObligacionME
         ImprimeObligacionSujetaEnc psOpeCod, pdFecIni, pdFecFin, 15
      Case RepCGEncBCRCredDeposi, RepCGEncBCRCredDeposiME
         ImprimeCreditosDepositosInterbancarios psOpeCod, pdFecIni, pdFecFin, 15, True
      Case RepCGEncBCRCredRecibi, RepCGEncBCRCredRecibiME
         ImprimeCreditosDepositosInterbancarios psOpeCod, pdFecIni, pdFecFin, 15, False
      Case RepCGEncBCRObligaExon, RepCGEncBCRObligaExonME
         ImprimeObligacionSujetaEnc psOpeCod, pdFecIni, pdFecFin, 15
      Case RepCGEncBCRLinCredExt, RepCGEncBCRLinCredExtME
         ImprimeLineaCreditoExterior psOpeCod, pdFecIni, pdFecFin, Right(psOpeCod, 1)
   End Select
ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
MsgBox "Archivo generado satisfactoriamente", vbInformation, "Aviso"
CargaArchivo lsArchivo, App.path & "\SPOOLER"
lbExcelOpen = False
Exit Sub
ImprimeEncajeBCRErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    If lbExcelOpen = True Then
         ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Sub

Private Sub IniciaHojaReporte(psOpeCod As String)
Dim Pos As Integer
If Mid(psOpeCod, 3, 1) = "1" Then
    lbMN = True
    lsCadenaMoneda = "(EN NUEVOS SOLES)"
Else
    lbMN = False
    lsCadenaMoneda = "(EN US$ DOLARES)"
End If
End Sub


Private Function ImprimeObligacionSujetaEnc(psOpeCod As String, pdFecIni As Date, pdFecFin As Date, FilaDetalle As Integer) As String
Dim Fila, Col     As Long
Dim j, i, L, m    As Long
Dim lnTotalEncaje As Currency
Dim lnResultado   As Currency
Dim ldFecha       As Date
Dim lsTotal1      As String
Dim lsOperador    As String
Dim lsTotalesCol() As String
Dim lsEncajeSoles As String
Dim lsTotalEncaje As String
Dim lsAnexo       As String

Dim lsMontoEncaje As String
Dim lsFondoEncaje As String
Dim R             As ADODB.Recordset
Dim MatTemp()     As String
Dim MatTemp2()    As String
Dim MatTemp3()    As String
Dim Total1        As Double
Dim Total2        As Double
Dim nCont       As Integer
nCont = 1
Dim rsEnc As ADODB.Recordset
Set rsEnc = New ADODB.Recordset
Dim oEncBcr As NEncajeBCR
Dim nIntNEnc2 As Integer
Set oBarra = New clsProgressBar
Set oEncBcr = New NEncajeBCR
'ALPA 20080811******************
Set rsEnc = oEncBcr.ObtenerRepColumnaCol(psOpeCod)
nIntNEnc = 0
nIntNEnc2 = 1
ReDim Preserve MNumEnc(0 To 0)
Do While Not rsEnc.EOF
nIntNEnc = nIntNEnc + 1

ReDim Preserve MNumEnc(nIntNEnc)
    MNumEnc(nIntNEnc) = rsEnc!nNroCol
rsEnc.MoveNext
'*********************************/
Loop
GeneraDatosColumnas pdFecIni, pdFecFin, psOpeCod, lsAnexo1, lsTotalesAnexo1
IniciaHojaReporte psOpeCod

ExcelAddHoja "Anx_" & Right(psOpeCod, 1), xlLibro, xlHoja1, False

ProgressShow oBarra, Me, ePCap_CaptionPercent, True
oBarra.Max = 6
oBarra.Progress 0, "ENCAJE BCR", "Generando Reporte", , vbBlue
lsAnexo = Right(psOpeCod, 1)
Select Case lsAnexo
    Case "1"
        
        CabeceraAnexo1 lsMoneda, pdFecIni, pdFecFin, psOpeCod
        If lsMoneda = "1" Then
        FilaDetalle = 16
        Else
        FilaDetalle = 16 'MIOL 20120712 16 reemplaza a 18
        End If
        
    Case "4"
        CabeceraAnexo4 pdFecIni, pdFecFin, psOpeCod
         If lsMoneda = "1" Then
            FilaDetalle = 18
        Else
            FilaDetalle = 17
        End If
End Select

Fila = FilaDetalle
ReDim lsTotalesCol(UBound(lsAnexo1, 1))

L = 0
ReDim MatTemp3(UBound(lsAnexo1, 2))
oBarra.Progress 1, "ENCAJE BCR", "Generando Reporte", , vbBlue
For i = 1 To UBound(lsAnexo1, 2)   'filas
    Fila = Fila + 1
    Col = 1
    xlHoja1.Cells(Fila, Col) = i
    lsTotal1 = ""
    If lsAnexo = "1" Then
        For j = 1 To UBound(lsAnexo1, 1)  'columnas
            Col = Col + 1
            If lsMoneda = "1" Then
                Select Case j
                    'MIOL 20120711, SEGUN RQ11370_RQ11371_RQ11372_RQ11373_RQ12147 ****
                    Case 1, 2, 4, 5, 7, 9, 10, 11, 12, 13, 14
                        lsOperador = "+"
                    Case 3, 6, 8
                        lsOperador = "-"
'                    Case 1, 2, 4, 5, 7, 10, 12, 14
'                        lsOperador = "+"
'                    Case 3, 6, 8, 9, 11, 13
'                        lsOperador = "-"
                    'END MIOL ********************************************************
                    Case Else
                        'Cambiado: Agregado
                        If nIntNEnc > 0 Then
                            If nIntNEnc >= nIntNEnc2 Then
                                If Col = MNumEnc(nIntNEnc2) Then
                                    nIntNEnc2 = nIntNEnc2 + 1
                                End If
                                'nIntNEnc = nIntNEnc + 1
                            End If
                        End If
                        'Fin Agregado
                End Select
            Else
                Select Case j
                    Case 1, 2, 4, 5, 7, 9, 10, 11, 12, 13, 14 'MIOL 20120712 Se Agrego 14
                        lsOperador = "+"
                    Case 3, 6, 8
                        lsOperador = "-"
                End Select
            End If
            If j = FilaEncaje Then
                xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Formula = "=Sum(" & Mid(lsTotal1, 2, Len(lsTotal1)) & ")"
                
                'xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Formula = lsTotal1
                If lsAnexo = "1" And lsMoneda = "2" Then
                    L = L + 1
                    MatTemp3(L) = xlHoja1.Cells(Fila, Col)
                End If
                
            Else
                lsTotal1 = lsTotal1 + lsOperador + xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Address(False, False)
                xlHoja1.Cells(Fila, Col) = lsAnexo1(j, i)
            End If
            lsTotalesCol(j) = lsTotalesCol(j) + xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Address(False, False) & "+"
        Next j
    Else
        For j = 1 To UBound(lsAnexo1, 1)  'columnas
            Col = Col + 1
'            If lsAnexo = "4" And Col = 9 Then
'               Col = Col + 2
'               xlHoja1.Cells(Fila, Col) = lsAnexo1(j, I)
'               lsTotalesCol(j) = lsTotalesCol(j) + xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Address(False, False) & "+"
'               Col = Col - 2
'            Else
               xlHoja1.Cells(Fila, Col) = lsAnexo1(j, i)
               lsTotalesCol(j) = lsTotalesCol(j) + xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Address(False, False) & "+"
'            End If
               
        Next j
    End If
Next i
Fila = Fila + 1
Col = 1
xlHoja1.Cells(Fila, Col) = "Total"
oBarra.Progress 2, "ENCAJE BCR", "Generando Reporte", , vbBlue

For j = 1 To UBound(lsTotalesCol)  'columnas
    Col = Col + 1
        xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Formula = "=Sum(" & Mid(lsTotalesCol(j), 1, Len(lsTotalesCol(j)) - 1) & ")"
    If j = FilaEncaje Then
        lsEncajeSoles = xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Address(False, False)
    End If
    '''CAMBIADO
    '''ORIGINAL
    'If J = UBound(lsTotalesCol) Then
    '    lsTotalEncaje = xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Address(False, False)
    'End If
    '''fin Original
    
    '''Nuevo
    If lsMoneda = "1" Then
        If j = UBound(lsTotalesCol) - 2 Then
            lsTotalEncaje = xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Address(False, False)
        End If
    ElseIf lsMoneda = "2" Then
        If j = UBound(lsTotalesCol) Then
        'MIOL 20120713 SEGUN RQ11370_RQ11371_RQ11372_RQ11373_RQ12147 Se cambio Col por Col -2 *********
            lsTotalEncaje = xlHoja1.Range(xlHoja1.Cells(Fila, Col - 2), xlHoja1.Cells(Fila, Col - 2)).Address(False, False)
        'END MIOL *****************************************************************************
        End If
    End If
    ''Fin Nuevo
    
Next j

oBarra.Progress 3, "ENCAJE BCR", "Generando Reporte", , vbBlue
For i = FilaDetalle To Fila
    For m = 1 To Col
        If m = 1 Then
            xlHoja1.Range(xlHoja1.Cells(i, m), xlHoja1.Cells(i, m)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        End If
        xlHoja1.Range(xlHoja1.Cells(i, m), xlHoja1.Cells(i, m)).Borders(xlEdgeRight).LineStyle = xlContinuous
        If i = Fila Then
            xlHoja1.Range(xlHoja1.Cells(i, m), xlHoja1.Cells(i, m)).Borders.LineStyle = xlContinuous
        End If
    Next m
Next i
oBarra.Progress 4, "ENCAJE BCR", "Generando Reporte", , vbBlue

If lsAnexo = "1" Then
   Dim oPara As New NEncajeBCR
   If lsMoneda = "2" Then
    Set R = oPara.GetParametroEncaje(pdFecIni)
      ReDim MatTemp(UBound(lsAnexo1, 2))
      ReDim MatTemp2(UBound(lsAnexo1, 2))
      Total1 = 0
      Total2 = 0
      R.Find "nCodigo=2"
      
      'Hallo Exceso diario EncajeActual - Base
      For i = 1 To UBound(lsAnexo1, 2)
          MatTemp(i) = CDbl(MatTemp3(i)) - R!nValor
      Next i
      
      Total1 = R!nValor
      
      R.Find "nCodigo=3"
      'Hallo El Basico : Base * TasaBase (21.6478)%
      For i = 1 To UBound(lsAnexo1, 2)
          MatTemp2(i) = Total1 * R!nValor
      Next i
      Total1 = 0#
      
      R.Find "nCodigo=4"
      'Encaje Marginal: Marginal por el Exeso * 20%
      For i = 1 To UBound(MatTemp)
          MatTemp(i) = CDbl(MatTemp(i)) * R!nValor
      Next i
    
      oBarra.Progress 5, "ENCAJE BCR", "Generando Reporte", , vbBlue
      
      'Total Encaje : En Basico mas Marginal
      For i = 1 To UBound(MatTemp)
          MatTemp2(i) = CDbl(MatTemp(i)) + CDbl(MatTemp2(i))
      Next i
      
      For i = 1 To UBound(MatTemp)
          Total1 = Total1 + CDbl(MatTemp2(i))
      Next i
      Total1 = CDbl(Format(Total1, "#0.00"))
          
      R.Close
      Set R = Nothing
   Else
      Set R = oPara.GetParametroEncaje(pdFecIni)
      R.Find "nCodigo=1"
      Total2 = R!nValor
      R.Close
      Set R = Nothing
   End If

   Fila = Fila + 2
   xlHoja1.Cells(Fila, 2) = "SITUACION DE ENCAJE "
   If lsMoneda = "1" Then
       lnTotalEncaje = CCur(lsTotalesAnexo1(FilaEncaje)) * Total2
   Else
       lnTotalEncaje = Total1
   End If
   Fila = Fila + 2
   xlHoja1.Cells(Fila, 2) = "ENCAJE EXIGIBLE" 'aca me qude ultimo
   If lsMoneda = "1" Then
       xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).Formula = "=(" & lsEncajeSoles & "*" & Format(Total2, "#0.00") & ")"
   Else
       xlHoja1.Cells(Fila, 4) = Format(lnTotalEncaje, "##,##0.00")
   End If
   xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).NumberFormat = "##,##0.00"
   lsMontoEncaje = xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).Address(False, False)
   
   Fila = Fila + 1
   xlHoja1.Cells(Fila, 2) = "FONDO DE ENCAJE"
   'xlHoja1.Cells(Fila, 4) = Format(lsTotalesAnexo1(UBound(lsTotalesAnexo1)), "##,##0.00")

   xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).Formula = "=" & lsTotalEncaje
   
   lsFondoEncaje = xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).Address(False, False)
   
   xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
   
   lnResultado = xlHoja1.Range(lsFondoEncaje) - xlHoja1.Range(lsMontoEncaje)
   
   Dim oEst As New NEstadisticas
   oEst.EliminaEstadAnexos pdFecFin, "ENCAJE_ACUMULA", Mid(psOpeCod, 3, 1)
   oEst.EliminaEstadAnexos pdFecFin, "ENCAJE_MIN_BCR", Mid(psOpeCod, 3, 1)
   oEst.InsertaEstadAnexos pdFecFin, "ENCAJE_ACUMULA", Mid(psOpeCod, 3, 1), Format(lnResultado, "#0.00")
   Dim nPos As Integer
   If FilaEncaje + IIf(Mid(psOpeCod, 3, 1) = "1", 2, 3) > UBound(lsTotalesAnexo1) Then
   nPos = UBound(lsTotalesAnexo1)
   Else
   nPos = IIf(Mid(psOpeCod, 3, 1) = "1", 2, 3) + FilaEncaje
   End If
   oEst.InsertaEstadAnexos pdFecFin, "ENCAJE_MIN_BCR", Mid(psOpeCod, 3, 1), Format(Round(lsTotalesAnexo1(nPos) * 0.01, 2), "#0.00")
   Set oEst = Nothing
   
   Fila = Fila + 1
   xlHoja1.Cells(Fila, 2) = "RESULTADO"
   xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).Formula = "=(" & lsFondoEncaje & "-" & lsMontoEncaje & ")"
   xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).Borders(xlEdgeBottom).LineStyle = xlDouble
Else
   Fila = Fila + 3
    If Mid(gsOpeCod, 3, 1) = "1" And gsOpeCod = "761201" Then
        xlHoja1.Range(xlHoja1.Cells(Fila, 2), xlHoja1.Cells(Fila, 8)).Merge True
        xlHoja1.Cells(Fila, 2) = "(1) Programa de Credito a la Microempresa, segón convenio con Foncodes"
        xlHoja1.Range(xlHoja1.Cells(Fila, Col - 2), xlHoja1.Cells(Fila, Col)).Merge True
        xlHoja1.Cells(Fila, Col - 2) = "Capital y Reservas"
        xlHoja1.Range(xlHoja1.Cells(Fila, Col - 2), xlHoja1.Cells(Fila, Col)).Borders.LineStyle = xlContinuous
        Fila = Fila + 1
    End If
    oBarra.Progress 5, "ENCAJE BCR", "Generando Reporte", , vbBlue
    If Mid(gsOpeCod, 3, 1) = "1" And gsOpeCod = "761201" Then
        xlHoja1.Range(xlHoja1.Cells(Fila, Col - 2), xlHoja1.Cells(Fila, Col)).Merge True
        xlHoja1.Cells(Fila, Col) = "Mes Precedente"
        xlHoja1.Range(xlHoja1.Cells(Fila, Col - 2), xlHoja1.Cells(Fila, Col)).Borders.LineStyle = xlContinuous
    End If
   
    If Mid(gsOpeCod, 3, 1) = "1" And gsOpeCod = "761201" Then
        xlHoja1.Range("J1:J1").EntireColumn.Insert
        xlHoja1.Range("L1:L1").EntireColumn.Insert
        xlHoja1.Range("M1:M1").EntireColumn.Insert
        xlHoja1.Range(xlHoja1.Cells(12, 11), xlHoja1.Cells(12, 12)).Merge True
        xlHoja1.Range(xlHoja1.Cells(12, 11), xlHoja1.Cells(12, 12)).Borders.LineStyle = xlContinuous
        xlHoja1.Range("B10:M11").Merge True
        'xlHoja1.Cells(13, 10) = "FIDEICOMISOS"
        xlHoja1.Cells(13, 12) = "PROGRAMA"
        xlHoja1.Cells(13, 13) = "FOCMAC"
    End If
   
End If
oBarra.Progress 6, "ENCAJE BCR", "Generando Reporte", , vbBlue
ProgressClose oBarra, Me, True
Set oBarra = Nothing

End Function


'
'Private Function ImprimeObligacionSujetaEnc(psOpeCod As String, pdFecIni As Date, pdFecFin As Date, FilaDetalle As Integer) As String
'Dim Fila, Col     As Integer
'Dim J, I, L, m    As Integer
'Dim lnTotalEncaje As Currency
'Dim lnResultado   As Currency
'Dim ldFecha       As Date
'Dim lsTotal1      As String
'Dim lsOperador    As String
'Dim lsTotalesCol() As String
'Dim lsEncajeSoles As String
'Dim lsTotalEncaje As String
'Dim lsAnexo       As String
'
'Dim lsMontoEncaje As String
'Dim lsFondoEncaje As String
'Dim R             As ADODB.Recordset
'Dim MatTemp()     As String
'Dim MatTemp2()    As String
'Dim MatTemp3()    As String
'Dim Total1        As Double
'Dim Total2        As Double
'
'Set oBarra = New clsProgressBar
'
'GeneraDatosColumnas pdFecIni, pdFecFin, psOpeCod, lsAnexo1, lsTotalesAnexo1
'
'ExcelAddHoja "Anx_" & Right(psOpeCod, 1), xlLibro, xlHoja1
'
'ProgressShow oBarra, Me, ePCap_CaptionPercent, True
'oBarra.Max = 6
'oBarra.Progress 0, "ENCAJE BCR", "Generando Reporte", , vbBlue
'lsAnexo = Right(psOpeCod, 1)
'Select Case lsAnexo
'    Case "1"
'        CabeceraAnexo1 lsMoneda, pdFecIni, pdFecFin
'    Case "4"
'        CabeceraAnexo4 pdFecIni, pdFecFin
'End Select
'
'Fila = FilaDetalle
'ReDim lsTotalesCol(UBound(lsAnexo1, 1))
'
'L = 0
'ReDim MatTemp3(UBound(lsAnexo1, 2))
'oBarra.Progress 1, "ENCAJE BCR", "Generando Reporte", , vbBlue
'For I = 1 To UBound(lsAnexo1, 2)   'filas
'    Fila = Fila + 1
'    Col = 1
'    xlHoja1.Cells(Fila, Col) = I
'    lsTotal1 = ""
'    If lsAnexo = "1" Then
'        For J = 1 To UBound(lsAnexo1, 1)  'columnas
'            Col = Col + 1
'            If lsMoneda = "1" Then
'                Select Case J
'                    Case 1, 2, 4, 5, 6, 8, 9, 10, 12, 14
'                        lsOperador = "+"
'                    Case 3, 7, 11, 13
'                        lsOperador = "-"
'                    Case Else
'                        'Cambiado: Agregado
'                        If Col = 17 Then
'                            lsTotal1 = ""
'                        End If
'                        'Fin Agregado
'                End Select
'            Else
'                Select Case J
'                    Case 1, 2, 4, 6, 7, 8, 10, 12
'                        lsOperador = "+"
'                    Case 3, 5, 9, 11
'                        lsOperador = "-"
'                End Select
'            End If
'            If J = FilaEncaje Then
'                xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Formula = "=Sum(" & Mid(lsTotal1, 2, Len(lsTotal1)) & ")"
'                If lsAnexo = "1" And lsMoneda = "2" Then
'                    L = L + 1
'                    MatTemp3(L) = xlHoja1.Cells(Fila, Col)
'                End If
'            Else
'                lsTotal1 = lsTotal1 + lsOperador + xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Address(False, False)
'                xlHoja1.Cells(Fila, Col) = lsAnexo1(J, I)
'            End If
'            lsTotalesCol(J) = lsTotalesCol(J) + xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Address(False, False) & "+"
'        Next J
'    Else
'        For J = 1 To UBound(lsAnexo1, 1)  'columnas
'            Col = Col + 1
'            xlHoja1.Cells(Fila, Col) = lsAnexo1(J, I)
'            lsTotalesCol(J) = lsTotalesCol(J) + xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Address(False, False) & "+"
'        Next J
'    End If
'Next I
'Fila = Fila + 1
'Col = 1
'xlHoja1.Cells(Fila, Col) = "Total"
'oBarra.Progress 2, "ENCAJE BCR", "Generando Reporte", , vbBlue
'
'For J = 1 To UBound(lsTotalesCol)  'columnas
'    Col = Col + 1
'        xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Formula = "=Sum(" & Mid(lsTotalesCol(J), 1, Len(lsTotalesCol(J)) - 1) & ")"
'    If J = FilaEncaje Then
'        lsEncajeSoles = xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Address(False, False)
'    End If
'    '''CAMBIADO
'    '''ORIGINAL
'    'If J = UBound(lsTotalesCol) Then
'    '    lsTotalEncaje = xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Address(False, False)
'    'End If
'    '''fin Original
'
'    '''Nuevo
'    If lsMoneda = "1" Then
'        If J = UBound(lsTotalesCol) - 2 Then
'            lsTotalEncaje = xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Address(False, False)
'        End If
'    ElseIf lsMoneda = "2" Then
'        If J = UBound(lsTotalesCol) Then
'            lsTotalEncaje = xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Address(False, False)
'        End If
'    End If
'    ''Fin Nuevo
'
'Next J
'
'oBarra.Progress 3, "ENCAJE BCR", "Generando Reporte", , vbBlue
'For I = FilaDetalle To Fila
'    For m = 1 To Col
'        If m = 1 Then
'            xlHoja1.Range(xlHoja1.Cells(I, m), xlHoja1.Cells(I, m)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'        End If
'        xlHoja1.Range(xlHoja1.Cells(I, m), xlHoja1.Cells(I, m)).Borders(xlEdgeRight).LineStyle = xlContinuous
'        If I = Fila Then
'            xlHoja1.Range(xlHoja1.Cells(I, m), xlHoja1.Cells(I, m)).Borders.LineStyle = xlContinuous
'        End If
'    Next m
'Next I
'oBarra.Progress 4, "ENCAJE BCR", "Generando Reporte", , vbBlue
'
'If lsAnexo = "1" Then
'   Dim oPara As New NEncajeBCR
'   If lsMoneda = "2" Then
'      Set R = oPara.GetParametroEncaje()
'      ReDim MatTemp(UBound(lsAnexo1, 2))
'      ReDim MatTemp2(UBound(lsAnexo1, 2))
'      Total1 = 0
'      Total2 = 0
'      R.Find "nCodigo=2"
'
'      'Hallo Exeso
'      For I = 1 To UBound(lsAnexo1, 2)
'          MatTemp(I) = Format(CDbl(MatTemp3(I)) - R!nValor, "0.00")
'      Next I
'
'      Total1 = R!nValor
'
'      R.Find "nCodigo=3"
'      'Hallo El Basico
'      For I = 1 To UBound(lsAnexo1, 2)
'          MatTemp2(I) = Format(Total1 * R!nValor, "0.00")
'      Next I
'      Total1 = 0#
'
'      R.Find "nCodigo=4"
'      'Encaje Marginal: Marginal por el Exeso
'      For I = 1 To UBound(MatTemp)
'          MatTemp(I) = Format(CDbl(MatTemp(I)) * R!nValor, "0.00")
'      Next I
'
'      oBarra.Progress 5, "ENCAJE BCR", "Generando Reporte", , vbBlue
'
'      'Total Encaje : En Basico mas Marginal
'      For I = 1 To UBound(MatTemp)
'          MatTemp2(I) = Format(CDbl(MatTemp(I)) + CDbl(MatTemp2(I)), "0.00")
'      Next I
'
'      For I = 1 To UBound(MatTemp)
'          Total1 = Total1 + CDbl(MatTemp2(I))
'      Next I
'      Total1 = CDbl(Format(Total1, "#0.00"))
'
'      R.Close
'      Set R = Nothing
'   Else
'      Set R = oPara.GetParametroEncaje()
'      R.Find "nCodigo=1"
'      Total2 = R!nValor
'      R.Close
'      Set R = Nothing
'   End If
'
'   Fila = Fila + 2
'   xlHoja1.Cells(Fila, 2) = "SITUACION DE ENCAJE "
'   If lsMoneda = "1" Then
'       lnTotalEncaje = CCur(lsTotalesAnexo1(FilaEncaje)) * Total2
'   Else
'       lnTotalEncaje = Total1
'   End If
'   Fila = Fila + 2
'   xlHoja1.Cells(Fila, 2) = "ENCAJE EXIGIBLE"
'   If lsMoneda = "1" Then
'   'xlHoja1.Cells(Fila, 4)== Format(lnTotalEncaje, "##,##0.00")
'       xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).Formula = "=(" & lsEncajeSoles & "*" & Format(Total2, "#0.00") & ")"
'   Else
'       xlHoja1.Cells(Fila, 4) = Format(lnTotalEncaje, "##,##0.00")
'   End If
'   lsMontoEncaje = xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).Address(False, False)
'
'   Fila = Fila + 1
'   xlHoja1.Cells(Fila, 2) = "FONDO DE ENCAJE"
'   'xlHoja1.Cells(Fila, 4) = Format(lsTotalesAnexo1(UBound(lsTotalesAnexo1)), "##,##0.00")
'
'   xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).Formula = "=" & lsTotalEncaje
'
'   lsFondoEncaje = xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).Address(False, False)
'
'   xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'
'   lnResultado = CCur(lsTotalesAnexo1(UBound(lsTotalesAnexo1))) - lnTotalEncaje
'
'   Fila = Fila + 1
'   xlHoja1.Cells(Fila, 2) = "RESULTADO"
'
'   xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).Formula = "=(" & lsFondoEncaje & "-" & lsMontoEncaje & ")"
'   xlHoja1.Range(xlHoja1.Cells(Fila, 4), xlHoja1.Cells(Fila, 4)).Borders(xlEdgeBottom).LineStyle = xlDouble
'Else
'   Fila = Fila + 3
'   xlHoja1.Cells(Fila, 2) = "(1) Programa de Credito a la Microempresa, segón convenio con Foncodes"
'   xlHoja1.Cells(Fila, Col - 1) = "Capital y Reservas"
'   xlHoja1.Range(xlHoja1.Cells(Fila, Col - 1), xlHoja1.Cells(Fila, Col)).Borders.LineStyle = xlContinuous
'   Fila = Fila + 1
'   oBarra.Progress 5, "ENCAJE BCR", "Generando Reporte", , vbBlue
'   xlHoja1.Cells(Fila, Col - 1) = "Mes Precedente"
'   xlHoja1.Range(xlHoja1.Cells(Fila, Col - 1), xlHoja1.Cells(Fila, Col)).Borders.LineStyle = xlContinuous
'
'End If
'oBarra.Progress 6, "ENCAJE BCR", "Generando Reporte", , vbBlue
'ProgressClose oBarra, Me, True
'Set oBarra = Nothing
'
'End Function

Private Function ImprimeCreditosDepositosInterbancarios(psOpeCod As String, pdFecIni As Date, pdFecFin As Date, FilaDetalle As Integer, lbCmact As Boolean) As String
Dim Fila, j, i, m As Integer
Dim lsMoneda As String
Set oBarra = New clsProgressBar

ProgressShow oBarra, Me, ePCap_CaptionPercent, True
oBarra.Max = 5
oBarra.Progress 0, "ENCAJE BCR", "Generando Reporte", , vbBlue

ExcelAddHoja "Anx_" & Right(psOpeCod, 1), xlLibro, xlHoja1
lsMoneda = Mid(psOpeCod, 3, 1)
oBarra.Progress 1, "ENCAJE BCR", "Generando Reporte", , vbBlue
GeneraDatosSubColumna psOpeCod, pdFecIni, pdFecFin, Int(Right(psOpeCod, 1)), lbCmact
Select Case Right(psOpeCod, 1)
    Case "2"
            Fila = 19
            CabeceraAnexo2 lsMoneda, pdFecIni, pdFecFin
    Case "3"
            Fila = 24
            CabeceraAnexo3 lsMoneda, pdFecIni, pdFecFin
End Select
oBarra.Progress 2, "ENCAJE BCR", "Generando Reporte", , vbBlue
'Fila = FilaDetalle
'Fila = 24
For i = 1 To UBound(lsAnexo2, 2)   'filas
    Col = 1
    Fila = Fila + 1
    xlHoja1.Cells(Fila, Col) = i
    xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).HorizontalAlignment = xlCenter
    For j = 1 To UBound(lsAnexo2, 1)  'columnas
        Col = Col + 1
        'If Col = 2 Then
        '    Col = 3
        'ElseIf Col = 3 Then
        '    Col = 2
        'End If
        xlHoja1.Cells(Fila, Col) = lsAnexo2(j, i)
        'If Col = 3 Then
        '    Col = 2
        'ElseIf Col = 2 Then
        '    Col = 3
        'End If
    Next j
Next i
oBarra.Progress 3, "ENCAJE BCR", "Generando Reporte", , vbBlue
Fila = Fila + 2
Col = 1
xlHoja1.Cells(Fila, Col) = "Total"
For j = 1 To UBound(lsTotalesAnexo2)  'columnas
    Col = Col + 1
    'If Col = 2 Then
    '   Col = 3
    'ElseIf Col = 3 Then
    '   Col = 2
    'End If
    xlHoja1.Cells(Fila, Col) = Format(lsTotalesAnexo2(j), "#,#0.00")
    'If Col = 3 Then
    '   Col = 2
    'ElseIf Col = 2 Then
    '   Col = 3
    'End If
Next j
oBarra.Progress 4, "ENCAJE BCR", "Generando Reporte", , vbBlue
For i = FilaDetalle - 1 To Fila
    For j = 1 To Col
        If j = 1 Then
            xlHoja1.Range(xlHoja1.Cells(i, j), xlHoja1.Cells(i, j)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        End If
        xlHoja1.Range(xlHoja1.Cells(i, j), xlHoja1.Cells(i, j)).Borders(xlEdgeRight).LineStyle = xlContinuous
        If i = Fila Then
            xlHoja1.Range(xlHoja1.Cells(i, j), xlHoja1.Cells(i, j)).Borders.LineStyle = xlContinuous
        End If
    Next j
Next i
oBarra.Progress 5, "ENCAJE BCR", "Generando Reporte", , vbBlue
ProgressClose oBarra, Me, True
Set oBarra = Nothing

End Function

Private Function ImprimeLineaCreditoExterior(psOpeCod As String, pdFecIni As Date, pdFecFin As Date, lsAnexo As String) As String
Dim fs As New Scripting.FileSystemObject
Dim i, j, Col, Fila As Integer
Dim lnTotalDias As Long
Dim lnTotalCols As Long
lnTotalDias = DateDiff("d", pdFecIni, pdFecFin)
lnTotalDias = lnTotalDias + 3
lnTotalCols = 8

ExcelAddHoja "Anx_" & Right(psOpeCod, 1), xlLibro, xlHoja1

xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.CenterHorizontally = True
xlHoja1.PageSetup.Zoom = 90

xlHoja1.Cells(1, 1) = "BANCO CENTRAL DE RESERVA DEL PERU - DEPARTAMENTO DE ENCAJE"
xlHoja1.Cells(4, 1) = "OBLIGACIONES EXONERADAS DE GUARDAR ENCAJE"
xlHoja1.Cells(6, 1) = UCase("INSTITUCION :" & gsNomCmac)
xlHoja1.Cells(8, 1) = "LINEAS DE CREDITOS Y CREDITOS PROVENIENTES DEL EXTERIOR"
xlHoja1.Cells(9, 1) = UCase("Periodo: " & EmitePeriodo(pdFecIni, pdFecFin))
Fila = 11
xlHoja1.Cells(Fila, 1) = "DIAS"
xlHoja1.Cells(Fila, 2) = "INSTITUCION FINANCIERA DEL EXTERIOR"
xlHoja1.Cells(Fila, 3) = "PAIS DE ORIGEN"
xlHoja1.Cells(Fila, 4) = "PLAZA"
xlHoja1.Cells(Fila, 5) = "MONEDA"
xlHoja1.Cells(Fila, 6) = "MONTO AUTORIZADO (Equivalente en US$)"
xlHoja1.Cells(Fila, 7) = "MONTO UTILIZADO (Equivalente en US$)"
xlHoja1.Cells(Fila, 8) = "FECHA DE VENCIMIENTO"
For j = 0 To lnTotalDias
   Fila = Fila + 1
   Col = 1
   For i = 1 To lnTotalCols
      If j = 0 Then
         If i < lnTotalCols Then
            xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Borders.LineStyle = xlContinuous
            xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Font.Size = 7
         End If
         xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).ColumnWidth = 12
         xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).RowHeight = 40
      Else
         If j = lnTotalDias Then
            If i < lnTotalCols Then
                xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Borders.LineStyle = xlContinuous
            End If
         End If
         If i < lnTotalCols Then
             xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Borders(xlEdgeRight).LineStyle = xlContinuous
         End If
      End If
      If Col = 1 Then
          xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).HorizontalAlignment = xlCenter
          xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).Borders(xlEdgeLeft).LineStyle = xlContinuous
          xlHoja1.Range(xlHoja1.Cells(Fila, Col), xlHoja1.Cells(Fila, Col)).ColumnWidth = 5
      End If
      Col = Col + 1
    Next i
Next j
End Function

Private Sub GeneraDatosColumnas(pdFecIni As Date, pdFecFin As Date, lsOpeCod As String, lsMatrizSaldos() As String, lsMatrizTotales() As String)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim rsT As New ADODB.Recordset
Dim ldFecha As Date
Dim Total As Currency
Dim TotalDias As Long
Dim j As Long
Dim lnColTotal As Long
Dim lsMoneda   As String
Dim m As Integer
Dim nSaldoAnt As Currency
Dim N1 As New nCajaGenImprimir

Dim ldFechaAnt As String

ldFechaAnt = DateAdd("d", -1, pdFecIni)
 
 
ldFecha = pdFecIni
Filas = 0
Col = 0
TotalColumnas = 0
TotalDias = DateDiff("d", pdFecIni, pdFecFin) + 1
j = 0

'***********CONTAMOS EL TOTAL DE COLUMNAS QUE POSEE EL REPORTE ********************
Dim oRep As New DRepCtaColumna
Dim oEst As New NEstadisticas
Set rs = oRep.CargaRepColumna(lsOpeCod, , , , , gbBitCentral)
TotalColumnas = rs.RecordCount
ReDim lsMatrizSaldos(rs.RecordCount, 0)
ReDim lsMatrizTotales(rs.RecordCount)
RSClose rs
lsMoneda = Mid(lsOpeCod, 3, 1)

'***************************Obtengo El Promedio de la Caja del Mes Pasado *************
'MIOL 20120710, SEGUN RQ11370_RQ11371_RQ11372_RQ11373_RQ12147*************
nSaldoAnt = oEst.GetCajaAnterior(ldFechaAnt, "761201", "32", gbBitCentral)
'nSaldoAnt = oEst.GetCajaAnterior(ldFechaAnt, "761201", "17", gbBitCentral)
'END MIOL ****************************************************************

'**************************************************************************************
oBarra.Max = pdFecFin - ldFecha
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.ShowForm Me

Do While ldFecha <= pdFecFin
   'oBarra.Progress J, "ENCAJE DEL BCR", "Generando datos de Obligaciones Sujetas", "Procesando " & ldFecha + J, vbBlue
    j = j + 1
    If j = 30 Then
       j = j
    End If
    Col = 0: Filas = Filas + 1    'inicializamos las columnas y aumentamos las filas
    lnColTotal = 0               'el marcador de columnas que son totales se inicializa a 1
    ReDim Preserve lsMatrizSaldos(TotalColumnas, Filas)
    Set rs = oRep.CargaRepColumna(lsOpeCod)
    Do While Not rs.EOF
        Col = Col + 1
        Total = 0
        'Select Case UCase(Trim(rs!cDescCol))
        Select Case UCase(Trim(rs!nNroCol))
        
            'Case "OBLIGACIONES INMEDIATAS"
                'Total = N1.SaldoObligInmediatas(1, lsOpeCod, "76" & lsMoneda & "201", ldFecha, pdFecFin, "")
            'Case "CHEQUES AHORRO"
            
            Case 8
                If lsOpeCod = "761201" Or lsOpeCod = "762201" Then
                    Total = oEst.GetEstadSaldoCheques(ldFecha, "Ahorros", lsMoneda, gbBitCentral)
                End If
            'Case "CHEQUES PLAZO"
            'Case "CHEQUES PLAZO"
            Case 6
                If lsOpeCod = "761201" Or lsOpeCod = "762201" Then
                    Total = oEst.GetEstadSaldoCheques(ldFecha, "Plazo", lsMoneda, gbBitCentral)
                End If
            'Case "FONDO MI VIVIENDA"
                'Total = oEst.GetEstadSaldoMiViv(gbBitCentral, ldFecha, lsMoneda)
            'Case "TOTAL CAJA PERIODO ANTERIOR"
            'MIOL 20120710 SEGUN RQ11370_RQ11371_RQ11372_RQ11373_RQ12147**********
            Case 28
'            Case 23
            'END MIOL ************************************************************
                If lsOpeCod = "761201" Then
                    Total = nSaldoAnt
                Else
                    Total = SaldoCtas(rs!nNroCol, lsOpeCod, ldFecha, pdFecFin, lnTCF, lnTCFF)
                End If
            'MIOL 20120710 SEGUN RQ11370_RQ11371_RQ11372_RQ11373_RQ12147**********
            Case 27
                If lsOpeCod = "761201" Or lsOpeCod = "762201" Then
                    Dim Col15 As Currency
                    Dim Col26 As Currency
                    Col15 = rs!nNroCol - 12
                    Col26 = rs!nNroCol - 1
                    Total = CCur(lsMatrizSaldos(Col15, Filas)) + CCur(lsMatrizSaldos(Col26, Filas))
                End If
            'END MIOL ************************************************************
            
            Case Else
                If Trim(rs!cTotal) = "1" Then
                    'MIOL 20120710 SEGUN RQ11370_RQ11371_RQ11372_RQ11373_RQ12147**********
                        Dim ColA As Currency
                        Dim ColB As Currency
                    If rs!nNroCol = "30" Then
                        ColA = rs!nNroCol - 2
                        ColB = rs!nNroCol - 1
                        Total = CCur(lsMatrizSaldos(ColA, Filas)) + CCur(lsMatrizSaldos(ColB, Filas))
                    ElseIf rs!nNroCol = "31" Then
                        ColA = rs!nNroCol - 2
                        ColB = rs!nNroCol - 1
                        Total = CCur(lsMatrizSaldos(ColA, Filas)) + CCur(lsMatrizSaldos(ColB, Filas))
                    'END MIOL ************************************************************
                    Else
                        For m = lnColTotal + 1 To rs!nNroCol - 1
                                Total = Total + CCur(lsMatrizSaldos(m, Filas))
                        Next m
                        lnColTotal = rs!nNroCol
                        If rs!nNroCol < UBound(lsMatrizTotales) And UCase(Trim(rs!cDescCol)) <> "FONDO TOTAL DE ENCAJE" Then
                            FilaEncaje = lnColTotal
                        End If
                    End If
                Else
                    'ok
                    Total = SaldoCtas(rs!nNroCol, lsOpeCod, ldFecha, pdFecFin, lnTCF, lnTCFF)
                End If
        End Select
        lsMatrizSaldos(Col, Filas) = Format(Total, "#,#0.00")
        lsMatrizTotales(Col) = Format(CCur(IIf(lsMatrizTotales(Col) = "", 0, lsMatrizTotales(Col))) + CCur(lsMatrizSaldos(Col, Filas)), "#,#0.00")
        rs.MoveNext
        'DoEvents
    Loop
    RSClose rs
    ldFecha = DateAdd("d", 1, ldFecha)
    oBarra.CloseForm Me
Loop
End Sub
'*****************************************************************
'Procedimiento para Generar datos de Reporte del Anexo 2 y 4
'*****************************************************************
Private Sub GeneraDatosSubColumna(psOpeCod As String, pdFecIni As Date, pdFecFin As Date, lnNroCol As Integer, lbCmact As Boolean)
Dim ldFecha As Date
Dim Total, TotalDias, j, m As Long
Dim i   As Integer
Dim Fil As Long
Dim ColTotal As Long
Dim Dias As Integer
Dim Col As Long
Dim z As Integer
On Error GoTo GeneraDatosSubColumanErr
Fil = 0
Columna = 0
If lnNroCol = 2 Then
    TotalColumnas = InicializaMatriz(psOpeCod, lsAnexo2Col, pdFecIni, pdFecFin, True, True)
    SubCuentasCol psOpeCod, pdFecIni, pdFecFin, lsAnexo2Col, TotalColumnas, True
Else
    TotalColumnas = InicializaMatriz(psOpeCod, lsAnexo2Col, pdFecIni, pdFecFin, False)
    SubCuentasCol psOpeCod, pdFecIni, pdFecFin, lsAnexo2Col, TotalColumnas, False
End If

'Redimensionamos la matriz al total de Columnas y SubColumnas
ReDim lsAnexo2(UBound(lsAnexo2Col, 2), 0)
ReDim lsTotalesAnexo2(UBound(lsAnexo2Col, 2))
ldFecha = pdFecIni
TotalDias = DateDiff("d", pdFecIni, pdFecFin) + 1
j = 0

ColTotal = 1
Col = 0
Dias = 0
Do While ldFecha <= pdFecFin
    Dias = Dias + 1
    j = j + 1
    Fil = Fil + 1
    Col = 0
    ColTotal = 1
    ReDim Preserve lsAnexo2(UBound(lsAnexo2Col, 2), Fil)
    lsAnexo2(0, Fil) = Fil
    For i = 1 To UBound(lsAnexo2Col, 1)
        For j = 1 To UBound(lsAnexo2Col, 2)
            Total = 0
            If Len(Trim(lsAnexo2Col(i, j, 1))) > 0 Then
                Col = Col + 1
                If lsAnexo2Col(i, j, 1) = "T" Then
                    'ColTotal = 15
                    For m = ColTotal To Col - 1
                        Total = Total + CCur(IIf(lsAnexo2(m, Fil) = "", "0", lsAnexo2(m, Fil)))
                    Next m
                    ColTotal = Col + 1
                Else
                    If i = 2 And lbCmact = True Then
                        For z = 0 To UBound(CuentasCmac) - 1
                            If Trim(CuentasCmac(z).cDescrip) = Trim(lsAnexo2Col(i, j, 2)) Then
                                Total = Total + SaldoCuentaCmacs(CuentasCmac(z).cCta, ldFecha)
                            End If
                        Next z
                    Else
                        For z = 0 To UBound(Creditos) - 1
                            If Trim(Creditos(z).cDescrip) = Trim(lsAnexo2Col(i, j, 2)) Then
                                Total = Total + SaldoCuenta(Creditos(z).cCta, Mid(psOpeCod, 3, 1), ldFecha, IIf(ldFecha = pdFecFin, lnTCFF, lnTCF))
                            End If
                        Next z
                    End If
                End If
                lsAnexo2(Col, Fil) = Format(Total, "#,#0.00")
                lsTotalesAnexo2(Col) = Format(CCur(IIf(lsTotalesAnexo2(Col) = "", "0", lsTotalesAnexo2(Col))) + CCur(IIf(lsAnexo2(Col, Fil) = "", "0", lsAnexo2(Col, Fil))), "#,#0.00")
            End If
'            DoEvents
        Next j
    Next i
    ldFecha = DateAdd("d", 1, ldFecha)
'    DoEvents
Loop
Exit Sub
GeneraDatosSubColumanErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
'ALPA 20080827**************
'***Se modifico el Reporte
Private Function SaldoCtas(lnNroCol As Long, lsOpeCod As String, ByVal ldFecha As Date, ByVal pdFecFin As Date, ByVal pnTCF As Currency, ByVal pnTCFF As Currency) As Currency
Dim rsC As New ADODB.Recordset
Dim oRep As New DRepCtaColumna
Set rsC = oRep.GetRepColumnaCtaSaldo(lsOpeCod, lnNroCol, ldFecha, gbBitCentral)
SaldoCtas = 0
If Not RSVacio(rsC) Then
    If Mid(lsOpeCod, 3, 1) = "1" Then
        SaldoCtas = IIf(IsNull(rsC!Total), 0, rsC!Total)
    Else
        'MIOL 20120713 SEGUN RQ11370_RQ11371_RQ11372_RQ11373_RQ12147********
        If lnNroCol = 29 Then
        'If lnNroCol = 18 Then
        'END MIOL **********************************************************
            SaldoCtas = Round(IIf(IsNull(rsC!TotalME), 0, rsC!TotalME), 1)
        Else
            If pdFecFin = ldFecha Then
               SaldoCtas = IIf(IsNull(rsC!Total), 0, Round(rsC!Total / pnTCFF, 2))
            Else
               SaldoCtas = IIf(IsNull(rsC!Total), 0, Round(rsC!Total / pnTCF, 2))
            End If
        End If
    End If
End If
Set oRep = Nothing
RSClose rsC
End Function
 'ALPA 20080810*************************************************************
Private Function devCelda(ByVal nNPos As Integer) As String
Dim sCellda As String
If nNPos = 1 Then
sCellda = "B"
ElseIf nNPos = 2 Then
sCellda = "C"
ElseIf nNPos = 3 Then
sCellda = "D"
ElseIf nNPos = 4 Then
sCellda = "E"
ElseIf nNPos = 5 Then
sCellda = "F"
ElseIf nNPos = 6 Then
sCellda = "G"
ElseIf nNPos = 7 Then
sCellda = "H"
ElseIf nNPos = 8 Then
sCellda = "I"
ElseIf nNPos = 9 Then
sCellda = "J"
ElseIf nNPos = 10 Then
sCellda = "K"
ElseIf nNPos = 11 Then
sCellda = "L"
ElseIf nNPos = 12 Then
sCellda = "M"
ElseIf nNPos = 13 Then
sCellda = "N"
ElseIf nNPos = 14 Then
sCellda = "O"
ElseIf nNPos = 15 Then
sCellda = "P"
ElseIf nNPos = 16 Then
sCellda = "Q"
ElseIf nNPos = 17 Then
sCellda = "R"
ElseIf nNPos = 18 Then
sCellda = "S"
ElseIf nNPos = 19 Then
sCellda = "T"
ElseIf nNPos = 20 Then
sCellda = "U"
ElseIf nNPos = 21 Then
sCellda = "V"
ElseIf nNPos = 22 Then
sCellda = "W"
ElseIf nNPos = 23 Then
sCellda = "X"
ElseIf nNPos = 24 Then
sCellda = "Y"
ElseIf nNPos = 25 Then
sCellda = "Z"
ElseIf nNPos = 26 Then
sCellda = "AA"
ElseIf nNPos = 27 Then
sCellda = "AB"
ElseIf nNPos = 28 Then
sCellda = "AC"
ElseIf nNPos = 29 Then
sCellda = "AD"
ElseIf nNPos = 30 Then
sCellda = "AE"
ElseIf nNPos = 31 Then
sCellda = "AF"
ElseIf nNPos = 32 Then
sCellda = "AG"
ElseIf nNPos = 33 Then
sCellda = "AH"
End If
devCelda = sCellda
End Function
Private Sub CabeceraAnexo1(lsMoneda As String, pdFecIni As Date, pdFecFin As Date, Optional gsCodOpe As String = "761201")
Dim i, j As Integer
Dim lnCol As Integer
Dim Col As Integer
xlAplicacion.Range("A1:R100").Font.Size = 10

xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.CenterHorizontally = True
xlHoja1.PageSetup.Zoom = 55


xlHoja1.Cells(1, 1) = "BANCO CENTRAL DE RESERVA DEL PERU-DEPARTAMENTO DE ENCAJE"
xlHoja1.Cells(2, 1) = lsCadenaMoneda
If lsMoneda = "1" Then
    xlHoja1.Cells(6, 1) = UCase("INSTITUCION : " & gsNomCmac)
    xlHoja1.Cells(7, 1) = UCase("Periodo: " & EmitePeriodo(pdFecIni, pdFecFin))
Else
    xlHoja1.Cells(5, 1) = UCase("INSTITUCION : " & gsNomCmac)
    xlHoja1.Cells(6, 1) = UCase("Periodo: " & EmitePeriodo(pdFecIni, pdFecFin))
End If
lnCol = UBound(lsAnexo1, 1)
Col = 1
For i = 1 To lnCol
    Col = Col + 1
    xlHoja1.Cells(11, Col) = "[" & i & "]"
Next i

xlHoja1.Range("A1:A1").ColumnWidth = 5
'MIOL 20120713 SEGUN RQ11370_RQ11371_RQ11372_RQ11373_RQ12147 *****
If lsMoneda = "1" Then
    xlHoja1.Range("B1:AG1").ColumnWidth = 13
    'xlHoja1.Range("B1:Z1").ColumnWidth = 13
Else
    xlHoja1.Range("B1:AH1").ColumnWidth = 13
End If
'END MIOL *********************************************************************

    Dim rs As New ADODB.Recordset
    Dim oRep As New DRepCtaColumna
    Set rs = oRep.CargaRepColumna(gsCodOpe, , , , , gbBitCentral)
    nCantE = rs.RecordCount
If lsMoneda = "1" Then
    xlHoja1.Cells(10, 1) = "OBLIGACIONES SUJETAS A ENCAJE MONEDA NACIONAL"
    xlHoja1.Cells(10, 29) = "FONDO DE ENCAJE" 'MIOL 20120712 29 reemplaza por 18

    xlHoja1.Range(xlHoja1.Cells(10, 1), xlHoja1.Cells(10, Col - 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(10, Col - 4), xlHoja1.Cells(10, Col)).Merge True

    ''''''''''
    xlAplicacion.Range(xlHoja1.Cells(12, 1), xlHoja1.Cells(13, Col)).Font.Size = 6
    xlAplicacion.Range(xlHoja1.Cells(10, 1), xlHoja1.Cells(13, Col)).HorizontalAlignment = xlHAlignCenter
    ''''''''

    ExcelCuadro xlHoja1, 1, 10, Col - 5, 14, True
    'Dim Columna As Integer
    ExcelCuadro xlHoja1, Col - 5, 10, CCur(Col), 14, True
    xlHoja1.Cells(12, 1) = "DIAS":
    Do While Not rs.EOF
        Columna = Columna + 1
        xlHoja1.Cells(12, rs!nNroCol + 1) = rs!cDescCol
        'xlHoja1.Range(devCelda(rs!nNroCol) & "(11) "B11:B11").EntireColumn.HorizontalAlignment = xlHAlignCenter
        xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol) & "14").Select
        xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol) & "14").MergeCells = True
        xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol) & "14").EntireColumn.HorizontalAlignment = xlHAlignJustify
        xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol) & "14").EntireColumn.VerticalAlignment = xlVAlignTop
        rs.MoveNext
    Loop
    'ALPA 20080810*************************************************************
    'rc.cDescCol
    RSClose rs
'    xlHoja1.Cells(12, 1) = "DIAS":
'    xlHoja1.Cells(12, 2) = "OBLIGACIONES":          xlHoja1.Cells(13, 2) = "INMEDIATAS"
'    xlHoja1.Cells(12, 3) = "A PLAZO HASTA":         xlHoja1.Cells(13, 3) = "30 DIAS"
'    xlHoja1.Cells(12, 4) = "CHEQUES A":             xlHoja1.Cells(13, 4) = "DEDUCIR"
'    xlHoja1.Cells(12, 5) = "A PLAZO MAYOR":         xlHoja1.Cells(13, 5) = "DE 30 DIAS"
'    xlHoja1.Cells(12, 6) = "DEUDA":                 xlHoja1.Cells(13, 6) = "SUBORDINADA"
'    xlHoja1.Cells(12, 7) = "OTRAS OBLIG.":          xlHoja1.Cells(13, 7) = "A PLAZO":
'    xlHoja1.Cells(12, 8) = "CHEQUES A":             xlHoja1.Cells(13, 8) = "DEDUCIR"
'    xlHoja1.Cells(12, 9) = "OTRAS OBLIG. A PLAZO":  xlHoja1.Cells(13, 9) = "SUJETAS A REAJ/VAC"
'    xlHoja1.Cells(12, 10) = "SWAP Y DEPOSITOS":     xlHoja1.Cells(13, 10) = "COMPRA FUTURO M/S"
'    xlHoja1.Cells(12, 11) = "OBLIG. EN FUNC.":      xlHoja1.Cells(13, 11) = "VARIAC.TC ME"
'    xlHoja1.Cells(12, 12) = "CHEQUES A ":           xlHoja1.Cells(13, 12) = "DEDUCIR"
'    xlHoja1.Cells(12, 13) = "AHORROS":              xlHoja1.Cells(13, 13) = " "
'
'    xlHoja1.Cells(12, 14) = "CHEQUES A":            xlHoja1.Cells(13, 14) = "DEDUCIR"
'    xlHoja1.Cells(12, 15) = "OBLIGACIONES POR":            xlHoja1.Cells(13, 15) = "COMISION DE COBRANZA"
'    xlHoja1.Cells(12, 16) = "OBLIGACIONES":            xlHoja1.Cells(13, 16) = "CON ENT.FIN.EXT"
'    xlHoja1.Cells(12, 17) = "TOTAL":                xlHoja1.Cells(13, 17) = " ":
'    xlHoja1.Cells(12, 18) = "TOTAL CAJA":            xlHoja1.Cells(13, 18) = "PERIODO ANTERIOR":
'    xlHoja1.Cells(12, 19) = "DEPOSITOS":                 xlHoja1.Cells(13, 19) = "EN EL BCR":
'    xlHoja1.Cells(12, 20) = "FONDO TOTAL":         xlHoja1.Cells(13, 20) = "DE ENCAJE"
'    xlHoja1.Cells(12, 21) = "PRESTAMOS DE CAJA":         xlHoja1.Cells(13, 21) = "PERIODO REPORTADO"
'    xlHoja1.Cells(12, 22) = "TOTAL CAJA":         xlHoja1.Cells(13, 22) = "PERIODO REPORTADO"
    
    
    ColumnaFinal = nCantE
Else
 'ALPA 20080810*************************************************************
    xlHoja1.Cells(10, 1) = "OBLIGACIONES SUJETAS A ENCAJE MONEDA EXTRANJERA"
    xlHoja1.Cells(10, 29) = "FONDO DE ENCAJE" 'MIOL 20120712 28 reemplaza por 15

    xlHoja1.Range(xlHoja1.Cells(10, 1), xlHoja1.Cells(10, Col - 6)).Merge True 'MIOL 20120712 6 reemplaza por 5
    xlHoja1.Range(xlHoja1.Cells(10, Col - 5), xlHoja1.Cells(10, Col)).Merge True

    xlAplicacion.Range(xlHoja1.Cells(12, 1), xlHoja1.Cells(13, Col)).Font.Size = 6
    xlAplicacion.Range(xlHoja1.Cells(10, 1), xlHoja1.Cells(13, Col)).HorizontalAlignment = xlHAlignCenter

    ExcelCuadro xlHoja1, 1, 10, Col - 4, 14, True

    ExcelCuadro xlHoja1, Col - 4, 10, CCur(Col), 14, True
'''    ALPA 20080801
    xlHoja1.Cells(12, 1) = "DIAS":
     Do While Not rs.EOF
        Columna = Columna + 1
        xlHoja1.Cells(12, rs!nNroCol + 1) = rs!cDescCol
        'xlHoja1.Range(devCelda(rs!nNroCol) & "(11) "B11:B11").EntireColumn.HorizontalAlignment = xlHAlignCenter
        xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol) & "14").Select
        xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol) & "14").MergeCells = True
        xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol) & "14").EntireColumn.HorizontalAlignment = xlHAlignJustify
        xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol) & "14").EntireColumn.VerticalAlignment = xlVAlignTop
        rs.MoveNext
    Loop
    RSClose rs
'''    xlHoja1.Cells(12, 2) = "OBLIGACIONES":          xlHoja1.Cells(13, 2) = "INMEDIATAS"
'''    xlHoja1.Cells(12, 3) = "A PLAZO HASTA":         xlHoja1.Cells(13, 3) = "30 DIAS"
'''    xlHoja1.Cells(12, 4) = "CHEQUES A":             xlHoja1.Cells(13, 4) = "DEDUCIR"
'''    xlHoja1.Cells(12, 5) = "A PLAZO MAYOR":         xlHoja1.Cells(13, 5) = "DE 30 DIAS"
'''
'''    xlHoja1.Cells(12, 6) = "VALORES COLOC":         xlHoja1.Cells(13, 6) = "EN EL EXTERIOR"
'''    xlHoja1.Cells(12, 7) = "CHEQUES A":             xlHoja1.Cells(13, 7) = "DEDUCIR PLAZO"
'''    xlHoja1.Cells(12, 8) = "AHORROS":               xlHoja1.Cells(13, 8) = " "
'''    xlHoja1.Cells(12, 9) = "CHEQUES A":             xlHoja1.Cells(13, 9) = "DEDUCIR AHORROS"
'''    xlHoja1.Cells(12, 10) = "OBLIGACIONES POR":     xlHoja1.Cells(13, 10) = "COMISIONES DE CONF."
'''    xlHoja1.Cells(12, 11) = "DEPOSITOS Y OTRAS":    xlHoja1.Cells(13, 11) = "OBLIGA. PROV, EXTER."
'''    xlHoja1.Cells(12, 12) = "OBLIGA.DERI.":         xlHoja1.Cells(13, 12) = "DE CREDITOS EXTERNOS"
'''    xlHoja1.Cells(12, 13) = "OBLIGA.CON.ENT.":      xlHoja1.Cells(13, 13) = "Y ORG. FIN DEL EXTER."
'''    xlHoja1.Cells(12, 14) = "OBLIGA.CON.SUC.":      xlHoja1.Cells(13, 14) = "DEL EXTERIOR"
'''    xlHoja1.Cells(12, 15) = "TOTAL":                xlHoja1.Cells(13, 15) = ""
'''
'''    xlHoja1.Cells(12, 16) = "PRESTAMOS":            xlHoja1.Cells(13, 16) = "DE CAJA"
'''    xlHoja1.Cells(12, 17) = "TOTAL":                xlHoja1.Cells(13, 17) = "CAJA"
'''    xlHoja1.Cells(12, 18) = "DEPOSITOS EN":         xlHoja1.Cells(13, 18) = "EL BCRP"
'''    xlHoja1.Cells(12, 19) = "TOTAL FONDOS":         xlHoja1.Cells(13, 19) = "DE ENCAJE"
'''    xlHoja1.Cells(12, 20) = "TOSE":                xlHoja1.Cells(13, 20) = "EUROS 12"
'''    xlHoja1.Cells(12, 21) = "TOSE":                xlHoja1.Cells(13, 21) = "EUROS"
'''    'xlHoja1.Cells(12, 6) = "DEUDA":                 xlHoja1.Cells(13, 6) = "SUBORDINADA"
'''    'xlHoja1.Cells(12, 7) = "OTRAS OBLIG. ":         xlHoja1.Cells(13, 7) = "A PLAZO"
    'xlHoja1.Cells(12, 8) = "BONOS COLOC.":          xlHoja1.Cells(13, 8) = "EN EL EXTERIOR":
    'xlHoja1.Cells(12, 9) = "CHEQUES A ":           xlHoja1.Cells(13, 9) = "DEDUCIR"
    'xlHoja1.Cells(12, 10) = "AHORROS":              xlHoja1.Cells(13, 10) = " "
    'xlHoja1.Cells(12, 11) = "CHEQUES A":            xlHoja1.Cells(13, 11) = "DEDUCIR"
    'xlHoja1.Cells(12, 12) = "OBLIGACIONES POR":     xlHoja1.Cells(13, 12) = "COMISIONES DE CONFIANZA"
    'xlHoja1.Cells(12, 13) = "OBLIGACIONES CON":     xlHoja1.Cells(13, 13) = "ENT.FIN.DEL EXTERIOR"
    
    'xlHoja1.Cells(12, 14) = "TOTAL":                xlHoja1.Cells(13, 14) = " ":
    'xlHoja1.Cells(12, 15) = "PRESTAMOS ":           xlHoja1.Cells(13, 15) = "DE CAJA":
    'xlHoja1.Cells(12, 16) = "TOTAL ":               xlHoja1.Cells(13, 16) = "CAJA ":
    'xlHoja1.Cells(12, 17) = "DEPOSITOS EN":         xlHoja1.Cells(13, 17) = "EL B.C.R"
    'xlHoja1.Cells(12, 18) = "TOTAL FONDOS":         xlHoja1.Cells(13, 18) = "DE ENCAJE"
    ColumnaFinal = nCantE
End If
RSClose rs
End Sub
 'ALPA 20080828****************************************************************
 'Se agrego el parametro lbCoop
 '*****************************************************************************
Private Function InicializaMatriz(lsOpeCod As String, lsMatriz() As String, pdFecIni As Date, pdFecFin As Date, Optional lbCmacs As Boolean = True, Optional lbCoop As Boolean = False) As Long
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim sql As String
Dim oRep As New DRepCtaColumna
Set rs = oRep.CargaRepColumna(lsOpeCod)
TotalCol = rs.RecordCount
RSClose rs
SubCol = 0

For i = 1 To TotalCol
   If Not (lbCmacs = True And i = 2) Then
      Set rs = oRep.CargaRepColumnaCta(lsOpeCod, i, , True)
      SubCol = SubCol + IIf(rs.RecordCount < 3, 4, rs.RecordCount + 1)
      rs.Close
   End If
Next

If lbCmacs = True Then
    Dim oCon As New DConecta
    Dim lsServConsol As String
    Dim lsCadena     As String
    
    If Not oCon.AbreConexion Then 'Remota(gsCodAge, True, , "03") Then
        Exit Function
    End If
    
    'Obtengo el servidor remoto
    Set rs = oCon.CargaRecordSet("select nconssisvalor from constsistema where nconssiscod=" & gConstSistServCentralRiesgos)
    If rs.BOF Then
    Else
        lsServConsol = rs!nConsSisValor
    End If
    RSClose rs

    
    'lsCadena = oCon.GetCadenaConexion(gsCodAge, "03")
    'lsServConsol = "[" & Mid(lsCadena, InStr(lsCadena, "SERVER=") + 7, 20) & "]."
    'lsServConsol = lsServConsol & Mid(lsCadena, InStr(lsCadena, "DATABASE=") + 9, InStr(lsCadena, "SERVER") - InStr(lsCadena, "DATABASE=") - 1) & ".dbo."
       
    'lsServConsol = "[128.107.2.102].DBConsolidada.dbo."
    
    'oCon.AbreConexion
        
   ' select rtrim(P.cPersNombre) as cNomPers, PPC.cCtaCod as cCodCta from DBCMACTAUX..persona P inner join productopersonaconsol PPC
   ' on P.cPersCod = PPC.cPersCod where P.nperspersoneria= 4
   ' GROUP BY PPC.cCtaCod, P.cPersNombre
    
        
    sql = "DELETE CtasCmacs"
    oCon.Ejecutar sql
 
    If gbBitCentral = True Then
    
        sql = " INSERT INTO CtasCmacs "
        sql = sql & " Select T.cPersNombre cObjetoDesc, T.cCtaCod From ( "
        sql = sql & " Select P.cPersNombre, PC.cCtaCod From Persona P INNER JOIN " & lsServConsol & "productopersonaconsol PC "
        sql = sql & " INNER JOIN " & lsServConsol & "AhorroCConsol A ON PC.cCtaCod = A.cCtaCod ON P.cPersCod = PC.cPersCod "
        sql = sql & " Where A.nPersoneria IN (4,5,7) And A.nEstCtaAC NOT IN (1400, 1300) "
        'ALPA 20080828*****************************
        If lbCoop = True And lsOpeCod = "761202" Then
            sql = sql & "and PC.cCtaCod not in ('109022321000055069','109012321000467294','109022321000001864','109032321000002445','109042321000004502','109062321000031348','109072321000006629','109092321000010243') "
        End If
        '******************************************
        sql = sql & " And PC.nPrdPersRelac = 10 Union "
        sql = sql & " Select P.cPersNombre, PC.cCtaCod From Persona P INNER JOIN " & lsServConsol & "productopersonaconsol PC "
        sql = sql & " INNER JOIN " & lsServConsol & "PlazoFijoConsol A ON PC.cCtaCod = A.cCtaCod ON P.cPersCod = PC.cPersCod "
        sql = sql & " Where A.nPersoneria IN (4,5) And A.nEstCtaPF NOT IN (1400, 1300) "
        sql = sql & " And PC.nPrdPersRelac =  10 ) T "
        sql = sql & " Group by  T.cPersNombre, T.cCtaCod Order by T.cPersNombre,T.cCtaCod "
        oCon.Ejecutar sql
             
        sql = "Select cPersNombre From CtasCmacs Where substring(cCtaCod,9,1) = '" & Mid(lsOpeCod, 3, 1) & "' Group by cPersNombre Order by cPersNombre "
     
    Else
     
        sql = "INSERT INTO CtasCmacs "
        sql = sql & " Select T.cNomPers cObjetoDesc, T.cCodCta From ( "
        sql = sql & " Select P.cNomPers, PC.cCodCta From DBPersona.dbo.Persona P INNER JOIN " & lsServConsol & "PersCuentaConsol PC INNER JOIN "
        sql = sql & lsServConsol & "AhorroCConsol A ON PC.cCodCta = A.cCodCta ON P.cCodPers = PC.cCodPers "
        sql = sql & " Where A.cPersoneria IN ('4','5') And A.cEstCtaAC NOT IN ('C','U') And PC.cRelaCta = 'TI' "
        sql = sql & " Union "
        sql = sql & " Select P.cNomPers, PC.cCodCta From DBPersona.dbo.Persona P INNER JOIN " & lsServConsol & "PersCuentaConsol PC INNER JOIN "
        sql = sql & lsServConsol & "PlazoFijoConsol A ON PC.cCodCta = A.cCodCta ON P.cCodPers = PC.cCodPers "
        sql = sql & " Where A.cPersoneria IN ('4','5') And A.cEstCtaPF NOT IN ('C','U') And PC.cRelaCta = 'TI' "
        sql = sql & "  ) T Group by  T.cNomPers, T.cCodCta Order by T.cNomPers,T.cCodCta "
        oCon.Ejecutar sql
             
        sql = "Select cNomPers From CtasCmacs Where substring(ccodcta,6,1) = '" & Mid(lsOpeCod, 3, 1) & "' Group by cNompers Order by cNomPers "
    
    End If
    Set rs = oCon.CargaRecordSet(sql)
    SubCol = SubCol + rs.RecordCount + 1
    RSClose rs
   Dim oIF As New DCajaCtasIF
   'Set rs = oIF.CargaCtasIF(gMonedaNacional, , MuestraInstituciones, , gTpoIFCmac)
   Set rs = oIF.GetCuentasCaptacionesCmacs(lsOpeCod, gbBitCentral, True)
   SubCol = SubCol + rs.RecordCount + 1
   RSClose rs
End If
ReDim lsMatriz(TotalCol, SubCol, 2)
InicializaMatriz = TotalCol
End Function

Private Sub SubCuentasCol(psOpeCod As String, pdFecIni As Date, pdFecFin As Date, lsMatriz() As String, lnNumColumnas As Long, Optional lbCmac As Boolean = False)
Dim i As Integer, j As Integer
Dim sql As String
Dim rs As New ADODB.Recordset
Dim TotalSub As Integer
Dim lnNroCol As Integer
Dim oRep As New DRepCtaColumna

'Determinamos las Cuentas de "Créditos que se han movido en el periodo
'para determinar el Numero de Columnas de Creditos
ReDim Creditos(0)
For lnNroCol = 1 To lnNumColumnas
   If lbCmac = True And lnNroCol = 2 Then
       SubCuentasCmac psOpeCod, pdFecIni, Format(pdFecFin, "dd/mm/yyyy"), 2, lsMatriz
   Else
      Set rs = oRep.GetRepColumnaCtaDesc(psOpeCod, lnNroCol, pdFecIni, True)
      Columna = 0
      TotalSub = rs.RecordCount
      If Not RSVacio(rs) Then
         lsMatriz(lnNroCol, Columna, 1) = UCase(Trim(rs!cDescCol))
         Do While Not rs.EOF
             Columna = Columna + 1
             lsMatriz(lnNroCol, Columna, 1) = "Credito"
             lsMatriz(lnNroCol, Columna, 2) = Trim(rs!cCtaContDesc)
             rs.MoveNext
         Loop
      
   
   '************************************************
   '*** Carga los Datos de Creditos Ctas y Descrip
   '************************************************
         rs.Close
         'MsgBox "Pepe"
         Set rs = oRep.GetRepColumnaCtaDesc(psOpeCod, lnNroCol, pdFecIni)
         ReDim Preserve Creditos(rs.RecordCount)
         Do While Not rs.EOF
             Creditos(rs.Bookmark - 1).cCta = Trim(rs!cCtaContCod)
             Creditos(rs.Bookmark - 1).cDescrip = Trim(rs!cCtaContDesc)
             rs.MoveNext
         Loop
      End If
      RSClose rs
      
      If TotalSub < 3 Then
         If TotalSub = 0 Then
            Set rs = oRep.CargaRepColumna(psOpeCod, lnNroCol)
            
            lsMatriz(lnNroCol, Columna, 1) = rs!cDescCol
            RSClose rs
         End If
        ' For I = TotalSub + 1 To 3
        For i = TotalSub + 1 To 2
            Columna = Columna + 1
            lsMatriz(lnNroCol, Columna, 1) = "I"
            lsMatriz(lnNroCol, Columna, 2) = "INSTITUCION"
         Next
      End If
      Columna = Columna + 1
      lsMatriz(lnNroCol, Columna, 1) = "T"
      lsMatriz(lnNroCol, Columna, 2) = "TOTAL"
   End If
Continuar:
Next
End Sub

Private Sub SubCuentasCmac(lsOpeCod As String, pdFecIni As Date, pdFecFin As String, lnNroCol As Long, lsMatriz() As String)
Dim rs As New ADODB.Recordset
Dim TotalSubColumn As Long
Dim ldFecha As Date
Dim rsSaldos As New ADODB.Recordset
Dim TotalCajas As String

Dim oRep As New DRepCtaColumna
Set rs = oRep.GetRepColumnaCmacs(CInt(Mid(lsOpeCod, 3, 1)), gbBitCentral)
TotalSubColumn = rs.RecordCount
Columna = 0
If Not RSVacio(rs) Then
    lsMatriz(lnNroCol, Columna, 1) = "DEPOSITOS"
    Do While Not rs.EOF
        Columna = Columna + 1
        lsMatriz(lnNroCol, Columna, 1) = "Cta"
        lsMatriz(lnNroCol, Columna, 2) = Trim(rs!cObjetoDesc)
        rs.MoveNext
    Loop
    Columna = Columna + 1
    lsMatriz(lnNroCol, Columna, 1) = "T"
    lsMatriz(lnNroCol, Columna, 2) = "TOTAL"
   rs.MoveFirst
End If

'*****************************************************
'********* Carga las Cuentas de Ahorros de Cada Cmact
'*****************************************************
Set rs = oRep.GetRepColumnaCmacsCuentas(CInt(Mid(lsOpeCod, 3, 1)), gbBitCentral)
If Not RSVacio(rs) Then
    ReDim CuentasCmac(rs.RecordCount)
    Do While Not rs.EOF
        CuentasCmac(rs.Bookmark - 1).cCta = Trim(rs!cCodCta)
        CuentasCmac(rs.Bookmark - 1).cDescrip = Trim(rs!cNomPers)
        rs.MoveNext
    Loop
End If
RSClose rs

End Sub
Private Function SaldoCuenta(lsCodCtaCont As String, psMoneda As String, ldFecha As Date, pnTipCambio As Currency) As Currency
Dim oSdo   As New NCtasaldo
Dim nSaldo As Currency
nSaldo = oSdo.GetCtaSaldo(lsCodCtaCont, Format(ldFecha, gsFormatoFecha), True)
    If psMoneda = "1" Then
        SaldoCuenta = nSaldo
    Else
        SaldoCuenta = Round(nSaldo / pnTipCambio, 2)
    End If

End Function

Private Function SaldoCuentaCmacs(lsCodCta As String, ldFecha As Date) As Currency
Dim sql As String
Dim prs  As ADODB.Recordset
Dim oConec As New DConecta
Dim cCodOpe As String
    SaldoCuentaCmacs = 0
   
    If gbBitCentral = True Then
   
        oConec.AbreConexion
  
'        sql = "SELECT  nSaldoContable "
'        sql = sql & " FROM MovCap Mc "
'        sql = sql & " WHERE   nMovNro = ( "
'        sql = sql & "                       SELECT   MAX(M.nMovNro) AS nMovNro "
'        sql = sql & "                       FROM    MovCap TA JOIN Mov M ON M.nMovNro = TA.nMovNro "
'        sql = sql & "                       Where TA.cCtaCod = mc.cCtaCod "
'        sql = sql & "                           And (M.nMovFlag IS NULL OR M.nMovFlag=" & gMovFlagVigente & ") "
'        sql = sql & "                           And LEFT(M.cMovNro,8) <= ''" & Format(ldFecha, "YYYYMMdd") & "'' "
'        sql = sql & "                           AND TA.cCtaCod = '" & lsCodCta & "' "
'        sql = sql & "                   ) "


        sql = "select nSaldCnt "
        sql = sql & " From capsaldosdiarios "
        sql = sql & " where datediff(day,dfecha,'" & Format(ldFecha, "MM/dd/YYYY") & "') = 0 and cCtaCod='" & lsCodCta & "' " 'and bInactiva=0"
   
    Else
        If Left(lsCodCta, 2) = "08" Then
            Exit Function
        End If
        If Not oConec.AbreConexion Then 'Remota(lsCodCta, False)
            MsgBox "El servidor de la Agencia Nº " & Left(lsCodCta, 2) & " no responde." & Chr(10) & "Consultar con Sistemas", vbInformation, "¡Aviso!"
            Exit Function
        End If
       
        sql = " SELECT  dbo.GetSaldoCapSaldosDiarios('" & lsCodCta & "', '" & Format(ldFecha, "mm/dd/yyyy") & "', 0) As nSaldCnt "
  
    End If
    
    Set prs = oConec.CargaRecordSet(sql)
    If Not RSVacio(prs) Then
        SaldoCuentaCmacs = prs!nSaldCnt
    End If
    RSClose prs
    oConec.CierraConexion
    Set oConec = Nothing
End Function


'Private Function SaldoCuentaCmacs(lsCodCta As String, ldFecha As Date) As Currency
'Dim sql As String
'Dim prs  As ADODB.Recordset
'Dim oConec As New DConecta
'Dim cCodOpe As String
'   SaldoCuentaCmacs = 0
'   If Left(lsCodCta, 2) = "08" Then
'    Exit Function
'   End If
'   If Not oConec.AbreConexionRemota(lsCodCta, False) Then
'     MsgBox "El servidor de la Agencia Nº " & Left(lsCodCta, 2) & " no responde." & Chr(10) & "Consultar con Sistemas", vbInformation, "¡Aviso!"
'    Exit Function
'   End If
'
'   'sql = "SELECT  nSaldoContable FROM MovCap Mc " _
'   '    & "WHERE   nMovNro = (  SELECT   MAX(nMovNro) AS nMovNro " _
'   '                          & "FROM    MovCap TA JOIN Mov M ON M.nMovNro = TA.nMovNro " _
'   '                          & "WHERE   TA.cCtaCod = mc.cCtaCod " _
'   '                          & "    And (M.cFlag IS NULL OR M.cFlag = " & gMovFlagVigente & ") " _
'   '                          & "    And LEFT(TA.cMovNro,8) <= '" & Format(ldFecha, gsFormatoMovFecha) & "' " _
'   '    & "    AND cCodCta = '" & lsCodCta & "'"
'
'   sql = " Update trandiariaconsol"
'   sql = sql & " Set nSaldCnt = 0.00 "
'   sql = sql & " where ccodcta = '" & lsCodCta & "'"
'   sql = sql & " and ccodope in ('210403','210301','210302','210401','210304','210305','210902','210903')"
'   sql = sql & " and nSaldCnt > 0 and ccodcta in (select ccodcta from plazofijo where cEstCtaPF = 'C' and ccodcta = '" & lsCodCta & "')"
'   oConec.ConexionActiva.Execute sql
'
'
'   sql = " Update TransAho"
'   sql = sql & " Set nSaldCnt = 0.00 "
'   sql = sql & " where ccodcta = '" & lsCodCta & "'"
'   sql = sql & " and ccodope in ('210403','210301','210302','210401','210304','210305','210902','210903')"
'   sql = sql & " and nSaldCnt > 0 and ccodcta in (select ccodcta from plazofijo where cEstCtaPF = 'C' and ccodcta = '" & lsCodCta & "')"
'   oConec.ConexionActiva.Execute sql
'
'
'   sql = "SELECT  cCodOpe, dFecTran, nSaldCnt FROM TransAho " _
'            & "Where nNumTran IN (   SELECT    MAX(TA.nNumTran) AS Fecha " _
'                                    & "FROM    TRANSAHO TA " _
'                                    & "WHERE   TA.cCodCta = '" & lsCodCta & "' " _
'                                    & "And     (TA.cFlag IS NULL OR SUBSTRING(TA.cFlag ,1,1) <> 'X') " _
'                                    & "AND     TA.dFectran<='" & Format(ldFecha, "mm/dd/yyyy") & " 23:59:59' And TA.cCodOpe not in ('260105','260106')) " _
'            & "AND CCODCTA='" & lsCodCta & "'"
'
'   Set prs = oConec.CargaRecordSet(sql)
'   If Not RSVacio(prs) Then
'       SaldoCuentaCmacs = prs!nSaldCnt
'   End If
'   RSClose prs
'   oConec.CierraConexion
'   Set oConec = Nothing
'End Function

Private Sub CabeceraAnexo2(lsMoneda As String, pdFecIni As Date, pdFecFin As Date)
Dim i, j As Integer
Dim Col As Long
Dim ColIni As Long
Dim TotalCol As Long
Dim nPosicion As Integer
xlAplicacion.Range("A1:R100").Font.Name = "Times New Roman"
xlAplicacion.Range("A1:R100").Font.Size = 12

xlHoja1.PageSetup.Orientation = xlLandscape

xlHoja1.PageSetup.CenterHorizontally = True

xlHoja1.PageSetup.Zoom = 60

xlHoja1.Range("A1:A1").ColumnWidth = 5
xlHoja1.Range("B1:Z1").ColumnWidth = 12

xlAplicacion.Range("A10:R14").HorizontalAlignment = xlHAlignCenter


xlHoja1.Cells(1, 1) = "BANCO CENTRAL DE RESERVA DEL PERU - DEPARTAMENTO DE ENCAJE"
xlHoja1.Cells(3, 1) = "OBLIGACIONES EXONERADAS DE GUARDAR ENCAJE CON INSTITUCIONES FINANCIERAS DEL PAIS"
xlHoja1.Cells(4, 1) = "REPORTE Nº 2"
xlHoja1.Cells(7, 1) = UCase("INSTITUCION : " & gsNomCmac)
xlHoja1.Cells(8, 1) = UCase("PERIODO: " & EmitePeriodo(pdFecIni, pdFecFin))




'xlHoja1.Cells(12, 1) = "CREDITOS, DEPOSITOS E INTERBANCARIOS RECIBIDOS DE INSTITUCIONES FINANCIERAS DEL PAIS"
Col = 1
If lsMoneda = "1" Then
    nPosicion = 16
    xlHoja1.Cells(9, 1) = "MONEDA         :  (EN NUEVOS SOLES)     "
Else
    xlHoja1.Cells(9, 1) = "MONEDA         :  (EN MONEDA EXTRANJERA)     "
    nPosicion = 16
End If
xlHoja1.Cells(nPosicion, Col) = "DIAS"
xlHoja1.Range(xlHoja1.Cells(11, Col), xlHoja1.Cells(11, Col)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range(xlHoja1.Cells(12, Col), xlHoja1.Cells(12, Col)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range(xlHoja1.Cells(12, Col), xlHoja1.Cells(12, Col)).Borders(xlEdgeBottom).LineStyle = xlContinuous
ColIni = 2
For i = 1 To UBound(lsAnexo2Col, 1)
    For j = 1 To UBound(lsAnexo2Col, 2)
        If lsAnexo2Col(i, j, 2) <> "" Then
            Col = Col + 1
            xlHoja1.Range(xlHoja1.Cells(12, Col), xlHoja1.Cells(12, Col)).Font.Size = 8
            xlHoja1.Range(xlHoja1.Cells(12, Col), xlHoja1.Cells(12, Col)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(12, Col), xlHoja1.Cells(12, Col)).Borders.LineStyle = xlContinuous
            
            'Original
                'xlHoja1.Cells(14, Col) = lsAnexo2Col(I, J, 2)
            'Fin Original
            
            'Cambiado
            If Left(lsAnexo2Col(i, j, 2), 10) = "NO SUBORDI" Then
                xlHoja1.Cells(12, Col) = "COFIDE"
            Else
                'If Col = 2 Then
                '    Col = 3
                'ElseIf Col = 3 Then
                '    Col = 2
                'End If
                xlHoja1.Cells(12, Col) = lsAnexo2Col(i, j, 2)
                'If Col = 3 Then
                '    Col = 2
                'ElseIf Col = 2 Then
                '    Col = 3
                'End If
            End If
            'Fin Cambio
            
            
        End If
    Next j
    xlHoja1.Cells(11, ColIni) = lsAnexo2Col(i, 0, 1)
    xlHoja1.Range(xlHoja1.Cells(11, ColIni), xlHoja1.Cells(11, Col)).Merge True
    xlHoja1.Range(xlHoja1.Cells(11, ColIni), xlHoja1.Cells(11, Col)).Borders.LineStyle = xlContinuous
    ColIni = Col + 1
Next i
'xlHoja1.Range(xlHoja1.Cells(12, 1), xlHoja1.Cells(12, Col)).Merge True
xlHoja1.Range(xlHoja1.Cells(11, 1), xlHoja1.Cells(19, Col)).Borders.LineStyle = xlContinuous
'xlHoja1.Range(xlHoja1.Cells(11, 1), xlHoja1.Cells(12, Col)).Borders.LineStyle = xlContinuous
End Sub

Private Sub CabeceraAnexo3(lsMoneda As String, pdFecIni As Date, pdFecFin As Date)
Dim i, j As Integer
Dim ColIni  As Long
xlAplicacion.Range("A1:R100").Font.Size = 9
xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.CenterHorizontally = True
xlHoja1.PageSetup.Zoom = 60

xlHoja1.Range("A1:A1").ColumnWidth = 5
xlAplicacion.Range("A14:R14").Font.Size = 6
xlAplicacion.Range("A10:R13").HorizontalAlignment = xlHAlignCenter

xlHoja1.Cells(1, 1) = "BANCO CENTRAL DE RESERVA DEL PERU-DEPARTAMENTO DE ENCAJE"
xlHoja1.Cells(3, 1) = "OBLIGACIONES EXONERADAS DE GUARDAR ENCAJE - EXTERIOR"
xlHoja1.Cells(4, 1) = lsCadenaMoneda

xlHoja1.Cells(7, 1) = UCase("INSTITUCION : " & gsNomCmac)
xlHoja1.Cells(8, 1) = UCase("Periodo: " & EmitePeriodo(pdFecIni, pdFecFin))

xlHoja1.Cells(12, 2) = "CREDITOS RECIBIDOS NO SUJETOS A ENCAJE"
Col = 1
'xlHoja1.Cells(13, Col) = "DIAS"
xlHoja1.Cells(21, Col) = "DIAS"
xlHoja1.Cells(18, Col) = "Cód. Operación"
xlHoja1.Cells(19, Col) = "Cód. Swift"
xlHoja1.Cells(20, Col) = "Destino Financiamiento"

xlHoja1.Cells(22, Col) = "Fecha Inicio"
xlHoja1.Cells(23, Col) = "Fecha Vcto."
xlHoja1.Cells(24, Col) = "PLAZO PROMEDIO"

'ALPA 20080804********************************************************
xlHoja1.Cells(13, 2) = "CREDITOS Y OTRAS OBLIG. DISTINTAS A DEPOSITOS"
xlHoja1.Range(xlHoja1.Cells(13, 2), xlHoja1.Cells(13, 10)).Merge True
xlHoja1.Range(xlHoja1.Cells(13, 2), xlHoja1.Cells(13, 10)).Borders.LineStyle = xlContinuous

xlHoja1.Cells(13, 11) = "DEPOSITOS"
xlHoja1.Range(xlHoja1.Cells(13, 11), xlHoja1.Cells(13, 13)).Merge True
xlHoja1.Range(xlHoja1.Cells(13, 11), xlHoja1.Cells(13, 13)).Borders.LineStyle = xlContinuous

xlHoja1.Cells(13, 14) = "PRESTAMOS SUBORDINADOS"
xlHoja1.Range(xlHoja1.Cells(12, 14), xlHoja1.Cells(13, 16)).Merge True
xlHoja1.Range(xlHoja1.Cells(12, 14), xlHoja1.Cells(13, 16)).Borders.LineStyle = xlContinuous

xlHoja1.Range(xlHoja1.Cells(18, 2), xlHoja1.Cells(18, 16)).Borders.LineStyle = xlContinuous
xlHoja1.Range(xlHoja1.Cells(12, 1), xlHoja1.Cells(21, 16)).Borders.LineStyle = xlContinuous

If lsMoneda = "1" Then
    xlHoja1.Cells(18, 2) = "100000"
    xlHoja1.Cells(18, 3) = "100000"
    xlHoja1.Cells(18, 4) = "101000"
    xlHoja1.Cells(18, 5) = "300100"
    xlHoja1.Cells(18, 6) = "300100"
    xlHoja1.Cells(18, 7) = "300110"
    xlHoja1.Cells(18, 8) = "100200"
    xlHoja1.Cells(18, 9) = "100200"
    xlHoja1.Cells(18, 10) = "100210"
    xlHoja1.Cells(18, 11) = "120000"
    xlHoja1.Cells(18, 12) = "120000"
    xlHoja1.Cells(18, 13) = "121000"
    xlHoja1.Cells(18, 14) = "500000"
    xlHoja1.Cells(18, 15) = "500000"
    xlHoja1.Cells(18, 16) = "501000"
    xlHoja1.Cells(19, 5) = "BOIDC0B1XXX"
    xlHoja1.Cells(20, 5) = "Capital de trabajo"

Else
    xlHoja1.Cells(18, 2) = "100000"
    xlHoja1.Cells(18, 3) = "100000"
    xlHoja1.Cells(18, 4) = "101000"
    xlHoja1.Cells(18, 5) = "400100"
    xlHoja1.Cells(18, 6) = "400100"
    xlHoja1.Cells(18, 7) = "400110"
    xlHoja1.Cells(18, 8) = "100200"
    xlHoja1.Cells(18, 9) = "100200"
    xlHoja1.Cells(18, 10) = "100210"
    xlHoja1.Cells(18, 11) = "120000"
    xlHoja1.Cells(18, 12) = "120000"
    xlHoja1.Cells(18, 13) = "121000"
    xlHoja1.Cells(18, 14) = "500000"
    xlHoja1.Cells(18, 15) = "500000"
    xlHoja1.Cells(18, 16) = "501000"
    xlHoja1.Cells(19, 5) = "ICROESM1XXX"
    xlHoja1.Cells(20, 5) = "Capital de trabajo"
End If
'********************************************1********************************
xlHoja1.Range(xlHoja1.Cells(13, Col), xlHoja1.Cells(13, Col)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range(xlHoja1.Cells(14, Col), xlHoja1.Cells(14, Col)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range(xlHoja1.Cells(14, Col), xlHoja1.Cells(14, Col)).Borders(xlEdgeBottom).LineStyle = xlContinuous
ColIni = 2
For i = 1 To UBound(lsAnexo2Col, 1)
    For j = 1 To UBound(lsAnexo2Col, 2)
        If lsAnexo2Col(i, j, 2) <> "" Then
            Col = Col + 1
            xlHoja1.Range(xlHoja1.Cells(14, Col), xlHoja1.Cells(17, Col)).Merge True
            xlHoja1.Range(xlHoja1.Cells(14, Col), xlHoja1.Cells(17, Col)).Font.Size = 8
            xlHoja1.Range(xlHoja1.Cells(14, Col), xlHoja1.Cells(17, Col)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(14, Col), xlHoja1.Cells(17, Col)).Borders.LineStyle = xlContinuous
            xlHoja1.Cells(14, Col) = lsAnexo2Col(i, j, 2)
        End If
    Next j
    'ALPA 20080805***************************************
    ''xlHoja1.Cells(13, ColIni) = lsAnexo2Col(I, 0, 1)
    ''xlHoja1.Range(xlHoja1.Cells(13, ColIni), xlHoja1.Cells(13, Col)).Merge True
    ''xlHoja1.Range(xlHoja1.Cells(13, ColIni), xlHoja1.Cells(13, Col)).Borders.LineStyle = xlContinuous
    ''ColIni = Col + 1
    '***************************************************
Next i
xlHoja1.Range(xlHoja1.Cells(12, 2), xlHoja1.Cells(12, 13)).Merge True
xlHoja1.Range(xlHoja1.Cells(12, 2), xlHoja1.Cells(12, 13)).Borders.LineStyle = xlContinuous

End Sub

Private Sub CabeceraAnexo4(pdFecIni As Date, pdFecFin As Date, Optional gsCodOpe As String = "761204")
Dim i, j As Integer
'xlAplicacion.Range("A1:R100").Font.FontStyle = "Arial"
xlAplicacion.Range("A1:R100").Font.Size = 12

xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.CenterVertically = True
xlHoja1.PageSetup.LeftMargin = xlAplicacion.CentimetersToPoints(1.6)
xlHoja1.PageSetup.RightMargin = xlAplicacion.CentimetersToPoints(1.2)
xlHoja1.PageSetup.Zoom = 60

xlHoja1.Range("A1:A1").ColumnWidth = 5
xlHoja1.Range("B1:Z1").ColumnWidth = 12

xlAplicacion.Range(xlHoja1.Cells(11, 1), xlHoja1.Cells(14, 21)).Font.Size = 6
xlAplicacion.Range(xlHoja1.Cells(10, 1), xlHoja1.Cells(14, 21)).HorizontalAlignment = xlHAlignCenter
xlAplicacion.Range(xlHoja1.Cells(10, 1), xlHoja1.Cells(14, 21)).VerticalAlignment = xlCenter

xlHoja1.Cells(1, 1) = "BANCO CENTRAL DE RESERVA DEL PERU-DEPARTAMENTO DE ENCAJE"

xlHoja1.Cells(3, 1) = "OBLIGACIONES EXONERADAS DE GUARDAR ENCAJE"
xlHoja1.Cells(4, 1) = UCase(lsCadenaMoneda)

xlHoja1.Cells(6, 1) = UCase("INSTITUCION : " & gsNomCmac)
xlHoja1.Cells(7, 1) = UCase("Periodo: " & EmitePeriodo(pdFecIni, pdFecFin))


'''xlHoja1.Cells(10, 2) = "OTRAS OBLIGACIONES EXONERADAS":
'''xlHoja1.Range(xlHoja1.Cells(10, 2), xlHoja1.Cells(11, 10)).Merge
'''ExcelCuadro xlHoja1, 2, 10, 10, 11


'''xlHoja1.Cells(10, 11) = "INFORMACION ADICIONAL":
'''xlHoja1.Range(xlHoja1.Cells(10, 11), xlHoja1.Cells(11, 15)).Merge


'''xlHoja1.Cells(12, 2) = "BONOS DE ":                     xlHoja1.Cells(13, 2) = "ARREND. FINANCIERO"

'''xlHoja1.Cells(12, 3) = "LETRAS":                        xlHoja1.Cells(13, 3) = "HIPOTECARIAS"
'''xlHoja1.Cells(12, 4) = "DEUDA SUBORDINADO":
''xlHoja1.Cells(14, 4) = "BONOS": xlHoja1.Cells(14, 5) = "BONOS":
'''xlHoja1.Range(xlHoja1.Cells(12, 4), xlHoja1.Cells(13, 5)).Merge
'''xlHoja1.Range(xlHoja1.Cells(12, 4), xlHoja1.Cells(13, 5)).Borders.LineStyle = xlContinuous
'xlHoja1.Range(xlHoja1.Cells(13, 4), xlHoja1.Cells(13, 5)).Merge True

'''xlHoja1.Cells(12, 6) = "OTROS":           xlHoja1.Cells(13, 6) = "BONOS"

'''xlHoja1.Cells(12, 7) = "CHEQUES DE GERENCIA":           xlHoja1.Cells(13, 7) = "A FAVOR DE ENTIDADES"
'''xlHoja1.Cells(12, 8) = "FONDO":                         xlHoja1.Cells(13, 8) = "MI VIVIENDA"
'ALPA***20080804******************
 Dim rs As New ADODB.Recordset
    Dim oRep As New DRepCtaColumna
    Set rs = oRep.CargaRepColumna(gsCodOpe, , , , , gbBitCentral)
    nCantE = rs.RecordCount
'*********************************

If Mid(gsOpeCod, 3, 1) = "1" Then
xlHoja1.Cells(18, 1) = "DIAS":
ExcelCuadro xlHoja1, 1, 12, 1, 18
ExcelCuadro xlHoja1, 2, 15, 21, 15
ExcelCuadro xlHoja1, 2, 12, 21, 13
ExcelCuadro xlHoja1, 2, 12, 21, 18
'''    'xlHoja1.Cells(12, 10) = "PROGRAMAS DE CREDITOS":
'''    'xlHoja1.Cells(13, 9) = "AGROBANCO": xlHoja1.Cells(13, 10) = "FONCODES"
'''    xlHoja1.Cells(12, 9) = "PROGRAMAS DE CREDITOS":
'''    xlHoja1.Cells(13, 9) = "FONCODES"
'''    xlHoja1.Range(xlHoja1.Cells(12, 9), xlHoja1.Cells(12, 10)).Merge True
'''    xlHoja1.Range(xlHoja1.Cells(13, 7), xlHoja1.Cells(13, 10)).Borders.LineStyle = xlContinuous
Do While Not rs.EOF
    Columna = Columna + 1
    If Trim(rs!cDescCol) = "Total" Then
    xlHoja1.Cells(14, rs!nNroCol + 1) = rs!cDescCol
    Else
    
    If rs!nNroCol < 17 Then
        xlHoja1.Cells(12, rs!nNroCol + 1) = rs!cDescCol
        xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol + 1) & "13").Select
        xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol + 1) & "13").MergeCells = True
        xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol + 1) & "13").EntireColumn.HorizontalAlignment = xlHAlignJustify
        xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol + 1) & "13").EntireColumn.VerticalAlignment = xlVAlignTop
    Else
        If rs!nNroCol = 17 Then
            xlHoja1.Cells(12, 18) = "OTROS"
            xlHoja1.Range("R12:U13").Select
            xlHoja1.Range("R12:U13").MergeCells = True
            xlHoja1.Range("R12:U13").EntireColumn.HorizontalAlignment = xlHAlignJustify
            xlHoja1.Range("R12:U13").EntireColumn.VerticalAlignment = xlVAlignTop
        End If
        xlHoja1.Cells(14, rs!nNroCol + 1) = rs!cDescCol
        xlHoja1.Range(devCelda(rs!nNroCol) & "13:" & devCelda(rs!nNroCol) & "13").Select
        xlHoja1.Range(devCelda(rs!nNroCol) & "13:" & devCelda(rs!nNroCol) & "13").MergeCells = True
        xlHoja1.Range(devCelda(rs!nNroCol) & "13:" & devCelda(rs!nNroCol) & "13").EntireColumn.HorizontalAlignment = xlHAlignJustify
        xlHoja1.Range(devCelda(rs!nNroCol) & "13:" & devCelda(rs!nNroCol) & "13").EntireColumn.VerticalAlignment = xlVAlignTop
    End If
    End If
    'xlHoja1.Range(devCelda(rs!nNroCol) & "(11) "B11:B11").EntireColumn.HorizontalAlignment = xlHAlignCenter
    
    rs.MoveNext
Loop
RSClose rs

Else
''''    xlHoja1.Cells(12, 9) = "PROGRAMAS DE CREDITOS":         xlHoja1.Cells(13, 9) = "PROGRAMA": xlHoja1.Cells(13, 10) = "PROGRAMA"
'''    xlHoja1.Range(xlHoja1.Cells(12, 9), xlHoja1.Cells(12, 10)).Merge True
'''    xlHoja1.Range(xlHoja1.Cells(12, 9), xlHoja1.Cells(12, 10)).Borders.LineStyle = xlContinuous
     xlHoja1.Cells(17, 1) = "DIAS":
    ExcelCuadro xlHoja1, 1, 12, 1, 17
    ExcelCuadro xlHoja1, 2, 12, 19, 17
    ExcelCuadro xlHoja1, 2, 13, 19, 16
    ExcelCuadro xlHoja1, 2, 14, 19, 14
    Do While Not rs.EOF
        Columna = Columna + 1
        If Trim(rs!cDescCol) = "Total" Or Trim(rs!cDescCol) = "" Then
            xlHoja1.Cells(13, rs!nNroCol + 1) = rs!cDescCol
        ElseIf rs!nNroCol = 10 Or rs!nNroCol = 13 Then
            xlHoja1.Cells(13, rs!nNroCol + 1) = rs!cDescCol
        ElseIf rs!nNroCol < 11 Then
            xlHoja1.Cells(12, rs!nNroCol + 1) = rs!cDescCol
            xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol + 2) & "12").Select
            xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol + 2) & "12").MergeCells = True
            xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol + 2) & "12").EntireColumn.HorizontalAlignment = xlHAlignJustify
            xlHoja1.Range(devCelda(rs!nNroCol) & "12:" & devCelda(rs!nNroCol + 2) & "12").EntireColumn.VerticalAlignment = xlVAlignTop
            Else
                If rs!nNroCol = 11 Then
                    xlHoja1.Cells(12, 11) = "DEUDA SUBORDINADA - OTROS 1"
                        xlHoja1.Range("k12:m12").Select
                        xlHoja1.Range("k12:m12").MergeCells = True
                        xlHoja1.Range("k12:m12").EntireColumn.HorizontalAlignment = xlHAlignJustify
                        xlHoja1.Range("k12:m12").EntireColumn.VerticalAlignment = xlVAlignTop
              
                        xlHoja1.Cells(12, 14) = "OTRAS 2"
                        xlHoja1.Range("N12:P12").Select
                        xlHoja1.Range("N12:P12").MergeCells = True
                        xlHoja1.Range("N12:P12").EntireColumn.HorizontalAlignment = xlHAlignJustify
                        xlHoja1.Range("N12:P12").EntireColumn.VerticalAlignment = xlVAlignTop
                        
                         xlHoja1.Cells(12, 17) = "TOTAL"
                        xlHoja1.Range("Q12:S12").Select
                        xlHoja1.Range("Q12:S12").MergeCells = True
                        xlHoja1.Range("Q12:S12").EntireColumn.HorizontalAlignment = xlHAlignJustify
                        xlHoja1.Range("Q12:S12").EntireColumn.VerticalAlignment = xlVAlignTop
                        xlHoja1.Cells(13, 19) = "TOTAL"
                    
''                End IfRR
''                    xlHoja1.Cells(14, rs!nNroCol + 1) = rs!cDescCol
''                    xlHoja1.Range(devCelda(rs!nNroCol) & "13:" & devCelda(rs!nNroCol) & "13").Select
''                    xlHoja1.Range(devCelda(rs!nNroCol) & "13:" & devCelda(rs!nNroCol) & "13").MergeCells = True
''                    xlHoja1.Range(devCelda(rs!nNroCol) & "13:" & devCelda(rs!nNroCol) & "13").EntireColumn.HorizontalAlignment = xlHAlignJustify
''                    xlHoja1.Range(devCelda(rs!nNroCol) & "13:" & devCelda(rs!nNroCol) & "13").EntireColumn.VerticalAlignment = xlVAlignTop
                End If
       End If
    rs.MoveNext
Loop
RSClose rs
End If
'ALPA 20080804*******************************
'''xlHoja1.Cells(12, 11) = "CHEQUES DE ":                  xlHoja1.Cells(13, 11) = "GERENCIA"
'''xlHoja1.Cells(12, 12) = "GIROS Y ":                     xlHoja1.Cells(13, 12) = "TRANSFERENCIAS POR PAGAR"
'''xlHoja1.Cells(12, 13) = "TRIBUTOS POR":                 xlHoja1.Cells(13, 13) = "PAGAR"
'''xlHoja1.Cells(12, 14) = "OPERACIONES":                  xlHoja1.Cells(13, 14) = "EN TRAMITE"
'''xlHoja1.Cells(12, 15) = "CUENTAS POR":                  xlHoja1.Cells(13, 15) = "PAGAR DIVERSAS":
'********************************************
ColumnaFinal = nCantE
End Sub

'Private Function CalculoEncajeDolares(pdFecIni As Date, pdFecFin As Date) As Currency
'Dim sql As String
'Dim rs As New ADODB.Recordset
'Dim lnPorcEncaje As Currency
'Dim lnCalculoEncajeDolares  As Currency
'
'gnEncajeBasico = MontoEncajeBase
'ReDim Encaje(0)
'For I = 1 To UBound(lsAnexo1, 2)
'    ReDim Preserve Encaje(I)
'    Encaje(I).lnTose = CCur(lsAnexo1(FilaEncaje, I))
'    Encaje(I).lnToseBase = gnEncajeBasico
'Next I
'
'For I = 1 To UBound(Encaje)
'    Encaje(I).lnExceso = Encaje(I).lnTose - Encaje(I).lnToseBase
'Next I
'
'lnPorcEncaje = Format(gnEncajeExig / gnTotalOblig, "#0.0000")
'For I = 1 To UBound(Encaje)
'    Encaje(I).lnEncBasico = IIf((Encaje(I).lnTose - Encaje(I).lnToseBase) > 0, Encaje(I).lnToseBase * lnPorcEncaje, Encaje(I).lnTose * lnPorcEncaje)
'Next I
'For I = 1 To UBound(Encaje)
'    Encaje(I).lnEncajeMarginal = IIf((Encaje(I).lnTose - Encaje(I).lnToseBase) > 0, Encaje(I).lnExceso * 0.2, 0)
'Next I
'
'For I = 1 To UBound(Encaje)
'    Encaje(I).lnTotalEncaje = Encaje(I).lnEncBasico + Encaje(I).lnEncajeMarginal
'Next I
'
'For I = 1 To UBound(Encaje)
'    Encaje(I).lnEncajeExig = Encaje(I).lnTotalEncaje
'Next I
'lnCalculoEncajeDolares = 0
'For I = 1 To UBound(Encaje)
'    lnCalculoEncajeDolares = lnCalculoEncajeDolares + Encaje(I).lnTotalEncaje
'Next I
'
'CalculoEncajeDolares = lnCalculoEncajeDolares
'End Function

Private Function EmitePeriodo(pdFecIni As Date, pdFecFin As Date) As String
Dim lsPeriodo As String
lsPeriodo = "#" & FillNum(Trim(Str(Month(pdFecIni))), 2, "0")
lsPeriodo = lsPeriodo & " (Del " & pdFecIni & " AL " & pdFecFin & " : Mes de " & Format(pdFecIni, "mmmm yyyy") & ")"
EmitePeriodo = lsPeriodo
End Function


Private Sub Form_Load()
nCantE = 0
End Sub
