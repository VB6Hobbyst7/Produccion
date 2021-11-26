VERSION 5.00
Begin VB.Form frmPresFinanExtInt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presupuesto de Financiamiento Externo e Interno a Largo Plazo (Desembolsos)"
   ClientHeight    =   720
   ClientLeft      =   960
   ClientTop       =   2235
   ClientWidth     =   2775
   Icon            =   "frmPresFinanExtInt.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmPresFinanExtInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type ltDesembolso
    lsCodEntidad As String
    lsNomEntidad As String
    lnPlazo      As Integer
    lsPlaza      As String
    Desem_T1_MN As Currency
    Desem_T1_ME As Currency
    Desem_T2_MN As Currency
    Desem_T2_ME As Currency
    Desem_T3_MN As Currency
    Desem_T3_ME As Currency
    Desem_T4_MN As Currency
    Desem_T4_ME As Currency
End Type
Dim ltDatos() As ltDesembolso

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lsArchivo As String
Dim lbExcel As Boolean
Dim nCont As Integer
Dim oCon As DConecta

Dim lnAnio    As Integer
Dim ldFechaAl As Date
Dim lnTCambio As Currency

Public Function ImprimeReporteFinanciamientoIE(pnAnio As Integer, pdFechaAl As Date, pnTipCambio As Currency)
lnAnio = pnAnio
ldFechaAl = pdFechaAl
lnTCambio = pnTipCambio

lbExcel = False
If FinanciamientoIE_GeneraDatos Then
    GeneraReporteExcel
End If
If lsArchivo <> "" Then
    '*******Carga el Archivo Excel a Objeto Ole ******
    CargaArchivo lsArchivo, App.path & "\SPOOLER\"
 End If
Exit Function
ErrorGenerarSD:
    MsgBox "Error N°[" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
    If lbExcel = True Then
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Function

Private Sub GeneraReporteExcel()
Dim fs As New Scripting.FileSystemObject
Dim lbExisteHoja As Boolean
Dim lnFila As Integer, lnCol As Integer
Dim I As Integer
Dim lsTotales As String
Dim Y1 As Currency, Y2 As Currency
Dim j As Integer, N As Integer
Dim lsDeudaVenPrinc As String
Dim lsDeudaVenInt As String
Dim lnFilaComun As Integer
Dim lnFilaInterno As Integer
Dim lnFilaExterno As Integer

Dim lsDatosComun() As String
Dim lsDatoInt() As String
Dim lsDatoExt() As String
Dim lsDirDatos() As String

    lsArchivo = "P_" & lnAnio & "_FinExtInt.XLS"
    ExcelBegin App.path & "\SPOOLER\" & lsArchivo, xlAplicacion, xlLibro, True
    lbExcel = True
    ExcelAddHoja "FinanEI", xlLibro, xlHoja1, True
    
    xlHoja1.PageSetup.Zoom = 50
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlAplicacion.Range("A1:Z1000").Font.Size = 8
    
    xlHoja1.Range("A1").ColumnWidth = 25
    xlHoja1.Range("B1").ColumnWidth = 14
    xlHoja1.Range("C1:Z1").ColumnWidth = 11
    lnFila = 1
    xlHoja1.Cells(lnFila, 1) = "EJECUCION INSTITUCIONAL DEL"
    xlHoja1.Cells(lnFila + 1, 1) = "PRESUPUESTO PARA EL"
    xlHoja1.Cells(lnFila + 2, 1) = "AÑO FISCAL " & lnAnio
    
    lnFila = lnFila + 3
    xlHoja1.Cells(lnFila, 2) = "EJECUCION DEL PRESUPUESTO DE FINANCIAMIENTO EXTERNO E INTERNO A LARGO Y CORTO PLAZO (DESEMBOLSOS)"
    xlHoja1.Cells(lnFila + 1, 5) = "AL " & IIf(Month(ldFechaAl) <= 3, "I", IIf(Month(ldFechaAl) <= 6, "II", IIf(Month(ldFechaAl) <= 9, "III", "IV"))) & " TRIMESTRE DEL AÑO " & lnAnio
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Font.Size = 12
    
    lnFila = lnFila + 3
    xlHoja1.Cells(lnFila, 1) = "RAZON SOCIAL : " & gsNomCmac:      xlHoja1.Cells(lnFila - 1, 17) = "ANEXO Nº E-1"
    xlAplicacion.Range(xlHoja1.Cells(1, 2), xlHoja1.Cells(lnFila, 18)).Font.Bold = True
    
    lnFila = lnFila + 1
    Y1 = lnFila
    xlHoja1.Cells(lnFila, 2) = "NOMBRE"
    xlHoja1.Cells(lnFila, 3) = "DISPOSITIVO"
    xlHoja1.Cells(lnFila, 5) = "MONTO PREVISTO AÑO " & (lnAnio - 1)
    xlHoja1.Cells(lnFila, 7) = "DESEMBOLSOS"
    xlHoja1.Cells(lnFila, 15) = "TOTAL"
    xlHoja1.Cells(lnFila, 17) = "AVANCE %"
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).Merge True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).Merge True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 14)).Merge True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 15), xlHoja1.Cells(lnFila, 16)).Merge True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 17), xlHoja1.Cells(lnFila, 18)).Merge True
    
    lnFila = lnFila + 1
    
    xlHoja1.Cells(lnFila, 1) = "ENTIDAD"
    xlHoja1.Cells(lnFila, 2) = "DEL"
    xlHoja1.Cells(lnFila, 3) = "LEGAL"
    xlHoja1.Cells(lnFila, 5) = "TOTAL (1)"
    xlHoja1.Cells(lnFila, 7) = "1er TRIMESTRE"
    xlHoja1.Cells(lnFila, 9) = "2do TRIMESTRE"
    xlHoja1.Cells(lnFila, 11) = "3er TRIMESTRE"
    xlHoja1.Cells(lnFila, 13) = "4to TRIMESTRE"
    xlHoja1.Cells(lnFila, 15) = "DESEMBOLSOS (%)"
    xlHoja1.Cells(lnFila, 17) = "(3) = (2)/(1)"
    
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).Merge True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).Merge True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 8)).Merge True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila, 10)).Merge True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 11), xlHoja1.Cells(lnFila, 12)).Merge True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 13), xlHoja1.Cells(lnFila, 14)).Merge True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 15), xlHoja1.Cells(lnFila, 16)).Merge True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 17), xlHoja1.Cells(lnFila, 18)).Merge True
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "ACREEDORA"
    xlHoja1.Cells(lnFila, 2) = "PROYECTO"
    xlHoja1.Cells(lnFila, 3) = "N°"
    xlHoja1.Cells(lnFila, 4) = "FECHA"
    xlHoja1.Cells(lnFila, 5) = gcME
    xlHoja1.Cells(lnFila, 6) = gcMN
    xlHoja1.Cells(lnFila, 7) = gcME
    xlHoja1.Cells(lnFila, 8) = gcMN
    xlHoja1.Cells(lnFila, 9) = gcME
    xlHoja1.Cells(lnFila, 10) = gcMN
    xlHoja1.Cells(lnFila, 11) = gcME
    xlHoja1.Cells(lnFila, 12) = gcMN
    xlHoja1.Cells(lnFila, 13) = gcME
    xlHoja1.Cells(lnFila, 14) = gcMN
    xlHoja1.Cells(lnFila, 15) = gcME
    xlHoja1.Cells(lnFila, 16) = gcMN
    xlHoja1.Cells(lnFila, 17) = gcME
    xlHoja1.Cells(lnFila, 18) = gcMN
    lnCol = 18
    xlAplicacion.Range(xlHoja1.Cells(Y1, 1), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
    xlAplicacion.Range(xlHoja1.Cells(Y1, 1), xlHoja1.Cells(lnFila, lnCol)).HorizontalAlignment = xlCenter
    xlAplicacion.Range(xlHoja1.Cells(Y1, 1), xlHoja1.Cells(lnFila, lnCol)).VerticalAlignment = xlCenter
    ExcelCuadro xlHoja1, 1, Y1, 18, CCur(lnFila)
    ExcelCuadro xlHoja1, 7, Y1 + 1, 14, Y1 + 1
    ExcelCuadro xlHoja1, 3, CCur(lnFila), 18, CCur(lnFila)

    Dim lsTotPlaza  As String
    Dim lnFilaPlazo As Integer
    Dim lsTotal     As String
    nCont = 1
    lnFilaPlazo = lnFila + 1
    lsTotPlaza = "+C" & (lnFila + 2)
    Llenavalores nCont, lnFila, "1", 0, "ENDEUDAMIENTO EXTERNO", "FINANCIAMIENTO A LARGO PLAZO"
    
    lnFila = lnFila + 1
    lsTotPlaza = "+C" & (lnFila + 1)
    xlHoja1.Range(xlHoja1.Cells(lnFilaPlazo, 3), xlHoja1.Cells(lnFilaPlazo, 3)).Formula = "=" & lsTotPlaza
    xlHoja1.Range(xlHoja1.Cells(lnFilaPlazo, 3), xlHoja1.Cells(lnFilaPlazo, 3)).AutoFill xlHoja1.Range("C" & lnFilaPlazo & ":R" & lnFilaPlazo), xlFillDefault
    Llenavalores nCont, lnFila, "0", 0, "ENDEUDAMIENTO INTERNO", ""
    lnFila = lnFila + 1
    ExcelCuadro xlHoja1, 1, CCur(lnFilaPlazo), 18, CCur(lnFila)
    lsTotal = "+C" & lnFilaPlazo
    
    lnFilaPlazo = lnFila + 1
    lsTotPlaza = "+B" & (lnFila + 2)
    Llenavalores nCont, lnFila, "1", 1, "ENDEUDAMIENTO EXTERNO", "FINANCIAMIENTO A CORTO PLAZO"
    lnFila = lnFila + 1
    lsTotPlaza = "+C" & (lnFila + 1)
    xlHoja1.Range(xlHoja1.Cells(lnFilaPlazo, 3), xlHoja1.Cells(lnFilaPlazo, 3)).Formula = "=" & lsTotPlaza
    xlHoja1.Range(xlHoja1.Cells(lnFilaPlazo, 3), xlHoja1.Cells(lnFilaPlazo, 3)).AutoFill xlHoja1.Range("C" & lnFilaPlazo & ":R" & lnFilaPlazo), xlFillDefault
    
    Llenavalores nCont, lnFila, "0", 1, "ENDEUDAMIENTO INTERNO", ""
    lnFila = lnFila + 1
    ExcelCuadro xlHoja1, 1, CCur(lnFilaPlazo), 18, CCur(lnFila)
    lsTotal = lsTotal & "+C" & lnFilaPlazo
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "TOTAL (A+B)"
    ExcelCuadro xlHoja1, 1, CCur(lnFila), 18, CCur(lnFila)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Formula = "=" & lsTotal
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).AutoFill xlHoja1.Range("C" & lnFila & ":R" & lnFila), xlFillDefault
    xlHoja1.Range(xlHoja1.Cells(11, 3), xlHoja1.Cells(lnFila, 18)).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 18)).Font.Bold = True
    
    ExcelEnd App.path & "\SPOOLER\" & lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
    
End Sub
    
Private Sub Llenavalores(ByRef nCont As Integer, ByRef lnFila As Integer, lsUltPlaza As String, lnUltPlazo As Integer, lsTextPlaza As String, lsTextPlazo As String)
Dim lnCol As Integer
Dim Y2    As Currency
Dim lnFilaPlaza As Integer
lnCol = 18
If lsTextPlazo <> "" Then
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = lsTextPlazo
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
    Y2 = lnFila
    ExcelCuadro xlHoja1, 2, Y2, CCur(lnCol), Y2
End If
If lsTextPlaza <> "" Then
    lnFila = lnFila + 1
    lnFilaPlaza = lnFila
    xlHoja1.Cells(lnFila, 1) = lsTextPlaza
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
End If
If UBound(ltDatos) > nCont Then
    Do While nCont <= UBound(ltDatos)
        If ltDatos(nCont).lsPlaza <> lsUltPlaza Or ltDatos(nCont).lnPlazo <> lnUltPlazo Then
            Exit Do
        End If
        lnFila = lnFila + 1
        xlHoja1.Cells(lnFila, 1) = ltDatos(nCont).lsNomEntidad
        xlHoja1.Cells(lnFila, 7) = ltDatos(nCont).Desem_T1_ME
        xlHoja1.Cells(lnFila, 8) = ltDatos(nCont).Desem_T1_MN
        xlHoja1.Cells(lnFila, 9) = ltDatos(nCont).Desem_T2_ME
        xlHoja1.Cells(lnFila, 10) = ltDatos(nCont).Desem_T2_MN
        xlHoja1.Cells(lnFila, 11) = ltDatos(nCont).Desem_T3_ME
        xlHoja1.Cells(lnFila, 12) = ltDatos(nCont).Desem_T3_MN
        xlHoja1.Cells(lnFila, 13) = ltDatos(nCont).Desem_T4_ME
        xlHoja1.Cells(lnFila, 14) = ltDatos(nCont).Desem_T4_MN
        xlHoja1.Range(xlHoja1.Cells(lnFila, 15), xlHoja1.Cells(lnFila, 15)).Formula = "=G" & lnFila & "+I" & lnFila & "+K" & lnFila & "+M" & lnFila
        xlHoja1.Range(xlHoja1.Cells(lnFila, 16), xlHoja1.Cells(lnFila, 16)).Formula = "=H" & lnFila & "+J" & lnFila & "+L" & lnFila & "+N" & lnFila
        nCont = nCont + 1
    Loop
    If lnFilaPlaza < lnFila Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaPlaza, 3), xlHoja1.Cells(lnFilaPlaza, 18)).Formula = "=SUM(C" & lnFilaPlaza + 1 & ":C" & lnFila & ")"
        xlHoja1.Range(xlHoja1.Cells(lnFilaPlaza, 3), xlHoja1.Cells(lnFilaPlaza, 3)).AutoFill xlHoja1.Range("C" & lnFilaPlaza & ":R" & lnFilaPlaza), xlFillDefault
    End If
End If
End Sub
   

Private Function FinanciamientoIE_GeneraDatos() As Boolean
Dim sSql    As String
Dim rs      As ADODB.Recordset
On Error GoTo ErrorGeneraDatos
Set oCon = New DConecta
oCon.AbreConexion

FinanciamientoIE_GeneraDatos = True
sSql = "SELECT ci.cIFTpo, ci.cPersCod, p.cPersNombre, Month(ci.dCtaIFAper) nMes, SubString(ci.cCtaIFCod,3,1) nMoneda, CASE WHEN nCtaIFPlazo * cia.nCtaIFCuotas > 365 THEN 0 ELSE 1 END nPlazo, cia.cPlaza, Sum(cia.nMontoPrestado) nPrestamo " _
     & "FROM Ctaif ci JOIN CtaifAdeudados cia ON cia.cPersCod = ci.cPersCod and " _
     & "     ci.cIFTpo = cia.cIFTpo And ci.cCtaIFCod = cia.cCtaIFCod " _
     & "     JOIN Persona p ON p.cPersCod = ci.cPersCod " _
     & "WHERE ci.cCtaIFEstado = '" & CGEstadoCtaIF.gEstadoCtaIFActiva & "' and Year(ci.dCtaIFAper) = " & lnAnio & " and ci.dCtaIFAper <= '" & Format(ldFechaAl, gsFormatoFecha) & "' " _
     & "GROUP BY ci.cIFTpo, ci.cPersCod, p.cPersNombre, month(ci.dCtaIFAper), SubString(ci.cCtaIFCod,3,1), CASE WHEN nCtaIFPlazo * cia.nCtaIFCuotas > 365 THEN 0 ELSE 1 END, cia.cPlaza " _
     & "ORDER BY nPlazo ASC, cia.cPlaza DESC, ci.cIFTpo, p.cPersNombre ASC "
Set rs = oCon.CargaRecordSet(sSql)
nCont = 0
Do While Not rs.EOF
    nCont = nCont + 1
    ReDim Preserve ltDatos(nCont)
    ltDatos(nCont).lsCodEntidad = rs!cPersCod
    ltDatos(nCont).lsNomEntidad = Trim(rs!cPersNombre)
    ltDatos(nCont).lnPlazo = rs!nPlazo
    ltDatos(nCont).lsPlaza = rs!cPlaza
    Do While rs!cPersCod = ltDatos(nCont).lsCodEntidad
        If rs!nMES <= 3 Then
            If rs!nMoneda = 1 Then
                ltDatos(nCont).Desem_T1_MN = rs!nPrestamo
            Else
                ltDatos(nCont).Desem_T1_ME = Round(rs!nPrestamo * lnTCambio, 2)
            End If
        ElseIf rs!nMES <= 6 Then
            If rs!nMoneda = 1 Then
                ltDatos(nCont).Desem_T2_MN = rs!nPrestamo
            Else
                ltDatos(nCont).Desem_T2_ME = Round(rs!nPrestamo * lnTCambio, 2)
            End If
        ElseIf rs!nMES <= 9 Then
            If rs!nMoneda = 1 Then
                ltDatos(nCont).Desem_T3_MN = rs!nPrestamo
            Else
                ltDatos(nCont).Desem_T3_ME = Round(rs!nPrestamo * lnTCambio, 2)
            End If
        Else
            If rs!nMoneda = 1 Then
                ltDatos(nCont).Desem_T4_MN = rs!nPrestamo
            Else
                ltDatos(nCont).Desem_T4_ME = Round(rs!nPrestamo * lnTCambio, 2)
            End If
        End If
        rs.MoveNext
        If rs.EOF Then
            Exit Do
        End If
    Loop
Loop
RSClose rs
oCon.CierraConexion
Set oCon = Nothing

Exit Function
ErrorGeneraDatos:
    FinanciamientoIE_GeneraDatos = False
    MsgBox "Error N° [" & Err.Number & "] " & Err.Description, vbInformation, "aviso"
End Function

