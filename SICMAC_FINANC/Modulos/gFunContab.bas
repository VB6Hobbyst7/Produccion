Attribute VB_Name = "gFunContab"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A837285037A"
Option Base 0
Option Explicit

' Para declarar en MODULO */
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

Public Sub CargaArchivo(lsArchivo As String, lsRutaArchivo As String)
    Dim X As Long
    Dim Temp As String
    Temp = GetActiveWindow()
    X = ShellExecute(Temp, "open", lsArchivo, "", lsRutaArchivo, 1)
    If X <= 32 Then
        If X = 2 Then
            MsgBox "No se encuentra el Archivo adjunto, " & vbCr & " verifique el servidor de archivos", vbInformation, " Aviso "
        ElseIf X = 8 Then
            MsgBox "Memoria insuficiente ", vbInformation, " Aviso "
        Else
            MsgBox "No se pudo abrir el Archivo adjunto", vbInformation, " Aviso "
        End If
    End If
  
End Sub

'********************************
' Adiciona Hoja a LibroExcel
'********************************
Public Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional pbActivaHoja As Boolean = True)
Dim lbExisteHoja As Boolean
Dim lbBorrarRangos As Boolean
On Error Resume Next
lbExisteHoja = False
lbBorrarRangos = False
activaHoja:
For Each xlHoja1 In xlLibro.Worksheets
    If UCase(xlHoja1.Name) = UCase(psHojName) Then
        If Not pbActivaHoja Then
            SendKeys "{ENTER}"
            xlHoja1.Delete
        Else
            xlHoja1.Activate
            If lbBorrarRangos Then xlHoja1.Range("A1:BZ1").EntireColumn.Delete
            lbExisteHoja = True
        End If
       Exit For
    End If
Next
If Not lbExisteHoja Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = psHojName
    If Err Then
        Err.Clear
        pbActivaHoja = True
        lbBorrarRangos = True
        GoTo activaHoja
    End If
End If
End Sub

'***********************************************************
' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
'***********************************************************
Public Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
        
Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application

If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
   Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
  ExcelBegin = False
End Function
'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Public Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Public Function ExcelColumnaString(pnCol As Integer) As String
Dim sTexto As String
Dim nLetra As Integer
   If pnCol + 64 <= 90 Then
      sTexto = Chr(pnCol + 64)
   ElseIf pnCol + 64 <= 740 Then
      nLetra = Int((pnCol - 26) / 26) + IIf((pnCol - 26) Mod 26 = 0, 0, 1)
      sTexto = Chr(nLetra + 64) & Chr(((pnCol - 26) Mod (26 + IIf((pnCol - 26) Mod 26 = 0, 1, 0))) + IIf((pnCol - 26) Mod 26 = 0, nLetra, 1) + 63)
   End If
   ExcelColumnaString = sTexto
End Function

Public Sub ExcelCuadro(xlHoja1 As Excel.Worksheet, ByVal X1 As Currency, ByVal Y1 As Currency, ByVal X2 As Currency, ByVal Y2 As Currency, Optional lbLineasVert As Boolean = True, Optional lbLineasHoriz As Boolean = False)
xlHoja1.Range(xlHoja1.Cells(Y1, X1), xlHoja1.Cells(Y2, X2)).BorderAround xlContinuous, xlThin
If lbLineasVert Then
   If X2 <> X1 Then
     xlHoja1.Range(xlHoja1.Cells(Y1, X1), xlHoja1.Cells(Y2, X2)).Borders(xlInsideVertical).LineStyle = xlContinuous
   End If
End If
If lbLineasHoriz Then
    If Y1 <> Y2 Then
        xlHoja1.Range(xlHoja1.Cells(Y1, X1), xlHoja1.Cells(Y2, X2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End If
End If
End Sub

Public Function LeeConstanteSist(psConst As ConstSistemas) As String
Dim oFun As NConstSistemas
Set oFun = New NConstSistemas
LeeConstanteSist = oFun.LeeConstSistema(psConst)
Set oFun = Nothing
End Function
Public Function AsignaCtaObj(ByVal psCtaContCod As String, ByRef lsRaiz As String, ByRef lnTipoObj As TpoObjetos) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
'Dim lsRaiz As String
Dim oDescObj As ClassDescObjeto
Dim UP As UPersona
Dim lsFiltro As String
Dim oRHAreas As DActualizaDatosArea
Dim oCtaCont As DCtaCont
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo


Set oEfect = New Defectivo
Set oCtaIf = New NCajaCtaIF
Set oRHAreas = New DActualizaDatosArea
Set oDescObj = New ClassDescObjeto
Set oCtaCont = New DCtaCont
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
lnTipoObj = -1
Set rs1 = oCtaCont.CargaCtaObj(psCtaContCod, , True)
If Not rs1.EOF And Not rs1.BOF Then
    lsRaiz = ""
    lsFiltro = ""
    lnTipoObj = Val(rs1!cObjetoCod)
    Select Case Val(rs1!cObjetoCod)
        Case ObjCMACAgencias
            Set rs = oRHAreas.GetAgencias(rs1!cCtaObjFiltro)
        Case ObjCMACAgenciaArea
            lsRaiz = "Unidades Organizacionales"
            Set rs = oRHAreas.GetAgenciasAreas(rs1!cCtaObjFiltro)
        Case ObjCMACArea
            Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
        Case ObjEntidadesFinancieras
            lsRaiz = "Cuentas de Entidades Financieras"
            'Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, psCtaContCod)
            Set rs = oCtaIf.CargaCtasIF(Mid(psCtaContCod, 3, 1), rs1!cCtaObjFiltro)
        Case ObjDescomEfectivo
            lsRaiz = "Denominación"
            Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
        Case ObjPersona
            Set rs = Nothing
        Case Else
            Set rs = GetObjetos(Val(rs1!cObjetoCod))
    End Select
End If
rs1.Close
Set rs1 = Nothing
Set AsignaCtaObj = rs

Set oDescObj = Nothing
Set UP = Nothing
Set oCtaCont = Nothing
Set oCtaIf = Nothing
Set oEfect = Nothing
End Function

Public Sub ImprimeAsientoContable(ByVal psMovNro As String, Optional ByVal psDocVoucher As String = "", _
                                  Optional ByVal psDocTpo As String = "", Optional ByVal psDocumento As String = "", _
                                  Optional ByVal pbEfectivo As Boolean = False, _
                                  Optional ByVal pbIngreso As Boolean = False, _
                                  Optional ByVal psGlosa As String, Optional ByVal psPersCod As String, _
                                  Optional ByVal pnImporte As Currency, Optional ByVal pnTipoArendir As ArendirTipo, _
                                  Optional ByVal psNroRecViaticos As String = "", Optional pbHabEfectivo As Boolean = False, _
                                  Optional ByVal pnNumCopiasAsiento As Integer = 1, Optional psTitulo As String = "", Optional lsPie As String = "", Optional psNotaAbonoCargo As String = "", Optional psCabeAdicional As String = "", Optional psDocVentanilla As String = "", _
                                  Optional pbHistoCtaCont As Boolean = False, Optional nTasa As Double = 0, Optional nDias As Integer = 0, Optional ByVal bMultipleOpe As Boolean = False)
Dim oContImp As NContImprimir
Dim oNContFunc As NContFunciones
Dim oPlant As dPlantilla
Dim oNPlant As NPlantilla

Set oContImp = New NContImprimir
Set oNContFunc = New NContFunciones
Set oPlant = New dPlantilla
Set oNPlant = New NPlantilla

Dim lsAsiento  As String
Dim lsTitulo As String
Dim lsVEOPSave As String
Dim lsRecibo As String
Dim lsOPSave As String
Dim lsHab As String

Dim lsOtraFirma As String
Dim I As Integer
Dim lsCopias As String
Dim lsCartas As String
Dim oBarra   As New clsProgressBar

'ANDE 20170912 ERS029-2017
Dim nCantMov As Integer, nCantDocTpo As Integer, nCantDocumento As Integer, nCantNotaCargo As Integer, j As Integer
Dim aMovNro() As String, aDocTpo() As String, aDocumento() As String, aNotaCargo() As String

If bMultipleOpe = True Then
    aMovNro = Split(psMovNro, ",")
    nCantMov = UBound(aMovNro)
    
    If InStr(1, psDocTpo, ",") > 0 Then
        aDocTpo = Split(psDocTpo, ",")
        nCantDocTpo = UBound(aDocTpo)
    Else
        ReDim aDocTpo(nCantMov)
        For I = 0 To nCantMov
            ReDim Preserve aDocTpo(I)
            aDocTpo(I) = ""
        Next I
    End If
    
    If InStr(1, psDocumento, ",") > 0 Then
        aDocumento = Split(psDocumento, ",")
        nCantDocumento = UBound(aDocumento)
    Else
        ReDim aDocumento(nCantMov)
        For I = 0 To nCantMov
            ReDim Preserve aDocumento(I)
            aDocumento(I) = ""
        Next I
    End If
    If InStr(1, psNotaAbonoCargo, ",") Then
        aNotaCargo = Split(psNotaAbonoCargo, ",")
        nCantNotaCargo = UBound(aNotaCargo)
    Else
        ReDim aNotaCargo(nCantMov)
        For I = 0 To nCantMov
            ReDim Preserve aNotaCargo(I)
            aNotaCargo(I) = ""
        Next I
    End If
Else
    ReDim aMovNro(0)
    aMovNro(0) = psMovNro
    nCantMov = 0
    ReDim aDocTpo(0)
    aDocTpo(0) = psDocTpo
    nCantDocTpo = 0
    ReDim aDocumento(0)
    aDocumento(0) = psDocumento
    nCantDocumento = 0
    ReDim aNotaCargo(0)
    aNotaCargo(0) = psNotaAbonoCargo
    nCantNotaCargo = 0
End If

lsTitulo = psTitulo
If psDocVoucher <> "" Then
    lsTitulo = " COMPROBANTE DE EGRESO N° " & psDocVoucher
End If

If pbEfectivo Then
    If pnImporte <> 0 Then
        lsRecibo = oContImp.ImprimeReciboIngresoEgreso(aMovNro(0), gdFecSis, psGlosa, _
                                                       gsNomCmac, gsOpeCod, psPersCod, _
                                                       pnImporte, gnColPage, pnTipoArendir, psNroRecViaticos, pbIngreso) 'ande psMovNro
    End If
    If pbIngreso Then
        lsTitulo = "I N G R E S O   D E   E F E C T I V O"
    Else
        lsTitulo = "S A L I D A   D E   E F E C T I V O"
   End If
   lsPie = "39"
End If
If pbHabEfectivo Then
    lsTitulo = "H A B I L I T A C I O N   D E   E F E C T I V O "
    lsHab = oContImp.ImprimeDocSalidaEfectivo(gnColPage, gdFecSis, gsOpeCod, aMovNro(0), gsNomCmac) 'ande
    lsPie = "158"
    lsOtraFirma = "RESPONSABLE TRASLADO"
End If
If lsPie = "" Then
    lsPie = "19"
End If

If InStr(1, "401522,402522,401620,402620", gsOpeCod) = 0 Then
    'ANDE 20170912 ERS029-2017
    If bMultipleOpe = True Then
        If gsOpeCod = "401125" Or gsOpeCod = "402125" Or gsOpeCod = "401225" Or gsOpeCod = "402225" Then
           nCantMov = nCantMov - 1
        End If 'NAGL ERS 075-2017 20171130
        For I = 0 To nCantMov
            lsAsiento = lsAsiento & oContImp.ImprimeAsientoContable(aMovNro(I), gnLinPage, gnColPage, lsTitulo, psCabeAdicional, lsPie, lsOtraFirma, gsNomCmac, pbHistoCtaCont) 'JACA 20111227 para agregar pbHistoCtaCont 'ande psMovNro
            lsAsiento = lsAsiento & oImpresora.gPrnSaltoPagina
        Next I
    Else
        lsAsiento = oContImp.ImprimeAsientoContable(aMovNro(0), gnLinPage, gnColPage, lsTitulo, psCabeAdicional, lsPie, lsOtraFirma, gsNomCmac, pbHistoCtaCont) 'JACA 20111227 para agregar pbHistoCtaCont 'ande psMovNro
    End If
    'end ande
Else
    lsAsiento = oContImp.ImprimeAsientoContable(aMovNro(0), gnLinPage, gnColPage, lsTitulo, psCabeAdicional, lsPie, lsOtraFirma, gsNomCmac, pbHistoCtaCont, nTasa, nDias)   'JACA 20111227 para agregar pbHistoCtaCont 'ande
    'RIRO20140620 Se agregaron los parametros nTasa, nDias
End If

Dim oPrevio As clsPrevioFinan
Set oPrevio = New clsPrevioFinan

If psDocTpo <> "" Then
    For I = 0 To nCantDocTpo
        'If psDocTpo = TpoDocOrdenPago And pbIngreso = False Then
        If aDocTpo(I) = TpoDocOrdenPago And pbIngreso = False Then
            lsOPSave = oNPlant.GetPlantillaDoc(IDPlantillaOP)
            'lsOPSave = lsOPSave & psDocumento
            lsOPSave = lsOPSave & aDocumento(I)
            
            oPlant.GrabaPlantilla "OPBatch", "Ordenes de Pago para impresiones en Batch", lsOPSave
            
            lsVEOPSave = oNPlant.GetPlantillaDoc(IDPlantillaVOP)
            lsVEOPSave = lsVEOPSave & IIf(lsVEOPSave = "", "", oImpresora.gPrnSaltoPagina) & lsAsiento
            oPlant.GrabaPlantilla "OPVEBatch", "Voucher de egresos de Ordenes de Pago para impresiones en Batch", lsVEOPSave
            If MsgBox(" ¿ Desea Imprimir Orden de Pago ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
                lsOPSave = oNPlant.GetPlantillaDoc(IDPlantillaOP)
                If ImprimeOrdenPago(lsOPSave) Then
                '    oPrevio.Show lsOPSave, gsOpeDesc, False, gnLinPage / 4
                    lsVEOPSave = oNPlant.GetPlantillaDoc(IDPlantillaVOP)
                    'oPrevio.Show lsVEOPSave, gsOpeDesc, False, gnLinPage
                    oPlant.GrabaPlantilla IDPlantillaOP, "Ordenes de Pago para impresiones en Batch", ""
                    oPlant.GrabaPlantilla IDPlantillaVOP, "Voucher de egresos de Ordenes de Pago para impresiones en Batch", ""
                End If
           End If
        Else
            'If Val(psDocTpo) = TpoDocNotaAbono Or Val(psDocTpo) = TpoDocNotaCargo Or psNotaAbonoCargo <> "" Then
            If Val(aDocTpo(I)) = TpoDocNotaAbono Or Val(aDocTpo(I)) = TpoDocNotaCargo Or aNotaCargo(I) <> "" Then
                MsgBox "Se va a Imprimir Boleta de Nota de Abono/Cargo." & vbCrLf & "Por favor prepare su impresora con papel boleta para realizar la Impresion ", vbExclamation, "Aviso"
                Dim lbimp As Boolean
                'If psNotaAbonoCargo = "" Then
                If aNotaCargo(I) = "" Then
                    'psNotaAbonoCargo = psDocumento
                    aNotaCargo(I) = aDocumento(I)
                End If
                lbimp = True
                Do While lbimp
                    'oPrevio.ShowImpreSpool psNotaAbonoCargo, False, 22
                    oPrevio.ShowImpreSpool aNotaCargo(I), False, 22
                    If MsgBox("Desea Reimprimir boleta de Nota de Abono/Cargo??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                        lbimp = False
                    End If
                Loop
                Dim lsDocVentanilla As String
                If psDocVentanilla <> "" Then
                    lsDocVentanilla = oNPlant.GetPlantillaDoc("RecCajero") & psDocVentanilla
                    If MsgBox(" ¿ Desea Imprimir Comprobante(s) de Ventanilla ? ", vbQuestion + vbYesNo, "¡Confirmación") = vbYes Then
                        lbimp = True
                        Do While lbimp
                            oPrevio.ShowImpreSpool lsDocVentanilla, False, 22
                            If MsgBox("Desea Reimprimir Comprobante de Ventanilla??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                                lbimp = False
                            End If
                        Loop
                        oPlant.GrabaPlantilla "RecCajero", "Documento de Ventanilla, para impresiones en Batch", ""
                    Else
                        oPlant.GrabaPlantilla "RecCajero", "Documento de Ventanilla, para impresiones en Batch", lsDocVentanilla
                    End If
                End If
            End If
        End If
    Next I
End If

For j = 0 To nCantDocTpo
    'Select Case Val(psDocTpo)
    Select Case Val(aDocTpo(j))
        Case TpoDocCheque  '  TpoDocCheque
            'If psDocumento <> "" Then
            If aDocumento(j) <> "" Then
                'lsAsiento = psDocumento & oImpresora.gPrnSaltoPagina & lsAsiento
                lsAsiento = aDocumento(j) & oImpresora.gPrnSaltoPagina & lsAsiento
            End If
            
            For I = 1 To pnNumCopiasAsiento - 1
                lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
            Next
            'lsAsiento = psDocumento & IIf(psDocumento = "", "", oImpresora.gPrnSaltoPagina) + lsAsiento + lsCopias
            lsAsiento = aDocumento(j) & IIf(aDocumento(j) = "", "", oImpresora.gPrnSaltoPagina) + lsAsiento + lsCopias
        Case TpoDocCarta  ' TpoDocCarta
            'If psDocumento <> "" Then
            If aDocumento(j) <> "" Then
                frmCopiasImp.Show 1
                For I = 1 To frmCopiasImp.CopiasCartas
                    'lsCartas = lsCartas & IIf(lsCartas <> "", oImpresora.gPrnSaltoPagina, "") + psDocumento
                    lsCartas = lsCartas & IIf(lsCartas <> "", oImpresora.gPrnSaltoPagina, "") + aDocumento(j)
                Next I
                'lsCartas = psDocumento + lsCartas
                pnNumCopiasAsiento = frmCopiasImp.CopiasAsientos
            End If
            For I = 1 To pnNumCopiasAsiento - 1
                lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
            Next
            lsAsiento = lsAsiento & lsCopias
            lsAsiento = IIf(lsCartas = "", "", lsCartas & oImpresora.gPrnSaltoPagina) + lsAsiento
            Set frmCopiasImp = Nothing
        Case TpoDocOrdenPago, TpoDocNotaCargo, TpoDocNotaAbono         'TpoDocOrdenPago, TpoDocNotaCargo, TpoDocNotaAbono
            'If psDocumento <> "" Then
            If aDocumento(j) <> "" Then
                'lsAsiento = psDocumento & oImpresora.gPrnSaltoPagina & lsAsiento
                'lsAsiento = psDocumento & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsAsiento
                lsAsiento = aDocumento(j) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsAsiento
            End If
            For I = 1 To pnNumCopiasAsiento - 1
                lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
            Next
            lsAsiento = lsAsiento & lsCopias
        Case Else
            If pbHabEfectivo Then
                For I = 1 To pnNumCopiasAsiento - 1
                    lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
                Next
                lsAsiento = lsAsiento & lsCopias
                If lsHab <> "" Then
                    lsAsiento = lsAsiento & oImpresora.gPrnSaltoPagina & lsHab
                End If
            Else
                For I = 1 To pnNumCopiasAsiento - 1
                    lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
                Next
                lsAsiento = lsAsiento & lsCopias
            End If
            If lsRecibo <> "" Then
                lsAsiento = lsAsiento & oImpresora.gPrnSaltoPagina & lsRecibo
            End If
    End Select
Next j
If Len(Trim(lsAsiento)) > 50 Then
    oPrevio.Show lsAsiento, gsOpeDesc, False, gnLinPage, gImpresora
End If

Set oPrevio = Nothing
Set oContImp = Nothing
Set oNContFunc = Nothing
End Sub

Public Sub ImprimeAsientoNoContable(ByVal psMovNro As String, Optional ByVal psDocVoucher As String = "", _
                                  Optional ByVal psDocTpo As String = "", Optional ByVal psDocumento As String = "", _
                                  Optional ByVal pbEfectivo As Boolean = False, _
                                  Optional ByVal pbIngreso As Boolean = False, _
                                  Optional ByVal psGlosa As String, Optional ByVal psPersCod As String, _
                                  Optional ByVal pnImporte As Currency, Optional ByVal pnTipoArendir As ArendirTipo, _
                                  Optional ByVal psNroRecViaticos As String = "", Optional pbHabEfectivo As Boolean = False, _
                                  Optional ByVal pnNumCopiasAsiento As Integer = 1, Optional psTitulo As String = "", Optional lsPie As String = "", Optional psNotaAbonoCargo As String = "", Optional psCabeAdicional As String = "", Optional psDocVentanilla As String = "")
Dim oContImp As NContImprimir
Dim oNContFunc As NContFunciones
Dim oPlant As dPlantilla
Dim oNPlant As NPlantilla

Set oContImp = New NContImprimir
Set oNContFunc = New NContFunciones
Set oPlant = New dPlantilla
Set oNPlant = New NPlantilla

Dim lsAsiento  As String
Dim lsTitulo As String
Dim lsVEOPSave As String
Dim lsRecibo As String
Dim lsOPSave As String
Dim lsHab As String

Dim lsOtraFirma As String
Dim I As Integer
Dim lsCopias As String
Dim lsCartas As String
Dim oBarra   As New clsProgressBar

lsTitulo = psTitulo
If psDocVoucher <> "" Then
    lsTitulo = " COMPROBANTE DE EGRESO N° " & psDocVoucher
End If

If pbEfectivo Then
    If pnImporte <> 0 Then
        lsRecibo = oContImp.ImprimeReciboIngresoEgreso(psMovNro, gdFecSis, psGlosa, _
                                                       gsNomCmac, gsOpeCod, psPersCod, _
                                                       pnImporte, gnColPage, pnTipoArendir, psNroRecViaticos, pbIngreso)
    End If
    If pbIngreso Then
        lsTitulo = "I N G R E S O   D E   E F E C T I V O"
    Else
        lsTitulo = "S A L I D A   D E   E F E C T I V O"
   End If
   lsPie = "39"
End If
If pbHabEfectivo Then
    lsTitulo = "H A B I L I T A C I O N   D E   E F E C T I V O "
    lsHab = oContImp.ImprimeDocSalidaEfectivo(gnColPage, gdFecSis, gsOpeCod, psMovNro, gsNomCmac)
    lsPie = "158"
    lsOtraFirma = "RESPONSABLE TRASLADO"
End If
If lsPie = "" Then
    lsPie = "19"
End If

lsAsiento = oContImp.ImprimeAsientoNoContable(psMovNro, gnLinPage, gnColPage, lsTitulo, psCabeAdicional, lsPie, lsOtraFirma, gsNomCmac)
Dim oPrevio As clsPrevioFinan
Set oPrevio = New clsPrevioFinan
If psDocTpo <> "" Then
    If psDocTpo = TpoDocOrdenPago And pbIngreso = False Then
        lsOPSave = oNPlant.GetPlantillaDoc(IDPlantillaOP)
        lsOPSave = lsOPSave & psDocumento
        oPlant.GrabaPlantilla "OPBatch", "Ordenes de Pago para impresiones en Batch", lsOPSave
        
        lsVEOPSave = oNPlant.GetPlantillaDoc(IDPlantillaVOP)
        lsVEOPSave = lsVEOPSave & IIf(lsVEOPSave = "", "", oImpresora.gPrnSaltoPagina) & lsAsiento
        oPlant.GrabaPlantilla "OPVEBatch", "Voucher de egresos de Ordenes de Pago para impresiones en Batch", lsVEOPSave
        If MsgBox(" ¿ Desea Imprimir Orden de Pago ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            lsOPSave = oNPlant.GetPlantillaDoc(IDPlantillaOP)
            If ImprimeOrdenPago(lsOPSave) Then
            '    oPrevio.Show lsOPSave, gsOpeDesc, False, gnLinPage / 4
                lsVEOPSave = oNPlant.GetPlantillaDoc(IDPlantillaVOP)
                'oPrevio.Show lsVEOPSave, gsOpeDesc, False, gnLinPage
                oPlant.GrabaPlantilla IDPlantillaOP, "Ordenes de Pago para impresiones en Batch", ""
                oPlant.GrabaPlantilla IDPlantillaVOP, "Voucher de egresos de Ordenes de Pago para impresiones en Batch", ""
            End If
       End If
    Else
        If Val(psDocTpo) = TpoDocNotaAbono Or Val(psDocTpo) = TpoDocNotaCargo Or psNotaAbonoCargo <> "" Then
            MsgBox "Se va a Imprimir Boleta de Nota de Abono/Cargo." & vbCrLf & "Por favor prepare su impresora con papel boleta para realizar la Impresion ", vbExclamation, "Aviso"
            Dim lbimp As Boolean
            If psNotaAbonoCargo = "" Then
                psNotaAbonoCargo = psDocumento
            End If
            lbimp = True
            Do While lbimp
                oPrevio.ShowImpreSpool psNotaAbonoCargo, False, 22
                If MsgBox("Desea Reimprimir boleta de Nota de Abono/Cargo??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                    lbimp = False
                End If
            Loop
            Dim lsDocVentanilla As String
            If psDocVentanilla <> "" Then
                lsDocVentanilla = oNPlant.GetPlantillaDoc("RecCajero") & psDocVentanilla
                If MsgBox(" ¿ Desea Imprimir Comprobante(s) de Ventanilla ? ", vbQuestion + vbYesNo, "¡Confirmación") = vbYes Then
                    lbimp = True
                    Do While lbimp
                        oPrevio.ShowImpreSpool lsDocVentanilla, False, 22
                        If MsgBox("Desea Reimprimir Comprobante de Ventanilla??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                            lbimp = False
                        End If
                    Loop
                    oPlant.GrabaPlantilla "RecCajero", "Documento de Ventanilla, para impresiones en Batch", ""
                Else
                    oPlant.GrabaPlantilla "RecCajero", "Documento de Ventanilla, para impresiones en Batch", lsDocVentanilla
                End If
            End If
        End If
    End If
End If
Select Case Val(psDocTpo)
    Case TpoDocCheque  '  TpoDocCheque
        If psDocumento <> "" Then
'            lsAsiento = psDocumento & oImpresora.gPrnSaltoPagina & lsAsiento
            lsAsiento = oImpresora.gPrnSaltoPagina & lsAsiento
            
        End If
        
        For I = 1 To pnNumCopiasAsiento - 1
            lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
        Next
'        lsAsiento = psDocumento & IIf(psDocumento = "", "", oImpresora.gPrnSaltoPagina) + lsAsiento + lsCopias
         lsAsiento = psDocumento & lsAsiento + lsCopias
    
    Case TpoDocCarta  ' TpoDocCarta
        If psDocumento <> "" Then
            frmCopiasImp.Show 1
            For I = 1 To frmCopiasImp.CopiasCartas
                lsCartas = lsCartas & IIf(lsCartas <> "", oImpresora.gPrnSaltoPagina, "") + psDocumento
            Next I
            'lsCartas = psDocumento + lsCartas
            pnNumCopiasAsiento = frmCopiasImp.CopiasAsientos
        End If
        For I = 1 To pnNumCopiasAsiento - 1
            lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
        Next
        lsAsiento = lsAsiento & lsCopias
        lsAsiento = IIf(lsCartas = "", "", lsCartas & oImpresora.gPrnSaltoPagina) + lsAsiento
        Set frmCopiasImp = Nothing
    Case TpoDocOrdenPago, TpoDocNotaCargo, TpoDocNotaAbono         'TpoDocOrdenPago, TpoDocNotaCargo, TpoDocNotaAbono
        If psDocumento <> "" Then
            lsAsiento = psDocumento & oImpresora.gPrnSaltoPagina & lsAsiento
        End If
        For I = 1 To pnNumCopiasAsiento - 1
            lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
        Next
        lsAsiento = lsAsiento & lsCopias
    Case Else
        If pbHabEfectivo Then
            For I = 1 To pnNumCopiasAsiento - 1
                lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
            Next
            lsAsiento = lsAsiento & lsCopias
            If lsHab <> "" Then
                lsAsiento = lsAsiento & oImpresora.gPrnSaltoPagina & lsHab
            End If
        Else
            For I = 1 To pnNumCopiasAsiento - 1
                lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
            Next
            lsAsiento = lsAsiento & lsCopias
        End If
        If lsRecibo <> "" Then
            lsAsiento = lsAsiento & oImpresora.gPrnSaltoPagina & lsRecibo
        End If
End Select
If Len(Trim(lsAsiento)) > 50 Then
    oPrevio.Show lsAsiento, gsOpeDesc, False, gnLinPage, gImpresora
End If

Set oPrevio = Nothing
Set oContImp = Nothing
Set oNContFunc = Nothing
End Sub


Public Function ValidaConfiguracionRegional() As Boolean
Dim nMoneda As Currency
Dim nMonto As Double
Dim sNumero As String, sFecha As String
Dim nPosPunto As Integer, nPosComa As Integer

'Inicializamos las variables
ValidaConfiguracionRegional = True
nMoneda = 1234567
nMonto = 1234567
'Validamos Configuración de punto y Coma de Moneda
sNumero = Format$(nMoneda, "#,##0.00")
nPosPunto = InStr(1, sNumero, ".", vbTextCompare)
nPosComa = InStr(1, sNumero, ",", vbTextCompare)

If nPosPunto < nPosComa Then
    ValidaConfiguracionRegional = False
    Exit Function
End If
'Validamos la configuración del punto y coma de los números
sNumero = Format$(nMonto, "#,##0.00")
nPosPunto = InStr(1, sNumero, ".", vbTextCompare)
nPosComa = InStr(1, sNumero, ",", vbTextCompare)
If nPosPunto < nPosComa Then
    ValidaConfiguracionRegional = False
    Exit Function
End If
'Validamos la fecha y la configuración de la hora
If Date <> Format$(Date, "dd/MM/yyyy") Then 'Validar el formato de la fecha
    ValidaConfiguracionRegional = False
    Exit Function
End If

sFecha = Format$(Date & " " & Time, "dd/mm/yyyy hh:mm:ss AMPM")
If InStr(1, sFecha, "A.M.", vbTextCompare) > 0 Then
    ValidaConfiguracionRegional = False
    Exit Function
End If
If InStr(1, sFecha, "P.M.", vbTextCompare) > 0 Then
    ValidaConfiguracionRegional = False
    Exit Function
End If
sFecha = Trim(Date)
If Day(Date) <> CInt(Mid(sFecha, 1, 2)) Then
    ValidaConfiguracionRegional = False
    Exit Function
End If
If Month(Date) <> CInt(Mid(sFecha, 4, 2)) Then
    ValidaConfiguracionRegional = False
    Exit Function
End If
If Year(Date) <> CInt(Mid(sFecha, 7, 4)) Then
    ValidaConfiguracionRegional = False
    Exit Function
End If

End Function

Public Function fgActualizaUltVersionEXE(psAgenciaCod As String) As Boolean
Dim fs As Scripting.FileSystemObject
Dim fCurrent As Scripting.Folder
Dim fi As Scripting.File
Dim fd As Scripting.File

Dim lsRutaUltActualiz As String
Dim lsRutaSICMACT As String
Dim lsFecUltModifLOCAL As String
Dim lsFecUltModifORIGEN As String
Dim lsFlagActualizaEXE As String

On Error GoTo ERROR
    fgActualizaUltVersionEXE = False
    Dim oCons As NConstSistemas
    Set oCons = New NConstSistemas
    
    lsRutaUltActualiz = oCons.GetRutaAcceso(psAgenciaCod)
    lsRutaSICMACT = App.path & "\"
    lsFlagActualizaEXE = oCons.LeeConstSistema(49)
    
    If lsFlagActualizaEXE = "0" Then  ' No Actualiza Ejecutable
        Exit Function
    End If
    
    If Dir(lsRutaSICMACT & "*.*") = "" Then
        Exit Function
    End If
    If Dir(lsRutaUltActualiz & "*.*") = "" Then
        Exit Function
    End If
 
    Set fs = New Scripting.FileSystemObject
    Set fCurrent = fs.GetFolder(lsRutaUltActualiz)
    For Each fi In fCurrent.Files
          If Right(UCase(fi.Name), 3) = "EXE" Or Right(UCase(fi.Name), 3) = "INI" Or Right(UCase(fi.Name), 3) = "DLL" Then
             lsFecUltModifORIGEN = Format(fi.DateLastModified, "yyyy/mm/dd hh:mm:ss")
             If Dir(lsRutaSICMACT & fi.Name) <> "" Then
                Set fd = fs.GetFile(lsRutaSICMACT & fi.Name)
                lsFecUltModifLOCAL = Format(fd.DateLastModified, "yyyy/mm/dd hh:mm:ss")
                If lsFecUltModifLOCAL < lsFecUltModifORIGEN And lsFecUltModifORIGEN <> "" Then ' ACTUALIZA
                    fgActualizaUltVersionEXE = True
                End If
             Else
                fgActualizaUltVersionEXE = True
             End If
             If fgActualizaUltVersionEXE = True Then
                Exit For
             End If
          End If
    Next
    If fgActualizaUltVersionEXE = True Then
'        frmHerActualizaSicmact.IniciaVariables True
'        frmHerActualizaSicmact.Show 1
    End If
    Exit Function

ERROR:
    MsgBox "No se puede acceder a la ruta de origen, de la Ultima Actualizacion. - " & lsRutaUltActualiz, vbInformation, "Aviso"
    fgActualizaUltVersionEXE = False
End Function

Public Function GetComprobRetencion(psMovNro As String) As String
    Dim oContImp As NContImprimir
    Set oContImp = New NContImprimir
    Dim oMov As DMov
    Set oMov = New DMov

    GetComprobRetencion = oContImp.ImprimeComprobanteRetencion(oMov.GetnMovNro(psMovNro))
End Function



Public Sub ImprimeAsientoContableNew(ByVal psMovNro As String, Optional ByVal psDocVoucher As String = "", _
                                  Optional ByVal psDocTpo As String = "", Optional ByVal psDocumento As String = "", _
                                  Optional ByVal pbEfectivo As Boolean = False, _
                                  Optional ByVal pbIngreso As Boolean = False, _
                                  Optional ByVal psGlosa As String, Optional ByVal psPersCod As String, _
                                  Optional ByVal pnImporte As Currency, Optional ByVal pnTipoArendir As ArendirTipo, _
                                  Optional ByVal psNroRecViaticos As String = "", Optional pbHabEfectivo As Boolean = False, _
                                  Optional ByVal pnNumCopiasAsiento As Integer = 1, Optional psTitulo As String = "", Optional lsPie As String = "", Optional psNotaAbonoCargo As String = "", Optional psCabeAdicional As String = "", Optional psDocVentanilla As String = "", Optional psMoneda As Integer)
Dim oContImp As NContImprimir
Dim oNContFunc As NContFunciones
Dim oPlant As dPlantilla
Dim oNPlant As NPlantilla

Set oContImp = New NContImprimir
Set oNContFunc = New NContFunciones
Set oPlant = New dPlantilla
Set oNPlant = New NPlantilla

Dim lsAsiento  As String
Dim lsTitulo As String
Dim lsVEOPSave As String
Dim lsRecibo As String
Dim lsOPSave As String
Dim lsHab As String

Dim lsOtraFirma As String
Dim I As Integer
Dim lsCopias As String
Dim lsCartas As String
Dim oBarra   As New clsProgressBar

lsTitulo = psTitulo
If psDocVoucher <> "" Then
    lsTitulo = " COMPROBANTE DE EGRESO N° " & psDocVoucher
End If

If pbEfectivo Then
    If pnImporte <> 0 Then
        lsRecibo = oContImp.ImprimeReciboIngresoEgreso(psMovNro, gdFecSis, psGlosa, _
                                                       gsNomCmac, gsOpeCod, psPersCod, _
                                                       pnImporte, gnColPage, pnTipoArendir, psNroRecViaticos, pbIngreso)
    End If
    If pbIngreso Then
        lsTitulo = "I N G R E S O   D E   E F E C T I V O"
    Else
        lsTitulo = "S A L I D A   D E   E F E C T I V O"
   End If
   lsPie = "39"
End If
If pbHabEfectivo Then
    lsTitulo = "H A B I L I T A C I O N   D E   E F E C T I V O "
    lsHab = oContImp.ImprimeDocSalidaEfectivo(gnColPage, gdFecSis, gsOpeCod, psMovNro, gsNomCmac)
    lsPie = "158"
    lsOtraFirma = "RESPONSABLE TRASLADO"
End If
If lsPie = "" Then
    lsPie = "19"
End If

'lsAsiento = oContImp.ImprimeAsientoContable(psMovNro, gnLinPage, gnColPage, lsTitulo, psCabeAdicional, lsPie, lsOtraFirma, gsNomCmac)
lsAsiento = oContImp.ImprimeAsientoContableNew(psMovNro, gnLinPage, gnColPage, lsTitulo, psCabeAdicional, lsPie, lsOtraFirma, gsNomCmac, pnImporte, psMoneda)
Dim oPrevio As clsPrevioFinan
Set oPrevio = New clsPrevioFinan
If psDocTpo <> "" Then
    If psDocTpo = TpoDocOrdenPago And pbIngreso = False Then
        lsOPSave = oNPlant.GetPlantillaDoc(IDPlantillaOP)
        lsOPSave = lsOPSave & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoPagina & psDocumento
        oPlant.GrabaPlantilla "OPBatch", "Ordenes de Pago para impresiones en Batch", lsOPSave
        
        lsVEOPSave = oNPlant.GetPlantillaDoc(IDPlantillaVOP)
        lsVEOPSave = lsVEOPSave & IIf(lsVEOPSave = "", "", oImpresora.gPrnSaltoPagina) & lsAsiento
        oPlant.GrabaPlantilla "OPVEBatch", "Voucher de egresos de Ordenes de Pago para impresiones en Batch", lsVEOPSave
        If MsgBox(" ¿ Desea Imprimir Orden de Pago ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            lsOPSave = oNPlant.GetPlantillaDoc(IDPlantillaOP)
            If ImprimeOrdenPago(lsOPSave) Then
            '    oPrevio.Show lsOPSave, gsOpeDesc, False, gnLinPage / 4
                lsVEOPSave = oNPlant.GetPlantillaDoc(IDPlantillaVOP)
                'oPrevio.Show lsVEOPSave, gsOpeDesc, False, gnLinPage
                oPlant.GrabaPlantilla IDPlantillaOP, "Ordenes de Pago para impresiones en Batch", ""
                oPlant.GrabaPlantilla IDPlantillaVOP, "Voucher de egresos de Ordenes de Pago para impresiones en Batch", ""
            End If
       End If
    Else
        If Val(psDocTpo) = TpoDocNotaAbono Or Val(psDocTpo) = TpoDocNotaCargo Or psNotaAbonoCargo <> "" Then
            MsgBox "Se va a Imprimir Boleta de Nota de Abono/Cargo." & vbCrLf & "Por favor prepare su impresora con papel boleta para realizar la Impresion ", vbExclamation, "Aviso"
            Dim lbimp As Boolean
            If psNotaAbonoCargo = "" Then
                psNotaAbonoCargo = psDocumento
            End If
            lbimp = True
            Do While lbimp
                oPrevio.ShowImpreSpool psNotaAbonoCargo, False, 22
                If MsgBox("Desea Reimprimir boleta de Nota de Abono/Cargo??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                    lbimp = False
                End If
            Loop
            Dim lsDocVentanilla As String
            If psDocVentanilla <> "" Then
                lsDocVentanilla = oNPlant.GetPlantillaDoc("RecCajero") & psDocVentanilla
                If MsgBox(" ¿ Desea Imprimir Comprobante(s) de Ventanilla ? ", vbQuestion + vbYesNo, "¡Confirmación") = vbYes Then
                    lbimp = True
                    Do While lbimp
                        oPrevio.ShowImpreSpool lsDocVentanilla, False, 22
                        If MsgBox("Desea Reimprimir Comprobante de Ventanilla??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                            lbimp = False
                        End If
                    Loop
                    oPlant.GrabaPlantilla "RecCajero", "Documento de Ventanilla, para impresiones en Batch", ""
                Else
                    oPlant.GrabaPlantilla "RecCajero", "Documento de Ventanilla, para impresiones en Batch", lsDocVentanilla
                End If
            End If
        End If
    End If
End If
Select Case Val(psDocTpo)
    Case TpoDocCheque  '  TpoDocCheque
        If psDocumento <> "" Then
            lsAsiento = psDocumento & oImpresora.gPrnSaltoPagina & lsAsiento
        End If
        
        For I = 1 To pnNumCopiasAsiento - 1
            lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
        Next
        lsAsiento = psDocumento & IIf(psDocumento = "", "", oImpresora.gPrnSaltoPagina) + lsAsiento + lsCopias
    Case TpoDocCarta  ' TpoDocCarta
        If psDocumento <> "" Then
            frmCopiasImp.Show 1
            For I = 1 To frmCopiasImp.CopiasCartas
                lsCartas = lsCartas & IIf(lsCartas <> "", oImpresora.gPrnSaltoPagina, "") + psDocumento
            Next I
            'lsCartas = psDocumento + lsCartas
            pnNumCopiasAsiento = frmCopiasImp.CopiasAsientos
        End If
        For I = 1 To pnNumCopiasAsiento - 1
            lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
        Next
        lsAsiento = lsAsiento & lsCopias
        lsAsiento = IIf(lsCartas = "", "", lsCartas & oImpresora.gPrnSaltoPagina) + lsAsiento
        Set frmCopiasImp = Nothing
    Case TpoDocOrdenPago, TpoDocNotaCargo, TpoDocNotaAbono         'TpoDocOrdenPago, TpoDocNotaCargo, TpoDocNotaAbono
        If psDocumento <> "" Then
            'lsAsiento = psDocumento & oImpresora.gPrnSaltoPagina & lsAsiento
            lsAsiento = psDocumento & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsAsiento
        End If
        For I = 1 To pnNumCopiasAsiento - 1
            lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
        Next
        lsAsiento = lsCopias & lsAsiento & lsCopias
    Case Else
        If pbHabEfectivo Then
            For I = 1 To pnNumCopiasAsiento - 1
                lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
            Next
            lsAsiento = lsAsiento & lsCopias
            If lsHab <> "" Then
                lsAsiento = lsAsiento & oImpresora.gPrnSaltoPagina & lsHab
            End If
        Else
            For I = 1 To pnNumCopiasAsiento - 1
                lsCopias = lsCopias & oImpresora.gPrnSaltoPagina & lsAsiento
            Next
            lsAsiento = lsAsiento & lsCopias
        End If
        If lsRecibo <> "" Then
            lsAsiento = lsAsiento & oImpresora.gPrnSaltoPagina & lsRecibo
        End If
End Select
If Len(Trim(lsAsiento)) > 50 Then
    Dim MSWord As Word.Application
    Dim MSWordSource As Word.Application
    Set MSWord = New Word.Application
    Set MSWordSource = New Word.Application
    
    Dim RangeSource As Word.Range
    
    MSWordSource.Documents.Open FileName:=App.path & "\SPOOLER\Orden_Pago.doc"
    
    Set RangeSource = MSWordSource.ActiveDocument.Content
    'Lo carga en Memoria
    MSWordSource.ActiveDocument.Content.Copy
    'MSWordSource.ActiveDocument
    'Crea Nuevo Documento
        
    MSWord.Documents.Add
    
    MSWord.Application.Selection.TypeParagraph
    MSWord.Application.Selection.Paste
    MSWord.Application.Selection.InsertBreak
    
    'MSWordSource.ActiveDocument.Close
    Set MSWordSource = Nothing
        
    MSWord.Selection.SetRange start:=MSWord.Selection.start, End:=MSWord.ActiveDocument.Content.End
    MSWord.Selection.MoveEnd
 
    MSWord.ActiveDocument.Range.InsertBefore lsAsiento
    
    MSWord.ActiveDocument.Select
    MSWord.ActiveDocument.Range.Font.Name = "Courier New"
    MSWord.ActiveDocument.Range.Font.Size = 6
    MSWord.ActiveDocument.Range.Paragraphs.Space1
    MSWord.ActiveDocument.PageSetup.TopMargin = CentimetersToPoints(0.8)
    MSWord.ActiveDocument.PageSetup.LeftMargin = CentimetersToPoints(0.5)
    MSWord.ActiveDocument.PageSetup.RightMargin = CentimetersToPoints(1)

    MSWord.Selection.MoveDown Unit:=wdLine, Count:=99999
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    MSWord.Selection.TypeBackspace
    
    
    MSWord.ActiveDocument.SaveAs App.path & "\SPOOLER\Orden_Pago" & gsCodUser & Format(Now, "yyyymmsshhmmss") & ".doc"
    MSWord.Visible = True
    Set MSWord = Nothing
    Set MSWordSource = Nothing

    'oPrevio.Show lsAsiento, gsOpeDesc, False, gnLinPage, gImpresora
End If

Set oPrevio = Nothing
Set oContImp = Nothing
Set oNContFunc = Nothing
End Sub


Public Sub ImprimeAsientoContableUltimo(ByVal psMovNro As String, Optional ByVal psDocVoucher As String = "", _
                                  Optional ByVal psDocTpo As String = "", Optional ByVal psDocumento As String = "", _
                                  Optional ByVal pbEfectivo As Boolean = False, _
                                  Optional ByVal pbIngreso As Boolean = False, _
                                  Optional ByVal psGlosa As String, Optional ByVal psPersCod As String, _
                                  Optional ByVal pnImporte As Currency, Optional ByVal pnTipoArendir As ArendirTipo, _
                                  Optional ByVal psNroRecViaticos As String = "", Optional pbHabEfectivo As Boolean = False, _
                                  Optional ByVal pnNumCopiasAsiento As Integer = 1, Optional psTitulo As String = "", Optional lsPie As String = "", Optional psNotaAbonoCargo As String = "", Optional psCabeAdicional As String = "", Optional psDocVentanilla As String = "", Optional psMoneda As Integer, Optional pnDocProv As String = "")
Dim oContImp As NContImprimir
Dim oNContFunc As NContFunciones
Dim oPlant As dPlantilla
Dim oNPlant As NPlantilla

Set oContImp = New NContImprimir
Set oNContFunc = New NContFunciones
Set oPlant = New dPlantilla
Set oNPlant = New NPlantilla

Dim lsAsiento  As String
Dim lsTitulo As String
Dim lsVEOPSave As String
Dim lsRecibo As String
Dim lsOPSave As String
Dim lsHab As String

Dim lsOtraFirma As String
Dim I As Integer
Dim lsCopias As String
Dim lsCartas As String
Dim oBarra   As New clsProgressBar

lsTitulo = psTitulo
If psDocVoucher <> "" Then
    lsTitulo = " COMPROBANTE DE EGRESO N° " & psDocVoucher
End If

'lsAsiento = oContImp.ImprimeAsientoContableNew(psMovNro, gnLinPage, gnColPage, lsTitulo, psCabeAdicional, lsPie, lsOtraFirma, gsNomCmac, pnImporte, psMoneda)
Dim lsnMovNro As Long
Dim lsFecha As String
Dim lsGlosa As String
Dim lsDia As String
Dim lsMes As String
Dim lsAno As String
Dim lsNombreCorto As String * 35
Dim lsMontoLetras As String
Dim lsOpeCod As String
Dim lnCambioFijo As Currency
Dim lnCambioVenta As Currency
Dim lnCambioCompra As Currency
Dim lnCambioVentaSBS As Currency
Dim lnCambioCompraSBS As Currency
Dim sDocAbrev As String
Dim sDocNro As String
Dim sDocFecha As String
Dim sDocAbrevOP As String
Dim sDocNroOP As String
Dim sDocFechaOP As String
Dim sNomPers As String
Dim sAsiento As String
Dim sAsientoME As String
Dim sSql  As String
Dim sSQL1 As String
Dim lsMovNro As String
Dim oConect As DConecta
Dim rsDoc As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim oTipoCamb As nTipoCambio
Dim lsMonedaMN As String
Dim lsMonedaME As String
Dim lnMonto As String
Dim MontLetras As String
Dim oMov  As DMov
Dim nDia As String
Dim nMes As String
Dim nAno As String
Dim sCabeceraMN As String
Dim SAsiento1 As String
Dim SAsiento2 As String
Dim SAsiento3 As String
Dim SAsiento4 As String
Dim SAsiento5 As String
Dim SAsiento6 As String

Dim sCabeceraME As String
Dim SAsientoME1 As String
Dim SAsientoME2 As String
Dim SAsientoME3 As String
Dim SAsientoME4 As String
Dim SAsientoME5 As String
Dim SAsientoME6 As String

Dim sTipoCambio As String
Dim sCuentaBanco As String

Set oTipoCamb = New nTipoCambio
Set oMov = New DMov
Set rsDoc = New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set oConect = New DConecta

If oConect.AbreConexion = False Then Exit Sub
' ******************Recupero Datos del Mov*****************************************************
lsFecha = GetFechaMov(psMovNro, True)

'sSql = "SELECT * FROM Mov WHERE cMovNro = '" & psMovNro & "'"

'Set rs1 = oConect.CargaRecordSet(sSql)
Set rs1 = oMov.GetTodoMov(psMovNro)
If rs1.EOF Then
   Exit Sub
End If
lsnMovNro = rs1!nMovNro
lsGlosa = rs1!cMovDesc
lsMovNro = psMovNro
lsOpeCod = rs1!cOpeCod
sTipoCambio = ""

If Mid(lsOpeCod, 3, 1) = 2 Then
   sCuentaBanco = "200119573"
Else
   sCuentaBanco = "200016650"
End If

If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Or Left(lsOpeCod, 4) = Left(gOpeME, 4) Or rs1!cOpeCod = gContAjusteTipoCambio Then
  lnCambioFijo = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCFijoMes)
  lnCambioVenta = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCVenta)
  lnCambioCompra = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCCompra)
  lnCambioVentaSBS = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCPondVenta)
  lnCambioCompraSBS = oTipoCamb.EmiteTipoCambio(CDate(lsFecha), TCPonderado)
End If

'Tipo de Cambio
 sSql = "SELECT nMovTpoCambio, m.cOpeCod FROM MovTpoCambio MC JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE M.cMovNro = '" & lsMovNro & "'"
 Set rs = oConect.CargaRecordSet(sSql)
 
If Not rs.EOF Then
    If rs!cOpeCod = gContAjusteTipoCambio Then
    Else
       sTipoCambio = sTipoCambio & "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "      T.C.Mercado: " & Format(rs!nMovTpoCambio, "##,###,##0.000")
    End If
Else
    If Mid(lsOpeCod, 3, 1) = gMonedaExtranjera Then
        sTipoCambio = sTipoCambio & "  T.C.Fijo : " & Format(lnCambioFijo, "##,###,##0.000") & "     T.C.Compra : " & Format(lnCambioCompra, "##,###,##0.000") & "     T.C.Venta : " & Format(lnCambioVenta, "##,###,##0.000") & "     T.C.C SBS : " & Format(lnCambioCompraSBS, "##,###,##0.000") & "     T.C.V SBS : " & Format(lnCambioVentaSBS, "##,###,##0.000")
    End If
End If


' *********************  FIn Datos Mov **********************************************************

sSql = ""
' ******************** Recupero Datos Documentos *********************************************
sSql = "SELECT b.cDocAbrev, a.cDocNro, a.dDocFecha,b.nDocTpo FROM MovDoc a join mov m on m.nmovnro =a.nmovnro  ,Documento b " _
       & "WHERE m.cMovNro = '" & lsMovNro & "' and b.nDocTpo = a.nDocTpo "
       
        Set rsDoc = oConect.CargaRecordSet(sSql)
        If Not rsDoc.EOF Then
          'sDoc = ""
          Do While Not rsDoc.EOF
             '---  Sacar Numero de Orden de Pago  JEOM
             If rsDoc!nDocTpo = 48 Then
                sDocAbrevOP = rsDoc!cDocAbrev
                sDocNroOP = rsDoc!cDocNro
                sDocFechaOP = rsDoc!dDocFecha
            Else
            '------------------------------------------
                sDocAbrev = rsDoc!cDocAbrev
                sDocNro = rsDoc!cDocNro
                sDocFecha = rsDoc!dDocFecha
            End If
            rsDoc.MoveNext
          Loop
' *********************************** FIN DATOS DOCUMENTOS ********************************************
sSql = ""
' ***************** Recuperar Datos de Persona o Proveedor ****************************
Dim NroMov As Long
'Dim sSQL1 As String
'Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
 
 sSQL1 = "SELECT p.cPersNombre,p.cPersCod FROM MovGasto m inner join persona p on p.cPErscod =m.cPerscod WHERE m.nMovNro = '" & lsnMovNro & "'"
 Set rs1 = oConect.CargaRecordSet(sSQL1)
 If rs1.EOF Then
    Exit Sub
 End If
 sNomPers = rs1!cPersNombre
 lsNombreCorto = Mid(sNomPers, 1, 35)
 
' ***************** Fin de Datos Persona  o Proveedor ****************************
sSql = ""
Dim nTot As Currency
Dim nToH As Currency

Dim nTotME As Currency
Dim nTotHME As Currency


' **************** Monto para MN *************************************************
sSql = "SELECT mc.nMovItem, mc.cCtaContCod, ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),'') cCtaContDesc, mc.nMovImporte " _
       & "FROM   MovCta mc JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO WHERE  M.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
    Set rs = oConect.CargaRecordSet(sSql)
    If rs.EOF Then
       Exit Sub
    End If
    
    sCabeceraMN = ""
    SAsiento1 = ""
    SAsiento2 = ""
    SAsiento3 = ""
    SAsiento4 = ""
    SAsiento5 = ""
    SAsiento6 = ""
    
    
    sCabeceraMN = sCabeceraMN & " EN MONEDA NACIONAL "  '& oImpresora.gPrnSaltoLinea
    'sAsiento = sAsiento & Space(9) & "--------------------" & oImpresora.gPrnSaltoLinea
        
    Dim j As Integer
    j = 1
    
    Do While Not rs.EOF
       DoEvents
       Select Case j
        Case 1
            SAsiento1 = SAsiento1 & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " & Justifica(rs!cCtaContDesc, 71) & " " _
                  & Right(space(16) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, "##,###,##0.00"), ""), 16) & " " _
                  & Right(space(16) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, "##,###,##0.00"), ""), 16) '_
                  '& oImpresora.gPrnSaltoLinea
             If rs!nMovImporte > 0 Then
                nTot = nTot + Val(rs!nMovImporte)
             Else
                nToH = nToH + Val(rs!nMovImporte) * -1
             End If
         Case 2
            SAsiento2 = SAsiento2 & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " & Justifica(rs!cCtaContDesc, 71) & " " _
                 & Right(space(16) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, "##,###,##0.00"), ""), 16) & " " _
                 & Right(space(16) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, "##,###,##0.00"), ""), 16) '_
                 '& oImpresora.gPrnSaltoLinea
            If rs!nMovImporte > 0 Then
               nTot = nTot + Val(rs!nMovImporte)
            Else
               nToH = nToH + Val(rs!nMovImporte) * -1
            End If
         Case 3
            SAsiento3 = SAsiento3 & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " & Justifica(rs!cCtaContDesc, 71) & " " _
                 & Right(space(16) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, "##,###,##0.00"), ""), 16) & " " _
                 & Right(space(16) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, "##,###,##0.00"), ""), 16) '_
                 '& oImpresora.gPrnSaltoLinea
            If rs!nMovImporte > 0 Then
               nTot = nTot + Val(rs!nMovImporte)
            Else
               nToH = nToH + Val(rs!nMovImporte) * -1
            End If
        Case 4
            SAsiento4 = SAsiento4 & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " & Justifica(rs!cCtaContDesc, 71) & " " _
                 & Right(space(16) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, "##,###,##0.00"), ""), 16) & " " _
                 & Right(space(16) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, "##,###,##0.00"), ""), 16) '_
                 '& oImpresora.gPrnSaltoLinea
            If rs!nMovImporte > 0 Then
               nTot = nTot + Val(rs!nMovImporte)
            Else
               nToH = nToH + Val(rs!nMovImporte) * -1
            End If
        
        Case 5
            SAsiento5 = SAsiento5 & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " & Justifica(rs!cCtaContDesc, 71) & " " _
                 & Right(space(16) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, "##,###,##0.00"), ""), 16) & " " _
                 & Right(space(16) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, "##,###,##0.00"), ""), 16) '_
                 '& oImpresora.gPrnSaltoLinea
            If rs!nMovImporte > 0 Then
               nTot = nTot + Val(rs!nMovImporte)
            Else
               nToH = nToH + Val(rs!nMovImporte) * -1
            End If
            
        Case 6
            SAsiento6 = SAsiento6 & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " & Justifica(rs!cCtaContDesc, 71) & " " _
                 & Right(space(16) & IIf(rs!nMovImporte > 0, Format(rs!nMovImporte, "##,###,##0.00"), ""), 16) & " " _
                 & Right(space(16) & IIf(rs!nMovImporte < 0, Format(rs!nMovImporte * -1, "##,###,##0.00"), ""), 16) '_
                 '& oImpresora.gPrnSaltoLinea
            If rs!nMovImporte > 0 Then
               nTot = nTot + Val(rs!nMovImporte)
            Else
               nToH = nToH + Val(rs!nMovImporte) * -1
            End If
                     
       End Select
       j = j + 1
       rs.MoveNext
    Loop
'*******************************FIN Monto MN******************************************

'***************** Monto para ME *************************************************
sCabeceraME = ""
SAsientoME1 = ""
SAsientoME2 = ""
SAsientoME3 = ""
SAsientoME4 = ""
SAsientoME5 = ""
SAsientoME6 = ""

If Mid(gsOpeCod, 3, 1) = 2 Then
sSql = "SELECT mc.nMovItem, mc.cCtaContCod, LTRIM(RTRIM(ISNULL(dbo.GetCtaContDesc(mc.cCtaContCod,2,1),''))) cCtaContDesc, me.nmovMEimporte " _
         & "FROM   mov m join MovCta mc on mc.nmovnro=m.nmovnro JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
         & "WHERE  m.cMovNro = '" & lsMovNro & "' order by mc.nMovItem "
    Set rs = oConect.CargaRecordSet(sSql)
    'sTexto = ""
    'nTot = 0
    'nToH = 0
        
    sCabeceraME = sCabeceraME & " EN MONEDA EXTRANJERA " '& oImpresora.gPrnSaltoLinea
    'sCabeceraME = sCabeceraME & Space(9) & "----------------------" & oImpresora.gPrnSaltoLinea
    Dim K As Integer
    
    K = 1
    Do While Not rs.EOF
       
       Select Case K
       Case 1
            SAsientoME1 = SAsientoME1 & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " _
                   & Justifica(rs!cCtaContDesc, 71) & " " _
                   & IIf(rs!nMovMEImporte > 0, PrnVal(rs!nMovMEImporte, 16, 2), space(16)) & " " _
                   & IIf(rs!nMovMEImporte < 0, PrnVal(rs!nMovMEImporte * -1, 16, 2), space(16)) '_
                   '& oImpresora.gPrnSaltoLinea
            If rs!nMovMEImporte > 0 Then
               nTotME = nTotME + Val(rs!nMovMEImporte)
            Else
               nTotHME = nTotHME + Val(rs!nMovMEImporte) * -1
            End If
            
       Case 2
            SAsientoME2 = SAsientoME2 & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " _
              & Justifica(rs!cCtaContDesc, 71) & " " _
              & IIf(rs!nMovMEImporte > 0, PrnVal(rs!nMovMEImporte, 16, 2), space(16)) & " " _
              & IIf(rs!nMovMEImporte < 0, PrnVal(rs!nMovMEImporte * -1, 16, 2), space(16)) '_
              '& oImpresora.gPrnSaltoLinea
            If rs!nMovMEImporte > 0 Then
               nTotME = nTotME + Val(rs!nMovMEImporte)
            Else
               nTotHME = nTotHME + Val(rs!nMovMEImporte) * -1
            End If
       
       Case 3
            SAsientoME3 = SAsientoME3 & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " _
              & Justifica(rs!cCtaContDesc, 71) & " " _
              & IIf(rs!nMovMEImporte > 0, PrnVal(rs!nMovMEImporte, 16, 2), space(16)) & " " _
              & IIf(rs!nMovMEImporte < 0, PrnVal(rs!nMovMEImporte * -1, 16, 2), space(16)) '_
              '& oImpresora.gPrnSaltoLinea
            If rs!nMovMEImporte > 0 Then
               nTotME = nTotME + Val(rs!nMovMEImporte)
            Else
               nTotHME = nTotHME + Val(rs!nMovMEImporte) * -1
            End If
       
       Case 4
            SAsientoME4 = SAsientoME4 & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " _
              & Justifica(rs!cCtaContDesc, 71) & " " _
              & IIf(rs!nMovMEImporte > 0, PrnVal(rs!nMovMEImporte, 16, 2), space(16)) & " " _
              & IIf(rs!nMovMEImporte < 0, PrnVal(rs!nMovMEImporte * -1, 16, 2), space(16)) '_
              '& oImpresora.gPrnSaltoLinea
            If rs!nMovMEImporte > 0 Then
               nTotME = nTotME + Val(rs!nMovMEImporte)
            Else
               nTotHME = nTotHME + Val(rs!nMovMEImporte) * -1
            End If
            
        Case 5
            SAsientoME5 = SAsientoME5 & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " _
              & Justifica(rs!cCtaContDesc, 71) & " " _
              & IIf(rs!nMovMEImporte > 0, PrnVal(rs!nMovMEImporte, 16, 2), space(16)) & " " _
              & IIf(rs!nMovMEImporte < 0, PrnVal(rs!nMovMEImporte * -1, 16, 2), space(16)) '_
              '& oImpresora.gPrnSaltoLinea
            If rs!nMovMEImporte > 0 Then
               nTotME = nTotME + Val(rs!nMovMEImporte)
            Else
               nTotHME = nTotHME + Val(rs!nMovMEImporte) * -1
            End If
        Case 6
            SAsientoME6 = SAsientoME6 & rs!nMovItem & " " & Justifica(rs!cCtaContCod, 22) & " " _
              & Justifica(rs!cCtaContDesc, 71) & " " _
              & IIf(rs!nMovMEImporte > 0, PrnVal(rs!nMovMEImporte, 16, 2), space(16)) & " " _
              & IIf(rs!nMovMEImporte < 0, PrnVal(rs!nMovMEImporte * -1, 16, 2), space(16)) '_
              '& oImpresora.gPrnSaltoLinea
            If rs!nMovMEImporte > 0 Then
               nTotME = nTotME + Val(rs!nMovMEImporte)
            Else
               nTotHME = nTotHME + Val(rs!nMovMEImporte) * -1
            End If
            
            End Select
            K = K + 1
            rs.MoveNext
    Loop
End If
' **************** Fin Monto ME **************************************************

'******************Arma Cheque ***************************************************

lsDia = Mid(lsFecha, 1, 2)
lsMes = Mid(lsFecha, 4, 2)
lsAno = Right(lsFecha, 4)

MontLetras = ConvNumLet(pnImporte, True, True)
lnMonto = Format(pnImporte, "#,##0.00")

'***************** FIN Arma Cheque **********************************************




If psDocTpo <> "" Then
    If psDocTpo = TpoDocOrdenPago And pbIngreso = False Then
        lsOPSave = oNPlant.GetPlantillaDoc(IDPlantillaOP)
        lsOPSave = lsOPSave & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoPagina & psDocumento
        oPlant.GrabaPlantilla "OPBatch", "Ordenes de Pago para impresiones en Batch", lsOPSave
        
        lsVEOPSave = oNPlant.GetPlantillaDoc(IDPlantillaVOP)
        lsVEOPSave = lsVEOPSave & IIf(lsVEOPSave = "", "", oImpresora.gPrnSaltoPagina) & lsAsiento
        oPlant.GrabaPlantilla "OPVEBatch", "Voucher de egresos de Ordenes de Pago para impresiones en Batch", lsVEOPSave
    End If
End If

Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Dim R2 As ADODB.Recordset

Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
Set oDoc = oWord.Documents.Open(App.path & "\SPOOLER\Orden_Pago_Caja.doc")

With oWord.Selection.Find
     .Text = "lsFecha"
     .Replacement.Text = lsFecha
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With
With oWord.Selection.Find
     .Text = "dia"
     .Replacement.Text = lsDia
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With
With oWord.Selection.Find
     .Text = "mes"
     .Replacement.Text = lsMes
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
     .Text = "ano"
     .Replacement.Text = lsAno
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With

If Mid(lsOpeCod, 3, 1) = 1 Then
   '''lsMonedaMN = "S/." 'MARG ERS044-2016
   lsMonedaMN = gcPEN_SIMBOLO 'MARG ERS044-2016
   
Else
   lsMonedaME = "$."
End If
If Mid(lsOpeCod, 3, 1) = 1 Then
    With oWord.Selection.Find
         .Text = "Moneda"
         .Replacement.Text = lsMonedaMN
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
Else
    With oWord.Selection.Find
         .Text = "Moneda"
         .Replacement.Text = lsMonedaME
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
End If

With oWord.Selection.Find
     .Text = "Monto"
     .Replacement.Text = lnMonto
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
     .Text = "nombreCorto"
     .Replacement.Text = lsNombreCorto
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
     .Text = "sCuentaBanco"
     .Replacement.Text = sCuentaBanco
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With


With oWord.Selection.Find
     .Text = "nombre"
     .Replacement.Text = sNomPers
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
     .Text = "MontLetras"
     .Replacement.Text = MontLetras
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
     .Text = "Glosa"
     .Replacement.Text = lsGlosa
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
     .Text = "DocProv"
     .Replacement.Text = pnDocProv
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With


With oWord.Selection.Find
     .Text = "AbrevO"
     .Replacement.Text = sDocAbrevOP
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
     .Text = "OrdenP"
     .Replacement.Text = sDocNroOP
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
     .Text = "OrdenFecha"
     .Replacement.Text = sDocFechaOP
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
     .Text = "AbrevV"
     .Replacement.Text = sDocAbrev
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
     .Text = "Voucher"
     .Replacement.Text = sDocNro
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With

With oWord.Selection.Find
     .Text = "VouchFecha"
     .Replacement.Text = sDocFecha
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With
sTipoCambio = Trim(sTipoCambio)
With oWord.Selection.Find
     .Text = "sTipoCambio"
     .Replacement.Text = sTipoCambio
     .Forward = True
     .Wrap = wdFindContinue
     .Format = False
     .Execute Replace:=wdReplaceAll
End With

'-------- Asiento en Soles --------------

   With oWord.Selection.Find
         .Text = "sCabeceraMN"
         .Replacement.Text = sCabeceraMN
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
         .Text = "SAsiento1"
         .Replacement.Text = SAsiento1
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
         .Text = "SAsiento2"
         .Replacement.Text = SAsiento2
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    
        With oWord.Selection.Find
         .Text = "SAsiento3"
         .Replacement.Text = SAsiento3
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
         .Text = "SAsiento4"
         .Replacement.Text = SAsiento4
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
         .Text = "SAsiento5"
         .Replacement.Text = SAsiento5
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
         .Text = "SAsiento6"
         .Replacement.Text = SAsiento6
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    
' Asiento Dolares
    With oWord.Selection.Find
         .Text = "sCabeceraME"
         .Replacement.Text = sCabeceraME
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
         .Text = "SAsientoME1"
         .Replacement.Text = SAsientoME1
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
        With oWord.Selection.Find
         .Text = "SAsientoME2"
         .Replacement.Text = SAsientoME2
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
        With oWord.Selection.Find
         .Text = "SAsientoME3"
         .Replacement.Text = SAsientoME3
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
         .Text = "SAsientoME4"
         .Replacement.Text = SAsientoME4
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
         .Text = "SAsientoME5"
         .Replacement.Text = SAsientoME5
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
         .Text = "SAsientoME6"
         .Replacement.Text = SAsientoME6
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    
    
    With oWord.Selection.Find
         .Text = "Stotal1"
         .Replacement.Text = nTot
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
         .Text = "sTotal2"
         .Replacement.Text = nToH
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
   
   Dim MME1 As String
   Dim MME2 As String
   Dim sLinea As String
 
If nTotME = 0 Or nTotHME = 0 Then
   MME1 = ""
   MME2 = ""
   sLinea = ""
Else
   
    sLinea = String(138, "-")
    MME1 = Format(nTotME, "#,##0.00")
    MME2 = Format(nTotHME, "#,##0.00")
End If
    
    With oWord.Selection.Find
         .Text = "StotaMEl"
         .Replacement.Text = MME1
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
         .Text = "sTotalME2"
         .Replacement.Text = MME2
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
         .Text = "sLinea"
         .Replacement.Text = sLinea
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    

Set oContImp = Nothing
Set oNContFunc = Nothing
End If
End Sub

'*** PEAC 20100904
Public Function ValidaCtaAge(psCtaContCod As String) As Boolean

Dim oConec As DConecta
Set oConec = New DConecta

Dim oCtaCont As DCtaCont
Set oCtaCont = New DCtaCont
 
Dim rs1 As ADODB.Recordset
Dim lsSql As String
 
    lsSql = " exec stp_sel_CargaCtaContAge '" & psCtaContCod & "'"
    
    oConec.AbreConexion
    Set rs1 = oConec.CargaRecordSet(lsSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If rs1.EOF And rs1.BOF Then
       ValidaCtaAge = False
    Else
       ValidaCtaAge = True
    End If
    rs1.Close
   
End Function

'*** PEAC 20100908
Public Function Devuelvemes(PsMes As Integer) As String
    Dim lsNomMes As String
    Select Case PsMes
        Case 1
            lsNomMes = "Enero"
        Case 2
            lsNomMes = "Febrero"
        Case 3
            lsNomMes = "Marzo"
        Case 4
            lsNomMes = "Abril"
        Case 5
            lsNomMes = "Mayo"
        Case 6
            lsNomMes = "Junio"
        Case 7
            lsNomMes = "Julio"
        Case 8
            lsNomMes = "Agosto"
        Case 9
            lsNomMes = "Setiembre"
        Case 10
            lsNomMes = "Octubre"
        Case 11
            lsNomMes = "Noviembre"
        Case 12
            lsNomMes = "Diciembre"
    End Select
    Devuelvemes = lsNomMes
End Function

'*** PEAC 20100921
Sub Llenar_Combo_Agencia_con_Recordset(prs As ADODB.Recordset, pcboObjeto As ComboBox)
    pcboObjeto.Clear
    Do While Not prs.EOF
        pcboObjeto.AddItem Trim(prs!cConsDescripcion) & space(100) & Trim(prs!nConsValor)
        prs.MoveNext
    Loop
    prs.Close
End Sub

'*** PEAC 20100921
Sub Llenar_Combo_con_Recordset(prs As ADODB.Recordset, pcboObjeto As ComboBox)
    pcboObjeto.Clear
    Do While Not prs.EOF
        pcboObjeto.AddItem Trim(prs!cDescripcion) & space(100) & Trim(Str(prs!nConsValor))
        prs.MoveNext
    Loop
    prs.Close
End Sub

