Attribute VB_Name = "CSimon"
Option Explicit

'Public gaObj() As String
'Variables de documentos
Public gcMovNro As String
Public gcGlosa As String
Public gcDocTpo As String
Public gcDocDesc As String
Public gcDocAbrev As String
Public gcMovEstado As String
Public gcEntiOrig As String
Public gcEntiDest As String
Public gcCtaEntiOrig As String
Public gcCtaEntiDest As String
Public gcPersona As String
Public gcObjetoCod As String
Public gcDirPers As String
Public gnSaldo As Currency
'Public gsOpeCod As String
'Public gsOpeDesc As String
'Public gsOpeDescHijo As String
'Public gsOpeDescPadre As String

Public Sub FlexBackColor(Flex As MSHFlexGrid, pnFil As Integer, pnColor As Double)
    Dim K     As Integer
    Dim lnCol As Integer
    Dim lnFil As Integer
    lnCol = Flex.Col
    lnFil = Flex.Row
    Flex.Row = pnFil
    For K = 1 To Flex.Cols - 1
       Flex.Col = K
       Flex.CellBackColor = pnColor
    Next
    Flex.Row = lnFil
    Flex.Col = lnCol
End Sub

Public Sub KeyUp_Flex(Flex As MSHFlexGrid, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyC And Shift = 2   '   Copiar  [Ctrl+C]
            Clipboard.Clear
            Clipboard.SetText Flex.Text
            KeyCode = 0
        Case vbKeyV And Shift = 2   '   Pegar  [Ctrl+V]
            Flex.Text = Clipboard.GetText
            KeyCode = 0
        Case vbKeyX And Shift = 2   '   Cortar  [Ctrl+X]
            Clipboard.Clear
            Clipboard.SetText Flex.Text
            Flex.Text = ""
            KeyCode = 0
        Case vbKeyDelete            '   Borrar [Delete]
            Flex.Text = ""
            KeyCode = 0
    End Select
End Sub

Public Sub Lin1(ByRef prtf As String, sTxt As String, Optional nLineas As Integer = 1)
    Dim N As Integer
    prtf = prtf & sTxt
    For N = 1 To nLineas
        prtf = prtf & oImpresora.gPrnSaltoLinea
    Next
End Sub

Public Function ImpreDetLog(sDetTxt As String, Optional sCant As String = "", Optional sMonto As String = "", Optional nLineas As Integer = 0) As String
    Dim sTexto As String
    
    sTexto = oImpresora.gPrnSaltoLinea & String(gnColPage + 1, "=") & oImpresora.gPrnSaltoLinea
    sTexto = sTexto & BON & CON & "Itm  Codigo               Descripcion" & Space(94 - Len(sCant) - Len(sMonto)) & IIf(Len(sCant) > 0, sCant, Space(8))
    sTexto = sTexto & "  " & IIf(Len(sMonto) > 0, sMonto, " ") & COFF & BOFF & oImpresora.gPrnSaltoLinea
    Lin1 sTexto, "", nLineas
    sTexto = sTexto & String(gnColPage + 1, "-") & oImpresora.gPrnSaltoLinea
    ImpreDetLog = sTexto & Replace(sDetTxt, "<<SALTO>>", sTexto) & _
                  String(gnColPage + 1, "=") & oImpresora.gPrnSaltoLinea
End Function
'641351

'Pie de Pagina
Public Function ImprimePiePagina(sPiePag) As String
   Dim sPie As String, sPiR As String
   Dim N  As Integer
   Dim nLenPie As Integer
   nLenPie = Len(sPiePag)
   ReDim aPiePag(nLenPie)
   nLenPie = (gnColPage + nLenPie) / nLenPie
   For N = 1 To Len(sPiePag)
      Select Case Mid(sPiePag, N, 1)
         Case 1:
            Lin1 sPiR, Centra("____________________", nLenPie), 0
            Lin1 sPie, Centra("     Vo Bo Caja      ", nLenPie), 0
         Case 2:
            Lin1 sPiR, Centra("___________________", nLenPie), 0
            Lin1 sPie, Centra("  Vo Bo Logistica  ", nLenPie), 0
         Case 3:
            Lin1 sPiR, Centra("____________________", nLenPie), 0
            Lin1 sPie, Centra("    Vo Bo Usuario   ", nLenPie), 0
         Case 4:
            Lin1 sPiR, Centra("___________________", nLenPie), 0
            Lin1 sPie, Centra("  Vo Bo Gerencia   ", nLenPie), 0
         Case 6:
            Lin1 sPiR, Centra("_____________________", nLenPie), 0
            Lin1 sPie, Centra("   Vo Bo Encargado   ", nLenPie), 0
         Case 7:
            Lin1 sPiR, Centra("_____________________", nLenPie), 0
            Lin1 sPie, Centra(" LE ________________ ", nLenPie), 0
         Case 8:
            Lin1 sPiR, Centra("____________________", nLenPie), 0
            Lin1 sPie, Centra("     PROVEEDOR      ", nLenPie), 0
         Case 9:
            Lin1 sPiR, Centra("____________________", nLenPie), 0
            Lin1 sPie, Centra(" Vo Bo Contabilidad ", nLenPie), 0
      End Select
   Next
   Lin1 ImprimePiePagina, "", 4
   Lin1 ImprimePiePagina, sPiR
   Lin1 ImprimePiePagina, sPie
   Lin1 ImprimePiePagina, "   "

End Function

'   ------------------------------------------------------------
'   Función     :   CabeRepo
'   Propósito   :   Define cabecera del reporte
'   Parámetro(s):   pnCarLin -> Caracteres x linea del reporte
'                   psSeccio -> Area responsable
'                   psTitRp1 -> Titulo 1 del reporte
'                   psTitRp2 -> Titulo 2 del reporte
'                   psMoneda -> Tipo de Moneda
'   Creado      :   02/07/1999  -   FAOS
'   Modificado  :   02/07/1999  -   FAOS
'   ------------------------------------------------------------
'   Formato     :   CabeRepo()
Public Function CabeRepoAnt(psCabe01 As String, psCabe02 As String, _
                         pnCarLin As Integer, psSeccio As String, _
                         psTitRp1 As String, psTitRp2 As String, _
                         psMoneda As String, psNumPag As String) As String

    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lsMoneda As String
    
    lsTitRp1 = "": lsTitRp2 = ""
    CabeRepoAnt = ""
    lsMoneda = ""
    lsMoneda = IIf(psMoneda = "", String(10, " "), " - " & psMoneda)
'    psCabe01 = ""
'    psCabe02 = ""
    
    '   Definición de Cabecera 1
 
    psCabe01 = FillText(UCase(Trim(gsNomAge)) & lsMoneda, 36, " ")
    psCabe01 = psCabe01 & Space((pnCarLin - 36) - (Len(psCabe01) - 2))
    psCabe01 = psCabe01 & "PAGINA: " & psNumPag
    psCabe01 = psCabe01 & Space(5) & "FECHA: " & Format(gdFecSis, gsFormatoFechaView)
    '   Definición de Cabecera 2
    psCabe02 = FillText(psSeccio, 19, " ")
    psCabe02 = psCabe02 & Space((pnCarLin - 19) - (Len(psCabe02) - 2))
    psCabe02 = psCabe02 & "HORA :   " & Format(Now(), "hh:mm:ss")
    '   Definición del Titulo del Reporte
    lsTitRp1 = String(Int((pnCarLin - Len(psTitRp1)) / 2), " ") & BON & psTitRp1 & BOFF
    lsTitRp2 = String(Int((pnCarLin - Len(psTitRp2)) / 2), " ") & BON & psTitRp2 & BOFF
    
    CabeRepoAnt = CabeRepoAnt & psCabe01 & oImpresora.gPrnSaltoLinea
    CabeRepoAnt = CabeRepoAnt & psCabe02 & oImpresora.gPrnSaltoLinea
    CabeRepoAnt = CabeRepoAnt & lsTitRp1 & oImpresora.gPrnSaltoLinea
    CabeRepoAnt = CabeRepoAnt & lsTitRp2
        
End Function

Public Sub EliminaRow2(fg As MSHFlexGrid, nItem As Integer, Optional nPrimerRow As Integer = 1)
Dim nPos As Integer
nPos = nItem
If fg.Rows > nPrimerRow + 1 Then
   fg.RemoveItem nPos
Else
   For nPos = 0 To fg.Cols - 1
       fg.TextMatrix(nPrimerRow, nPos) = ""
   Next
End If
End Sub

