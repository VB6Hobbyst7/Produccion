Attribute VB_Name = "gImprimir"
Option Explicit
Public ArcSal As Integer
Public lnNumCopias As Integer
Public lbCancela As Boolean
'***************************************************
'* Inicia una impresi�n - Cabecera
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'Referencia : Global ArcSal As Integer
'***************************************************
Public Sub ImpreBegin(pbCondensado As Boolean, nLineas As Integer)
    ArcSal = FreeFile
    
    Open sLPT For Output As ArcSal
    Print #ArcSal, oImpresora.gPrnInicializa;            'Inicializa Impresora
    If pbCondensado Then
       Print #ArcSal, oImpresora.gPrnMargenIzq00; 'Tipo letra : 0,1,2 - Roman,SansS,Courier
       Print #ArcSal, oImpresora.gPrnTamLetra10CPI;            'Tama�o  : 80, 77, 103
       Print #ArcSal, oImpresora.gPrnCondensadaON;                        'Retorna al tipo de letra normal
    Else
       Print #ArcSal, oImpresora.gPrnCondensadaOFF;
    End If
    Print #ArcSal, oImpresora.gPrnEspaLineaN;            'Espaciamiento lineas 1/6 pulg.1
    Print #ArcSal, oImpresora.gPrnEspaLineaValor & Chr$(nLineas);  '   Chr$(nLineas); 'Longitud de p�gina a 66 l�neas
    If Not pbCondensado Then
       Print #ArcSal, oImpresora.gPrnTpoLetraCurier;  'Tipo letra : 0,1,2 - Roman,SansS,Courier
       Print #ArcSal, oImpresora.gPrnTamLetra10CPI;            'Tama�o  : 80, 77, 103
    End If
    Print #ArcSal, oImpresora.gPrnTpoLetraRoman1P;   'Draf : 1 pasada
   
End Sub
'***************************************************
'* Termina un impresi�n - Cola
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'Referencia : Global ArcSal As Integer
'***************************************************
Public Sub ImpreEnd()
    Print #ArcSal, oImpresora.gPrnSaltoPagina;   'Nueva p�gina
    Print #ArcSal, oImpresora.gPrnCondensadaOFF;    'Retorna al tipo de letra normal
    Close ArcSal
End Sub
'***************************************************
'* Genera nueva p�gina
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'Referencia : Global ArcSal As Integer
'***************************************************
Public Sub ImpreNewPage()
    Print #ArcSal, oImpresora.gPrnSaltoPagina;   'Nueva p�gina
End Sub
'Prepara una cadena especial (cadena con caracteres con tilde y/o otros)
' para que se imprima en el modo FREEFILE.
Public Function ImpreCarEsp(ByVal vCadena As String) As String
    vCadena = Replace(vCadena, "�", Chr(160), , , vbTextCompare)
    vCadena = Replace(vCadena, "�", Chr(130), , , vbTextCompare)
    vCadena = Replace(vCadena, "�", Chr(161), , , vbTextCompare)
    vCadena = Replace(vCadena, "�", Chr(162), , , vbTextCompare)
    vCadena = Replace(vCadena, "�", Chr(163), , , vbTextCompare)
    vCadena = Replace(vCadena, "�", Chr(164), , , vbTextCompare)
    vCadena = Replace(vCadena, "�", Chr(165), , , vbTextCompare)
    vCadena = Replace(vCadena, "�", Chr(248), , , vbTextCompare)
    vCadena = Replace(vCadena, "�", Chr(179), , , vbTextCompare)
    ImpreCarEsp = vCadena
End Function
'Verifica la corrceta habilitaci�n de la impresora
Public Function ImpreSensa() As Boolean
Dim lbArchAbierto As Boolean
On Error GoTo ControlError
    ArcSal = FreeFile
    lbArchAbierto = True
    Open sLPT For Output As ArcSal
    Print #ArcSal, oImpresora.gPrnInicializa;            'Inicializa Impresora
    Close ArcSal
    lbArchAbierto = False
    ImpreSensa = True
    Exit Function
ControlError:   ' Rutina de control de errores.
    If lbArchAbierto Then
        Close ArcSal
    End If
    MsgBox "Impresora no Encontrada � Inactiva" & vbCr & "Verifique que la Conexi�n sea Correcta", vbExclamation, "Aviso de Precauci�n"
    ImpreSensa = False
End Function

Public Function ImpreMovNro(sMov As String) As String
ImpreMovNro = Mid(sMov, 1, 8) & "-" & Mid(sMov, 9, 6) & "-" & Right(sMov, 4)
End Function

Public Function PrnSet(Code As String, Optional nValor As Integer) As String
If nValor = 12 Or nValor = 10 Then
   nValor = nValor - 1
End If
Select Case Code
 Case "B+": PrnSet = oImpresora.gPrnBoldON 'Bold On
 Case "B-": PrnSet = oImpresora.gPrnBoldOFF 'Bold Off
 Case "U+": PrnSet = oImpresora.gPrnUnderLineONOFF   'Underline On
 Case "U-": PrnSet = oImpresora.gPrnUnderLineONOFF  'Underline Off
 Case "I+": PrnSet = oImpresora.gPrnItalicON 'Italic On
 Case "I-": PrnSet = oImpresora.gPrnItalicOFF 'Italic Off
 Case "W+": PrnSet = oImpresora.gPrnDblAnchoON 'Doble Ancho On
 Case "W-": PrnSet = oImpresora.gPrnDblAnchoOFF 'Doble Ancho Off
 Case "C+": PrnSet = oImpresora.gPrnCondensadaON 'Condensado On
 Case "C-": PrnSet = oImpresora.gPrnCondensadaOFF 'Condensado Off
 Case "Rm": PrnSet = oImpresora.gPrnTpoLetraRoman  'Roman
 Case "Ss": PrnSet = oImpresora.gPrnTpoLetraSansSerif  'Sans Serif
 Case "Co": PrnSet = oImpresora.gPrnTpoLetraCurier  'Courier
 Case "1.5": PrnSet = oImpresora.gPrnUnoMedioEspacio  ' 1 1/2 espacios
 Case "MI": PrnSet = oImpresora.gPrnMargenIzqCab & Chr$(nValor) 'Margen Izquierdo
 Case "MD": PrnSet = oImpresora.gPrnMargenDerCab & Chr$(nValor)   'Margen Derecho
 Case "10CPI": PrnSet = oImpresora.gPrnTamLetra12CPI
 Case "12CPI": PrnSet = oImpresora.gPrnTamLetra10CPI
 Case "15CPI": PrnSet = oImpresora.gPrnTamLetra15CPI
 Case "EspN": PrnSet = oImpresora.gPrnEspaLineaN     'Espaciado Normal 4.5/72
 Case "Esp":  PrnSet = oImpresora.gPrnEspaLineaValor & Chr$(nValor)  'Espaciado nValor/72 pulg.
End Select
End Function

Public Function ImpreFormat(ByVal pNumero As Variant, ByVal pLongitudEntera As Integer, _
        Optional ByVal pLongitudDecimal As Integer = 2, _
        Optional ByVal pMoneda As Boolean = False) As String
Dim vPosPto As Integer
Dim vParEnt As String
Dim vParDec As String
Dim vLonEnt As Integer
Dim vLonDec As Integer
Dim x As Integer

vParDec = ""
If IsNull(pNumero) Then
    If pLongitudDecimal > 0 Then vParDec = "." & String(pLongitudDecimal, "0")
    If pLongitudEntera <= 0 Then pLongitudEntera = 1
    ImpreFormat = String(pLongitudEntera - 1, " ") & "0" & vParDec
ElseIf VarType(pNumero) = 8 Then
    pNumero = Trim(pNumero)
    vLonEnt = Len(pNumero)
    If vLonEnt > pLongitudEntera Then
        pNumero = Left(pNumero, pLongitudEntera)
        vLonEnt = pLongitudEntera
    End If
    ImpreFormat = String(pLongitudDecimal, " ") & pNumero & String(pLongitudEntera - vLonEnt, " ")
Else
    vPosPto = InStr(Trim(CStr(pNumero)), ".")
    If vPosPto > 0 Then
        vParEnt = Trim(CStr(Left(pNumero, vPosPto - 1)))
        vParDec = Trim(CStr(Mid(pNumero, vPosPto + 1)))
        vLonEnt = Len(vParEnt)
        vLonDec = Len(vParDec)
    Else
        vParEnt = Trim(Str(pNumero))
        vParDec = ""
        vLonEnt = Len(vParEnt)
        vLonDec = 0
    End If
    If pMoneda And vLonEnt > 3 Then
        vParEnt = Format(vParEnt, "#,###,###")
        For x = 1 To Len(vParEnt)
            If Mid(vParEnt, x, 1) = "," Then pLongitudEntera = pLongitudEntera - 1
        Next x
    End If
    If vLonEnt > pLongitudEntera Then pLongitudEntera = vLonEnt + 1
    If vLonDec > pLongitudDecimal Then
        vLonDec = pLongitudDecimal
        vParDec = Left(vParDec, vLonDec)
    End If
    ImpreFormat = String(pLongitudEntera - vLonEnt, " ") & vParEnt
    If pLongitudDecimal > 0 Then
        ImpreFormat = ImpreFormat & "." & vParDec & String(pLongitudDecimal - vLonDec, "0")
    End If
End If

End Function
Public Function CabeRepo(psEmpresa As String, psAgencia As String, _
                         psSeccion As String, psMoneda As String, psFecha As String, _
                         psTitulo1 As String, psTitulo2 As String, _
                         psCabe1 As String, psCabe2 As String, _
                         pnNumPag As Integer, pnColRep As Integer) As String

   Dim lsImpre  As String
   Dim lsMoneda As String
   lsImpre = ""
   lsMoneda = IIf(psMoneda = "", String(10, " "), " - " & psMoneda)
   
   If pnNumPag > 0 Then
      Linea lsImpre, oImpresora.gPrnSaltoPagina
   End If
   pnNumPag = pnNumPag + 1
   
   Linea lsImpre, FillText(UCase(psEmpresa), pnColRep - 27, " ") & "Fecha : " & psFecha & " " & Format(Now(), "hh:mm:ss")
   Linea lsImpre, FillText(UCase(Trim(psAgencia)) & " - " & psSeccion & lsMoneda, pnColRep - 27, " ") + PrnSet("I+") + "Pagina: " & Format(pnNumPag, "000") + PrnSet("I-")
   Linea lsImpre, ""
   Linea lsImpre, PrnSet("B+") + Centra(psTitulo1, pnColRep) + PrnSet("B-")
   Linea lsImpre, PrnSet("B+") + Centra(psTitulo2, pnColRep) + PrnSet("B-"), 2
   If psCabe1 <> "" Then
      Linea lsImpre, Centra(psCabe1, pnColRep)
   End If
   If psCabe2 <> "" Then
      Linea lsImpre, Centra(psCabe2, pnColRep), 2
   End If
   CabeRepo = lsImpre
End Function

Public Function Linea(psVarImpre As String, psTexto As String, Optional pnLineas As Integer = 1, Optional ByRef pnLinCnt As Integer = 0) As String
Dim K As Integer
psVarImpre = psVarImpre & psTexto
For K = 1 To pnLineas
   psVarImpre = psVarImpre & oImpresora.gPrnSaltoLinea
   pnLinCnt = pnLinCnt + 1
Next
End Function

'   ------------------------------------------------------------
'   Funci�n     :   FillText
'   Prop�sito   :   Rellena a la derecha un campo texto con un
'                   caracter especificado
'   ------------------------------------------------------------
'   Formato: FillText("HOLA", 20, " ")
Public Function FillText(psCadena As String, pnLenTex As Integer, ChrFil As String) As String
    If pnLenTex > Len(Trim(psCadena)) Then
       FillText = Trim(psCadena) & String((pnLenTex - Len(Trim(psCadena))), ChrFil)
    End If
End Function

Public Function Encabezado(psCadena As String, pnItem As Long) As String
    Dim lsCadena As String
    Dim lsCampo As String
    Dim lnLonCampo As Long
    Dim lnTotalLinea As Long
    Dim lnPos As Long
    Dim lsResultado As String
    Dim I As Long
    Dim lsLineas As String
    
    lsResultado = ""
    lnTotalLinea = 0
        
    lsCadena = psCadena
    pnItem = pnItem + 3
    
    While lsCadena <> ""
        lnPos = InStr(1, lsCadena, ";", vbTextCompare)
        lsCampo = Left(lsCadena, lnPos - 1)
        lsCadena = Mid(lsCadena, lnPos + 1)
        lnPos = InStr(1, lsCadena, ";", vbTextCompare)
        
        lnLonCampo = CCur(Left(lsCadena, lnPos - 1))
        lsCadena = Mid(lsCadena, lnPos + 1)
        lnTotalLinea = lnTotalLinea + lnLonCampo
        lsResultado = lsResultado & Space(lnLonCampo - Len(lsCampo)) & lsCampo
    Wend
        
    lsResultado = lsResultado & oImpresora.gPrnSaltoLinea
    lsLineas = String(lnTotalLinea + 1, "-") & oImpresora.gPrnSaltoLinea
    lsResultado = lsLineas + lsResultado + lsLineas
    
    Encabezado = lsResultado
End Function

'***************************************************
'* Funcion que centra una cadena en una dimensi�n dada, si la cadena es
'* menor a esta dimensi�n la centra y rellena los espacios del caracter
'* espacio, caso contrario devuelve "No valido" y muestra un mensaje con el error
'***************************************************
'FECHA CREACION : 24/06/99  -   JPNZ
'MODIFICACION:
'***************************************************
Public Function CentrarCadena(psCadena As String, pnNroLineas As Long, Optional lnEspaciosIzq As Integer = 0, Optional lsCarImp As String = " ") As String
    Dim psNinf As Long
    Dim lnPosIni As Long
    
    psCadena = Trim(psCadena)
    If Len(psCadena) > pnNroLineas Then
        'psCadena = Left(psCadena, pnNroLineas)
        'MsgBox "EL valor de la Cadena enviada es mayor al espacio destinado", vbInformation, "Aviso"
        psCadena = Left(psCadena, pnNroLineas)
    End If
    'Else
    psNinf = Len(psCadena) / 2
    lnPosIni = Int(pnNroLineas / 2) - Int(Len(psCadena) / 2)
    
    'psCadena = String((pnNroLineas / 2) - psNinf, " ") & psCadena & String(pnNroLineas - Len(psCadena), " ")
    psCadena = String(lnEspaciosIzq, " ") & String(lnPosIni, lsCarImp) & psCadena & String(lnPosIni, lsCarImp)
    CentrarCadena = psCadena
   'End If
End Function
Public Function ImpreCabAsiento(ByVal pnColPage As Integer, ByVal pdFecha As Date, ByVal psEmpresaLogo As String, ByVal psOpeCod As String, ByVal psMovNro As String, Optional ByVal sTit As String, Optional ByVal lMoneda As Boolean = True, Optional lFecha As String = "", Optional lOpe As Boolean = True, Optional nLin As Integer = 4) As String
Dim nItem As Integer
Dim sTexto As String
Dim sImpre As String
Dim N As Integer
Dim BON As String, BOFF As String
Dim COFF As String
Dim CON As String
Dim lsArea As String
BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")


If lFecha = "" Then
   lFecha = pdFecha
End If
If Len(Trim(sTit)) = 0 Then
   sTit = "A S I E N T O   C O N T A B L E"
End If
' Definici�n de Cabecera 1
  sImpre = PrnSet("Co") & COFF
  For N = 1 To nLin
      sImpre = sImpre & oImpresora.gPrnSaltoLinea
  Next
  If psMovNro <> "" And lOpe Then
     sImpre = sImpre & CON & BON & ImpreFormat(psEmpresaLogo, 55) & COFF & " " & PrnSet("I+") & "Operaci�n: " & psOpeCod & " # " & psMovNro & PrnSet("I-") & BOFF & oImpresora.gPrnSaltoLinea

    Dim oGen As New DGeneral
    Dim rs   As ADODB.Recordset
    Set rs = oGen.GetDataUser(Right(psMovNro, 4))
    If Not rs.EOF And Not rs.BOF Then
        lsArea = Trim(rs!cDescAgActual) & " - " & Trim(rs!cAreaDescripcion)
    End If
    RSClose rs
    Set oGen = Nothing
    sImpre = sImpre & CON & lsArea & COFF & oImpresora.gPrnSaltoLinea
    nLin = nLin + 2
  Else
     sImpre = sImpre & BON & ImpreFormat(psEmpresaLogo, 57) & pdFecha & "-" & Format(Time, "hh:mm:ss") & BOFF & oImpresora.gPrnSaltoLinea
     nLin = nLin + 1
  End If
  sImpre = sImpre & String(Int((pnColPage - Len(sTit)) / 2), " ")
  sImpre = sImpre & BON & sTit & BOFF & oImpresora.gPrnSaltoLinea
  nLin = nLin + 1
  
  If lMoneda Then
     sTexto = "M O N E D A   " & IIf(Mid(psOpeCod, 3, 1) = gMonedaExtranjera, "E X T R A N J E R A", "N A C I O N A L")
     sImpre = sImpre & String(Int((pnColPage - Len(sTexto)) / 2), " ")
     sImpre = sImpre & BON & sTexto & BOFF & oImpresora.gPrnSaltoLinea
     nLin = nLin + 1
  End If
  ImpreCabAsiento = sImpre
End Function

Public Function ImpreGlosa(psGlosa As String, pnColPage As Integer, Optional psTitGlosa As String = "  GLOSA      : ", Optional pnCols As Integer = 0, Optional lbEnterFinal As Boolean = True, Optional ByRef nLin As Integer = 0) As String
Dim sImpre As String
Dim sTexto As String, N As Integer
Dim nLen As Integer
  nLen = Len(psTitGlosa)
  sTexto = JustificaTexto(psGlosa, IIf(pnCols = 0, pnColPage, pnCols) - nLen)
  sImpre = psTitGlosa
  N = 0
  Do While True
     N = InStr(sTexto, oImpresora.gPrnSaltoLinea)
     If N > 0 Then
        sImpre = sImpre & Mid(sTexto, 1, N - 1) & oImpresora.gPrnSaltoLinea & Space(nLen)
        sTexto = Mid(sTexto, N + 1, Len(sTexto))
        nLin = nLin + 1
     End If
     If N = 0 Then
        sImpre = sImpre & Justifica(sTexto, IIf(pnCols = 0, pnColPage, pnCols) - nLen) & IIf(lbEnterFinal, oImpresora.gPrnSaltoLinea, "")
        If lbEnterFinal Then
            nLin = nLin + 1
        End If
        Exit Do
     End If
  Loop
  ImpreGlosa = sImpre
End Function
Public Function Centra(psCad As String, Optional pnAncho As Integer = 80) As String
Dim N As Integer, m As Integer, I As Integer
N = Len(Trim(psCad))
m = (pnAncho - N) / 2
Centra = ""
If pnAncho < m + N Then
   pnAncho = m + N
End If
Centra = Space(m) & Trim(psCad) & Space(pnAncho - m - N)
End Function
Public Function JustificaTexto(sTemp As String, lnColPage As Integer, Optional lsEspIzq As Integer = 0) As String
Dim vTextFin As String
Dim Letra As String * 1, I As Integer, K As Integer, N As Integer
Dim nVeces As Long, m As Integer, Fin As Integer, Ini As Integer
Dim nAncho1 As Integer, nSpa As Integer
I = 0
K = 0
N = Len(sTemp)
nAncho1 = lnColPage
Do While I <= N
   K = K + 1
   I = I + 1
   If I > N Then
      Exit Do
   End If
   Letra = Mid(sTemp, I, 1)
   If Letra = Chr$(27) Then
      vTextFin = vTextFin & Letra & Mid(sTemp, I + 1, 1)
      I = I + 1
      K = K + 1
      nAncho1 = nAncho1 + 2
   Else
      If Asc(Letra) <> 13 And Asc(Letra) <> 10 Then
         If K > nAncho1 Then
            m = 0
            If Mid(sTemp, I, 1) = Chr(32) Then
               vTextFin = Trim(vTextFin)
            Else
               m = InStrRev(vTextFin, " ", , vbTextCompare)
               If m = 0 Then    'ANGC20210918
                    m = 1
                End If
               If InStr(Mid(vTextFin, m, Len(vTextFin)), Chr$(27)) Then
                  nAncho1 = nAncho1 - 2
               End If
               I = I - (nAncho1 + 1 - m)
               vTextFin = Mid(vTextFin, 1, m - 1)
            End If
            nSpa = nAncho1 - Len(Trim(vTextFin))
            vTextFin = Trim(vTextFin)
            If nSpa <> 0 Then
               Fin = 1
               nVeces = 0
               m = 1
               Do While m <= nSpa
                  Ini = InStr(Fin, vTextFin, " ", vbTextCompare)
                  If Ini = 0 Then
                     Fin = 1
                     nVeces = nVeces + 1
                     m = m + 1
                  Else
                      vTextFin = Mid(vTextFin, 1, Ini) & " " & RTrim(Mid(vTextFin, Ini + 1, nAncho1))
                      Fin = Ini + 2 + nVeces
                      m = m + 1
                  End If
               Loop
            End If
            vTextFin = vTextFin & oImpresora.gPrnSaltoLinea
            JustificaTexto = JustificaTexto & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin))
            nAncho1 = lnColPage
            vTextFin = ""
            Letra = ""
            K = 0
         Else
            vTextFin = vTextFin & Letra
         End If
      Else
        If I < N Then
            If Asc(Mid(sTemp, I + 1, 1)) = 13 Or Asc(Mid(sTemp, I + 1, 1)) = 10 Then
                I = I + 1
            End If
        End If
         JustificaTexto = JustificaTexto & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin)) & oImpresora.gPrnSaltoLinea
         nAncho1 = lnColPage
         vTextFin = ""
         Letra = ""
         K = 0
      End If
   End If
Loop
JustificaTexto = JustificaTexto & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin))
End Function
Public Function Justifica(sCad As String, nAncho As Integer)
Justifica = Mid(sCad & Space(nAncho), 1, nAncho)
End Function
Public Function ImprePiePag(pnColPage As Integer, Optional ByVal psPiePag As String = "9", Optional psOtraDesc As String = "") As String
'Pie de Pagina
Dim sPie As String, sPiR As String
Dim nLenPie As Integer
Dim N As Integer
Dim lsPiePag As String
       
nLenPie = Len(psPiePag)
ReDim aPiePag(nLenPie)
nLenPie = pnColPage / nLenPie
For N = 1 To Len(psPiePag)
   Select Case Mid(psPiePag, N, 1)
      Case 1:
         sPiR = sPiR + Centra("______________________", nLenPie)
         sPie = sPie + Centra("      Vo Bo Caja      ", nLenPie)
      Case 2:
         sPiR = sPiR + Centra("______________________", nLenPie)
         sPie = sPie + Centra("   Vo Bo Logistica    ", nLenPie)
      Case 3:
         sPiR = sPiR + Centra("______________________", nLenPie)
         sPie = sPie + Centra("    Vo Bo Usuario     ", nLenPie)
      Case 5:
         sPiR = sPiR + Centra("______________________", nLenPie)
         sPie = sPie + Centra("  Vo Bo Administrador ", nLenPie)
      Case 6:
         sPiR = sPiR + Centra("______________________", nLenPie)
         sPie = sPie + Centra("   Vo Bo Encargado   ", nLenPie)
      Case 7:
         sPiR = sPiR + Centra("______________________", nLenPie)
         sPie = sPie + Centra(" DNI _________________ ", nLenPie)
      Case 9:
         sPiR = sPiR + Centra("______________________", nLenPie)
         sPie = sPie + Centra("  Vo Bo Contabilidad ", nLenPie)
       Case 8
            If psOtraDesc <> "" Then
                sPiR = sPiR + Centra("______________________", nLenPie)
                sPie = sPie + Centra(psOtraDesc, nLenPie)
            End If
   End Select
Next
lsPiePag = lsPiePag + "" + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea '+ oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
lsPiePag = lsPiePag + sPiR + oImpresora.gPrnSaltoLinea
lsPiePag = lsPiePag + sPie + oImpresora.gPrnSaltoLinea
ImprePiePag = lsPiePag
End Function
Public Function Cabecera(sTit As String, P As Integer, psSimbolo As String, pnColPage As Integer, Optional sCabe As String = "", Optional sFecha As String = "", Optional psEmprLogo As String = " CMAC MAYNAS S.A.") As String
Dim BON As String
Dim BOFF As String
Dim CON As String
Dim COFF As String

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

Cabecera = ""
gsRUC = "20103845328"
If sFecha = "" Then
   sFecha = Date
End If
   If P > 0 Then Cabecera = oImpresora.gPrnSaltoPagina
   P = P + 1
   Cabecera = Cabecera + ImpreFormat(psEmprLogo, 90) & sFecha & "-" & Format(Time, "hh:mm:ss") & oImpresora.gPrnSaltoLinea
   Cabecera = Cabecera + "RUC: " + gsRUC + Space(82) & "Pag. " & Format(P, "000") & oImpresora.gPrnSaltoLinea
   'By Capi 01102008
   'Cabecera = Cabecera + BON & Centra(sTit, pnColPage) & BOFF & oImpresora.gPrnSaltoLinea
   Cabecera = Cabecera + Centra(sTit, pnColPage) & oImpresora.gPrnSaltoLinea
   If psSimbolo <> "" Then
      'By Capi 07102008
      'Cabecera = Cabecera + BON & Centra(" M O N E D A   " & IIf(psSimbolo = gcMN, "N A C I O N A L ", "E X T R A N J E R A "), pnColPage) & BOFF & oImpresora.gPrnSaltoLinea
      Cabecera = Cabecera + Centra(" M O N E D A   " & IIf(psSimbolo = gcMN, "N A C I O N A L ", "E X T R A N J E R A "), pnColPage) & oImpresora.gPrnSaltoLinea
   End If
   If sCabe <> "" Then
      Cabecera = Cabecera & sCabe
   End If
End Function

Public Function CabeceraCusco(sTit As String, P As Integer, psSimbolo As String, pnColPage As Integer, Optional sCabe As String = "", Optional sFecha As String = "", Optional psEmprLogo As String = " CMAC MAYNAS S.A. ") As String
Dim BON As String
Dim BOFF As String
Dim CON As String
Dim COFF As String

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

CabeceraCusco = ""
If sFecha = "" Then
   sFecha = Date
End If
   If P > 0 Then CabeceraCusco = oImpresora.gPrnSaltoPagina
   P = P + 1
   CabeceraCusco = CabeceraCusco + ImpreFormat(psEmprLogo, 80) & sFecha & "-" & Format(Time, "hh:mm:ss") & "            Pag. " & Format(P, "000") & oImpresora.gPrnSaltoLinea
   CabeceraCusco = CabeceraCusco + Space(72) & oImpresora.gPrnSaltoLinea
   'By capi 14112008
   'CabeceraCusco = CabeceraCusco + BON & Centra(sTit, pnColPage) & BOFF & oImpresora.gPrnSaltoLinea
   CabeceraCusco = CabeceraCusco + Centra(sTit, pnColPage) & oImpresora.gPrnSaltoLinea
   If psSimbolo <> "" Then
      'By capi 14112008
      'CabeceraCusco = CabeceraCusco + BON & Centra(" M O N E D A   " & IIf(psSimbolo = gcMN, "N A C I O N A L ", "E X T R A N J E R A "), pnColPage) & BOFF & oImpresora.gPrnSaltoLinea
      CabeceraCusco = CabeceraCusco + Centra(" M O N E D A   " & IIf(psSimbolo = gcMN, "N A C I O N A L ", "E X T R A N J E R A "), pnColPage) & oImpresora.gPrnSaltoLinea
   End If
   If sCabe <> "" Then
      CabeceraCusco = CabeceraCusco & sCabe
   End If
End Function





Public Function EmiteBoleta(ByVal sMsgProd As String, ByVal sMsgOpe As String, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double, _
            ByVal nIntMes As Double, ByVal nExtracto As Long, Optional bDocumento As Boolean = False, Optional nDocumento As TpoDoc = TpoDocCheque, _
            Optional sNroDoc As String = "", Optional dFechaValor As Date, Optional bImpSaldos As Boolean = True, _
            Optional pdFecsis As Date, Optional psNomAge As String = "", Optional psCodUser As String = "") As String

Dim bReImp As Boolean
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = sMsgProd
Dim clsMant As NCapMantenimiento
Set clsMant = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsMant.GetNombreTitulares(sCuenta))
Set clsMant = Nothing
Dim oImp As NContImprimir
Set oImp = New NContImprimir
bReImp = False
MsgBox "Se va imprimir Boleta  de Operacion por favor verifique su impresora por favor", vbExclamation, "Aviso"
Do
    If bDocumento Then
        Select Case nDocumento
            Case TpoDocCheque
                oImp.ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, gsFormatoFechaView), nSaldoDisp, nIntMes, "Fecha Valor", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecsis, psNomAge, psCodUser
            Case TpoDocNotaAbono, TpoDocNotaCargo
                oImp.ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecsis, psNomAge, psCodUser
            Case TpoDocOrdenPago
                oImp.ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecsis, psNomAge, psCodUser
        End Select
    End If
    If MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        bReImp = True
    Else
        bReImp = False
    End If
Loop Until Not bReImp
End Function
 
Public Sub GrabaArchivo(ByVal sCadena As String, ByVal sArchivo As String)
    Open sArchivo For Output As #1
    Print #1, sCadena
    Close #1
End Sub
