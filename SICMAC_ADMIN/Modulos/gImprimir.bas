Attribute VB_Name = "gImprimir"
Option Explicit
Public ArcSal As Integer
'Public sLpt As String
Public lnNumCopias As Integer
Public lbCancela As Boolean
Public TextoFinLen As Integer
'***************************************************
'* Inicia una impresión - Cabecera
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'Referencia : Global ArcSal As Integer
'***************************************************
Public Sub ImpreBegin(pbCondensado As Boolean, nlineas As Integer)
    ArcSal = FreeFile
    
    Open sLpt For Output As ArcSal
    Print #ArcSal, oImpresora.gPrnInicializa;             'Inicializa Impresora
    If pbCondensado Then
       Print #ArcSal, oImpresora.gPrnMargenIzq00;  'Tipo letra : 0,1,2 - Roman,SansS,Courier
       Print #ArcSal, oImpresora.gPrnTamLetra10CPI;            'Tamaño  : 80, 77, 103
       Print #ArcSal, oImpresora.gPrnCondensadaON;                        'Retorna al tipo de letra normal
    Else
       Print #ArcSal, oImpresora.gPrnCondensadaOFF;
    End If
    Print #ArcSal, oImpresora.gPrnEspaLineaN;             'Espaciamiento lineas 1/6 pulg.1
    Print #ArcSal, oImpresora.gPrnEspaLineaValor & Chr$(nlineas);  '   Chr$(nLineas); 'Longitud de página a 66 líneas
    If Not pbCondensado Then
       Print #ArcSal, oImpresora.gPrnTpoLetraCurier;  'Tipo letra : 0,1,2 - Roman,SansS,Courier
       Print #ArcSal, oImpresora.gPrnTamLetra10CPI;            'Tamaño  : 80, 77, 103
    End If
    Print #ArcSal, oImpresora.gPrnTpoLetraRoman1P;  'Draf : 1 pasada
   
End Sub
'***************************************************
'* Termina un impresión - Cola
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'Referencia : Global ArcSal As Integer
'***************************************************
Public Sub ImpreEnd()
    Print #ArcSal, oImpresora.gPrnSaltoPagina;    'Nueva página
    Print #ArcSal, oImpresora.gPrnCondensadaOFF;   'Retorna al tipo de letra normal
    Close ArcSal
End Sub
'***************************************************
'* Genera nueva página
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'Referencia : Global ArcSal As Integer
'***************************************************
Public Sub ImpreNewPage()
    Print #ArcSal, oImpresora.gPrnSaltoPagina;    'Nueva página
End Sub
'Prepara una cadena especial (cadena con caracteres con tilde y/o otros)
' para que se imprima en el modo FREEFILE.
Public Function ImpreCarEsp(ByVal vCadena As String) As String
    vCadena = Replace(vCadena, "á", Chr(160), , , vbTextCompare)
    vCadena = Replace(vCadena, "é", Chr(130), , , vbTextCompare)
    vCadena = Replace(vCadena, "í", Chr(161), , , vbTextCompare)
    vCadena = Replace(vCadena, "ó", Chr(162), , , vbTextCompare)
    vCadena = Replace(vCadena, "ú", Chr(163), , , vbTextCompare)
    vCadena = Replace(vCadena, "ñ", Chr(164), , , vbTextCompare)
    vCadena = Replace(vCadena, "Ñ", Chr(165), , , vbTextCompare)
    vCadena = Replace(vCadena, "°", Chr(248), , , vbTextCompare)
    vCadena = Replace(vCadena, "¦", Chr(179), , , vbTextCompare)
    ImpreCarEsp = vCadena
End Function
'Verifica la corrceta habilitación de la impresora
Public Function ImpreSensa() As Boolean
Dim lbArchAbierto As Boolean
On Error GoTo ControlError
    ArcSal = FreeFile
    lbArchAbierto = True
    Open sLpt For Output As ArcSal
    Print #ArcSal, oImpresora.gPrnInicializa;             'Inicializa Impresora
    Close ArcSal
    lbArchAbierto = False
    ImpreSensa = True
    Exit Function
ControlError:   ' Rutina de control de errores.
    If lbArchAbierto Then
        Close ArcSal
    End If
    MsgBox "Impresora no Encontrada ó Inactiva" & vbCr & "Verifique que la Conexión sea Correcta", vbExclamation, "Aviso de Precaución"
    ImpreSensa = False
End Function
Public Function PrnSet(Code As String, Optional nValor As Integer) As String
If nValor = 12 Or nValor = 10 Then
   nValor = nValor - 1
End If
Select Case Code
 Case "B+": PrnSet = oImpresora.gPrnBoldOFF 'Bold On
 Case "B-": PrnSet = oImpresora.gPrnBoldOFF 'Bold Off
 Case "U+": PrnSet = oImpresora.gPrnUnderLineONOFF   'Underline On
 Case "U-": PrnSet = oImpresora.gPrnUnderLineONOFF 'Underline Off
 Case "I+": PrnSet = oImpresora.gPrnItalicON  'Italic On
 Case "I-": PrnSet = oImpresora.gPrnItalicOFF  'Italic Off
 Case "W+": PrnSet = oImpresora.gPrnDblAnchoON  'Doble Ancho On
 Case "W-": PrnSet = oImpresora.gPrnDblAnchoOFF  'Doble Ancho Off
 Case "C+": PrnSet = oImpresora.gPrnCondensadaON  'Condensado On
 Case "C-": PrnSet = oImpresora.gPrnCondensadaOFF 'Condensado Off
 Case "Rm": PrnSet = oImpresora.gPrnTpoLetraRoman  'Roman
 Case "Ss": PrnSet = oImpresora.gPrnTpoLetraSansSerif  'Sans Serif
 Case "Co": PrnSet = oImpresora.gPrnTpoLetraCurier  'Courier
 Case "1.5": PrnSet = oImpresora.gPrnUnoMedioEspacio  ' 1 1/2 espacios
 Case "MI": PrnSet = oImpresora.gPrnMargenIzqCab & Chr$(nValor)  'Margen Izquierdo
 Case "MD": PrnSet = oImpresora.gPrnMargenDerCab & Chr$(nValor)   'Margen Derecho
 Case "10CPI": PrnSet = oImpresora.gPrnTamLetra12CPI
 Case "12CPI": PrnSet = oImpresora.gPrnTamLetra10CPI
 Case "15CPI": PrnSet = oImpresora.gPrnTamLetra15CPI
 Case "EspN": PrnSet = oImpresora.gPrnEspaLineaN      'Espaciado Normal 4.5/72
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
Dim X As Integer

'On Error GoTo ErrHandler
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
        For X = 1 To Len(vParEnt)
            If Mid(vParEnt, X, 1) = "," Then pLongitudEntera = pLongitudEntera - 1
        Next X
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
'Exit Function

'ErrHandler:     ' Errores obtenidos
'    MsgBox " Operación no válida " & vbCr & _
        " Error " & Err.Number & " : " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " ! Aviso ! "
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

Public Function Linea(psVarImpre As String, psTexto As String, Optional pnLineas As Integer = 1) As String
Dim K As Integer
psVarImpre = psVarImpre & psTexto
For K = 1 To pnLineas
   psVarImpre = psVarImpre & oImpresora.gPrnSaltoLinea
Next
End Function
'   ------------------------------------------------------------
'   Función     :   FillText
'   Propósito   :   Rellena a la derecha un campo texto con un
'                   caracter especificado
'   Uso         :   Reportes en general
'   Parámetro(s):   pnCampos -> Importe
'                   pnLongit -> Longitud del Campo
'   Creado      :   02/07/1999  -   FAOS
'   Modificado  :   02/07/1999  -   FAOS
'   ------------------------------------------------------------
'   Formato: FillText("HOLA", 20, " ")
Public Function FillText(psCadena As String, pnLenTex As Integer, ChrFil As String) As String
    If pnLenTex > Len(Trim(psCadena)) Then
       FillText = Trim(psCadena) & String((pnLenTex - Len(Trim(psCadena))), ChrFil)
    End If
End Function

Public Function CabeceraPagina(ByVal psTitulo As String, pnPagina As Long, pnItem As Long, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, Optional psMoneda As String = "1") As String
    Dim lsC1 As String
    Dim lsC2 As String
    Dim lsC3 As String
    Dim lsCadena As String
    
    If pnItem >= 66 Then
        pnItem = 0
    End If
    
    pnPagina = pnPagina + 1
    pnItem = 5
    lsCadena = ""

    lsC1 = Format(pgdFecSis, gsFormatoFechaView)
    lsC2 = Format(Time, "hh:mm:ss AMPM")
    lsC3 = "PAGINA Nro. " & Format(pnPagina, "000")
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & pgsEmpresa & Space(39 - Len(lsC3) + 10 - Len(Trim(pgsEmpresa))) & lsC3 & Space(70 - Len(lsC1)) & lsC1 & oImpresora.gPrnSaltoLinea
  
    If psMoneda = "" Then
        lsCadena = lsCadena & pgsNomAge & "      " & Space(114 - Len(lsC2) - Len(lsC2) + 10 - Len(pgsNomAge)) & lsC2 & oImpresora.gPrnSaltoLinea
    ElseIf psMoneda = "1" Then
        '''lsCadena = lsCadena & Trim(pgsNomAge) & "- Soles" + Space(109 - Len("- Soles") - Len(lsC2) + 10 - Len(Trim(pgsNomAge))) & lsC2 & oImpresora.gPrnSaltoLinea 'marg ers044-2016
        lsCadena = lsCadena & Trim(pgsNomAge) & "- " & StrConv(gcPEN_PLURAL, vbProperCase) + Space(109 - Len("- " & StrConv(gcPEN_PLURAL, vbProperCase)) - Len(lsC2) + 10 - Len(Trim(pgsNomAge))) & lsC2 & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    Else
        lsCadena = lsCadena & Trim(pgsNomAge) & "- Dolares" + Space(109 - Len("- Dolares") - Len(lsC2) + 10 - Len(Trim(pgsNomAge))) & lsC2 & oImpresora.gPrnSaltoLinea
    End If
    
    lsCadena = lsCadena & CentrarCadena(psTitulo, 104) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        
    CabeceraPagina = lsCadena
End Function

Public Function CabeceraPaginaHE(ByVal psTitulo As String, pnPagina As Long, pnItem As Long, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, Optional psMoneda As String = "1") As String
    Dim lsC1 As String
    Dim lsC2 As String
    Dim lsC3 As String
    Dim lsCadena As String
    
    If pnItem >= 66 Then
        pnItem = 0
    End If
    
    pnPagina = pnPagina + 1
    pnItem = 5
    lsCadena = ""

    lsC1 = Format(pgdFecSis, gsFormatoFechaView)
    lsC2 = Format(Time, "hh:mm:ss AMPM")
    lsC3 = "PAGINA Nro. " & Format(pnPagina, "000")
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & pgsEmpresa & Space(39 - Len(lsC3) + 10 - Len(Trim(pgsEmpresa))) & lsC3 & Space(70 - Len(lsC1)) & lsC1 & oImpresora.gPrnSaltoLinea
  
    If psMoneda = "" Then
        lsCadena = lsCadena & pgsNomAge & "      " & Space(114 - Len(lsC2) - Len(lsC2) + 10 - Len(pgsNomAge)) & lsC2 & oImpresora.gPrnSaltoLinea
    ElseIf psMoneda = "1" Then
        '''lsCadena = lsCadena & Trim(pgsNomAge) & "- Soles" + Space(109 - Len("- Soles") - Len(lsC2) + 10 - Len(Trim(pgsNomAge))) & lsC2 & oImpresora.gPrnSaltoLinea 'marg ers044-2016
        lsCadena = lsCadena & Trim(pgsNomAge) & "- " & StrConv(gcPEN_PLURAL, vbProperCase) + Space(109 - Len("- " & StrConv(gcPEN_PLURAL, vbProperCase)) - Len(lsC2) + 10 - Len(Trim(pgsNomAge))) & lsC2 & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    Else
        lsCadena = lsCadena & Trim(pgsNomAge) & "- Dolares" + Space(109 - Len("- Dolares") - Len(lsC2) + 10 - Len(Trim(pgsNomAge))) & lsC2 & oImpresora.gPrnSaltoLinea
    End If
    
    'lsCadena = lsCadena & CentrarCadena(psTitulo, 104) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & psTitulo & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        
    CabeceraPaginaHE = lsCadena
End Function
Public Function Encabezado(psCadena As String, pnItem As Long, Optional pbLineaSimple As Boolean = True) As String
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
    If pbLineaSimple Then
        lsLineas = String(lnTotalLinea + 1, "=") & oImpresora.gPrnSaltoLinea
    Else
        lsLineas = String(lnTotalLinea + 1, "-") & oImpresora.gPrnSaltoLinea
    End If
    
    lsResultado = lsLineas + lsResultado + lsLineas
    
    Encabezado = lsResultado
End Function

Public Function Encabezado1(psCadena As String, pnItem As Long, Optional pbLineaSimple As Boolean = True) As String
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
    If pbLineaSimple Then
        lsLineas = Space(2) & String(lnTotalLinea + 1, "=") & oImpresora.gPrnSaltoLinea
    Else
        lsLineas = Space(2) & String(lnTotalLinea + 1, "-") & oImpresora.gPrnSaltoLinea
    End If
    
    lsResultado = Space(2) & lsLineas + Space(5) & lsResultado + Space(2) & lsLineas
    
    Encabezado1 = lsResultado
End Function

Public Function CabeceraPagina1(ByVal psTitulo As String, pnPagina As Long, pnItem As Long, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, Optional psMoneda As String = "1") As String
    Dim lsC1 As String
    Dim lsC2 As String
    Dim lsC3 As String
    Dim lsCadena As String
    
    If pnItem >= 66 Then
        pnItem = 0
    End If
    
    pnPagina = pnPagina + 1
    pnItem = 5
    lsCadena = ""

    lsC1 = Format(pgdFecSis, gsFormatoFechaView)
    lsC2 = Format(Time, "hh:mm:ss AMPM")
    lsC3 = "PAGINA Nro. " & Format(pnPagina, "000")
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & pgsEmpresa & Space(39 - Len(lsC3) + 10 - Len(Trim(pgsEmpresa))) & lsC3 & Space(70 - Len(lsC1)) & lsC1 & oImpresora.gPrnSaltoLinea
  
    If psMoneda = "" Then
        lsCadena = lsCadena & Space(5) & pgsNomAge & "      " & Space(114 - Len(lsC2) - Len(lsC2) + 10 - Len(pgsNomAge)) & lsC2 & oImpresora.gPrnSaltoLinea
    ElseIf psMoneda = "1" Then
        '''lsCadena = lsCadena & Space(5) & Trim(pgsNomAge) & "- Soles" + Space(109 - Len("- Soles") - Len(lsC2) + 10 - Len(Trim(pgsNomAge))) & lsC2 & oImpresora.gPrnSaltoLinea 'marg ers044-2016
        lsCadena = lsCadena & Space(5) & Trim(pgsNomAge) & "- " & StrConv(gcPEN_PLURAL, vbProperCase) + Space(109 - Len(" - " & StrConv(gcPEN_PLURAL, vbProperCase)) - Len(lsC2) + 10 - Len(Trim(pgsNomAge))) & lsC2 & oImpresora.gPrnSaltoLinea 'marg ers044-2016"
    Else
        lsCadena = lsCadena & Space(5) & Trim(pgsNomAge) & "- Dolares" + Space(109 - Len("- Dolares") - Len(lsC2) + 10 - Len(Trim(pgsNomAge))) & lsC2 & oImpresora.gPrnSaltoLinea
    End If
    
    lsCadena = lsCadena & CentrarCadena(psTitulo, 104) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        
    CabeceraPagina1 = lsCadena
End Function


'***************************************************
'* Funcion que centra una cadena en una dimensión dada, si la cadena es
'* menor a esta dimensión la centra y rellena los espacios del caracter
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
' Definición de Cabecera 1
  sImpre = PrnSet("Co") & COFF
  For N = 1 To nLin
      sImpre = sImpre & oImpresora.gPrnSaltoLinea
  Next
  If psMovNro <> "" And lOpe Then
     sImpre = sImpre & BON & ImpreFormat(psEmpresaLogo, 55) & PrnSet("I+") & "Operación : " & psOpeCod & " # " & psMovNro & PrnSet("I-") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
  Else
     sImpre = sImpre & BON & ImpreFormat(psEmpresaLogo, 54) & pdFecha & "-" & Time & BOFF & oImpresora.gPrnSaltoLinea
  End If
  sImpre = sImpre & String(Int((pnColPage - Len(sTit)) / 2), " ")
  sImpre = sImpre & BON & sTit & BOFF & oImpresora.gPrnSaltoLinea
  If lMoneda Then
     sTexto = "M O N E D A   " & IIf(Mid(psOpeCod, 3, 1) = gMonedaNacional, "N A C I O N A L", "E X T R A N J E R A")
     sImpre = sImpre & String(Int((pnColPage - Len(sTexto)) / 2), " ")
     sImpre = sImpre & BON & sTexto & BOFF & oImpresora.gPrnSaltoLinea
  End If
  sImpre = sImpre & oImpresora.gPrnSaltoLinea
  ImpreCabAsiento = sImpre
End Function

Public Function ImpreGlosa(Optional sTitGlosa As String = "  GLOSA      : ", Optional nCols As Integer = 0, Optional lbEnterFinal As Boolean = True, Optional ByRef nLin As Integer = 0) As String
Dim sImpre As String
Dim sTexto As String, N As Integer
Dim nLen As Integer
  nLen = Len(sTitGlosa)
  sTexto = JustificaTexto(gsGlosa, IIf(nCols = 0, gnColPage, nCols) - nLen)
  sImpre = sTitGlosa
  N = 0
  Do While True
     N = InStr(sTexto, oImpresora.gPrnSaltoLinea)
     If N > 0 Then
        sImpre = sImpre & Mid(sTexto, 1, N - 1) & oImpresora.gPrnSaltoLinea & Space(nLen)
        sTexto = Mid(sTexto, N + 1, Len(sTexto))
        nLin = nLin + 1
     End If
     If N = 0 Then
        If sTexto <> "" Then
            sImpre = sImpre & Justifica(sTexto, IIf(nCols = 0, gnColPage, nCols) - nLen) & IIf(lbEnterFinal, oImpresora.gPrnSaltoLinea, "")
        End If
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
Dim letra As String * 1, I As Integer, K As Integer, N As Integer
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
   letra = Mid(sTemp, I, 1)
   If letra = Chr$(27) Then
      vTextFin = vTextFin & letra & Mid(sTemp, I + 1, 1)
      I = I + 1
      K = K + 1
      nAncho1 = nAncho1 + 2
   Else
      If Asc(letra) <> 13 And Asc(letra) <> 10 Then
         If K > nAncho1 Then
            m = 0
            If Mid(sTemp, I, 1) = Chr(32) Then
               vTextFin = Trim(vTextFin)
            Else
               m = InStrRev(vTextFin, " ", , vbTextCompare)
               If m = 0 Then m = 1
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
            letra = ""
            K = 0
         Else
            vTextFin = vTextFin & letra
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
         letra = ""
         K = 0
      End If
   End If
Loop
JustificaTexto = JustificaTexto & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin))
End Function

Public Function JustificaTextoCadena(sTemp As String, lnColPage As Integer, Optional lsEspIzq As Integer = 0) As String
Dim vTextFin As String
Dim letra As String * 1, I As Integer, K As Integer, N As Integer
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
   letra = Mid(sTemp, I, 1)
   If letra = Chr$(27) Then
      vTextFin = vTextFin & letra & Mid(sTemp, I + 1, 1)
      I = I + 1
      K = K + 1
      nAncho1 = nAncho1 + 2
   Else
      If Asc(letra) <> 13 And Asc(letra) <> 10 Then
         If K > nAncho1 Then
            m = 0
            If Mid(sTemp, I, 1) = Chr(32) Then
               vTextFin = Trim(vTextFin)
            Else
               m = InStrRev(vTextFin, " ", , vbTextCompare)
               If m = 0 Then m = 1
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
            JustificaTextoCadena = JustificaTextoCadena & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin))
            nAncho1 = lnColPage
            vTextFin = ""
            letra = ""
            K = 0
         Else
            vTextFin = vTextFin & letra
         End If
      Else
        If I < N Then
            If Asc(Mid(sTemp, I + 1, 1)) = 13 Or Asc(Mid(sTemp, I + 1, 1)) = 10 Then
                I = I + 1
            End If
        End If
         JustificaTextoCadena = JustificaTextoCadena & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin)) & oImpresora.gPrnSaltoLinea
         nAncho1 = lnColPage
         vTextFin = ""
         letra = ""
         K = 0
      End If
   End If
Loop
JustificaTextoCadena = JustificaTextoCadena & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin))
End Function

Public Function JustificaTextoCadenaOrdenCompra(sTemp As String, lnColPage As Integer, Optional lsEspIzq As Integer = 0) As String
Dim vTextFin As String
Dim letra As String * 1, I As Integer, K As Integer, N As Integer
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
   letra = Mid(sTemp, I, 1)
   If letra = Chr$(27) Then
      vTextFin = vTextFin & letra & Mid(sTemp, I + 1, 1)
      I = I + 1
      K = K + 1
      nAncho1 = nAncho1 + 2
   Else
      If Asc(letra) <> 13 And Asc(letra) <> 10 Then
         If K > nAncho1 Then
            m = 0
            If Mid(sTemp, I, 1) = Chr(32) Then
               vTextFin = Trim(vTextFin)
            Else
               m = InStrRev(vTextFin, " ", , vbTextCompare)
               If m = 0 Then m = 1
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
            JustificaTextoCadenaOrdenCompra = JustificaTextoCadenaOrdenCompra & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin))
            nAncho1 = lnColPage
            vTextFin = ""
            letra = ""
            K = 0
            
            
            
         Else
            vTextFin = vTextFin & letra
            TextoFinLen = Len(Trim(vTextFin))
         End If
      Else
        If I < N Then
            If Asc(Mid(sTemp, I + 1, 1)) = 13 Or Asc(Mid(sTemp, I + 1, 1)) = 10 Then
                I = I + 1
            End If
        End If
         JustificaTextoCadenaOrdenCompra = JustificaTextoCadenaOrdenCompra & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin)) & oImpresora.gPrnSaltoLinea
         nAncho1 = lnColPage
         vTextFin = ""
         letra = ""
         K = 0
      End If
   End If
Loop
JustificaTextoCadenaOrdenCompra = JustificaTextoCadenaOrdenCompra & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin))
End Function
'PASIERS0772014
Public Function JustificaTextoCadenaPASI(sTemp As String, lnColPage As Integer, Optional lsEspIzq As Integer = 0) As String
Dim vTextFin As String
Dim letra As String * 1, I As Integer, K As Integer, N As Integer
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
   letra = Mid(sTemp, I, 1)
   If letra = Chr$(27) Then
      vTextFin = vTextFin & letra & Mid(sTemp, I + 1, 1)
      I = I + 1
      K = K + 1
      nAncho1 = nAncho1 + 2
   Else
      If Asc(letra) <> 13 And Asc(letra) <> 10 Then
         If K > nAncho1 Then
            m = 0
            If Mid(sTemp, I, 1) = Chr(32) Then
               vTextFin = Trim(vTextFin)
            Else
               m = InStrRev(vTextFin, " ", , vbTextCompare)
               If m = 0 Then m = 1
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
            JustificaTextoCadenaPASI = JustificaTextoCadenaPASI & Space(lsEspIzq) & Trim((vTextFin))
            nAncho1 = lnColPage
            vTextFin = ""
            letra = ""
            K = 0
            
            
            
         Else
            vTextFin = vTextFin & letra
            TextoFinLen = Len(Trim(vTextFin))
         End If
      Else
        If I < N Then
            If Asc(Mid(sTemp, I + 1, 1)) = 13 Or Asc(Mid(sTemp, I + 1, 1)) = 10 Then
                I = I + 1
            End If
        End If
         JustificaTextoCadenaPASI = JustificaTextoCadenaPASI & Space(lsEspIzq) & Trim(vTextFin) & oImpresora.gPrnSaltoLinea
         nAncho1 = lnColPage
         vTextFin = ""
         letra = ""
         K = 0
      End If
   End If
Loop
JustificaTextoCadenaPASI = JustificaTextoCadenaPASI & Space(lsEspIzq) & Trim((vTextFin))
End Function
'end PASI


Public Function Justifica(sCad As String, nAncho As Integer)
Justifica = Mid(sCad & Space(nAncho), 1, nAncho)
End Function
Public Function PrnVal(pnVal As Currency, pnLen As Integer, pnDec As Integer, Optional lCero As Boolean = True) As String
Dim sFormat As String
 sFormat = "###,###,###,##0" & IIf(pnDec > 0, "." & String(pnDec, "0"), "")
 PrnVal = Right(Space(pnLen) & IIf(Not IsNull(pnVal) And (pnVal <> 0 Or lCero), Format(pnVal, sFormat), ""), pnLen)
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
         sPie = sPie + Centra(" LE _________________ ", nLenPie)
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
lsPiePag = lsPiePag + "" + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
lsPiePag = lsPiePag + sPiR + oImpresora.gPrnSaltoLinea
lsPiePag = lsPiePag + sPie + oImpresora.gPrnSaltoLinea
ImprePiePag = lsPiePag
End Function
Public Function Cabecera(sTit As String, P As Integer, psSimbolo As String, pnColPage As Integer, Optional sCabe As String = "", Optional sFecha As String = "") As String
Dim BON As String
Dim BOFF As String
Dim CON As String
Dim COFF As String

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

Cabecera = ""
If sFecha = "" Then
   sFecha = Date
End If
   If P > 0 Then Cabecera = oImpresora.gPrnSaltoPagina
   P = P + 1
   Cabecera = Cabecera + " CMAC - TRUJILLO " & Space(42) & sFecha & " - " & Format(Time, "hh:mm:ss") & oImpresora.gPrnSaltoLinea
   Cabecera = Cabecera + Space(72) & "Pag. " & Format(P, "000") & oImpresora.gPrnSaltoLinea
   Cabecera = Cabecera + BON & Centra(sTit, pnColPage) & BOFF & oImpresora.gPrnSaltoLinea
   If psSimbolo <> "" Then
      '''Cabecera = Cabecera + BON & Centra(" M O N E D A   " & IIf(psSimbolo = "S/.", "N A C I O N A L ", "E X T R A N J E R A "), pnColPage) & BOFF & oImpresora.gPrnSaltoLinea 'marg ers044-2016
      Cabecera = Cabecera + BON & Centra(" M O N E D A   " & IIf(psSimbolo = gcPEN_SIMBOLO, "N A C I O N A L ", "E X T R A N J E R A "), pnColPage) & BOFF & oImpresora.gPrnSaltoLinea 'marg ers044-2016
   End If
   If sCabe <> "" Then
      Cabecera = Cabecera & sCabe
   End If
End Function
Public Function EmiteBoleta(ByVal sMsgProd As String, ByVal sMsgOpe As String, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double, _
            ByVal nIntMes As Double, ByVal nExtracto As Long, Optional bDocumento As Boolean = False, Optional nDocumento As TpoDoc = TpoDocCheque, _
            Optional sNroDoc As String = "", Optional dFechaValor As Date, Optional bImpSaldos As Boolean = True, _
            Optional pdFecSis As Date, Optional psNomAge As String = "", Optional psCodUser As String = "") As String

Dim bReImp As Boolean
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
'''sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES") 'marg ers044-2016
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES")
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
            '    oImp.ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, gsFormatoFechaView), nSaldoDisp, nIntMes, "Fecha Valor", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecsis, psNomAge, psCodUser
            Case TpoDocNotaAbono, TpoDocNotaCargo
            '    oImp.ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecsis, psNomAge, psCodUser
            Case TpoDocOrdenPago
            '    oImp.ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecsis, psNomAge, psCodUser
        End Select
    End If
    If MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        bReImp = True
    Else
        bReImp = False
    End If
Loop Until Not bReImp
End Function




