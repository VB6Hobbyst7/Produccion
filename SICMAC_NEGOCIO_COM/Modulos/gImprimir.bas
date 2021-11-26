Attribute VB_Name = "gImprimir"
Option Explicit
Public ArcSal As Integer
'Public sLpt As String
Public lnNumCopias As Integer
Public lbCancela As Boolean


'***************************************************
'* Inicia una impresión - Cabecera
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'Referencia : Global ArcSal As Integer
'***************************************************
Public Sub ImpreBegin(pbCondensado As Boolean, nLineas As Integer)
    ArcSal = FreeFile
    
    Open sLpt For Output As ArcSal
    Print #ArcSal, Chr$(27) & Chr$(64);            'Inicializa Impresora
    If pbCondensado Then
       Print #ArcSal, Chr$(27) & Chr$(108) & Chr$(0); 'Tipo letra : 0,1,2 - Roman,SansS,Courier
       Print #ArcSal, Chr$(27) & Chr$(77);            'Tamaño  : 80, 77, 103
       Print #ArcSal, Chr$(15);                       'Retorna al tipo de letra normal
    Else
       Print #ArcSal, Chr$(18);
    End If
    Print #ArcSal, Chr$(27) & Chr$(50);            'Espaciamiento lineas 1/6 pulg.1
    Print #ArcSal, Chr$(27) & Chr$(67) & Chr$(nLineas); '   Chr$(nLineas); 'Longitud de página a 66 líneas
    If Not pbCondensado Then
       Print #ArcSal, Chr$(27) & Chr$(107) & Chr$(2); 'Tipo letra : 0,1,2 - Roman,SansS,Courier
       Print #ArcSal, Chr$(27) & Chr$(77);            'Tamaño  : 80, 77, 103
    End If
    Print #ArcSal, Chr$(27) & Chr$(120) & Chr$(0);  'Draf : 1 pasada
   
End Sub
'***************************************************
'* Termina un impresión - Cola
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'Referencia : Global ArcSal As Integer
'***************************************************
Public Sub ImpreEnd()
    Print #ArcSal, Chr$(12);   'Nueva página
    Print #ArcSal, Chr$(18);   'Retorna al tipo de letra normal
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
    Print #ArcSal, Chr$(12);   'Nueva página
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
Dim oImp As ContsImp.clsConstImp
On Error GoTo ControlError
    ArcSal = FreeFile
    lbArchAbierto = True
    Set oImp = New ContsImp.clsConstImp
    Open sLpt For Output As ArcSal
    Print #ArcSal, oImp.gPrnInicializa;              'Inicializa Impresora
    Close ArcSal
    lbArchAbierto = False
    ImpreSensa = True
    Set oImp = Nothing
    Exit Function
ControlError:   ' Rutina de control de errores.
    If lbArchAbierto Then
        Close ArcSal
    End If
    MsgBox "Impresora no Encontrada ó Inactiva" & vbCr & "Verifique que la Conexión sea Correcta", vbExclamation, "Aviso de Precaución"
    ImpreSensa = False
    Set oImp = Nothing
End Function


Public Function PrnSet(Code As String, Optional nValor As Integer) As String
If nValor = 12 Or nValor = 10 Then
   nValor = nValor - 1
End If
Select Case Code
 Case "B+": PrnSet = Chr$(27) & Chr$(69) 'Bold On
 Case "B-": PrnSet = Chr$(27) & Chr$(70) 'Bold Off
 Case "U+": PrnSet = Chr$(27) & Chr$(45)  'Underline On
 Case "U-": PrnSet = Chr$(27) & Chr$(46) 'Underline Off
 Case "I+": PrnSet = Chr$(27) & Chr$(52) 'Italic On
 Case "I-": PrnSet = Chr$(27) & Chr$(53) 'Italic Off
 Case "W+": PrnSet = Chr$(27) & Chr$(87) 'Doble Ancho On
 Case "W-": PrnSet = Chr$(27) & Chr$(20) 'Doble Ancho Off
 Case "C+": PrnSet = Chr$(27) & Chr$(15) 'Condensado On
 Case "C-": PrnSet = Chr$(27) & Chr$(18) 'Condensado Off
 Case "Rm": PrnSet = Chr$(27) & Chr$(107) & Chr$(0) 'Roman
 Case "Ss": PrnSet = Chr$(27) & Chr$(107) & Chr$(1) 'Sans Serif
 Case "Co": PrnSet = Chr$(27) & Chr$(107) & Chr$(2) 'Courier
 Case "1.5": PrnSet = Chr$(27) & Chr$(48) ' 1 1/2 espacios
 Case "MI": PrnSet = Chr$(27) & Chr$(108) & Chr$(nValor) 'Margen Izquierdo
 Case "MD": PrnSet = Chr$(27) & Chr$(81) & Chr$(nValor)  'Margen Derecho
 Case "10CPI": PrnSet = Chr$(27) & Chr$(80)
 Case "12CPI": PrnSet = Chr$(27) & Chr$(77)
 Case "15CPI": PrnSet = Chr$(27) & Chr$(103)
 Case "EspN": PrnSet = Chr$(27) & Chr$(50)     'Espaciado Normal 4.5/72
 Case "Esp":  PrnSet = Chr$(27) & Chr$(65) & Chr$(nValor) 'Espaciado nValor/72 pulg.
End Select
End Function

'Devuelve un string formateado de acuerdo a los parametros ingresados
' se utiliza con numeros y caracteres
Public Function ImpreFormat(ByVal pNumero As Variant, ByVal pLongitudEntera As Integer, _
        Optional ByVal pLongitudDecimal As Integer = 2, _
        Optional ByVal pMoneda As Boolean = False) As String
Dim vPosPto As Integer
Dim vParEnt As String
Dim vParDec As String
Dim vLonEnt As Integer
Dim vLonDec As Integer
Dim X As Integer

On Error GoTo ErrHandler
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
Exit Function

ErrHandler:     ' Errores obtenidos
    MsgBox " Operación no válida " & vbCr & _
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
      Linea lsImpre, Chr(12)
   End If
   pnNumPag = pnNumPag + 1
   
   Linea lsImpre, Space(40) & FillText(UCase(psEmpresa), pnColRep - 27, " ") & "Fecha : " & psFecha & " " & Format(Now(), "hh:mm:ss")
   Linea lsImpre, Space(40) & FillText(UCase(Trim(psAgencia)) & " - " & psSeccion & lsMoneda, pnColRep - 27, " ") + PrnSet("I+") + "Pagina: " & Format(pnNumPag, "000") + PrnSet("I-")
   Linea lsImpre, ""
'   Linea lsImpre, PrnSet("B+") + Centra(Space(40) & psTitulo1, pnColRep) + PrnSet("B-")
'   Linea lsImpre, PrnSet("B+") + Centra(Space(40) & psTitulo2, pnColRep) + PrnSet("B-"), 2
   Linea lsImpre, PrnSet("B+") + psTitulo1 + PrnSet("B-")
   Linea lsImpre, PrnSet("B+") + psTitulo2 + PrnSet("B-"), 2
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
   psVarImpre = psVarImpre & Chr(10)
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
    Else
       FillText = Trim(Left(psCadena, pnLenTex)) & String((pnLenTex - Len(Trim(psCadena))), ChrFil)
    End If
End Function

Public Function CabeceraPagina(ByVal psTitulo As String, pnPagina As Long, pnItem As Long, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, Optional psMoneda As Moneda = gMonedaNacional, Optional pbConMoneda As Boolean = True) As String
    Dim lsC1 As String
    Dim lsC2 As String
    Dim lsC3 As String
    Dim lscadena As String
    
    If pnItem >= 66 Then
        pnItem = 0
    End If
    
    pnPagina = pnPagina + 1
    pnItem = 5
    lscadena = ""

    lsC1 = Format(pgdFecSis, gsFormatoFechaView)
    lsC2 = Format(Time, "hh:mm:ss AMPM")
    lsC3 = "PAGINA Nro. " & Format(pnPagina, "000")
    lscadena = lscadena & Chr(10)
    lscadena = lscadena & pgsEmpresa & Space(39 - Len(lsC3) + 10 - Len(pgsEmpresa)) & lsC3 & Space(70 - Len(lsC1)) & lsC1 & Chr(10)
  
    If Not pbConMoneda Then
        lscadena = lscadena & pgsNomAge & "      " & Space(114 - Len(lsC2) - Len(lsC2) + 10 - Len(pgsNomAge)) & lsC2 & Chr(10)
    ElseIf psMoneda = gMonedaNacional Then
        lscadena = lscadena & Trim(pgsNomAge) & "- Soles" + Space(109 - Len("- Soles") - Len(lsC2) + 10 - Len(Trim(pgsNomAge))) & lsC2 & Chr(10)
    Else
        lscadena = lscadena & Trim(pgsNomAge) & "- Dolares" + Space(109 - Len("- Dolares") - Len(lsC2) + 10 - Len(Trim(pgsNomAge))) & lsC2 & Chr(10)
    End If
    
    lscadena = lscadena & CentrarCadena(psTitulo, 104) & Chr(10) & Chr(10)
        
    CabeceraPagina = lscadena
End Function


Public Function Encabezado(psCadena As String, pnItem As Long, Optional pbLineaSimple As Boolean = True) As String
    Dim lscadena As String
    Dim lsCampo As String
    Dim lnLonCampo As Long
    Dim lnTotalLinea As Long
    Dim lnPos As Long
    Dim lsResultado As String
    Dim i As Long
    Dim lsLineas As String
    
    lsResultado = ""
    lnTotalLinea = 0
        
    lscadena = psCadena
    pnItem = pnItem + 3
    
    While lscadena <> ""
        lnPos = InStr(1, lscadena, ";", vbTextCompare)
        lsCampo = Left(lscadena, lnPos - 1)
        lscadena = Mid(lscadena, lnPos + 1)
        lnPos = InStr(1, lscadena, ";", vbTextCompare)
        
        lnLonCampo = CCur(Left(lscadena, lnPos - 1))
        lscadena = Mid(lscadena, lnPos + 1)
        lnTotalLinea = lnTotalLinea + lnLonCampo
        
        If lnLonCampo - Len(lsCampo) > 0 Then
            lsResultado = lsResultado & Space(lnLonCampo - Len(lsCampo)) & lsCampo
        Else
            lsResultado = lsResultado & lsCampo
        End If
    Wend
        
    lsResultado = lsResultado & Chr(10)
    If pbLineaSimple Then
        lsLineas = String(lnTotalLinea + 1, "=") & Chr(10)
    Else
        lsLineas = String(lnTotalLinea + 1, "-") & Chr(10)
    End If
    
    lsResultado = lsLineas + lsResultado + lsLineas
    
    Encabezado = lsResultado
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
Dim Con As String
BON = PrnSet("B+")
BOFF = PrnSet("B-")
Con = PrnSet("C+")
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
      sImpre = sImpre & Chr$(10)
  Next
  If psMovNro <> "" And lOpe Then
     sImpre = sImpre & BON & ImpreFormat(psEmpresaLogo, 55) & PrnSet("I+") & "Operación : " & psOpeCod & " # " & psMovNro & PrnSet("I-") & BOFF & Chr$(10) & Chr$(10)
  Else
     sImpre = sImpre & BON & ImpreFormat(psEmpresaLogo, 54) & pdFecha & "-" & Time & BOFF & Chr(10)
  End If
  sImpre = sImpre & String(Int((pnColPage - Len(sTit)) / 2), " ")
  sImpre = sImpre & BON & sTit & BOFF & Chr$(10)
  If lMoneda Then
     sTexto = "M O N E D A   " & IIf(Mid(psOpeCod, 3, 1) = gMonedaNacional, "N A C I O N A L", "E X T R A N J E R A")
     sImpre = sImpre & String(Int((pnColPage - Len(sTexto)) / 2), " ")
     sImpre = sImpre & BON & sTexto & BOFF & Chr$(10)
  End If
  sImpre = sImpre & Chr$(10)
  ImpreCabAsiento = sImpre
End Function



Public Function ImpreGlosa(psGlosa As String, pnColPage As Integer, Optional psTitGlosa As String = "  GLOSA      : ", Optional pnCols As Integer = 0) As String
Dim sImpre As String
Dim sTexto As String, N As Integer
Dim nLen As Integer
  nLen = Len(psTitGlosa)
  sTexto = JustificaTexto(psGlosa, IIf(pnCols = 0, pnColPage, pnCols) - nLen)
  sImpre = psTitGlosa
  N = 0
  Do While True
     N = InStr(sTexto, Chr$(10))
     If N > 0 Then
        sImpre = sImpre & Mid(sTexto, 1, N - 1) & Chr$(10) & Space(nLen)
        sTexto = Mid(sTexto, N + 1, Len(sTexto))
     End If
     If N = 0 Then
        sImpre = sImpre & sTexto & Chr$(10)
        Exit Do
     End If
  Loop
  ImpreGlosa = sImpre
End Function
Public Function Centra(psCad As String, Optional pnAncho As Integer = 80) As String
Dim N As Integer, M As Integer, i As Integer
N = Len(Trim(psCad))
M = (pnAncho - N) / 2
Centra = ""
If pnAncho < M + N Then
   pnAncho = M + N
End If
Centra = Space(M) & Trim(psCad) & Space(pnAncho - M - N)
End Function
'
Public Function JustificaTexto(ByVal sTemp As String, ByVal lnColPage As Integer, Optional lsEspIzq As Integer = 0) As String
Dim vTextFin As String
Dim Letra As String * 1, i As Integer, K As Integer, N As Integer
Dim nVeces As Long, M As Integer, Fin As Integer, Ini As Integer
Dim nAncho1 As Integer, nSpa As Integer
i = 0
K = 0
N = Len(sTemp)
nAncho1 = lnColPage
Do While i <= N
   K = K + 1
   i = i + 1
   If i > N Then
      Exit Do
   End If
   Letra = Mid(sTemp, i, 1)
   If Letra = Chr$(27) Then
      vTextFin = vTextFin & Letra & Mid(sTemp, i + 1, 1)
      i = i + 1
      K = K + 1
      nAncho1 = nAncho1 + 2
   Else
      If Asc(Letra) <> 13 And Asc(Letra) <> 10 Then
         If K > nAncho1 Then
            M = 0
            If Mid(sTemp, i, 1) = Chr(32) Then
               vTextFin = Trim(vTextFin)
            Else
               M = InStrRev(vTextFin, " ", , vbTextCompare)
               If M = 0 Then M = 1
               If InStr(Mid(vTextFin, M, Len(vTextFin)), Chr$(27)) Then
                  nAncho1 = nAncho1 - 2
               End If
               i = i - (nAncho1 + 1 - M)
               vTextFin = Mid(vTextFin, 1, M - 1)
            End If
            nSpa = nAncho1 - Len(Trim(vTextFin))
            vTextFin = Trim(vTextFin)
            If nSpa <> 0 Then
               Fin = 1
               nVeces = 0
               M = 1
               Do While M <= nSpa
                  Ini = InStr(Fin, vTextFin, " ", vbTextCompare)
                  If Ini = 0 Then
                     Fin = 1
                     nVeces = nVeces + 1
                     M = M + 1
                  Else
                      vTextFin = Mid(vTextFin, 1, Ini) & " " & RTrim(Mid(vTextFin, Ini + 1, nAncho1))
                      Fin = Ini + 2 + nVeces
                      M = M + 1
                  End If
               Loop
            End If
            vTextFin = vTextFin & Chr(10)
            JustificaTexto = JustificaTexto & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin))
            nAncho1 = lnColPage
            vTextFin = ""
            Letra = ""
            K = 0
         Else
            vTextFin = vTextFin & Letra
         End If
      Else
        If i < N Then
            If Asc(Mid(sTemp, i + 1, 1)) = 13 Or Asc(Mid(sTemp, i + 1, 1)) = 10 Then
                i = i + 1
            End If
        End If
         JustificaTexto = JustificaTexto & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin)) & Chr(10)
         nAncho1 = lnColPage
         vTextFin = ""
         Letra = ""
         K = 0
      End If
   End If
Loop
JustificaTexto = JustificaTexto & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin))
End Function

Public Function JustificaTextoCadena(sTemp As String, lnColPage As Integer, Optional lsEspIzq As Integer = 0) As String
Dim vTextFin As String
Dim Letra As String * 1, i As Integer, K As Integer, N As Integer
Dim nVeces As Long, M As Integer, Fin As Integer, Ini As Integer
Dim nAncho1 As Integer, nSpa As Integer
i = 0
K = 0
N = Len(sTemp)
nAncho1 = lnColPage
Do While i <= N
   K = K + 1
   i = i + 1
   If i > N Then
      Exit Do
   End If
   Letra = Mid(sTemp, i, 1)
   If Letra = Chr$(27) Then
      vTextFin = vTextFin & Letra & Mid(sTemp, i + 1, 1)
      i = i + 1
      K = K + 1
      nAncho1 = nAncho1 + 2
   Else
      If Asc(Letra) <> 13 And Asc(Letra) <> 10 Then
         If K > nAncho1 Then
            M = 0
            If Mid(sTemp, i, 1) = Chr(32) Then
               vTextFin = Trim(vTextFin)
            Else
               M = InStrRev(vTextFin, " ", , vbTextCompare)
               If M = 0 Then M = 1
               If InStr(Mid(vTextFin, M, Len(vTextFin)), Chr$(27)) Then
                  nAncho1 = nAncho1 - 2
               End If
               i = i - (nAncho1 + 1 - M)
               vTextFin = Mid(vTextFin, 1, M - 1)
            End If
            nSpa = nAncho1 - Len(Trim(vTextFin))
            vTextFin = Trim(vTextFin)
            If nSpa <> 0 Then
               Fin = 1
               nVeces = 0
               M = 1
               Do While M <= nSpa
                  Ini = InStr(Fin, vTextFin, " ", vbTextCompare)
                  If Ini = 0 Then
                     Fin = 1
                     nVeces = nVeces + 1
                     M = M + 1
                  Else
                      vTextFin = Mid(vTextFin, 1, Ini) & " " & RTrim(Mid(vTextFin, Ini + 1, nAncho1))
                      Fin = Ini + 2 + nVeces
                      M = M + 1
                  End If
               Loop
            End If
            vTextFin = vTextFin & Chr(10)
            JustificaTextoCadena = JustificaTextoCadena & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin))
            nAncho1 = lnColPage
            vTextFin = ""
            Letra = ""
            K = 0
         Else
            vTextFin = vTextFin & Letra
         End If
      Else
        If i < N Then
            If Asc(Mid(sTemp, i + 1, 1)) = 13 Or Asc(Mid(sTemp, i + 1, 1)) = 10 Then
                i = i + 1
            End If
        End If
         JustificaTextoCadena = JustificaTextoCadena & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin)) & Chr(10)
         nAncho1 = lnColPage
         vTextFin = ""
         Letra = ""
         K = 0
      End If
   End If
Loop
JustificaTextoCadena = JustificaTextoCadena & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin))
End Function


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
         '***Modificado por ELRO el 20111102, según Acta 277-2011/TI-D
         'sPie = sPie + Centra(" LE _________________ ", nLenPie)
         sPie = sPie + Centra(" DNI _________________ ", nLenPie)
         '***Fin Modificado por ELRO
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
lsPiePag = lsPiePag + "" + Chr$(10) + Chr$(10) + Chr$(10) + Chr$(10)
lsPiePag = lsPiePag + sPiR + Chr$(10)
lsPiePag = lsPiePag + sPie + Chr$(10)
ImprePiePag = lsPiePag
End Function
Public Function Cabecera(sTit As String, P As Integer, psSimbolo As String, pnColPage As Integer, Optional sCabe As String = "", Optional sFecha As String = "") As String
Dim BON As String
Dim BOFF As String
Dim Con As String
Dim COFF As String

BON = PrnSet("B+")
BOFF = PrnSet("B-")
Con = PrnSet("C+")
COFF = PrnSet("C-")

Cabecera = ""
If sFecha = "" Then
   sFecha = Date
End If
   If P > 0 Then Cabecera = Chr$(12)
   P = P + 1
   Cabecera = Cabecera + " CMAC - TRUJILLO " & Space(42) & sFecha & " - " & Format(Time, "hh:mm:ss") & Chr$(10)
   Cabecera = Cabecera + Space(72) & "Pag. " & Format(P, "000") & Chr$(10)
   Cabecera = Cabecera + BON & Centra(sTit, pnColPage) & BOFF & Chr$(10)
   If psSimbolo <> "" Then
      Cabecera = Cabecera + BON & Centra(" M O N E D A   " & IIf(psSimbolo = "S/.", "N A C I O N A L ", "E X T R A N J E R A "), pnColPage) & BOFF & Chr$(10)
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


Public Function JustificaTextoSinCarEsp(sTemp As String, lnColPage As Integer, Optional lsEspIzq As Integer = 0) As String
Dim vTextFin As String
Dim Letra As String * 1, i As Integer, K As Integer, N As Integer
Dim nVeces As Long, M As Integer, Fin As Integer, Ini As Integer
Dim nAncho1 As Integer, nSpa As Integer
i = 0
K = 0
N = Len(sTemp)
nAncho1 = lnColPage
Do While i <= N
   K = K + 1
   i = i + 1
   If i > N Then
      Exit Do
   End If
   Letra = Mid(sTemp, i, 1)
   If Letra = Chr$(27) Then
      vTextFin = vTextFin & Letra & Mid(sTemp, i + 1, 1)
      i = i + 1
      K = K + 1
      nAncho1 = nAncho1 + 2
   Else
      If Asc(Letra) <> 13 And Asc(Letra) <> 10 Then
         If K > nAncho1 Then
            M = 0
            If Mid(sTemp, i, 1) = Chr(32) Then
               vTextFin = Trim(vTextFin)
            Else
               M = InStrRev(vTextFin, " ", , vbTextCompare)
               If M = 0 Then M = 1
               If InStr(Mid(vTextFin, M, Len(vTextFin)), Chr$(27)) Then
                  nAncho1 = nAncho1 - 2
               End If
               i = i - (nAncho1 + 1 - M)
               vTextFin = Mid(vTextFin, 1, M - 1)
            End If
            nSpa = nAncho1 - Len(Trim(vTextFin))
            vTextFin = Trim(vTextFin)
            If nSpa <> 0 Then
               Fin = 1
               nVeces = 0
               M = 1
               Do While M <= nSpa
                  Ini = InStr(Fin, vTextFin, " ", vbTextCompare)
                  If Ini = 0 Then
                     Fin = 1
                     nVeces = nVeces + 1
                     M = M + 1
                  Else
                      vTextFin = Mid(vTextFin, 1, Ini) & " " & RTrim(Mid(vTextFin, Ini + 1, nAncho1))
                      Fin = Ini + 2 + nVeces
                      M = M + 1
                  End If
               Loop
            End If
            vTextFin = vTextFin & Chr(10)
            JustificaTextoSinCarEsp = JustificaTextoSinCarEsp & Space(lsEspIzq) & Trim(vTextFin)
            nAncho1 = lnColPage
            vTextFin = ""
            Letra = ""
            K = 0
         Else
            vTextFin = vTextFin & Letra
         End If
      Else
        If i < N Then
            If Asc(Mid(sTemp, i + 1, 1)) = 13 Or Asc(Mid(sTemp, i + 1, 1)) = 10 Then
                i = i + 1
            End If
        End If
         JustificaTextoSinCarEsp = JustificaTextoSinCarEsp & Space(lsEspIzq) & Trim(vTextFin) & Chr(10)
         nAncho1 = lnColPage
         vTextFin = ""
         Letra = ""
         K = 0
      End If
   End If
Loop
JustificaTextoSinCarEsp = JustificaTextoSinCarEsp & Space(lsEspIzq) & Trim(vTextFin)
End Function

