Attribute VB_Name = "gGeneral"
Option Explicit
Global Const gnColPage = 79
'*******VARIABLES DE MONEDA****************
'MARG ERS044-2016
Global Const gcPEN_SINGULAR = "SOL"
Global Const gcPEN_PLURAL = "SOLES"
Global Const gcPEN_SIMBOLO = "S/"

Public oImpresora As New ContsImp.clsConstImp
Public gImpresora As Impresoras

Dim unidad(0 To 9) As String
Dim decena(0 To 9) As String
Dim centena(0 To 10) As String
Dim deci(0 To 9) As String
Dim otros(0 To 15) As String

'Public Function GeneraDocNro(psDocTpo As String, psMoneda As String, Optional psDocSerie As String = "") As String
'    On Error GoTo GeneraMovNroErr
'    Dim sSql As String
'    Dim rs As New ADODB.Recordset
'    Dim oConect As DConecta
'
'    Set oConect = New DConecta
'    If oConect.AbreConexion = False Then Exit Function
'
'    If psDocTpo <> "" Then
'        If psDocSerie <> "" Then
'            sSql = "SELECT max(cDocNro) AS cDocCorrela FROM Movdoc md JOIN Mov m ON m.cMovNro = md.cMovNro " _
'                & "WHERE   m.cMovFlag <> 'X' and cDocTpo = '" & psDocTpo & "' and substring(cDocNro,1," & Len(psDocSerie) & ") = '" & psDocSerie & "'"
'        Else
'            sSql = "SELECT  cDocNro AS cDocCorrela , md.cMovnro " _
'                & " FROM    movdoc md " _
'                & " WHERE   cDocTpo = '" & psDocTpo & "' " _
'                & "         and md.cmovnro = (  Select Max(MD1.cmovnro) " _
'                & "                             From MovDoc MD1 JOIN Mov M on M.cMovNro= MD1.cMovnro " _
'                & "                             WHERE MD1.cDocTpo = '" & psDocTpo & "' and m.cMovFlag NOT IN ('X') and Substring(M.cOpeCod,3,1) ='" & psMoneda & "') "
'       End If
'       Set rs = oConect.CargaRecordSet(sSql)
'       If Not IsNull(rs!cDocCorrela) Then
'          If psDocSerie <> "" Then
'             GeneraDocNro = psDocSerie & "-" & Format(Val(Mid(rs!cDocCorrela, Len(psDocSerie) + 2, 20)) + 1, String(8, "0"))
'          Else
'             If Mid(rs!cDocCorrela, 4, 1) = "-" Then
'                GeneraDocNro = Mid(rs!cDocCorrela, 1, 3) & "-" & Format(Val(Mid(rs!cDocCorrela, 5, 20)) + 1, String(8, "0"))
'             Else
'                GeneraDocNro = Format(Val(rs!cDocCorrela) + 1, String(8, "0"))
'             End If
'          End If
'       Else
'          GeneraDocNro = "00000001"
'          If psDocSerie <> "" Then
'             GeneraDocNro = psDocSerie & "-" & GeneraDocNro
'          End If
'       End If
'    End If
'    rs.Close
'    Set rs = Nothing
'    oConect.CierraConexion
'    Set oConect = Nothing
'    Exit Function
'GeneraMovNroErr:
'    MsgBox "Erro N°[" & Err.Number & "] " & Err.Description & vbCrLf & Err.Source, vbInformation, "Aviso"
'End Function

Public Function GetFecha(dtmFechas As Date) As String
Dim txtMeses As String
txtMeses = Choose(Month(dtmFechas), "Enero", "Febrero", "Marzo", "Abril", _
                                    "Mayo", "Junio", "Julio", "Agosto", _
                                    "Setiembre", "Octubre", "Noviembre", "Diciembre")
If dtmFechas <= Date Then
   GetFecha = Day(dtmFechas) & " de " & txtMeses & " de " & Year(dtmFechas)
Else
   GetFecha = Day(Date) & " de " & txtMeses & " de " & Year(Date)
End If
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
         JustificaTexto = JustificaTexto & Space(lsEspIzq) & RTrim(ImpreCarEsp(vTextFin)) & oImpresora.gPrnSaltoLinea
         nAncho1 = lnColPage
         vTextFin = ""
         Letra = ""
         K = 0
      End If
   End If
Loop
JustificaTexto = JustificaTexto & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin))
End Function
Public Function ConvNumLet(nNumero As Currency, lsMoneda As String, Optional lSoloText As Boolean = True, Optional lSinMoneda As Boolean = False) As String
Dim sCent As String
Dim xValor As Single
Dim vMoneda As String
Dim cNumero As String
Dim lsSimbolo As String
cNumero = Format(nNumero, "#,##0.00")
xValor = nNumero - Int(nNumero)
If xValor = 0 Then
   sCent = " Y 00/100 "
Else
   sCent = " Y " & Right(Trim(cNumero), 2) & "/100 "
End If
vMoneda = IIf(lsMoneda = "1", StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES AMERICANOS") 'marg ers044-2016
lsSimbolo = IIf(lsMoneda = "1", gcPEN_SIMBOLO, "$.") 'marg ers044-2016
If Not lSoloText Then
   ConvNumLet = Trim(lsSimbolo) & " " & Trim(Format(nNumero, "###,###,##0.00#")) & " ("
End If
ConvNumLet = ConvNumLet & Trim(UCase(NumLet(CStr(nNumero), 0))) & sCent & " " & IIf(lSinMoneda, "", Trim(vMoneda)) & IIf(lSoloText, "", ")")
End Function



'***************************************************

Public Function NumLet(ByVal strNum As String, Optional ByVal vLo)   '  , Optional ByVal vMoneda, Optional ByVal vCentimos) As String
    '----------------------------------------------------------
    ' Convierte el número strNum en letras          (28/Feb/91)
    '----------------------------------------------------------
    Dim I As Integer
    Dim Lo As Integer
    Dim iHayDecimal As Integer          'Posición del signo decimal
    Dim sDecimal As String              'Signo decimal a usar
    Dim sEntero As String
    Dim sFraccion As String
    Dim fFraccion As Single
    Dim sNumero As String
    '
    Dim sMoneda As String
    Dim sCentimos As String
    
    'Averiguar el signo decimal
    sNumero = Format$(25.5, "#.#")
    If InStr(sNumero, ".") Then
        sDecimal = "."
    Else
        sDecimal = ","
    End If
    'Si no se especifica el ancho...
    If IsMissing(vLo) Then
        Lo = 0
    Else
        Lo = vLo
    End If
    '
    If Lo Then
        sNumero = Space$(Lo)
    Else
        sNumero = ""
    End If
    'Quitar los espacios que haya por medio
    
    Do
        I = InStr(strNum, " ")
        If I = 0 Then Exit Do
        strNum = Left$(strNum, I - 1) & Mid$(strNum, I + 1)
    Loop
    
    'Comprobar si tiene decimales
    iHayDecimal = InStr(strNum, sDecimal)
    If iHayDecimal Then
        sEntero = Left$(strNum, iHayDecimal - 1)
        sFraccion = Mid$(strNum, iHayDecimal + 1) & "00"
        'obligar a que tenga dos cifras
        sFraccion = Left$(sFraccion, 2)
        fFraccion = Val(sFraccion)
        
        'Si no hay decimales... no agregar nada...
        If fFraccion < 1 Then
            strNum = RTrim$(UnNumero(sEntero) & sMoneda)
            If Lo Then
                LSet sNumero = strNum
            Else
                sNumero = strNum
            End If
            NumLet = sNumero
            Exit Function
        End If
        
        sEntero = UnNumero(sEntero)
        sFraccion = sFraccion & "/100"
        strNum = sEntero
        If Lo Then
            LSet sNumero = RTrim$(strNum)
        Else
            sNumero = RTrim$(strNum)
        End If
        NumLet = sNumero
    Else
        strNum = RTrim$(UnNumero(strNum) & sMoneda)
        If Lo Then
            LSet sNumero = strNum
        Else
            sNumero = strNum
        End If
        NumLet = sNumero
    End If
End Function


Public Function UnNumero(ByVal strNum As String) As String
    '----------------------------------------------------------
    'Esta es la rutina principal                    (10/Jul/97)
    'Está separada para poder actuar con decimales
    '----------------------------------------------------------
    
    Dim lngA As Double
    Dim Negativo As Boolean
    Dim l As Integer
    Dim Una As Boolean
    Dim Millon As Boolean
    Dim Millones As Boolean
    Dim vez As Integer
    Dim MaxVez As Integer
    Dim K As Integer
    Dim strQ As String
    Dim strB As String
    Dim strU As String
    Dim strD As String
    Dim strC As String
    Dim iA As Integer
    '
    Dim strN() As String
    
    'Si se amplia este valor... no se manipularán bien los números
    Const cAncho = 12
    Const cGrupos = cAncho \ 3
    '
    If unidad(1) <> "una" Then
        InicializarArrays
    End If
    'Si se produce un error que se pare el mundo!!!
    On Local Error GoTo 0
    
    lngA = Abs(CDbl(strNum))
    Negativo = (lngA <> CDbl(strNum))
    strNum = LTrim$(RTrim$(Str$(lngA)))
    l = Len(strNum)
    
    If lngA < 1 Then
        UnNumero = "cero"
        Exit Function
    End If
    '
    Una = True
    Millon = False
    Millones = False
    If l < 4 Then Una = False
    If lngA > 999999 Then Millon = True
    If lngA > 1999999 Then Millones = True
    strB = ""
    strQ = strNum
    vez = 0
    
    ReDim strN(1 To cGrupos)
    strQ = Right$(String$(cAncho, "0") & strNum, cAncho)
    For K = Len(strQ) To 1 Step -3
        vez = vez + 1
        strN(vez) = Mid$(strQ, K - 2, 3)
    Next
    MaxVez = cGrupos
    For K = cGrupos To 1 Step -1
        If strN(K) = "000" Then
            MaxVez = MaxVez - 1
        Else
            Exit For
        End If
    Next
    For vez = 1 To MaxVez
        strU = "": strD = "": strC = ""
        strNum = strN(vez)
        l = Len(strNum)
        K = Val(Right$(strNum, 2))
        If Right$(strNum, 1) = "0" Then
            K = K \ 10
            strD = decena(K)
        ElseIf K > 10 And K < 16 Then
            K = Val(Mid$(strNum, l - 1, 2))
            strD = otros(K)
        Else
            strU = unidad(Val(Right$(strNum, 1)))
            If l - 1 > 0 Then
                K = Val(Mid$(strNum, l - 1, 1))
                strD = deci(K)
            End If
        End If
        
        If l - 2 > 0 Then
            K = Val(Mid$(strNum, l - 2, 1))
            'Con esto funcionará bien el 100100, por ejemplo...
            If K = 1 Then
                If Val(strNum) = 100 Then
                    K = 10
                End If
            End If
            strC = centena(K) & " "
        End If
        '------
        If strU = "uno" And Left$(strB, 4) = " mil" Then strU = ""
        strB = strC & strD & strU & " " & strB
    
        If (vez = 1 Or vez = 3) Then
            If strN(vez + 1) <> "000" Then strB = " mil " & strB
        End If
        If vez = 2 And Millon Then
            If Millones Then
                strB = " millones " & strB
            Else
                strB = "un millón " & strB
            End If
        End If
    Next
    strB = Trim$(strB)
    If Right$(strB, 3) = "uno" Then strB = Left$(strB, Len(strB) - 1) & "a"
    Do                              'Quitar los espacios que haya por medio
        iA = InStr(strB, "  ")
        If iA = 0 Then Exit Do
        strB = Left$(strB, iA - 1) & Mid$(strB, iA + 1)
    Loop
    If Left$(strB, 6) = "un  un" Then strB = Mid$(strB, 5)
    If Left$(strB, 5) = "un un" Then strB = Mid$(strB, 4)
    If Left$(strB, 6) = "un mil" Then strB = Mid$(strB, 4)
    If Left$(strB, 7) = "un  mil" Then strB = Mid$(strB, 5)
    If Right$(strB, 16) <> "millones mil un " Then
        iA = InStr(strB, "millones mil un ")
        If iA Then strB = Left$(strB, iA + 8) & Mid$(strB, iA + 13)
    End If
    If Right$(strB, 6) = "ciento" Then strB = Left$(strB, Len(strB) - 2)
    If Negativo Then strB = "menos " & strB
    
    UnNumero = Trim$(strB)
End Function

'***************************************************
'* Funcion:  Es llamada desde UnNumero
'***************************************************
'FECHA CREACION : 24/06/99  -   FAOS
'MODIFICACION:
'***************************************************
Public Sub InicializarArrays()
    'Asignar los valores
    unidad(1) = "un"
    unidad(2) = "dos"
    unidad(3) = "tres"
    unidad(4) = "cuatro"
    unidad(5) = "cinco"
    unidad(6) = "seis"
    unidad(7) = "siete"
    unidad(8) = "ocho"
    unidad(9) = "nueve"
    '
    decena(1) = "diez"
    decena(2) = "veinte"
    decena(3) = "treinta"
    decena(4) = "cuarenta"
    decena(5) = "cincuenta"
    decena(6) = "sesenta"
    decena(7) = "setenta"
    decena(8) = "ochenta"
    decena(9) = "noventa"
    '
    centena(1) = "ciento"
    centena(2) = "doscientos"
    centena(3) = "trescientos"
    centena(4) = "cuatrocientos"
    centena(5) = "quinientos"
    centena(6) = "seiscientos"
    centena(7) = "setecientos"
    centena(8) = "ochocientos"
    centena(9) = "novecientos"
    centena(10) = "cien"                'Parche
    '
    deci(1) = "dieci"
    deci(2) = "veinti"
    deci(3) = "treinta y "
    deci(4) = "cuarenta y "
    deci(5) = "cincuenta y "
    deci(6) = "sesenta y "
    deci(7) = "setenta y "
    deci(8) = "ochenta y "
    deci(9) = "noventa y "
    '
    otros(1) = "1"
    otros(2) = "2"
    otros(3) = "3"
    otros(4) = "4"
    otros(5) = "5"
    otros(6) = "6"
    otros(7) = "7"
    otros(8) = "8"
    otros(9) = "9"
    otros(10) = "10"
    otros(11) = "once"
    otros(12) = "doce"
    otros(13) = "trece"
    otros(14) = "catorce"
    otros(15) = "quince"
End Sub

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
Public Function NumerosEnteros(intTecla As Integer, Optional pbNegativos As Boolean = False) As Integer
Dim cValidar As String
    If pbNegativos = False Then
        cValidar = "0123456789"
    Else
        cValidar = "0123456789-"
    End If
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) = 0 Then
            intTecla = 0
            Beep
        End If
    End If
    NumerosEnteros = intTecla
End Function
Public Function FormaOrdenPago(psPersona As String, pdFecha As Date, pnImporte As Currency, ByVal psMoneda As String) As String
Dim sTexto As String, N As Integer
Dim sDatos As String

Dim BON As String
Dim BOFF As String
Dim COFF As String
Dim CON As String

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

sDatos = ""
sDatos = sDatos & Space(40) & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & Space(3) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
sDatos = sDatos & PrnSet("Esp", 18) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
sDatos = sDatos & Space(11) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(59, "*"), 1, IIf(Len(psPersona) > 59, 89, 59)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
sTexto = ConvNumLet(pnImporte, psMoneda, True)
sDatos = sDatos & Space(5) & IIf(Len(sTexto) > 57, CON, "") & Mid(Trim(sTexto) & String(57, "*"), 1, IIf(Len(sTexto) > 57, 85, 57)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
FormaOrdenPago = ImpreCarEsp(sDatos)
End Function
Public Function FormaCheque(sBanco As String, psEmpresaRuc As String, psPersona As String, pdFecha As Date, pnImporte As Currency, ByVal psMoneda As String) As String
Dim sTexto As String, N As Integer
Dim sDatos As String
Dim BON As String
Dim BOFF As String
Dim CON As String
Dim COFF As String
Dim Caja As Integer
Dim psCodCmac As String
Dim sql As String

Caja = 0
BON = oImpresora.gPrnBoldON
BOFF = oImpresora.gPrnBoldOFF
CON = oImpresora.gPrnCondensadaON
COFF = oImpresora.gPrnCondensadaOFF
'sDatos = oImpresora.gPrnEspaLineaCab
sDatos = oImpresora.gPrnEspaLineaN

Dim oConst As New NConstSistemas
psCodCmac = oConst.LeeConstSistema(gConstSistCodCMAC)
Set oConst = Nothing

If psCodCmac = "112" Then    'TRUJILLO
    Select Case sBanco
       Case "01"      'Banco Central de Reserva 'NO TIENE
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "02"      'Banco de la Nación
          sDatos = sDatos & PrnSet("Esp", 6) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(30) & BON & "TRUJILLO, " & Mid(pdFecha, 1, 2) & "   " & Mid(pdFecha, 4, 2) & "   " & Right(pdFecha, 4) & Space(6) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(7) & BON & IIf(Len(Trim(psPersona)) > 73, CON, "") & Mid(Trim(psPersona) & String(73, "*"), 1, IIf(Len(psPersona) > 73, 108, 73)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(1) & IIf(Len(sTexto) > 70, CON, "") & Mid(Trim(sTexto) & String(70, "*"), 1, IIf(Len(sTexto) > 70, 105, 70)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(20) & BON & " RUC :" & Trim(psEmpresaRuc) & BOFF & oImpresora.gPrnSaltoLinea
       Case "03"      'Banco de Credito
          sDatos = sDatos & PrnSet("Esp", 6) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(31) & CON & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & COFF & Space(11) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 4) & oImpresora.gPrnSaltoLinea & CON & BON & "  RUC :" & Trim(psEmpresaRuc) & COFF & BOFF
          sDatos = sDatos & PrnSet("Esp", 22) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          'sDatos = sDatos & PrnSet("EspN")
          sDatos = sDatos & Space(13) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(57, "*"), 1, IIf(Len(psPersona) > 57, 86, 57)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(3) & IIf(Len(sTexto) > 57, CON, "") & Mid(Trim(sTexto) & String(57, "*"), 1, IIf(Len(sTexto) > 57, 85, 57)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "04"      'BANCO INTERNACIONAL
          sDatos = sDatos & Space(33) & CON & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & COFF & Space(8) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(6) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(59, "*"), 1, IIf(Len(psPersona) > 59, 89, 59)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(2) & IIf(Len(sTexto) > 54, CON, "") & Mid(Trim(sTexto) & String(54, "*"), 1, IIf(Len(sTexto) > 54, 79, 54)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(15) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea
    
       Case "06"      'BANCO CONTINENTAL  1/2 OK
          sDatos = sDatos & PrnSet("Esp", 7) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(36) & BON & Mid(pdFecha, 1, 2) & " " & Mid(pdFecha, 4, 2) & " " & Right(pdFecha, 4) & Space(5) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 5) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(9) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(59, "*"), 1, IIf(Len(psPersona) > 59, 89, 59)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(3) & IIf(Len(sTexto) > 57, CON, "") & Mid(Trim(sTexto) & String(57, "*"), 1, IIf(Len(sTexto) > 57, 85, 57)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(8) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea
       'Case "05"      'BANCO DE LIMA - SUDAMERIS
       '   sDatos = sDatos & Space(37) & CON & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & COFF & Space(9) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       '   sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
       '   sDatos = sDatos & Space(8) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(59, "*"), 1, IIf(Len(psPersona) > 59, 89, 59)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       '   sTexto = ConvNumLet(pnImporte, psMoneda, True)
       '   sDatos = sDatos & Space(8) & IIf(Len(sTexto) > 54, CON, "") & Mid(Trim(sTexto) & String(54, "*"), 1, IIf(Len(sTexto) > 54, 79, 54)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       '   sDatos = sDatos & Space(15) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea
          
       Case "21"      'BID BANCO INTERAMERICANO DE FINANZAS
          sDatos = sDatos & PrnSet("Esp", 6) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(28) & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & Space(8) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(8) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(59, "*"), 1, IIf(Len(psPersona) > 59, 89, 59)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(4) & IIf(Len(sTexto) > 57, CON, "") & Mid(Trim(sTexto) & String(57, "*"), 1, IIf(Len(sTexto) > 57, 85, 57)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(15) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea
          
       Case "24"      'BANCO DEL TRABAJO
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "25"      'BANCO SOLVENTA
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "26"      'BANCO SERBANCO
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "11"      'BANCO LATINO  OK
          sDatos = sDatos & Space(27) & BON & "TRUJILLO, " & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & " " & Right(pdFecha, 4) & Space(3) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 4) & oImpresora.gPrnSaltoLinea & BON & "  RUC :" & Trim(psEmpresaRuc) & PrnSet("EspN") & BOFF
          sDatos = sDatos & PrnSet("Esp", 18) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(8) & BON & IIf(Len(Trim(psPersona)) > 57, CON, "") & Mid(Trim(psPersona) & String(57, "*"), 1, IIf(Len(psPersona) > 57, 87, 57)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(2) & IIf(Len(sTexto) > 55, CON, "") & Mid(Trim(sTexto) & String(55, "*"), 1, IIf(Len(sTexto) > 55, 83, 55)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "18"      'BANCO NUEVO MUNDO OK
          sDatos = sDatos & PrnSet("Esp", 8) & oImpresora.gPrnSaltoLinea & PrnSet("EspN") & Space(36) & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & Space(4) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & CON & BON & "  RUC :" & Trim(psEmpresaRuc) & COFF & BOFF
          sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(9) & BON & IIf(Len(Trim(psPersona)) > 57, CON, "") & Mid(Trim(psPersona) & String(57, "*"), 1, IIf(Len(psPersona) > 57, 87, 57)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(3) & IIf(Len(sTexto) > 54, CON, "") & Mid(Trim(sTexto) & String(54, "*"), 1, IIf(Len(sTexto) > 54, 84, 54)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "19"      'BANCO STANDARD CHARTERED
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "20"
            sDatos = sDatos & Space(26) & "TRUJILLO " & Mid(pdFecha, 1, 2) & " " & Mid(pdFecha, 4, 2) & " " & Right(pdFecha, 2) & Space(5) & Format(pnImporte, "###,###,##0.00") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            sDatos = sDatos & PrnSet("Esp", 18) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
            sDatos = sDatos & Space(5) & BON & IIf(Len(Trim(psPersona)) > 57, CON, "") & Mid(Trim(psPersona) & String(57, "*"), 1, IIf(Len(psPersona) > 57, 87, 57)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            sTexto = ConvNumLet(pnImporte, psMoneda, True)
            sDatos = sDatos & Space(5) & IIf(Len(sTexto) > 55, CON, "") & Mid(Trim(sTexto) & String(55, "*"), 1, IIf(Len(sTexto) > 55, 83, 55)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        Case "05"   'BANCO WIESSE SUDAMERIS
            sDatos = sDatos & Space(28) & CON & BON & Mid(pdFecha, 1, 2) & "   " & Mid(pdFecha, 4, 2) & "    " & Right(pdFecha, 4) & COFF & Space(12) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
            sDatos = sDatos & Space(3) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(59, "*"), 1, IIf(Len(psPersona) > 59, 89, 59)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            sTexto = ConvNumLet(pnImporte, psMoneda, True)
            sDatos = sDatos & Space(3) & IIf(Len(sTexto) > 54, CON, "") & Mid(Trim(sTexto) & String(54, "*"), 1, IIf(Len(sTexto) > 54, 79, 54)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            sDatos = sDatos & Space(10) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    End Select
End If

If psCodCmac = "102" Then     '********************** L I M A ********************
    Select Case sBanco
       Case "01"      'Banco Central de Reserva 'NO TIENE
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "02"      'Banco de la Nación
          sDatos = sDatos & PrnSet("Esp", 6) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(30) & BON & "LIMA, " & Mid(pdFecha, 1, 2) & "   " & Mid(pdFecha, 4, 2) & "   " & Right(pdFecha, 4) & Space(6) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(7) & BON & IIf(Len(Trim(psPersona)) > 73, CON, "") & Mid(Trim(psPersona) & String(73, "*"), 1, IIf(Len(psPersona) > 73, 108, 73)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(1) & IIf(Len(sTexto) > 70, CON, "") & Mid(Trim(sTexto) & String(70, "*"), 1, IIf(Len(sTexto) > 70, 105, 70)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(20) & BON & " RUC :" & Trim(psEmpresaRuc) & BOFF & oImpresora.gPrnSaltoLinea
       Case "03"      'Banco de Credito
          sDatos = sDatos & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(38) & CON & BON & Mid(pdFecha, 1, 2) & "   " & Mid(pdFecha, 4, 2) & "   " & Right(pdFecha, 4) & COFF & Space(5) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 22) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(10) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(50, "*"), 1, IIf(Len(psPersona) > 50, 79, 50)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True, True)
          sDatos = sDatos & Space(3) & IIf(Len(sTexto) > 57, CON, "") & Mid(Trim(sTexto) & String(50, "*"), 1, IIf(Len(sTexto) > 50, 78, 50)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "04"      'BANCO INTERNACIONAL
          sDatos = sDatos & Space(33) & CON & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & COFF & Space(8) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(6) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(59, "*"), 1, IIf(Len(psPersona) > 59, 89, 59)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(2) & IIf(Len(sTexto) > 54, CON, "") & Mid(Trim(sTexto) & String(54, "*"), 1, IIf(Len(sTexto) > 54, 79, 54)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(15) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea
    
       Case "05"   'BANCO WIESSE SUDAMERIS
            If psMoneda = 1 Then        '****** SOLES
                sDatos = sDatos & Space(33) & CON & BON & Mid(pdFecha, 1, 2) & "   " & Mid(pdFecha, 4, 2) & "    " & Right(pdFecha, 4) & COFF & Space(10) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
                sDatos = sDatos & Space(8) & BON & IIf(Len(Trim(psPersona)) > 52, CON, "") & Mid(Trim(psPersona) & String(52, "*"), 1, IIf(Len(psPersona) > 52, 82, 52)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                sTexto = ConvNumLet(pnImporte, psMoneda, True, True)
                sDatos = sDatos & Space(8) & IIf(Len(sTexto) > 54, CON, "") & Mid(Trim(sTexto) & String(48, "*"), 1, IIf(Len(sTexto) > 48, 73, 48)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                sDatos = sDatos & Space(6) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            ElseIf psMoneda = 2 Then    '****** DOLARES
                sDatos = sDatos & Space(35) & CON & BON & Mid(pdFecha, 1, 2) & "   " & Mid(pdFecha, 4, 2) & "    " & Right(pdFecha, 4) & COFF & Space(10) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
                sDatos = sDatos & Space(8) & BON & IIf(Len(Trim(psPersona)) > 52, CON, "") & Mid(Trim(psPersona) & String(52, "*"), 1, IIf(Len(psPersona) > 52, 82, 52)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                sTexto = ConvNumLet(pnImporte, psMoneda, True, True)
                sDatos = sDatos & Space(8) & IIf(Len(sTexto) > 54, CON, "") & Mid(Trim(sTexto) & String(48, "*"), 1, IIf(Len(sTexto) > 48, 73, 48)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                sDatos = sDatos & Space(6) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            End If
       Case "06"      'BANCO CONTINENTAL  1/2 OK
          sDatos = sDatos & PrnSet("Esp", 7) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(36) & BON & Mid(pdFecha, 1, 2) & " " & Mid(pdFecha, 4, 2) & " " & Right(pdFecha, 4) & Space(5) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 5) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(9) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(59, "*"), 1, IIf(Len(psPersona) > 59, 89, 59)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(3) & IIf(Len(sTexto) > 57, CON, "") & Mid(Trim(sTexto) & String(57, "*"), 1, IIf(Len(sTexto) > 57, 85, 57)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(8) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea
       Case "21"      'BID BANCO INTERAMERICANO DE FINANZAS
          sDatos = sDatos & PrnSet("Esp", 6) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(28) & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & Space(8) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(8) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(59, "*"), 1, IIf(Len(psPersona) > 59, 89, 59)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(4) & IIf(Len(sTexto) > 57, CON, "") & Mid(Trim(sTexto) & String(57, "*"), 1, IIf(Len(sTexto) > 57, 85, 57)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(15) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea
       Case "24"      'BANCO DEL TRABAJO
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "25"      'BANCO SOLVENTA
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "26"      'BANCO SERBANCO
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "11"      'BANCO LATINO  OK
          sDatos = sDatos & Space(27) & BON & "LIMA, " & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & " " & Right(pdFecha, 4) & Space(3) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 4) & oImpresora.gPrnSaltoLinea & BON & "  RUC :" & Trim(psEmpresaRuc) & PrnSet("EspN") & BOFF
          sDatos = sDatos & PrnSet("Esp", 18) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(8) & BON & IIf(Len(Trim(psPersona)) > 57, CON, "") & Mid(Trim(psPersona) & String(57, "*"), 1, IIf(Len(psPersona) > 57, 87, 57)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(2) & IIf(Len(sTexto) > 55, CON, "") & Mid(Trim(sTexto) & String(55, "*"), 1, IIf(Len(sTexto) > 55, 83, 55)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "18"      'BANCO NUEVO MUNDO OK
          sDatos = sDatos & PrnSet("Esp", 8) & oImpresora.gPrnSaltoLinea & PrnSet("EspN") & Space(36) & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & Space(4) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & CON & BON & "  RUC :" & Trim(psEmpresaRuc) & COFF & BOFF
          sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(9) & BON & IIf(Len(Trim(psPersona)) > 57, CON, "") & Mid(Trim(psPersona) & String(57, "*"), 1, IIf(Len(psPersona) > 57, 87, 57)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(3) & IIf(Len(sTexto) > 54, CON, "") & Mid(Trim(sTexto) & String(54, "*"), 1, IIf(Len(sTexto) > 54, 84, 54)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "19"      'BANCO STANDARD CHARTERED
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "20"
            sDatos = sDatos & Space(26) & "LIMA " & Mid(pdFecha, 1, 2) & " " & Mid(pdFecha, 4, 2) & " " & Right(pdFecha, 2) & Space(5) & Format(pnImporte, "###,###,##0.00") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            sDatos = sDatos & PrnSet("Esp", 18) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
            sDatos = sDatos & Space(5) & BON & IIf(Len(Trim(psPersona)) > 57, CON, "") & Mid(Trim(psPersona) & String(57, "*"), 1, IIf(Len(psPersona) > 57, 87, 57)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            sTexto = ConvNumLet(pnImporte, psMoneda, True)
            sDatos = sDatos & Space(5) & IIf(Len(sTexto) > 55, CON, "") & Mid(Trim(sTexto) & String(55, "*"), 1, IIf(Len(sTexto) > 55, 83, 55)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    End Select

End If

If psCodCmac = "231" Then     '********************** I C A ********************
    Select Case sBanco
       Case "01"      'Banco Central de Reserva 'NO TIENE
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "02"      'Banco de la Nación
          sDatos = sDatos & PrnSet("Esp", 6) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(30) & BON & "EDPYME CONFIANZA, " & Mid(pdFecha, 1, 2) & "   " & Mid(pdFecha, 4, 2) & "   " & Right(pdFecha, 4) & Space(6) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(7) & BON & IIf(Len(Trim(psPersona)) > 73, CON, "") & Mid(Trim(psPersona) & String(73, "*"), 1, IIf(Len(psPersona) > 73, 108, 73)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(1) & IIf(Len(sTexto) > 70, CON, "") & Mid(Trim(sTexto) & String(70, "*"), 1, IIf(Len(sTexto) > 70, 105, 70)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(20) & BON & " RUC :" & Trim(psEmpresaRuc) & BOFF & oImpresora.gPrnSaltoLinea
       Case "03"      'Banco de Credito
          sDatos = sDatos & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(38) & CON & BON & Mid(pdFecha, 1, 2) & "   " & Mid(pdFecha, 4, 2) & "   " & Right(pdFecha, 4) & COFF & Space(5) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 22) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(10) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(50, "*"), 1, IIf(Len(psPersona) > 50, 79, 50)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True, True)
          sDatos = sDatos & Space(3) & IIf(Len(sTexto) > 57, CON, "") & Mid(Trim(sTexto) & String(50, "*"), 1, IIf(Len(sTexto) > 50, 78, 50)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "04"      'BANCO INTERNACIONAL
          sDatos = sDatos & Space(33) & CON & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & COFF & Space(8) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(6) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(59, "*"), 1, IIf(Len(psPersona) > 59, 89, 59)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(2) & IIf(Len(sTexto) > 54, CON, "") & Mid(Trim(sTexto) & String(54, "*"), 1, IIf(Len(sTexto) > 54, 79, 54)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(15) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea
    
       Case "05"   'BANCO WIESSE SUDAMERIS
            If psMoneda = 1 Then        '****** SOLES
                sDatos = sDatos & Space(33) & CON & BON & Mid(pdFecha, 1, 2) & "   " & Mid(pdFecha, 4, 2) & "    " & Right(pdFecha, 4) & COFF & Space(10) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
                sDatos = sDatos & Space(8) & BON & IIf(Len(Trim(psPersona)) > 52, CON, "") & Mid(Trim(psPersona) & String(52, "*"), 1, IIf(Len(psPersona) > 52, 82, 52)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                sTexto = ConvNumLet(pnImporte, psMoneda, True, True)
                sDatos = sDatos & Space(8) & IIf(Len(sTexto) > 54, CON, "") & Mid(Trim(sTexto) & String(48, "*"), 1, IIf(Len(sTexto) > 48, 73, 48)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                sDatos = sDatos & Space(6) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            ElseIf psMoneda = 2 Then    '****** DOLARES
                sDatos = sDatos & Space(35) & CON & BON & Mid(pdFecha, 1, 2) & "   " & Mid(pdFecha, 4, 2) & "    " & Right(pdFecha, 4) & COFF & Space(10) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
                sDatos = sDatos & Space(8) & BON & IIf(Len(Trim(psPersona)) > 52, CON, "") & Mid(Trim(psPersona) & String(52, "*"), 1, IIf(Len(psPersona) > 52, 82, 52)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                sTexto = ConvNumLet(pnImporte, psMoneda, True, True)
                sDatos = sDatos & Space(8) & IIf(Len(sTexto) > 54, CON, "") & Mid(Trim(sTexto) & String(48, "*"), 1, IIf(Len(sTexto) > 48, 73, 48)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                sDatos = sDatos & Space(6) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            End If
       Case "06"      'BANCO CONTINENTAL  1/2 OK
          sDatos = sDatos & PrnSet("Esp", 7) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(36) & BON & Mid(pdFecha, 1, 2) & " " & Mid(pdFecha, 4, 2) & " " & Right(pdFecha, 4) & Space(5) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 5) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(9) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(59, "*"), 1, IIf(Len(psPersona) > 59, 89, 59)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(3) & IIf(Len(sTexto) > 57, CON, "") & Mid(Trim(sTexto) & String(57, "*"), 1, IIf(Len(sTexto) > 57, 85, 57)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(8) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea
       Case "21"      'BID BANCO INTERAMERICANO DE FINANZAS
          sDatos = sDatos & PrnSet("Esp", 6) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(28) & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & Space(8) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(8) & BON & IIf(Len(Trim(psPersona)) > 59, CON, "") & Mid(Trim(psPersona) & String(59, "*"), 1, IIf(Len(psPersona) > 59, 89, 59)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(4) & IIf(Len(sTexto) > 57, CON, "") & Mid(Trim(sTexto) & String(57, "*"), 1, IIf(Len(sTexto) > 57, 85, 57)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & Space(15) & CON & BON & " RUC :" & Trim(psEmpresaRuc) & COFF & BOFF & oImpresora.gPrnSaltoLinea
       Case "24"      'BANCO DEL TRABAJO
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "25"      'BANCO SOLVENTA
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "26"      'BANCO SERBANCO
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "11"      'BANCO LATINO  OK
          sDatos = sDatos & Space(27) & BON & "ICA, " & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & " " & Right(pdFecha, 4) & Space(3) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & PrnSet("Esp", 4) & oImpresora.gPrnSaltoLinea & BON & "  RUC :" & Trim(psEmpresaRuc) & PrnSet("EspN") & BOFF
          sDatos = sDatos & PrnSet("Esp", 18) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(8) & BON & IIf(Len(Trim(psPersona)) > 57, CON, "") & Mid(Trim(psPersona) & String(57, "*"), 1, IIf(Len(psPersona) > 57, 87, 57)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(2) & IIf(Len(sTexto) > 55, CON, "") & Mid(Trim(sTexto) & String(55, "*"), 1, IIf(Len(sTexto) > 55, 83, 55)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "18"      'BANCO NUEVO MUNDO OK
          sDatos = sDatos & PrnSet("Esp", 8) & oImpresora.gPrnSaltoLinea & PrnSet("EspN") & Space(36) & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & Space(4) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sDatos = sDatos & CON & BON & "  RUC :" & Trim(psEmpresaRuc) & COFF & BOFF
          sDatos = sDatos & PrnSet("Esp", 24) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
          sDatos = sDatos & Space(9) & BON & IIf(Len(Trim(psPersona)) > 57, CON, "") & Mid(Trim(psPersona) & String(57, "*"), 1, IIf(Len(psPersona) > 57, 87, 57)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
          sTexto = ConvNumLet(pnImporte, psMoneda, True)
          sDatos = sDatos & Space(3) & IIf(Len(sTexto) > 54, CON, "") & Mid(Trim(sTexto) & String(54, "*"), 1, IIf(Len(sTexto) > 54, 84, 54)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "19"      'BANCO STANDARD CHARTERED
          sDatos = " * NO EXISTE FORMATO DE CHEQUE * " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       Case "20"
            sDatos = sDatos & Space(26) & "ICA " & Mid(pdFecha, 1, 2) & " " & Mid(pdFecha, 4, 2) & " " & Right(pdFecha, 2) & Space(5) & Format(pnImporte, "###,###,##0.00") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            sDatos = sDatos & PrnSet("Esp", 18) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
            sDatos = sDatos & Space(5) & BON & IIf(Len(Trim(psPersona)) > 57, CON, "") & Mid(Trim(psPersona) & String(57, "*"), 1, IIf(Len(psPersona) > 57, 87, 57)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            sTexto = ConvNumLet(pnImporte, psMoneda, True)
            sDatos = sDatos & Space(5) & IIf(Len(sTexto) > 55, CON, "") & Mid(Trim(sTexto) & String(55, "*"), 1, IIf(Len(sTexto) > 55, 83, 55)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    End Select

End If

FormaCheque = ImpreCarEsp(sDatos)
End Function
Public Function PrnSet(Code As String, Optional nValor As Integer) As String
If nValor = 12 Or nValor = 10 Then
   nValor = nValor - 1
End If
Select Case Code
 Case "B+": PrnSet = oImpresora.gPrnBoldON 'Bold On
 Case "B-": PrnSet = oImpresora.gPrnBoldOFF 'Bold Off
 Case "U+": PrnSet = oImpresora.gPrnUnderLineONOFF  'Underline On
 Case "U-": PrnSet = oImpresora.gPrnUnderLineONOFF 'Chr$(27) & Chr$(46) 'Underline Off
 Case "I+": PrnSet = oImpresora.gPrnItalicON 'Italic On
 Case "I-": PrnSet = oImpresora.gPrnItalicOFF 'Italic Off
 Case "W+": PrnSet = oImpresora.gPrnDblAnchoON 'Doble Ancho On
 Case "W-": PrnSet = oImpresora.gPrnDblAnchoOFF 'Doble Ancho Off
 Case "C+": PrnSet = oImpresora.gPrnCondensadaON 'Condensado On
 Case "C-": PrnSet = oImpresora.gPrnCondensadaOFF 'Condensado Off
 Case "Rm": PrnSet = oImpresora.gPrnTpoLetraRoman 'Roman
 Case "Ss": PrnSet = oImpresora.gPrnTpoLetraSansSerif 'Sans Serif
 Case "Co": PrnSet = oImpresora.gPrnTpoLetraCurier 'Courier
 Case "1.5": PrnSet = oImpresora.gPrnUnoMedioEspacio ' 1 1/2 espacios
 Case "MI": PrnSet = oImpresora.gPrnMargenIzqCab 'Margen Izquierdo
 Case "MD": PrnSet = oImpresora.gPrnMargenDerCab & Chr$(nValor)   'Margen Derecho
 Case "10CPI": PrnSet = oImpresora.gPrnTamLetra12CPI
 Case "12CPI": PrnSet = oImpresora.gPrnTamLetra10CPI
 Case "15CPI": PrnSet = oImpresora.gPrnTamLetra15CPI
 Case "EspN": PrnSet = oImpresora.gPrnEspaLineaN     'Espaciado Normal 4.5/72
 Case "Esp":  PrnSet = oImpresora.gPrnEspaLineaValor & Chr$(nValor)  'Espaciado nValor/72 pulg.
End Select
End Function

Public Function ValFecha(lsControl As Control) As Boolean
   If Mid(lsControl, 1, 2) > 0 And Mid(lsControl, 1, 2) <= 31 Then
        If Mid(lsControl, 4, 2) > 0 And Mid(lsControl, 4, 2) <= 12 Then
            If Mid(lsControl, 7, 4) >= 1900 And Mid(lsControl, 7, 4) <= 9999 Then
               If IsDate(lsControl) = False Then
                    ValFecha = False
                    MsgBox "Formato de fecha no es válido", vbInformation, "Aviso"
                    lsControl.SetFocus
                    Exit Function
               Else
                    ValFecha = True
               End If
            Else
                ValFecha = False
                MsgBox "Año de Fecha no es válido", vbInformation, "Aviso"
                lsControl.SetFocus
                lsControl.SelStart = 6
                lsControl.SelLength = 4
                Exit Function
            End If
        Else
            ValFecha = False
            MsgBox "Mes de Fecha no es válido", vbInformation, "Aviso"
            lsControl.SetFocus
            lsControl.SelStart = 3
            lsControl.SelLength = 2
            Exit Function
        End If
    Else
        ValFecha = False
        MsgBox "Dia de Fecha no es válido", vbInformation, "Aviso"
        lsControl.SetFocus
        lsControl.SelStart = 0
        lsControl.SelLength = 2
        Exit Function
    End If
End Function

Public Function FormaOrdenPagoTalon(ByVal psPersNombre As String, ByVal pdFecha As Date, ByVal pnImporte As Currency, ByVal psMoneda As String) As String
Dim sTexto As String, N As Integer
Dim sDatos As String
Dim BON As String
Dim BOFF As String
Dim COFF As String
Dim CON As String

BON = PrnSet("B+")
BOFF = PrnSet("B-")
CON = PrnSet("C+")
COFF = PrnSet("C-")

'Formato anterior
'sDatos = PrnSet("Esp", 19) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
'sDatos = sDatos & Space(61) & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & Space(3) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'sDatos = sDatos & PrnSet("Esp", 15) & oImpresora.gPrnSaltoLinea
'sDatos = sDatos & Space(5) & BON & Format(pnImporte, "###,###,##0.00") & BOFF
'sDatos = sDatos & PrnSet("Esp", 4) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
'sDatos = sDatos & Space(32) & BON & IIf(Len(Trim(psPersNombre)) > 59, CON, "") & Mid(Trim(psPersNombre) & String(59, "*"), 1, IIf(Len(psPersNombre) > 59, 89, 59)) & COFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'sTexto = ConvNumLet(pnImporte, psMoneda, True)
'sDatos = sDatos & Space(26) & IIf(Len(sTexto) > 57, CON, "") & Mid(Trim(sTexto) & String(57, "*"), 1, IIf(Len(sTexto) > 57, 85, 57)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'sDatos = sDatos & Space(5) & CON & Mid(psPersNombre, 1, 20) & COFF & PrnSet("Esp", 7.5) & oImpresora.gPrnSaltoLinea & PrnSet("EspN") & oImpresora.gPrnSaltoLinea
'sDatos = sDatos & Space(5) & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & BOFF & PrnSet("Esp", 4.5) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
'sDatos = sDatos & String(5, oImpresora.gPrnSaltoLinea)
'FormaOrdenPagoTalon = ImpreCarEsp(sDatos)

'Nuevo Formato
sDatos = PrnSet("Esp", 11) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
sDatos = sDatos & PrnSet("Esp", 1) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
sDatos = sDatos & Space(53) & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & Space(6) & Format(pnImporte, "###,###,##0.00") & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
sDatos = sDatos & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
sDatos = sDatos & PrnSet("Esp", 13) & oImpresora.gPrnSaltoLinea
sDatos = sDatos & Space(27) & BON & IIf(Len(Trim(psPersNombre)) > 59, CON, "") & Mid(Trim(psPersNombre) & String(59, "*"), 1, IIf(Len(psPersNombre) > 59, 89, 59)) & COFF
sDatos = sDatos & PrnSet("Esp", 8) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
sDatos = sDatos & Space(5) & BON & Format(pnImporte, "###,###,##0.00") & BOFF
sTexto = ConvNumLet(pnImporte, psMoneda, True)
sDatos = sDatos & PrnSet("Esp", 2) & oImpresora.gPrnSaltoLinea & PrnSet("EspN") & oImpresora.gPrnSaltoLinea
sDatos = sDatos & Space(21) & IIf(Len(sTexto) > 60, CON, "") & Mid(Trim(sTexto) & String(60, "*"), 1, IIf(Len(sTexto) > 60, 90, 60)) & COFF & BOFF & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
sDatos = sDatos & PrnSet("Esp", 9) & oImpresora.gPrnSaltoLinea & PrnSet("EspN")
sDatos = sDatos & Space(1) & CON & Mid(psPersNombre, 1, 28) & COFF & PrnSet("Esp", 5) & oImpresora.gPrnSaltoLinea & PrnSet("EspN") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
sDatos = sDatos & Space(5) & BON & Mid(pdFecha, 1, 2) & "  " & Mid(pdFecha, 4, 2) & "  " & Right(pdFecha, 4) & BOFF & PrnSet("Esp", 2.5) & PrnSet("EspN")
sDatos = sDatos & String(4, oImpresora.gPrnSaltoLinea)
FormaOrdenPagoTalon = ImpreCarEsp(sDatos)

End Function
