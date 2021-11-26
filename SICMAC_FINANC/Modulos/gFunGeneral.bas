Attribute VB_Name = "gFunGeneral"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A7EDE77033C"
'Módulo de datos de Contabilidad
Option Base 0
Option Explicit
Dim unidad(0 To 9) As String
Dim decena(0 To 9) As String
Dim centena(0 To 10) As String
Dim deci(0 To 9) As String
Dim otros(0 To 15) As String

Public Sub CentraForm(frmCentra As Form)
    frmCentra.Move (Screen.Width - frmCentra.Width) / 2, (Screen.Height - frmCentra.Height) / 2, frmCentra.Width, frmCentra.Height
    If Dir(App.Path & gsRutaIcono) <> "" Then
       ' frmCentra.Icon = LoadPicture(App.path & "\BMP\ec.ico")
        frmCentra.Icon = LoadPicture(App.Path & gsRutaIcono)
        'Me.Icon = LoadPicture(App.path & gsRutaIcono)
    End If
End Sub

Public Sub CargaCombo(ByRef Combo As ComboBox, rs As ADODB.Recordset, Optional pbLimpiaCombo As Boolean = True)
Dim Campo As ADODB.Field
Dim lsDato As String
If rs Is Nothing Then Exit Sub
If pbLimpiaCombo Then Combo.Clear
Do While Not rs.EOF
    lsDato = ""
    For Each Campo In rs.Fields
        lsDato = lsDato & Campo.value & Space(50)
    Next
    lsDato = Mid(lsDato, 1, Len(lsDato) - 50)
    Combo.AddItem lsDato
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub

Public Sub fEnfoque(ctrControl As Control)
    ctrControl.SelStart = 0
    ctrControl.SelLength = Len(ctrControl.Text)
End Sub

Function ValidaFecha(cadfec As String) As String
Dim I As Integer
    If Len(cadfec) <> 10 Then
        ValidaFecha = "Fecha No Valida"
        Exit Function
    End If
    For I = 1 To 10
        If I = 3 Or I = 6 Then
            If Mid(cadfec, I, 1) <> "/" Then
                ValidaFecha = "Fecha No Valida"
                Exit Function
            End If
        Else
            If Asc(Mid(cadfec, I, 1)) < 48 Or Asc(Mid(cadfec, I, 1)) > 57 Then
                ValidaFecha = "Fecha No Valida"
                Exit Function
            End If
        End If
    Next I
'validando dia
If Val(Mid(cadfec, 1, 2)) < 1 Or Val(Mid(cadfec, 1, 2)) > 31 Then
    ValidaFecha = "Dia No Valido"
    Exit Function
End If
'validando mes
If Val(Mid(cadfec, 4, 2)) < 1 Or Val(Mid(cadfec, 4, 2)) > 12 Then
    ValidaFecha = "Mes No Valido"
    Exit Function
End If
'validando año
If Val(Mid(cadfec, 7, 4)) < 1900 Or Val(Mid(cadfec, 7, 4)) > 9972 Then
    ValidaFecha = "Año No Valido"
    Exit Function
End If
'validando con isdate
If IsDate(cadfec) = False Then
    ValidaFecha = "Mes o Dia No Valido"
    Exit Function
End If
ValidaFecha = ""
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

Public Function NumerosDecimales(cTexto As TextBox, intTecla As Integer, _
    Optional nLongitud As Integer = 12, Optional nDecimal As Integer = 2, Optional pbSoloPositivos As Boolean = True) As Integer
    Dim cValidar As String
    Dim cCadena As String
    cCadena = cTexto
    cValidar = "-0123456789."
    If pbSoloPositivos Then
        cValidar = "0123456789."
    End If
    
    If InStr(".", Chr(intTecla)) <> 0 Then
        If InStr(cCadena, ".") <> 0 Then
            intTecla = 0
            Beep
        ElseIf intTecla > 26 Then
            If InStr(cValidar, Chr(intTecla)) = 0 Then
                intTecla = 0
                Beep
            End If
        End If
    ElseIf intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) = 0 Then
            intTecla = 0
            Beep
        End If
    End If
    Dim vPosCur As Byte
    Dim vPosPto As Byte
    Dim vNumDec As Byte
    Dim vNumLon As Byte
    
    vPosPto = InStr(cTexto.Text, ".")
    vPosCur = cTexto.SelStart
    vNumLon = Len(cTexto)
    If vPosPto > 0 Then
        vNumDec = Len(Mid(cTexto, vPosPto + 1))
    End If
    If vPosPto > 0 Then
        If cTexto.SelLength <> Len(cTexto) Then
        If ((vNumDec >= nDecimal And cTexto.SelStart >= vPosPto) Or _
        (vNumLon >= nLongitud)) _
        And intTecla <> vbKeyBack And intTecla <> vbKeyDecimal And intTecla <> vbKeyReturn Then
            intTecla = 0
            Beep
        End If
        End If
    Else
        If vNumLon >= nLongitud And intTecla <> vbKeyBack _
        And intTecla <> vbKeyReturn Then
            intTecla = 0
            Beep
        End If
        If (vNumLon - cTexto.SelStart) > nDecimal And intTecla = 46 Then
            intTecla = 0
            Beep
        End If
    End If
    NumerosDecimales = intTecla
End Function

Public Function NumerosDecimales4(cTexto As TextBox, intTecla As Integer, _
    Optional nLongitud As Integer = 12, Optional nDecimal As Integer = 4, Optional pbSoloPositivos As Boolean = True) As Integer
    Dim cValidar As String
    Dim cCadena As String
    cCadena = cTexto
    cValidar = "-0123456789."
    If pbSoloPositivos Then
        cValidar = "0123456789."
    End If
    
    If InStr(".", Chr(intTecla)) <> 0 Then
        If InStr(cCadena, ".") <> 0 Then
            intTecla = 0
            Beep
        ElseIf intTecla > 26 Then
            If InStr(cValidar, Chr(intTecla)) = 0 Then
                intTecla = 0
                Beep
            End If
        End If
    ElseIf intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) = 0 Then
            intTecla = 0
            Beep
        End If
    End If
    Dim vPosCur As Byte
    Dim vPosPto As Byte
    Dim vNumDec As Byte
    Dim vNumLon As Byte
    
    vPosPto = InStr(cTexto.Text, ".")
    vPosCur = cTexto.SelStart
    vNumLon = Len(cTexto)
    If vPosPto > 0 Then
        vNumDec = Len(Mid(cTexto, vPosPto + 1))
    End If
    If vPosPto > 0 Then
        If cTexto.SelLength <> Len(cTexto) Then
        If ((vNumDec >= nDecimal And cTexto.SelStart >= vPosPto) Or _
        (vNumLon >= nLongitud)) _
        And intTecla <> vbKeyBack And intTecla <> vbKeyDecimal And intTecla <> vbKeyReturn Then
            intTecla = 0
            Beep
        End If
        End If
    Else
        If vNumLon >= nLongitud And intTecla <> vbKeyBack _
        And intTecla <> vbKeyReturn Then
            intTecla = 0
            Beep
        End If
        If (vNumLon - cTexto.SelStart) > nDecimal And intTecla = 46 Then
            intTecla = 0
            Beep
        End If
    End If
    NumerosDecimales4 = intTecla
End Function '***NAGL ERS 079-2016 20170407

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
'EJVG20111207 *******************************************
Public Function DigitoRPM(intTecla As Integer) As Boolean
    Dim cValidar As String
    cValidar = "#*"
    DigitoRPM = True
    If InStr(cValidar, Chr(intTecla)) = 0 Then
        DigitoRPM = False
        Beep
    End If
End Function
'*******************************************************
'RUTINA VALIDA EL INGRESO DE UN NUMERO MAXIMO DE LINEAS
'*******************************************************
'FECHA CREACION : 24/06/99  -   MAVF
'MODIFICACION:
'**********************************************
Public Function intfLineas(cCadena As String, intTecla As Integer, intLinea As Integer) As Integer
Dim vLineas As Byte
Dim x As Byte
    If intTecla = 13 Then
        For x = 1 To Len(cCadena)
            If Mid(cCadena, x, 1) = Chr(13) Then
                vLineas = vLineas + 1
            End If
        Next x
        If vLineas >= intLinea Then
            MsgBox " No se permite mas lineas ", vbInformation, " Aviso "
            intTecla = 0
            Beep
        End If
    End If
    intfLineas = intTecla
End Function
Public Function Letras(intTecla As Integer, Optional lbMayusculas As Boolean = True) As Integer
If lbMayusculas Then
    Letras = Asc(UCase(Chr(intTecla)))
Else
    Letras = Asc(LCase(Chr(intTecla)))
End If
End Function
Public Function SoloLetras(intTecla As Integer) As Integer
Dim cValidar  As String
    cValidar = "0123456789+:;'<>?_=+[]{}|!@#$%^&()*"
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) <> 0 Then
            intTecla = 0
            Beep
        End If
    End If
    SoloLetras = intTecla
End Function
Public Function LetrasNumeros(intTecla As Integer) As Integer
Dim cValidar  As String
Dim cNumero As String
    cValidar = "+:;-/\'<>?_=+[]{}|!@#$%^&()*°"
    cNumero = "0123456789"
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) <> 0 Then
            intTecla = 0
            Beep
        End If
    End If
    If InStr(cNumero, Chr(intTecla)) = 0 Then
        LetrasNumeros = Asc(UCase(Chr(intTecla)))
    Else
        LetrasNumeros = intTecla
    End If
End Function '***NAGL ERS012-2017 20170710
'***************************************************
'* Funcion:  Convierte un valor Numerico a su corres
'*           pondiente descripción alfabetica
'***************************************************
'FECHA CREACION : 24/06/99  -   FAOS
'MODIFICACION:
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
'***************************************************
'* Funcion:  Es llamada desde NumLet
'***************************************************
'FECHA CREACION : 24/06/99  -   FAOS
'MODIFICACION:
'***************************************************
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
'***************************************************
'* Funcion:  Convierte un valor Fecha a su corres
'*           pondiente descripción alfabetica
'***************************************************
'FECHA CREACION : 24/06/99  -   FAOS
'MODIFICACION:
'***************************************************
Public Function ArmaFecha(dtmFechas As Date) As String
    Dim txtMeses As String
    txtMeses = Choose(Month(dtmFechas), "Enero", "Febrero", "Marzo", "Abril", _
                                        "Mayo", "Junio", "Julio", "Agosto", _
                                        "Setiembre", "Octubre", "Noviembre", "Diciembre")
                                        
    ArmaFecha = Day(dtmFechas) & " de " & txtMeses & " de " & Year(dtmFechas)

End Function

'***************************************************
'* VALIDA LA HORA INGRESADA EN 23 HORAS, 59 SEGUNDOS
'***************************************************
'FECHA CREACION : 25/07/99  -   MAVF
'MODIFICACION:
'***************************************************
Public Function ValidaHora(lsControl As Control) As Boolean
   If Mid(lsControl, 1, 2) >= 0 And Mid(lsControl, 1, 2) <= 23 Then
        If Mid(lsControl, 4, 2) >= 0 And Mid(lsControl, 4, 2) <= 59 Then
            ValidaHora = True
        Else
            ValidaHora = False
            MsgBox "Minuto no es válido", vbInformation, "Aviso"
            lsControl.SetFocus
            lsControl.SelStart = 3
            lsControl.SelLength = 2
            Exit Function
        End If
    Else
        ValidaHora = False
        MsgBox "Hora no es válido", vbInformation, "Aviso"
        lsControl.SetFocus
        lsControl.SelStart = 0
        lsControl.SelLength = 2
        Exit Function
    End If
End Function
Public Function Residuo(Dividendo As Currency, Divisor As Currency) As Boolean
Dim x As Currency
x = Round(Dividendo / Divisor, 0)
Residuo = True
x = x * Divisor
If x <> Dividendo Then
   Residuo = False
End If
End Function
Public Function ConvNumLet(nNumero As Currency, Optional lSoloText As Boolean = True, Optional lSinMoneda As Boolean = False, Optional pnMoneda As Moneda = 0) As String
Dim sCent As String
Dim xValor As Single
Dim vMoneda As String
Dim cNumero As String
cNumero = Format(nNumero, gsFormatoNumeroView)
xValor = nNumero - Int(nNumero)
If xValor = 0 Then
   sCent = " Y 00/100 "
Else
   sCent = " Y " & Right(Trim(cNumero), 2) & "/100 "
End If
If pnMoneda <> 0 Then
    '''vMoneda = IIf(pnMoneda = gMonedaNacional, "NUEVOS SOLES", "DOLARES AMERICANOS") 'marg ers044-2016
    vMoneda = IIf(pnMoneda = gMonedaNacional, StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES AMERICANOS") 'marg ers044-2016
End If
If Not lSoloText Then
   ConvNumLet = Trim(gsSimbolo) & " " & Trim(Format(nNumero, "###,###,##0.00#")) & " ("
End If
ConvNumLet = ConvNumLet & Trim(UCase(NumLet(CStr(nNumero), 0))) & sCent & " " & IIf(lSinMoneda, "", Trim(vMoneda)) & IIf(lSoloText, "", ")")
End Function

Public Function BuscaDato(ByVal Criterio As String, rsAdo As ADODB.Recordset, ByVal start As Long, ByVal lMsg As Boolean) As Boolean
Dim Pos As Variant
On Error GoTo Errbusq
   BuscaDato = False
   Pos = rsAdo.Bookmark
   rsAdo.Find Criterio, IIf(start = 1, 0, start + 1), adSearchForward, 1
   If rsAdo.EOF Then
      rsAdo.Bookmark = Pos
      If lMsg Then
         MsgBox " ! Dato no encontrado... ! ", vbExclamation, "Error de Busqueda"
         BuscaDato = False
      End If
   Else
      BuscaDato = True
   End If
Exit Function
Errbusq:
   MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Function

Public Function FechaHora(psFecha As Date) As String
    FechaHora = Format(psFecha & Space(1) & GetHoraServer, gsFormatoFechaHora)
End Function
Public Function FillNum(intNumero As String, intLenNum As Integer, ChrFil As String) As String
  FillNum = Left(String(intLenNum, ChrFil), (Len(String(intLenNum, ChrFil)) - Len(Trim(intNumero)))) + Trim(intNumero)
End Function
Public Sub RSClose(rs As ADODB.Recordset)
If Not rs Is Nothing Then
   If rs.State = adStateOpen Then
       rs.Close
       Set rs = Nothing
   End If
End If
End Sub
Public Function RSVacio(rs1 As ADODB.Recordset) As Boolean
 RSVacio = (rs1.BOF And rs1.EOF)
End Function
Public Function TextErr(sMsg As String) As String
Dim nLen As Integer
Dim lsDato As String
If InStr(sMsg, "PRIMARY KEY") > 0 Then
    sMsg = "Imposible Insertar Clave Duplicada..."
End If
If InStr(sMsg, "COLUMN REFERENCE") > 0 And InStr(sMsg, "DELETE") > 0 Then
    lsDato = Mid(sMsg, InStr(sMsg, "tabla '") + 7)
    lsDato = Mid(lsDato, 1, InStr(lsDato, "'") - 1)
    sMsg = "Imposible Eliminar Dato. Codigo ha sido utilizado en la tabla <" & UCase(lsDato) & ">"
End If
nLen = InStr(1, sMsg, "*", vbTextCompare)

TextErr = Mid(sMsg, nLen + 1, Len(sMsg))
End Function
Public Function PstaNombre(psNombre As String, Optional pbNombApell As Boolean = False) As String
Dim Total As Long
Dim Pos As Long
Dim CadAux As String
Dim lsApellido As String
Dim lsNOmbre As String
Dim lsMaterno As String
Dim lsConyugue As String
Dim CadAux2 As String
Dim posAux As Integer
Dim lbVda As Boolean
lbVda = False
Total = Len(Trim(psNombre))
Pos = InStr(psNombre, "/")
If Pos <> 0 Then
    lsApellido = Left(psNombre, Pos - 1)
    CadAux = Mid(psNombre, Pos + 1, Total)
    Pos = InStr(CadAux, "\")
    If Pos <> 0 Then
        lsMaterno = Left(CadAux, Pos - 1)
        CadAux = Mid(CadAux, Pos + 1, Total)
        Pos = InStr(CadAux, ",")
        If Pos > 0 Then
            CadAux2 = Left(CadAux, Pos - 1)
            posAux = InStr(CadAux, "VDA")
            If posAux = 0 Then
                lsConyugue = CadAux2
            Else
                lbVda = True
                lsConyugue = CadAux2
            End If
        Else
            lsMaterno = CadAux
        End If
    Else
        CadAux = Mid(CadAux, Pos + 1, Total)
        Pos = InStr(CadAux, ",")
        If Pos <> 0 Then
            lsMaterno = Left(CadAux, Pos - 1)
            lsConyugue = ""
        Else
            lsMaterno = CadAux
        End If
    End If
    lsNOmbre = Mid(CadAux, Pos + 1, Total)
    If pbNombApell = True Then
        If Len(Trim(lsConyugue)) > 0 Then
            PstaNombre = Trim(lsNOmbre) & " " & Trim(lsApellido) & " " & Trim(lsMaterno) & IIf(lbVda = False, " DE ", " ") & Trim(lsConyugue)
        Else
            PstaNombre = Trim(lsNOmbre) & " " & Trim(lsApellido) & " " & Trim(lsMaterno)
        End If
    Else
        If Len(Trim(lsConyugue)) > 0 Then
            PstaNombre = Trim(lsApellido) & " " & Trim(lsMaterno) & IIf(lbVda = False, " DE ", " ") & Trim(lsConyugue) & " " & Trim(lsNOmbre)
        Else
            PstaNombre = Trim(lsApellido) & " " & Trim(lsMaterno) & " " & Trim(lsNOmbre)
        End If
    End If
Else
    PstaNombre = Trim(psNombre)
End If
End Function
Public Function GetTipCambio(pdFecha As Date) As Boolean
    Dim oDGeneral As nTipoCambio
    On Error GoTo GetTipCambioErr
    Set oDGeneral = New nTipoCambio
    GetTipCambio = True
    gnTipCambio = 0
    gnTipCambioV = 0
    gnTipCambioC = 0
    gnTipCambioVE = 0
    gnTipCambioCE = 0
    gnTipCambioPonderado = 0
    gnTipCambioPonderadoVenta = 0
    
    gnTipCambio = oDGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
    gnTipCambioV = oDGeneral.EmiteTipoCambio(pdFecha, TCVenta)
    gnTipCambioC = oDGeneral.EmiteTipoCambio(pdFecha, TCCompra)
    gnTipCambioVE = oDGeneral.EmiteTipoCambio(pdFecha, TCVentaEsp)
    gnTipCambioCE = oDGeneral.EmiteTipoCambio(pdFecha, TCCompraEsp)
    gnTipCambioPonderado = oDGeneral.EmiteTipoCambio(pdFecha, TCPonderado)
    gnTipCambioPonderadoVenta = oDGeneral.EmiteTipoCambio(pdFecha, 7)
    
    If gnTipCambio = 0 Then
        MsgBox "Tipo de Cambio aun no definido", vbInformation, "Aviso"
        GetTipCambio = False
    End If
     Exit Function
GetTipCambioErr:
        MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Function
Public Function GeneraMovNroActualiza(pdFecha As Date, psCodUser As String, psCodCmac As String, psCodAge As String) As String
GeneraMovNroActualiza = Format(pdFecha & " " & GetHoraServer, gsFormatoMovFechaHora) & psCodCmac & Right(psCodAge, 2) & "00" & psCodUser
End Function

Public Sub EnviaPrevio(psImpre As String, psTitulo As String, ByVal pnLinPage As Integer, Optional plCondensado As Boolean = False)
Dim clsPrevioFinan As clsPrevioFinan
Set clsPrevioFinan = New clsPrevioFinan
clsPrevioFinan.Show psImpre, psTitulo, plCondensado, pnLinPage, gImpresora
Set clsPrevioFinan = Nothing
End Sub

Public Function nVal(psImporte As String) As Currency
nVal = 0
If psImporte <> "" Then
   nVal = Format(psImporte, gsFormatoNumeroDato)
End If
End Function

Public Sub CambiaTamañoCombo(ByRef cboCombo As ComboBox, Optional nTamaño As Long = 200)
SendMessage cboCombo.hwnd, CB_SETDROPPEDWIDTH, nTamaño, 0
End Sub

Public Function ValidaAnio(nAnio As Integer) As Boolean
ValidaAnio = False
If nAnio > Year(gdFecSis) Then
   MsgBox "Año no puede ser mayor a Periodo Actual", vbInformation, "Aviso"
   Exit Function
End If
If Year(gdFecSis) - nAnio > 5 Then
   MsgBox "El Sistema sólo permite procesos hasta 5 años anteriores", vbInformation, "Aviso"
   Exit Function
End If
ValidaAnio = True
End Function
Public Function GetObjetosOpeCta(ByVal psOpeCod As String, ByVal psObjetoOrden As String, _
                                ByVal psCtaContCod As String, ByRef psRaiz As String, _
                                Optional ByVal psFiltro As String = "", Optional ByVal psFiltroAdd As String = "") As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim lsFiltro As String

Dim oCtaCont As DCtaCont
Dim oOpe As DOperacion

Set oOpe = New DOperacion
Set oCtaCont = New DCtaCont

Set rs1 = New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs1 = oOpe.GetOpeObj(psOpeCod, psObjetoOrden, psCtaContCod, psFiltro, psFiltroAdd)
If rs1.State = adStateOpen Then
    If Not rs1.EOF And Not rs1.BOF Then
        Set rs = rs1
    End If
Else
    Set rs1 = oCtaCont.CargaCtaObj(psCtaContCod, , True)
    If Not rs1.EOF And Not rs1.BOF Then
        If psFiltro <> "" Then
            lsFiltro = psFiltro & psFiltroAdd
        Else
            lsFiltro = Trim(rs1!cCtaObjFiltro) & psFiltroAdd
        End If
        Set rs = GetrsObjetos(Val(rs1!cObjetoCod), psCtaContCod, lsFiltro, psRaiz)
    End If
    rs1.Close
    Set rs1 = Nothing
End If
Set oCtaCont = Nothing
Set oOpe = Nothing
Set GetObjetosOpeCta = rs
End Function
Public Function GetrsObjetos(ByVal psObjetoCod As TpoObjetos, ByVal psCtaContCod As String, _
                            ByVal lsFiltro As String, ByRef psRaiz As String) As ADODB.Recordset

Dim oRHAreas As DActualizaDatosArea
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo
Dim oContFunct As NContFunciones
Dim rs As ADODB.Recordset
Set oEfect = New Defectivo
Set oCtaIf = New NCajaCtaIF
Set oRHAreas = New DActualizaDatosArea
Set oContFunct = New NContFunciones
Set rs = New ADODB.Recordset

Select Case Val(psObjetoCod)
    Case ObjCMACAgencias
        Set rs = oRHAreas.GetAgencias(lsFiltro)
    Case ObjCMACAgenciaArea
        psRaiz = "Unidades Organizacionales"
        Set rs = oRHAreas.GetAgenciasAreas(lsFiltro)
    Case ObjCMACArea
        Set rs = oRHAreas.GetAreas(lsFiltro)
    Case ObjEntidadesFinancieras
        psRaiz = "Cuentas de Entidades Financieras"
        Set rs = oCtaIf.GetCtasInstFinancieras(lsFiltro, psCtaContCod)
    Case ObjDescomEfectivo
        Set rs = oEfect.GetBilletajes(lsFiltro)
    Case ObjPersona
        Set rs = Nothing
    Case Else
        Set rs = oContFunct.GetObjetos(psObjetoCod)
End Select

Set GetrsObjetos = rs

Set oCtaIf = Nothing
Set oEfect = Nothing
Set oRHAreas = Nothing
Set oContFunct = Nothing
End Function
Public Sub RSLlenaCombo(prs As ADODB.Recordset, psCombo As ComboBox, Optional pnPosCod As Integer = 0, Optional pnPosDes As Integer = 1, Optional pbPresentaCodigo As Boolean = True)
If Not prs Is Nothing Then
   If Not prs.EOF Then
      psCombo.Clear
      Do While Not prs.EOF
         If pbPresentaCodigo Then
            psCombo.AddItem Trim(prs(pnPosDes)) & Space(100) & Trim(prs(pnPosCod))
         Else
            psCombo.AddItem Trim(prs(pnPosCod)) & "  " & Trim(prs(pnPosDes))
         End If
         prs.MoveNext
      Loop
   End If
End If
End Sub

Public Function GetHoraServer() As String
Dim oConect As DConecta
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
GetHoraServer = oConect.GetHoraServer()
oConect.CierraConexion
Set oConect = Nothing
End Function

Public Function GetFechaHoraServer() As String
Dim oConect As DConecta
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
GetFechaHoraServer = oConect.GetFechaHoraServer()
oConect.CierraConexion
Set oConect = Nothing
End Function

Public Sub EnviaImpresion(psImpre As String, ByVal pnLinPage As Integer, Optional plCondensado As Boolean = False)
Dim clsPrevioFinan As New PrevioFinan.clsPrevioFinan
clsPrevioFinan.ShowImpreSpool psImpre, plCondensado, pnLinPage
Set clsPrevioFinan = Nothing
End Sub

Public Function CodigoOperacion(psOpeCod, pnMoneda As Integer) As String
CodigoOperacion = Mid(psOpeCod, 1, 2) & pnMoneda & Mid(psOpeCod, 4, 3)
End Function

Public Function GetCaptionForm(psTexto As String, psOpeCod, pnAncho As Long) As String
Dim lsTexto As String
Dim lnSpace As Integer
lnSpace = Int(pnAncho / 82) - 12 - Len(psTexto)
If Mid(gsOpeCod, 3, 1) = "2" Then
   lsTexto = psTexto & Space(lnSpace) & "M.Extranjera"
ElseIf Mid(gsOpeCod, 3, 1) = "1" Then
   lsTexto = psTexto & Space(lnSpace) & "M.Nacional"
Else
   lsTexto = psTexto
End If
GetCaptionForm = lsTexto
End Function

Public Function RSMuestraLista(prs As ADODB.Recordset, Optional pnCol As Integer = 0) As String
Dim lsLista As String
If Not prs Is Nothing Then
   lsLista = ""
   Do While Not prs.EOF
      lsLista = lsLista & "'" & prs(pnCol) & "',"
      prs.MoveNext
   Loop
   If lsLista <> "" Then
      lsLista = Left(lsLista, Len(lsLista) - 1)
   End If
End If
RSMuestraLista = lsLista
End Function

Public Function RSMuestraListaCadenaNew(prs As ADODB.Recordset, Optional pnCol As Integer = 0) As String
Dim lsLista As String
If Not prs Is Nothing Then
   lsLista = ""
   Do While Not prs.EOF
      lsLista = lsLista & prs(pnCol) & ","
      prs.MoveNext
   Loop
   If lsLista <> "" Then
      lsLista = Left(lsLista, Len(lsLista) - 1)
   End If
End If
RSMuestraListaCadenaNew = lsLista
End Function 'NAGL 20190612 Según INC1903010011

Public Function ClaveCorrectaNT(ByVal psTxtUserName As String, ByVal psTxtPass As String, ByVal psTxtDominio As String) As Boolean
    Dim ClsNT As DLLWinNT.ClsWinNT
    Set ClsNT = New DLLWinNT.ClsWinNT
    ClaveCorrectaNT = ClsNT.SSPValidateUser(psTxtUserName, psTxtDominio, psTxtPass)
    Set ClsNT = Nothing
End Function

Public Function GetFechaMov(cMovNro, lDia As Boolean) As String
Dim lFec As Date
lFec = Mid(cMovNro, 7, 2) & "/" & Mid(cMovNro, 5, 2) & "/" & Mid(cMovNro, 1, 4)
If lDia Then
   GetFechaMov = Format(lFec, gsFormatoFechaView)
Else
   GetFechaMov = Format(lFec, gsFormatoFecha)
End If
End Function

Public Function RecordSetAdiciona(ByRef prsDat As ADODB.Recordset, ByVal prs As ADODB.Recordset)
Dim nCol As Integer
RecordSetDefineCampos prsDat, prs
Do While Not prs.EOF
    prsDat.AddNew
    For nCol = 0 To prs.Fields.Count - 1
        If prs.Fields(nCol).Type = adVarChar Then
            prsDat.Fields(nCol).value = Left(prs.Fields(nCol).value, prsDat.Fields(nCol).DefinedSize)
        Else
            prsDat.Fields(nCol).value = prs.Fields(nCol).value
        End If
    Next
    prsDat.Update
    prs.MoveNext
Loop
End Function

Public Function RecordSetDefineCampos(ByRef prsDat As ADODB.Recordset, ByVal prs As ADODB.Recordset)
Dim nCol As Integer

If prsDat Is Nothing Then
    Set prsDat = New ADODB.Recordset
End If
If prsDat.State = adStateClosed Then
    For nCol = 0 To prs.Fields.Count - 1
        With prs.Fields(nCol)
            If .Type = adVarChar Then
                prsDat.Fields.Append .Name, .Type, 250, .Attributes
            Else
                prsDat.Fields.Append .Name, .Type, .DefinedSize, .Attributes
            End If
        End With
    Next
    prsDat.Open
End If
End Function

Public Function PrnVal(pnVal As Currency, pnLen As Integer, pnDec As Integer, Optional lCero As Boolean = True) As String
Dim sFormat As String
 sFormat = "###,###,###,##0" & IIf(pnDec > 0, "." & String(pnDec, "0"), "")
 PrnVal = Right(Space(pnLen) & IIf(Not IsNull(pnVal) And (pnVal <> 0 Or lCero), Format(pnVal, sFormat), ""), pnLen)
End Function

Public Function AdicionaRecordSet(ByRef prsDat As ADODB.Recordset, ByVal prs As ADODB.Recordset)
Dim nCol As Integer
Do While Not prs.EOF
    If Not prsDat Is Nothing Then
        If prsDat.State = adStateClosed Then
            For nCol = 0 To prs.Fields.Count - 1
                With prs.Fields(nCol)
                    prsDat.Fields.Append .Name, .Type, .DefinedSize, .Attributes
                End With
            Next
            prsDat.Open
        End If
        prsDat.AddNew
        For nCol = 0 To prs.Fields.Count - 1
            prsDat.Fields(nCol).value = prs.Fields(nCol).value
        Next
        prsDat.Update
    End If
    prs.MoveNext
Loop
If Not prsDat Is Nothing Then
    If prsDat.RecordCount > 0 Then
        prsDat.MoveFirst
    End If
End If
End Function

Public Function MuestraListaRecordSet(prs As ADODB.Recordset, Optional pnCol As Integer = 0) As String
Dim lsLista As String
If Not prs Is Nothing Then
   lsLista = ""
   Do While Not prs.EOF
      lsLista = lsLista & "'" & prs(pnCol) & "',"
      prs.MoveNext
   Loop
   If lsLista <> "" Then
      lsLista = Left(lsLista, Len(lsLista) - 1)
   End If
End If
MuestraListaRecordSet = lsLista
End Function

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

Public Function LlenaCerosSUCAVE(ByVal sNumero As Double, Optional nTipoFormato As Integer = 1)
 
Dim sTempo As String
Dim I  As Integer

If sNumero = 0 Then
    If nTipoFormato = 1 Then
        sTempo = "              0"
    ElseIf nTipoFormato = 2 Then
        sTempo = "             0"
    ElseIf nTipoFormato = 3 Then
        sTempo = "     0"
    End If
    
Else
    If nTipoFormato = 1 Then
        'antes: sTempo = Format(sNumero, "###########0.00")
        
        sTempo = Format(sNumero, "############.00")
       
        For I = 1 To 16 - Len(sTempo)
            sTempo = " " & sTempo
        Next
    ElseIf nTipoFormato = 2 Then
        sTempo = Format(sNumero, "###########0")
        For I = 1 To 15 - Len(sTempo)
            sTempo = " " & sTempo
        Next
    ElseIf nTipoFormato = 3 Then
        'antes:  sTempo = Format(sNumero, "###0.00")
        
        sTempo = Format(sNumero, "####.00")
        
        For I = 1 To 7 - Len(sTempo)
            sTempo = " " & sTempo
        Next
    End If
    
     
    If nTipoFormato = 1 Then
        sTempo = Mid(sTempo, 1, 13) & Mid(sTempo, 15, 2)
    ElseIf nTipoFormato = 3 Then
        sTempo = Mid(sTempo, 1, 4) & Mid(sTempo, 6, 2)
    End If
    
End If

LlenaCerosSUCAVE = sTempo
End Function

Public Function TipoCambioCierre(pnAnio As Integer, pnMes As Integer, Optional pbMesCerrado As Boolean = True) As Currency
Dim oCambio As nTipoCambio
Dim sFecha  As Date
    If pnMes <= 0 Or pnMes > 12 Or pnAnio < 1900 Then
        Exit Function
    End If
    sFecha = CDate("01/" & Format(pnMes, "00") & "/" & Trim(pnAnio))
    sFecha = DateAdd("m", 1, sFecha)
    If Not pbMesCerrado Then
       sFecha = sFecha - 1
    End If
    Set oCambio = New nTipoCambio
    TipoCambioCierre = Format(oCambio.EmiteTipoCambio(sFecha, TCFijoMes), "#,##0.0000")
    Set oCambio = Nothing
End Function


Public Function fgFechaHoraGrab(ByVal psMovNro As String) As String
    fgFechaHoraGrab = Mid(psMovNro, 1, 4) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 7, 2) & " " & Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2)
End Function



Public Function ReemplazaApostrofe(ByVal lsCadena As String) As String
    ReemplazaApostrofe = Replace(lsCadena, "'", "''", , , vbTextCompare)
End Function

Public Function CadDerecha(psCadena As String, lsTam As Integer) As String
    CadDerecha = Format(psCadena, "!" & String(lsTam, "@"))
End Function

Public Function GetRsNewDeListView(lstLista As ListView, psFormato As String, Optional pnColIni As Integer = 1) As ADODB.Recordset
'JHVP ==> Adeudos
'Formatos:  0-General
'           1-Solo Letras
'           2-Decimales
'           3-Enteros
'           4-Billetaje ¿?
'           5-FechaHora
'           6-Hora

'Col Alignment

' 0 El contenido de la celda se alinea a la izquierda, arriba.
' 1 Predeterminado para las cadenas. El contenido de la celda se alinea a la izquierda, centro.
' 2 El contenido de la celda se alinea a la izquierda, abajo.
' 3 El contenido de la celda se alinea al centro y arriba.
' 4 El contenido de la celda se alinea al centro, centro.
' 5 El contenido de la celda se alinea al centro, abajo.
' 6 El contenido de la celda se alinea a la derecha, arriba.
' 7 Predeterminado para los números. El contenido de la celda se alinea a la derecha, centro.
' 8 El contenido de la celda se alinea a la derecha, abajo.
' 9 El contenido de la celda tiene la alineación general. Esta corresponde a "izquierda, centro" para las cadenas y "derecha, centro" para los números.



Dim I As Long
Dim j As Long
Dim rsAux As ADODB.Recordset
Dim lnFila As Long
Dim lncol As Long
Dim lsTipoDato As DataTypeEnum
Dim lsTamCampo As Long
Dim lnFormatoCol As Long
If lstLista.ListItems(1).SubItems(1) <> "" Then
    lnFila = 0
    'formamos generamos del recordset
    Set rsAux = New ADODB.Recordset
    For I = pnColIni To lstLista.ColumnHeaders.Count - 1
        If Len(Trim(lstLista.ListItems(lnFila + 1).SubItems(I))) >= 16 Then
            lsTipoDato = adVarChar
        Else
            lnFormatoCol = DeterminaFormato(psFormato, I, lstLista.ColumnHeaders.Count)
            If (lnFormatoCol = 2 Or lnFormatoCol = 3 Or lnFormatoCol = 5) Then ' And lstLista.ColumnHeaders(i).Alignment >= 7 Then
                lsTipoDato = adDouble
            Else
                If ValidaFecha(lstLista.ListItems(lnFila + 1).SubItems(I)) = "" Then
                    lsTipoDato = adDate
                Else
                    lsTipoDato = adVarChar
                End If
            End If
        End If
        If lsTipoDato = adVarChar Then
            rsAux.Fields.Append lstLista.ColumnHeaders(I + 1).Text, lsTipoDato, 400, adFldMayBeNull
        Else
            rsAux.Fields.Append lstLista.ColumnHeaders(I + 1).Text, lsTipoDato, , adFldMayBeNull
        End If
    Next
    rsAux.Open

    For I = 1 To lstLista.ListItems.Count
        rsAux.AddNew
        'columnas
        For j = pnColIni To lstLista.ColumnHeaders.Count - 1
            'No consideramos checkboz
            If rsAux.Fields(lstLista.ColumnHeaders(j + 1).Text).Type = adDouble Then
                rsAux.Fields(lstLista.ColumnHeaders(j + 1).Text) = CCur(IIf(lstLista.ListItems(I).SubItems(j) = "", "0", lstLista.ListItems(I).SubItems(j)))
            Else
                If rsAux.Fields(lstLista.ColumnHeaders(j + 1).Text).Type = adDate Then
                    rsAux.Fields(lstLista.ColumnHeaders(j + 1).Text) = IIf(lstLista.ListItems(I).SubItems(j) = "", Null, lstLista.ListItems(I).SubItems(j))
                Else
                    rsAux.Fields(lstLista.ColumnHeaders(j + 1).Text) = lstLista.ListItems(I).SubItems(j)
                End If
            End If
        Next
        rsAux.Update
    Next
    rsAux.MoveFirst
    Set GetRsNewDeListView = rsAux
End If
End Function

Private Function DeterminaFormato(pssformato As String, lncol As Long, Cols As Integer) As Long

Dim vPos As Integer
Dim pFormatos As String
Dim x As Long
Dim lFormatos() As String
Dim lnNroFormato As Long
pFormatos = pssformato
If Len(Trim(pFormatos)) > 0 Then
    For x = 0 To Cols - 1
        vPos = InStr(1, pFormatos, "-", vbTextCompare)
        ReDim Preserve lFormatos(x)
        If vPos > 0 Then
            lFormatos(x) = Mid(pFormatos, 1, IIf(vPos > 0, vPos - 1, Len(pFormatos)))
        Else
            If pFormatos <> "" Then
                lFormatos(x) = pFormatos
                pFormatos = ""
            End If
        End If
        If pFormatos <> "" Then
            pFormatos = Mid(pFormatos, IIf(vPos > 0, vPos + 1, Len(pFormatos)))
        End If
        If lncol = x Then
            lnNroFormato = Val(lFormatos(x))
            Exit For
        End If
    Next x
End If
DeterminaFormato = lnNroFormato
End Function

'Public Function fgTruncar(pnNumero As Double, pnDecimales As Integer) As String
'    Dim lnValor As Currency
'    Dim lnRes As Currency
'
'    lnValor = 10 ^ pnDecimales
'
'    'lnRes = Int(pnNumero) + Int((pnNumero - Int(pnNumero)) * lnValor) / lnValor
'
'    lnRes = CDbl(pnNumero) + CDbl((pnNumero - CDbl(pnNumero)) * lnValor) / lnValor
'
'    fgTruncar = lnRes
'End Function

'Public Function fgTruncar(pnNumero As Double, pnDecimales As Integer) As String
'    Dim intpos  As Integer
'    Dim lnDecimal As Double
'    Dim lsDec As String
'    Dim lnEntero As Long
'    Dim lnPos As Long
'
'    lnEntero = Int(pnNumero)
'    lnDecimal = Round(pnNumero - Int(lnEntero), 6)
'    lnPos = InStr(1, Trim(Str(lnDecimal)), ".")
'    If lnPos > 0 Then
'        lsDec = Mid(Trim(Str(lnDecimal)), lnPos + 1, 2)
'        lsDec = IIf(Len(lsDec) = 1, lsDec * 10, lsDec)
'        lnDecimal = Val(lsDec) / 100
'        fgTruncar = lnEntero + lnDecimal
'    Else
'        lnDecimal = 0
'        fgTruncar = lnEntero
'    End If
'End Function


Public Function fgTruncar(pnNumero As Double, pnDecimales As Integer) As String

    Dim I As Integer
    Dim sEnt As String
    Dim sDec As String
    Dim sNum As String
    Dim sPunto As String
    Dim sResul As String
    
    sNum = Trim(Str(pnNumero))
    sDec = ""
    sPunto = ""
    sEnt = ""
    For I = 1 To Len(Trim(sNum))
        If Mid(sNum, I, 1) = "." Then
            sPunto = "."
        Else
            If sPunto = "" Then
                sEnt = sEnt & Mid(sNum, I, 1)
            Else
                sDec = sDec & Mid(sNum, I, 1)
            End If
        End If
    Next I
    If sDec = "" Then
        sDec = "00"
    End If
    sResul = sEnt & "." & Left(sDec, 2)
    fgTruncar = sResul
    
End Function
'**DAOR 20070209
'**Función que permite transponer(Reemplazar caracteres)
Public Function CHRTRAN(psCadena As String, psChrBuscar As String, psChrReemplazo As String) As String
Dim I As Integer
Dim nLenB As Integer, nLenR As Integer, nLenC As Integer
Dim nPosiR As Integer
    nLenB = Len(psChrBuscar)
    nLenR = Len(psChrReemplazo)
    nLenC = Len(psCadena)
    If nLenC > 0 And nLenB > 0 Then
        For I = 1 To nLenB
            If I > nLenR Then
                psCadena = Replace(psCadena, Mid$(psChrBuscar, I, 1), "")
            Else
                psCadena = Replace(psCadena, Mid$(psChrBuscar, I, 1), Mid$(psChrReemplazo, I, 1))
            End If
        Next
    End If
    CHRTRAN = psCadena
End Function
'EJVG20111123 *******************************
Public Function FechaEsFinMes(ByVal pdFecha As Date) As Boolean
    Dim ldFechaFinMes As Date
    ldFechaFinMes = DateAdd("M", 1, pdFecha)
    ldFechaFinMes = Month(ldFechaFinMes) & "/01/" & Year(ldFechaFinMes)
    ldFechaFinMes = DateAdd("D", -1, ldFechaFinMes)
    If DateDiff("D", pdFecha, ldFechaFinMes) = 0 Then
        FechaEsFinMes = True
    Else
        FechaEsFinMes = False
    End If
End Function
'EJVG20111219 *********************************************************
Public Function EsEmailValido(ByVal psEmail As String) As Boolean
On Error GoTo ErrFunction
    Dim oReg As RegExp
    Set oReg = New RegExp
    ' Expresión regular
    oReg.Pattern = "^[\w-\.]{3,}@([\w-]{2,}\.)*([\w-]{2,}\.)[\w-]{2,4}$" '"^[\w-\.]+@\w+\.\w+$"
    ' Comprueba y Retorna True o false
    EsEmailValido = oReg.Test(psEmail)
    Set oReg = Nothing
Exit Function
ErrFunction:
    MsgBox Err.Description, vbCritical
    If Not oReg Is Nothing Then
        Set oReg = Nothing
    End If
End Function

'ALPA 20120402***********************************************************
Public Function TiposCambiosCierreMensual(pnAnio As Integer, pnMes As Integer, Optional pbMesCerrado As Boolean = True, Optional pnTipo As Integer = 0, Optional psOpeCod As String = "") As Currency
'*************************
    Dim sFecha  As Date
    Dim oCambio As nTipoCambio
    Dim oConect As DConecta
    Set oConect = New DConecta
    sFecha = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Format(pnMes, "00") & "/" & Trim(pnAnio))))
    Dim nTCFijo, nTCMerc As Currency
    Dim nTCVent As Currency
    Dim nTCComp As Currency
    Dim sSQL As String
    Dim rs   As New ADODB.Recordset
    Dim dFecha As Date
    Dim nImporteD As Currency
    Dim nImporteS As Currency
    'dFecha = CDate(GetFechaMov(psMovNro, True))
    dFecha = CDate(sFecha)
    
    Dim oTC As New nTipoCambio
    nTCFijo = oTC.EmiteTipoCambio(dFecha, TCFijoDia)
    
    If nTCMerc = 0 Then
            nTCVent = oTC.EmiteTipoCambio(dFecha, TCPondVenta)          'X
            nTCComp = oTC.EmiteTipoCambio(dFecha, TCPonderado)          'X
    Else
            nTCVent = nTCMerc
            nTCComp = nTCMerc
    End If
    oConect.CierraConexion
    Set oTC = Nothing
'*************************

    Set oCambio = New nTipoCambio
    If pnTipo = 1 Then
        TiposCambiosCierreMensual = Format(nTCFijo, "#,##0.0000")
    ElseIf pnTipo = 2 Then
        TiposCambiosCierreMensual = Format(nTCVent, "#,##0.0000")
    ElseIf pnTipo = 3 Then
        TiposCambiosCierreMensual = Format(nTCComp, "#,##0.0000")
    End If
    Set oCambio = Nothing
End Function

'************************************************************************
'EJVG20120814 ***
Public Function obtenerFechaFinMes(ByVal pnMes As Integer, ByVal pnAnio As Integer) As Date
    Dim dFecha  As Date
    dFecha = CDate("01/" & Format(pnMes, "00") & "/" & Format(pnAnio, "0000"))
    dFecha = DateAdd("m", 1, dFecha)
    dFecha = DateAdd("d", -1, dFecha)
    obtenerFechaFinMes = dFecha
End Function
'END EJVG *******
'EJVG20121113 ***
Public Function dameNombreMes(ByVal pnMes As Integer, Optional pbMayuscula As Boolean = False) As String
    dameNombreMes = Choose(pnMes, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    If pbMayuscula Then
        dameNombreMes = UCase(dameNombreMes)
    End If
End Function
'END EJVG *******
'ALPA20131218***************************************************************
Public Function LlenaCerosSUCAVESinVacios(ByVal sNumero As Double, Optional nTipoFormato As Integer = 1)
 
Dim sTempo As String
Dim I  As Integer

    If nTipoFormato = 1 Then
        'antes: sTempo = Format(sNumero, "###########0.00")
        
        sTempo = Format(sNumero, "############.00")
       
        For I = 1 To 16 - Len(sTempo)
            sTempo = " " & sTempo
        Next
    ElseIf nTipoFormato = 2 Then
        sTempo = Format(sNumero, "###########0")
        For I = 1 To 15 - Len(sTempo)
            sTempo = " " & sTempo
        Next
    ElseIf nTipoFormato = 3 Then
        'antes:  sTempo = Format(sNumero, "###0.00")
        
        sTempo = Format(sNumero, "####.00")
        
        For I = 1 To 7 - Len(sTempo)
            sTempo = " " & sTempo
        Next
    End If
    
     
    If nTipoFormato = 1 Then
        sTempo = Mid(sTempo, 1, 13) & Mid(sTempo, 15, 2)
    ElseIf nTipoFormato = 3 Then
        sTempo = Mid(sTempo, 1, 4) & Mid(sTempo, 6, 2)
    End If
    


LlenaCerosSUCAVESinVacios = Right(String(15, "0") & Trim(sTempo), 15)
End Function
'*************************************************************************************
'FRHU 20140315 RQ13659 - TI-ERS068-2013
Public Function NumerosEnterosSignosMasyMenos(intTecla As Integer, Optional pbNegativos As Boolean = False) As Integer
Dim cValidar As String
    If pbNegativos = False Then
        cValidar = "0123456789+-"
    Else
        cValidar = "0123456789-"
    End If
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) = 0 Then
            intTecla = 0
            Beep
        End If
    End If
    NumerosEnterosSignosMasyMenos = intTecla
End Function
'FIN FRHU 20140315
'PASIERS02420215
Public Function ArchivoEstaAbierto(ByVal Ruta As String) As Boolean
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
'END PASI
