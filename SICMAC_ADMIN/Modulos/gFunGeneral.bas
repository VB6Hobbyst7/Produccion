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


Global grCtasControl As New ADODB.Recordset
Global gnCostosControlNiv As Integer
'ALPA 20100322************************************************
Public Function GetAgencias(ByVal cCodAge As String) As String
        Dim oCon As DConecta
        Dim sSQL As String
        Dim rs As ADODB.Recordset
        Set oCon = New DConecta
        If oCon.AbreConexion = False Then Exit Function

        sSQL = "Select cAgeDescripcion From Agencias where cAgeCod='" & cCodAge & "'"
        Set rs = New ADODB.Recordset
        Set rs = oCon.CargaRecordSet(sSQL)
        If Not (rs.EOF Or rs.BOF) Then
            GetAgencias = rs!cAgeDescripcion & Space(20)
        Else
            GetAgencias = Space(20)
        End If
        
        'Set GetAgencias = oCon.CargaRecordSet(sSql)
        
        oCon.CierraConexion
        Set oCon = Nothing
End Function
'***************************************************************

Public Sub CentraForm(frmCentra As Form)
    frmCentra.Move (Screen.Width - frmCentra.Width) / 2, (Screen.Height - frmCentra.Height) / 2, frmCentra.Width, frmCentra.Height
End Sub
'##ModelId=3A7EDEA302BF
Public Function LlenaCombo(sSQL As String) As ComboBox
    On Error GoTo LlenaComboErr

    'your code goes here...

    Exit Function
LlenaComboErr:
    Call RaiseError(MyUnhandledError, "DFunGeneral:LlenaCombo Method")
End Function


Public Sub CambiaTamañoCombo(ByRef cboCombo As ComboBox, Optional nTamaño As Long = 200)
    SendMessage cboCombo.hwnd, CB_SETDROPPEDWIDTH, nTamaño, 0
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
Public Function valFecha(lsControl As Control) As Boolean
   If Mid(lsControl, 1, 2) > 0 And Mid(lsControl, 1, 2) <= 31 Then
        If Mid(lsControl, 4, 2) > 0 And Mid(lsControl, 4, 2) <= 12 Then
            If Mid(lsControl, 7, 4) >= 1900 And Mid(lsControl, 7, 4) <= 9999 Then
               If IsDate(lsControl) = False Then
                    valFecha = False
                    MsgBox "Formato de fecha no es válido", vbInformation, "Aviso"
                    lsControl.SetFocus
                    Exit Function
               Else
                    valFecha = True
               End If
            Else
                valFecha = False
                MsgBox "Año de Fecha no es válido", vbInformation, "Aviso"
                lsControl.SetFocus
                lsControl.SelStart = 6
                lsControl.SelLength = 4
                Exit Function
            End If
        Else
            valFecha = False
            MsgBox "Mes de Fecha no es válido", vbInformation, "Aviso"
            lsControl.SetFocus
            lsControl.SelStart = 3
            lsControl.SelLength = 2
            Exit Function
        End If
    Else
        valFecha = False
        MsgBox "Dia de Fecha no es válido", vbInformation, "Aviso"
        lsControl.SetFocus
        lsControl.SelStart = 0
        lsControl.SelLength = 2
        Exit Function
    End If
End Function
Public Function NumerosDecimales(cTexto As TextBox, intTecla As Integer, _
    Optional nLongitud As Integer = 8, Optional nDecimal As Integer = 2) As Integer
    Dim cValidar As String
    Dim cCadena As String
    cCadena = cTexto
    cValidar = "-0123456789."
    
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
'PASI20151126 ERS0782015
Public Function TextBox_SoloNumeros(ByVal KeyAscii As Integer) As Integer
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        TextBox_SoloNumeros = 0
    Else
        TextBox_SoloNumeros = KeyAscii
    End If
    If KeyAscii = 8 Then TextBox_SoloNumeros = KeyAscii
    If KeyAscii = 13 Then TextBox_SoloNumeros = KeyAscii
End Function
'end PASI**
'*******************************************************
'RUTINA VALIDA EL INGRESO DE UN NUMERO MAXIMO DE LINEAS
'*******************************************************
'FECHA CREACION : 24/06/99  -   MAVF
'MODIFICACION:
'**********************************************
Public Function intfLineas(cCadena As String, intTecla As Integer, intLinea As Integer) As Integer
Dim vLineas As Byte
Dim X As Byte
    If intTecla = 13 Then
        For X = 1 To Len(cCadena)
            If Mid(cCadena, X, 1) = Chr(13) Then
                vLineas = vLineas + 1
            End If
        Next X
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
'EJVG20120914 ***
Public Function LetrasNumeros(intTecla As Integer) As Integer
Dim cValidar  As String
Dim cNumero As String
    cValidar = "+:;-/\'´.¿<>?_=+[]{}|!@#$%^&()*,°"
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

Public Function SoloLetras2(intTecla As Integer, Optional lbMayusculas As Boolean = False) As Integer
Dim cValidar  As String
    cValidar = "0123456789+:;'<>?_=+[]{}|!@#$%^&()*./\ç¨-,´`¡¿Çºª""·"
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) <> 0 Then
            intTecla = 0
            Beep
        End If
    End If
    If lbMayusculas Then
        SoloLetras2 = Asc(UCase(Chr(intTecla)))
    Else
        SoloLetras2 = intTecla
    End If
End Function
'END EJVG *******
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
    Dim L As Integer
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
    L = Len(strNum)
    
    If lngA < 1 Then
        UnNumero = "cero"
        Exit Function
    End If
    '
    Una = True
    Millon = False
    Millones = False
    If L < 4 Then Una = False
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
        L = Len(strNum)
        K = Val(Right$(strNum, 2))
        If Right$(strNum, 1) = "0" Then
            K = K \ 10
            strD = decena(K)
        ElseIf K > 10 And K < 16 Then
            K = Val(Mid$(strNum, L - 1, 2))
            strD = otros(K)
        Else
            strU = unidad(Val(Right$(strNum, 1)))
            If L - 1 > 0 Then
                K = Val(Mid$(strNum, L - 1, 1))
                strD = deci(K)
            End If
        End If
        
        If L - 2 > 0 Then
            K = Val(Mid$(strNum, L - 2, 1))
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
'EJVG20111219 *******************************************
Public Function DigitoRPM(intTecla As Integer) As Boolean
    Dim cValidar As String
    cValidar = "#*"
    DigitoRPM = True
    If InStr(cValidar, Chr(intTecla)) = 0 Then
        DigitoRPM = False
        Beep
    End If
End Function
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
Dim X As Currency
X = Round(Dividendo / Divisor, 0)
Residuo = True
X = X * Divisor
If X <> Dividendo Then
   Residuo = False
End If
End Function
Public Function ConvNumLet(ByVal nNumero As Currency, Optional ByVal lSoloText As Boolean = True, Optional lSinMoneda As Boolean = False) As String
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
'''vMoneda = IIf(gsSimbolo = gcMN, "NUEVOS SOLES", "DOLARES AMERICANOS") 'marg ers044-2016
vMoneda = IIf(gsSimbolo = gcPEN_SIMBOLO, StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES AMERICANOS") 'marg ers044-2016
If Not lSoloText Then
   ConvNumLet = Trim(gsSimbolo) & " " & Trim(Format(nNumero, "###,###,##0.00#")) & " ("
End If
ConvNumLet = ConvNumLet & Trim(UCase(NumLet(CStr(nNumero), 0))) & sCent & " " & IIf(lSinMoneda, "", Trim(vMoneda)) & IIf(lSoloText, "", ")")
End Function

'Public Function ReadVarSis(txtCodPro As String, txtNomVar As String) As String
'    Dim RecVar As New ADODB.Recordset
'    Dim qryVar As String
'    On Error GoTo ERROR
'    Dim oConec As DConecta
'
'    Set oConec = New DConecta
'
'    qryVar = "SELECT cCodProd, cValorVar, cNomVar FROM VarSistema WHERE cCodProd = '" & txtCodPro & "' AND cNomVar = '" & txtNomVar & "'"
'    Set RecVar = oConec.CargaRecordSet(qryVar)
'    RecVar.Open qryVar, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If Not RecVar.EOF Then
'      ReadVarSis = Trim(RecVar!cValorVar)
'    Else
'      MsgBox "Variable " & txtNomVar & " no esta definido en el Sistema!. Por favor Consultar con Sistemas", vbInformation, "!Aviso!"
'    End If
'    RecVar.Close
'    Set RecVar = Nothing
'    oConec.CierraConexion
'    Set oConec = Nothing
'    Exit Function
'ERROR:
'    MsgBox "Error en Conexión + " + Err.Description, vbCritical, "Aviso"
'End Function

Public Sub CargaVarSis()
    Dim lsQrySis As String
    Dim rsQrySis As New ADODB.Recordset
    Dim oConect As DConecta
    Dim VSQL As String
    Dim lnStrConn As String
    Dim lnPosIni As Integer
    Dim lnPosFin As Integer
    Dim lnStr As String
    Set oConect = New DConecta
    
    If oConect.AbreConexion(gsConnection) = False Then
        Exit Sub
    End If
    
    lsQrySis = "SELECT cCodProd, cNomVar, cValorVar, cDescVar FROM VarSistema " _
             & "WHERE cCodProd in ('ADM','AHO') AND cNomVar IN ('dFecSis','cCodAge','cNomCMAC','cDirBackup','cCodCMAC') "
    Set rsQrySis = oConect.CargaRecordSet(lsQrySis)
    If rsQrySis.BOF Or rsQrySis.EOF Then
       rsQrySis.Close
       Set rsQrySis = Nothing
       MsgBox "Tabla VarSistema está vacia", vbInformation, "Aviso"
       gsCodAge = ""
       gsNomAge = ""
       gdFecSis = ""
       gsInstCmac = ""
       gsNomCmac = ""
       gsCodCMAC = ""
       Exit Sub
    End If
    Do While Not rsQrySis.EOF
        Select Case Trim(rsQrySis!cNomVar)
                Case "dFecSis"
                        gdFecSis = CDate(Trim(rsQrySis!cValorVar))
                Case "cCodAge"
                        gsCodAge = Trim(rsQrySis!cValorVar)
                        gsNomAge = Trim(rsQrySis!cDescVar)
                Case "cNomCMAC"
                        gsInstCmac = Trim(rsQrySis!cValorVar)
                        gsNomCmac = Trim(rsQrySis!cDescVar)
                Case "cCodCMAC"
                        gsCodCMAC = Trim(rsQrySis!cValorVar)
                Case "cDirBackup"
                        gsDirBackup = Trim(rsQrySis!cDescVar)
        End Select
        rsQrySis.MoveNext
    Loop
    rsQrySis.Close
    Set rsQrySis = Nothing
    
    'Deduce el nombre del Servidor
    
    gsServerName = oConect.servername
    'Deduce el nombre de la Base de Datos
    gsDBName = oConect.DatabaseName
    lnStrConn = oConect.CadenaConexion
    'Deduce el nombre de usuario
    lnPosIni = InStr(1, lnStrConn, "UID=", vbTextCompare)
    If lnPosIni > 0 Then
        lnPosFin = InStr(lnPosIni, lnStrConn, ";", vbTextCompare)
        lnStr = Mid(lnStrConn, lnPosIni, lnPosFin - lnPosIni)
        lnPosIni = InStr(1, lnStr, "=", vbTextCompare)
        gsUID = Mid(lnStr, lnPosIni + 1, Len(lnStr) - lnPosIni)
    Else
        gsUID = ""
    End If
    'Deduce el password
    lnPosIni = InStr(1, lnStrConn, "PWD=", vbTextCompare)
    If lnPosIni > 0 Then
        lnPosFin = InStr(lnPosIni, lnStrConn, ";", vbTextCompare)
        lnStr = Mid(lnStrConn, lnPosIni, lnPosFin - lnPosIni)
        lnPosIni = InStr(1, lnStr, "=", vbTextCompare)
        gsPWD = Mid(lnStr, lnPosIni + 1, Len(lnStr) - lnPosIni)
    Else
        gsPWD = ""
    End If
    oConect.CierraConexion
    Set oConect = Nothing
End Sub
Public Function CargaVarContab() As Boolean
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim cVal As String
Dim dFecInicio As Date, dCieCont As Date
Dim oConecta As DConecta

Exit Function

Set oConecta = New DConecta

oConecta.AbreConexion
sSQL = "select cNomVar, cValorVar from varsistema where cCodProd = 'CON' and cTipDat = 'D' or cTipDat = 'T' "
Set rs = oConecta.CargaRecordSet(sSQL)
If rs.EOF Then
   MsgBox "Datos de Variables del Sistema no encontrados. Por favor Consultar con Sistemas", vbCritical, "Error"
   Exit Function
End If
Do While Not rs.EOF
   cVal = Trim(rs!cValorVar)
   Select Case Trim(rs!cNomVar)
     Case "cIGV":           gcCtaIGV = cVal
     Case "cTpoFactura":    gcDocTpoFac = cVal
     Case "nTasaCajaCh":    gnTasaCajaCh = Val(Format(cVal, gcFormDato)) / 100
     Case "cTpoOrdenPago":  gcDocTpoOPago = cVal
     Case "cTpoCargo":      gcDocTpoCargo = cVal
     Case "cTpoCarta":      gcDocTpoCarta = cVal
     Case "cTpoAbono":      gcDocTpoAbono = cVal
     Case "cTpoCheque":     gcDocTpoCheque = cVal
     Case "cMonedaN":       gcMN = cVal
     Case "cMonedaE":       gcME = cVal
     Case "nMargSup":       gnMgSup = cVal
     Case "nMargIzq":       gnMgIzq = cVal
     Case "nMargDer":       gnMgDer = cVal
     Case "nLinPage":       gnLinPage = cVal
     Case "nTopeArendir":   gnArendirImporte = Val(cVal)
     Case "nLinPageOP":     gnLinPageOP = cVal
     Case "cConvMED":       gcConvMED = cVal
     Case "cConvMES":       gcConvMES = cVal
     Case "cConvTipo":      gcConvTipo = cVal
     Case "cCtaCaja":       gcCtaCaja = cVal
     Case "cCCHCta":        gcCCHCta = cVal
     Case "nEncajeExig":    gnEncajeExig = cVal
     Case "nTotalOblig":    gnTotalOblig = cVal
     Case "cCtaBancoMN":   gsCtaBancoMN = cVal
     Case "cCtaBancoME": gsCtaBancoME = cVal
     Case "cCtaBCRMN":   gsCtaBCRMN = cVal
     Case "cCtaBCRME":   gsCtaBCRME = cVal
     Case "cDirPlantillas": gsDirPlantillas = cVal
     'Case "cTpoRecEgreso": gsReciboEgreso = cVal
 End Select
   rs.MoveNext
Loop
glDiaCerrado = False

'gsFormatoFecha = gsFormatoFecha
sSQL = "SELECT cNomVar, cValorVar, cDescVar FROM varsistema WHERE cCodProd = 'ADM' and cNomVar IN ('dCieCont','dFecInicio','cEmpresa','cEmpresaRUC','cTitModulo','cFormatoFecha') "
Set rs = oConecta.CargaRecordSet(sSQL)
Do While Not rs.EOF
   Select Case Trim(rs!cNomVar)
   Case "cEmpresa":     gcEmpresa = rs!cDescVar
   Case "cEmpresaRUC":  gcEmpresaRUC = rs!cDescVar
   Case "dCieCont":     dCieCont = CDate(rs!cValorVar)
   Case "dFecInicio":   dFecInicio = CDate(rs!cValorVar)
   Case "cTitModulo":   gcTitModulo = Trim(rs!cValorVar) & " - " & rs!cDescVar
   Case "cFormatoFecha": 'FECHA XXX gcFormatoFecha = Trim(rs!cValorVar)
   Case "": gcEmpresaLogo = Trim(rs!cValorVar)
   End Select
   rs.MoveNext
Loop
rs.Close
Set rs = Nothing
oConecta.CierraConexion
Set oConecta = Nothing
CargaVarContab = True
End Function
Public Function CargaVarSistema1(lbContab As Boolean) As Boolean
On Error GoTo ErrorCarga
Dim oCon As NConstSistemas
Set oCon = New NConstSistemas
CargaVarSistema1 = True
CargaVarSis
GetTipCambio gdFecSis, Not IIf(oCon.LeeConstSistema(gConstSistBitCentral) = "1", True, False)
If lbContab Then
    CargaVarContab
End If
Exit Function

ErrorCarga:
    CargaVarSistema1 = False
    MsgBox "Error N° [" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
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

Public Function GetHoraServer() As String
Dim sql As String
Dim rsH As New ADODB.Recordset
Dim oConect As DConecta

Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
sql = "Select Convert(varchar(10),getdate(),108) as sHora"
Set rsH = oConect.CargaRecordSet(sql)
If Not rsH.EOF Then
   GetHoraServer = rsH!sHora
Else
   GetHoraServer = Format(Time, "hh:mm:ss")
End If
RSClose rsH

oConect.CierraConexion
Set oConect = Nothing

End Function
Public Function CaracteresFuncionales(intTecla As Integer) As Integer
Dim cValidar  As String
    cValidar = "0123456789+:;'<>?_=+[]{}|!@#$^&()*"
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) <> 0 Then
            intTecla = 0
            'Beep
        End If
    End If
    CaracteresFuncionales = intTecla
End Function
Public Function FechaHora(psFecha As Date) As String
    FechaHora = Format(psFecha & Space(1) & GetHoraServer, gsFormatoFechaHora)
End Function
Public Function FillNum(intNumero As String, intLenNum As Integer, ChrFil As String) As String
  FillNum = Left(String(intLenNum, ChrFil), (Len(String(intLenNum, ChrFil)) - Len(Trim(intNumero)))) + Trim(intNumero)
End Function
Public Sub RSClose(rs As ADODB.Recordset)
If rs.State = adStateOpen Then
    rs.Close
    Set rs = Nothing
End If
End Sub
Public Function RSVacio(rs1 As ADODB.Recordset) As Boolean
 RSVacio = (rs1.BOF And rs1.EOF)
End Function
Public Function TextErr(sMsg As String) As String
Dim nLen As Integer
nLen = InStr(1, sMsg, "*", vbTextCompare)
TextErr = Mid(sMsg, nLen + 1, Len(sMsg))
End Function
Public Function Encripta(psTexto As String, Optional psValor As Boolean = True) As String
'true = encripta
'false = desencripta
Dim oEnc As cEncrypt
Set oEnc = New cEncrypt
Encripta = oEnc.ConvertirClave(psTexto, , psValor)
End Function
Public Function PstaNombre(psNombre As String, Optional pbNombApell As Boolean = False) As String
Dim Total As Long
Dim Pos As Long
Dim CadAux As String
Dim lsApellido As String
Dim lsNombre As String
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
    lsNombre = Mid(CadAux, Pos + 1, Total)
    If pbNombApell = True Then
        If Len(Trim(lsConyugue)) > 0 Then
            PstaNombre = Trim(lsNombre) & " " & Trim(lsApellido) & " " & Trim(lsMaterno) & IIf(lbVda = False, " DE ", " ") & Trim(lsConyugue)
        Else
            PstaNombre = Trim(lsNombre) & " " & Trim(lsApellido) & " " & Trim(lsMaterno)
        End If
    Else
        If Len(Trim(lsConyugue)) > 0 Then
            PstaNombre = Trim(lsApellido) & " " & Trim(lsMaterno) & IIf(lbVda = False, " DE ", " ") & Trim(lsConyugue) & " " & Trim(lsNombre)
        Else
            PstaNombre = Trim(lsApellido) & " " & Trim(lsMaterno) & " " & Trim(lsNombre)
        End If
    End If
Else
    PstaNombre = Trim(psNombre)
End If
End Function

Public Function GetTipCambio(pdFecha As Date, Optional LeedeAdmin As Boolean = True) As Boolean
    Dim oDGeneral As DGeneral
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    
    If LeedeAdmin Then
        oCon.AbreConexion 'Remota "07", , , "01"
        oCon.Ejecutar "Set dateformat mdy"
        
        sql = " Select top 1 nValVent,nValComp,nValFijo from dbcomunes..tipcambio where dFecCamb Between '" & Format(pdFecha, gcFormatoFecha) & "' And '" & Format(pdFecha, gcFormatoFecha) & " 23:59:59' order by dFecCamb desc"
        Set rs = oCon.CargaRecordSet(sql)
        
        If rs.EOF And rs.BOF Then
            gnTipCambio = 0
            gnTipCambioV = 0
            gnTipCambioC = 0
            
            gnTipCambioVE = 0
            gnTipCambioCE = 0
            gnTipCambioPonderado = 0
            
            MsgBox "Tipo de cambio no definido.", vbInformation, "Aviso"
        Else
            gnTipCambio = rs!nValFijo
            gnTipCambioV = rs!nValVent
            gnTipCambioC = rs!nValComp
        End If
        
        rs.Close
    Else
    
        Set oDGeneral = New DGeneral
        GetTipCambio = True
        gnTipCambio = 0
        gnTipCambioV = 0
        gnTipCambioC = 0
        
        gnTipCambioVE = 0
        gnTipCambioCE = 0
        gnTipCambioPonderado = 0
        
        gnTipCambio = oDGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
        gnTipCambioV = oDGeneral.EmiteTipoCambio(pdFecha, TCVENTA)
        gnTipCambioC = oDGeneral.EmiteTipoCambio(pdFecha, TCCOMPRA)
        
        gnTipCambioVE = oDGeneral.EmiteTipoCambio(pdFecha, TCVentaEsp)
        gnTipCambioCE = oDGeneral.EmiteTipoCambio(pdFecha, TCCompraEsp)
        gnTipCambioPonderado = oDGeneral.EmiteTipoCambio(pdFecha, TCPonderado)
        
        If gnTipCambio = 0 Then
            MsgBox "Tipo de Cambio aun no definido", vbInformation, "Aviso"
            GetTipCambio = False
        End If
    End If
End Function

Public Function GeneraMovNroActualiza(pdFecha As Date, psCodUser As String, psCodCmac As String, psCodAge As String) As String
GeneraMovNroActualiza = Format(pdFecha & " " & GetHoraServer, gsFormatoMovFechaHora) & psCodCmac & psCodAge & "00" & psCodUser
End Function
Public Sub EnviaPrevio(psImpre As String, psTitulo As String, ByVal pnLinPage As Integer, Optional plCondensado As Boolean = False)
Dim clsPrev As New clsPrevio
clsPrev.Show psImpre, psTitulo, plCondensado, pnLinPage
Set clsPrev = Nothing
End Sub
Public Function nVal(psImporte As String) As Currency
    If IsNumeric(psImporte) Then
        nVal = Format(psImporte, gsFormatoNumeroDato)
    Else
        nVal = Format(0, gsFormatoNumeroDato)
    End If
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

Public Function ClaveCorrectaNT(ByVal psTxtUserName As String, ByVal psTxtPass As String, psTxtDominio As String) As Boolean
Dim oAcceso As UAcceso
    Set oAcceso = New UAcceso
    ClaveCorrectaNT = oAcceso.ClaveIncorrectaNT(psTxtUserName, psTxtPass, psTxtDominio)
    Set oAcceso = Nothing
End Function

'Convertir un Número a su denominacion en Letras
Public Function ConversNL(ByVal nMoneda As Moneda, ByVal nMonto As Double) As String
    Dim Numero As String
    Dim Letras As String
    Dim sCent As String
    Dim sMoneda As String
    Dim xValor As Single
    xValor = nMonto - Int(nMonto)
    If xValor = 0 Then
        sCent = " Y 00/100"
    Else
        sCent = " Y " & Right(Trim(Val(xValor)), 2) & "/100"
    End If
    Numero = CStr(nMonto)
    sMoneda = IIf(nMoneda = gMonedaNacional, " NUEVOS SOLES", " DOLARES")
    ConversNL = Trim(UCase(NumLet(Numero, 0))) & sCent & sMoneda
End Function

'FreeFile de impresión
'Global ArcSal As Integer



Public Function JDNum(pnCampos As String, pnLongit As Integer, _
                      pbComass As Boolean, pnDigEnt As Integer, _
                      pnDigDec As Integer)
                    
Dim formato As String, I As Integer, lnPosDig As Integer
If pnCampos = "0.00" Then
   JDNum = Format(Trim(pnCampos), String(pnLongit, "@"))
   Exit Function
End If
If pbComass Then
   lnPosDig = 0
   For I = 1 To pnDigEnt
       lnPosDig = lnPosDig + 1
       Select Case lnPosDig
          Case 1
               formato = "0" & formato
          Case 4, 7, 10
               formato = "#," & formato
          Case Else
               formato = "#" & formato
       End Select
   Next I
   If pnDigDec > 0 Then
      formato = formato & "." & String(pnDigDec, "0")
   End If
Else
   For I = 1 To pnDigEnt
       formato = IIf(I = 1, "0", "#") & formato
   Next I
   If pnDigDec > 0 Then
      formato = formato & "." & String(pnDigDec, "0")
   End If
End If
pnCampos = Format(pnCampos, formato)
JDNum = Format(Trim(pnCampos), String(pnLongit, "@"))
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

Public Function GetCtasControlCostoRS() As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sql As String
    
    sql = " Select b.cctacontcod from CostosCtaContControl a " _
        & " inner join ctacont b on b.cctacontcod like a.cctacontcod + '%'"
    
    oCon.AbreConexion
    
    Set GetCtasControlCostoRS = oCon.CargaRecordSet(sql)
End Function

Public Function GetCtasControlCostoBit(psCtaContCod As String) As Boolean
    grCtasControl.MoveFirst
    grCtasControl.Find "cCtaContCod='" & psCtaContCod & "'", , adSearchForward, 0
    
    If grCtasControl.EOF Then
        GetCtasControlCostoBit = False
    Else
        GetCtasControlCostoBit = True
    End If
End Function

Public Function GetOpeControlCostoTipo(psUsuario As String, psDominio As String) As CostosControlNiv
    Dim oAcceso As UAcceso
    Set oAcceso = New UAcceso
    
    If oAcceso.TienePermisoUsuario(psUsuario, OpeCostosControlOpeNivArea, psDominio) Then
        GetOpeControlCostoTipo = OpeCostosControlNivArea
    ElseIf oAcceso.TienePermisoUsuario(psUsuario, OpeCostosControlOpeNivAge, psDominio) Then
        GetOpeControlCostoTipo = OpeCostosControlNivAge
    ElseIf oAcceso.TienePermisoUsuario(psUsuario, OpeCostosControlOpeNivCmac, psDominio) Then
        GetOpeControlCostoTipo = OpeCostosControlNivCmac
    Else
        GetOpeControlCostoTipo = OpeCostosControlNivUsu
    End If
    
End Function

Public Sub ubicar_ano(codigo As String, combo As ComboBox)
Dim I As Integer
For I = 0 To combo.ListCount
If combo.List(I) = codigo Then
    combo.ListIndex = I
    Exit For
    End If
Next
End Sub

Public Function GetDescAgencia(psCodagencia) As String
Dim sql As String
Dim rsH As New ADODB.Recordset
Dim oConect As DConecta

Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
sql = " select   cAgeDescripcion from agencias where  cAgeCod ='" & psCodagencia & "'"
Set rsH = oConect.CargaRecordSet(sql)
If Not rsH.EOF Then
   GetDescAgencia = UCase(rsH!cAgeDescripcion)
Else
   GetDescAgencia = ""
End If
RSClose rsH

oConect.CierraConexion
Set oConect = Nothing

End Function

Public Function GetDescBancoProv(ByVal psPersCodProv As String, ByVal pnMoneda As Integer, Optional ByVal bDetrac As Boolean = False) As String
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim lsCodBanco As String
Dim oConect As DConecta
Set oConect = New DConecta

' Recuperamos codigo del Banco para obtener nombre
'*************************************************
If oConect.AbreConexion = False Then Exit Function
If pnMoneda = 1 Then
   sql = "Select isnull(cPERscodBancoMN,'')cPERscodBancoMN FROM Proveedor where cPersCod  ='" & psPersCodProv & "' "
   Set rs = oConect.CargaRecordSet(sql)
   If Not rs.EOF And Not rs.BOF Then
       lsCodBanco = rs!cPERscodBancoMN
   End If
   
Else
   sql = "Select isnull(cPERscodBancoME,'')cPERscodBancoME FROM Proveedor where cPersCod  ='" & psPersCodProv & "' "
   Set rs = oConect.CargaRecordSet(sql)
   If Not rs.EOF And Not rs.BOF Then
       lsCodBanco = rs!cPERscodBancoME
   End If
End If

'PASIERS0472015
If bDetrac Then
    sql = "Select isnull(cPersCodBancoDetracMN,'')cPersCodBancoDetracMN FROM Proveedor where cPersCod  ='" & psPersCodProv & "' "
   Set rs = oConect.CargaRecordSet(sql)
   If Not rs.EOF And Not rs.BOF Then
       lsCodBanco = rs!cPersCodBancoDetracMN
   End If
End If
'end PASI
'****************************************

sql = "Select cPersNombre from Persona where cPersCod  ='" & lsCodBanco & "'"
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetDescBancoProv = rs!cPersNombre
End If
rs.Close
Set rs = Nothing

End Function


Public Function GetCtaBancoProv(ByVal psPersCodProv As String, ByVal pnMoneda As Integer) As ADODB.Recordset
Dim sql As String
Dim rs As New ADODB.Recordset
Dim oConect As DConecta
Set oConect = New DConecta

If oConect.AbreConexion = False Then Exit Function

If pnMoneda = 1 Then
   sql = "Select isnull(cPersCodBancoMN,'')cPersCodBancoMN,isnull(cctaCodMN,'')cctaCodMN,isnull(cCtaCCIMN,'')cCtaCCIMN FROM Proveedor where cPersCod  ='" & psPersCodProv & "' " 'PASIERS0472015 agrego cCtaCCIMN
   Set rs = oConect.CargaRecordSet(sql)
ElseIf pnMoneda = 2 Then
   sql = "Select isnull(cPersCodBancoME,'')cPersCodBancoME,isnull(cCtaCodME,'')cCtaCodME,isnull(cCtaCCIME,'')cCtaCCIME FROM Proveedor where cPersCod  ='" & psPersCodProv & "' " 'PASIERS0472015 agrego cCtaCCIME
   Set rs = oConect.CargaRecordSet(sql)
Else 'PASIERS0472015 **********************
    sql = "Select isnull(cPersCodBancoDetracMN,'')cPersCodBancoDetracMN,isnull(cCtaCodDetracMN,'')cCtaCodDetracMN FROM Proveedor where cPersCod  ='" & psPersCodProv & "' "
   Set rs = oConect.CargaRecordSet(sql)
   'end PASI**************************
End If

Set GetCtaBancoProv = rs
oConect.CierraConexion
Set oConect = Nothing
End Function
'ALPA 20080903*********************************************
'Se copio el procedimiento del negocio
Public Sub GeneraReporteEnArchivoExcel(ByVal psNomCmac As String, ByVal psNomAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, ByVal psTitulo As String, ByVal psSubTitulo As String, _
                                    ByVal psNomArchivo As String, ByVal pMatCabeceras As Variant, ByVal prRegistros As ADODB.Recordset, _
                                    Optional pnNumDecimales As Integer, Optional Visible As Boolean = False, Optional psNomHoja As String = "", _
                                    Optional pbSinFormatDeReg As Boolean = False, _
                                    Optional pbUsarCabecerasDeRS As Boolean = False)
    Dim rs As ADODB.Recordset
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim liLineas As Integer, I As Integer
    Dim fs As Scripting.FileSystemObject
    Dim lnNumColumns As Integer


    If Not (prRegistros.EOF And prRegistros.BOF) Then
        If pbUsarCabecerasDeRS = True Then
            lnNumColumns = prRegistros.Fields.Count
        Else
            lnNumColumns = UBound(pMatCabeceras)
            lnNumColumns = IIf(prRegistros.Fields.Count < lnNumColumns, prRegistros.Fields.Count, prRegistros.Fields.Count)
        End If

        If psNomHoja = "" Then psNomHoja = psNomArchivo
        psNomArchivo = psNomArchivo & "_" & psCodUser & ".xls"

        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application
        If fs.FileExists(App.path & "\Spooler\" & psNomArchivo) Then
            fs.DeleteFile (App.path & "\Spooler\" & psNomArchivo)
        End If
        Set xlLibro = xlAplicacion.Workbooks.Add
        Set xlHoja1 = xlLibro.Worksheets.Add

        xlHoja1.Name = psNomHoja
        xlHoja1.Cells.Select
        'xlHoja1.Cells.NumberFormat = "@"

        'Cabeceras
        xlHoja1.Cells(1, 1) = psNomCmac
        xlHoja1.Cells(1, lnNumColumns) = Trim(Format(pdFecSis, "dd/mm/yyyy hh:mm:ss"))
        xlHoja1.Cells(2, 1) = psNomAge
        xlHoja1.Cells(2, lnNumColumns) = psCodUser
        xlHoja1.Cells(4, 1) = psTitulo
        xlHoja1.Cells(5, 1) = psSubTitulo
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, lnNumColumns)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(5, lnNumColumns)).HorizontalAlignment = xlCenter

        liLineas = 6
        If pbUsarCabecerasDeRS = True Then
            For I = 0 To prRegistros.Fields.Count - 1
                xlHoja1.Cells(liLineas, I + 1) = prRegistros.Fields(I).Name
            Next I
        Else
            For I = 0 To lnNumColumns - 1
                If (I + 1) > UBound(pMatCabeceras) Then
                    xlHoja1.Cells(liLineas, I + 1) = prRegistros.Fields(I).Name
                Else
                    xlHoja1.Cells(liLineas, I + 1) = pMatCabeceras(I, 0)
                End If
            Next I
        End If

        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).Cells.Interior.Color = RGB(220, 220, 220)
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).HorizontalAlignment = xlCenter

        If pbSinFormatDeReg = False Then
            liLineas = liLineas + 1
            While Not prRegistros.EOF
                For I = 0 To lnNumColumns - 1
                    If pMatCabeceras(I, 1) = "" Then  'Verificamos si tiene tipo
                        xlHoja1.Cells(liLineas, I + 1) = prRegistros(I)
                    Else
                        Select Case pMatCabeceras(I, 1)
                            Case "S"
                                xlHoja1.Cells(liLineas, I + 1) = prRegistros(I)
                            Case "N"
                                xlHoja1.Cells(liLineas, I + 1) = Format(prRegistros(I), "#0.00")
                            Case "D"
                                xlHoja1.Cells(liLineas, I + 1) = IIf(Format(prRegistros(I), "yyyymmdd") = "19000101", "", Format(prRegistros(I), "dd/mm/yyyy"))
                        End Select
                    End If
                Next I
                liLineas = liLineas + 1
                prRegistros.MoveNext
            Wend
        Else
            xlHoja1.Range("A7").CopyFromRecordset prRegistros 'Copia el contenido del recordset a excel
        End If

        xlHoja1.SaveAs App.path & "\Spooler\" & psNomArchivo
        MsgBox "Se ha generado el Archivo en " & App.path & "\Spooler\" & psNomArchivo

        If Visible Then
            xlAplicacion.Visible = True
            xlAplicacion.Windows(1).Visible = True
        'By Capi 19082008 se modifico para que se visualice correctamente
        Else

            xlLibro.Close
            xlAplicacion.Quit
        End If
            'xlLibro.Close
            'xlAplicacion.Quit
        '
        
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing

    End If
End Sub

