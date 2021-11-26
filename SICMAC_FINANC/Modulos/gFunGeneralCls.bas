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

Function ValidaFecha(cadfec As String) As String
Dim i As Integer
    If Len(cadfec) <> 10 Then
        ValidaFecha = "Fecha No Valida"
        Exit Function
    End If
    For i = 1 To 10
        If i = 3 Or i = 6 Then
            If Mid(cadfec, i, 1) <> "/" Then
                ValidaFecha = "Fecha No Valida"
                Exit Function
            End If
        Else
            If Asc(Mid(cadfec, i, 1)) < 48 Or Asc(Mid(cadfec, i, 1)) > 57 Then
                ValidaFecha = "Fecha No Valida"
                Exit Function
            End If
        End If
    Next i
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
If Val(Mid(cadfec, 7, 4)) < 1950 Or Val(Mid(cadfec, 7, 4)) > 9972 Then
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
Dim cValidar As String
    cValidar = "0123456789+:;'<>?_=+[]{}|!@#$%^&()*"
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) <> 0 Then
            intTecla = 0
            Beep
        End If
    End If
    SoloLetras = intTecla
End Function
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
    Dim i As Integer
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
        i = InStr(strNum, " ")
        If i = 0 Then Exit Do
        strNum = Left$(strNum, i - 1) & Mid$(strNum, i + 1)
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
    Dim k As Integer
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
    For k = Len(strQ) To 1 Step -3
        vez = vez + 1
        strN(vez) = Mid$(strQ, k - 2, 3)
    Next
    MaxVez = cGrupos
    For k = cGrupos To 1 Step -1
        If strN(k) = "000" Then
            MaxVez = MaxVez - 1
        Else
            Exit For
        End If
    Next
    For vez = 1 To MaxVez
        strU = "": strD = "": strC = ""
        strNum = strN(vez)
        l = Len(strNum)
        k = Val(Right$(strNum, 2))
        If Right$(strNum, 1) = "0" Then
            k = k \ 10
            strD = decena(k)
        ElseIf k > 10 And k < 16 Then
            k = Val(Mid$(strNum, l - 1, 2))
            strD = otros(k)
        Else
            strU = unidad(Val(Right$(strNum, 1)))
            If l - 1 > 0 Then
                k = Val(Mid$(strNum, l - 1, 1))
                strD = deci(k)
            End If
        End If
        
        If l - 2 > 0 Then
            k = Val(Mid$(strNum, l - 2, 1))
            'Con esto funcionará bien el 100100, por ejemplo...
            If k = 1 Then
                If Val(strNum) = 100 Then
                    k = 10
                End If
            End If
            strC = centena(k) & " "
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
Public Function ConvNumLet(psOpeCod As String, nNumero As Currency, Optional lSoloText As Boolean = True, Optional lSinMoneda As Boolean = False) As String
Dim sCent As String
Dim xValor As Single
Dim vMoneda As String
Dim cNumero As String
Dim lsSimbolo As String

lsSimbolo = IIf(Mid(psOpeCod, 3, 1) = gMonedaNacional, "S/.", "$")
cNumero = Format(nNumero, "#,#0.00")
xValor = nNumero - Int(nNumero)
If xValor = 0 Then
   sCent = " Y 00/100 "
Else
   sCent = " Y " & Right(Trim(cNumero), 2) & "/100 "
End If
vMoneda = IIf(Mid(psOpeCod, 3, 1) = gMonedaNacional, "NUEVOS SOLES", "DOLARES AMERICANOS")
If Not lSoloText Then
   ConvNumLet = Trim(lsSimbolo) & " " & Trim(Format(nNumero, "###,###,##0.00#")) & " ("
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
'Public Sub CargaVarSis()
'    Dim lsQrySis As String
'    Dim rsQrySis As New ADODB.Recordset
'    Dim oConect As DConecta
'    Dim VSQL As String
'    Dim lnStrConn As String
'    Dim lnPosIni As Integer
'    Dim lnPosFin As Integer
'    Dim lnStr As String
'    Set oConect = New DConecta
'
'    If oConect.AbreConexion(gsConnection) = False Then
'        Exit Sub
'    End If
'
'    lsQrySis = "SELECT cCodProd, cNomVar, cValorVar, cDescVar FROM VarSistema " _
'             & "WHERE cCodProd in ('ADM','AHO') AND cNomVar IN ('dFecSis','cCodAge','cNomCMAC','cDirBackup','cCodCMAC') "
'
'    Set rsQrySis = oConect.CargaRecordSet(lsQrySis)
'    If rsQrySis.BOF Or rsQrySis.EOF Then
'       rsQrySis.Close
'       Set rsQrySis = Nothing
'       MsgBox "Tabla VarSistema está vacia", vbInformation, "Aviso"
'       gsCodAge = ""
'       gsNomAge = ""
'       gdFecSis = ""
'       gsInstCmac = ""
'       gsNomCmac = ""
'       gsCodCMAC = ""
'       Exit Sub
'    End If
'    Do While Not rsQrySis.EOF
'        Select Case Trim(rsQrySis!cNomVar)
'                Case "dFecSis"
'                        gdFecSis = CDate(Trim(rsQrySis!cValorVar))
'                Case "cCodAge"
'                        gsCodAge = Trim(rsQrySis!cValorVar)
'                        gsNomAge = Trim(rsQrySis!cDescVar)
'                Case "cNomCMAC"
'                        gsInstCmac = Trim(rsQrySis!cValorVar)
'                        gsNomCmac = Trim(rsQrySis!cDescVar)
'                Case "cCodCMAC"
'                        gsCodCMAC = Trim(rsQrySis!cValorVar)
'                Case "cDirBackup"
'                        gsDirBackup = Trim(rsQrySis!cDescVar)
'        End Select
'        rsQrySis.MoveNext
'    Loop
'    rsQrySis.Close
'    Set rsQrySis = Nothing
'
'    'Deduce el nombre del Servidor
'
'    gsServerName = oConect.ServerName
'    'Deduce el nombre de la Base de Datos
'    gsDBName = oConect.DatabaseName
'    lnStrConn = oConect.CadenaConexion
'    'Deduce el nombre de usuario
'    lnPosIni = InStr(1, lnStrConn, "UID=", vbTextCompare)
'    If lnPosIni > 0 Then
'        lnPosFin = InStr(lnPosIni, lnStrConn, ";", vbTextCompare)
'        lnStr = Mid(lnStrConn, lnPosIni, lnPosFin - lnPosIni)
'        lnPosIni = InStr(1, lnStr, "=", vbTextCompare)
'        gsUID = Mid(lnStr, lnPosIni + 1, Len(lnStr) - lnPosIni)
'    Else
'        gsUID = ""
'    End If
'    'Deduce el password
'    lnPosIni = InStr(1, lnStrConn, "PWD=", vbTextCompare)
'    If lnPosIni > 0 Then
'        lnPosFin = InStr(lnPosIni, lnStrConn, ";", vbTextCompare)
'        lnStr = Mid(lnStrConn, lnPosIni, lnPosFin - lnPosIni)
'        lnPosIni = InStr(1, lnStr, "=", vbTextCompare)
'        gsPWD = Mid(lnStr, lnPosIni + 1, Len(lnStr) - lnPosIni)
'    Else
'        gsPWD = ""
'    End If
'    oConect.CierraConexion
'    Set oConect = Nothing
'End Sub
'Public Function CargaVarContab() As Boolean
'Dim sSQL As String
'Dim rs As New ADODB.Recordset
'Dim cVal As String
'Dim dFecInicio As Date, dCieCont As Date
'Dim oConecta As DConecta
'
'Set oConecta = New DConecta
'
'oConecta.AbreConexion
'sSQL = "select cNomVar, cValorVar from varsistema where cCodProd = 'CON' and cTipDat = 'D' or cTipDat = 'T' "
'Set rs = oConecta.CargaRecordSet(sSQL)
'If rs.EOF Then
'   MsgBox "Datos de Variables del Sistema no encontrados. Por favor Consultar con Sistemas", vbCritical, "Error"
'   Exit Function
'End If
'Do While Not rs.EOF
'   cVal = Trim(rs!cValorVar)
'   Select Case Trim(rs!cNomVar)
'     Case "cIGV":           gcCtaIGV = cVal
'     Case "cTpoFactura":    gnDocTpoFac = cVal
'     Case "nTasaCajaCh":    gnTasaCajaCh = Val(Format(cVal, gcFormDato)) / 100
'     Case "cTpoOrdenPago":  gnDocTpoOPago = cVal
'     Case "cTpoCargo":      gnDocTpoCargo = cVal
'     Case "cTpoCarta":      gnDocTpoCarta = cVal
'     Case "cTpoAbono":      gnDocTpoAbono = cVal
'     Case "cTpoCheque":     gnDocTpoCheque = cVal
'     Case "cMonedaN":       gcMN = cVal
'     Case "cMonedaE":       gcME = cVal
'     Case "nMargSup":       gnMgSup = cVal
'     Case "nMargIzq":       gnMgIzq = cVal
'     Case "nMargDer":       gnMgDer = cVal
'     Case "nLinPage":       gnLinPage = cVal
'     Case "nTopeArendir":   gnArendirImporte = Val(cVal)
'     Case "nLinPageOP":     gnLinPageOP = cVal
'     Case "cConvMED":       gcConvMED = cVal
'     Case "cConvMES":       gcConvMES = cVal
'     Case "cConvTipo":      gcConvTipo = cVal
'     Case "cCtaCaja":       gcCtaCaja = cVal
'     Case "cCCHCta":        gcCCHCta = cVal
'     Case "nEncajeExig":    gnEncajeExig = cVal
'     Case "nTotalOblig":    gnTotalOblig = cVal
'     Case "cCtaBancoMN":   gsCtaBancoMN = cVal
'     Case "cCtaBancoME": gsCtaBancoME = cVal
'     Case "cCtaBCRMN":   gsCtaBCRMN = cVal
'     Case "cCtaBCRME":   gsCtaBCRME = cVal
' End Select
'   rs.MoveNext
'Loop
'glDiaCerrado = False
'
'
'sSQL = "SELECT cNomVar, cValorVar, cDescVar FROM varsistema WHERE cCodProd = 'ADM' and cNomVar IN ('dCieCont','dFecInicio','cEmpresa','cEmpresaRUC','cTitModulo','cFormatoFecha') "
'Set rs = oConecta.CargaRecordSet(sSQL)
'Do While Not rs.EOF
'   Select Case Trim(rs!cNomVar)
'   Case "cEmpresa":     gcEmpresa = rs!cDescVar
'   Case "cEmpresaRUC":  gcEmpresaRUC = rs!cDescVar
'   Case "dCieCont":     dCieCont = CDate(rs!cValorVar)
'   Case "dFecInicio":   dFecInicio = CDate(rs!cValorVar)
'   Case "cTitModulo":   gcTitModulo = Trim(rs!cValorVar) & " - " & rs!cDescVar
'   Case "cFormatoFecha": gsFormatoFecha = Trim(rs!cValorVar)
'   Case "": gcEmpresaLogo = Trim(rs!cValorVar)
'   End Select
'   rs.MoveNext
'Loop
'rs.Close
'Set rs = Nothing
''If dCieCont = gdFecSis Then
''   MsgBox "Cierre del Día ya se realizó. Las Operaciones que haga no afectaran los Saldos !!! ", vbInformation, "Aviso"
''   glDiaCerrado = True
''End If
''If gdFecSis < dFecInicio Then
''   If MsgBox(" ¿ DESEA INICIAR DIA " & dFecInicio & " ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
''      CargaVarContab = False
''      Exit Function
''   End If
''   gdFecSis = dFecInicio
''   sSql = "UPDATE VarSistema SET cValorVar = '" & gdFecSis & "', cCodUsu = '" & gsCodUser & "', dUltMod = '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "' WHERE cNomVar = 'dFecSis'"
''   oConecta.Ejecutar sSql
''   glDiaCerrado = False
''End If
''If oConecta.AbreConexion(Right(gsCodAge, 2)) Then
''   sSql = "SELECT rtrim(cNomVar) as cNomVar, rtrim(cValorVar) as cValorVar FROM VarSistema WHERE cNomVar IN ('dFecSis','dFecCierre')"
''   If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
''   rs.CursorLocation = adUseClient
''   rs.Open sSql, dbCmactN, adOpenStatic, adLockReadOnly, adCmdText
''   Do While Not rs.EOF
''      If rs!cNomVar = "dFecSis" Then
''         If CDate(rs!cValorVar) > gdFecSis Then
''            gdFecSis = CDate(rs!cValorVar)
''         End If
''      End If
''      If rs!cNomVar = "dFecCierre" Then
''         If CDate(rs!cValorVar) >= gdFecSis Then
''            If MsgBox("Cierre de Agencias ya se realizó (" & rs!cValorVar & ") " & Chr(10) & "Es necesario que se realice el Cierre Diario del Módulo Administrativo. " & Chr(10) & "¿ Desea trabajar con fecha del Sistema " & gdFecSis & "? ", vbYesNo, "!Aviso!") = vbNo Then
''               gdFecSis = CDate(rs!cValorVar) + 1
''               sSql = "UPDATE VarSistema SET cValorVar = '" & gdFecSis & "', cCodUsu = '" & gsCodUser & "', dUltMod = '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "' WHERE cNomVar = 'dFecSis'"
''               dbCmact.Execute sSql
''               sSql = "UPDATE VarSistema SET cValorVar = '" & gdFecSis - 1 & "', cCodUsu = '" & gsCodUser & "', dUltMod = '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "' WHERE cNomVar = 'dFecCierre'"
''               dbCmact.Execute sSql
''            Else
''               Exit Do
''            End If
''         End If
''      End If
''      rs.MoveNext
''   Loop
''   CierraConeccion
''End If
''sSql = "SELECT rtrim(cValor) as cCtacod, nRanIniTab as nTipoBala, rtrim(cAbrev) as cTipoCta " _
''     & "FROM " & gcCentralCom & "TablaCod WHERE ccodtab LIKE 'C0__'"
''If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
''rs.Open sSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
''Do While Not rs.EOF
''   If rs!cTipoCta = "D" Then
''      gcCtaDeudora = gcCtaDeudora & "'" & rs!cCtaCod & "',"
''   Else
''      gcCtaAcreedora = gcCtaAcreedora & "'" & rs!cCtaCod & "',"
''   End If
''   If rs!nTipoBala = 1 Then
''      gcCtaBalaMN = gcCtaBalaMN & "'" & rs!cCtaCod & "',"
''   End If
''   rs.MoveNext
''Loop
''If gcCtaDeudora <> "" Then
''   gcCtaDeudora = Mid(gcCtaDeudora, 1, Len(gcCtaDeudora) - 1)
''End If
''If gcCtaAcreedora <> "" Then
''   gcCtaAcreedora = Mid(gcCtaAcreedora, 1, Len(gcCtaAcreedora) - 1)
''End If
''If gcCtaBalaMN <> "" Then
''   gcCtaBalaMN = Mid(gcCtaBalaMN, 1, Len(gcCtaBalaMN) - 1)
''End If
''rs.Close: Set rs = Nothing
'oConecta.CierraConexion
'Set oConecta = Nothing
'CargaVarContab = True
'End Function
'Public Function GetTipCambio(dFecha As Date) As Boolean
'Dim sSQL As String
'Dim rs As New ADODB.Recordset
'Dim oConect As DConecta
'
'Set oConect = New DConecta
'
'GetTipCambio = False
'gnTipCambio = 0
'gnTipCambioV = 0
'gnTipCambioC = 0
'
'oConect.AbreConexion
'sSQL = "SELECT * FROM " & gcCentralCom & "tipcambio WHERE dFecCamb BETWEEN '" & Format(dFecha, gsFormatoFecha) & "' and '" & Format(dFecha + 1, gsFormatoFecha) & "'"
'Set rs = oConect.CargaRecordSet(sSQL)
'If rs.EOF Then
'   sSQL = "SELECT nValFijoDia, nValComp, nValVent FROM " & gcCentralCom & "tipcambio WHERE datepart(mm,dfeccamb) = " & Mid(dFecha, 4, 2) _
'        & " and datepart(yy,dfeccamb) = " & Mid(dFecha, 7, 4)
'   Set rs = oConect.CargaRecordSet(sSQL)
'   If rs.EOF And rs.BOF Then
'      MsgBox "Tipo de Cambio aún no definido...!Por favor Ingréselo!", vbCritical, "Error"
'      Exit Function
'   End If
'End If
'rs.MoveLast
'gnTipCambio = Format(rs!nValFijoDia, "###,###,##0.000")
'gnTipCambioV = Format(rs!nValVent, "###,###,##0.000")
'gnTipCambioC = Format(rs!nValComp, "###,###,##0.000")
'GetTipCambio = True
'rs.Close: Set rs = Nothing
'oConect.CierraConexion
'Set oConect = Nothing
'End Function
'Public Sub CargaValoresCentral()
'Dim sSQL As String
'Dim R As New ADODB.Recordset
'Dim cCen As String
'Dim oConect As DConecta
'
'gcCentralPers = ""
'gcCentralImg = ""
'gcCentralCom = ""
'
'Set oConect = New DConecta
'oConect.AbreConexion
'
'    sSQL = "Select cValorVar,cNomVar From VarSistema Where cCodProd = 'ADM' And cNomVar IN ('cCentralPers','cCentralImg','cCentralCom')"
'    Set R = oConect.CargaRecordSet(sSQL)
'    Do While Not R.EOF
'        Select Case Trim(R!cNomVar)
'            Case "cCentralPers"
'                gcCentralPers = Trim(R!cValorVar)
'            Case "cCentralImg"
'                gcCentralImg = Trim(R!cValorVar)
'            Case "cCentralCom"
'                gcCentralCom = Trim(R!cValorVar)
'        End Select
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'    oConect.CierraConexion
'    Set oConect = Nothing
'End Sub
'Public Function CargaVarSistema(lbContab As Boolean) As Boolean
'On Error GoTo ErrorCarga
'CargaVarSistema = True
'
'CargaVarSis
'CargaValoresCentral
''GetTipCambio gdFecSis
'If lbContab Then
'    CargaVarContab
'End If
'Exit Function
'
'ErrorCarga:
'    CargaVarSistema = False
'    MsgBox "Error N° [" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
'End Function
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
Public Function FechaHora(psFecha As Date) As String
    FechaHora = Format(psFecha & Space(1) & GetHoraServer, "mm/dd/yyyy hh:mm:ss")
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

Public Function GetFechaMov(cMovNro, lDia As Boolean) As String
Dim lFec As Date
lFec = Mid(cMovNro, 7, 2) & "/" & Mid(cMovNro, 5, 2) & "/" & Mid(cMovNro, 1, 4)
If lDia Then
   GetFechaMov = Format(lFec, gsFormatoFechaView)
Else
   GetFechaMov = Format(lFec, gsFormatoFecha)
End If
End Function
