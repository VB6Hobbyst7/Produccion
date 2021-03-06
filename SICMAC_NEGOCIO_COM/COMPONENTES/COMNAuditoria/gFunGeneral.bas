Attribute VB_Name = "gFunGeneral"
'M?dulo de datos de Contabilidad
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
        frmCentra.Icon = LoadPicture(App.Path & gsRutaIcono)
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
'validando a?o
If Val(Mid(cadfec, 7, 4)) < 1900 Or Val(Mid(cadfec, 7, 4)) > 9972 Then
    ValidaFecha = "A?o No Valido"
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
                    MsgBox "Formato de fecha no es v?lido", vbInformation, "Aviso"
                    lsControl.SetFocus
                    Exit Function
               Else
                    ValFecha = True
               End If
            Else
                ValFecha = False
                MsgBox "A?o de Fecha no es v?lido", vbInformation, "Aviso"
                lsControl.SetFocus
                lsControl.SelStart = 6
                lsControl.SelLength = 4
                Exit Function
            End If
        Else
            ValFecha = False
            MsgBox "Mes de Fecha no es v?lido", vbInformation, "Aviso"
            lsControl.SetFocus
            lsControl.SelStart = 3
            lsControl.SelLength = 2
            Exit Function
        End If
    Else
        ValFecha = False
        MsgBox "Dia de Fecha no es v?lido", vbInformation, "Aviso"
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

Public Function NumLet(ByVal strNum As String, Optional ByVal vLo)   '  , Optional ByVal vMoneda, Optional ByVal vCentimos) As String
    Dim I As Integer
    Dim Lo As Integer
    Dim iHayDecimal As Integer          'Posici?n del signo decimal
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
    Dim strN() As String
    
    Const cAncho = 12
    Const cGrupos = cAncho \ 3
    If unidad(1) <> "una" Then
        InicializarArrays
    End If
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
            'Con esto funcionar? bien el 100100, por ejemplo...
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
                strB = "un mill?n " & strB
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

Public Sub InicializarArrays()
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

Public Function ArmaFecha(dtmFechas As Date) As String
    Dim txtMeses As String
    txtMeses = Choose(Month(dtmFechas), "Enero", "Febrero", "Marzo", "Abril", _
                                        "Mayo", "Junio", "Julio", "Agosto", _
                                        "Setiembre", "Octubre", "Noviembre", "Diciembre")
                                        
    ArmaFecha = Day(dtmFechas) & " de " & txtMeses & " de " & Year(dtmFechas)

End Function

Public Function ValidaHora(lsControl As Control) As Boolean
   If Mid(lsControl, 1, 2) >= 0 And Mid(lsControl, 1, 2) <= 23 Then
        If Mid(lsControl, 4, 2) >= 0 And Mid(lsControl, 4, 2) <= 59 Then
            ValidaHora = True
        Else
            ValidaHora = False
            MsgBox "Minuto no es v?lido", vbInformation, "Aviso"
            lsControl.SetFocus
            lsControl.SelStart = 3
            lsControl.SelLength = 2
            Exit Function
        End If
    Else
        ValidaHora = False
        MsgBox "Hora no es v?lido", vbInformation, "Aviso"
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

'Public Function ConvNumLet(nNumero As Currency, Optional lSoloText As Boolean = True, Optional lSinMoneda As Boolean = False, Optional pnMoneda As Moneda = 0) As String
'Dim sCent As String
'Dim xValor As Single
'Dim vMoneda As String
'Dim cNumero As String
'cNumero = Format(nNumero, gsFormatoNumeroView)
'xValor = nNumero - Int(nNumero)
'If xValor = 0 Then
'   sCent = " Y 00/100 "
'Else
'   sCent = " Y " & Right(Trim(cNumero), 2) & "/100 "
'End If
'If pnMoneda <> 0 Then
'    vMoneda = IIf(pnMoneda = gMonedaNacional, "NUEVOS SOLES", "DOLARES AMERICANOS")
'End If
'If Not lSoloText Then
'   ConvNumLet = Trim(gsSimbolo) & " " & Trim(Format(nNumero, "###,###,##0.00#")) & " ("
'End If
'ConvNumLet = ConvNumLet & Trim(UCase(NumLet(CStr(nNumero), 0))) & sCent & " " & IIf(lSinMoneda, "", Trim(vMoneda)) & IIf(lSoloText, "", ")")
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

Public Function GeneraMovNroActualiza(pdFecha As Date, psCodUser As String, psCodCmac As String, psCodAge As String) As String
GeneraMovNroActualiza = Format(pdFecha & " " & GetHoraServer, gsFormatoMovFechaHora) & psCodCmac & Right(psCodAge, 2) & "00" & psCodUser
End Function

Public Function nVal(psImporte As String) As Currency
nVal = 0
If psImporte <> "" Then
   nVal = Format(psImporte, gsFormatoNumeroDato)
End If
End Function

Public Function ValidaAnio(nAnio As Integer) As Boolean
ValidaAnio = False
If nAnio > Year(gdFecSis) Then
   MsgBox "A?o no puede ser mayor a Periodo Actual", vbInformation, "Aviso"
   Exit Function
End If
If Year(gdFecSis) - nAnio > 5 Then
   MsgBox "El Sistema s?lo permite procesos hasta 5 a?os anteriores", vbInformation, "Aviso"
   Exit Function
End If
ValidaAnio = True
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
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    GetHoraServer = oConect.GetHoraServer()
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function GetFechaHoraServer() As String
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    GetFechaHoraServer = oConect.GetFechaHoraServer()
    oConect.CierraConexion
    Set oConect = Nothing
End Function

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

Public Function fgFechaHoraGrab(ByVal psMovNro As String) As String
    fgFechaHoraGrab = Mid(psMovNro, 1, 4) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 7, 2) & " " & Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2)
End Function

Public Function ReemplazaApostrofe(ByVal lsCadena As String) As String
    ReemplazaApostrofe = Replace(lsCadena, "'", "''", , , vbTextCompare)
End Function

Public Function CadDerecha(psCadena As String, lsTam As Integer) As String
    CadDerecha = Format(psCadena, "!" & String(lsTam, "@"))
End Function

Private Function DeterminaFormato(pssformato As String, lnCol As Long, Cols As Integer) As Long

Dim vPos As Integer
Dim pFormatos As String
Dim X As Long
Dim lFormatos() As String
Dim lnNroFormato As Long
pFormatos = pssformato
If Len(Trim(pFormatos)) > 0 Then
    For X = 0 To Cols - 1
        vPos = InStr(1, pFormatos, "-", vbTextCompare)
        ReDim Preserve lFormatos(X)
        If vPos > 0 Then
            lFormatos(X) = Mid(pFormatos, 1, IIf(vPos > 0, vPos - 1, Len(pFormatos)))
        Else
            If pFormatos <> "" Then
                lFormatos(X) = pFormatos
                pFormatos = ""
            End If
        End If
        If pFormatos <> "" Then
            pFormatos = Mid(pFormatos, IIf(vPos > 0, vPos + 1, Len(pFormatos)))
        End If
        If lnCol = X Then
            lnNroFormato = Val(lFormatos(X))
            Exit For
        End If
    Next X
End If
DeterminaFormato = lnNroFormato
End Function

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
