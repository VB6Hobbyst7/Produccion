Attribute VB_Name = "gFunGeneral"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A7EDE77033C"
'Módulo de datos de Contabilidad
Option Base 0
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

Dim unidad(0 To 9) As String
Dim decena(0 To 9) As String
Dim centena(0 To 10) As String
Dim deci(0 To 9) As String
Dim otros(0 To 15) As String


Dim cad As String
Dim Cadd As String


'FreeFile de impresión
'Global ArcSal As Integer



Public Function JDNum(pnCampos As String, pnLongit As Integer, _
                      pbComass As Boolean, pnDigEnt As Integer, _
                      pnDigDec As Integer)
                      

Dim Formato As String, i As Integer, lnPosDig As Integer
If pnCampos = "0.00" Then
   JDNum = Format(Trim(pnCampos), String(pnLongit, "@"))
   Exit Function
End If
If pbComass Then
   lnPosDig = 0
   For i = 1 To pnDigEnt
       lnPosDig = lnPosDig + 1
       Select Case lnPosDig
          Case 1
               Formato = "0" & Formato
          Case 4, 7, 10
               Formato = "#," & Formato
          Case Else
               Formato = "#" & Formato
       End Select
   Next i
   If pnDigDec > 0 Then
      Formato = Formato & "." & String(pnDigDec, "0")
   End If
Else
   For i = 1 To pnDigEnt
       Formato = IIf(i = 1, "0", "#") & Formato
   Next i
   If pnDigDec > 0 Then
      Formato = Formato & "." & String(pnDigDec, "0")
   End If
End If
pnCampos = Format(pnCampos, Formato)
JDNum = Format(Trim(pnCampos), String(pnLongit, "@"))
End Function

Public Sub CentraForm(frmCentra As Form)
    frmCentra.Move (Screen.Width - frmCentra.Width) / 2, (Screen.Height - frmCentra.Height) / 2, frmCentra.Width, frmCentra.Height
End Sub
'##ModelId=3A7EDEA302BF
Public Sub CargaCombo(ByRef Combo As ComboBox, rs As ADODB.Recordset)
Dim Campo As ADODB.Field
Dim lsDato As String
If rs Is Nothing Then Exit Sub
Combo.Clear
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


Public Function EliminaPunto(lnNumero As Currency) As Currency
Dim Pos As Long
Dim CadAux As String
Dim CadAux1 As String
Dim lsNumero As String
lsNumero = Trim(Str(lnNumero))
If Val(lsNumero) > 0 Then
    Pos = InStr(1, lsNumero, ".", vbTextCompare)
    If Pos > 0 Then
        CadAux = Mid(lsNumero, 1, Pos - 1)
        CadAux1 = Mid(lsNumero, Pos + 1, Len(Trim(lsNumero)))
        If Len(Trim(CadAux1)) = 1 Then
            CadAux1 = CadAux1 & "0"
        End If
        EliminaPunto = CCur(CadAux & CadAux1)
    Else
        EliminaPunto = lnNumero & "00"
    End If
Else
    EliminaPunto = lnNumero
End If
End Function
Public Function NumerosDecimales(cTexto As TextBox, intTecla As Integer, _
    Optional nLongitud As Integer = 8, Optional nDecimal As Integer = 2, _
    Optional pbNegativos As Boolean = False) As Integer
    Dim cValidar As String
    Dim cCadena As String
    cCadena = cTexto
    If pbNegativos Then
        cValidar = "-0123456789."
    Else
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
'Public Function ConvNumLet(nNumero As Currency, Optional lSoloText As Boolean = True, Optional lSinMoneda As Boolean = False) As String
'Dim sCent As String
'Dim xValor As Single
'Dim vMoneda As String
'Dim cNumero As String
'cNumero = Format(nNumero, gcFormView)
'xValor = nNumero - Int(nNumero)
'If xValor = 0 Then
'   sCent = " Y 00/100 "
'Else
'   sCent = " Y " & Right(Trim(cNumero), 2) & "/100 "
'End If
'vMoneda = IIf(gsSimbolo = gcMN, "NUEVOS SOLES", "DOLARES AMERICANOS")
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

Public Function GetHoraServer() As String
Dim Sql As String
Dim rsH As New ADODB.Recordset
Dim oConect As DConecta

Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
Sql = "Select Convert(varchar(10),getdate(),108) as sHora"
Set rsH = oConect.CargaRecordSet(Sql)
If Not rsH.EOF Then
   GetHoraServer = rsH!sHora
Else
   GetHoraServer = Format(Time, "hh:mm:ss")
End If
RSClose rsH

oConect.CierraConexion
Set oConect = Nothing

End Function

Public Function FechaHora(ByVal psFecha As Date) As String
    FechaHora = Format(psFecha & Space(1) & GetHoraServer, "mm/dd/yyyy hh:mm:ss")
End Function

Public Function FillNum(intNumero As String, intLenNum As Integer, ChrFil As String) As String
'On Error Resume Next
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

'Public Function nVal(psImporte As String) As Currency
'nVal = Format(psImporte, gsFormatoNumeroDato)
'End Function

'Public Sub ImprimeAsientoContable(ByVal psMovNro As String, Optional ByVal psDocVoucher As String = "", _
'                                  Optional ByVal psDocTpo As String = "", Optional ByVal psDocumento As String = "", _
'                                  Optional ByVal pbEfectivo As Boolean = False, _
'                                  Optional ByVal pbIngreso As Boolean = False, _
'                                  Optional ByVal psGlosa As String, Optional ByVal psPersCod As String, _
'                                  Optional ByVal pnImporte As Currency, Optional ByVal pnTipoArendir As ArendirTipo, _
'                                  Optional ByVal psNroRecViaticos As String = "", Optional pbHabEfectivo As Boolean = False, _
'                                  Optional ByVal pnNumCopiasAsiento As Integer = 2)
'Dim oContImp As NContImprimir
'Dim oNContFunc As NContFunciones
'Dim oPlant As dPlantilla
'Dim oNPlant As NPlantilla
'
'Set oContImp = New NContImprimir
'Set oNContFunc = New NContFunciones
'Set oPlant = New dPlantilla
'Set oNPlant = New NPlantilla
'
'Dim lsAsiento  As String
'Dim lsTitulo As String
'Dim lsVEOPSave As String
'Dim lsRecibo As String
'Dim lsOPSave As String
'Dim lsHab As String
'Dim lsPie As String
'Dim lsOtraFirma As String
'Dim I As Integer
'Dim lsCopias As String
'Dim lsCartas As String
'
'lsTitulo = ""
'If psDocVoucher <> "" Then
'    lsTitulo = " COMPROBANTE DE EGRESO N° " & psDocVoucher
'End If
'If pbEfectivo Then
'    lsRecibo = oContImp.ImprimeReciboIngresoEgreso(psMovNro, gdFecSis, psGlosa, _
'                                                   gcEmpresaLogo, gsOpeCod, psPersCod, _
'                                                   pnImporte, gnColPage, pnTipoArendir, psNroRecViaticos, pbIngreso)
'    If pbIngreso Then
'        lsTitulo = "I N G R E S O   D E   E F E C T I V O"
'    Else
'        lsTitulo = "S A L I D A   D E   E F E C T I V O"
'   End If
'End If
'lsPie = "179"
'If pbHabEfectivo Then
'    lsTitulo = "H A B I L I T A C I O N   D E   E F E C T I V O "
'    lsHab = oContImp.ImprimeDocSalidaEfectivo(gnColPage, gdFecSis, gsOpeCod, psMovNro)
'    lsPie = "158"
'    lsOtraFirma = "RESPONSABLE TRASLADO"
'End If
'lsAsiento = oContImp.ImprimeAsientoContable(psMovNro, gnLinPage, gnColPage, lsTitulo, , lsPie, lsOtraFirma)
'Select Case Val(psDocTpo)
'    Case gnDocTpoCheque
'        If psDocumento <> "" Then
'            lsAsiento = psDocumento & lsAsiento
'        End If
'        For I = 1 To pnNumCopiasAsiento - 1
'            lsCopias = lsCopias & Chr$(12) & lsAsiento
'        Next
'        lsAsiento = psDocumento & Chr$(12) + lsAsiento & lsCopias
'    Case gnDocTpoCarta
'        If psDocumento <> "" Then
'            frmCopiasImp.Show 1
'            For I = 1 To frmCopiasImp.CopiasCartas - 1
'                lsCartas = Chr$(12) + psDocumento
'            Next I
'            lsCartas = psDocumento + lsCartas
'            pnNumCopiasAsiento = frmCopiasImp.CopiasAsientos
'        End If
'        For I = 1 To pnNumCopiasAsiento - 1
'            lsCopias = lsCopias & Chr$(12) & lsAsiento
'        Next
'        lsAsiento = lsAsiento & lsCopias
'        lsAsiento = IIf(lsCartas = "", "", lsCartas & Chr$(12)) + lsAsiento
'        Set frmCopiasImp = Nothing
'    Case gnDocTpoOPago, TpoDocNotaCargo, TpoDocNotaAbono
'        If psDocumento <> "" Then
'            lsAsiento = psDocumento & lsAsiento
'        End If
'        For I = 1 To pnNumCopiasAsiento - 1
'            lsCopias = lsCopias & Chr$(12) & lsAsiento
'        Next
'        lsAsiento = lsAsiento & lsCopias
'    Case Else
'        If pbHabEfectivo Then
'            For I = 1 To pnNumCopiasAsiento - 1
'                lsCopias = lsCopias & Chr$(12) & lsAsiento
'            Next
'            lsAsiento = lsAsiento & lsCopias
'            If lsHab <> "" Then
'                lsAsiento = lsAsiento & Chr$(12) & lsHab
'            End If
'        Else
'            For I = 1 To pnNumCopiasAsiento - 1
'                lsCopias = lsCopias & Chr$(12) & lsAsiento
'            Next
'            lsAsiento = lsAsiento & lsCopias
'        End If
'        If lsRecibo <> "" Then
'            lsAsiento = lsAsiento & Chr$(12) & lsRecibo
'        End If
'End Select
'Dim oPrevio As clsPrevio
'Set oPrevio = New clsPrevio
'If psDocTpo = gnDocTpoOPago And pbIngreso = False Then
'    lsOPSave = oNPlant.GetPlantillaDoc(IDPlantillaOP)
'    lsOPSave = lsOPSave & IIf(lsOPSave = "", "", Chr$(12)) & psDocumento
'
'    oPlant.GrabaPlantilla "OPBatch", "Ordenes de Pago para impresiones en Batch", lsOPSave
'
'    lsVEOPSave = oNPlant.GetPlantillaDoc(IDPlantillaVOP)
'    lsVEOPSave = lsVEOPSave & IIf(lsVEOPSave = "", "", Chr$(12)) & lsAsiento
'    oPlant.GrabaPlantilla "OPVEBatch", "Voucher de egresos de Ordenes de Pago para impresiones en Batch", lsVEOPSave
'    If MsgBox(" ¿ Desea Imprimir Orden de Pago ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
'        lsOPSave = oNPlant.GetPlantillaDoc(IDPlantillaOP)
'        If ImprimeOrdenPago(lsOPSave) Then
'            lsVEOPSave = oNPlant.GetPlantillaDoc(IDPlantillaVOP)
'            oPrevio.Show lsVEOPSave, gsOpeDesc, False, gnLinPage
'            oPlant.GrabaPlantilla IDPlantillaOP, "Ordenes de Pago para impresiones en Batch", ""
'            oPlant.GrabaPlantilla IDPlantillaVOP, "Voucher de egresos de Ordenes de Pago para impresiones en Batch", ""
'      End If
'   End If
'Else
'   oPrevio.Show lsAsiento, gsOpeDesc, False, gnLinPage
'End If
'Set oPrevio = Nothing
'Set oContImp = Nothing
'Set oNContFunc = Nothing
'End Sub
Public Function GetTipCambio(pdFecha As Date) As Boolean
Dim oDGeneral As nTipoCambio
Set oDGeneral = New nTipoCambio
GetTipCambio = True
gnTipCambio = 0
gnTipCambioV = 0
gnTipCambioC = 0

 gnTipCambio = oDGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
 gnTipCambioV = oDGeneral.EmiteTipoCambio(pdFecha, TCVenta)
 gnTipCambioC = oDGeneral.EmiteTipoCambio(pdFecha, TCCompra)

If gnTipCambio = 0 Then
    MsgBox "Tipo de Cambio aun no definido", vbInformation, "Aviso"
    GetTipCambio = False
End If
End Function

Public Function GeneraMovNroActualiza(pdFecha As Date, psCodUser As String, psCodCmac As String, psCodAge As String) As String
GeneraMovNroActualiza = Format(pdFecha & " " & GetHoraServer, gsFormatoMovFechaHora) & psCodCmac & psCodAge & "00" & psCodUser
End Function

Public Sub EnviaPrevio(psImpre As String, psTitulo As String, ByVal pnLinPage As Integer, Optional plCondensado As Boolean = False)
Dim clsPrevio As New Previo.clsPrevio
clsPrevio.Show psImpre, psTitulo, plCondensado, pnLinPage
Set clsPrevio = Nothing
End Sub

Public Function nVal(psImporte As String) As Currency
nVal = Format(psImporte, gsFormatoNumeroDato)
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

Public Sub CambiaTamañoCombo(ByRef cboCombo As ComboBox, Optional nTamaño As Long = 200)
SendMessage cboCombo.hwnd, CB_SETDROPPEDWIDTH, nTamaño, 0
End Sub


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
If Not rs1 Is Nothing And rs1.State = adStateOpen Then
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

Public Sub RSLlenaCombo(prs As Recordset, psCombo As ComboBox)
If Not prs Is Nothing Then
   If Not prs.EOF Then
      psCombo.Clear
      Do While Not prs.EOF
         psCombo.AddItem Trim(prs(1)) & Space(100) & Trim(prs(0))
         prs.MoveNext
      Loop
   End If
End If
End Sub

'Convertir un Número a su denominacion en Letras
Public Function ConversNL(ByVal nmoneda As Moneda, ByVal nMonto As Double) As String
Dim Numero As String, sDecimal As String
Dim Letras As String
Dim sCent As String
Dim sMoneda As String
Dim xValor As Single
xValor = nMonto - Int(nMonto)
sDecimal = Right(Format$(nMonto, "#0.00"), 2)
If xValor = 0 Then
    sCent = " Y 00/100"
Else
    sCent = " Y " & sDecimal & "/100"
End If
Numero = CStr(nMonto)
sMoneda = IIf(nmoneda = gMonedaNacional, " NUEVOS SOLES", " DOLARES")
ConversNL = Trim(UCase(NumLet(Numero, 0))) & sCent & sMoneda
End Function

Public Function ConvNumLet(nNumero As Currency, Optional lSoloText As Boolean = True, Optional lSinMoneda As Boolean = False) As String
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
vMoneda = IIf(gsSimbolo = gcMN, "NUEVOS SOLES", "DOLARES AMERICANOS")
If Not lSoloText Then
   ConvNumLet = Trim(gsSimbolo) & " " & Trim(Format(nNumero, "###,###,##0.00#")) & " ("
End If
ConvNumLet = ConvNumLet & Trim(UCase(NumLet(CStr(nNumero), 0))) & sCent & " " & IIf(lSinMoneda, "", Trim(vMoneda)) & IIf(lSoloText, "", ")")
End Function


Public Sub ImprimeAsientoContable(ByVal psMovNro As String, Optional ByVal psDocVoucher As String = "", _
                                  Optional ByVal psDocTpo As String = "", Optional ByVal psDocumento As String = "", _
                                  Optional ByVal pbEfectivo As Boolean = False, _
                                  Optional ByVal pbIngreso As Boolean = False, _
                                  Optional ByVal psGlosa As String, Optional ByVal psPersCod As String, _
                                  Optional ByVal pnImporte As Currency, Optional ByVal pnTipoArendir As ArendirTipo, _
                                  Optional ByVal psNroRecViaticos As String = "", Optional pbHabEfectivo As Boolean = False, _
                                  Optional ByVal pnNumCopiasAsiento As Integer = 2)
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
Dim lsPie As String
Dim lsOtraFirma As String
Dim i As Integer
Dim lsCopias As String
Dim lsCartas As String

lsTitulo = ""
If psDocVoucher <> "" Then
    lsTitulo = " COMPROBANTE DE EGRESO N° " & psDocVoucher
End If
If pbEfectivo Then
    lsRecibo = oContImp.ImprimeReciboIngresoEgreso(psMovNro, gdFecSis, psGlosa, _
                                                   gsNomCmac, gsOpeCod, psPersCod, _
                                                   pnImporte, gnColPage, pnTipoArendir, psNroRecViaticos, pbIngreso)
    If pbIngreso Then
        lsTitulo = "I N G R E S O   D E   E F E C T I V O"
    Else
        lsTitulo = "S A L I D A   D E   E F E C T I V O"
   End If
End If
lsPie = "179"
If pbHabEfectivo Then
    lsTitulo = "H A B I L I T A C I O N   D E   E F E C T I V O "
    lsHab = oContImp.ImprimeDocSalidaEfectivo(gnColPage, gdFecSis, gsOpeCod, psMovNro, gsNomCmac)
    lsPie = "158"
    lsOtraFirma = "RESPONSABLE TRASLADO"
End If
'lsAsiento = oContImp.ImprimeAsientoContable(psMovNro, gnLinPage, gnColPage, lsTitulo, , lsPie, lsOtraFirma)
lsAsiento = oContImp.ImprimeAsientoContable(psMovNro, gnLinPage, gnColPage, lsTitulo, , lsPie)
Select Case Val(psDocTpo)
    Case TpoDocCheque  '  gnDocTpoCheque
        If psDocumento <> "" Then
            lsAsiento = psDocumento & lsAsiento
        End If
        For i = 1 To pnNumCopiasAsiento - 1
            lsCopias = lsCopias & Chr$(12) & lsAsiento
        Next
        lsAsiento = psDocumento & Chr$(12) + lsAsiento & lsCopias
    Case TpoDocCarta  ' gnDocTpoCarta
        If psDocumento <> "" Then
            frmCopiasImp.Show 1
            For i = 1 To frmCopiasImp.CopiasCartas - 1
                lsCartas = Chr$(12) + psDocumento
            Next i
            lsCartas = psDocumento + lsCartas
            pnNumCopiasAsiento = frmCopiasImp.CopiasAsientos
        End If
        For i = 1 To pnNumCopiasAsiento - 1
            lsCopias = lsCopias & Chr$(12) & lsAsiento
        Next
        lsAsiento = lsAsiento & lsCopias
        lsAsiento = IIf(lsCartas = "", "", lsCartas & Chr$(12)) + lsAsiento
        Set frmCopiasImp = Nothing
    Case TpoDocOrdenPago, TpoDocNotaCargo, TpoDocNotaAbono        'gnDocTpoOPago, TpoDocNotaCargo, TpoDocNotaAbono
        If psDocumento <> "" Then
            lsAsiento = psDocumento & lsAsiento
        End If
        For i = 1 To pnNumCopiasAsiento - 1
            lsCopias = lsCopias & Chr$(12) & lsAsiento
        Next
        lsAsiento = lsAsiento & lsCopias
    Case Else
        If pbHabEfectivo Then
            For i = 1 To pnNumCopiasAsiento - 1
                lsCopias = lsCopias & Chr$(12) & lsAsiento
            Next
            lsAsiento = lsAsiento & lsCopias
            If lsHab <> "" Then
                lsAsiento = lsAsiento & Chr$(12) & lsHab
            End If
        Else
            For i = 1 To pnNumCopiasAsiento - 1
                lsCopias = lsCopias & Chr$(12) & lsAsiento
            Next
            lsAsiento = lsAsiento & lsCopias
        End If
        If lsRecibo <> "" Then
            lsAsiento = lsAsiento & Chr$(12) & lsRecibo
        End If
End Select
Dim oPrevio As clsPrevio
Set oPrevio = New clsPrevio
If psDocTpo <> "" Then
    If psDocTpo = TpoDocOrdenPago And pbIngreso = False Then
        lsOPSave = oNPlant.GetPlantillaDoc(IDPlantillaOP)
        lsOPSave = lsOPSave & IIf(lsOPSave = "", "", Chr$(12)) & psDocumento
        
        oPlant.GrabaPlantilla "OPBatch", "Ordenes de Pago para impresiones en Batch", lsOPSave
        
        lsVEOPSave = oNPlant.GetPlantillaDoc(IDPlantillaVOP)
        lsVEOPSave = lsVEOPSave & IIf(lsVEOPSave = "", "", Chr$(12)) & lsAsiento
        oPlant.GrabaPlantilla "OPVEBatch", "Voucher de egresos de Ordenes de Pago para impresiones en Batch", lsVEOPSave
        If MsgBox(" ¿ Desea Imprimir Orden de Pago ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            lsOPSave = oNPlant.GetPlantillaDoc(IDPlantillaOP)
            If oContImp.ImprimeOrdenPago(lsOPSave) Then
                lsVEOPSave = oNPlant.GetPlantillaDoc(IDPlantillaVOP)
                oPrevio.Show lsVEOPSave, gsOpeDesc, False, gnLinPage
                oPlant.GrabaPlantilla IDPlantillaOP, "Ordenes de Pago para impresiones en Batch", ""
                oPlant.GrabaPlantilla IDPlantillaVOP, "Voucher de egresos de Ordenes de Pago para impresiones en Batch", ""
          End If
       End If
    Else
       oPrevio.Show lsAsiento, gsOpeDesc, False, gnLinPage
    End If
Else
    oPrevio.Show lsAsiento, gsOpeDesc, False, gnLinPage
End If
Set oPrevio = Nothing
Set oContImp = Nothing
Set oNContFunc = Nothing
End Sub

Public Function Encripta(pnTexto As String, Valor As Boolean) As String
'true = encripta
'false = desencripta
Dim MiClase As cEncrypt
Set MiClase = New cEncrypt
Encripta = MiClase.ConvertirClave(pnTexto, , Valor)
End Function

Public Function AdicionaRecordSet(ByRef prsDat As ADODB.Recordset, ByVal prs As ADODB.Recordset)
Dim nCol As Integer
Do While Not prs.EOF
    If Not prsDat Is Nothing Then
        If prsDat.State = adStateClosed Then
            For nCol = 0 To prs.Fields.Count - 1
                With prs.Fields(nCol)
                    prsDat.Fields.Append .name, .Type, .DefinedSize, .Attributes
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

Public Function RecordSetAdiciona(ByRef prsDat As ADODB.Recordset, ByVal prs As ADODB.Recordset)
Dim nCol As Integer
RecordSetDefineCampos prsDat, prs
Do While Not prs.EOF
    prsDat.AddNew
    For nCol = 0 To prs.Fields.Count - 1
        prsDat.Fields(nCol).value = prs.Fields(nCol).value
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
            prsDat.Fields.Append .name, .Type, .DefinedSize, .Attributes
        End With
    Next
    prsDat.Open
End If
End Function

Public Function ValidaConfiguracionRegional() As Boolean
Dim nmoneda As Currency
Dim nMonto As Double
Dim sNumero As String, sFecha As String
Dim nPosPunto As Integer, nPosComa As Integer

'Inicializamos las variables
ValidaConfiguracionRegional = True
nmoneda = 1234567
nMonto = 1234567
'Validamos Configuración de punto y Coma de Moneda
sNumero = Format$(nmoneda, "#,##0.00")
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

Public Sub UbicaCombo(pCombo As ComboBox, psDato As String, Optional pbBuscaFinal As Boolean = True, Optional pnNumComp As Integer = 7)
    Dim i As Integer
    Dim lbBan As Boolean
    lbBan = False
    
    If pbBuscaFinal Then
        For i = 0 To pCombo.ListCount - 1
            If Trim(Right(pCombo.List(i), pnNumComp)) = Trim(Right(psDato, pnNumComp)) Then
                lbBan = True
                pCombo.ListIndex = i
                i = pCombo.ListCount
            End If
        Next i
    Else
        For i = 0 To pCombo.ListCount - 1
            If Trim(Left(pCombo.List(i), pnNumComp)) = Trim(Left(psDato, pnNumComp)) Then
                lbBan = True
                pCombo.ListIndex = i
                i = pCombo.ListCount
            End If
        Next i
    End If
    
    If Not lbBan Then pCombo.ListIndex = -1
End Sub

Public Function ReemplazaApostrofe(ByVal lsCadena As String) As String
    ReemplazaApostrofe = Replace(lsCadena, "'", "''", , , vbTextCompare)
End Function
Public Function CadDerecha(psCadena As String, lsTam As Integer) As String
    CadDerecha = Format(psCadena, "!" & String(lsTam, "@"))
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

On Error GoTo Error
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
          If Right(UCase(fi.name), 3) = "EXE" Or Right(UCase(fi.name), 3) = "INI" Or Right(UCase(fi.name), 3) = "DLL" Then
             lsFecUltModifORIGEN = Format(fi.DateLastModified, "yyyy/mm/dd hh:mm:ss")
             If Dir(lsRutaSICMACT & fi.name) <> "" Then
                Set fd = fs.GetFile(lsRutaSICMACT & fi.name)
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
        frmHerActualizaSicmact.IniciaVariables True
        frmHerActualizaSicmact.Show 1
    End If
    Exit Function

Error:
    MsgBox "No se puede acceder a la ruta de origen, de la Ultima Actualizacion. - " & lsRutaUltActualiz, vbInformation, "Aviso"
    fgActualizaUltVersionEXE = False
End Function

Public Function fgFechaHoraGrab(ByVal psMovNro As String) As String
    fgFechaHoraGrab = Mid(psMovNro, 1, 4) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 7, 2) & " " & Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2)
End Function

Public Function fgFechaHoraPrend(ByVal psMovNro As String) As String
    fgFechaHoraPrend = Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4) & " " & Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2)
End Function

'********************************
' Adiciona Hoja a LibroExcel
'********************************
Public Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional pbActivaHoja As Boolean = True)
Dim lbExisteHoja As Boolean
Dim lbBorrarRangos As Boolean
'On Error Resume Next
lbExisteHoja = False
lbBorrarRangos = False
activaHoja:
For Each xlHoja1 In xlLibro.Worksheets
    If UCase(xlHoja1.name) = UCase(psHojName) Then
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
    xlHoja1.name = psHojName
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

'--*** Cadena de Conexion Servidor Consol -- Utiliza el mismo usuario y clave del Ini
Public Function fgCadenaConexConsol(ByVal psServer As String, ByVal psDataBase As String) As String

Dim lsPassword As String, lsUser As String
    Call fgObtenerDatosConexion(lsPassword, lsUser)
    
    'PROVIDER=SQLOLEDB;User ID=dbaccess;Password=cmacica;INITIAL CATALOG=dbcmacicamig;DATA SOURCE=01SRVSICMAC02
   fgCadenaConexConsol = "PROVIDER=SQLOLEDB;" & "User ID=" + Trim(lsUser) & ";Password=" & Trim(lsPassword) & ";INITIAL CATALOG=" & Trim(psDataBase) & ";DATA SOURCE=" & Trim(psServer)

End Function
' --*** Devuelve el password y usuario de la Cadena de Conexion.
Public Sub fgObtenerDatosConexion(psPassword As String, psUsuario As String)
Dim loConec As DConecta
Dim lsCadenaConexion As String
Dim lintPosI As Integer
Dim lintPosF As Integer
Set loConec = New DConecta
    loConec.AbreConexion
        lsCadenaConexion = loConec.CadenaConexion
    loConec.CierraConexion
    Set loConec = Nothing
    '*** Password
    lintPosI = InStr(1, lsCadenaConexion, "Password")
    lintPosF = InStr(lintPosI, lsCadenaConexion, ";")
    psPassword = Mid(lsCadenaConexion, lintPosI + Len("Password="), lintPosF - (lintPosI + Len("Password=")))
    '*** User ID
    lintPosI = InStr(1, lsCadenaConexion, "User ID=")
    lintPosF = InStr(lintPosI, lsCadenaConexion, ";")
    psUsuario = Mid(lsCadenaConexion, lintPosI + Len("User ID="), lintPosF - (lintPosI + Len("User ID=")))
End Sub

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

