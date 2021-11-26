Attribute VB_Name = "gFunGeneral"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A7EDE77033C"
'Módulo de datos de Contabilidad
Option Base 0
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
'By capi 28102008
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
'

Dim unidad(0 To 9) As String
Dim decena(0 To 9) As String
Dim centena(0 To 10) As String
Dim deci(0 To 9) As String
Dim otros(0 To 15) As String
'add jhcu encuesta 09-01-2020
Dim cUsuVisto As String
Dim res As Integer
'fin jhcu
Dim cad As String
Dim Cadd As String
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018


'FreeFile de impresión*
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
Public Sub CargaCombo(ByRef combo As ComboBox, rs As ADODB.Recordset)
Dim Campo As ADODB.Field
Dim lsDato As String
If rs Is Nothing Then Exit Sub
combo.Clear
Do While Not rs.EOF
    lsDato = ""
    For Each Campo In rs.Fields
        lsDato = lsDato & Campo.value & space(50)
    Next
    lsDato = Mid(lsDato, 1, Len(lsDato) - 50)
    combo.AddItem lsDato
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub

'WIOR 20120705 SEGUN OYP-RFC060-2012 ***************************************************
Public Function SoloLetras2(intTecla As Integer, _
                           Optional lbMayusculas As Boolean = False) As Integer
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
'WIOR FIN********************************************************************************

Public Function SoloLetras3(intTecla As Integer, Optional lbMayusculas As Boolean = False) As Integer 'LUCV20160913, Según ERS004-2016
Dim cValidar  As String
    cValidar = "+'<>?_=+[]{}|!@#$%^&()*/\ç¨-´`¡¿Çºª""·"
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) <> 0 Then
            intTecla = 0
            Beep
        End If
    End If
    If lbMayusculas Then
        SoloLetras3 = Asc(UCase(Chr(intTecla)))
    Else
        SoloLetras3 = intTecla
    End If
End Function
Public Function SoloLetras4(intTecla As Integer, Optional lbMayusculas As Boolean = False) As Integer 'LUCV20170304, ANEXO 001-2017
Dim cValidar  As String
    cValidar = "'<>?_|!@#$%&\ç¨´`¡¿Çºª""·"
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) <> 0 Then
            intTecla = 0
            Beep
        End If
    End If
    If lbMayusculas Then
        SoloLetras4 = Asc(UCase(Chr(intTecla)))
    Else
        SoloLetras4 = intTecla
    End If
End Function
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
Dim pos As Long
Dim CadAux As String
Dim CadAux1 As String
Dim lsNumero As String
lsNumero = Trim(Str(lnNumero))
If Val(lsNumero) > 0 Then
    pos = InStr(1, lsNumero, ".", vbTextCompare)
    If pos > 0 Then
        CadAux = Mid(lsNumero, 1, pos - 1)
        CadAux1 = Mid(lsNumero, pos + 1, Len(Trim(lsNumero)))
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

Public Function SoloLetras(intTecla As Integer, _
                           Optional lbMayusculas As Boolean = False) As Integer
Dim cValidar  As String
    cValidar = "0123456789+:;'<>?_=+[]{}|!@#$%^&()*"
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) <> 0 Then
            intTecla = 0
            Beep
        End If
    End If
    If lbMayusculas Then
        SoloLetras = Asc(UCase(Chr(intTecla)))
    Else
        SoloLetras = intTecla
    End If
End Function
Public Sub RecuperaTimeOutPinPadAG()
Dim sql As String
Dim rsTO As New ADODB.Recordset
Dim oConect As DConecta

    Set oConect = New DConecta
    oConect.AbreConexion
    
    sql = "Exec ATM_RecuperaTimeOutPinPad '" & gsCodAge & "'"
    Set rsTO = oConect.CargaRecordSet(sql)
    If Not rsTO.EOF Then
       gnTimeOutAg = rsTO!nTimeOutAGE
       gnCodOpeTarj = rsTO!nRetOblTarjeta
    Else
       gnTimeOutAg = 700
       gnCodOpeTarj = 0
    End If
    RSClose rsTO
    
    oConect.CierraConexion
    Set oConect = Nothing

End Sub

Public Function ValidaDevTarjetas() As Boolean
Dim sql As String
Dim rs As New ADODB.Recordset
Dim oConect As DConecta
Dim lnValDev As Integer
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
    
    Set oConect = New DConecta
    oConect.AbreConexion
    'Dim loConec As New DConecta

    Set Cmd = New ADODB.Command
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCodAge", adVarChar, adParamInput, 2, gsCodAge)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pdFecha", adDBDate, adParamInput, , Format(gdFecSis, "YYYY/MM/DD"))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psUser", adVarChar, adParamInput, 4, gsCodUser)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnVal", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
        
    oConect.AbreConexion
    Cmd.ActiveConnection = oConect.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "stp_sel_UserNoDevolvieronTarjeta"
    Cmd.Execute
    
    'Sql = "Exec stp_sel_UserNoDevolvieronTarjeta '" & gsCodAge & "','" & Format(gdFecSis, "YYYY/MM/DD") & "','" & gsCodUser & "', @aaa Output"
    'Set rs = oConect.CargaRecordSet(Sql)
    
    If Cmd.Parameters(3).value = 0 Then
       MsgBox "Tienes Tarjetas para devolver al supervisor", vbInformation, "MENSAJE DEL SISTEMA"
       ValidaDevTarjetas = False
    Else
       ValidaDevTarjetas = True
    End If
    
    'RSClose rs
    Set Prm = Nothing
    Set Cmd = Nothing
    
    oConect.CierraConexion
    Set oConect = Nothing

End Function

Public Function ValidaConfDevTarjetas() As Boolean
Dim sql As String
Dim rs As New ADODB.Recordset
Dim oConect As DConecta
Dim lsUser As String


    Set oConect = New DConecta
    oConect.AbreConexion
    
    sql = "Exec stp_sel_UserNoConfDevolucionTarjeta '" & gsCodAge & "','" & Format(gdFecSis, "YYYY/MM/DD") & "'"
    Set rs = oConect.CargaRecordSet(sql)
    
    If Not rs.EOF Then
       lsUser = rs!cUserConfDev
       MsgBox lsUser & ": Tienes pendiente la confirmacion de la devolucion Tarjetas", vbInformation, "MENSAJE DEL SISTEMA"
       ValidaConfDevTarjetas = False
    Else
       ValidaConfDevTarjetas = True
    End If
    
    RSClose rs
    
    oConect.CierraConexion
    Set oConect = Nothing

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
Public Function EsEmailValido(ByVal psEmail As String) As Boolean
On Error GoTo ErrFunction
    Dim oreg As RegExp
    Set oreg = New RegExp
    ' Expresión regular
    'oreg.Pattern = "^[\w-\.]{3,}@([\w-]{2,}\.)*([\w-]{2,}\.)[\w-]{2,4}$" '"^[\w-\.]+@\w+\.\w+$"
    oreg.Pattern = "^[ñ\w-\.]{3,}@([\w-]{2,}\.)*([\w-]{2,}\.)[\w-]{2,4}$" 'JOEP20201102
    ' Comprueba y Retorna True o false
    EsEmailValido = oreg.Test(psEmail)
    Set oreg = Nothing
Exit Function
ErrFunction:
    MsgBox err.Description, vbCritical
    If Not oreg Is Nothing Then
        Set oreg = Nothing
    End If
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
        sNumero = space$(Lo)
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
Dim pos As Variant
On Error GoTo Errbusq
   BuscaDato = False
   pos = rsAdo.Bookmark
   rsAdo.Find Criterio, IIf(start = 1, 0, start + 1), adSearchForward, 1
   If rsAdo.EOF Then
      rsAdo.Bookmark = pos
      If lMsg Then
         MsgBox " ! Dato no encontrado... ! ", vbExclamation, "Error de Busqueda"
         BuscaDato = False
      End If
   Else
      BuscaDato = True
   End If
Exit Function
Errbusq:
   MsgBox TextErr(err.Description), vbInformation, "Aviso"
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

Public Function FechaHora(ByVal psFecha As Date) As String
    FechaHora = Format(psFecha & space(1) & GetHoraServer, "mm/dd/yyyy hh:mm:ss")
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
Dim pos As Long
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
pos = InStr(psNombre, "/")
If pos <> 0 Then
    lsApellido = Left(psNombre, pos - 1)
    CadAux = Mid(psNombre, pos + 1, Total)
    pos = InStr(CadAux, "\")
    If pos <> 0 Then
        lsMaterno = Left(CadAux, pos - 1)
        CadAux = Mid(CadAux, pos + 1, Total)
        pos = InStr(CadAux, ",")
        If pos > 0 Then
            CadAux2 = Left(CadAux, pos - 1)
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
        CadAux = Mid(CadAux, pos + 1, Total)
        pos = InStr(CadAux, ",")
        If pos <> 0 Then
            lsMaterno = Left(CadAux, pos - 1)
            lsConyugue = ""
        Else
            lsMaterno = CadAux
        End If
    End If
    lsNombre = Mid(CadAux, pos + 1, Total)
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
Public Function GetTipCambio(pdFecha As Date) As Boolean
Dim oDGeneral As COMDConstSistema.NCOMTipoCambio
Set oDGeneral = New COMDConstSistema.NCOMTipoCambio
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

Public Function GeneraMovNroActualiza(pdFecha As Date, psCodUser As String, psCodCMAC As String, psCodAge As String) As String
GeneraMovNroActualiza = Format(pdFecha & " " & GetHoraServer, gsFormatoMovFechaHora) & psCodCMAC & psCodAge & "00" & psCodUser
End Function

Public Sub EnviaPrevio(psImpre As String, psTitulo As String, ByVal pnLinPage As Integer, Optional plCondensado As Boolean = False)
Dim clsprevio As New previo.clsprevio
clsprevio.Show psImpre, psTitulo, plCondensado, pnLinPage
Set clsprevio = Nothing
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
    Case objPersona
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

Public Sub RSLlenaCombo(pRs As Recordset, psCombo As ComboBox)
If Not pRs Is Nothing Then
   If Not pRs.EOF Then
      psCombo.Clear
      Do While Not pRs.EOF
         psCombo.AddItem Trim(pRs(1)) & space(100) & Trim(pRs(0))
         pRs.MoveNext
      Loop
   End If
End If
End Sub

'Convertir un Número a su denominacion en Letras
Public Function ConversNL(ByVal nMoneda As Moneda, ByVal nMonto As Double) As String
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
sMoneda = IIf(nMoneda = gMonedaNacional, " SOLES", " DOLARES")
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
vMoneda = IIf(gsSimbolo = gcMN, "SOLES", "DOLARES AMERICANOS")
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
Dim oPrevio As clsprevio
Set oPrevio = New clsprevio
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

'JUEZ 20150310 ********************************
Public Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
    'permite que solo sean ingresados los numeros, el ENTER y el RETROCESO
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        SoloNumeros = 0
    Else
        SoloNumeros = KeyAscii
    End If
    ' teclas especiales permitidas
    If KeyAscii = 8 Then SoloNumeros = KeyAscii ' borrado atras
    If KeyAscii = 13 Then SoloNumeros = KeyAscii 'Enter
End Function
'END JUEZ *************************************

'ARCV 25-04-2007 : Comenta para generar el Proyecto Clases
Public Function Encripta(pnTexto As String, valor As Boolean) As String
'true = encripta
'false = desencripta
Dim MiClase As cEncrypt
Set MiClase = New cEncrypt
Encripta = MiClase.ConvertirClave(pnTexto, , valor)
End Function

Public Function AdicionaRecordSet(ByRef prsDat As ADODB.Recordset, ByVal pRs As ADODB.Recordset)
Dim nCol As Integer
Do While Not pRs.EOF
    If Not prsDat Is Nothing Then
        If prsDat.State = adStateClosed Then
            For nCol = 0 To pRs.Fields.count - 1
                With pRs.Fields(nCol)
                    prsDat.Fields.Append .Name, .Type, .DefinedSize, .Attributes
                End With
            Next
            prsDat.Open
        End If
        prsDat.AddNew
        For nCol = 0 To pRs.Fields.count - 1
            prsDat.Fields(nCol).value = pRs.Fields(nCol).value
        Next
        prsDat.Update
    End If
    pRs.MoveNext
Loop
If Not prsDat Is Nothing Then
    If prsDat.RecordCount > 0 Then
        prsDat.MoveFirst
    End If
End If
End Function

Public Function RecordSetAdiciona(ByRef prsDat As ADODB.Recordset, ByVal pRs As ADODB.Recordset)
Dim nCol As Integer
RecordSetDefineCampos prsDat, pRs
Do While Not pRs.EOF
    prsDat.AddNew
    For nCol = 0 To pRs.Fields.count - 1
        prsDat.Fields(nCol).value = pRs.Fields(nCol).value
    Next
    prsDat.Update
    pRs.MoveNext
Loop
End Function


Public Function RecordSetDefineCampos(ByRef prsDat As ADODB.Recordset, ByVal pRs As ADODB.Recordset)
Dim nCol As Integer

If prsDat Is Nothing Then
    Set prsDat = New ADODB.Recordset
End If
If prsDat.State = adStateClosed Then
    For nCol = 0 To pRs.Fields.count - 1
        With pRs.Fields(nCol)
            prsDat.Fields.Append .Name, .Type, .DefinedSize, .Attributes
        End With
    Next
    prsDat.Open
End If
End Function

Public Function ValidaConfiguracionRegional() As Boolean
Dim nMoneda As Currency
Dim nMonto As Double
Dim sNumero As String, sFecha As String
Dim nPosPunto As Integer, nPosComa As Integer

'Inicializamos las variables
ValidaConfiguracionRegional = True
nMoneda = 1234567
nMonto = 1234567
'Validamos Configuración de punto y Coma de Moneda
sNumero = Format$(nMoneda, "#,##0.00")
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
If Date <> Format$(Date, "dd/MM/yyyy") Then 'Validar el formato de la fecha
    ValidaConfiguracionRegional = False
    Exit Function
End If

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



Public Function fgActualizaUltVersionEXE(ByVal psAgenciaCod As String, _
                                        ByVal psRutaUltActualiz As String, _
                                        ByVal psFlagActualizaEXE As String) As Boolean
Dim fs As Scripting.FileSystemObject
Dim fCurrent As Scripting.Folder
Dim fi As Scripting.file
Dim fd As Scripting.file

Dim lsRutaUltActualiz As String
Dim lsRutaSICMACT As String
Dim lsFecUltModifLOCAL As String
Dim lsFecUltModifORIGEN As String
Dim lsFlagActualizaEXE As String

On Error GoTo Error
    fgActualizaUltVersionEXE = False
    'Dim oCons As NConstSistemas
    'Set oCons = New NConstSistemas
    
    lsRutaUltActualiz = psRutaUltActualiz 'oCons.GetRutaAcceso(psAgenciaCod)
    lsRutaSICMACT = App.Path & "\"
    lsFlagActualizaEXE = psFlagActualizaEXE 'oCons.LeeConstSistema(49)
    
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
          If Right(UCase(fi.Name), 3) = "EXE" Or Right(UCase(fi.Name), 3) = "INI" Or Right(UCase(fi.Name), 3) = "DLL" Then
             lsFecUltModifORIGEN = Format(fi.DateLastModified, "yyyy/mm/dd hh:mm:ss")
             If Dir(lsRutaSICMACT & fi.Name) <> "" Then
                Set fd = fs.GetFile(lsRutaSICMACT & fi.Name)
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
        'frmHerActualizaSicmact.IniciaVariables True
        'frmHerActualizaSicmact.Show 1
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

On Error GoTo ErrorInicioExcel 'pti1 26072018

Dim lbExisteHoja As Boolean
Dim lbBorrarRangos As Boolean
'On Error Resume Next
lbExisteHoja = False
lbBorrarRangos = False

'MsgBox "PASO 1.1.1 concluido", vbInformation, "!Exito!" ' pti1 comentado pti1 15/01/2019

activaHoja:

'MsgBox "PASO 1.1.2 concluido", vbInformation, "!Exito!" ' pti1 comentado pti1 15/01/2019

For Each xlHoja1 In xlLibro.Worksheets
    If UCase(xlHoja1.Name) = UCase(psHojName) Then
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

'MsgBox "PASO 1.1.3 concluido", vbInformation, "!Exito!" ' pti1  comentado por pti1 15/01/2019

If Not lbExisteHoja Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = psHojName
    If err Then
        err.Clear
        pbActivaHoja = True
        lbBorrarRangos = True
        GoTo activaHoja
    End If
End If

'MsgBox "PASO 1.1.4 concluido", vbInformation, "!Exito!" ' pti1  comentado por pti1 15/01/2019

Exit Sub 'pti1 26072018
ErrorInicioExcel:         'pti1 26072018
MsgBox err.Description + "Error 1.1: Error al iniciar la creación del excel comunicar a TI", vbInformation, "Error" 'pti1 26072018
    
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
  MsgBox TextErr(err.Description), vbInformation, "Aviso"
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
   MsgBox TextErr(err.Description), vbInformation, "Aviso"
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

Public Sub ExcelCuadro(xlHoja1 As Excel.Worksheet, ByVal x1 As Currency, ByVal Y1 As Currency, ByVal x2 As Currency, ByVal Y2 As Currency, Optional lbLineasVert As Boolean = True, Optional lbLineasHoriz As Boolean = False)
xlHoja1.Range(xlHoja1.Cells(Y1, x1), xlHoja1.Cells(Y2, x2)).BorderAround xlContinuous, xlThin
If lbLineasVert Then
   If x2 <> x1 Then
     xlHoja1.Range(xlHoja1.Cells(Y1, x1), xlHoja1.Cells(Y2, x2)).Borders(xlInsideVertical).LineStyle = xlContinuous
   End If
End If
If lbLineasHoriz Then
    If Y1 <> Y2 Then
        xlHoja1.Range(xlHoja1.Cells(Y1, x1), xlHoja1.Cells(Y2, x2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
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

'**DAOR 20070209
'**Función que permite transponer(Reemplazar caracteres)
Public Function CHRTRAN(psCadena As String, psChrBuscar As String, psChrReemplazo As String) As String
Dim i As Integer
Dim nLenB As Integer, nLenR As Integer, nLenC As Integer
Dim nPosiR As Integer
    nLenB = Len(psChrBuscar)
    nLenR = Len(psChrReemplazo)
    nLenC = Len(psCadena)
    If nLenC > 0 And nLenB > 0 Then
        For i = 1 To nLenB
            If i > nLenR Then
                psCadena = Replace(psCadena, Mid$(psChrBuscar, i, 1), "")
            Else
                psCadena = Replace(psCadena, Mid$(psChrBuscar, i, 1), Mid$(psChrReemplazo, i, 1))
            End If
        Next
    End If
    CHRTRAN = psCadena
End Function

Public Sub GeneraReporte108337(ByVal psNomCmac As String, ByVal psNomAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, ByVal psTitulo As String, ByVal psSubTitulo As String, _
                               ByVal psNomArchivo As String, ByVal prRegistros As ADODB.Recordset, Optional psNomHoja As String = "", Optional Visible As Boolean = False)

    Dim rs As ADODB.Recordset
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim liLineas As Integer, NumFilas As Integer
    Dim fs As Scripting.FileSystemObject
    Dim lnNumColumns As Integer


    If Not (prRegistros.EOF And prRegistros.BOF) Then
        If psNomHoja = "" Then psNomHoja = psNomArchivo
        psNomArchivo = psNomArchivo & "_" & psCodUser & ".xls"

        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application
        If fs.FileExists(App.Path & "\Spooler\" & psNomArchivo) Then
            fs.DeleteFile (App.Path & "\Spooler\" & psNomArchivo)
        End If
        Set xlLibro = xlAplicacion.Workbooks.Add
        Set xlHoja1 = xlLibro.Worksheets.Add
        
        NumFilas = prRegistros.RecordCount
        
        xlHoja1.Name = psNomHoja
        xlHoja1.Cells.Select
    
        xlHoja1.Cells(1, 1) = psNomCmac
        xlHoja1.Cells(1, 11) = Trim(Format(pdFecSis, "dd/mm/yyyy hh:mm:ss"))
        xlHoja1.Cells(2, 1) = psNomAge
        xlHoja1.Cells(2, 11) = psCodUser
        xlHoja1.Cells(4, 2) = psTitulo
        xlHoja1.Cells(5, 2) = "Estadísticas de oficinas compartidas del BN"
        xlHoja1.Cells(NumFilas + 11, 2) = "Análisis de la cartera de oficinas compartidas del BN"

        xlHoja1.Range(xlHoja1.Cells(1, 2), xlHoja1.Cells(5, 11)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 11)).Merge True
        xlHoja1.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(5, 11)).Merge True
        xlHoja1.Range(xlHoja1.Cells(NumFilas + 11, 2), xlHoja1.Cells(NumFilas + 11, 11)).Merge True
        xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(5, 11)).HorizontalAlignment = xlCenter

        xlHoja1.Range("B8") = "UOB"
        xlHoja1.Range("C8") = "Oficina Compartida"
        xlHoja1.Range("D7") = "Número de créditos"
        xlHoja1.Range("D8") = "Nuevos"
        xlHoja1.Range("E8") = "Représtamos"
        xlHoja1.Range("F8") = "Total"
        xlHoja1.Range("D7:F7").Merge True
        
        xlHoja1.Range("H7") = "Monto en soles"
        xlHoja1.Range("H8") = "Nuevos"
        xlHoja1.Range("I8") = "Représtamos"
        xlHoja1.Range("J8") = "Total"
        xlHoja1.Range("H7:J7").Merge True
        
        xlHoja1.Range("K8") = "Nro. de Analistas"
        
        liLineas = 9
        
        xlHoja1.Range("B" & liLineas + NumFilas + 5) = "UOB"
        xlHoja1.Range("C" & liLineas + NumFilas + 5) = "Oficina Compartida"
        xlHoja1.Range("D" & liLineas + NumFilas + 4) = "Número de créditos"
        xlHoja1.Range("D" & liLineas + NumFilas + 5) = "Vigentes"
        xlHoja1.Range("E" & liLineas + NumFilas + 5) = "Refinanciados"
        xlHoja1.Range("F" & liLineas + NumFilas + 5) = "Atrasados"
        xlHoja1.Range("G" & liLineas + NumFilas + 5) = "Total"
        xlHoja1.Range("B" & liLineas + NumFilas + 4 & ":B" & liLineas + NumFilas + 5).Merge True
        xlHoja1.Range("C" & liLineas + NumFilas + 4 & ":C" & liLineas + NumFilas + 5).Merge True
        xlHoja1.Range("D" & liLineas + NumFilas + 4 & ":G" & liLineas + NumFilas + 4).Merge True
        
        xlHoja1.Range("H" & liLineas + NumFilas + 4) = "Monto en soles"
        xlHoja1.Range("H" & liLineas + NumFilas + 5) = "Vigentes"
        xlHoja1.Range("I" & liLineas + NumFilas + 5) = "Refinanciados"
        xlHoja1.Range("J" & liLineas + NumFilas + 5) = "Atrasados"
        xlHoja1.Range("K" & liLineas + NumFilas + 5) = "Total"
        xlHoja1.Range("H" & liLineas + NumFilas + 4 & ":K" & liLineas + NumFilas + 4).Merge True
        
        xlHoja1.Range("B7:K8").Cells.Interior.Color = RGB(220, 220, 220)
        xlHoja1.Range("B7:K8").HorizontalAlignment = xlCenter
        xlHoja1.Range("B7:K8").Font.Bold = True
        xlHoja1.Range("B:K").EntireColumn.AutoFit
        
        xlHoja1.Range("B" & liLineas + NumFilas + 4 & ":K" & liLineas + NumFilas + 5).Cells.Interior.Color = RGB(220, 220, 220)
        xlHoja1.Range("B" & liLineas + NumFilas + 4 & ":K" & liLineas + NumFilas + 5).HorizontalAlignment = xlCenter
        xlHoja1.Range("B" & liLineas + NumFilas + 4 & ":K" & liLineas + NumFilas + 5).Font.Bold = True
        
        xlHoja1.Range("H9:J" & liLineas + NumFilas).Style = "Comma"
        xlHoja1.Range("H" & liLineas + NumFilas + 6 & ":K" & liLineas + NumFilas * 2 + 6).Style = "Comma"
        
       
        While Not prRegistros.EOF
            
            xlHoja1.Range("B" & liLineas) = CInt(prRegistros!cCodAge)
            xlHoja1.Range("C" & liLineas) = prRegistros!vAgencia
            xlHoja1.Range("D" & liLineas) = prRegistros!nNumNuevos
            xlHoja1.Range("E" & liLineas) = prRegistros!nNumRepres
            xlHoja1.Range("F" & liLineas) = prRegistros!nNumTotalMes
            xlHoja1.Range("H" & liLineas) = Format(prRegistros!nSalNuevos, "#.00")
            xlHoja1.Range("I" & liLineas) = Format(prRegistros!nSalRepres, "#.00")
            xlHoja1.Range("J" & liLineas) = Format(prRegistros!nSalTotalMes, "#.00")
            xlHoja1.Range("K" & liLineas) = prRegistros!Analistas
            
            xlHoja1.Range("B" & liLineas + NumFilas + 6) = CInt(prRegistros!cCodAge)
            xlHoja1.Range("C" & liLineas + NumFilas + 6) = prRegistros!vAgencia
            xlHoja1.Range("D" & liLineas + NumFilas + 6) = prRegistros!nNumVig
            xlHoja1.Range("E" & liLineas + NumFilas + 6) = prRegistros!nNumRef
            xlHoja1.Range("F" & liLineas + NumFilas + 6) = prRegistros!nNumAtr
            xlHoja1.Range("G" & liLineas + NumFilas + 6) = prRegistros!nNumTotal
            xlHoja1.Range("H" & liLineas + NumFilas + 6) = Format(prRegistros!nSalVig, "#.00")
            xlHoja1.Range("I" & liLineas + NumFilas + 6) = Format(prRegistros!nSalRef, "#.00")
            xlHoja1.Range("J" & liLineas + NumFilas + 6) = Format(prRegistros!nSalAtr, "#.00")
            xlHoja1.Range("K" & liLineas + NumFilas + 6) = Format(prRegistros!nSalTotal, "#.00")
            
            prRegistros.MoveNext
            liLineas = liLineas + 1
        Wend
        
       
        xlHoja1.Range("C" & liLineas) = "TOTALES"
        xlHoja1.Range("D" & liLineas).Formula = "= SUM(D9:D" & liLineas - 1 & ")"
        xlHoja1.Range("E" & liLineas).Formula = "= SUM(E9:E" & liLineas - 1 & ")"
        xlHoja1.Range("F" & liLineas).Formula = "= SUM(F9:F" & liLineas - 1 & ")"
        xlHoja1.Range("H" & liLineas).Formula = "= SUM(H9:H" & liLineas - 1 & ")"
        xlHoja1.Range("I" & liLineas).Formula = "= SUM(I9:I" & liLineas - 1 & ")"
        xlHoja1.Range("J" & liLineas).Formula = "= SUM(J9:J" & liLineas - 1 & ")"
        xlHoja1.Range("K" & liLineas).Formula = "= SUM(K9:K" & liLineas - 1 & ")"
        xlHoja1.Range("B" & liLineas & ":K" & liLineas).Cells.Interior.Color = RGB(220, 220, 220)
        xlHoja1.Range("B" & liLineas & ":K" & liLineas).Font.Bold = True


        xlHoja1.Range("C" & liLineas + NumFilas + 6) = "TOTALES"
        xlHoja1.Range("D" & liLineas + NumFilas + 6).Formula = "= SUM(D" & liLineas + 6 & ":D" & liLineas + NumFilas + 5 & ")"
        xlHoja1.Range("E" & liLineas + NumFilas + 6).Formula = "= SUM(E" & liLineas + 6 & ":E" & liLineas + NumFilas + 5 & ")"
        xlHoja1.Range("F" & liLineas + NumFilas + 6).Formula = "= SUM(F" & liLineas + 6 & ":F" & liLineas + NumFilas + 5 & ")"
        xlHoja1.Range("G" & liLineas + NumFilas + 6).Formula = "= SUM(G" & liLineas + 6 & ":G" & liLineas + NumFilas + 5 & ")"
        xlHoja1.Range("H" & liLineas + NumFilas + 6).Formula = "= SUM(H" & liLineas + 6 & ":H" & liLineas + NumFilas + 5 & ")"
        xlHoja1.Range("I" & liLineas + NumFilas + 6).Formula = "= SUM(I" & liLineas + 6 & ":I" & liLineas + NumFilas + 5 & ")"
        xlHoja1.Range("J" & liLineas + NumFilas + 6).Formula = "= SUM(J" & liLineas + 6 & ":J" & liLineas + NumFilas + 5 & ")"
        xlHoja1.Range("K" & liLineas + NumFilas + 6).Formula = "= SUM(K" & liLineas + 6 & ":K" & liLineas + NumFilas + 5 & ")"
        xlHoja1.Range("B" & liLineas + NumFilas + 6 & ":K" & liLineas + NumFilas + 6).Cells.Interior.Color = RGB(220, 220, 220)
        xlHoja1.Range("B" & liLineas + NumFilas + 6 & ":K" & liLineas + NumFilas + 6).Font.Bold = True

        xlHoja1.Range("B7:B8").Merge True
        xlHoja1.Range("C7:C8").Merge True
        
        xlHoja1.Range("B" & liLineas + 4 & ":K" & liLineas + NumFilas + 6).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range("B" & liLineas + 4 & ":K" & liLineas + NumFilas + 6).Borders(xlEdgeLeft).Weight = xlThin
        xlHoja1.Range("B" & liLineas + 4 & ":K" & liLineas + NumFilas + 6).Borders(xlEdgeRight).LineStyle = xlContinuous
        xlHoja1.Range("B" & liLineas + 4 & ":K" & liLineas + NumFilas + 6).Borders(xlEdgeRight).Weight = xlThin
        xlHoja1.Range("B" & liLineas + 4 & ":K" & liLineas + NumFilas + 6).Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlHoja1.Range("B" & liLineas + 4 & ":K" & liLineas + NumFilas + 6).Borders(xlEdgeBottom).Weight = xlThin
        xlHoja1.Range("B" & liLineas + 4 & ":K" & liLineas + NumFilas + 6).Borders(xlEdgeTop).LineStyle = xlContinuous
        xlHoja1.Range("B" & liLineas + 4 & ":K" & liLineas + NumFilas + 6).Borders(xlEdgeTop).Weight = xlThin
        xlHoja1.Range("B" & liLineas + 4 & ":K" & liLineas + NumFilas + 6).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range("B" & liLineas + 4 & ":K" & liLineas + NumFilas + 6).Borders(xlInsideVertical).Weight = xlThin
        xlHoja1.Range("B" & liLineas & ":K" & liLineas + NumFilas + 6).Borders(xlEdgeTop).LineStyle = xlContinuous
        xlHoja1.Range("B" & liLineas & ":K" & liLineas + NumFilas + 6).Borders(xlEdgeTop).Weight = xlThin
        xlHoja1.Range("D" & liLineas + 4 & ":K" & liLineas + 4).Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlHoja1.Range("D" & liLineas + 4 & ":K" & liLineas + 4).Borders(xlEdgeBottom).Weight = xlThin

        
        xlHoja1.Range("B7:K" & liLineas).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range("B7:K" & liLineas).Borders(xlEdgeLeft).Weight = xlThin
        xlHoja1.Range("B7:K" & liLineas).Borders(xlEdgeRight).LineStyle = xlContinuous
        xlHoja1.Range("B7:K" & liLineas).Borders(xlEdgeRight).Weight = xlThin
        xlHoja1.Range("B7:K" & liLineas).Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlHoja1.Range("B7:K" & liLineas).Borders(xlEdgeBottom).Weight = xlThin
        xlHoja1.Range("B7:K" & liLineas).Borders(xlEdgeTop).LineStyle = xlContinuous
        xlHoja1.Range("B7:K" & liLineas).Borders(xlEdgeTop).Weight = xlThin
        xlHoja1.Range("B7:K" & liLineas).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range("B7:K" & liLineas).Borders(xlInsideVertical).Weight = xlThin
        xlHoja1.Range("B" & liLineas & ":K" & liLineas).Borders(xlEdgeTop).LineStyle = xlContinuous
        xlHoja1.Range("B" & liLineas & ":K" & liLineas).Borders(xlEdgeTop).Weight = xlThin
        xlHoja1.Range("B8:K8").Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlHoja1.Range("B8:K8").Borders(xlEdgeBottom).Weight = xlThin
        xlHoja1.Range("D8:J8").Borders(xlEdgeTop).LineStyle = xlContinuous
        xlHoja1.Range("D8:J8").Borders(xlEdgeTop).Weight = xlThin
        
        xlHoja1.SaveAs App.Path & "\Spooler\" & psNomArchivo
        MsgBox "Se ha generado el Archivo en " & App.Path & "\Spooler\" & psNomArchivo

        If Visible Then
            xlAplicacion.Visible = True
            xlAplicacion.Windows(1).Visible = True
        Else
            xlLibro.Close
            xlAplicacion.Quit
        End If

        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing

    End If

End Sub

'WIOR 20130829 *****************************************
Public Function EdadPersona(ByVal pdFecNac As Date, Optional ByVal pdFecha As Date = "01/01/1900") As Integer
Dim nEdad As Integer

nEdad = DateDiff("yyyy", pdFecNac, pdFecha)

If Month(pdFecNac) >= Month(pdFecha) Then
    If Month(pdFecNac) = Month(pdFecha) Then
        If Day(pdFecNac) > Day(pdFecha) Then
            nEdad = nEdad - 1
        End If
    Else
        nEdad = nEdad - 1
    End If
End If

EdadPersona = nEdad
End Function

'**DAOR Comentar para compilar Clases.vbp
'By Capi Set 07 Planeamieno
Public Function SayInformacionFepCmac(ByVal psInformacion As String, ByVal pnTipoCambio As Currency, ByVal pnMes As Integer, ByVal pnAno As Integer)

    Dim oDCred As COMDCredito.DCOMCredDoc
    Dim fs As Scripting.FileSystemObject
    Dim pRs As ADODB.Recordset
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsArchivo2 As String
    Dim lbLibroOpen As Boolean
    Dim lsNomHoja  As String
    Dim lsMes As String


    Set pRs = New ADODB.Recordset
    Set oDCred = New COMDCredito.DCOMCredDoc
    Set pRs = oDCred.GetInformacionFepCmac(psInformacion, pnTipoCambio, pnMes, pnAno)

    If pRs Is Nothing Then
        MsgBox "No existe Informacion del Periodo", vbInformation, "Aviso"
        Exit Function
    End If

    'Determinando Descripcion del Mes
    Select Case pnMes
        Case 1
            lsMes = "ENERO"
        Case 2
            lsMes = "FEBRERO"
        Case 3
            lsMes = "MARZO"
        Case 4
            lsMes = "ABRIL"
        Case 5
            lsMes = "MAYO"
        Case 6
            lsMes = "JUNIO"
        Case 7
            lsMes = "JULIO"
        Case 8
            lsMes = "AGOSTO"
        Case 9
            lsMes = "SETIEMBRE"
        Case 10
            lsMes = "OCTUBRE"
        Case 11
            lsMes = "NOVIEMBRE"
        Case 12
            lsMes = "DICIEMBRE"
    End Select

    'Determinando que Archivo y hoja Excel se debe abrir de acuerdo a eleccion del usuario

    Select Case psInformacion
        Case gColCredRepFepInforme01
            lsArchivo1 = "FepCmacInforme01"
            lsNomHoja = "FepCmacInforme01"
        Case gCapCredRepFepInforme02
            lsArchivo1 = "FepCmacInforme02"
            lsNomHoja = "FepCmacInforme02"
        Case gColCredRepFepInforme03
            lsArchivo1 = "FepCmacInforme03"
            lsNomHoja = "FepCmacInforme03"
        Case gColCredRepFepInforme3a
            lsArchivo1 = "FepCmacInforme3a"
            lsNomHoja = "FepCmacInforme3a"
        Case gColCredRepFepInforme3b
            lsArchivo1 = "FepCmacInforme3b"
            lsNomHoja = "FepCmacInforme3b"
        Case gColCredRepFepInforme3c
            lsArchivo1 = "FepCmacInforme3c"
            lsNomHoja = "FepCmacInforme3c"
        Case gColCredRepFepInforme3d
            lsArchivo1 = "FepCmacInforme3d"
            lsNomHoja = "FepCmacInforme3d"
        Case gColCredRepFepInforme04
            lsArchivo1 = "FepCmacInforme04"
            lsNomHoja = "FepCmacInforme04"
        Case gColCredRepFepInforme06
            lsArchivo1 = "FepCmacInforme06"
            lsNomHoja = "FepCmacInforme06"
        Case gColCredRepFepEntorno
            lsArchivo1 = "FepCmacEntorno"
            lsNomHoja = "FepCmacEntorno"
    End Select

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application

    If fs.FileExists(App.Path & "\FormatoCarta\" & lsArchivo1 & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsArchivo1 & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Function
    End If

    lsArchivo2 = lsArchivo1 & "_" & gsCodUser & "_" & Format$(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS")

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
'   lsArchivo1 = "Say_" & lsArchivo1
'   Call lsArchivo1(prs, xlHoja1, pnTipoCambio, lsMes, pnAno)

    Select Case psInformacion
        Case gCapCredRepFepInforme02
            Call Say_FepCmacInforme02(pRs, xlHoja1, pnTipoCambio, lsMes, pnAno)

        Case gColCredRepFepInforme03

            Call Say_FepCmacInforme03(pRs, xlHoja1, pnTipoCambio, lsMes, pnAno)
        Case gColCredRepFepInforme3a
            Call Say_FepCmacInforme3a(pRs, xlHoja1, pnTipoCambio, lsMes, pnAno)
        Case gColCredRepFepInforme3b
            Call Say_FepCmacInforme3b(pRs, xlHoja1, pnTipoCambio, lsMes, pnAno)
        Case gColCredRepFepInforme3c
            Call Say_FepCmacInforme3c(pRs, xlHoja1, pnTipoCambio, lsMes, pnAno)
        Case gColCredRepFepInforme3d
            Call Say_FepCmacInforme3d(pRs, xlHoja1, pnTipoCambio, lsMes, pnAno)
        'Case gColCredRepFepInforme04
        '   Call Say_FepCmacInforme04(prs, xlHoja1, pnTipoCambio, lsMes, pnAno)
        Case gColCredRepFepInforme06
            Call Say_FepCmacInforme06(pRs, xlHoja1, pnTipoCambio, lsMes, pnAno)
        Case gColCredRepFepEntorno
            Call Say_FepCmacEntorno(pRs, xlHoja1, pnTipoCambio, lsMes, pnAno)

    End Select

    xlHoja1.SaveAs App.Path & "\Spooler\" & lsArchivo2 & ".xls"
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

End Function
Public Sub Say_FepCmacInforme03(ByRef pRs As ADODB.Recordset, ByRef xlHoja1 As Excel.Worksheet, ByVal pnTipoCambio, ByVal plsMes, ByVal pnAno)

xlHoja1.Cells(7, 3) = Trim(plsMes)
xlHoja1.Cells(7, 4) = Str(pnAno)
xlHoja1.Cells(7, 6) = Str(pnTipoCambio)

Do While Not pRs.EOF
    If Mid(pRs!Producto, 1, 5) = "Micro" Then
        xlHoja1.Cells(13, 3) = pRs!Cantidad
        xlHoja1.Cells(13, 4) = pRs!Desembolso
        xlHoja1.Cells(14, 3) = pRs!Rango11
        xlHoja1.Cells(14, 4) = pRs!Rango21
        xlHoja1.Cells(15, 3) = pRs!Rango12
        xlHoja1.Cells(15, 4) = pRs!Rango22
        xlHoja1.Cells(16, 3) = pRs!Rango13
        xlHoja1.Cells(16, 4) = pRs!Rango23
        xlHoja1.Cells(17, 3) = pRs!Rango14
        xlHoja1.Cells(17, 4) = pRs!Rango24
        xlHoja1.Cells(18, 3) = pRs!Rango15
        xlHoja1.Cells(18, 4) = pRs!Rango25
        xlHoja1.Cells(19, 3) = pRs!Rango16
        xlHoja1.Cells(19, 4) = pRs!Rango26
        xlHoja1.Cells(20, 3) = pRs!Rango17
        xlHoja1.Cells(20, 4) = pRs!Rango27

        xlHoja1.Cells(23, 3) = pRs!Cantidad
        xlHoja1.Cells(23, 4) = pRs!Desembolso
        xlHoja1.Cells(24, 3) = pRs!Plazo11
        xlHoja1.Cells(24, 4) = pRs!Plazo21
        xlHoja1.Cells(25, 3) = pRs!Plazo12
        xlHoja1.Cells(25, 4) = pRs!Plazo22
        xlHoja1.Cells(26, 3) = pRs!Plazo13
        xlHoja1.Cells(26, 4) = pRs!Plazo23
        xlHoja1.Cells(27, 3) = pRs!Plazo14
        xlHoja1.Cells(27, 4) = pRs!Plazo24
        xlHoja1.Cells(28, 3) = pRs!Plazo15
        xlHoja1.Cells(28, 4) = pRs!Plazo25


        xlHoja1.Cells(31, 3) = pRs!Cantidad
        xlHoja1.Cells(31, 4) = pRs!Desembolso
        xlHoja1.Cells(32, 3) = pRs!Sector11
        xlHoja1.Cells(32, 4) = pRs!Sector21
        xlHoja1.Cells(33, 3) = pRs!Sector12
        xlHoja1.Cells(33, 4) = pRs!Sector22
        xlHoja1.Cells(34, 3) = pRs!Sector13
        xlHoja1.Cells(34, 4) = pRs!Sector23
        xlHoja1.Cells(35, 3) = pRs!Sector14
        xlHoja1.Cells(35, 4) = pRs!Sector24
        xlHoja1.Cells(36, 3) = pRs!Sector15
        xlHoja1.Cells(36, 4) = pRs!Sector25

        xlHoja1.Cells(39, 3) = pRs!Cantidad
        xlHoja1.Cells(39, 4) = pRs!Desembolso
        xlHoja1.Cells(40, 3) = pRs!Ktrabajo11
        xlHoja1.Cells(40, 4) = pRs!Ktrabajo21
        xlHoja1.Cells(41, 3) = pRs!Afijo11
        xlHoja1.Cells(41, 4) = pRs!Afijo21
        xlHoja1.Cells(42, 3) = pRs!Otros11
        xlHoja1.Cells(42, 4) = pRs!Otros21


    Else

        xlHoja1.Cells(13, 5) = pRs!Cantidad
        xlHoja1.Cells(13, 6) = pRs!Desembolso
        xlHoja1.Cells(14, 5) = pRs!Rango11
        xlHoja1.Cells(14, 6) = pRs!Rango12
        xlHoja1.Cells(15, 5) = pRs!Rango12
        xlHoja1.Cells(15, 6) = pRs!Rango22
        xlHoja1.Cells(16, 5) = pRs!Rango13
        xlHoja1.Cells(16, 6) = pRs!Rango23
        xlHoja1.Cells(17, 5) = pRs!Rango14
        xlHoja1.Cells(17, 6) = pRs!Rango24
        xlHoja1.Cells(18, 5) = pRs!Rango15
        xlHoja1.Cells(18, 6) = pRs!Rango25
        xlHoja1.Cells(19, 5) = pRs!Rango16
        xlHoja1.Cells(19, 6) = pRs!Rango26
        xlHoja1.Cells(20, 5) = pRs!Rango17
        xlHoja1.Cells(20, 6) = pRs!Rango27

        xlHoja1.Cells(23, 5) = pRs!Cantidad
        xlHoja1.Cells(23, 6) = pRs!Desembolso


        xlHoja1.Cells(24, 5) = pRs!Plazo11
        xlHoja1.Cells(24, 6) = pRs!Plazo21
        xlHoja1.Cells(25, 5) = pRs!Plazo12
        xlHoja1.Cells(25, 6) = pRs!Plazo22
        xlHoja1.Cells(26, 5) = pRs!Plazo13
        xlHoja1.Cells(26, 6) = pRs!Plazo23
        xlHoja1.Cells(27, 5) = pRs!Plazo14
        xlHoja1.Cells(27, 6) = pRs!Plazo24
        xlHoja1.Cells(28, 5) = pRs!Plazo15
        xlHoja1.Cells(28, 6) = pRs!Plazo25


        xlHoja1.Cells(31, 5) = pRs!Cantidad
        xlHoja1.Cells(31, 6) = pRs!Desembolso
        xlHoja1.Cells(32, 5) = pRs!Sector11
        xlHoja1.Cells(32, 6) = pRs!Sector21
        xlHoja1.Cells(33, 5) = pRs!Sector12
        xlHoja1.Cells(33, 6) = pRs!Sector22
        xlHoja1.Cells(34, 5) = pRs!Sector13
        xlHoja1.Cells(34, 6) = pRs!Sector23
        xlHoja1.Cells(35, 5) = pRs!Sector14
        xlHoja1.Cells(35, 6) = pRs!Sector24
        xlHoja1.Cells(36, 5) = pRs!Sector15
        xlHoja1.Cells(36, 6) = pRs!Sector25


        xlHoja1.Cells(39, 5) = pRs!Cantidad
        xlHoja1.Cells(39, 6) = pRs!Desembolso
        xlHoja1.Cells(40, 5) = pRs!Ktrabajo11
        xlHoja1.Cells(40, 6) = pRs!Ktrabajo21
        xlHoja1.Cells(41, 5) = pRs!Afijo11
        xlHoja1.Cells(41, 6) = pRs!Afijo21
        xlHoja1.Cells(42, 5) = pRs!Otros11
        xlHoja1.Cells(42, 6) = pRs!Otros21


    End If
    pRs.MoveNext
    Loop
End Sub
Public Sub Say_FepCmacInforme3a(ByRef pRs As ADODB.Recordset, ByRef xlHoja1 As Excel.Worksheet, ByVal pnTipoCambio, ByVal plsMes, ByVal pnAno)

Dim lnFila As Integer
Dim lnContador As Integer
Dim lsRango As String
Dim lsAgencia As String

xlHoja1.Cells(5, 2) = Trim(plsMes)
xlHoja1.Cells(5, 3) = Str(pnAno)
xlHoja1.Cells(8, 3) = Str(pnTipoCambio)

lnFila = 12
lnContador = 1
Do While Not pRs.EOF
    lsAgencia = pRs!Agencia
    If lnContador > 1 Then
        'Que Proceda a Copiar el cuadro de acuerdo a Formato
        If Mid(pRs!Producto, 1, 5) = "Micro" Then
            xlHoja1.Range("A12:D22").Copy
            lsRango = "A" & Trim(Str((lnFila * lnContador)))
            xlHoja1.Range(lsRango).PasteSpecial
        Else
            xlHoja1.Range("F12:H22").Copy
            lsRango = "F" & Trim(Str((lnFila * lnContador)))
            xlHoja1.Range(lsRango).PasteSpecial
        End If
    End If
    xlHoja1.Cells(lnFila * lnContador, 1) = pRs!AgenciaDesc
    If Mid(pRs!Producto, 1, 5) = "Micro" Then
        xlHoja1.Cells(lnFila * lnContador + 2, 2) = pRs!Rango11
        xlHoja1.Cells(lnFila * lnContador + 2, 3) = pRs!Rango21
        xlHoja1.Cells(lnFila * lnContador + 2, 4) = pRs!Rango31
        xlHoja1.Cells(lnFila * lnContador + 3, 2) = pRs!Rango12
        xlHoja1.Cells(lnFila * lnContador + 3, 3) = pRs!Rango22
        xlHoja1.Cells(lnFila * lnContador + 3, 4) = pRs!Rango32
        xlHoja1.Cells(lnFila * lnContador + 4, 2) = pRs!Rango13
        xlHoja1.Cells(lnFila * lnContador + 4, 3) = pRs!Rango23
        xlHoja1.Cells(lnFila * lnContador + 4, 4) = pRs!Rango33
        xlHoja1.Cells(lnFila * lnContador + 5, 2) = pRs!Rango14
        xlHoja1.Cells(lnFila * lnContador + 5, 3) = pRs!Rango24
        xlHoja1.Cells(lnFila * lnContador + 5, 4) = pRs!Rango34
        xlHoja1.Cells(lnFila * lnContador + 6, 2) = pRs!Rango15
        xlHoja1.Cells(lnFila * lnContador + 6, 3) = pRs!Rango25
        xlHoja1.Cells(lnFila * lnContador + 6, 4) = pRs!Rango35
        xlHoja1.Cells(lnFila * lnContador + 7, 2) = pRs!Rango16
        xlHoja1.Cells(lnFila * lnContador + 7, 3) = pRs!Rango26
        xlHoja1.Cells(lnFila * lnContador + 7, 4) = pRs!Rango36
        xlHoja1.Cells(lnFila * lnContador + 8, 2) = pRs!Rango17
        xlHoja1.Cells(lnFila * lnContador + 8, 3) = pRs!Rango27
        xlHoja1.Cells(lnFila * lnContador + 8, 4) = pRs!Rango37
        xlHoja1.Cells(lnFila * lnContador + 9, 2) = pRs!Cantidad
        xlHoja1.Cells(lnFila * lnContador + 9, 3) = pRs!Desembolso
        xlHoja1.Cells(lnFila * lnContador + 9, 4) = pRs!Saldo_Capital

        xlHoja1.Cells(lnFila * lnContador + 10, 3) = pRs!Desembolso * pnTipoCambio
        xlHoja1.Cells(lnFila * lnContador + 10, 4) = pRs!Saldo_Capital * pnTipoCambio
    Else
        xlHoja1.Cells(lnFila * lnContador + 2, 6) = pRs!Rango11
        xlHoja1.Cells(lnFila * lnContador + 2, 7) = pRs!Rango21
        xlHoja1.Cells(lnFila * lnContador + 2, 8) = pRs!Rango31
        xlHoja1.Cells(lnFila * lnContador + 3, 6) = pRs!Rango12
        xlHoja1.Cells(lnFila * lnContador + 3, 7) = pRs!Rango22
        xlHoja1.Cells(lnFila * lnContador + 3, 8) = pRs!Rango32
        xlHoja1.Cells(lnFila * lnContador + 4, 6) = pRs!Rango13
        xlHoja1.Cells(lnFila * lnContador + 4, 7) = pRs!Rango23
        xlHoja1.Cells(lnFila * lnContador + 4, 8) = pRs!Rango33
        xlHoja1.Cells(lnFila * lnContador + 5, 6) = pRs!Rango14
        xlHoja1.Cells(lnFila * lnContador + 5, 7) = pRs!Rango24
        xlHoja1.Cells(lnFila * lnContador + 5, 8) = pRs!Rango34
        xlHoja1.Cells(lnFila * lnContador + 6, 6) = pRs!Rango15
        xlHoja1.Cells(lnFila * lnContador + 6, 7) = pRs!Rango25
        xlHoja1.Cells(lnFila * lnContador + 6, 8) = pRs!Rango35
        xlHoja1.Cells(lnFila * lnContador + 7, 6) = pRs!Rango16
        xlHoja1.Cells(lnFila * lnContador + 7, 7) = pRs!Rango26
        xlHoja1.Cells(lnFila * lnContador + 7, 8) = pRs!Rango36
        xlHoja1.Cells(lnFila * lnContador + 8, 6) = pRs!Rango17
        xlHoja1.Cells(lnFila * lnContador + 8, 7) = pRs!Rango27
        xlHoja1.Cells(lnFila * lnContador + 8, 8) = pRs!Rango37
        xlHoja1.Cells(lnFila * lnContador + 9, 6) = pRs!Cantidad
        xlHoja1.Cells(lnFila * lnContador + 9, 7) = pRs!Desembolso
        xlHoja1.Cells(lnFila * lnContador + 9, 8) = pRs!Saldo_Capital

        xlHoja1.Cells(lnFila * lnContador + 10, 7) = pRs!Desembolso * pnTipoCambio
        xlHoja1.Cells(lnFila * lnContador + 10, 8) = pRs!Saldo_Capital * pnTipoCambio
    End If
    pRs.MoveNext
    If pRs.EOF Then
        Exit Do
    ElseIf pRs!Agencia <> lsAgencia Then
       lnContador = lnContador + 1
    End If
Loop
End Sub

Public Sub Say_FepCmacInforme3b(ByRef pRs As ADODB.Recordset, ByRef xlHoja1 As Excel.Worksheet, ByVal pnTipoCambio, ByVal plsMes, ByVal pnAno)

Dim lnFila As Integer
Dim lnContador As Integer
Dim lsRango As String
Dim lsAgencia As String

xlHoja1.Cells(5, 2) = Trim(plsMes)
xlHoja1.Cells(5, 3) = Str(pnAno)
xlHoja1.Cells(8, 3) = Str(pnTipoCambio)

lnFila = 10
lnContador = 1
Do While Not pRs.EOF
    lsAgencia = pRs!Agencia
    If lnContador > 1 Then
        'Que Proceda a Copiar el cuadro de acuerdo a Formato
        If Mid(pRs!Producto, 1, 5) = "Micro" Then
            xlHoja1.Range("A10:D18").Copy
            lsRango = "A" & Trim(Str((lnFila * lnContador)))
            xlHoja1.Range(lsRango).PasteSpecial
        Else
            xlHoja1.Range("F10:H18").Copy
            lsRango = "F" & Trim(Str((lnFila * lnContador)))
            xlHoja1.Range(lsRango).PasteSpecial
        End If
    End If
    xlHoja1.Cells(lnFila * lnContador, 1) = pRs!AgenciaDesc
    If Mid(pRs!Producto, 1, 5) = "Micro" Then
        xlHoja1.Cells(lnFila * lnContador + 2, 2) = pRs!Plazo11
        xlHoja1.Cells(lnFila * lnContador + 2, 3) = pRs!Plazo21
        xlHoja1.Cells(lnFila * lnContador + 2, 4) = pRs!Plazo31
        xlHoja1.Cells(lnFila * lnContador + 3, 2) = pRs!Plazo12
        xlHoja1.Cells(lnFila * lnContador + 3, 3) = pRs!Plazo22
        xlHoja1.Cells(lnFila * lnContador + 3, 4) = pRs!Plazo32
        xlHoja1.Cells(lnFila * lnContador + 4, 2) = pRs!Plazo13
        xlHoja1.Cells(lnFila * lnContador + 4, 3) = pRs!Plazo23
        xlHoja1.Cells(lnFila * lnContador + 4, 4) = pRs!Plazo33
        xlHoja1.Cells(lnFila * lnContador + 5, 2) = pRs!Plazo14
        xlHoja1.Cells(lnFila * lnContador + 5, 3) = pRs!Plazo24
        xlHoja1.Cells(lnFila * lnContador + 5, 4) = pRs!Plazo34
        xlHoja1.Cells(lnFila * lnContador + 6, 2) = pRs!Plazo15
        xlHoja1.Cells(lnFila * lnContador + 6, 3) = pRs!Plazo25
        xlHoja1.Cells(lnFila * lnContador + 6, 4) = pRs!Plazo35
        xlHoja1.Cells(lnFila * lnContador + 7, 2) = pRs!Cantidad
        xlHoja1.Cells(lnFila * lnContador + 7, 3) = pRs!Desembolso
        xlHoja1.Cells(lnFila * lnContador + 7, 4) = pRs!Saldo_Capital

        xlHoja1.Cells(lnFila * lnContador + 8, 3) = pRs!Desembolso * pnTipoCambio
        xlHoja1.Cells(lnFila * lnContador + 8, 4) = pRs!Saldo_Capital * pnTipoCambio
    Else
        xlHoja1.Cells(lnFila * lnContador + 2, 6) = pRs!Plazo11
        xlHoja1.Cells(lnFila * lnContador + 2, 7) = pRs!Plazo21
        xlHoja1.Cells(lnFila * lnContador + 2, 8) = pRs!Plazo31
        xlHoja1.Cells(lnFila * lnContador + 3, 6) = pRs!Plazo12
        xlHoja1.Cells(lnFila * lnContador + 3, 7) = pRs!Plazo22
        xlHoja1.Cells(lnFila * lnContador + 3, 8) = pRs!Plazo32
        xlHoja1.Cells(lnFila * lnContador + 4, 6) = pRs!Plazo13
        xlHoja1.Cells(lnFila * lnContador + 4, 7) = pRs!Plazo23
        xlHoja1.Cells(lnFila * lnContador + 4, 8) = pRs!Plazo33
        xlHoja1.Cells(lnFila * lnContador + 5, 6) = pRs!Plazo14
        xlHoja1.Cells(lnFila * lnContador + 5, 7) = pRs!Plazo24
        xlHoja1.Cells(lnFila * lnContador + 5, 8) = pRs!Plazo34
        xlHoja1.Cells(lnFila * lnContador + 6, 6) = pRs!Plazo15
        xlHoja1.Cells(lnFila * lnContador + 6, 7) = pRs!Plazo25
        xlHoja1.Cells(lnFila * lnContador + 6, 8) = pRs!Plazo35
        xlHoja1.Cells(lnFila * lnContador + 7, 6) = pRs!Cantidad
        xlHoja1.Cells(lnFila * lnContador + 7, 7) = pRs!Desembolso
        xlHoja1.Cells(lnFila * lnContador + 7, 8) = pRs!Saldo_Capital

        xlHoja1.Cells(lnFila * lnContador + 8, 7) = pRs!Desembolso * pnTipoCambio
        xlHoja1.Cells(lnFila * lnContador + 8, 8) = pRs!Saldo_Capital * pnTipoCambio
    End If
    pRs.MoveNext
    If pRs.EOF Then
        Exit Do
    ElseIf pRs!Agencia <> lsAgencia Then
       lnContador = lnContador + 1
    End If
Loop
End Sub

Public Sub Say_FepCmacInforme3c(ByRef pRs As ADODB.Recordset, ByRef xlHoja1 As Excel.Worksheet, ByVal pnTipoCambio, ByVal plsMes, ByVal pnAno)

Dim lnFila As Integer
Dim lnContador As Integer
Dim lsRango As String
Dim lsAgencia As String

xlHoja1.Cells(5, 2) = Trim(plsMes)
xlHoja1.Cells(5, 3) = Str(pnAno)
xlHoja1.Cells(8, 3) = Str(pnTipoCambio)

lnFila = 13
lnContador = 1
Do While Not pRs.EOF
    lsAgencia = pRs!Agencia
    If lnContador > 1 Then
        'Que Proceda a Copiar el cuadro de acuerdo a Formato
        If Mid(pRs!Producto, 1, 5) = "Micro" Then
            xlHoja1.Range("A13:D24").Copy
            lsRango = "A" & Trim(Str((lnFila * lnContador)))
            xlHoja1.Range(lsRango).PasteSpecial
        Else
            xlHoja1.Range("F13:H24").Copy
            lsRango = "F" & Trim(Str((lnFila * lnContador)))
            xlHoja1.Range(lsRango).PasteSpecial
        End If
    End If
    xlHoja1.Cells(lnFila * lnContador, 1) = pRs!AgenciaDesc
    If Mid(pRs!Producto, 1, 5) = "Micro" Then
        xlHoja1.Cells(lnFila * lnContador + 2, 2) = pRs!Sector11
        xlHoja1.Cells(lnFila * lnContador + 2, 3) = pRs!Sector21
        xlHoja1.Cells(lnFila * lnContador + 2, 4) = pRs!sector31
        xlHoja1.Cells(lnFila * lnContador + 3, 2) = pRs!Sector12
        xlHoja1.Cells(lnFila * lnContador + 3, 3) = pRs!Sector22
        xlHoja1.Cells(lnFila * lnContador + 3, 4) = pRs!sector32
        xlHoja1.Cells(lnFila * lnContador + 4, 2) = pRs!Sector13
        xlHoja1.Cells(lnFila * lnContador + 4, 3) = pRs!Sector23
        xlHoja1.Cells(lnFila * lnContador + 4, 4) = pRs!sector33
        xlHoja1.Cells(lnFila * lnContador + 5, 2) = pRs!Sector14
        xlHoja1.Cells(lnFila * lnContador + 5, 3) = pRs!Sector24
        xlHoja1.Cells(lnFila * lnContador + 5, 4) = pRs!sector34
        xlHoja1.Cells(lnFila * lnContador + 6, 2) = pRs!Sector15
        xlHoja1.Cells(lnFila * lnContador + 6, 3) = pRs!Sector25
        xlHoja1.Cells(lnFila * lnContador + 6, 4) = pRs!sector35
        xlHoja1.Cells(lnFila * lnContador + 7, 2) = pRs!sector16
        xlHoja1.Cells(lnFila * lnContador + 7, 3) = pRs!sector26
        xlHoja1.Cells(lnFila * lnContador + 7, 4) = pRs!sector36
        xlHoja1.Cells(lnFila * lnContador + 8, 2) = pRs!sector17
        xlHoja1.Cells(lnFila * lnContador + 8, 3) = pRs!sector27
        xlHoja1.Cells(lnFila * lnContador + 8, 4) = pRs!sector37
        xlHoja1.Cells(lnFila * lnContador + 9, 2) = pRs!sector18
        xlHoja1.Cells(lnFila * lnContador + 9, 3) = pRs!sector28
        xlHoja1.Cells(lnFila * lnContador + 9, 4) = pRs!sector38
        xlHoja1.Cells(lnFila * lnContador + 10, 2) = pRs!Cantidad
        xlHoja1.Cells(lnFila * lnContador + 10, 3) = pRs!Desembolso
        xlHoja1.Cells(lnFila * lnContador + 10, 4) = pRs!Saldo_Capital

        xlHoja1.Cells(lnFila * lnContador + 11, 3) = pRs!Desembolso * pnTipoCambio
        xlHoja1.Cells(lnFila * lnContador + 11, 4) = pRs!Saldo_Capital * pnTipoCambio
    Else
        xlHoja1.Cells(lnFila * lnContador + 2, 6) = pRs!Sector11
        xlHoja1.Cells(lnFila * lnContador + 2, 7) = pRs!Sector21
        xlHoja1.Cells(lnFila * lnContador + 2, 8) = pRs!sector31
        xlHoja1.Cells(lnFila * lnContador + 3, 6) = pRs!Sector12
        xlHoja1.Cells(lnFila * lnContador + 3, 7) = pRs!Sector22
        xlHoja1.Cells(lnFila * lnContador + 3, 8) = pRs!sector32
        xlHoja1.Cells(lnFila * lnContador + 4, 6) = pRs!Sector13
        xlHoja1.Cells(lnFila * lnContador + 4, 7) = pRs!Sector23
        xlHoja1.Cells(lnFila * lnContador + 4, 8) = pRs!sector33
        xlHoja1.Cells(lnFila * lnContador + 5, 6) = pRs!Sector14
        xlHoja1.Cells(lnFila * lnContador + 5, 7) = pRs!Sector24
        xlHoja1.Cells(lnFila * lnContador + 5, 8) = pRs!sector34
        xlHoja1.Cells(lnFila * lnContador + 6, 6) = pRs!Sector15
        xlHoja1.Cells(lnFila * lnContador + 6, 7) = pRs!Sector25
        xlHoja1.Cells(lnFila * lnContador + 6, 8) = pRs!sector35
        xlHoja1.Cells(lnFila * lnContador + 7, 6) = pRs!sector16
        xlHoja1.Cells(lnFila * lnContador + 7, 7) = pRs!sector26
        xlHoja1.Cells(lnFila * lnContador + 7, 8) = pRs!sector36
        xlHoja1.Cells(lnFila * lnContador + 8, 6) = pRs!sector17
        xlHoja1.Cells(lnFila * lnContador + 8, 7) = pRs!sector27
        xlHoja1.Cells(lnFila * lnContador + 8, 8) = pRs!sector37
        xlHoja1.Cells(lnFila * lnContador + 9, 6) = pRs!sector18
        xlHoja1.Cells(lnFila * lnContador + 9, 7) = pRs!sector28
        xlHoja1.Cells(lnFila * lnContador + 9, 8) = pRs!sector38
        xlHoja1.Cells(lnFila * lnContador + 10, 6) = pRs!Cantidad
        xlHoja1.Cells(lnFila * lnContador + 10, 7) = pRs!Desembolso
        xlHoja1.Cells(lnFila * lnContador + 10, 8) = pRs!Saldo_Capital

        xlHoja1.Cells(lnFila * lnContador + 11, 7) = pRs!Desembolso * pnTipoCambio
        xlHoja1.Cells(lnFila * lnContador + 11, 8) = pRs!Saldo_Capital * pnTipoCambio
    End If
    pRs.MoveNext
    If pRs.EOF Then
        Exit Do
    ElseIf pRs!Agencia <> lsAgencia Then
       lnContador = lnContador + 1
    End If
Loop
End Sub
Public Sub Say_FepCmacInforme3d(ByRef pRs As ADODB.Recordset, ByRef xlHoja1 As Excel.Worksheet, ByVal pnTipoCambio, ByVal plsMes, ByVal pnAno)

Dim lnFila As Integer
Dim lnContador As Integer
Dim lsRango As String
Dim lsAgencia As String

xlHoja1.Cells(5, 2) = Trim(plsMes)
xlHoja1.Cells(5, 3) = Str(pnAno)
xlHoja1.Cells(8, 3) = Str(pnTipoCambio)

lnFila = 10
lnContador = 1
Do While Not pRs.EOF
    lsAgencia = pRs!Agencia
    If lnContador > 1 Then
        'Que Proceda a Copiar el cuadro de acuerdo a Formato
        If Mid(pRs!Producto, 1, 5) = "Micro" Then
            xlHoja1.Range("A10:D18").Copy
            lsRango = "A" & Trim(Str((lnFila * lnContador)))
            xlHoja1.Range(lsRango).PasteSpecial
        Else
            xlHoja1.Range("F10:H18").Copy
            lsRango = "F" & Trim(Str((lnFila * lnContador)))
            xlHoja1.Range(lsRango).PasteSpecial
        End If
    End If
    xlHoja1.Cells(lnFila * lnContador, 1) = pRs!AgenciaDesc
    If Mid(pRs!Producto, 1, 5) = "Micro" Then
        xlHoja1.Cells(lnFila * lnContador + 2, 2) = pRs!Afijo11
        xlHoja1.Cells(lnFila * lnContador + 2, 3) = pRs!Afijo21
        xlHoja1.Cells(lnFila * lnContador + 2, 4) = pRs!Afijo31
        xlHoja1.Cells(lnFila * lnContador + 3, 2) = pRs!Ktrabajo11
        xlHoja1.Cells(lnFila * lnContador + 3, 3) = pRs!Ktrabajo21
        xlHoja1.Cells(lnFila * lnContador + 3, 4) = pRs!Ktrabajo31
        xlHoja1.Cells(lnFila * lnContador + 4, 2) = pRs!Otros11
        xlHoja1.Cells(lnFila * lnContador + 4, 3) = pRs!Otros21
        xlHoja1.Cells(lnFila * lnContador + 4, 4) = pRs!Otros31
        xlHoja1.Cells(lnFila * lnContador + 5, 2) = pRs!Cantidad
        xlHoja1.Cells(lnFila * lnContador + 5, 3) = pRs!Desembolso
        xlHoja1.Cells(lnFila * lnContador + 5, 4) = pRs!Saldo_Capital

        xlHoja1.Cells(lnFila * lnContador + 6, 3) = pRs!Desembolso * pnTipoCambio
        xlHoja1.Cells(lnFila * lnContador + 6, 4) = pRs!Saldo_Capital * pnTipoCambio
    Else
        xlHoja1.Cells(lnFila * lnContador + 2, 6) = pRs!Afijo11
        xlHoja1.Cells(lnFila * lnContador + 2, 7) = pRs!Afijo21
        xlHoja1.Cells(lnFila * lnContador + 2, 8) = pRs!Afijo31
        xlHoja1.Cells(lnFila * lnContador + 3, 6) = pRs!Ktrabajo11
        xlHoja1.Cells(lnFila * lnContador + 3, 7) = pRs!Ktrabajo21
        xlHoja1.Cells(lnFila * lnContador + 3, 8) = pRs!Ktrabajo31
        xlHoja1.Cells(lnFila * lnContador + 4, 6) = pRs!Otros11
        xlHoja1.Cells(lnFila * lnContador + 4, 7) = pRs!Otros21
        xlHoja1.Cells(lnFila * lnContador + 4, 8) = pRs!Otros31
        xlHoja1.Cells(lnFila * lnContador + 5, 6) = pRs!Cantidad
        xlHoja1.Cells(lnFila * lnContador + 5, 7) = pRs!Desembolso
        xlHoja1.Cells(lnFila * lnContador + 5, 8) = pRs!Saldo_Capital

        xlHoja1.Cells(lnFila * lnContador + 6, 7) = pRs!Desembolso * pnTipoCambio
        xlHoja1.Cells(lnFila * lnContador + 6, 8) = pRs!Saldo_Capital * pnTipoCambio
    End If
    pRs.MoveNext
    If pRs.EOF Then
        Exit Do
    ElseIf pRs!Agencia <> lsAgencia Then
       lnContador = lnContador + 1
    End If
Loop
End Sub


Public Sub Say_FepCmacInforme06(ByRef pRs As ADODB.Recordset, ByRef xlHoja1 As Excel.Worksheet, ByVal pnTipoCambio, ByVal plsMes, ByVal pnAno)

xlHoja1.Cells(8, 4) = plsMes
xlHoja1.Cells(8, 5) = Str(pnAno)
xlHoja1.Cells(9, 5) = Str(pnTipoCambio)


Do While Not pRs.EOF
    Select Case Trim(pRs!Sexo_Ubigeo)
        Case "Hombres"
            xlHoja1.Cells(14, 4) = pRs!Cantidad
            xlHoja1.Cells(14, 5) = pRs!Desembolso
        Case "Mujeres"
            xlHoja1.Cells(15, 4) = pRs!Cantidad
            xlHoja1.Cells(15, 5) = pRs!Desembolso
        Case "Sexo_ND"
            xlHoja1.Cells(16, 4) = pRs!Cantidad
            xlHoja1.Cells(16, 5) = pRs!Desembolso

        Case "Ubigeo_ND"
            xlHoja1.Cells(24, 4) = pRs!Cantidad
            xlHoja1.Cells(24, 5) = pRs!Desembolso

        Case "Urbano"
            xlHoja1.Cells(22, 4) = pRs!Cantidad
            xlHoja1.Cells(22, 5) = pRs!Desembolso
        Case "Rural"
            xlHoja1.Cells(23, 4) = pRs!Cantidad
            xlHoja1.Cells(23, 5) = pRs!Desembolso
    End Select
    pRs.MoveNext
Loop

End Sub
Public Sub Say_FepCmacInforme02(ByRef pRs As ADODB.Recordset, ByRef xlHoja1 As Excel.Worksheet, ByVal pnTipoCambio, ByVal plsMes, ByVal pnAno)
Dim lnFila As Integer
Dim lnContador As Integer
Dim lsRango As String
Dim lsAgencia As String
Dim lnUbicacion As Integer


xlHoja1.Cells(4, 2) = Trim(plsMes)
xlHoja1.Cells(4, 3) = Str(pnAno)
xlHoja1.Cells(7, 3) = Str(pnTipoCambio)

lnFila = 10
lnContador = 1
Do While Not pRs.EOF
    lsAgencia = pRs!Agencia

    xlHoja1.Cells(lnFila * lnContador + (lnContador * 5) - 5, 1) = pRs!AgenciaDesc
    lnUbicacion = lnFila * lnContador + ((lnContador - 1) * 5)
    If Mid(pRs!Producto, 1, 5) = "AHORR" Then
        xlHoja1.Cells(lnUbicacion + 2, 2) = pRs!Rango11
        xlHoja1.Cells(lnUbicacion + 2, 3) = pRs!Rango21
        xlHoja1.Cells(lnUbicacion + 3, 2) = pRs!Rango12
        xlHoja1.Cells(lnUbicacion + 3, 3) = pRs!Rango22
        xlHoja1.Cells(lnUbicacion + 4, 2) = pRs!Rango13
        xlHoja1.Cells(lnUbicacion + 4, 3) = pRs!Rango23
        xlHoja1.Cells(lnUbicacion + 5, 2) = pRs!Rango14
        xlHoja1.Cells(lnUbicacion + 5, 3) = pRs!Rango24
        xlHoja1.Cells(lnUbicacion + 6, 2) = pRs!Rango15
        xlHoja1.Cells(lnUbicacion + 6, 3) = pRs!Rango25
        xlHoja1.Cells(lnUbicacion + 7, 2) = pRs!Rango16
        xlHoja1.Cells(lnUbicacion + 7, 3) = pRs!Rango26
        xlHoja1.Cells(lnUbicacion + 8, 2) = pRs!Rango17
        xlHoja1.Cells(lnUbicacion + 8, 3) = pRs!Rango27
        xlHoja1.Cells(lnUbicacion + 9, 2) = pRs!Rango18
        xlHoja1.Cells(lnUbicacion + 9, 3) = pRs!Rango28
        xlHoja1.Cells(lnUbicacion + 10, 2) = pRs!Rango19
        xlHoja1.Cells(lnUbicacion + 10, 3) = pRs!Rango29
        xlHoja1.Cells(lnUbicacion + 11, 2) = pRs!Rango1A
        xlHoja1.Cells(lnUbicacion + 11, 3) = pRs!Rango2A
        xlHoja1.Cells(lnUbicacion + 12, 2) = pRs!Numero
        xlHoja1.Cells(lnUbicacion + 12, 3) = pRs!Volumen

        xlHoja1.Cells(lnUbicacion + 13, 3) = pRs!Volumen * pnTipoCambio
    Else
        xlHoja1.Cells(lnUbicacion + 2, 4) = pRs!Rango11
        xlHoja1.Cells(lnUbicacion + 2, 5) = pRs!Rango21
        xlHoja1.Cells(lnUbicacion + 3, 4) = pRs!Rango12
        xlHoja1.Cells(lnUbicacion + 3, 5) = pRs!Rango22
        xlHoja1.Cells(lnUbicacion + 4, 4) = pRs!Rango13
        xlHoja1.Cells(lnUbicacion + 4, 5) = pRs!Rango23
        xlHoja1.Cells(lnUbicacion + 5, 4) = pRs!Rango14
        xlHoja1.Cells(lnUbicacion + 5, 5) = pRs!Rango24
        xlHoja1.Cells(lnUbicacion + 6, 4) = pRs!Rango15
        xlHoja1.Cells(lnUbicacion + 6, 5) = pRs!Rango25
        xlHoja1.Cells(lnUbicacion + 7, 4) = pRs!Rango16
        xlHoja1.Cells(lnUbicacion + 7, 5) = pRs!Rango26
        xlHoja1.Cells(lnUbicacion + 8, 4) = pRs!Rango17
        xlHoja1.Cells(lnUbicacion + 8, 5) = pRs!Rango27
        xlHoja1.Cells(lnUbicacion + 9, 4) = pRs!Rango18
        xlHoja1.Cells(lnUbicacion + 9, 5) = pRs!Rango28
        xlHoja1.Cells(lnUbicacion + 10, 4) = pRs!Rango19
        xlHoja1.Cells(lnUbicacion + 10, 5) = pRs!Rango29
        xlHoja1.Cells(lnUbicacion + 11, 4) = pRs!Rango1A
        xlHoja1.Cells(lnUbicacion + 11, 5) = pRs!Rango2A
        xlHoja1.Cells(lnUbicacion + 12, 4) = pRs!Numero
        xlHoja1.Cells(lnUbicacion + 12, 5) = pRs!Volumen

        xlHoja1.Cells(lnUbicacion + 13, 5) = pRs!Volumen * pnTipoCambio
    End If
    pRs.MoveNext
    If pRs.EOF Then
        Exit Do
    ElseIf pRs!Agencia <> lsAgencia Then
       lnContador = lnContador + 1

       'Que Proceda a Copiar el cuadro de acuerdo a Formato
       xlHoja1.Range("A10:E23").Copy
       lsRango = "A" & Trim(Str((lnFila * lnContador + ((lnContador - 1) * 5))))
       xlHoja1.Range(lsRango).PasteSpecial

    End If
Loop
End Sub

Public Sub Say_FepCmacEntorno(ByRef pRs As ADODB.Recordset, ByRef xlHoja1 As Excel.Worksheet, ByVal pnTipoCambio, ByVal plsMes, ByVal pnAno)
Dim lnFila As Integer
Dim lnContador As Integer
Dim lsRango As String
Dim lsAgencia As String
Dim lnUbicacion As Integer


xlHoja1.Cells(8, 2) = Trim(plsMes)
xlHoja1.Cells(8, 3) = Str(pnAno)
xlHoja1.Cells(8, 6) = Str(pnTipoCambio)

lnFila = 20
lnContador = 1
Do While Not pRs.EOF

    xlHoja1.Cells(lnFila + lnContador, 2) = pRs!Institucion
    xlHoja1.Cells(lnFila + lnContador, 5) = pRs!Cantidad
    xlHoja1.Cells(lnFila + lnContador, 6) = pRs!Total_Saldo
    xlHoja1.Cells(lnFila + lnContador, 7) = pRs!Cal_ONORMAL
    xlHoja1.Cells(lnFila + lnContador, 8) = pRs!CAL_1CPP
    xlHoja1.Cells(lnFila + lnContador, 9) = pRs!CAL_2DEFICIENTE
    xlHoja1.Cells(lnFila + lnContador, 10) = pRs!CAL_3DUDOSO
    xlHoja1.Cells(lnFila + lnContador, 11) = pRs!CAL_4PERDIDO
    pRs.MoveNext
    If pRs.EOF Then
        Exit Do
    End If
    lnContador = lnContador + 1
    'Que Proceda a Copiar el cuadro de acuerdo a Formato
     xlHoja1.Range("A20:IV20").Copy
       lsRango = ("A" & Trim(Str(lnFila + lnContador)) & ":IV" & Trim(Str(lnFila + lnContador)))
       xlHoja1.Range(lsRango).Insert
       xlHoja1.Range(lsRango).PasteSpecial

Loop
End Sub


'**DAOR 20070927, Funciòn que genera archivo excel
Public Sub GeneraArchivoExcel(psNomArchivo As String, pMatCabeceras As Variant, prRegistros As ADODB.Recordset, _
    Optional pnNumDecimales As Integer, Optional Visible As Boolean = False, Optional psNomHoja As String = "")
    Dim rs As ADODB.Recordset
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim liLineas As Integer, i As Integer
    Dim fs As Scripting.FileSystemObject
    Dim lnNumColumns As Integer


    If Not (prRegistros.EOF And prRegistros.BOF) Then

        lnNumColumns = UBound(pMatCabeceras)
        lnNumColumns = IIf(prRegistros.Fields.count < lnNumColumns, prRegistros.Fields.count, lnNumColumns)

        If psNomHoja = "" Then psNomHoja = psNomArchivo
        psNomArchivo = psNomArchivo & ".XLS"

        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application
        If fs.FileExists(App.Path & "\FormatoCarta\" & psNomArchivo) Then
            fs.DeleteFile (App.Path & "\FormatoCarta\" & psNomArchivo)
        End If
        Set xlLibro = xlAplicacion.Workbooks.Add
        Set xlHoja1 = xlLibro.Worksheets.Add

        xlHoja1.Name = psNomHoja
        xlHoja1.Cells.Select
        'xlHoja1.Cells.NumberFormat = "@"

        'Cabeceras
        liLineas = 1
        For i = 0 To lnNumColumns - 1
            xlHoja1.Cells(liLineas, i + 1) = pMatCabeceras(i, 0)
        Next i

        liLineas = liLineas + 1
        xlHoja1.Range("A2").CopyFromRecordset prRegistros 'Copia el contenido del recordset a excel

        xlHoja1.SaveAs App.Path & "\FormatoCarta\" & psNomArchivo
        MsgBox "Se ha generado el Archivo en " & App.Path & "\FormatoCarta\" & psNomArchivo

        If Visible Then
            xlAplicacion.Visible = True
            xlAplicacion.Windows(1).Visible = True
        Else
            xlLibro.Close
            xlAplicacion.Quit
        End If


        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing

    End If
End Sub

'**DAOR 20071124, Funciòn que genera reporte en archivo excel
Public Sub GeneraReporteEnArchivoExcel(ByVal psNomCmac As String, _
                                       ByVal psNomAge As String, _
                                       ByVal psCodUser As String, _
                                       ByVal pdFecSis As Date, _
                                       ByVal psTitulo As String, _
                                       ByVal psSubTitulo As String, _
                                       ByVal psNomArchivo As String, _
                                       ByVal pMatCabeceras As Variant, _
                                       ByVal prRegistros As ADODB.Recordset, _
                                       Optional pnNumDecimales As Integer, _
                                       Optional Visible As Boolean = False, _
                                       Optional psNomHoja As String = "", _
                                       Optional pbSinFormatDeReg As Boolean = False, _
                                       Optional pbUsarCabecerasDeRS As Boolean = False, _
                                       Optional psRuta As String = "")
    Dim rs As ADODB.Recordset
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim liLineas As Integer, i As Integer
    Dim fs As Scripting.FileSystemObject
    Dim lnNumColumns As Integer


    If Not (prRegistros.EOF And prRegistros.BOF) Then
        If pbUsarCabecerasDeRS = True Then
            lnNumColumns = prRegistros.Fields.count
        Else
            lnNumColumns = UBound(pMatCabeceras)
            lnNumColumns = IIf(prRegistros.Fields.count < lnNumColumns, prRegistros.Fields.count, prRegistros.Fields.count)
        End If

        If psNomHoja = "" Then psNomHoja = psNomArchivo
        psNomArchivo = psNomArchivo & "_" & psCodUser & ".xls"

        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application

        '**************************************************************
        '**Modificado por ELRO 20110714, según acta 158-2011/TI-D

        If psRuta = "" Then
            If fs.FileExists(App.Path & "\Spooler\" & psNomArchivo) Then
            fs.DeleteFile (App.Path & "\Spooler\" & psNomArchivo)
            End If
        Else
            If fs.FileExists(psRuta & psNomArchivo) Then
                fs.DeleteFile (psRuta & psNomArchivo)
            End If
        End If



        '**************************************************************

        Set xlLibro = xlAplicacion.Workbooks.Add
        Set xlHoja1 = xlLibro.Worksheets.Add

        xlHoja1.Name = psNomHoja
        xlHoja1.Cells.Select

        'Cabeceras
        xlHoja1.Cells(2, 1) = psNomCmac
        xlHoja1.Cells(2, lnNumColumns) = Trim(Format(pdFecSis, "dd/mm/yyyy hh:mm:ss"))
        xlHoja1.Cells(2, 1) = psNomAge
        xlHoja1.Cells(2, lnNumColumns) = psCodUser
        xlHoja1.Cells(4, 1) = psTitulo
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, lnNumColumns)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(5, lnNumColumns)).HorizontalAlignment = xlCenter

        liLineas = 6
        If pbUsarCabecerasDeRS = True Then
            For i = 0 To prRegistros.Fields.count - 1
                xlHoja1.Cells(liLineas, i + 1) = prRegistros.Fields(i).Name
            Next i
        Else
            For i = 0 To lnNumColumns - 1
                If (i + 1) > UBound(pMatCabeceras) Then
                    xlHoja1.Cells(liLineas, i + 1) = prRegistros.Fields(i).Name
                Else
                    xlHoja1.Cells(liLineas, i + 1) = pMatCabeceras(i, 0)
                End If
            Next i
        End If

        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).Cells.Interior.Color = RGB(220, 220, 220)
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).HorizontalAlignment = xlCenter

        If pbSinFormatDeReg = False Then
            liLineas = liLineas + 1
            While Not prRegistros.EOF
                For i = 0 To lnNumColumns - 1
                    If pMatCabeceras(i, 1) = "" Then  'Verificamos si tiene tipo
                        xlHoja1.Cells(liLineas, i + 1) = prRegistros(i)
                    Else
                        Select Case pMatCabeceras(i, 1)
                            Case "S"
                                xlHoja1.Cells(liLineas, i + 1) = prRegistros(i)
                            Case "N"
                                xlHoja1.Cells(liLineas, i + 1) = Format(prRegistros(i), "#0.00")
                            Case "D"
                                xlHoja1.Cells(liLineas, i + 1) = IIf(Format(prRegistros(i), "yyyymmdd") = "19000101", "", Format(prRegistros(i), "dd/mm/yyyy"))
                        End Select
                    End If
                Next i
                liLineas = liLineas + 1
                prRegistros.MoveNext
            Wend
        Else
            xlHoja1.Range("A7").CopyFromRecordset prRegistros 'Copia el contenido del recordset a excel
        End If

        '**************************************************************
        '**Modificado por ELRO 20110714, según acta 158-2011/TI-D

        If psRuta = "" Then
            xlHoja1.SaveAs App.Path & "\Spooler\" & psNomArchivo
            MsgBox "Se ha generado el Archivo en " & App.Path & "\Spooler\" & psNomArchivo

        Else
            xlHoja1.SaveAs psRuta & psNomArchivo
            MsgBox "Se ha generado el Archivo en " & psRuta & psNomArchivo
        End If



        '**************************************************************

        If Visible Then
            xlAplicacion.Visible = True
            xlAplicacion.Windows(1).Visible = True
        'By Capi 19082008 se modifico para que se visualice correctamente
        Else
            xlLibro.Close
            xlAplicacion.Quit
        End If

        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing

    End If
End Sub

'ALPA 20081010
Public Function devCelda(ByVal nNPos As Integer) As String
    Dim sCellda As String
    Dim sMatrixCeldas() As String
    Dim nContador As Integer
    ReDim Preserve sMatrixCeldas(0 To 0)
    nContador = 0
    Dim i As Integer
    For i = 97 To 122
        nContador = nContador + 1
        ReDim Preserve sMatrixCeldas(0 To nContador)
        sMatrixCeldas(nContador) = Chr(i)
    Next i
    For i = 1 To nContador
    If nNPos >= (nContador * i - nContador) And nNPos <= (nContador * i) Then
        If nNPos > nContador Then
            sCellda = sMatrixCeldas(Round(IIf(nNPos >= nContador + (nContador / 2) + 1, Round(nNPos / nContador) - 1, Round(nNPos / nContador)))) & sMatrixCeldas(nNPos - (nContador * (i - 1)))
            Exit For
        Else
            sCellda = sMatrixCeldas(nNPos)
             Exit For
        End If
    End If
    Next i
    devCelda = sCellda
End Function
'By capi 28102008
Public Function GetMaquinaUsuario() As String  'Para obtener la Maquina del Usuario
    Dim buffMaq As String
    Dim lSizeMaq As Long
    buffMaq = space(255)
    lSizeMaq = Len(buffMaq)
    GetComputerName buffMaq, lSizeMaq
    GetMaquinaUsuario = Trim(Left$(buffMaq, lSizeMaq))
End Function

Public Sub RecuperaCodigoOpeTarjeta()

    gnCodOpeTarj = CInt(LeeConstanteSist(413))

End Sub

Public Function CalculaComisionPreCancelacion(ByVal pnSaldoCanc As Currency, ByVal psCtaCod As String) As Double
Dim nMontoMin As Double, nPorc As Double, nMontoCom As Double
Dim rs As ADODB.Recordset, rsTC As ADODB.Recordset, nTC As Currency
Dim oDCred As COMDCredito.DCOMCredito
Dim oParam As COMDColocPig.DCOMColPCalculos
Set oDCred = New COMDCredito.DCOMCredito
Set oParam = New COMDColocPig.DCOMColPCalculos

    Set rsTC = oDCred.DevolverTCMoneda(gdFecSis)
    nTC = rsTC!nVenta
    nMontoMin = Format(oParam.dObtieneColocParametro(gColPParamMontoMinPreCancPersJur), "#,##0.00")
    nPorc = Format(oParam.dObtieneColocParametro(gColPParamPorcPreCancPersJur), "#,##0.00")

    nMontoMin = nMontoMin / IIf(Mid(psCtaCod, 9, 1) = "1", 1, nTC)
    nMontoCom = pnSaldoCanc * (nPorc / 100)
    CalculaComisionPreCancelacion = Format(IIf(nMontoCom < nMontoMin, nMontoMin, nMontoCom), "#,##0.00")
End Function
'EJVG20140408 ***
Public Function DeducirMontoxITF(ByVal pnMonto As Currency, Optional ByRef pnITF As Currency = 0) As Currency
    Dim oNCred As New COMNCredito.NCOMCredito
    Dim lnMonto As Currency, ITF As Currency

    lnMonto = pnMonto
    Do While lnMonto > 0
        ITF = oNCred.DameMontoITF(lnMonto)
        If (lnMonto + ITF) = pnMonto Then
            Exit Do
        ElseIf (lnMonto + ITF) < pnMonto Then
            lnMonto = Round(lnMonto, 2)
            Exit Do
        End If
        lnMonto = lnMonto - 0.005
    Loop
    DeducirMontoxITF = lnMonto
    pnITF = ITF
    Set oNCred = Nothing
End Function
'END EJVG *******
'EJVG20140506 ***
Public Sub IniciarVerDocsPendiente()
    On Error GoTo ErrIniciarVerDocsPendiente
    Dim oNGar As New COMNCredito.NCOMGarantia
    Dim oform As New frmCredVerDocsPendiente
    Dim lbMostrarDocsPendiente As Boolean
    Dim lnNivelVerDocsPendiente As Integer
    lbMostrarDocsPendiente = oNGar.MostrarDocsPendiente(gdFecSis)
    If lbMostrarDocsPendiente Then
        lnNivelVerDocsPendiente = oNGar.NivelVerDocsPendiente(gsGruposUser, gsCodUser, gsCodCargo)
        If lnNivelVerDocsPendiente > 0 Then
            oform.Inicio IIf(lnNivelVerDocsPendiente = 1, False, True)
        End If
    End If
    Set oform = Nothing
    Set oNGar = Nothing
    Exit Sub
ErrIniciarVerDocsPendiente:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
'END EJVG *******
'FRHU ERS077-2015 20151130

Public Sub VerSiClienteActualizoAutorizoSusDatos(ByVal psPersCod As String, Optional ByVal psCodOpe As String = "0")
    Dim oNPersona As New COMNPersona.NCOMPersona
    Dim oform As New frmPersonaActualizacionDatos

    Dim lbClienteActualizoAutorizoSusDatos As ADODB.Recordset 'add pti1
    'Dim lbClienteActualizoAutorizoSusDatos As Boolean 'comentado por pti1
    Dim lbOpeBlackList As Boolean 'add por pti1
    Dim lbClienteEsPersonaNatural As Boolean
    Dim lbValidarOperacion As Boolean
    Dim lbValidarCargoUsuario As Boolean 'FRHU 20151210 TIC1512100001
On Error GoTo ErrorVerSiClienteActualizoAutorizoSusDatos

    lbOpeBlackList = oNPersona.codOpeBlackList(psCodOpe, "ERS0702018") 'ADD PTI1 ERS070-2018 22/12/2018
    If lbOpeBlackList Then 'ADD PTI1 ERS070-2018 22/12/2018
     Exit Sub
    End If

    If psCodOpe = "0" Then
        lbValidarOperacion = True
    Else
        lbValidarOperacion = oNPersona.ValidarOperacionActAutoDatos(psCodOpe)
    End If
    lbValidarCargoUsuario = oNPersona.ValidarCargoParaActAutoDatos(gsCodCargo) 'FRHU 20151210 TIC1512100001
    'If lbValidarOperacion Then
    If lbValidarOperacion And lbValidarCargoUsuario Then 'FRHU 20151210 TIC1512100001
        lbClienteEsPersonaNatural = oNPersona.ClienteEsPersonaNatural(psPersCod) 'VERFICAMOS SI ES CLIENTE NATURAL
        If lbClienteEsPersonaNatural Then
            Set lbClienteActualizoAutorizoSusDatos = oNPersona.ClienteActualizoAutorizoSusDatos(psPersCod) 'add pti1 2018/12/06 ERS070-2018
            'lbClienteActualizoAutorizoSusDatos = oNPersona.ClienteActualizoAutorizoSusDatos(psPersCod) 'comentado por pti1 2018/12/06
            If Not (lbClienteActualizoAutorizoSusDatos.EOF And lbClienteActualizoAutorizoSusDatos.BOF) Then
                'ya se registro el usuario
                Dim autorizos As Integer
                autorizos = lbClienteActualizoAutorizoSusDatos!nAutorizaUsoDatos

                     'el cliente ya autorizo sus datos entonces solo se actualizará
                     If (Trim(lbClienteActualizoAutorizoSusDatos!Mes) > 5) Then
                       oform.Inicio psPersCod, psCodOpe, "1", autorizos
                     End If



            Else
                'aun no se registra, se procederá con la autorizacion
                     oform.Inicio psPersCod, psCodOpe, "0", "0"
            End If

'            If Not lbClienteActualizoAutorizoSusDatos Then comentado por pti1
'                oForm.Inicio (psPersCod)
'            End If 'fin comentado por pti1
        End If
    End If
    Set oform = Nothing
    Set oNPersona = Nothing
    Exit Sub
ErrorVerSiClienteActualizoAutorizoSusDatos:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
'FIN FRHU 20151130
'EJVG20150317 ***
Public Function fgFechaHoraMovDate(ByVal psMovNro As String) As Date
    fgFechaHoraMovDate = CDate(Mid(psMovNro, 1, 4) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 7, 2) & " " & Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2))
End Function
Public Function TieneGarantiasPendienteMigracion(ByVal psCtaCod As String, Optional ByVal pbMensaje As Boolean = True) As Boolean
    Dim oGarantia As New COMDCredito.DCOMGarantia
    Dim lsCadCredMigrar As String

    lsCadCredMigrar = oGarantia.CadenaGarantiasPendienteMigracion(psCtaCod)
    Set oGarantia = Nothing

    If Len(lsCadCredMigrar) > 0 Then
        TieneGarantiasPendienteMigracion = True

        If pbMensaje Then
            MsgBox lsCadCredMigrar, vbInformation, "Aviso"
        End If
    End If
End Function
Public Function RecalcularCoberturaGarantias(ByVal psCtaCod As String, ByVal pbLeasing As Boolean, _
                                                ByVal psTpoProdCod As String, ByVal psTpoProdDesc, _
                                                ByVal pnMonto As Currency, _
                                                ByRef pvGravamen() As tGarantiaGravamen) As Boolean
    Dim frm As frmGarantiaCobertura
    Dim lbAfectaProducto As Boolean, lbAfectaMonto As Boolean
    Dim lsTpoProdCod As String
    Dim lnMonto As Currency

    Dim oGarantia As COMDCredito.DCOMGarantia
    Dim rsGarantia As ADODB.Recordset
    Dim bColocGarantia As Boolean

    RecalcularCoberturaGarantias = False

    ReDim pvGravamen(0)

    'Verificar que estén actualizadas las Garantías
    Do While (TieneGarantiasPendienteMigracion(psCtaCod, False))
        If MsgBox("El Préstamo está coberturado por Garantías migradas." & Chr(13) & "Es necesario que vuelva a actualizar las Garantías y Registrar las Coberturas." & Chr(13) & Chr(13) & "¿Desea realizar la actualización ahora?", vbInformation + vbYesNo, "Actualizar información de Garantías") = vbYes Then
            Set frm = New frmGarantiaCobertura
            '**ARLO20180712 ERS042 - 2018
            Dim bCredito As Boolean '**ARLO20180712 ERS042 - 2018
            Set objProducto = New COMDCredito.DCOMCredito
            If objProducto.GetResultadoCondicionCatalogo("N0000086", psTpoProdCod) Then
                bCredito = True
            Else
                bCredito = False
            End If
            'RecalcularCoberturaGarantias = frm.Inicio(InicioGravamenxAjuste, IIf(psTpoProdCod <> "514", Credito, CartaFianza), psCtaCod, pbLeasing, lbAfectaProducto, psTpoProdCod, psTpoProdDesc, pnMonto, pvGravamen)
            RecalcularCoberturaGarantias = frm.Inicio(InicioGravamenxAjuste, IIf(bCredito, Credito, CartaFianza), psCtaCod, pbLeasing, lbAfectaProducto, psTpoProdCod, psTpoProdDesc, pnMonto, pvGravamen)
            '**ARLO20180712 ERS042 - 2018
            Set frm = Nothing

            If RecalcularCoberturaGarantias Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    Loop

    Set oGarantia = New COMDCredito.DCOMGarantia
    Do
        Set rsGarantia = oGarantia.RecuperaColocGravamen(psCtaCod)
        If rsGarantia.RecordCount = 0 Then
            If MsgBox("Se necesita Registrar las Coberturas del Préstamo." & Chr(13) & Chr(13) & "¿Desea realizar el Registro de Coberturas ahora?", vbInformation + vbYesNo, "Registro de Cobertura") = vbYes Then
                Set frm = New frmGarantiaCobertura
                RecalcularCoberturaGarantias = frm.Inicio(InicioGravamenxAjuste, IIf(psTpoProdCod <> "514", Credito, CartaFianza), psCtaCod, pbLeasing, lbAfectaProducto, psTpoProdCod, psTpoProdDesc, pnMonto, pvGravamen)
                Set frm = Nothing

                If RecalcularCoberturaGarantias Then
                    RSClose rsGarantia
                    Set oGarantia = Nothing
                    Exit Function
                End If
            Else
                RSClose rsGarantia
                Set oGarantia = Nothing
                Exit Function
            End If
        Else
            lsTpoProdCod = rsGarantia!cTpoProdCod
            lnMonto = rsGarantia!nMontoColocado
            bColocGarantia = True
        End If
    Loop While (Not bColocGarantia)
    RSClose rsGarantia
    Set oGarantia = Nothing

    'Ajuste x Cambio de datos ahora con un proceso anterior
    RecalcularCoberturaGarantias = True

    If psTpoProdCod <> lsTpoProdCod Then
        lbAfectaProducto = True
    End If
    If pnMonto <> lnMonto Then
        lbAfectaMonto = True
    End If

    If lbAfectaProducto Or lbAfectaMonto Then
        Set frm = New frmGarantiaCobertura
        If lbAfectaProducto Then
            MsgBox "El Producto del Crédito está cambiando, para lo cual deberá volver a registrar el Gravamen de las Garantías", vbInformation, "Aviso"
        Else
            MsgBox "Los cambios en este proceso han variado con relación al Registro de Cobertura." & Chr(13) & "Para continuar debe ajustar las coberturas.", vbInformation, "Aviso"
        End If
        RecalcularCoberturaGarantias = frm.Inicio(InicioGravamenxAjuste, IIf(psTpoProdCod <> "514", Credito, CartaFianza), psCtaCod, pbLeasing, lbAfectaProducto, psTpoProdCod, psTpoProdDesc, pnMonto, pvGravamen)
    End If
    Set frm = Nothing
End Function
Public Sub VerificarFechaSistema(ByRef obj As Object, Optional ByVal pbSalirSistema As Boolean = False)
    Dim oSis As New NConstSistemas
    Dim ldFechaSistema As Date

    ldFechaSistema = CDate(oSis.LeeConstSistema(gConstSistFechaSistema))
    Set oSis = Nothing

    If gdFecSis <> ldFechaSistema Then
        MsgBox "La Fecha de tu sesión en el Negocio no coincide con la fecha del Sistema", vbCritical, "Aviso"
        If pbSalirSistema Then
            Call SalirSICMACMNegocio
            Unload obj
            End
        End If
    End If
End Sub
Private Sub SalirSICMACMNegocio()
    Dim oSeguridad As New COMManejador.Pista
    Call oSeguridad.InsertarPista(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, "Salida del " & frmLogin.gsSistDescripcion & " Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-" & frmLogin.gsFechaVersion)
     If oSeguridad.ValidaAccesoPistaRF(gsCodUser) Then
            Call oSeguridad.InsertarPistaSesion(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, frmLogin.gnSistCod)
     End If
    Set oSeguridad = Nothing
End Sub
Public Function EsAgenciaConNivApr(ByVal psCodAge As String) As Boolean
    Dim oCS As New NCOMConstSistema
    Dim sAgencias As String
    Dim sCod() As String
    Dim i As Integer

    On Error GoTo ErrEsAgenciaConNivApr

    sAgencias = oCS.LeeConstSistema(gConstSistAgenciaHabNivelesApr)
    Set oCS = Nothing

    sCod = Split(sAgencias, ",")
    For i = 0 To UBound(sCod)
        If psCodAge = sCod(i) Then
            EsAgenciaConNivApr = True
            Exit Function
        End If
    Next
    Exit Function
ErrEsAgenciaConNivApr:
    MsgBox err.Description, vbCritical, "Error"
End Function
Public Function TieneNivelAprobacionPendiente(ByVal psCtaCod As String, Optional ByVal psAgeCod As String = "", Optional ByVal pbMensaje As Boolean = True, _
    Optional ByVal psPersCod As String = "") As Boolean ' RECO20160421 SE AGREGO PARAMETRO psPersCod
    Dim oDNiv As COMDCredito.DCOMNivelAprobacion
    Dim oDCOL As COMDCredito.DCOMCredito
    Dim rs As ADODB.Recordset
    Dim lsAgeCod As String

    Dim oCred As New COMDCredito.DCOMCredito 'RECO20160421
    'Dim oCliPre As New COMNCredito.NCOMCredito 'RECO20160421 'COMENTADO POR ARLO 20170722
    Dim bClientPref As Boolean 'RECO20160421


    lsAgeCod = psAgeCod
    If Len(lsAgeCod) = 0 Then
        Set oDCOL = New COMDCredito.DCOMCredito
        Set rs = New ADODB.Recordset

        Set rs = oDCOL.RecuperaColocaciones(psCtaCod)
        If Not rs.EOF Then
            lsAgeCod = rs!cAgeCodAct
        End If
    End If

    Set oDCOL = Nothing
    RSClose rs
    'RECO20160421    **************************************************
    If psPersCod = "" Then
        psPersCod = oCred.RecuperaTitularCredito(psCtaCod)
    End If

    'bClientPref = oCliPre.ValidarClientePreferencial(psPersCod)    'COMENTADO POR ARLO 20170722
    bClientPref = False 'ARLO 20170722
    'RECO FIN *********************************************************

    If EsAgenciaConNivApr(lsAgeCod) Then
        Set oDNiv = New COMDCredito.DCOMNivelAprobacion
        If oDNiv.ExisteAprobacionCredNivelesPendientes(psCtaCod, IIf(bClientPref = True, 2, 1)) Then 'RECO20160421
        'If oDNiv.ExisteAprobacionCredNivelesPendientes(psCtaCod) Then
            TieneNivelAprobacionPendiente = True

            If pbMensaje Then
                MsgBox "El crédito aún no ha recibido los V°B° de todos los niveles de aprobación.", vbInformation, "Aviso"
            End If
        End If
    End If
    Set oDNiv = Nothing
    Set oCred = Nothing
End Function
'END EJVG *******

'JOEP 20160706 ERS004-2016
Public Sub CargarComboBox(ByVal lrDatos As ADODB.Recordset, ByVal cboControl As ComboBox)
    Do Until lrDatos.EOF
        cboControl.AddItem "" & lrDatos!cDescripcion
        cboControl.ItemData(cboControl.NewIndex) = "" & lrDatos!cValor
        lrDatos.MoveNext
    Loop
    Set lrDatos = Nothing

    cboControl.ListIndex = 0
End Sub

'JOEP20171015 Flujo de Caja
'***********************************************************
' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
'***********************************************************
Public Function ExcelInicio(psArchivo As String, _
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
ExcelInicio = True
Exit Function
ErrBegin:
  MsgBox TextErr(err.Description), vbInformation, "Aviso"
  ExcelInicio = False
End Function
'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Public Sub ExcelFin(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
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
   MsgBox TextErr(err.Description), vbInformation, "Aviso"
End Sub

Public Sub CuadroExcel(xlHoja1 As Excel.Worksheet, x1 As Integer, Y1 As Integer, x2 As Integer, Y2 As Integer, Optional lbLineasVert As Boolean = False)
Dim i, J As Integer

For i = x1 To x2
    xlHoja1.Range(xlHoja1.Cells(Y1, i), xlHoja1.Cells(Y1, i)).Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(Y2, i), xlHoja1.Cells(Y2, i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Next i
If lbLineasVert = False Then
    For i = x1 To x2
        For J = Y1 To Y2
            xlHoja1.Range(xlHoja1.Cells(J, i), xlHoja1.Cells(J, i)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Next J
    Next i
End If
If lbLineasVert Then
    For J = Y1 To Y2
        xlHoja1.Range(xlHoja1.Cells(J, x1), xlHoja1.Cells(J, x1)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Next J
End If
For J = Y1 To Y2
    xlHoja1.Range(xlHoja1.Cells(J, x2), xlHoja1.Cells(J, x2)).Borders(xlEdgeRight).LineStyle = xlContinuous
Next J
End Sub

Public Sub AbrirArchivo(lsArchivo As String, lsRutaArchivo As String)
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
'JOEP20171015 Flujo de Caja

'CTI1 20180719 ***
Public Function ReDimPreserve(MyArray As Variant, nNewFirstUBound As Long, nNewLastUBound As Long) As Variant
    Dim i, J As Long
    Dim nOldFirstUBound, nOldLastUBound, nOldFirstLBound, nOldLastLBound As Long
    Dim TempArray() As Variant 'Change this to "String" or any other data type if want it to work for arrays other than Variants. MsgBox UCase(TypeName(MyArray))
'---------------------------------------------------------------
'COMMENT THIS BLOCK OUT IF YOU CHANGE THE DATA TYPE OF TempArray
    If InStr(1, UCase(TypeName(MyArray)), "VARIANT") = 0 Then
        MsgBox "This function only works if your array is a Variant Data Type." & vbNewLine & _
               "You have two choice:" & vbNewLine & _
               " 1) Change your array to a Variant and try again." & vbNewLine & _
               " 2) Change the DataType of TempArray to match your array and comment the top block out of the function ReDimPreserve" _
                , vbCritical, "Invalid Array Data Type"
        End
    End If
'---------------------------------------------------------------
    ReDimPreserve = False
    'check if its in array first
    If Not IsArray(MyArray) Then MsgBox "You didn't pass the function an array.", vbCritical, "No Array Detected": End

    'get old lBound/uBound
    nOldFirstUBound = UBound(MyArray, 1): nOldLastUBound = UBound(MyArray, 2)
    nOldFirstLBound = LBound(MyArray, 1): nOldLastLBound = LBound(MyArray, 2)
    'create new array
    ReDim TempArray(nOldFirstLBound To nNewFirstUBound, nOldLastLBound To nNewLastUBound)
    'loop through first
    For i = LBound(MyArray, 1) To nNewFirstUBound
        For J = LBound(MyArray, 2) To nNewLastUBound
            'if its in range, then append to new array the same way
            If nOldFirstUBound >= i And nOldLastUBound >= J Then
                TempArray(i, J) = MyArray(i, J)
            End If
        Next
    Next
    'return the array redimmed
    If IsArray(TempArray) Then ReDimPreserve = TempArray
End Function
'CTI1 ************

'INICIO EAAS20181010 SEGÚN
Public Sub VerificarFechaSistemaAntesDelExtorno(ByRef obj As Object, Optional ByVal pbSalirSistema As Boolean = False)
    Dim oSis As New NConstSistemas
    Dim ldFechaSistema As Date

    ldFechaSistema = CDate(oSis.LeeConstSistema(gConstSistFechaSistema))
    Set oSis = Nothing

    If gdFecSis <> ldFechaSistema Then
        MsgBox "La Fecha de tu sesión en el SICMACM-NEGOCIO no coincide con la fecha del servidor. Debe volver a iniciar sesión.", vbInformation, "Aviso"
        If pbSalirSistema Then
            Call SalirSICMACMNegocio
            Unload obj
            End
        End If
    End If
End Sub
'FIN EAAS20181010
'***JGPA20190815 ACTA N° 106 - 2019
Public Function FormateaTexto(ByVal psCadena As String) As String
    Dim sCaracter As String
    sCaracter = "'"
    If Len(psCadena) > 0 Then
        If InStr(psCadena, sCaracter) > 0 Then
                FormateaTexto = Replace(psCadena, sCaracter, "")
        Else
            FormateaTexto = psCadena
        End If
    End If
End Function
'***End JGPA20190815
'JHCU ENCUESTA 16-10-2019
Public Sub Encuestas(ByRef sUser As String, ByRef sCodage As String, ByRef sCodEncuesta As String, ByRef sCodOpe As Variant)
    Dim oCons As COMDConstantes.DCOMConstantes
    Dim R As ADODB.Recordset
    Dim sCodOpeconv As String

    Dim nValor As Integer

    Dim cOpeConv As String
    On Error GoTo ErrFunction
    Set oCons = New COMDConstantes.DCOMConstantes
       sCodOpeconv = sCodOpe
    Set R = oCons.EncuestaPinPads(sUser, sCodage, sCodEncuesta, sCodOpeconv)

    If Not R.EOF Then
       nValor = R!ENCUESTA
    Else
       Exit Sub
    End If

    If nValor = 1 Then

        EncuestaVisto
        oCons.EncuestaPinPadsRes sUser, sCodage, sCodEncuesta, sCodOpeconv, res, cUsuVisto
        Set oCons = Nothing
    End If

    Exit Sub
ErrFunction:
        MsgBox err.Description, vbCritical
        If Not R Is Nothing Then
            Set R = Nothing
        End If

End Sub

Public Function MyMsgBox(Msg As String, Title As String, Optional Command1 As String = "Aceptar", Optional Command2 As String = "Cancelar") As Integer
    'Esto coloca el formulario en el centro de la pantalla
    frmMyMsgBox.Move Int(Screen.Width / 2) - Int(frmMyMsgBox.Width / 2), Int(Screen.Height / 2) - Int(frmMyMsgBox.Height / 2)

frmMyMsgBox.Caption = Title 'Se le asigna el título al Caption del formulario
    frmMyMsgBox.lblMessage.Caption = Msg 'Se le asigna el mensaje al lblMessage del formulario
    frmMyMsgBox.cmdOk.Caption = Command1 'Se le asigna el Caption del Command1 al cmdOk
    frmMyMsgBox.cmdCancel.Caption = Command2 'Se le asigna el Caption del Command1 al cmdOk

    frmMyMsgBox.Show vbModal 'Se muestra el formulario de formal Modal
    MyMsgBox = frmMyMsgBox.valor 'Se asigna al valor de retorno el botón pulsado
End Function
'SUBIR CADA VEZ QUE MODIFICA EL FORM DE MENSAJE DE ENCUESTAS
Public Sub EncuestaVisto()
  res = MyMsgBox("Operación requiere encuesta de satisfacción", "Encuesta de satisfacción", "Pedir Encuesta", "Omitir Encuesta") 'Llamada a la función
    If res = -1 Then 'Se comprueba el botón pulsado
        Dim oVisto As frmVistoElectronico
        Dim bResultadoVisto As Boolean
        Set oVisto = New frmVistoElectronico
        bResultadoVisto = oVisto.Inicio(6)
        If bResultadoVisto Then
         cUsuVisto = oVisto.ObtieneUsuarioVisto
        Else
        EncuestaVisto
        End If
'    Else
'        If Res = -2 Then
'             MsgBox "No existe conexión con el PinPad"
'        Else
'             MsgBox "El usuario te califico con"
'        End If
    End If
End Sub
'FIN JHCU


