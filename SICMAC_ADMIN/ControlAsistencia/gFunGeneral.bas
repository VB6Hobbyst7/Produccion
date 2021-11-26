Attribute VB_Name = "gFunGeneral"
''   Declaración de Arrays para conversión de Numeros a Letras
    Dim unidad(0 To 9) As String
    Dim decena(0 To 9) As String
    Dim centena(0 To 10) As String
    Dim deci(0 To 9) As String
    Dim otros(0 To 15) As String
    Dim lbSqlCnx As Boolean
    
'****************************************************
'GetCadenaConexion:Funcion que devuelve la cadena de conección de la
'agencia segun el codigo que se ingresa
'****************************************************
Public Function GetCadenaConexion(psCuenta As String, psNroServidor As String) As String
    Dim RegSistema As New ADODB.Recordset
    Dim sql1 As String
    
    On Error GoTo Error
    
    sql1 = "Select * From Servidor Where substring(cCodAge,4,2) = '" + Mid(psCuenta, 1, 2) + "' AND cEstado = 'A' And cNroSer = '" & psNroServidor & "'"
    
    RegSistema.Open sql1, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    
    If RegSistema.EOF And RegSistema.BOF Then
       RegSistema.Close
       Set RegSistema = Nothing
        MsgBox "No hay una conexion Activa con Agencia " & Mid(psCuenta, 1, 2), vbInformation, "Aviso"
        GetCadenaConexion = ""
        Exit Function
    End If
    
    GetCadenaConexion = "DSN=" + Trim(RegSistema!cDsn) + ";uid=" + Trim(RegSistema!cLogin) + ";pwd=" + Trim(RegSistema!cpass) + ";DATABASE=" + Trim(RegSistema!cDataBase) + ";SERVER=" + Trim(RegSistema!cNomSer)
    
    gsCodAgeN = RegSistema!cCodAge
    gsCodAgeNAdd = RegSistema!cCodAge
    
    RegSistema.Close
    Set RegSistema = Nothing

    Exit Function
Error:
    MsgBox "Error de Conexión " + Err.Description, vbCritical, "Servicio Tecnico"
    
End Function


'****************************************
'AbreConeccion:Función que devvuelve True si la se realiza la nueva conexión y false
'si la nueva conexion falla, se ingresa el codigo de agencia y False
'o los dos numeros iniciales del codigo de agencia y true
'****************************************

'**************************************************************
'* Retorna la Sentencia SQl que busca los datos del cliente
'* Dado su codigo
'* Uso :Formulario frmBuscacli
'**************************************************************
Public Function BuscaPersCodigo(psCodigo As String) As String
    Dim Cons As String
        Cons = "SELECT persona.cnompers, " _
                    & "persona.cdirpers, " _
                    & "persona.ccodpers, " _
                    & "persona.cNudoci, " _
                    & "persona.cNudotr, " _
                    & "persona.cTelpers, " _
                    & "persona.dFecNac, " _
                    & "persona.cCodZon " _
                    & "FROM persona " _
                    & "WHERE persona.cCodpers = '" & Trim(psCodigo) & "' " _
                    & "ORDER BY Persona.cNomPers"
        BuscaPersCodigo = Cons
End Function

'**********************************************
'Busca en la tabla el dato y se ubica en el reguistro
'**********************************************
'FECHA CREACION : 25/07/99  -   JPNZ
'MODIFICACION:
'****************************************
Function BuscaTablaCad(Record As ADODB.Recordset, DatoBuscado As String, Cad As String)
    Dim Ban As Boolean
    Ban = True
    
    If Not Record.BOF Then
        Record.MoveFirst
    End If
    
    Cad = Trim(Cad)
    While Not Record.EOF And Ban
        If Trim(Record.Fields(Cad)) = Trim(DatoBuscado) Then
            Ban = False
        Else
            Record.MoveNext
        End If
    Wend
    If Ban Then
        BuscaTablaCad = 0
    Else
        BuscaTablaCad = 1
    End If
    End Function

'**********************************************
' Busca en la tabla el dato y se ubica en el reguistro
' y devuelve el campo que se le pide en la segunda cadena
'**********************************************
'FECHA CREACION : 25/07/99  -   JPNZ
'MODIFICACION:
'****************************************
Public Function BuscaTablaCadDev(Record As ADODB.Recordset, DatoBuscado As String, Cad As String, CadDev As String) As String
    Dim Ban As Boolean
    Ban = True

    If Not Record.BOF Then
        Record.MoveFirst
    End If

    While Not Record.EOF And Ban
        If Trim(Record.Fields(Cad)) = Trim(DatoBuscado) Then
            Ban = False
        Else
            Record.MoveNext
        End If
    Wend
    If Ban Then
        BuscaTablaCadDev = ""
    Else
        BuscaTablaCadDev = IIf(IsNull(Record.Fields(CadDev)), "", Record.Fields(CadDev))
    End If
End Function


'**********************************************
'RUTINA VALIDA EL INGRESO DE NUMEROS DECIMALES
'**********************************************
'FECHA CREACION : 24/06/99  -   MAVF
'MODIFICACION:
'Agradecimiento a EJRS por su función inicial
'**********************************************
Public Function intfNumDec(cTexto As TextBox, intTecla As Integer, _
    Optional nLongitud As Integer = 8, Optional nDecimal As Integer = 2) As Integer
    Dim cValidar As String
    cCadena = cTexto
    cValidar = "-0123456789."
    
    If InStr(".", Chr(intTecla)) <> 0 Then
        If InStr(cCadena & Chr(intecla), ".") <> 0 Then
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
        'vNumEnt = Len(Left(cTexto, vPosPto - 1))
    End If
    If vPosPto > 0 Then
        'If cTexto.SelStart + 1 = vPosPto And intTecla = vbKeyDelete Then
        '    intTecla = 0
        '    Beep
        'End If
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
    intfNumDec = intTecla
End Function

'**********************************************
'RUTINA VALIDA EL INGRESO DE NUMEROS ENTEROS
'**********************************************
'FECHA CREACION : 24/06/99  -   MAVF
'MODIFICACION:
'Agradecimiento a EJRS por su función inicial
'**********************************************
Public Function intfNumEnt(intTecla As Integer) As Integer
Dim cValidar As String
    cValidar = "0123456789"
    If intTecla > 27 Then
        If InStr(cValidar, Chr(intTecla)) = 0 Then
            intTecla = 0
            Beep
        End If
    End If
    intfNumEnt = intTecla
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

'*********************************************************************
'RUTINA VALIDA EL INGRESO DE SOLO TEXTO Y EN MAYUSCULAS EN UN TEXTBOX
'*********************************************************************
'FECHA CREACION : 24/06/99  -   EJRS
'MODIFICACION:  11/07/99 - NSSE
'**********************************************
Public Function intfLetrasMay(intTecla As Integer) As Integer
If intTecla > 26 Then
    If Chr(intTecla) >= "a" And Chr(intTecla) <= "z" Then
        intTecla = intTecla - 32
    End If
    If Chr(intTecla) >= "A" And Chr(intTecla) <= "Z" Then
        intTecla = intTecla
    Else
        intTecla = intfLetras(intTecla)
    End If
    Cad = Chr(intTecla)
    ' Ñ , ñ ,
    If intTecla = 209 Or intTecla = 241 Or intTecla = 8 Or intTecla = 32 Then
        intfLetrasMay = Asc(UCase(Chr(intTecla)))
        Exit Function
    End If
End If
intfLetrasMay = intTecla
End Function

'***********************************************
'RUTINA QUE VALIDA EL INGRESO DE TEXTO SOLAMENTE EN UN TEXTCONTROL
'***********************************************
Private Function intfLetras(intTecla As Integer) As Integer
'    cValidar = "0123456789+:;'<>.?_=+[]{}|!@#$%^&()*"
    cValidar = "0123456789+:;'<>?_=+[]{}|!@#$%^()*"
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) <> 0 Then
            intTecla = 0
            Beep
        End If
    End If
    intfLetras = intTecla
End Function
'***********************************
'RUTINA DA EL ENFOQUE A UN CONTROL
'***********************************
'FECHA CREACION : 24/06/99  -   EJRS
'MODIFICACION:
'**********************************************
Public Sub fEnfoque(ctrControl As Control)
    ctrControl.SelStart = 0
    ctrControl.SelLength = Len(ctrControl.Text)
   ' ctrControl.SetFocus
End Sub
Public Function CentraSdi(frmCentra As Form) As Integer
    frmCentra.Move (Screen.Width - frmCentra.Width) / 2, (Screen.Height - frmCentra.Height) / 2, frmCentra.Width, frmCentra.Height
    CentraSdi = 1
End Function

'***************************************
'CENTRA CUALQUIER FORMULARIO MDICHILD
'***************************************
'FECHA CREACION : 24/06/99  -   EJRS
'MODIFICACION:
'****************************************
Public Function intfCentrado(frmCentra As Form) As Integer
    frmCentra.Move (Screen.Width - frmCentra.Width) / 2, (Screen.Height - frmCentra.Height) / 2 - 1000, frmCentra.Width, frmCentra.Height
    fCentrado = 1
End Function

'**************************************************************
'* Asigna los respectivos datos de la Person a los
'* Controles Text o Label de un Formulario.
'* Observacion :
'* En cada control del formulario asigne en su propieda TAG los valores del
'* Select Case segun el valor que dese Mostrar(TXTNOMBRE,TXTCODIGO,etc....)
'* Uso :Formulario frmBuscacli
'**************************************************************
'FECHA CREACION : 24/06/99  -   EJRS
'MODIFICACION:
'****************************************
Public Sub Asigna(frmFormulario As Form)
Dim ctrControl As Control
Dim strTag As String

For Each ctrControl In frmFormulario.Controls
    strTag = UCase(Trim(ctrControl.Tag))
    Select Case strTag
        Case "TXTNOMBRE"
            ctrControl = NomGrid
        Case "TXTDOCUMENTO"
            ctrControl = DNatGrid
        Case "TXTDIRECCION"
            ctrControl = DirGrid
        Case "TXTCODIGO"
            ctrControl = CodGrid
        Case "TXTCODIGO1"
            ctrControl = CodGrid
        Case "TXTTRIBUTARIO"
            ctrControl = DTriGrid
        Case "TXTTELEFONO"
            ctrControl = TelGrid
        Case "TXTNACIMIENTO"
            ctrControl = NacGrid
        Case "TXTZONA"
            ctrControl = ZonGrid
    End Select
Next
End Sub

'**********************************************************
'***  Retorna un adodb.RecordSet cargado dada la sentencia SQL
'**********************************************************
'FECHA CREACION : 24/06/99  -   EJRS
'MODIFICACION:
'****************************************

Public Function CargaRecord(sql As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
    
    If rs.State = adStateOpen Then
       rs.Close
    End If
    dbCmact.CommandTimeout = 90
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'    rs.Open sql, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText

    Set CargaRecord = rs
    Set rs = Nothing

End Function

'**************************************************************
'* Muestra el nombre de una Persona en su forma Real
'* por ejemplo Ingresa:
'*   Nombre Tabla:Perez/Castro\De Jimenes,Maria
'* Retornar:
'*   Maria Perez Castro de Jimenez
'**********************************************************
'FECHA CREACION : 24/06/99  -   EJRS
'MODIFICACION:
'****************************************
Public Function PstaNombre(psNombre As String, Optional psEstado As Boolean = True, Optional lbNombICC As Boolean = False) As String
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
    If lbNombICC = False Then
        If psEstado = True Then
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
        If Len(Trim(lsConyugue)) > 0 Then
            PstaNombre = Trim(lsConyugue) & " " & Trim(lsNombre) & " " & Trim(lsApellido) & " " & Trim(lsMaterno) & " DE "
        Else
            PstaNombre = Trim(lsNombre) & " " & Trim(lsApellido) & " " & Trim(lsMaterno)
        End If
    End If
Else
    PstaNombre = Trim(psNombre)
End If
End Function


'**************************************************************
'* Procedimiento que busca los datos de la tabla de codigos
'* Pasando como parametro el codigo de la tabla Codigos, el valor
'* a Buscar y los espacios definidos en el llenado del combo que desea
'* Retornar la Informacion
'**********************************************************
'FECHA CREACION : 24/06/99  -   EJRS
'MODIFICACION:
'****************************************
Public Function DatoTablaCodigo(lsCodTabla As String, lsValorBusc As String) As String
Dim sql As String
Dim rsB As New ADODB.Recordset

    'SQL = SQLTablaCod(Trim(lsCodTabla))
    Set rsB = CargaRecord(sql)
    If RSVacio(rsB) Then
        rsB.Close
        Set rsB = Nothing
    Else
        Do While Not rsB.EOF
            If Trim(rsB!cValor) = Trim(lsValorBusc) Then
                DatoTablaCodigo = Trim(rsB!cNomTab) & Space(100) & Trim(rsB!cValor)
                Exit Do
            Else
                DatoTablaCodigo = ""
            End If
            rsB.MoveNext
        Loop
        rsB.Close
        Set rsB = Nothing
    End If
End Function

'************************************************
'* Funcion:  Encripta / Desemcripta la Clave de Usuario
'************************************************
'FECHA CREACION : 25/07/99  -   VALV
'MODIFICACION:
'****************************************

Public Function Encripta(pnTexto As String, Valor As Boolean) As String
'true = encripta
'false = desencripta
Dim MiClase As cEncrypt
Set MiClase = New cEncrypt
Encripta = MiClase.ConvertirClave(pnTexto, , Valor)
End Function



'************************************************
'* Funcion:  Determina el ESTADO de un Crédito
'************************************************
'FECHA CREACION : 25/07/99  -   MAVF
'MODIFICACION:
'****************************************
Public Function Estado(vTabla As String, vCamCod As String, vCodigo As String) As String
    Dim RegTabla As New ADODB.Recordset
    sSql = "SELECT cestado FROM " & vTabla & " WHERE " & vCamCod & " = '" & vCodigo & "'"
    RegTabla.Open sSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    If RegTabla.EOF And RegTabla.BOF Then
        MsgBox " Estado de " & vCodigo & " no se encuentra ", vbInformation, " Aviso "
        Estado = ""
    Else
        Estado = RegTabla!cestado
    End If
    RegTabla.Close
End Function


'************************************************
'* Funcion:  Determina si un adodb.RecordSet esta Vacio
'************************************************
'FECHA CREACION : 24/06/99  -   FAOS
'MODIFICACION:
'****************************************
Public Function RSVacio(rs1 As ADODB.Recordset) As Boolean
 RSVacio = (rs1.BOF And rs1.EOF)
End Function

'*********************************************************************
'RUTINA QUE VALIDA EL INGRESO DE TEXTO EN MAYUSCULAS EN UN TEXTCONTROL
'**********************************************************************
Public Function intfMayusculas(intTecla As Integer) As Integer
 If Chr(intTecla) >= "a" And Chr(intTecla) <= "z" Then
    intTecla = intTecla - 32
 End If
 If intTecla = 39 Then
    intTecla = 0
 End If
 If intTecla = 209 Or intTecla = 241 Or intTecla = 8 Or intTecla = 32 Then
    intfMayusculas = Asc(UCase(Chr(intTecla)))
     Exit Function
 End If
 intfMayusculas = intTecla
End Function
'************************************************
'* Funcion:  Rellena hacia la izquierda de ceros
'*           un numero string
'************************************************
'FECHA CREACION : 24/06/99  -   FAOS
'MODIFICACION:
'************************************************
Public Function FillNum(intNumero As String, intLenNum As Integer, ChrFil As String) As String
  FillNum = Left(String(intLenNum, ChrFil), (Len(String(intLenNum, ChrFil)) - Len(Trim(intNumero)))) + Trim(intNumero)
End Function

'************************************************
'* Funcion:  Carga datos de una tabla de Codigos
'*           especifica a un adodb.RecordSet
'************************************************
'FECHA CREACION : 24/06/99  -   FAOS
'MODIFICACION:
'************************************************
Public Function LoadTabCod(CodTab As String) As ADODB.Recordset
 Dim qryTabCod As String
 Dim rs As New ADODB.Recordset
 If CodTab = "" Then
    Set LoadTabCod = Null
    Exit Function
 End If
 qryTabCod = ""
 qryTabCod = "SELECT cCodTab, cNomTab, cValor FROM TablaCod WHERE cCodTab LIKE '" & CodTab & "__'"
 rs.Open qryTabCod, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
 Set LoadTabCod = rs
 Set rs = Nothing
End Function

'************************************************
'* Funcion:  Rellena el Contenido de un adodb.RecordSet en
'*           un control ComboBox
'************************************************
'FECHA CREACION : 24/06/99  -   FAOS
'MODIFICACION:
'************************************************
Public Sub FillCboTab(cboBoxTmp As ComboBox, rs As ADODB.Recordset)
    
    With rs
        Do While Not .EOF
           cboBoxTmp.AddItem Left(!cNomTab, 40) & " " & Trim(!cCodTab) & " " & Trim(!cValor)
           .MoveNext
        Loop
    End With
    
End Sub

'***************************************************
'* Funcion:  Recupera el valor del campo cValorVar
'*           de Tabla VarSistema segun el ccodProd y
'*           cNomVar
'***************************************************
'FECHA CREACION : 24/06/99  -   FAOS
'MODIFICACION:
'***************************************************
Public Function ReadVarSis(txtCodPro As String, txtNomVar As String) As String
    Dim RecVar As New ADODB.Recordset
    Dim qryVar As String
    On Error GoTo Error
    qryVar = "SELECT cCodProd, cValorVar, cNomVar FROM VarSistema WHERE cCodProd = '" & txtCodPro & "' AND cNomVar = '" & txtNomVar & "'"
    RecVar.Open qryVar, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    If Not RecVar.EOF Then
      ReadVarSis = Trim(RecVar!cValorVar)
    Else
      MsgBox "No se encontro '" & txtNomVar & "' en Tabla de Variables del Sistema", vbCritical, "Error"
    End If
    RecVar.Close
    Set RecVar = Nothing
    Exit Function
Error:
    MsgBox "Error en Conexión + " + Err.Description, vbCritical, "Aviso"
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
'* Funcion que centra una cadena en una dimensión dada, si la cadena es
'* menor a esta dimensión la centra y rellena los espacios del caracter
'* espacio, caso contrario devuelve "No valido" y muestra un mensaje con el error
'***************************************************
'FECHA CREACION : 24/06/99  -   JPNZ
'MODIFICACION:
'***************************************************
Public Function CentrarCadena(psCadena As String, psNro As Long) As String
    Dim psNinf As Long
    Dim i As Long
    
    psCadena = Trim(psCadena)
    If Len(psCadena) > psNro Then
        MsgBox "EL valor de la Cadena enviada es mayor al espacio destinado", vbInformation, frmRealizaOpe.Caption
        CentrarCadena = "No valido"
    Else
        psNinf = Len(psCadena) / 2
        
        For i = 1 To (psNro / 2) - psNinf
            psCadena = " " + psCadena
        Next i
        
        For i = Len(psCadena) To psNro
            psCadena = psCadena + " "
        Next i
        CentrarCadena = psCadena
   End If

End Function

'***************************************************
'* GENERA EL DIGITO DE CHEQUEO DE LOS CODIGO DE CUENTA
'* UTILIZA EL MODULO 11
'***************************************************
'FECHA CREACION : 24/06/99  -   JPNZ
'MODIFICACION:
'***************************************************
Public Function GetDigitoChequeo(ByVal psCadena As String) As Integer
Dim liFactor As Variant
Dim liCadena(1 To 5) As Integer
Dim liSum, i As Integer
Dim lnDigito As Integer
liFactor = Array(6, 5, 4, 3, 2)
liCadena(1) = Val(Mid(psCadena, 1, 1))
liCadena(2) = Val(Mid(psCadena, 2, 1))
liCadena(3) = Val(Mid(psCadena, 3, 1))
liCadena(4) = Val(Mid(psCadena, 4, 1))
liCadena(5) = Val(Mid(psCadena, 5, 1))
liSum = 0
For i = 1 To 5
    liSum = liSum + liCadena(i) * liFactor(i - 1)
Next i
lnDigito = 11 - (liSum Mod 11)
If lnDigito = 10 Then
    GetDigitoChequeo = 0
ElseIf lnDigito = 11 Then
    GetDigitoChequeo = 1
Else
    GetDigitoChequeo = lnDigito
End If

End Function

'***************************************************
'* VERIFICA QUE SEA CORRECTO EL NRO DE CUENTA
'* VALIDANDO EL DIGITO DE CHEQUEO
'***************************************************
'FECHA CREACION : 24/06/99  -   JPNZ
'MODIFICACION:
'***************************************************
Public Function EsValido(ByVal psCadena As String) As Boolean
Dim liDigito As Integer
'Validar una Cadena con Módulo 11
liDigito = Val(Mid(psCadena, 6, 1))
If GetDigitoChequeo(psCadena) = liDigito Then
    EsValido = True
Else
    EsValido = False
End If
End Function

'***************************************************
'* CARGA UN adodb.RecordSet APARTIR DE UNA SENTENCIA SQL
'***************************************************
'FECHA CREACION : 24/06/99  -   CAFF
'MODIFICACION:
'***************************************************
Public Function GetObjeto(VSQL As String) As Object
Set MiReg = New ADODB.Recordset

MiReg.Open VSQL, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
If (MiReg.BOF And MiReg.EOF) Then
    Set MiReg = Nothing
Else
   Set GetObjeto = MiReg
End If

End Function


'***************************************************
'*  INICIA UN FLEXGRID CON NOMBRE DE SUS CABECERAS (5)
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'***************************************************

'Inicia un FlexGrid con nombre de sus cabeceras (5)
Public Sub FlexInicia(vFle As Object, ByVal vCol As Integer, ByVal vRow As Integer, _
    ByVal vHe1 As String, ByVal vHe2 As String, ByVal vHe3 As String, _
    ByVal vHe4 As String, ByVal vHe5 As String)
With vFle
    .Cols = vCol
    .Rows = vRow
    .ColWidth(0) = 2500
    .ColWidth(1) = 2500
    .ColWidth(2) = 1000
    .ColWidth(3) = 1250
    .ColWidth(4) = 1000
    .BandDisplay = flexBandDisplayHorizontal
    For X = 0 To vCol
        .ColHeader(X) = flexColHeaderOn
        .ColAlignmentHeader(X) = flexAlignCenterCenter
    Next X
    .ColHeaderCaption(0, 0) = vHe1
    .ColHeaderCaption(0, 1) = vHe2
    .ColHeaderCaption(0, 2) = vHe3
    .ColHeaderCaption(0, 3) = vHe4
    .ColHeaderCaption(0, 4) = vHe5
End With
'flexAviso.AddItem (Format(txtFecAviso, "dd/mm/yyyy") & vbTab & vNombre & vbTab & vDireccion & vbTab & !cCodcta)
End Sub


'***************************************************
'*  Indica la posisición donde de llena le FlexGrid
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'***************************************************
Function FlexIndice(vFle As Object, ByVal vFil As Integer, ByVal vCol As Integer) As Long
    FlexIndice = vFil * vFle.Cols + vCol
End Function

'***************************************************
'* Cambia cantidad de Filas de un FlexGrid
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'***************************************************
Public Sub FlexRediFila(vFle As Object, ByVal vFil As Integer)
    vFle.Rows = vFil
End Sub


'***************************************************
'* Carga un FlexGrid con los datos ingresados
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'***************************************************
Public Sub FlexCarga(vFle As Object, ByVal vFil As Integer, ByVal vCol As Integer, _
    ByVal cCadena)
    vFle.TextArray(FlexIndice(vFle, vFil, vCol)) = cCadena
End Sub

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
    'Print #ArcSal, Chr$(27) & Chr$(18);            'Desactiva condensado
    Print #ArcSal, Chr$(27) & Chr$(50);            'Espaciamiento lineas 1/6 pulg.1
    Print #ArcSal, Chr$(27) & Chr$(67) & Chr$(nLineas); '   Chr$(nLineas); 'Longitud de página a 66 líneas
    'Print #ArcSal, Chr$(27) & Chr$(87);            'Longitud de página a doble ancho
    If Not pbCondensado Then
'        Print #ArcSal, Chr$(27) & Chr$(108) & Chr$(0); 'Tipo letra : 0,1,2 - Roman,SansS,Courier
'        Print #ArcSal, Chr$(27) & Chr$(103);            'Tamaño  : 80, 77, 103
       Print #ArcSal, Chr$(27) & Chr$(107) & Chr$(2); 'Tipo letra : 0,1,2 - Roman,SansS,Courier
       Print #ArcSal, Chr$(27) & Chr$(77);            'Tamaño  : 80, 77, 103
    End If
    Print #ArcSal, Chr$(27) & Chr$(120) & Chr$(0);  'Draf : 1 pasada
    'Print #ArcSal, Chr$(27) & Chr$(150) '& "160"
   
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

'***************************************************
'* Valida que una fecha es correcta (entre 1900 - 9999)
'***************************************************
'FECHA CREACION : 11/07/99  -   MAVF
'MODIFICACION:
'***************************************************
Public Function ValFecha(lscontrol As Control) As Boolean
   If Mid(lscontrol, 1, 2) > 0 And Mid(lscontrol, 1, 2) <= 31 Then
        If Mid(lscontrol, 4, 2) > 0 And Mid(lscontrol, 4, 2) <= 12 Then
            If Mid(lscontrol, 7, 4) >= 1900 And Mid(lscontrol, 7, 4) <= 9999 Then
               If IsDate(lscontrol) = False Then
                    ValFecha = False
                    MsgBox "Formato de fecha no es válido", vbInformation, "Aviso"
                    lscontrol.SetFocus
                    Exit Function
               Else
                    ValFecha = True
               End If
            Else
                ValFecha = False
                MsgBox "Año de Fecha no es válido", vbInformation, "Aviso"
                lscontrol.SetFocus
                lscontrol.SelStart = 6
                lscontrol.SelLength = 4
                Exit Function
            End If
        Else
            ValFecha = False
            MsgBox "Mes de Fecha no es válido", vbInformation, "Aviso"
            lscontrol.SetFocus
            lscontrol.SelStart = 3
            lscontrol.SelLength = 2
            Exit Function
        End If
    Else
        ValFecha = False
        MsgBox "Dia de Fecha no es válido", vbInformation, "Aviso"
        lscontrol.SetFocus
        lscontrol.SelStart = 0
        lscontrol.SelLength = 2
        Exit Function
    End If
End Function

'***************************************************
'* Devuelve true si una fecha determinada se encuentra en la tabla feriados
'***************************************************
'FECHA CREACION : 11/07/99  -   CAFF
'MODIFICACION:
'***************************************************

Function VerSiFeriado(ByVal FecVer As String) As Boolean
Dim RegVerFec As New ADODB.Recordset
Dim Cad As String

Cad = ValidaFecha(FecVer)
If Cad = "" Then
    ' Determina si una fecha es feriado
    FecVer = Format(FecVer, "mm/dd/yyyy")
    
    VSQL = "select dFeriado from Feriado where dFeriado = '" & FecVer & "' "
    
    RegVerFec.Open VSQL, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    
    If RegVerFec.EOF And RegVerFec.BOF Then
        VerSiFeriado = False
    Else
        VerSiFeriado = True
    End If
    RegVerFec.Close
    Set RegVerFec = Nothing

Else
    MsgBox Cad, vbInformation, "Aviso"
    VerSiFeriado = False
End If

End Function

'***************************************************
'*  CABECERA PARA UNA PAGINA ENTERA
'***************************************************
'FECHA CREACION : 11/07/99  -   NSSE
'MODIFICACION:
'***************************************************
Sub LPT_PaginaEntera()
With frmimp
    vnegritaon = Chr$(27) + Chr$(69)
    vnegritaoff = Chr$(27) + Chr$(70)
    nFicSal = FreeFile
    Open mislpt For Output As nFicSal
    Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(65);  'Longitud de página a 22 líneas'
    Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
    Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(1);     'Tipo de Letra Sans Serif
    Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
    Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
End With
End Sub

'***************************************************
'* FUNCION QUE VERIFICA SI UNA FECHA ES VALIDA
'***************************************************
'FECHA CREACION : 11/07/99  -   NSSE
'MODIFICACION:
' si vacia ok
'***************************************************
Function ValidaFecha(cadfec As String) As String
Dim i As Integer
'validando longitud de fecha
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


'***************************************************
'* VALIDA LA HORA INGRESADA EN 23 HORAS, 59 SEGUNDOS
'***************************************************
'FECHA CREACION : 25/07/99  -   MAVF
'MODIFICACION:
'***************************************************
Public Function ValidaHora(lscontrol As Control) As Boolean
   If Mid(lscontrol, 1, 2) >= 0 And Mid(lscontrol, 1, 2) <= 23 Then
        If Mid(lscontrol, 4, 2) >= 0 And Mid(lscontrol, 4, 2) <= 59 Then
            ValidaHora = True
        Else
            ValidaHora = False
            MsgBox "Minuto no es válido", vbInformation, "Aviso"
            lscontrol.SetFocus
            lscontrol.SelStart = 3
            lscontrol.SelLength = 2
            Exit Function
        End If
    Else
        ValidaHora = False
        MsgBox "Hora no es válido", vbInformation, "Aviso"
        lscontrol.SetFocus
        lscontrol.SelStart = 0
        lscontrol.SelLength = 2
        Exit Function
    End If
End Function

'***************************************************
'* PROCEDIMIENTO QUE PERMITE SOLO EL INGRESO DE MAYUSCULAS Y NUMEROS
'***************************************************
'FECHA CREACION : 11/07/99  -   NSSE
'MODIFICACION:
' si vacia ok
'***************************************************
Sub SoloMayNum(KeyAscii As Integer)
    If KeyAscii < 48 Or Chr(KeyAscii) > 57 Then
            KeyAscii = intfLetrasMay(KeyAscii)
    End If
End Sub

'***************************************************
'* CONCATENA LA FECHA CON LA HORA ACTUAL
'***************************************************
'FECHA CREACION : 11/07/99  -   EJRS
'MODIFICACION:
'***************************************************
Public Function FechaHora(Psfecha As Date) As String
    FechaHora = Format(Psfecha & Space(1) & Time, "mm/dd/yyyy hh:mm:ss")
End Function

'   ------------------------------------------------------------
'   Función     :   LeeParam
'   Propósito   :   Recupera el valor del campo nValor1 de Tabla
'                   Parametro según el valor de ccodPar
'   Parámetro(s):   psCodPar -> Código del Parámetro
'   Creado      :   02/07/1999  -   FAOS
'   Modificado  :   02/07/1999  -   FAOS
'   ------------------------------------------------------------
'
Public Function LeeParam(psCodPar As String) As String

    Dim RecPar As New ADODB.Recordset
    Dim qryPar As String
    Dim lbVacio As Boolean
    
    qryPar = "SELECT nValor1, nValor2 FROM Parametro WHERE cCodPar = '" & psCodPar & "'"
    RecPar.Open qryPar, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    
    lbVacio = ((RecPar.BOF) Or (RecPar.BOF))
    LeeParam = IIf(lbVacio, 0, RecPar!nValor1)

    RecPar.Close
    Set RecPar = Nothing
    
End Function
'   ------------------------------------------------------------
'   Función     :   TituForm
'   Propósito   :   Retira el "&" del caption del menú de opcio
'                   nes que será el nuevo caption del formulario
'                   activo.
'   Parámetro(s):   psCapMen -> Captión de la Opción del Menú
'   Creado      :   02/07/1999  -   FAOS
'   Modificado  :   02/07/1999  -   FAOS
'   ------------------------------------------------------------
'   Form1.Caption = TituForm(frmMdiMain.mnuAhorrosProcEspApert.Caption)
Public Function TituForm(psCapMen As String) As String
    Dim i As Integer
    
    TituForm = ""
    For i = 1 To Len(psCapMen)
        If Mid$(psCapMen, i, 1) <> "&" Then
           TituForm = TituForm + Mid$(psCapMen, i, 1)
        End If
    Next i
End Function


'***************************************************
'* MUESTRA MENSAJE AVISO TEMPORAL
'***************************************************
'FECHA CREACION : 11/07/99  -   VALV
'MODIFICACION:
'***************************************************
Public Sub Mensaje(cadena As String, Valor As Variant)
frmMensaje.Show
frmMensaje.Mensaje.Caption = cadena
frmMensaje.Timer1.Interval = Valor
End Sub

' *************************************************************
'   Propósito   :   Recupera algunos datos del registro de la
'                   Tabla VarSistema
'   Creado      :   02/07/1999  -   FAOS
'   Modificado  :   02/07/1999  -   FAOS
' *************************************************************
'Public Sub CargaVarSis()
'    Dim lsQrySis As String
'    Dim rsQrySis As New ADODB.Recordset
'    lsQrySis = ""
'    lsQrySis = "SELECT cCodProd, cNomVar, cValorVar, cDescVar FROM VarSistema " _
'             & "WHERE cCodProd in ('ADM','AHO') AND cNomVar IN ('dFecSis','cCodAge','cNomCMAC','cDirBackup','cTitModulo') "
'    'If rsQrySis.State = adStateOpen Then rsQrySis.Close: Set rsQrySis = Nothing
'    rsQrySis.Open lsQrySis, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
'    If rsQrySis.BOF Or rsQrySis.EOF Then
'       rsQrySis.Close
'       Set rsQrySis = Nothing
'       MsgBox "Tabla VarSistema está vacia", vbInformation, "Aviso"
'       gsCodAge = ""
'       gsNomAge = ""
'       gdFecSis = ""
'       Exit Sub
'    End If
'    Do While Not rsQrySis.EOF
'       If Trim(rsQrySis!cNomVar) = "dFecSis" Then
'          gdFecSis = CDate(Trim(rsQrySis!cValorVar))
'       ElseIf Trim(rsQrySis!cNomVar) = "cCodAge" Then
'          gsCodAge = Trim(rsQrySis!cValorVar)
'          gsNomAge = Trim(rsQrySis!cDescVar)
'       ElseIf Trim(rsQrySis!cNomVar) = "cNomCMAC" Then
'          gsInstCmac = Trim(rsQrySis!cValorVar)
'          gsNomCmac = Trim(rsQrySis!cDescVar)
'       ElseIf Trim(rsQrySis!cNomVar) = "cDirBackup" Then
'          gsDirBackup = Trim(rsQrySis!cDescVar)
'       ElseIf Trim(rsQrySis!cNomVar) = "cTitModulo" Then
'         gcEmpresaLogo = Trim(rsQrySis!cValorVar)
'         gcTitModulo = Trim(rsQrySis!cValorVar) & " " & Trim(rsQrySis!cDescVar)
'       End If
'      rsQrySis.MoveNext
'    Loop
'    If gcTitModulo = "" Then
'      gcEmpresaLogo = "SICMACT"
'      gcTitModulo = "SICMACT Módulo Administrativo"
'   End If
'    rsQrySis.Close
'    Set rsQrySis = Nothing
'
'    'Deduce el nombre del Servidor
'    Dim lnPosIni As Integer, lnPosFin As Integer
'    Dim lnStr As String, lnStrConn
'    lnStrConn = dbCmact.ConnectionString
'    lnStrConn = lnStrConn & ";"
'    lnPosIni = InStr(1, lnStrConn, "SERVER=", vbTextCompare)
'    If lnPosIni <> 0 Then
'      lnPosFin = InStr(lnPosIni, lnStrConn, ";", vbTextCompare)
'      lnStr = Mid(lnStrConn, lnPosIni, lnPosFin - lnPosIni)
'      lnPosIni = InStr(1, lnStr, "=", vbTextCompare)
'      gsServerName = Mid(lnStr, lnPosIni + 1, Len(lnStr) - lnPosIni)
'      If Right(gsServerName, 1) = """" Then
'          gsServerName = Mid(gsServerName, 1, Len(gsServerName) - 1)
'      End If
'    Else
'      'Deduce el nombre del Servidor
'      Dim rsServ As New ADODB.Recordset
'      VSQL = "Select cNomSer = @@ServerName"
'      rsServ.Open VSQL, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'      If RSVacio(rsServ) Then
'          gsServerName = ""
'      Else
'          gsServerName = IIf(IsNull(rsServ!cNomSer), "", Trim(rsServ!cNomSer))
'      End If
'      rsServ.Close
'      Set rsServ = Nothing
'    End If
'
'    'Deduce el nombre de la Base de Datos
'    lnPosIni = InStr(1, lnStrConn, "DATABASE=", vbTextCompare)
'    lnPosFin = InStr(lnPosIni, lnStrConn, ";", vbTextCompare)
'    lnStr = Mid(lnStrConn, lnPosIni, lnPosFin - lnPosIni)
'    lnPosIni = InStr(1, lnStr, "=", vbTextCompare)
'    gsDBName = Mid(lnStr, lnPosIni + 1, Len(lnStr) - lnPosIni)
'    If Right(gsDBName, 1) = """" Then
'        gsDBName = Mid(gsDBName, 1, Len(gsDBName) - 1)
'    End If
'
'    'Deduce el nombre de usuario
'    lnPosIni = InStr(1, lnStrConn, "UID=", vbTextCompare)
'    lnPosFin = InStr(lnPosIni, lnStrConn, ";", vbTextCompare)
'    lnStr = Mid(lnStrConn, lnPosIni, lnPosFin - lnPosIni)
'    lnPosIni = InStr(1, lnStr, "=", vbTextCompare)
'    gsUID = Mid(lnStr, lnPosIni + 1, Len(lnStr) - lnPosIni)
'
'    'Deduce el password
'    lnPosIni = InStr(1, lnStrConn, "PWD=", vbTextCompare)
'    lnPosFin = InStr(lnPosIni, lnStrConn, ";", vbTextCompare)
'    lnStr = Mid(lnStrConn, lnPosIni, lnPosFin - lnPosIni)
'    lnPosIni = InStr(1, lnStr, "=", vbTextCompare)
'    gsPWD = Mid(lnStr, lnPosIni + 1, Len(lnStr) - lnPosIni)
'End Sub

'***************************************************
'* CARGA UN FLEX A PARTIR DE UNA SENTENCIA SQL
'***************************************************
'FECHA CREACION : 11/07/99  -   VALV
'MODIFICACION:
'***************************************************
Public Sub CargaFlex(sql As String, Flex As Control, foco As Boolean)
Dim rs As New ADODB.Recordset
rs.Open sql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
If RSVacio(rs) Then
   'MsgBox "No se Encontrarón Datos", vbInformation, "Aviso"
   Exit Sub
Else
Flex.Rows = rs.RecordCount + 1
Flex.Cols = rs.Fields.Count + 2
Flex.ColWidth(1) = 0
i = 1
Do While Not rs.EOF
        For j = 1 To rs.Fields.Count
            Flex.Row = i
            Flex.Col = j + 1
            If IsNull(rs(j - 1)) Then
                rs(j - 1) = 0
            End If
            Flex.Text = rs(j - 1)
        Next
         i = i + 1
         rs.MoveNext
Loop
        If foco = True Then
            Flex.SetFocus
        End If
End If
rs.Close
Set rs = Nothing
End Sub

'   ------------------------------------------------------------
'   Función     :   JustDerRep
'   Propósito   :   Justifica a la derecha un campo numérico en
'                   un reporte
'   Uso         :   Reportes en general
'   Parámetro(s):   pnCampos -> Importe
'                   pnLongit -> Longitud del Campo
'   Creado      :   02/07/1999  -   FAOS
'   Modificado  :   02/07/1999  -   FAOS
'   ------------------------------------------------------------
'
Public Function JustDerRep(pnCampos As String, pnLongit As Integer)
    JustDerRep = Format(Trim(pnCampos), String(pnLongit, "@"))
End Function

'Prepara una cadena especial (cadena con caracteres con tilde y/o otros)
' para que se imprima en el modo FREEFILE.
Public Function ImpreCarEsp(ByVal vCadena As String) As String
    vCadena = Replace(vCadena, "á", Chr(160), , , vbBinaryCompare)
    vCadena = Replace(vCadena, "É", Chr(144), , , vbBinaryCompare)
    vCadena = Replace(vCadena, "é", Chr(130), , , vbBinaryCompare)
    vCadena = Replace(vCadena, "í", Chr(161), , , vbBinaryCompare)
    vCadena = Replace(vCadena, "ó", Chr(162), , , vbBinaryCompare)
    vCadena = Replace(vCadena, "ú", Chr(163), , , vbBinaryCompare)
    vCadena = Replace(vCadena, "ñ", Chr(164), , , vbBinaryCompare)
    vCadena = Replace(vCadena, "Ñ", Chr(165), , , vbBinaryCompare)
    vCadena = Replace(vCadena, "°", Chr(248), , , vbBinaryCompare)
    vCadena = Replace(vCadena, "¦", Chr(179), , , vbBinaryCompare)
    ImpreCarEsp = vCadena
End Function

'   ------------------------------------------------------------
'   Función     :   CreaDevice
'   Propósito   :   Define device para efectuar la copia de segu
'                   ridad de la base de datos del servidor antes
'                   indicado
'   Uso         :   General
'   Parámetro(s):   Ninguno
'   Creado      :   01/07/1999  -   FAOS
'   Modificado  :   01/07/1999  -   FAOS
'   ------------------------------------------------------------
'
'Public Sub CreaDevice(psNomDev As String, psNomFis As String)
'
'       Dim ConBackUp As New ADODB.Connection
'       Dim lsConBackUp As String
'       Dim cmd As New ADODB.Command
'       Dim prm As New ADODB.Parameter
'
'       lsConBackUp = "DSN=DSNCmact;UID=sa;PWD=dba"
'       ConBackUp.CommandTimeout = 100
'       ConBackUp.Open lsConBackUp
'       cmd.CommandText = "sp_addumpdevice"
'       cmd.CommandType = adCmdStoredProc
'       cmd.Name = "spBUp"
'       Set prm = cmd.CreateParameter("MedioFis", adChar, adParamInput, 10)
'       cmd.Parameters.Append prm
'       Set prm = cmd.CreateParameter("NomLogico", adChar, adParamInput, 30)
'       cmd.Parameters.Append prm
'       Set prm = cmd.CreateParameter("NomFisico", adChar, adParamInput, 127)
'       cmd.Parameters.Append prm
'       Set cmd.ActiveConnection = ConBackUp
'       cmd.CommandTimeout = 720
'       cmd.Parameters.Refresh
'       ConBackUp.spBUp "Disk", psNomDev, psNomFis & ".DAT"
'       Set cmd = Nothing
'       ConBackUp.Close
'       Set ConBackUp = Nothing
'End Sub

'   ------------------------------------------------------------
'   Función     :   RealizaBackUp
'   Propósito   :   Efectua la copia de seguridad de la base de
'                   datos del servidor antes indicado
'   Uso         :   General
'   Parámetro(s):   Ninguno
'   Creado      :   01/07/1999  -   FAOS
'   Modificado  :   01/07/1999  -   FAOS
'   ------------------------------------------------------------
'Public Function RealizaBackup() As Boolean
'    Dim lsBackUp As String
'    Dim lsNomDev As String
'    Dim MyBackup As New SQLOLE.Backup           ' Crea el Objeto Back Up
'    lsNomDev = GetNomBacKUp(Trim(Format(gdFecSis, "dd/mm/yyyy")))
'    lsBackUp = gsDirBackup & "\" & lsNomDev
'    If Not ExisteDeviceBkup(lsNomDev) Then
'        CreaDevice lsNomDev, lsBackUp
'    End If
'    MyBackup.DumpDevices = lsNomDev               'Asigna nombre del device para Back Up
'    MyBackup.DumpInitDeviceBefore = 0
'    MousePointer = vbHourglass
'    On Error Resume Next
'    oServer.Databases(gsDBName).Dump MyBackup    ' Ejecuta el Back Up
'    If Err <> 0 Then                             ' Se produjo error(es) durante el back up
'          MsgBox Err.Description, 16, Err.Source & " Error", _
'          Err.HelpFile, Err.HelpContext
'          MousePointer = vbArrow
'          RealizaBackup = False
'          Exit Function
'    End If
'    On Error GoTo 0
'    RealizaBackup = True
'    MousePointer = vbArrow                       ' Proceso finalizado
'End Function
'
''   ------------------------------------------------------------
''   Función     :   ConectaBackUp
''   Propósito   :   Efectua la conexión a un Servidor específico
''                   y evalua los parametros
''   Uso         :   General
''   Parámetro(s):   psServer -> Nombre del Servidor
''                   psLogin  -> Nombre de Acceso del Usuario
''                   psClave  -> Clave de Acceso del Usuario
''   Creado      :   01/07/1999  -   FAOS
''   Modificado  :   01/07/1999  -   FAOS
''   ------------------------------------------------------------
'Public Function ConectaBackUp(psServer As String, psLogin As String, psClave As String) As Boolean
'
'    lbSqlCnx = False        ' No conectado aún
'    If psServer = "" Or psLogin = "" Or psClave = "" Then
'        MsgBox "Ingrese los Datos Correctos", vbInformation, "Aviso"
'        Exit Function
'    End If
'    On Error Resume Next
'    ' Conectando el Servidor
'    oServer.Connect ServerName:=psServer, _
'                        Login:=psLogin, _
'                        Password:=psClave
'    If Err <> 0 Then    ' Se detectaron errores en la conexión
'          MsgBox Err.Description, 16, Err.Source & " Error", _
'          Err.HelpFile, Err.HelpContext
'          MousePointer = vbArrow
'          ConectaBackUp = False
'          Exit Function
'    End If
'
'    On Error GoTo 0
'    lbSqlCnx = True
'    ConectaBackUp = True
'End Function

'   ------------------------------------------------------------
'   Función     :   DesconectaBackUp
'   Propósito   :   Efectua la desconexión a un Servidor específico
'                   luego de efectuar el proceso de Back Up
'   Uso         :   General
'   Parámetro(s):   Ninguno
'   Creado      :   01/07/1999  -   FAOS
'   Modificado  :   01/07/1999  -   FAOS
'   ------------------------------------------------------------
'
'Public Sub DesconectaBackUp()
'    If lbSqlCnx = True Then               ' Si esta Conectado
'       oServer.Disconnect                 ' efectua la desconexion
'       lbSqlCnx = False
'    End If
'End Sub

'Public Function ExisteDeviceBkup(psNomDev As String) As Boolean
'    Dim exito As Boolean
'    Dim Devices As Object
'    Set Devices = CreateObject("SQLOLE.Device")
'    exito = False
'    For Each Devices In oServer.Devices
'        If Devices.Name = psNomDev Then
'           exito = True
'           Exit For
'        End If
'    Next
'    ExisteDeviceBkup = exito
'End Function
'Verifica la corrceta habilitación de la impresora
Public Function ImpreSensa() As Boolean
Dim lbArchAbierto As Boolean
On Error GoTo ControlError
    frmMdiMain.staMain.Panels(2).Text = "Verificando Conexión con Impresora"
    ArcSal = FreeFile
    lbArchAbierto = True
    Open sLpt For Output As ArcSal
    Print #ArcSal, Chr$(27) & Chr$(64);            'Inicializa Impresora
    Close ArcSal
    lbArchAbierto = False
    ImpreSensa = True
    frmMdiMain.staMain.Panels(2).Text = ""
    Exit Function
ControlError:   ' Rutina de control de errores.
    frmMdiMain.staMain.Panels(2).Text = ""
    If lbArchAbierto Then
        Close ArcSal
    End If
    MsgBox "Impresora no Encontrada ó Inactiva" & vbCr & "Verifique que la Conexión sea Correcta", vbExclamation, "Aviso de Precaución"
    ImpreSensa = False
End Function

Public Function intfNumEntNeg(cTexto As TextBox, intTecla As Integer) As Integer
    Dim cValidar As String
    cCadena = cTexto
    cValidar = "0123456789-"
    
    If InStr("-", Chr(intTecla)) <> 0 Then
        If InStr(cCadena & Chr(intecla), "-") <> 0 Then
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
    intfNumEntNeg = intTecla
End Function


