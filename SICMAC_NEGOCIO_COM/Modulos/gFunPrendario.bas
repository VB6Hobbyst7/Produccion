Attribute VB_Name = "gFunPrendario"
'Modificacion de Bases: CASL 05.12.2000
'---------------------------------------'

'FUNCIONES PARA CALCULOS DEL CONTRATO DE DEUDAS DE UN PRESTAMO PRENDARIO
'*****************************************************
'FECHA CREACION : 03/07/99  -   LAYG
'MODIFICACION:
'**********************************************

Public Function CalculaInteresAdelantado(SaldoCapital As Double, _
    TasaInteres As Double, Plazo As Integer) As Double
CalculaInteresAdelantado = SaldoCapital * (1 - (1 / ((1 + TasaInteres) ^ (Plazo / 360))))
End Function

Public Function CalculaInteresMoratorio(SaldoCapital As Double, _
    TasaInteresMoratorio As Double, DiasAtraso As Double) As Double
CalculaInteresMoratorio = SaldoCapital * ((1 + TasaInteresMoratorio) ^ (DiasAtraso / 360) - 1)
End Function

Public Function CalculaCostoCustodia(ValorTasacion As Double, _
    TasaCustodia As Double, Plazo As Integer) As Double
CalculaCostoCustodia = ValorTasacion * TasaCustodia * (Plazo / 30)
End Function

Public Function CalculaCostoCustodiaMoratorio(ValorTasacion As Double, _
    TasaCustodiaMoratoria As Double, DiasAtraso As Double) As Double
CalculaCostoCustodiaMoratorio = ValorTasacion * (((1 + TasaCustodiaMoratoria) ^ (DiasAtraso / 30)) - 1)
End Function

Public Function CalculaCostoCustodiaDiferida(ValorTasacion As Double, _
    DiasTranscurridos As Double, PorcentajeCustodiaDiferida As Double, _
    IGV As Double) As Double
CalculaCostoCustodiaDiferida = Format((PorcentajeCustodiaDiferida / 30) * (1 + IGV) * ValorTasacion * (DiasTranscurridos - 30), "#0.00")
End Function

Public Function CalculaDeudaPrendario(lnSaldoCapital As Double, _
    ldFecVencimiento As Date, lnValorTasacion As Double, _
    lnTasaInteresVencido As Double, lnCostoCustodiaMoratoria As Double, _
    lnTasaImpuesto As Double, lsEstado As String, _
    lnCostoPreparacionRemate As Double, Optional ldFecParaDeuda As String) As Double
Dim vDiasAtra As Double
Dim vDeud As Double
Dim vInteMora As Double
Dim vImpu As Double
Dim vCostCustMora As Double
If Len(ldFecParaDeuda) <= 0 Then ldFecParaDeuda = gdFecSis
vDiasAtra = IIf(DateDiff("d", ldFecVencimiento, ldFecParaDeuda) <= 0, 0, DateDiff("d", ldFecVencimiento, ldFecParaDeuda))
If vDiasAtra = 0 Then
   vInteMora = 0
   vCostCustMora = 0
   vImpu = 0
Else
   vInteMora = CalculaInteresMoratorio(lnSaldoCapital, lnTasaInteresVencido, vDiasAtra)
   vInteMora = Round(vInteMora, 2)
   vCostCustMora = CalculaCostoCustodiaMoratorio(lnValorTasacion, lnCostoCustodiaMoratoria, vDiasAtra)
   vCostCustMora = Round(vCostCustMora, 2)
   vImpu = (vInteMora + vCostCustMora) * lnTasaImpuesto
   vImpu = Round(vImpu, 2)
End If
vDeud = Round(lnSaldoCapital, 2) + vInteMora + vCostCustMora + vImpu
If lsEstado = "6" Or lsEstado = "8" Then  ' Si esta en via de Remate
   vDeud = vDeud + Round((lnCostoPreparacionRemate * lnValorTasacion), 2)
End If
CalculaDeudaPrendario = vDeud
End Function

Public Function CalculaPrecioVentaRemate(vValor As Double) As Double
If (vValor - Int(vValor)) = 0 Then  ' Si es entero
    If (vValor Mod 5) = 0 Then
        CalculaPrecioVentaRemate = vValor
    Else
        CalculaPrecioVentaRemate = (Int(vValor / 5) * 5) + 5
    End If
Else    'Si no es entero
    CalculaPrecioVentaRemate = (Int(vValor / 5) * 5) + 5
End If
End Function

Public Function RedondeaPrecio(pValor As Currency) As Currency
If (pValor - Int(pValor)) = 0 Then  ' Si es entero
    If (pValor Mod 5) = 0 Then
        RedondeaPrecio = pValor
    Else
        RedondeaPrecio = (Int(pValor / 5) * 5) + 5
    End If
Else    'Si no es entero
    RedondeaPrecio = (Int(pValor / 5) * 5) + 5
End If
End Function

'***********************************************************************************
'FUNCIONES Y PROCEDIMIENTOS PARA MANEJO DE LOS PROCESOS
'Habilita los valores (check) de una lista, al parametro indicado
Public Sub OpcListado(ByRef pLista As ListBox, ByVal pValor As Boolean)
Dim X As Byte
For X = 0 To pLista.ListCount - 1
    pLista.Selected(X) = pValor
Next X
End Sub



'Carga el Código de Contrato Antiguo si tuviera
Public Function ContratoAntiguo(ByVal pCodCtaNue As String) As String
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    tmpSql = " SELECT cCodAnt FROM RelconNueAntPrend WHERE cCodNue = '" & pCodCtaNue & "'"
    tmpReg.Open tmpSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    If (tmpReg.BOF Or tmpReg.EOF) Then
        ContratoAntiguo = ""
    Else
        ContratoAntiguo = Trim(tmpReg!cCodAnt)
    End If
    tmpReg.Close
    Set tmpReg = Nothing
End Function

'Indica el Estado del Contrato
Public Function ContratoEstado(ByVal pnEstado As Integer) As String
'Falta modificar los estados
psEstado = Trim(psEstado)
If Len(psEstado) > 0 Then
    ContratoEstado = Switch( _
        psEstado = gPigEstReg, "Emitido ", psEstado = gPigEstDesemb, "Desembolsado ", _
        psEstado = gPigEstCanc, "Cancelado ", psEstado = gPigEstEntJoya, "Rescatado ", _
        psEstado = gPigEstVenc, "Vencida ", psEstado = gPigEstRemate, "Rematada ", _
        psEstado = gPigEstRemate, "Para Remate ", psEstado = gPigEstRenov, "Renovado ", _
        psEstado = gPigEstAdjud, "Adjudicado ", psEstado = gPigEstSubast, "Subastado ", _
        psEstado = gPigEstAnulNoDesemb, "Anulado ", psEstado = gPigEstBarras, "En Barras ", _
        psEstado = gPigEstAnulNoDesemb, "Anulado por No Desembolso ")
Else
    ContratoEstado = ""
End If
End Function

'Señala el Estado del documento por su Código de Transacción
Public Function ContratoOperacion(ByVal vCodigoOperacion As String) As String
vCodigoOperacion = Trim(vCodigoOperacion)
If Len(vCodigoOperacion) > 0 Then
    ContratoOperacion = Switch( _
        vCodigoOperacion = gsRegContrato, "Emisión ", _
        vCodigoOperacion = gsDesPrestamo, "Desembolso ", _
        vCodigoOperacion = gsCanNorPrestamo, "Cancelación Normal ", vCodigoOperacion = gsCanNorEnOtCjPrestamo, "Cancelación Normal EO CMAC", _
        vCodigoOperacion = gsCanMorPrestamo, "Cancelación Morosa ", vCodigoOperacion = gsCanMorEnOtCjPrestamo, "Cancelación Morosa EO CMAC", _
        vCodigoOperacion = gsRenPrestamo, "Renovación ", _
        vCodigoOperacion = gsAnuContrato, "Anulación ", _
        vCodigoOperacion = gsImpDuplicado, "Duplicado ", vCodigoOperacion = gsImpDuplicadoEOA, "Duplicado ", _
        vCodigoOperacion = gsDevJoyas, "Rescate (Entrega de Joya) ", vCodigoOperacion = gsDevJoyasEOA, "Rescate (Entrega de Joya) ", _
        vCodigoOperacion = gsModContrato, "Modificación de la Descripción ", _
        vCodigoOperacion = gsPagSobrante, "Pago de Sobrante ", _
        vCodigoOperacion = gsRemContrato, "Ingreso a Remate ", _
        vCodigoOperacion = gsVtaRemate, "Venta en Remate ", _
        vCodigoOperacion = gsAdjudica, "Adjudicación ", vCodigoOperacion = gsRenEnOtCj, "Renovacion en CMAC ", _
        vCodigoOperacion = gsVtaSubasta, "Venta en Subasta ", _
        vCodigoOperacion = gsCobCusDiferida, "Cobro Custodia ", vCodigoOperacion = gsCobCusDiferidaEOA, "Cobro Custodia EOA", _
        vCodigoOperacion = gsRenEnOtAg, "Renovado en otra Agencia ", _
        vCodigoOperacion = gsCanNorEnOtAgPrestamo, "Cancelación Normal en otra Agencia ", _
        vCodigoOperacion = gsCanMorEnOtAgPrestamo, "Cancelación Morosa en otra Agencia ", _
        vCodigoOperacion = gsRenDeOtAg, "Renovado de otra Agencia ", _
        vCodigoOperacion = gsCanNorDeOtAg, "Cancelación de otra Agencia ", _
        vCodigoOperacion = gsCanMorDeOtAg, "Cancelación Morosa de otra Agencia ", _
        vCodigoOperacion = gsExtCanPrestamo, "Extorno de Cancelación ", _
        vCodigoOperacion = gsExtRenPrestamo, "Extorno de Renovación ", _
        vCodigoOperacion = gsExtDevJoyas, "Extorno de Rescate (Entrega de Joyas) ")
Else
    ContratoOperacion = ""
End If
End Function

'Indica el tipo de Documento Civil
Public Function TipoDoCi(ByVal vTipoDocuCivil As String) As String
vTipoDocuCivil = Trim(vTipoDocuCivil)
If Len(Trim(vTipoDocuCivil)) > 0 Then
    TipoDoCi = Switch( _
        vTipoDocuCivil = "1", "LE", vTipoDocuCivil = "2", "CE", _
        vTipoDocuCivil = "3", "CFP", vTipoDocuCivil = "4", "CFA", _
        vTipoDocuCivil = "5", "PAS", vTipoDocuCivil = "6", "PN", _
        vTipoDocuCivil = "7", "BM", vTipoDocuCivil = "8", "LM", _
        vTipoDocuCivil = "9", "NI")
Else
    TipoDoCi = ""
End If
End Function

'Indica el tipo de Documento Tributario
Public Function TipoDoTr(ByVal vTipoDocuTributario As String) As String
vTipoDocuTributario = Trim(vTipoDocuTributario)
If Len(Trim(vTipoDocuTributario)) > 0 Then
    TipoDoTr = Switch( _
    vTipoDocuTributario = "1", "LT", vTipoDocuTributario = "2", "RUC", _
    vTipoDocuTributario = "3", "RUS", vTipoDocuTributario = "4", "NI")
Else
    TipoDoTr = ""
End If
End Function

'Rutina para quebrar un texto multiline en una impresión
Public Function QuiebreTexto(ByVal vTexto As String, ByVal vFila As Byte) As String
Dim vLinea As String
Dim X As Integer
Dim paso As Integer
paso = 0
vLinea = ""
'MsgBox vTexto
For X = 1 To Len(vTexto)
    If Asc(Mid(vTexto, X, 1)) = 13 Or X = Len(vTexto) Then
        paso = paso + 1
        If paso = vFila Then
            If Len(Trim(vLinea)) < 2 Then
                QuiebreTexto = " " 'IIf(paso = 1, vLinea, Right(vLinea, Len(vLinea)))
            Else
                QuiebreTexto = IIf(paso = 1, vLinea, Right(vLinea, Len(vLinea) - 2))
            End If
            Exit Function
        End If
        vLinea = ""
    End If
    vLinea = vLinea & Mid(vTexto, X, 1)
Next X
End Function

' Mostrar datos del cliente en el ListView indicado
'   se envia el ListView por REFERENCIA
Public Function MostrarCliente(lstCliente As ListView, ByVal prPers As ADODB.Recordset) As Boolean
Dim RegPersona As New ADODB.Recordset
Dim lstTmpCliente As ListItem
    ssql = "SELECT Persona.ccodpers, Persona.cnompers, Persona.cdirpers, Persona.ctelpers, " & _
        " Persona.ccodzon, Persona.cTipPers, Persona.ctidoci, Persona.cnudoci, Persona.ctidotr, Persona.cnudotr " & _
        " FROM PersCuenta INNER JOIN " & gcCentralPers & "Persona Persona ON Persona.cCodPers = PersCuenta.cCodPers " & _
        " WHERE PersCuenta.cRelaCta = 'TI' and PersCuenta.cCodCta = '" & vContrato & "'"
    RegPersona.Open ssql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (RegPersona.BOF Or RegPersona.EOF) Then
        RegPersona.Close
        Set RegPersona = Nothing
        MsgBox " Error al mostrar datos del cliente ", vbCritical, " Aviso "
        MostrarCliente = False
    Else
        lstCliente.ListItems.Clear
        Do While Not RegPersona.EOF
            ' Verifica si Documento de Identidad esta correcto
            If fgVerIdentificacionPers(RegPersona!cTipPers, RegPersona!ctidoci, IIf(IsNull(RegPersona!cnudoci), "null", RegPersona!cnudoci)) = False Then
                MsgBox "Por favor, Actualizar el Documento de Identidad de " & vbCr & Trim(RegPersona!cNomPers) & vbCr & "Consulte al Administrador o Asistente", vbInformation, "Aviso"
                MostrarCliente = False
                Exit Function
            End If
            
            Set lstTmpCliente = lstCliente.ListItems.Add(, , Trim(RegPersona!cCodPers))
                lstTmpCliente.SubItems(1) = Trim(PstaNombre(RegPersona!cNomPers, False))
                lstTmpCliente.SubItems(2) = Trim(RegPersona!cDirPers)
                lstTmpCliente.SubItems(3) = Trim(RegPersona!cTelPers & "")
                lstTmpCliente.SubItems(4) = Trim(ClienteZona(RegPersona!cCodZon))
                lstTmpCliente.SubItems(5) = Trim(ClienteCiudad(RegPersona!cCodZon))
                lstTmpCliente.SubItems(6) = TipoDoCi(RegPersona!ctidoci & "")
                lstTmpCliente.SubItems(7) = Trim(RegPersona!cnudoci & "")
                lstTmpCliente.SubItems(8) = TipoDoTr(RegPersona!cTidoTr & "")
                lstTmpCliente.SubItems(9) = Trim(RegPersona!cNudoTr & "")
            RegPersona.MoveNext
        Loop
        RegPersona.Close
        Set RegPersona = Nothing
        MostrarCliente = True
    End If
End Function

'Nombre de un Cliente
Public Function ClienteNombre(ByVal pContrato As String, ByVal pConexion As ADODB.Connection) As String
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    tmpSql = " SELECT p.cNomPers FROM " & gcCentralPers & "Persona AS P, PersCuenta AS PC" & _
        " WHERE p.ccodpers = pc.ccodpers AND pc.ccodcta = '" & pContrato & "'"
    tmpReg.Open tmpSql, pConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (tmpReg.BOF Or tmpReg.EOF) Then
        MsgBox " No existe la Persona ", vbInformation, " Aviso "
        ClienteNombre = ""
    Else
        ClienteNombre = Trim(tmpReg!cNomPers)
    End If
    tmpReg.Close
    Set tmpReg = Nothing
End Function

Public Function ClienteDocum(ByVal vContrato As String, ByVal pConexion As ADODB.Connection) As String
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    tmpSql = " SELECT p.cnudoci, p.cnudotr FROM " & gcCentralPers & "Persona AS P, PersCuenta AS PC" & _
        " WHERE p.ccodpers = pc.ccodpers AND pc.ccodcta = '" & vContrato & "'"
    tmpReg.Open tmpSql, pConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (tmpReg.BOF Or tmpReg.EOF) Then
        MsgBox " No existe la Persona, para obtener documento", vbInformation, " Aviso "
        ClienteDocum = ""
    Else
        ClienteDocum = Trim(tmpReg!cnudoci) & " " & Trim(tmpReg!cNudoTr)
    End If
    tmpReg.Close
    Set tmpReg = Nothing
End Function

'Zona de un cliente
Public Function ClienteZona(ByVal vCodigoZona As String) As String
Dim RegZona As New ADODB.Recordset
Dim vNombreZona As String
    ssql = " SELECT cDesZon FROM " & gcCentralCom & "Zonas where ccodzon = '" & vCodigoZona & "'"
    RegZona.Open ssql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    If (RegZona.BOF Or RegZona.EOF) Then
        'MsgBox " No existe la zona ", vbInformation, " Aviso "
        vNombreZona = ""
    Else
        vNombreZona = Trim(RegZona!cDesZon)
    End If
    RegZona.Close
    Set RegZona = Nothing
    ClienteZona = vNombreZona
End Function

'Ciudad de un Cliente
Public Function ClienteCiudad(ByVal vCodigoZona As String) As String
Dim RegZona As New ADODB.Recordset
Dim vNombreCiudad As String
    ssql = " SELECT cDesZon FROM " & gcCentralCom & "Zonas where ccodzon IN ('" & "1" & Mid(vCodigoZona, 2, 2) & "000000000" & "','" & "2" & Mid(vCodigoZona, 2, 4) & "0000000" & "')"
    RegZona.Open ssql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    If (RegZona.BOF Or RegZona.EOF) Then
        'MsgBox " No existe la Ciudad ", vbInformation, " Aviso "
        vNombreCiudad = ""
    Else
        Do While Not RegZona.EOF
            vNombreCiudad = vNombreCiudad & " " & Trim(RegZona!cDesZon)
            RegZona.MoveNext
        Loop
    End If
    RegZona.Close
    Set RegZona = Nothing
    ClienteCiudad = vNombreCiudad
End Function

Public Function ClienteCodigoZona(ByVal pContrato As String) As String
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    tmpSql = " SELECT p.cCodZon FROM " & gcCentralPers & "Persona AS P, PersCuenta AS PC" & _
        " WHERE p.ccodpers = pc.ccodpers AND pc.ccodcta = '" & pContrato & "' AND pc.cRelaCta = 'TI'"
    tmpReg.Open tmpSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    If (tmpReg.BOF Or tmpReg.EOF) Then
        MsgBox " No existe la Persona ", vbInformation, " Aviso "
        ClienteCodigo = ""
    Else
        ClienteCodigo = tmpReg!cCodZon
    End If
    tmpReg.Close
    Set tmpReg = Nothing
End Function

'Numero del último duplicado de un contrato
Public Function NumUltDuplic(pCodCta As String, Optional pLocal As Boolean = True) As Integer
Dim RegCP As New ADODB.Recordset
Dim tmpSql As String
Dim vConexion As ADODB.Connection
If pLocal Then
    Set vConexion = dbCmact
Else
    Set vConexion = dbCmactN
End If
tmpSql = "SELECT nnroduplic FROM credprenda  WHERE ccodcta = '" & pCodCta & "'"
RegCP.Open tmpSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
If (RegCP.BOF Or RegCP.EOF) Then
    NumUltDuplic = 0
Else
    NumUltDuplic = IIf(IsNull(RegCP!nNroDuplic) = True, 0, RegCP!nNroDuplic)
End If
RegCP.Close
Set RegCP = Nothing
End Function

'Numero de última transacción de un contrato en TransPrenda
Public Function NumUltTransac(pCodCta As String, Optional pLocal As Boolean = True) As Integer
Dim RegCP As New ADODB.Recordset
Dim tmpSql As String
Dim vConexion As ADODB.Connection
If pLocal Then
    Set vConexion = dbCmact
Else
    Set vConexion = dbCmactN
End If
tmpSql = " SELECT nNumTran FROM CredPrenda where ccodcta = '" & pCodCta & "' ORDER BY nnumtran DESC"
RegCP.Open tmpSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
If (RegCP.BOF Or RegCP.EOF) Then
    NumUltTransac = 0
Else
    NumUltTransac = RegCP!nNumTran
End If
RegCP.Close
Set RegCP = Nothing
End Function

'Numero de última renovación de un contrato
Public Function NumUltRenov(pCodCta As String, Optional pLocal As Boolean = True) As Integer
Dim RegCP As New ADODB.Recordset
Dim tmpSql As String
Dim vConexion As ADODB.Connection
If pLocal Then
    Set vConexion = dbCmact
Else
    Set vConexion = dbCmactN
End If
tmpSql = "SELECT nnumrenov FROM credprenda  WHERE ccodcta = '" & pCodCta & "'"
RegCP.Open tmpSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
If (RegCP.BOF Or RegCP.EOF) Then
    NumUltRenov = 0
Else
    NumUltRenov = RegCP!nnumrenov
End If
RegCP.Close
Set RegCP = Nothing
End Function

'Procedimiento de verificación que el numero de documento
' no se encuentre duplicado
Public Function ExisBolVtaRem(ByVal pNroDoc As String) As Boolean
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    'Verifica la existencia en DetRemate
    tmpSql = " SELECT cnrodocum FROM detremate where cnrodocum = '" & pNroDoc & "' and cestado not in ('X')"
    tmpReg.Open tmpSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    If (tmpReg.BOF Or tmpReg.EOF) Then
        tmpReg.Close
        Set tmpReg = Nothing
        'Verifica la existencia en RemSubRemota
        tmpSql = " SELECT cnrodocum FROM RemSubRemota " & _
            " WHERE cnrodocum = '" & pNroDoc & "' AND ctipo = 'R' and cestado not in ('X')"
        tmpReg.Open tmpSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
        If (tmpReg.BOF Or tmpReg.EOF) Then
            ExisBolVtaRem = False
        Else
            ExisBolVtaRem = True
        End If
    Else
        ExisBolVtaRem = True
    End If
    tmpReg.Close
    Set tmpReg = Nothing
End Function
Public Function ExisBolVtaSub(ByVal pNroDoc As String) As Boolean
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    'Verifica la existencia en DetRemate
    tmpSql = " SELECT cnrodocum FROM detSubas where cnrodocum = '" & pNroDoc & "'"
    tmpReg.Open tmpSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    If (tmpReg.BOF Or tmpReg.EOF) Then
        tmpReg.Close
        Set tmpReg = Nothing
        'Verifica la existencia en RemSubRemota
        tmpSql = " SELECT cnrodocum FROM RemSubRemota " & _
            " WHERE cnrodocum = '" & pNroDoc & "' AND ctipo = 'S'"
        tmpReg.Open tmpSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
        If (tmpReg.BOF Or tmpReg.EOF) Then
            ExisBolVtaSub = False
        Else
            ExisBolVtaSub = True
        End If
    Else
        ExisBolVtaSub = True
    End If
    tmpReg.Close
    Set tmpReg = Nothing
End Function


'Permite obtener la suma de las operaciones
'de la tabla de TranDiaria de acuerdo a la condición deseada
Public Function DiaSumOpe(ByVal pCodOpe As String, Optional ByVal pCodOpe2 As String = "000000", _
    Optional ByVal pCodOpe3 As String = "000000", Optional ByVal pCodOpe4 As String = "000000", _
    Optional ByVal pCodOpe5 As String = "000000", Optional ByVal pCodOpe6 As String = "000000", _
    Optional ByVal pCodOpe7 As String = "000000", Optional ByVal pCodOpe8 As String = "000000") As Currency
Dim RegProc As New ADODB.Recordset
Dim tmpSql As String
tmpSql = "SELECT sum(abs(nMonTran)) AS Suma FROM TranDiaria " & _
    " WHERE ccodope IN ('" & pCodOpe & "','" & pCodOpe2 & "','" & pCodOpe3 & "','" & pCodOpe4 & "','" & pCodOpe5 & "','" & pCodOpe6 & "','" & pCodOpe7 & "','" & pCodOpe8 & "')" & _
    " AND cflag IS NULL "
    'If pbListGene = False Then      'Verifica el listado (General o no)
    '    tmpSql = tmpSql & " AND ccodusu = '" & gsCodUser & "' "
    'ElseIf cboUsuario <> "<Consolidado>" Then
    '    tmpSql = tmpSql & " AND ccodusu = '" & Trim(cboUsuario.Text) & "' "
    'End If
RegProc.Open tmpSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
If (RegProc.BOF Or RegProc.EOF) Then
    DiaSumOpe = Format(0, "#0.00")
Else
    DiaSumOpe = IIf(IsNull(RegProc!Suma) = True, 0, Format(RegProc!Suma, "#0.00"))
End If
RegProc.Close
Set RegProc = Nothing
End Function

'Sumas generales de saldos y creditos para la estadística
Public Function SalNroEst(ByVal pBoveda As String, pConexPrend As ADODB.Connection, ByVal pEst1 As String, _
    Optional ByVal pEst2 As String = "", Optional ByVal pEst3 As String = "", _
    Optional ByVal pEst4 As String = "", Optional ByVal pEst5 As String = "", _
    Optional ByVal pEst6 As String = "", Optional ByVal pEst7 As String = "") As Currency
Dim tmpReg As New ADODB.Recordset
Dim tmpSql As String
tmpSql = "SELECT Count(cCodCta) AS Nume FROM CredPrenda " & _
    " WHERE cestado IN ('" & pEst1 & "','" & pEst2 & "','" & pEst3 & "','" & pEst4 & "'," & _
    " '" & pEst5 & "','" & pEst6 & "','" & pEst7 & "') "
If Len(Trim(pBoveda)) > 0 Then  ' Boveda
    tmpSql = tmpSql & " AND cAgeBoveda in " & pBoveda & " "
End If
tmpReg.Open tmpSql, pConexPrend, adOpenStatic, adLockOptimistic, adCmdText
If (tmpReg.BOF Or tmpReg.EOF) Then
    SalNroEst = Format(0, "#0")
Else
    SalNroEst = IIf(IsNull(tmpReg!nume) = True, 0, Format(tmpReg!nume, "#0"))
End If
tmpReg.Close
Set tmpReg = Nothing
End Function

Public Function SalSumEst(ByVal pBoveda As String, pConexPrend As ADODB.Connection, ByVal pNomReg As String, ByVal pEst1 As String, _
    Optional ByVal pEst2 As String = "", Optional ByVal pEst3 As String = "", _
    Optional ByVal pEst4 As String = "", Optional ByVal pEst5 As String = "", _
    Optional ByVal pEst6 As String = "", Optional ByVal pEst7 As String = "") As Currency
Dim tmpReg As New ADODB.Recordset
Dim tmpSql As String
tmpSql = "SELECT sum(abs(" & pNomReg & ")) AS Suma FROM CredPrenda " & _
    " WHERE cestado IN ('" & pEst1 & "','" & pEst2 & "','" & pEst3 & "','" & pEst4 & "'," & _
    " '" & pEst5 & "','" & pEst6 & "','" & pEst7 & "')"
If Len(Trim(pBoveda)) > 0 Then
    tmpSql = tmpSql & " AND cAgeBoveda in " & pBoveda & " "
End If
tmpReg.Open tmpSql, pConexPrend, adOpenStatic, adLockOptimistic, adCmdText
If (tmpReg.BOF Or tmpReg.EOF) Then
    SalSumEst = Format(0, "#0.00")
Else
    SalSumEst = IIf(IsNull(tmpReg!Suma) = True, 0, Format(tmpReg!Suma, "#0.00"))
End If
tmpReg.Close
Set tmpReg = Nothing
End Function



'Impresión para Contratos y Boletas
Public Sub ImpreBegChe(pbCondensado As Boolean, nLineas As Integer)
    ArcSal = FreeFile
    Open sLpt For Output As ArcSal
    Print #ArcSal, Chr$(27) & Chr$(64);            'Inicializa Impresora
    If pbCondensado Then
       Print #ArcSal, Chr$(27) & Chr$(108) & Chr$(0); 'Tipo letra : 0,1,2 - Roman,SansS,Courier
       Print #ArcSal, Chr$(27) & Chr$(77);            'Tamaño  : 80, 77, 103
       Print #ArcSal, Chr$(15);
    End If
    Print #ArcSal, Chr$(27) & Chr$(50);            'Espaciamiento lineas 1/6 pulg.1
    Print #ArcSal, Chr$(27) & Chr$(67) & Chr$(nLineas); '   Chr$(nLineas); 'Longitud de página a 66 líneas
    'Print #ArcSal, Chr$(27) & Chr$(87);            'Longitud de página a doble ancho
    If Not pbCondensado Then
        Print #ArcSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
        Print #ArcSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
        Print #ArcSal, Chr$(27) + Chr$(18); ' cancela condensada
        Print #ArcSal, Chr$(27) + Chr$(72); ' desactiva negrita
    End If
    Print #ArcSal, Chr$(27) & Chr$(120) & Chr$(0);  'Draf : 1 pasada
End Sub

'Procedimiento de impresión del recibo de comision por Remate
Public Sub ImprimirComision(pCodCta As String, pNomAdj As String, pMonCom As Currency, Optional ByVal PVENBRUTA As Currency, Optional ByVal PVENNETA As Currency, Optional ByVal PIGV As Currency)
    Dim vNombre As String * 27
    Dim vespacio As Integer
    vNombre = pNomAdj
    MousePointer = 11
    'vEspacio = 9
    vespacio = 34
    ImpreBegChe True, 22
        Print #ArcSal, "": Print #ArcSal, ""
        Print #ArcSal, Chr$(27) & Chr$(69);    'Activa Negrita
        Print #ArcSal, Tab(15); " CREDITO PIGNORATICIO" & Space(23 + vespacio) & "CREDITO PIGNORATICIO"
        Print #ArcSal, ""
        Print #ArcSal, Chr$(27) & Chr$(70);    'Desactiva Negrita
        Print #ArcSal, Tab(5); ImpreFormat(gsNomAge, 26, 0) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & Space(vespacio) & ImpreFormat(gsNomAge, 26, 0) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm")
        Print #ArcSal, Chr$(27) & Chr$(69);    'Activa Negrita
        Print #ArcSal, Tab(4); " CONTRATO     :" & Space(28 - Len(pCodCta)) & pCodCta & Space(vespacio) & "CONTRATO     :" & Space(28 - Len(pCodCta)) & pCodCta
        Print #ArcSal, Tab(4); "ADJUDICATARIO: " & vNombre & Space(vespacio) & "ADJUDICATARIO: " & vNombre
        Print #ArcSal, ""
        Print #ArcSal, Tab(4); "COMISION DE REMATE:" & ImpreFormat(pMonCom, 20, , True) & Space(vespacio) & "COMISION DE REMATE:" & ImpreFormat(pMonCom, 20, , True)
        Print #ArcSal, Tab(4); "SUBTOTAL          :" & ImpreFormat(PVENBRUTA, 20, , True) & Space(vespacio) & "SUBTOTAL          :" & ImpreFormat(PVENBRUTA, 20, , True)
        Print #ArcSal, Tab(4); "IGV               :" & ImpreFormat(PIGV, 20, , True) & Space(vespacio) & "IGV               :" & ImpreFormat(PIGV, 20, , True)
        Print #ArcSal, Tab(4); "VENTA NETA        :" & ImpreFormat(PVENNETA, 20, , True) & Space(vespacio) & "VENTA NETA        :" & ImpreFormat(PVENNETA, 20, , True)
        Print #ArcSal, ""
        Print #ArcSal, Tab(4); "Venta en Remate   " & Space(20) & Format(gsCodUser, "@@@@"); Space(vespacio) & "Venta en Remate   " & Space(20) & Format(gsCodUser, "@@@@")
        Print #ArcSal, Chr$(27) & Chr$(70);    'Desactiva Negrita
    ImpreEnd
    MousePointer = 0
End Sub

'Procedimiento de impresión del recibo de venta en subasta
Public Sub ImpRecVtaSub(pCodCta As String, pNomAdj As String, pMonCom As Currency)
    Dim vNombre As String * 28
    Dim vespacio As Integer
    vNombre = pNomAdj
    MousePointer = 11
    vespacio = 6
    ImpreBegChe False, 22
        Print #ArcSal, Chr$(27) & Chr$(77)
        Print #ArcSal, "": Print #ArcSal, ""
        Print #ArcSal, Chr$(27) & Chr$(69);    'Activa Negrita
        Print #ArcSal, Tab(15); " CREDITO PIGNORATICIO" & Space(23 + vespacio) & "CREDITO PIGNORATICIO"
        Print #ArcSal, ""
        Print #ArcSal, Chr$(27) & Chr$(70);    'Desactiva Negrita
        Print #ArcSal, ImpreFormat(gsNomAge, 26, 0) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & Space(vespacio) & ImpreFormat(gsNomAge, 26, 0) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm")
        Print #ArcSal, Chr$(27) & Chr$(69);    'Activa Negrita
        Print #ArcSal, "CONTRATO     :" & Space(28 - Len(pCodCta)) & pCodCta & Space(vespacio + 1) & "CONTRATO     :" & Space(28 - Len(pCodCta)) & pCodCta
        Print #ArcSal, "ADJUDICATARIO:" & vNombre & Space(vespacio) & "ADJUDICATARIO:" & vNombre
        Print #ArcSal, ""
        Print #ArcSal, "MONTO DE VENTA    :" & ImpreFormat(pMonCom, 20, , True) & Space(vespacio) & "MONTO DE VENTA    :" & ImpreFormat(pMonCom, 20, , True)
        Print #ArcSal, ""
        Print #ArcSal, "Venta en Subasta  " & Space(20) & Format(gsCodUser, "@@@@"); Space(vespacio) & "Venta en Subasta   " & Space(20) & Format(gsCodUser, "@@@@")
        Print #ArcSal, Chr$(27) & Chr$(70);    'Desactiva Negrita
    ImpreEnd
    MousePointer = 0
End Sub


'Procedimiento de impresión del recibo de venta en subasta
Public Sub ImpRecupSub(ByVal pCodCta As String, ByVal pNomAdj As String, ByVal pMonRecup As Currency, ByVal pnTotRecup As Currency, ByVal pnMonITF As Currency)
    Dim vNombre As String * 28
    Dim vespacio As Integer
    vNombre = pNomAdj
    MousePointer = 11
    vespacio = 9
    ImpreBegChe False, 22
        Print #ArcSal, "": Print #ArcSal, ""
        Print #ArcSal, Chr$(27) & Chr$(69);    'Activa Negrita
        Print #ArcSal, Tab(15); " CREDITO PIGNORATICIO" & Space(23 + vespacio) & "CREDITO PIGNORATICIO"
        Print #ArcSal, ""
        Print #ArcSal, Chr$(27) & Chr$(70);    'Desactiva Negrita
        Print #ArcSal, ImpreFormat(gsNomAge, 26, 0) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & Space(vespacio) & ImpreFormat(gsNomAge, 26, 0) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm")
        Print #ArcSal, Chr$(27) & Chr$(69);    'Activa Negrita
        Print #ArcSal, "CONTRATO     :" & Space(28 - Len(pCodCta)) & pCodCta & Space(vespacio) & "CONTRATO     :" & Space(28 - Len(pCodCta)) & pCodCta
        Print #ArcSal, "CLIENTE      :" & vNombre & Space(vespacio) & "CLIENTE      :" & vNombre
        'Print #ArcSal, ""
        Print #ArcSal, "MONTO DE RECUPERACION:" & ImpreFormat(pMonRecup, 15, , True) & Space(vespacio + 2) & "MONTO DE RECUPERACION:" & ImpreFormat(pMonRecup, 15, , True)
        Print #ArcSal, "ITF EFECTIVO         :" & ImpreFormat(pnMonITF, 15, , True) & Space(vespacio + 2) & "ITF EFECTIVO         :" & ImpreFormat(pnMonITF, 15, , True)
        Print #ArcSal, "TOTAL DE RECUPERACION:" & ImpreFormat(pnTotRecup, 15, , True) & Space(vespacio + 2) & "TOTAL DE RECUPERACION:" & ImpreFormat(pnTotRecup, 15, , True)
      '  Print #ArcSal, ""
        Print #ArcSal, "JOYA(S) ADJUDICADA RECUPERADA" & Space(8) & Format(gsCodUser, "@@@@"); Space(vespacio + 1) & "JOYA(S)ADJUDICADA RECUPERADA" & Space(8) & Format(gsCodUser, "@@@@")
        
        Print #ArcSal, Chr$(27) & Chr$(70);    'Desactiva Negrita
    ImpreEnd
    MousePointer = 0
End Sub

'Carga en un List las Agencias de la CMACT
Public Sub CargaAgencias(ByRef pLista As ListBox)
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    tmpSql = " SELECT cValor, cNomTab FROM " & gcCentralCom & "TablaCod WHERE substring(cValor,1,3) in ('112') AND " & _
        " substring(cCodTab,1,2) in ('47') "
    tmpReg.Open tmpSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    If (tmpReg.BOF Or tmpReg.EOF) Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        pLista.Clear
        With tmpReg
            Do While Not .EOF
                pLista.AddItem Mid(!cValor, 4, 2) & " " & Trim(!cNomTab)
                If !cValor = gsCodAge Then
                    pLista.Selected(pLista.ListCount - 1) = True
                End If
                .MoveNext
            Loop
        End With
    End If
    tmpReg.Close
    Set tmpReg = Nothing
End Sub

'Devuelve valor deseado.
'Se le envia una cadena SQL y que devuelva Campo
Public Function FuncGnral(pSql As String, Optional pLocal As Boolean = True) As Variant
Dim tmpReg As New ADODB.Recordset
Dim vConexion As ADODB.Connection
If pLocal Then
    Set vConexion = dbCmact
Else
    Set vConexion = dbCmactN
End If
tmpReg.Open pSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
If (tmpReg.BOF Or tmpReg.EOF) Then
    FuncGnral = ""
Else
    FuncGnral = Trim(tmpReg!Campo)
End If
tmpReg.Close
Set tmpReg = Nothing
End Function

Public Function IsCtaBlo(ByVal pCodCta As String, pConex As ADODB.Connection) As Boolean
Dim tmpReg As New ADODB.Recordset
Dim tmpSql As String

tmpSql = "SELECT * FROM HistBloqueo " & _
    " WHERE cCodCta = '" & pCodCta & "' AND nEstBlq = 1 AND dDesBlq IS NULL "
tmpReg.Open tmpSql, pConex, adOpenStatic, adLockOptimistic, adCmdText
If (tmpReg.BOF Or tmpReg.EOF) Then
    IsCtaBlo = False
Else
    IsCtaBlo = True
End If
tmpReg.Close
Set tmpReg = Nothing

End Function


Public Function InicialBoveda(pCodAge As String) As String
InicialBoveda = ""
Select Case pCodAge
    Case "11201"
        InicialBoveda = "Pi"
    Case "11202"
        InicialBoveda = "Zo"
    Case "11203"
        InicialBoveda = "Po"
    Case "11204"
        InicialBoveda = "SD"
    Case "11205"
        InicialBoveda = "Es"
    Case "11206"
        InicialBoveda = "Ch"
    Case "11207"
        InicialBoveda = "Se"
    Case "11208"
        InicialBoveda = "Hu"
    Case "11209"
        InicialBoveda = "He"
    Case "11210"
        InicialBoveda = "Vi"
End Select

End Function

'-- Carga las Agencias Seleccionadas
'--
Public Function fgCargaAgenciasSelec() As String
Dim lnCont As Integer
Dim lsAge As String
Dim lbFlag As Boolean
lsAge = "("
lbFlag = False
For lnCont = 1 To frmSelectAgencias.List1.ListCount
    If frmSelectAgencias.List1.Selected(lnCont - 1) = True Then
        lbFlag = True
        lsAge = lsAge & "'" & Mid(frmSelectAgencias.List1.List(lnCont - 1), 1, 2) & "',"
    End If
Next lnCont
If lbFlag = True Then
    lsAge = Mid(lsAge, 1, Len(lsAge) - 1) & ")"
Else
   'lsBov = "('" & gsCodAge & "')"
   lsBov = ""
End If
fgCargaAgenciasSelec = lsAge
End Function

'-- Carga las Bovedas Seleccionadas
'--
Public Function fgCargaBovedaSelec() As String
Dim lnCont As Integer
Dim lsBov As String
Dim lbFlag As Boolean
lsBov = "("
lbFlag = False
For lnCont = 1 To frmColPSelectBoveda.List1.ListCount
    If frmColPSelectBoveda.List1.Selected(lnCont - 1) = True Then
        lbFlag = True
        lsBov = lsBov & "'" & Mid(frmColPSelectBoveda.List1.List(lnCont - 1), 1, 2) & "',"
    End If
Next lnCont
If lbFlag = True Then
    lsBov = Mid(lsBov, 1, Len(lsBov) - 1) & ")"
Else
   'lsBov = "('" & gsCodAge & "')"
   lsBov = ""
End If
fgCargaBovedaSelec = lsBov
End Function

Public Function RenovacionEnDia(ByVal pContrato As String, pConexion As ADODB.Connection) As String
Dim lsSQL As String
Dim rr As New ADODB.Recordset

lsSQL = "SELECT dFecha,cCodAge FROM TransPrenda " & _
        "WHERE cCodCta = '" & pContrato & "' " & _
        "AND cCodTran in ('030500','031800','032000','032050','032100','032200') " & _
        "AND  datediff(day,dfecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 AND cFlag is null "
rr.Open lsSQL, pConexion, adOpenStatic, adLockReadOnly, adCmdText
If rr.BOF And rr.EOF Then
    RenovacionEnDia = ""
Else
    RenovacionEnDia = " " & Format(rr!dFecha, "dd/mm/yyyy hh:mm") & "  en Ag. " & IIf(IsNull(rr!cCodAge), "", rr!cCodAge)
End If
rr.Close
Set rr = Nothing
End Function

'*** FUNCION QUE VERIFICA QUE DOCUMENTOS DE IDENTIDAD ESTEN CORRECTOS ******
'****
Public Function fgVerIdentificacionPers(ByVal pcTipPers As String, ByVal pcTidoci As String, ByVal pcNudoci As String) As Boolean
Dim lbOk As Boolean
lbOk = True
If pcTipPers = "1" Then
   If (pcTidoci = "1" And pcNudoci = "null") _
      Or (pcTidoci = "1" And Trim(pcNudoci) = "00000000") _
      Or (pcTidoci = "1" And Len(Trim(pcNudoci)) <> 8) Then
     lbOk = False
   End If
   If pcTidoci = "9" Then
     lbOk = False
   End If
End If
fgVerIdentificacionPers = lbOk
End Function

'*** FUNCION QUE VERIFICA QUE DOCUMENTOS DE IDENTIDAD ESTEN CORRECTOS ******
'****
Public Function fgVerificaDocIdPers(ByVal pCodPers As String) As Boolean
Dim lr As New ADODB.Recordset
Dim lsSQL As String
Dim lbOk As Boolean
lbOk = True
lsSQL = "SELECT cTipPers,cTidoci,cNudoCi,cTidotr,cNudotr FROM " & gcCentralPers & "Persona  " & _
        "WHERE cCodPers ='& pCodPers & " ' "
lr.Open lsSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
If lr.BOF And lr.EOF Then
    lbOk = False
Else
    If lr!cTipPers = "1" Then
       If (lr!ctidoci = "1" And lr!cnudoci = "null") _
          Or (lr!ctidoci = "1" And Trim(lr!cnudoci) = "00000000") _
          Or (lr!ctidoci = "1" And Len(Trim(lr!cnudoci)) <> 8) Then
         lbOk = False
       End If
       If pcTidoci = "9" Then
         lbOk = False
       End If
       If lr!ctidoci = "6" Or lr!ctidoci = "7" Or lr!ctidoci = "8" Then
         lbOk = False
       End If
       
    End If
    If lr!cTipPers = "2" Then
    
    End If

End If

End Function


