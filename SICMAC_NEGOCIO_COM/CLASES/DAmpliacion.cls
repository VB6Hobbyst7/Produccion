VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAmpliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function ListaDatosAmpliacion(ByVal psctacod As String) As Recordset
    Dim sSql As String
    Dim oCn As DConecta
    
    sSql = " Select Pe.cPersNombre as Nombre,"
    sSql = sSql & " TipoProducto=Case Substring(P.cCtaCod,6,3)"
    sSql = sSql & " When '101' Then 'COMERCIAL EMPRESARIAL'"
    sSql = sSql & " When '102' Then 'COMERCIAL AGRICOLA'"
    sSql = sSql & " When '103' Then 'COMERCIAL PESQUERO'"
    sSql = sSql & " When '121' Then 'COMERCIAL CARTA FIANZA'"
    sSql = sSql & " When '201' Then 'MES EMPRESARIAL'"
    sSql = sSql & " When '202' Then 'MES AGRICOLA'"
    sSql = sSql & " When '301' Then 'CONSUMO DESCUENTO POR PLANILLA'"
    sSql = sSql & " When '302' Then 'CONSUMO PLAZO FIJO'"
    sSql = sSql & " When '303' Then 'CONSUMO CTS'"
    sSql = sSql & " When '304' Then 'CONSUMO PERSONAL DIRECTO'"
    sSql = sSql & " When '320' Then 'CONSUMO A DIRECTORES Y TRABAJADORES'"
    sSql = sSql & " When '401' Then ' HIPOTECARIO PARA VIVIENDA'"
    sSql = sSql & " When '403' Then 'HIPOTECARIO DIRECTORES Y TRABAJADORES'"
    sSql = sSql & " When '423' Then 'HIPOTECARIO MIVIVIENDA'"
    sSql = sSql & " end,"
    sSql = sSql & " Moneda=Case Substring(P.cCtaCod,9,1)"
    sSql = sSql & " When '1' Then 'Soles'"
    sSql = sSql & " When '2' Then 'Dolares'"
    sSql = sSql & " End"
    sSql = sSql & " From  Producto P"
    sSql = sSql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod"
    sSql = sSql & " Inner Join Persona Pe on Pe.cPersCod=PP.cPerscod and PP.nPrdPersRelac=20"
    sSql = sSql & " Where P.cCtaCod='" & psctacod & "'"
    
        
    Set oCn = New DConecta
    oCn.AbreConexion
    Set ListaDatosAmpliacion = oCn.CargaRecordSet(sSql)
    oCn.CierraConexion
    Set oConec = Nothing
End Function

Public Function VerificaAmpliado(ByVal psctacod As String) As Boolean
    Dim sSql As String
    Dim oCn As DConecta
    Dim rs As ADODB.Recordset
    Dim nCantidad As Integer
    
    sSql = "Select Count(*) as Cantidad From ColocacAmpliado Where cCtaCod='" & psctacod & "'"
    
    Set oCn = New DConecta
    oCn.AbreConexion
    
    Set rs = oCn.CargaRecordSet(sSql)
    oCn.CierraConexion
    Set cn = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        nCantidad = rs!cantidad
    End If
    
    If nCantidad = 1 Then
        VerificaAmpliado = False
    Else
        VerificaAmpliado = True
    End If
End Function

Public Function ValidaMontoAmpliado(ByVal nMontoAmpliado As Double, ByVal nMonedaAmpliado As Integer, _
                                    ByVal nMonto As Double, ByVal nMoneda As Integer, ByVal dFecha As Date) As Boolean
     
     'Definiciones
     'nMontoAmpliado Cuenta antigua
     'nMonto Cuenta Nueva
     
     Dim oGeneral As DGeneral
     Dim nTipoCambioFijo As Double
     Dim bValida As Boolean
     Dim lnMonto As Double
     
     bValida = True
    
     Set oGeneral = New DGeneral
     nTipoCambioFijo = oGeneral.EmiteTipoCambio(dFecha, TCFijoMes)
     nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))

     If nMonedaAmpliado = 1 Then ' Moneda Nacional
        If nMoneda = 1 Then
            If nMontoAmpliado > nMonto Then
                bValida = False
            End If
        Else
            If (nMonto * nTipoCambioFijo) < nMontoAmpliado Then
                bValida = False
            End If
        End If
        
    Else
        If nMoneda = 1 Then
            If (nMonto / nTipoCambioFijo) < nMontoAmpliado Then
                 bValida = False
            End If
       Else
             If nMontoAmpliado > nMonto Then
                bValida = False
             End If
       End If
    End If
        
     
   ValidaMontoAmpliado = bValida

End Function
Public Function ValidaMontoAmpliadoLista(ByVal rsAmpliado As ADODB.Recordset, ByVal nMontoSol As Double, _
                                         ByVal nMonedaSol As Integer) As Boolean
    'llevando todo a dolares
     Dim oGeneral As DGeneral
     Dim nTipoCambioFijo As Double
     Dim bValida As Boolean
     Dim lnMonto As Double
     Dim sMonto As Double
     
     bValida = True
'     CDbl(rsAmpliado(3)), CInt(Mid(rsAmpliado(0), 9, 1)), Val(Me.txtMontoSol), Mid(psNuevaCta, 9, 1)
     
     Set oGeneral = New DGeneral
     nTipoCambioFijo = oGeneral.EmiteTipoCambio(dFecha, TCFijoMes)
     nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
     
     If nMonedaSol = 2 Then
        lnMonto = CDbl(nMontoSol / nTipoCambio)
     Else
        lnMonto = nMontoSol
     End If
     
     rsAmpliado.MoveFirst
     sMonto = 0
     Do Until rsAmpliado.EOF
        
        If Mid(rsAmpliado(0), 9, 1) = "2" Then
            sMonto = sMonto + (rsAmpliado(3) / nTipoCambio)
        Else
            sMonto = sMonto + CDbl(rsAmpliado(3))
        End If
        rsAmpliado.MoveNext
     Loop
     Set rsAmpliado = Nothing
     
     If lnMonto >= sMonto Then
        ValidaMontoAmpliadoLista = True
     Else
        ValidaMontoAmpliadoLista = False
     End If
End Function

Public Function ValidacionCredito(ByVal psctacod As String) As Boolean
    Dim sSql As String
    Dim oCn  As DConecta
    Dim rs As ADODB.Recordset
    Dim nCantidad As Integer
    
    
    sSql = sSql & " Select Count(*) as Cantidad"
    sSql = sSql & " From Producto"
    sSql = sSql & " Where cCtaCod='" & psctacod & "' and"
    sSql = sSql & " nPrdEstado in (2020)"

    Set oCn = New DConecta
    oCn.AbreConexion
    Set rs = oCn.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        nCantidad = rs!cantidad
    End If
    Set rs = Nothing
    
    If nCantidad = 1 Then
        ValidacionCredito = True
    Else
        ValidacionCredito = False
    End If
    
End Function


Public Function ListaCreditoPorAmpliar(ByVal cPersCod As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select distinct P.cCtaCod,Pers.cPersNombre"
    sSql = sSql & " From Persona Pers"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cPersCod=Pers.cPersCod"
    sSql = sSql & " Inner Join Producto P on PP.cCtaCod=P.cCtaCod"
    sSql = sSql & " Where P.nPrdEstado in (2020) and PP.nPrdPersRelac=20"
    sSql = sSql & " and PP.cPersCod='" & cPersCod & "'"
        
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCreditoPorAmpliar = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

' Function que emula la liberacion de la garantia para un credito ampliado

Public Function EmularGarantia(ByVal psctacod As String, ByVal nMoneda As Integer) As Double
    Dim oConec As DConecta
    Dim sSql As String
    Dim nMontoCol As Double
    Dim nMontoK As Double
    Dim rs As ADODB.Recordset
    Dim oCred As DCredito
    Dim nSumaTotalGarant As Double
    Dim nPago As Double
    Dim nMontoLiberar As Double
    Dim rsAmpliado As ADODB.Recordset
    Dim nTotalGarant As Double
    ' se obtiene el monto colocado
    
    sSql = "Select nMontoCol From Colocaciones Where cCtaCod='" & psctacod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        nMontoCol = rs!nMontoCol
    End If
    
    
    ' Se obtiene la sumatoria capitales pendientes
   
    sSql = "Select Sum(CD.nMonto-CD.nMontoPagado) as nMontoCapital"
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join ColocCalendario CC on P.cCtaCod=CC.cCtaCod"
    sSql = sSql & " Inner Join ColocCalendDet CD on CD.cCtaCod=CC.cCtaCod and CC.nNroCalen=CD.nNroCalen  and"
    sSql = sSql & " CC.nColocCalendApl = CD.nColocCalendApl And CC.nCuota = CD.nCuota"
    sSql = sSql & " Where P.cCtaCod='" & psctacod & "' and CD.nPrdConceptoCod=1000 and CC.nColocCalendEstado=0 and"
    sSql = sSql & " CD.nColocCalendApl=1 and"
    sSql = sSql & " CC.nNroCalen=(Select Max(nNroCalen) From ColocCalendario Where cCtaCod='" & psctacod & "')"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.BOF And Not rs.EOF Then
        nMontoK = rs!nMontoCapital
    End If
    Set rs = Nothing
    
    'Obtiene el monto Total de la garantia
    Set oCred = New DCredito
    Set rs = oCred.RecuperaColocGarantia(psctacod)
    Do While Not rs.EOF
        If rs!nMoneda = nMoneda Then
            nSumaTotalGarant = nSumaTotalGarant + rs!nGravado
        End If
        rs.MoveNext
    Loop
    nSumaTotalGarant = Format(nSumaTotalGarant, "#0.00")
    
    ' se obtiene el porcentaje de pago de la cancelacion de la deuda
    nPago = Format(nMontoK * nSumaTotalGarant / nMontoCol, "#0.00")
    Set rs = Nothing
    
    Set rs = oCred.RecuperaColocGarantia(psctacod)
    Set oCred = Nothing
    
    ' se suma los disponible con la emulacion de la liberacion de la garantia
    Do Until rs.EOF
        If rs!nMoneda = nMoneda Then
            nMontoLiberar = nPago * (rs!nGravado / nSumaTotalGarant)
            sSql = "Select (G.nPorGravar-G.nGravament) as nDisponible"
            sSql = sSql & " From ColocGarantia CG"
            sSql = sSql & " Inner Join Garantias G on CG.cNumGarant=G.cNumGarant"
            sSql = sSql & " Inner Join Producto P on CG.cCtaCod=P.cCtaCod"
            sSql = sSql & " Where P.cCtaCod='" & psctacod & "'" & " and  G.cNumGarant='" & rs!cNumGarant & "'"
            
            Set rsAmpliado = New ADODB.Recordset
            
            Set oConec = New DConecta
            oConec.AbreConexion
            Set rsAmpliado = oConec.CargaRecordSet(sSql)
            oConec.CierraConexion
            Set oConec = Nothing
            
            If Not rsAmpliado.EOF And Not rsAmpliado.BOF Then
                nTotalGarant = nTotalGarant + (rsAmpliado!nDisponible + nMontoLiberar)
            End If
            Set rsAmpliado = Nothing
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
    
'    If nTotalGarant >= nMonto Then
'        EmularGarantia = True
'    Else
'        EmularGarantia = False
'    End If
    EmularGarantia = nTotalGarant
End Function


Public Function ValidaCreditoaAmpliar(ByVal psctacod As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim oConec As DConecta
    Dim sSql As String
    Dim nCantidad As Integer
    
    sSql = "Select Count(*) as nCantidad From ColocacAmpliado Where cCtaCod='" & psctacod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        nCantidad = rs!nCantidad
    End If
    
    Set rs = Nothing
    
    If nCantidad > 0 Then
        ValidaCreditoaAmpliar = True
    Else
        ValidaCreditoaAmpliar = False
    End If
End Function


Public Function GetcCtaCodPorAmpliar(ByVal psCtaCodNew As String) As String
    Dim rs As ADODB.Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select cCtaCodAmp From ColocacAmpliado Where cCtaCod='" & psCtaCodNew & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    If Not rs.EOF And Not rs.BOF Then
        GetcCtaCodPorAmpliar = IIf(IsNull(rs!cCtaCodAmp), "", rs!cCtaCodAmp)
    End If
    Set rs = Nothing
End Function

Public Function AsientoGarantiaAmpliado(ByVal psctacod As String) As Boolean
    Dim oConec As DConecta
    Dim oGeneral As DGeneral
    
    Dim sSql As String
    Dim sCtaCodNew As String
    Dim sNumGarant As String
    
    Dim nMonto As Double
    Dim nTipoCambioFijo As Double
    Dim nMontoDispGarant As Double
    Dim nMontoAsig As Double
    
    Dim rs As ADODB.Recordset
    Dim rsGarant As ADODB.Recordset
    
    Dim bEstado As Boolean
    
    Dim nMonedaGarant As Integer
    
    
    'Se obtiene el Codigo del Cuenta anterior que se va ampliar
    
    sSql = "Select cCtaCodAmp From ColocacAmpliado Where cCtaCod='" & psctacod & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not rs.BOF And Not rs.EOF Then
        sCtaCodNew = rs!cCtaCodAmp
    End If
    
    Set rs = Nothing
    
    'obtener el monto de la aprobacion para comenzar a coberturar la garantia pe!!!!
    sSql = "Select nMonto From ColocacEstado Where cCtaCod='" & psctacod & "'"
    
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        nMonto = rs!nMonto
    End If
    
    ' obtener el tipo de cambio
    Set oGeneral = New DGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
    nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
    Set oGeneral = Nothing
    
    'obtener las garantias relacionadas con el nuevo credito
    
    sSql = "Select cNumGarant,nGravado From ColocGarantia Where cCtacod='" & psctacod & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Do Until rs.EOF
        If Mid(psctacod, 9, 1) = "1" Then 'Soles
            If rs!nMoneda = 1 Then
                nMonto = nMonto - rs!nGravado
            Else
                'cambiado al tipo de moneda del credito
                'convirtiendo a dolares
                nMonto = nMonto - Val(rs!nGravado) * nTipoCambioFijo
            End If
        Else
            'dolares
            If rs!nMoneda = 2 Then
                nMonto = nMonto - rs!nGravado
            Else
                nMonto = nMonto - (Val(rs!nGravado) / nTipoCambioFijo)
            End If
            
        End If
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    
    'Obteniendo las garantias del credito antiguo relacionado
    sSql = "Select * From ColocGarantia Where cCtaCod='" & psCtaCodNew & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Do Until rs.EOF
        'calculado el monto disponible de la garantia
        sSql = "Select nPorGravar-nGravament as nDisponible,nMoneda From Garantias Where cNumGarant='" & rs!cNumGarant & "'"
        oConec.AbreConexion
        Set rsGarant = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        
        If Not rsGarant.EOF And Not rsGarant.BOF Then
            nMontoDispGarant = rsGarant!nDisponible
        End If
        
        'insertando el colocgarantia y actualizando la garantia
        If nMonto > 0 Then
            'calculando el monto de la garantia que se va asignar
            If nMonto >= ObtieneMonto(Mid(psctacod, 9, 1), rs!nMoneda, nMontoDispGarant) Then
               nMontoAsig = nMontoDispGarant
               nMonto = nMonto - ObtieneMonto(Mid(psctacod, 9, 1), rs!nMoneda, nMontoAsig)
            Else
               nMontoAsig = nMonto
               nMonto = nMonto - ObtieneMonto(Mid(psctacod, 9, 1), rs!nMoneda, nMontoAsig)
            End If
            oConec.AbreConexion
            oConec.ConexionActiva.BeginTrans
            sSql = "Insert Into ColocGarantia Values('" & rs!cNumGarant & "','" & psctacod & "'," & rs!nMoneda & "," & nMontoAsig & ",1)"
            oConec.ConexionActiva.Execute sSql
            'actualizando la garantia
            sSql = "Update ColocGarantia"
            sSql = sSql & " set nGravament=nGravament+" & nMontoAsig & ",nPorGravar=nPorGravar-" & nMontoAsig
            sSql = sSql & " Where cNumGarant='" & rs!cNumGarant
            oConec.ConexionActiva.Execute sSql
            oConec.CommitTrans
        End If
        rs.MoveNext
    Loop
End Function

Function ObtieneMonto(ByVal pnMonedaCuentas As Integer, ByVal pnMonedaGarantia, pnMontoGarantia As Double) As Double
    Dim oGeneral As DGeneral
    Dim nTipoCambioFijo As Double
    
    Set oGeneral = New DGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
    nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
    Set oGeneral = Nothing
    
    If pnMonedaCuentas = 1 Then
        If pnMonedaGarantia = 1 Then
            ObtieneMonto = pnMontoGarantia
        Else
            ObtieneMonto = pnMontoGarantia * nTipoCambio
        End If
    Else
        If pnMonedaGarantia = 1 Then
            ObtieneMonto = (pnMontoGarantia / nTipoCambio)
        Else
            ObtieneMonto = pnMontoGarantia
        End If
    End If
    
End Function

Public Function ListaCreditosAmpliadosByNombre(ByVal psNombre As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    Set oConec = New DConecta
    sSql = "Select P.cPersNombre,CA.cCtaCod"
    sSql = sSql & " From Persona P"
    sSql = sSql & " Inner Join ProductoPersona PP on P.cPersCod=PP.cPersCod and PP.nPrdPersRelac=20"
    sSql = sSql & " Inner Join ColocacAmpliado CA on CA.cCtaCodAmp=PP.cCtaCod"
    sSql = sSql & " Inner Join Producto PRO on PRO.cCtaCod=CA.cCtaCodAmp"
    sSql = sSql & " Where P.cPersNombre like '%" & psNombre & "%' and PRO.nPrdEstado=2020"
    sSql = sSql & "Group By P.cPersNombre,CA.cCtaCod"
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCreditosAmpliadosByNombre = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaCreditosBycCtaCodNew(ByVal psctacod As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select cCtaCodAmp From ColocacAmpliado Where cCtaCod='" & psctacod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCreditosBycCtaCodNew = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ValidarCreditoGarantizar(ByVal psctacod As String) As Boolean
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim bValor As Boolean
    
    bValor = False
    
    sSql = "Select Count(*) as nCantidad"
    sSql = sSql & " From ColocacAmpliado CA"
    sSql = sSql & " Inner Join Producto P on P.cCtaCod=CA.cCtaCod"
    sSql = sSql & " Where P.nPrdEstado in (2001,2000) and cCtaCod='" & psctacod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        If rs!nCantidad > 0 Then
            bValor = True
        Else
            bValor = False
        End If
    End If
    Set rs = Nothing
    ValidarCreditoGarantizar = bValor
End Function

Public Function ListaGarantiaCtaAmpliado(ByVal psctacod As String, ByVal psCtaCodAmp As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select G.cNumGarant,cDescripcion,CN.cConsDescripcion,CG.nGravado"
    sSql = sSql & " From ColocacAmpliado CA"
    sSql = sSql & " Inner Join ColocGarantia CG on CG.cCtaCod=CA.cCtaCodAmp"
    sSql = sSql & " Inner Join Garantias G on G.cNumGarant=CG.cNumGarant"
    sSql = sSql & " Inner Join Constante CN on CN.nConsCod=1027 and CN.nConsValor<>1027 and CN.nConsValor=G.nTpoGarantia"
    sSql = sSql & " Where CA.cCtaCod='" & psctacod & "' and CA.cCtaCodAmp='" & psCtaCodAmp & "'"

    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaGarantiaCtaAmpliado = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
End Function

Public Function InsertarGarantiaAmpliado(ByVal pRs As ADODB.Recordset) As Boolean
    Dim oConec As DConecta
    Dim sSql As String
    Dim sCtaCod As String
    Dim sNumGarant As String
    Dim nGravado As Double
    
    On Error GoTo ErrHandler
    
    Set oConec = New DConecta
    oConec.AbreConexion
    
    If Not pRs.EOF And Not pRs.BOF Then
        pRs.MoveFirst
        sNumGarant = pRs(0)
        sCtaCod = pRs(1)
        oConec.ConexionActiva.BeginTrans
         Do Until pRs.EOF
             If sNumGarant = pRs(0) And sCtaCod = pRs(1) Then
                nGravado = nGravado + Val(pRs(3))
             Else
                sSql = "Insert into GarantAmpliado Values('" & sNumGarant & "','" & sCtaCod & "'," & nGravado & ")"
                oConec.ConexionActiva.Execute sSql
                nGravado = 0
                sNumGarant = pRs(0)
                sCtaCod = pRs(1)
             End If
            pRs.MoveNext
         Loop
         sSql = "Insert into GarantAmpliado Values('" & sNumGarant & "','" & sCtaCod & "'," & nGravado & ")"
         oConec.ConexionActiva.Execute sSql
         oConec.ConexionActiva.CommitTrans
    End If
    oConec.CierraConexion
    Set oConec = Nothing
    InsertarGarantiaAmpliado = True
    Exit Function
ErrHandler:
    oConec.ConexionActiva.RollbackTrans
    InsertarGarantiaAmpliado = False
End Function


Public Function ObtenerMontoColocar(ByVal psctacod As String) As Double
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    
    sSql = "Select Top 1 C.nMontoCol"
    sSql = sSql & " From Colocaciones C"
    sSql = sSql & " Inner Join Colocacampliado CA on CA.cCtaCod=C.cCtaCod"
    sSql = sSql & " Where C.cCtaCod='" & psctacod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    If Not rs.EOF And Not rs.BOF Then
        ObtenerMontoColocar = IIf(IsNull(rs!nMontoCol), 0, rs!nMontoCol)
    Else
        ObtenerMontoColocar = 0
    End If
    Set rs = Nothing
End Function

Public Function ObtenerMonedaGarantia(ByVal pscNumGarant As String) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    
    sSql = "Select nMoneda From Garantias Where cNumGarant='" & pscNumGarant & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenerMonedaGarantia = IIf(IsNull(rs!nMoneda), 0, rs!nMoneda)
    End If
    Set rs = Nothing
End Function

Public Function ObteneMontoGarant(ByVal psNumGarant As String, ByVal pcCtaCod As String) As Double
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    
    sSql = "Select nGravado  From ColocGarantia where cCtaCod='" & pcCtaCod & "' And cNumGarant='" & psNumGarant & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    If Not rs.EOF And Not rs.BOF Then
        ObteneMontoGarant = IIf(IsNull(rs!nGravado), 0, rs!nGravado)
    End If
    Set rs = Nothing
End Function

Public Function ObtenerListaGarantiasAmpliadas(ByVal psctacod As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select * From GarantAmpliado Where cCtaCod='" & psctacod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ObtenerListaGarantiasAmpliadas = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function EliminarGarantiaAmpliada(ByVal psctacod As String) As Boolean
    Dim oConec As DConecta
    Dim sSql As String
    
    On Error GoTo ErrHandler
    
    sSql = "Delete From GarantAmpliado Where cCtaCod='" & psctacod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    oConec.ConexionActiva.Execute sSql
    oConec.CierraConexion
    Set oConec = Nothing
    EliminarGarantiaAmpliada = True
    Exit Function
ErrHandler:
    EliminarGarantiaAmpliada = False
End Function


Public Function ObtenerMontoGarantAmpliado(ByVal psctacod As String, ByVal pnMoneda As Integer) As Double
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select Sum(GA.nMontoGravado) as nGravado"
    sSql = sSql & " From GarantAmpliado GA"
    sSql = sSql & " Inner Join Garantias G on G.cNumGarant=GA.cNumGarant"
    sSql = sSql & " Where cCtaCod='" & psctacod & "' and G.nMoneda=" & pnMoneda
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenerMontoGarantAmpliado = IIf(IsNull(rs!nGravado), 0, rs!nGravado)
    Else
        ObtenerMontoGarantAmpliado = 0
    End If
    Set rs = Nothing
    
End Function

Public Sub ActualizandoGarantiasAmpliadas(ByVal psctacod As String)
   Dim oConec As DConecta
   Dim sSql As String
   Dim rs As ADODB.Recordset
        
        
   sSql = "Select GA.*,G.nMoneda"
   sSql = sSql & " From GarantAmpliado GA"
   sSql = sSql & " Inner Join Garantias G on G.cNumGarant=GA.cNumGarant"
   sSql = sSql & " Where cCtaCod='" & psctacod & "'"
   
   Set oConec = New DConecta
   oConec.AbreConexion
   Set rs = oConec.CargaRecordSet(sSql)
   
   If Not rs.EOF And Not rs.BOF Then
    oConec.ConexionActiva.BeginTrans
   Else
      Exit Sub
   End If
    
    Do Until rs.EOF
        sSql = "Insert Into ColocGarantia Values('" & rs!cNumGarant & "','" & psctacod & "'," & rs!nMoneda & "," & rs!nMontoGravado & ",1)"
        oConec.ConexionActiva.Execute sSql
        sSql = "Update Garantias"
        sSql = sSql & " Set nGravament=nGravament + " & rs!nMontoGravado
        sSql = sSql & " Where cNumGarant='" & rs!cNumGarant & "'"
        oConec.ConexionActiva.Execute sSql
     rs.MoveNext
    Loop
    Set rs = Nothing
    oConec.CommitTrans
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
