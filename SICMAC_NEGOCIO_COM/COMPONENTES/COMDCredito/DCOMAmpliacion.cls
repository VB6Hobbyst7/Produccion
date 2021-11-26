VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMAmpliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ListaDatosAmpliacion(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCn As COMConecta.DCOMConecta
    
    sSql = " Select Pe.cPersNombre as Nombre,"
'    sSql = sSql & " TipoProducto=Case Substring(P.cCtaCod,6,3)"
'    sSql = sSql & " When '101' Then 'COMERCIAL EMPRESARIAL'"
'    sSql = sSql & " When '102' Then 'COMERCIAL AGRICOLA'"
'    sSql = sSql & " When '103' Then 'COMERCIAL PESQUERO'"
'    sSql = sSql & " When '121' Then 'COMERCIAL CARTA FIANZA'"
'    sSql = sSql & " When '201' Then 'MES EMPRESARIAL'"
'    sSql = sSql & " When '202' Then 'MES AGRICOLA'"
'    sSql = sSql & " When '301' Then 'CONSUMO DESCUENTO POR PLANILLA'"
'    sSql = sSql & " When '302' Then 'CONSUMO PLAZO FIJO'"
'    sSql = sSql & " When '303' Then 'CONSUMO CTS'"
'    sSql = sSql & " When '304' Then 'CONSUMO PERSONAL DIRECTO'"
'    sSql = sSql & " When '320' Then 'CONSUMO A DIRECT. Y TRABAJAD.'"
'    sSql = sSql & " When '401' Then ' HIPOTECARIO PARA VIVIENDA'"
'    sSql = sSql & " When '403' Then 'HIPOTECARIO DIRECTORES Y TRABAJADORES'"
'    sSql = sSql & " When '423' Then 'HIPOTECARIO MIVIVIENDA'"
'    sSql = sSql & " end,"
    sSql = sSql & " TipoProducto=(Select cConsDescripcion from constante where nConsCod=3033 and nConsValor=C.cTpoProdCod),"
    sSql = sSql & " Moneda=Case Substring(P.cCtaCod,9,1)"
    sSql = sSql & " When '1' Then 'Soles'"
    sSql = sSql & " When '2' Then 'Dolares'"
    sSql = sSql & " End"
    sSql = sSql & " ,IdCampana=ISNULL(CC.IdCampana,0) " 'ARCV 10-03-2007
    sSql = sSql & " From  Producto P"
    sSql = sSql & " INNER JOIN ColocacCred CC ON P.cCtaCod=CC.cCtaCod"   'ARCV 10-03-2007
    sSql = sSql & " INNER JOIN Colocaciones C ON CC.cCtaCod=C.cCtaCod"   'ARCV 10-03-2007
    sSql = sSql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod"
    sSql = sSql & " Inner Join Persona Pe on Pe.cPersCod=PP.cPerscod and PP.nPrdPersRelac=20"
    sSql = sSql & " Where P.cCtaCod='" & psCtaCod & "'"
    
        
    Set oCn = New COMConecta.DCOMConecta
    oCn.AbreConexion
    Set ListaDatosAmpliacion = oCn.CargaRecordSet(sSql)
    oCn.CierraConexion
    Set oCn = Nothing
End Function

Public Function VerificaAmpliado(ByVal psCtaCod As String) As Boolean
    Dim sSql As String
    Dim oCn As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim nCantidad As Integer
    
    sSql = "Select Count(*) as Cantidad From ColocacAmpliado Where cCtaCod='" & psCtaCod & "'"
    
    Set oCn = New COMConecta.DCOMConecta
    oCn.AbreConexion
    
    Set rs = oCn.CargaRecordSet(sSql)
    oCn.CierraConexion
    Set oCn = Nothing
    
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
     
     Dim oGeneral As COMDConstSistema.DCOMGeneral
     Dim nTipoCambioFijo As Double
     Dim bValida As Boolean
     Dim lnMonto As Double
     
     bValida = True
    
     Set oGeneral = New COMDConstSistema.DCOMGeneral
     nTipoCambioFijo = oGeneral.EmiteTipoCambio(dFecha, COMDConstantes.TCFijoMes)
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
                                         ByVal nMonedaSol As Integer, ByVal pdFecha As Date) As Boolean
    'llevando todo a dolares
     Dim oGeneral As COMDConstSistema.DCOMGeneral
     Dim nTipoCambioFijo As Double
     Dim bValida As Boolean
     Dim lnMonto As Double
     Dim sMonto As Double
     
     bValida = True
'     CDbl(rsAmpliado(3)), CInt(Mid(rsAmpliado(0), 9, 1)), Val(Me.txtMontoSol), Mid(psNuevaCta, 9, 1)
     
     Set oGeneral = New COMDConstSistema.DCOMGeneral
     nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdFecha, COMDConstantes.TCFijoMes)
     nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
     
     If nMonedaSol = 2 Then
        lnMonto = CDbl(nMontoSol / nTipoCambioFijo)
     Else
        lnMonto = nMontoSol
     End If
     
     rsAmpliado.MoveFirst
     sMonto = 0
     Do Until rsAmpliado.EOF
        
        If Mid(rsAmpliado(0), 9, 1) = "2" Then
            sMonto = sMonto + (rsAmpliado(3) / nTipoCambioFijo)
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

Public Function ValidacionCredito(ByVal psCtaCod As String) As Boolean
    Dim sSql As String
    Dim oCn  As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim nCantidad As Integer
    
    
    sSql = sSql & " Select Count(*) as Cantidad"
    sSql = sSql & " From Producto"
    sSql = sSql & " Where cCtaCod='" & psCtaCod & "' and"
    sSql = sSql & " nPrdEstado in (2020)"

    Set oCn = New COMConecta.DCOMConecta
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

'************************ RECO 20130913 INC: INC1309120009**********************
Public Function ObtenerEstadoProducto(ByVal psCtaCod As String) As String
    Dim sSql As String
    Dim oCn  As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    sSql = "exec stp_sel_ObtenerEstadoProducto '" & psCtaCod & "'"
    
    Set oCn = New COMConecta.DCOMConecta
    oCn.AbreConexion
    Set rs = oCn.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        ObtenerEstadoProducto = rs!cConsDescripcion
    End If
    Set rs = Nothing
End Function
'*********************END RECO*************************************************


Public Function ListaCreditoPorAmpliar(ByVal cPersCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select distinct P.cCtaCod,Pers.cPersNombre"
    sSql = sSql & " From Persona Pers"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cPersCod=Pers.cPersCod"
    sSql = sSql & " Inner Join Producto P on PP.cCtaCod=P.cCtaCod"
    sSql = sSql & " Where P.nPrdEstado in (2020) and PP.nPrdPersRelac=20"
    sSql = sSql & " and PP.cPersCod='" & cPersCod & "'"
        
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCreditoPorAmpliar = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

' Function que emula la liberacion de la garantia para un credito ampliado

Public Function EmularGarantia(ByVal psCtaCod As String, ByVal nMoneda As Integer) As Double
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim nMontoCol As Double
    Dim nMontoK As Double
    Dim rs As ADODB.Recordset
    Dim oCred As DCOMCredito
    Dim nSumaTotalGarant As Double
    Dim nPago As Double
    Dim nMontoLiberar As Double
    Dim rsAmpliado As ADODB.Recordset
    Dim nTotalGarant As Double
    ' se obtiene el monto colocado
    
    sSql = "Select nMontoCol From Colocaciones Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
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
    sSql = sSql & " Where P.cCtaCod='" & psCtaCod & "' and CD.nPrdConceptoCod=1000 and CC.nColocCalendEstado=0 and"
    sSql = sSql & " CD.nColocCalendApl=1 and"
    sSql = sSql & " CC.nNroCalen=(Select Max(nNroCalen) From ColocCalendario Where cCtaCod='" & psCtaCod & "')"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.BOF And Not rs.EOF Then
        nMontoK = rs!nMontoCapital
    End If
    Set rs = Nothing
    
    'Obtiene el monto Total de la garantia
    Set oCred = New DCOMCredito
    Set rs = oCred.RecuperaColocGarantia(psCtaCod)
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
    
    Set rs = oCred.RecuperaColocGarantia(psCtaCod)
    Set oCred = Nothing
    
    ' se suma los disponible con la emulacion de la liberacion de la garantia
    Do Until rs.EOF
        If rs!nMoneda = nMoneda Then
            nMontoLiberar = nPago * (rs!nGravado / nSumaTotalGarant)
            sSql = "Select (G.nPorGravar-G.nGravament) as nDisponible"
            sSql = sSql & " From ColocGarantia CG"
            sSql = sSql & " Inner Join Garantias G on CG.cNumGarant=G.cNumGarant"
            sSql = sSql & " Inner Join Producto P on CG.cCtaCod=P.cCtaCod"
            sSql = sSql & " Where P.cCtaCod='" & psCtaCod & "'" & " and  G.cNumGarant='" & rs!cNumGarant & "'"
            
            Set rsAmpliado = New ADODB.Recordset
            
            Set oConec = New COMConecta.DCOMConecta
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


Public Function ValidaCreditoaAmpliar(ByVal psCtaCod As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim nCantidad As Integer
    
    sSql = "Select Count(*) as nCantidad From ColocacAmpliado Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
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
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select cCtaCodAmp From ColocacAmpliado Where cCtaCod='" & psCtaCodNew & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    If Not rs.EOF And Not rs.BOF Then
        GetcCtaCodPorAmpliar = IIf(IsNull(rs!cCtaCodAmp), "", rs!cCtaCodAmp)
    End If
    Set rs = Nothing
End Function

Public Function AsientoGarantiaAmpliado(ByVal psCtaCod As String, ByVal pdFecha As Date, ByVal psCtaCodNew As String) As Boolean

    Dim oConec As COMConecta.DCOMConecta
    Dim oGeneral As COMDConstSistema.DCOMGeneral
    
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
    
    sSql = "Select cCtaCodAmp From ColocacAmpliado Where cCtaCod='" & psCtaCod & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not rs.BOF And Not rs.EOF Then
        sCtaCodNew = rs!cCtaCodAmp
    End If
    
    Set rs = Nothing
    
    'obtener el monto de la aprobacion para comenzar a coberturar la garantia pe!!!!
    sSql = "Select nMonto From ColocacEstado Where cCtaCod='" & psCtaCod & "'"
    
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        nMonto = rs!nMonto
    End If
    
    ' obtener el tipo de cambio
    Set oGeneral = New COMDConstSistema.DCOMGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdFecha, COMDConstantes.TCFijoMes)
    nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
    Set oGeneral = Nothing
    
    'obtener las garantias relacionadas con el nuevo credito
    
    sSql = "Select cNumGarant,nGravado From ColocGarantia Where cCtacod='" & psCtaCod & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Do Until rs.EOF
        If Mid(psCtaCod, 9, 1) = "1" Then 'Soles
            If rs!nMoneda = 1 Then
                nMonto = nMonto - rs!nGravado
            Else
                'cambiado al tipo de moneda del credito
                'convirtiendo a dolares
                nMonto = nMonto - val(rs!nGravado) * nTipoCambioFijo
            End If
        Else
            'dolares
            If rs!nMoneda = 2 Then
                nMonto = nMonto - rs!nGravado
            Else
                nMonto = nMonto - (val(rs!nGravado) / nTipoCambioFijo)
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
            If nMonto >= ObtieneMonto(Mid(psCtaCod, 9, 1), rs!nMoneda, nMontoDispGarant, pdFecha) Then
               nMontoAsig = nMontoDispGarant
               nMonto = nMonto - ObtieneMonto(Mid(psCtaCod, 9, 1), rs!nMoneda, nMontoAsig, pdFecha)
            Else
               nMontoAsig = nMonto
               nMonto = nMonto - ObtieneMonto(Mid(psCtaCod, 9, 1), rs!nMoneda, nMontoAsig, pdFecha)
            End If
            oConec.AbreConexion
            oConec.ConexionActiva.BeginTrans
            sSql = "Insert Into ColocGarantia Values('" & rs!cNumGarant & "','" & psCtaCod & "'," & rs!nMoneda & "," & nMontoAsig & ",1)"
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

Function ObtieneMonto(ByVal pnMonedaCuentas As Integer, ByVal pnMonedaGarantia, pnMontoGarantia As Double, pdFecha As Date) As Double
    Dim oGeneral As COMDConstSistema.DCOMGeneral
    Dim nTipoCambioFijo As Double
    
    Set oGeneral = New COMDConstSistema.DCOMGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdFecha, COMDConstantes.TCFijoMes)
    nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
    Set oGeneral = Nothing
    
    If pnMonedaCuentas = 1 Then
        If pnMonedaGarantia = 1 Then
            ObtieneMonto = pnMontoGarantia
        Else
            ObtieneMonto = pnMontoGarantia * nTipoCambioFijo
        End If
    Else
        If pnMonedaGarantia = 1 Then
            ObtieneMonto = (pnMontoGarantia / nTipoCambioFijo)
        Else
            ObtieneMonto = pnMontoGarantia
        End If
    End If
    
End Function

Public Function ListaCreditosAmpliadosByNombre(ByVal psNombre As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    Set oConec = New COMConecta.DCOMConecta
    sSql = "Select P.cPersNombre,CA.cCtaCod"
    sSql = sSql & " From Persona P"
    sSql = sSql & " Inner Join ProductoPersona PP on P.cPersCod=PP.cPersCod and PP.nPrdPersRelac=20"
    sSql = sSql & " Inner Join ColocacAmpliado CA on CA.cCtaCodAmp=PP.cCtaCod"
    sSql = sSql & " Inner Join Producto PRO on PRO.cCtaCod=CA.cCtaCodAmp"
    sSql = sSql & " Where P.cPersNombre like '%" & psNombre & "%' and PRO.nPrdEstado=2020"
    sSql = sSql & "Group By P.cPersNombre,CA.cCtaCod"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCreditosAmpliadosByNombre = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaCreditosBycCtaCodNew(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select cCtaCodAmp From ColocacAmpliado Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCreditosBycCtaCodNew = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ValidarCreditoGarantizar(ByVal psCtaCod As String) As Boolean
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim bValor As Boolean
    
    bValor = False
    
    sSql = "Select Count(*) as nCantidad"
    sSql = sSql & " From ColocacAmpliado CA"
    sSql = sSql & " Inner Join Producto P on P.cCtaCod=CA.cCtaCod"
    sSql = sSql & " Where P.nPrdEstado in (2001,2000) and CA.cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
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

Public Function ListaGarantiaCtaAmpliado(ByVal psCtaCod As String, ByVal psCtaCodAmp As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select G.cNumGarant,cDescripcion,CN.cConsDescripcion,CG.nGravado"
    sSql = sSql & " From ColocacAmpliado CA"
    sSql = sSql & " Inner Join ColocGarantia CG on CG.cCtaCod=CA.cCtaCodAmp"
    sSql = sSql & " Inner Join Garantias G on G.cNumGarant=CG.cNumGarant"
    sSql = sSql & " Inner Join Constante CN on CN.nConsCod=1027 and CN.nConsValor<>1027 and CN.nConsValor=G.nTpoGarantia"
    sSql = sSql & " Where CA.cCtaCod='" & psCtaCod & "' and CA.cCtaCodAmp='" & psCtaCodAmp & "'"

    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaGarantiaCtaAmpliado = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
End Function

Public Function InsertarGarantiaAmpliado(ByVal prs As ADODB.Recordset) As Boolean
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim sCtaCod As String
    Dim sNumGarant As String
    Dim nGravado As Double
    
    On Error GoTo ErrHandler
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    
    If Not prs.EOF And Not prs.BOF Then
        prs.MoveFirst
        sNumGarant = prs(0)
        sCtaCod = prs(1)
        oConec.ConexionActiva.BeginTrans
         Do Until prs.EOF
             If sNumGarant = prs(0) And sCtaCod = prs(1) Then
                nGravado = nGravado + val(prs(3))
             Else
                sSql = "Insert into GarantAmpliado Values('" & sNumGarant & "','" & sCtaCod & "'," & nGravado & ")"
                oConec.ConexionActiva.Execute sSql
                nGravado = 0
                sNumGarant = prs(0)
                sCtaCod = prs(1)
             End If
            prs.MoveNext
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


Public Function ObtenerMontoColocar(ByVal psCtaCod As String) As Double
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    
    sSql = "Select Top 1 C.nMontoCol"
    sSql = sSql & " From Colocaciones C"
    sSql = sSql & " Inner Join Colocacampliado CA on CA.cCtaCod=C.cCtaCod"
    sSql = sSql & " Where C.cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
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
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    
    sSql = "Select nMoneda From Garantias Where cNumGarant='" & pscNumGarant & "'"
    
    Set oConec = New COMConecta.DCOMConecta
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
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    
    sSql = "Select nGravado  From ColocGarantia where cCtaCod='" & pcCtaCod & "' And cNumGarant='" & psNumGarant & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    If Not rs.EOF And Not rs.BOF Then
        ObteneMontoGarant = IIf(IsNull(rs!nGravado), 0, rs!nGravado)
    End If
    Set rs = Nothing
End Function

Public Function ObtenerListaGarantiasAmpliadas(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select * From GarantAmpliado Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerListaGarantiasAmpliadas = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function EliminarGarantiaAmpliada(ByVal psCtaCod As String) As Boolean
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    On Error GoTo ErrHandler
    
    sSql = "Delete From GarantAmpliado Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.ConexionActiva.Execute sSql
    oConec.CierraConexion
    Set oConec = Nothing
    EliminarGarantiaAmpliada = True
    Exit Function
ErrHandler:
    EliminarGarantiaAmpliada = False
End Function


Public Function ObtenerMontoGarantAmpliado(ByVal psCtaCod As String, ByVal pnMoneda As Integer) As Double
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select Sum(GA.nMontoGravado) as nGravado"
    sSql = sSql & " From GarantAmpliado GA"
    sSql = sSql & " Inner Join Garantias G on G.cNumGarant=GA.cNumGarant"
    sSql = sSql & " Where cCtaCod='" & psCtaCod & "' and G.nMoneda=" & pnMoneda
    
    Set oConec = New COMConecta.DCOMConecta
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

Public Sub ActualizandoGarantiasAmpliadas(ByVal psCtaCod As String)
   Dim oConec As COMConecta.DCOMConecta
   Dim sSql As String
   Dim rs As ADODB.Recordset
        
        
   sSql = "Select GA.*,G.nMoneda"
   sSql = sSql & " From GarantAmpliado GA"
   sSql = sSql & " Inner Join Garantias G on G.cNumGarant=GA.cNumGarant"
   sSql = sSql & " Where cCtaCod='" & psCtaCod & "'"
   
   Set oConec = New COMConecta.DCOMConecta
   oConec.AbreConexion
   Set rs = oConec.CargaRecordSet(sSql)
   
   If Not rs.EOF And Not rs.BOF Then
    oConec.ConexionActiva.BeginTrans
   Else
      Exit Sub
   End If
    
    Do Until rs.EOF
        sSql = "Insert Into ColocGarantia Values('" & rs!cNumGarant & "','" & psCtaCod & "'," & rs!nMoneda & "," & rs!nMontoGravado & ",1)"
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

'MADM 20110408
Public Function ObtenerCreditoPrincipalPol(ByVal psCtaCod As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim nCantidad As Integer
        
    sSql = "exec stp_Sel_CreditoParaAmpliado '" & psCtaCod & "'"
    
     Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerCreditoPrincipalPol = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'END MADM

'FRHU 20140424 TI-ERS015-2014
Public Function ObtenerCreditoXPersona(ByVal psPersCod As String, ByVal psMoneda As String, Optional ByVal cCodProd As String = "") As ADODB.Recordset 'JOEP20190919 ERS042 CP-2018 Agrego cCodProd
    Dim rs As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim nCantidad As Integer
        
    sSql = "exec stp_sel_CreditosParaAmpliacionPorCliente '" & psPersCod & "','" & psMoneda & "','" & cCodProd & "'" 'JOEP20190919 ERS042 CP-2018 Agrego cCodProd
    
     Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerCreditoXPersona = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'FIN FRHU 20140424
'INICIO EAAS20180924 SEGUN ACTA 103-2018
Public Function ValidaCreditoParaAmpliar(ByVal psCtaCod As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim nCantidad As Integer
    
    sSql = "EXEC stp_sel_ValidaCreditoParaAmpliar '" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        nCantidad = rs!nCantidad
    End If
    
    Set rs = Nothing
    
    If nCantidad > 0 Then
        ValidaCreditoParaAmpliar = True
    Else
        ValidaCreditoParaAmpliar = False
    End If
End Function
'FIN EAAS20180924 SEGUN ACTA 103-2018
