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
Public Function ListaDatosAmpliacion(ByVal psCtaCod As String) As Recordset
    Dim sSQL As String
    Dim oCn As DConecta
    
    sSQL = " Select Pe.cPersNombre as Nombre,"
    sSQL = sSQL & " TipoProducto=Case Substring(P.cCtaCod,6,3)"
    sSQL = sSQL & " When '101' Then 'COMERCIAL EMPRESARIAL'"
    sSQL = sSQL & " When '102' Then 'COMERCIAL AGRICOLA'"
    sSQL = sSQL & " When '103' Then 'COMERCIAL PESQUERO'"
    sSQL = sSQL & " When '121' Then 'COMERCIAL CARTA FIANZA'"
    sSQL = sSQL & " When '201' Then 'MES EMPRESARIAL'"
    sSQL = sSQL & " When '202' Then 'MES AGRICOLA'"
    sSQL = sSQL & " When '301' Then 'CONSUMO DESCUENTO POR PLANILLA'"
    sSQL = sSQL & " When '302' Then 'CONSUMO PLAZO FIJO'"
    sSQL = sSQL & " When '303' Then 'CONSUMO CTS'"
    sSQL = sSQL & " When '304' Then 'CONSUMO PERSONAL DIRECTO'"
    sSQL = sSQL & " When '320' Then 'CONSUMO A DIRECTORES Y TRABAJADORES'"
    sSQL = sSQL & " When '401' Then ' HIPOTECARIO PARA VIVIENDA'"
    sSQL = sSQL & " When '403' Then 'HIPOTECARIO DIRECTORES Y TRABAJADORES'"
    sSQL = sSQL & " When '423' Then 'HIPOTECARIO MIVIVIENDA'"
    sSQL = sSQL & " end,"
    sSQL = sSQL & " Moneda=Case Substring(P.cCtaCod,9,1)"
    sSQL = sSQL & " When '1' Then 'Soles'"
    sSQL = sSQL & " When '2' Then 'Dolares'"
    sSQL = sSQL & " End"
    sSQL = sSQL & " From  Producto P"
    sSQL = sSQL & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod"
    sSQL = sSQL & " Inner Join Persona Pe on Pe.cPersCod=PP.cPerscod and PP.nPrdPersRelac=20"
    sSQL = sSQL & " Where P.cCtaCod='" & psCtaCod & "'"
    
        
    Set oCn = New DConecta
    oCn.AbreConexion
    Set ListaDatosAmpliacion = oCn.CargaRecordSet(sSQL)
    oCn.CierraConexion
    Set oConec = Nothing
End Function

Public Function VerificaAmpliado(ByVal psCtaCod As String) As Boolean
    Dim sSQL As String
    Dim oCn As DConecta
    Dim rs As ADODB.Recordset
    Dim nCantidad As Integer
    
    sSQL = "Select Count(*) as Cantidad From ColocacAmpliado Where cCtaCod='" & psCtaCod & "'"
    
    Set oCn = New DConecta
    oCn.AbreConexion
    
    Set rs = oCn.CargaRecordSet(sSQL)
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

Public Function ValidacionCredito(ByVal psCtaCod As String) As Boolean
    Dim sSQL As String
    Dim oCn  As DConecta
    Dim rs As ADODB.Recordset
    Dim nCantidad As Integer
    
    
    sSQL = sSQL & " Select Count(*) as Cantidad"
    sSQL = sSQL & " From Producto"
    sSQL = sSQL & " Where cCtaCod='" & psCtaCod & "' and"
    sSQL = sSQL & " nPrdEstado in (2020)"

    Set oCn = New DConecta
    oCn.AbreConexion
    Set rs = oCn.CargaRecordSet(sSQL)
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
    Dim sSQL As String
    
    sSQL = "Select distinct P.cCtaCod,Pers.cPersNombre"
    sSQL = sSQL & " From Persona Pers"
    sSQL = sSQL & " Inner Join ProductoPersona PP on PP.cPersCod=Pers.cPersCod"
    sSQL = sSQL & " Inner Join Producto P on PP.cCtaCod=P.cCtaCod"
    sSQL = sSQL & " Where P.nPrdEstado in (2020) and PP.nPrdPersRelac=20"
    sSQL = sSQL & " and PP.cPersCod='" & cPersCod & "'"
        
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCreditoPorAmpliar = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

' Function que emula la liberacion de la garantia para un credito ampliado

Public Function EmularGarantia(ByVal psCtaCod As String, ByVal nMoneda As Integer) As Double
    Dim oConec As DConecta
    Dim sSQL As String
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
    
    sSQL = "Select nMontoCol From Colocaciones Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        nMontoCol = rs!nMontoCol
    End If
    
    
    ' Se obtiene la sumatoria capitales pendientes
   
    sSQL = "Select Sum(CD.nMonto-CD.nMontoPagado) as nMontoCapital"
    sSQL = sSQL & " From Producto P"
    sSQL = sSQL & " Inner Join ColocCalendario CC on P.cCtaCod=CC.cCtaCod"
    sSQL = sSQL & " Inner Join ColocCalendDet CD on CD.cCtaCod=CC.cCtaCod and CC.nNroCalen=CD.nNroCalen  and"
    sSQL = sSQL & " CC.nColocCalendApl = CD.nColocCalendApl And CC.nCuota = CD.nCuota"
    sSQL = sSQL & " Where P.cCtaCod='" & psCtaCod & "' and CD.nPrdConceptoCod=1000 and CC.nColocCalendEstado=0 and"
    sSQL = sSQL & " CD.nColocCalendApl=1 and"
    sSQL = sSQL & " CC.nNroCalen=(Select Max(nNroCalen) From ColocCalendario Where cCtaCod='" & psCtaCod & "')"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.BOF And Not rs.EOF Then
        nMontoK = rs!nMontoCapital
    End If
    Set rs = Nothing
    
    'Obtiene el monto Total de la garantia
    Set oCred = New DCredito
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
            sSQL = "Select (G.nPorGravar-G.nGravament) as nDisponible"
            sSQL = sSQL & " From ColocGarantia CG"
            sSQL = sSQL & " Inner Join Garantias G on CG.cNumGarant=G.cNumGarant"
            sSQL = sSQL & " Inner Join Producto P on CG.cCtaCod=P.cCtaCod"
            sSQL = sSQL & " Where P.cCtaCod='" & psCtaCod & "'" & " and  G.cNumGarant='" & rs!cNumGarant & "'"
            
            Set rsAmpliado = New ADODB.Recordset
            
            Set oConec = New DConecta
            oConec.AbreConexion
            Set rsAmpliado = oConec.CargaRecordSet(sSQL)
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
    Dim oConec As DConecta
    Dim sSQL As String
    Dim nCantidad As Integer
    
    sSQL = "Select Count(*) as nCantidad From ColocacAmpliado Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
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
    Dim sSQL As String
    
    sSQL = "Select cCtaCodAmp From ColocacAmpliado Where cCtaCod='" & psCtaCodNew & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    If Not rs.EOF And Not rs.BOF Then
        GetcCtaCodPorAmpliar = IIf(IsNull(rs!cCtaCodAmp), "", rs!cCtaCodAmp)
    End If
    Set rs = Nothing
End Function

Public Function AsientoGarantiaAmpliado(ByVal psCtaCod As String) As Boolean
    Dim oConec As DConecta
    Dim oGeneral As DGeneral
    
    Dim sSQL As String
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
    
    sSQL = "Select cCtaCodAmp From ColocacAmpliado Where cCtaCod='" & psCtaCod & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    
    If Not rs.BOF And Not rs.EOF Then
        sCtaCodNew = rs!cCtaCodAmp
    End If
    
    Set rs = Nothing
    
    'obtener el monto de la aprobacion para comenzar a coberturar la garantia pe!!!!
    sSQL = "Select nMonto From ColocacEstado Where cCtaCod='" & psCtaCod & "'"
    
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        nMonto = rs!nMonto
    End If
    
    ' obtener el tipo de cambio
    Set oGeneral = New DGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdfecha, TCFijoMes)
    nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
    Set oGeneral = Nothing
    
    'obtener las garantias relacionadas con el nuevo credito
    
    sSQL = "Select cNumGarant,nGravado From ColocGarantia Where cCtacod='" & psCtaCod & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    
    Do Until rs.EOF
        If Mid(psCtaCod, 9, 1) = "1" Then 'Soles
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
    sSQL = "Select * From ColocGarantia Where cCtaCod='" & psCtaCodNew & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    
    Do Until rs.EOF
        'calculado el monto disponible de la garantia
        sSQL = "Select nPorGravar-nGravament as nDisponible,nMoneda From Garantias Where cNumGarant='" & rs!cNumGarant & "'"
        oConec.AbreConexion
        Set rsGarant = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
        
        If Not rsGarant.EOF And Not rsGarant.BOF Then
            nMontoDispGarant = rsGarant!nDisponible
        End If
        
        'insertando el colocgarantia y actualizando la garantia
        If nMonto > 0 Then
            'calculando el monto de la garantia que se va asignar
            If nMonto >= ObtieneMonto(Mid(psCtaCod, 9, 1), rs!nMoneda, nMontoDispGarant) Then
               nMontoAsig = nMontoDispGarant
               nMonto = nMonto - ObtieneMonto(Mid(psCtaCod, 9, 1), rs!nMoneda, nMontoAsig)
            Else
               nMontoAsig = nMonto
               nMonto = nMonto - ObtieneMonto(Mid(psCtaCod, 9, 1), rs!nMoneda, nMontoAsig)
            End If
            oConec.AbreConexion
            oConec.ConexionActiva.BeginTrans
            sSQL = "Insert Into ColocGarantia Values('" & rs!cNumGarant & "','" & psCtaCod & "'," & rs!nMoneda & "," & nMontoAsig & ",1)"
            oConec.ConexionActiva.Execute sSQL
            'actualizando la garantia
            sSQL = "Update ColocGarantia"
            sSQL = sSQL & " set nGravament=nGravament+" & nMontoAsig & ",nPorGravar=nPorGravar-" & nMontoAsig
            sSQL = sSQL & " Where cNumGarant='" & rs!cNumGarant
            oConec.ConexionActiva.Execute sSQL
            oConec.CommitTrans
        End If
        rs.MoveNext
    Loop
End Function

Function ObtieneMonto(ByVal pnMonedaCuentas As Integer, ByVal pnMonedaGarantia, pnMontoGarantia As Double) As Double
    Dim oGeneral As DGeneral
    Dim nTipoCambioFijo As Double
    
    Set oGeneral = New DGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdfecha, TCFijoMes)
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
    Dim sSQL As String
    
    Set oConec = New DConecta
    sSQL = "Select P.cPersNombre,CA.cCtaCod"
    sSQL = sSQL & " From Persona P"
    sSQL = sSQL & " Inner Join ProductoPersona PP on P.cPersCod=PP.cPersCod and PP.nPrdPersRelac=20"
    sSQL = sSQL & " Inner Join ColocacAmpliado CA on CA.cCtaCodAmp=PP.cCtaCod"
    sSQL = sSQL & " Inner Join Producto PRO on PRO.cCtaCod=CA.cCtaCodAmp"
    sSQL = sSQL & " Where P.cPersNombre like '%" & psNombre & "%' and PRO.nPrdEstado=2020"
    sSQL = sSQL & "Group By P.cPersNombre,CA.cCtaCod"
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCreditosAmpliadosByNombre = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaCreditosBycCtaCodNew(ByVal psCtaCod As String) As Recordset
    Dim oConec As DConecta
    Dim sSQL As String
    
    sSQL = "Select cCtaCodAmp From ColocacAmpliado Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCreditosBycCtaCodNew = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ValidarCreditoGarantizar(ByVal psCtaCod As String) As Boolean
    Dim oConec As DConecta
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim bValor As Boolean
    
    bValor = False
    
    sSQL = "Select Count(*) as nCantidad"
    sSQL = sSQL & " From ColocacAmpliado CA"
    sSQL = sSQL & " Inner Join Producto P on P.cCtaCod=CA.cCtaCod"
    sSQL = sSQL & " Where P.nPrdEstado in (2001,2000) and cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
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

Public Function ListaGarantiaCtaAmpliado(ByVal psCtaCod As String, ByVal psCtaCodAmp As String) As Recordset
    Dim oConec As DConecta
    Dim sSQL As String
    
    sSQL = "Select G.cNumGarant,cDescripcion,CN.cConsDescripcion,CG.nGravado"
    sSQL = sSQL & " From ColocacAmpliado CA"
    sSQL = sSQL & " Inner Join ColocGarantia CG on CG.cCtaCod=CA.cCtaCodAmp"
    sSQL = sSQL & " Inner Join Garantias G on G.cNumGarant=CG.cNumGarant"
    sSQL = sSQL & " Inner Join Constante CN on CN.nConsCod=1027 and CN.nConsValor<>1027 and CN.nConsValor=G.nTpoGarantia"
    sSQL = sSQL & " Where CA.cCtaCod='" & psCtaCod & "' and CA.cCtaCodAmp='" & psCtaCodAmp & "'"

    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaGarantiaCtaAmpliado = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
End Function

Public Function InsertarGarantiaAmpliado(ByVal prs As ADODB.Recordset) As Boolean
    Dim oConec As DConecta
    Dim sSQL As String
    Dim sCtaCod As String
    Dim sNumGarant As String
    Dim nGravado As Double
    
    On Error GoTo ErrHandler
    
    Set oConec = New DConecta
    oConec.AbreConexion
    
    If Not prs.EOF And Not prs.BOF Then
        prs.MoveFirst
        sNumGarant = prs(0)
        sCtaCod = prs(1)
        oConec.ConexionActiva.BeginTrans
         Do Until prs.EOF
             If sNumGarant = prs(0) And sCtaCod = prs(1) Then
                nGravado = nGravado + Val(prs(3))
             Else
                sSQL = "Insert into GarantAmpliado Values('" & sNumGarant & "','" & sCtaCod & "'," & nGravado & ")"
                oConec.ConexionActiva.Execute sSQL
                nGravado = 0
                sNumGarant = prs(0)
                sCtaCod = prs(1)
             End If
            prs.MoveNext
         Loop
         sSQL = "Insert into GarantAmpliado Values('" & sNumGarant & "','" & sCtaCod & "'," & nGravado & ")"
         oConec.ConexionActiva.Execute sSQL
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
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "Select Top 1 C.nMontoCol"
    sSQL = sSQL & " From Colocaciones C"
    sSQL = sSQL & " Inner Join Colocacampliado CA on CA.cCtaCod=C.cCtaCod"
    sSQL = sSQL & " Where C.cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
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
    Dim sSQL As String
    
    sSQL = "Select nMoneda From Garantias Where cNumGarant='" & pscNumGarant & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenerMonedaGarantia = IIf(IsNull(rs!nMoneda), 0, rs!nMoneda)
    End If
    Set rs = Nothing
End Function

Public Function ObteneMontoGarant(ByVal psNumGarant As String, ByVal pcCtacod As String) As Double
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "Select nGravado  From ColocGarantia where cCtaCod='" & pcCtacod & "' And cNumGarant='" & psNumGarant & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    If Not rs.EOF And Not rs.BOF Then
        ObteneMontoGarant = IIf(IsNull(rs!nGravado), 0, rs!nGravado)
    End If
    Set rs = Nothing
End Function

Public Function ObtenerListaGarantiasAmpliadas(ByVal psCtaCod As String) As Recordset
    Dim oConec As DConecta
    Dim sSQL As String
    
    sSQL = "Select * From GarantAmpliado Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ObtenerListaGarantiasAmpliadas = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function EliminarGarantiaAmpliada(ByVal psCtaCod As String) As Boolean
    Dim oConec As DConecta
    Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    sSQL = "Delete From GarantAmpliado Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    oConec.ConexionActiva.Execute sSQL
    oConec.CierraConexion
    Set oConec = Nothing
    EliminarGarantiaAmpliada = True
    Exit Function
ErrHandler:
    EliminarGarantiaAmpliada = False
End Function


Public Function ObtenerMontoGarantAmpliado(ByVal psCtaCod As String, ByVal pnMoneda As Integer) As Double
    Dim oConec As DConecta
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    sSQL = "Select Sum(GA.nGravado) as nGravado"
    sSQL = sSQL & " From GarantAmpliado GA"
    sSQL = sSQL & " Inner Join Garantias G on G.cNumGarant=GA.cNumGarant"
    sSQL = sSQL & " Where cCtaCod='" & psCtaCod & "' and G.nMoneda=" & pnMoneda
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
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
   Dim oConec As DConecta
   Dim sSQL As String
   Dim rs As ADODB.Recordset
        
        
   sSQL = "Select GA.*,G.nMoneda"
   sSQL = sSQL & " From GarantAmpliado GA"
   sSQL = sSQL & " Inner Join Garantias G on G.cNumGarant=GA.cNumGarant"
   sSQL = sSQL & " Where cCtaCod='" & psCtaCod & "'"
   
   Set oConec = New DConecta
   oConec.AbreConexion
   Set rs = oConec.CargaRecordSet(sSQL)
   
   If Not rs.EOF And Not rs.BOF Then
    oConec.ConexionActiva.BeginTrans
   Else
      Exit Sub
   End If
    
    Do Until rs.EOF
        sSQL = "Insert Into ColocGarantia Values('" & rs!cNumGarant & "','" & psCtaCod & "'," & rs!nMoneda & "," & rs!nGravado & ",1)"
        oConec.ConexionActiva.Execute sSQL
        sSQL = "Update Garantias"
        sSQL = sSQL & " Set nGravament=nGravament + " & rs!nGravado
        sSQL = sSQL & " Where cNumGarant='" & rs!cNumGarant & "'"
        oConec.ConexionActiva.Execute sSQL
     rs.MoveNext
    Loop
    Set rs = Nothing
    oConec.CommitTrans
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
