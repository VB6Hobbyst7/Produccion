VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCredReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RecuperaUltimoDiaMes(ByVal pdFecha As Date) As Integer
 
    Dim nUltimoDiaMes As Integer
    Dim nMes  As Integer
    Dim nAnio As Integer
    Dim nDia   As Integer
    
    nDia = Day(pdFecha)
    nMes = Month(pdFecha)
    nAnio = Year(pdFecha)

    nUltimoDiaMes = 0

    Select Case nMes
    Case 1, 3, 5, 7, 8, 10, 12
        nUltimoDiaMes = 31
    Case 4, 6, 9, 11
        nUltimoDiaMes = 30
    Case 2
        If nAnio Mod 4 = 0 Then  'Bisiesto
            nUltimoDiaMes = 29
        Else
            nUltimoDiaMes = 28
        End If
    End Select
    RecuperaUltimoDiaMes = nUltimoDiaMes
End Function

Public Function RecuperaListadoMorosos(ByVal psMonedas As String, ByVal psProductos As String, ByVal psAnalistas As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
     
    On Error GoTo ErrorRecuperaAgencias
 
    sSql = " SELECT C1.nMoneda, C1.cMoneda, C1.cCtaCod, c1Analista.CodAnalista, c1Analista.Analista, c1Analista.cNomAnalista, "
    sSql = sSql & " C1.cPersNombre, C1.cPersDireccDomicilio, C1.cUbiGeoDescripcion, C1.cPersTelefono, "
    sSql = sSql & " C1.nSaldo, C1.nDiasAtraso, C1FteIng.cDirFteIngreso, C1FteIng.cFonoFteIngreso, C1FteIng.cZonaFteIngreso,"
    sSql = sSql & " ISNULL(c1Garante.cNomGarante, '') AS cNomGarante, ISNULL(c1Garante.cDirGarante, '') AS cDirGarante, "
    sSql = sSql & " ISNULL(c1Garante.cZonaGarante, '') AS cZonaGarante "
    sSql = sSql & " From "
    sSql = sSql & " ("
    sSql = sSql & " SELECT SUBSTRING(C.cCtaCod, 9, 1) AS nMoneda, "
    sSql = sSql & " cMoneda = CASE WHEN SUBSTRING(C.cCtaCod, 9, 1) = '1' THEN 'SOLES' ELSE 'DOLARES' END, "
    sSql = sSql & " C.cCtaCod, PER.cPersNombre, PER.cPersDireccDomicilio, U.cUbiGeoDescripcion, PER.cPersTelefono, "
    sSql = sSql & " P.nSaldo , CC.nDiasAtraso "
    sSql = sSql & " FROM Colocaciones C INNER JOIN ColocacCred CC ON C.cCtaCod = CC.cCtaCod INNER JOIN"
    sSql = sSql & " Producto P ON C.cCtaCod = P.cCtaCod INNER JOIN "
    sSql = sSql & " ProductoPersona PP ON P.cCtaCod = PP.cCtaCod INNER JOIN "
    sSql = sSql & " Persona PER ON PP.cPersCod = PER.cPersCod INNER JOIN "
    sSql = sSql & " UbicacionGeografica U ON PER.cPersDireccUbiGeo = U.cUbiGeoCod "
    sSql = sSql & " Where (PP.NPRDPERSRELAC=" & gColRelPersTitular & ") "
    sSql = sSql & " AND (CC.nDiasAtraso > 0) "
    sSql = sSql & " AND (P.nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & "," & gColocEstRefNorm & "," & gColocEstRefMor & "," & gColocEstRefVenc & ")) "
    If Len(Trim(psMonedas)) > 0 Then
        sSql = sSql & " AND (SUBSTRING(C.cCtaCod, 9, 1) IN(" & psMonedas & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        sSql = sSql & " AND (SUBSTRING(C.cCtaCod, 6, 3) IN(" & psProductos & ")) "
    End If
    sSql = sSql & " ) C1 "
    
    sSql = sSql & " Inner Join "
    sSql = sSql & " ( SELECT CFI.cCtaCod, PFI.cRazSocDirecc AS cDirFteIngreso, PFI.cRazSocTelef AS cFonoFteIngreso, U.cUbiGeoDescripcion  as cZonaFteIngreso "
    sSql = sSql & " FROM ColocFteIngreso CFI INNER JOIN "
    sSql = sSql & " PersFteIngreso PFI ON CFI.cNumFuente = PFI.cNumFuente INNER JOIN "
    sSql = sSql & " UbicacionGeografica U ON PFI.cRazSocUbiGeo = U.cUbiGeoCod "
    If Len(Trim(psMonedas)) > 0 Then
        sSql = sSql & " WHERE (SUBSTRING(CFI.cCtaCod, 9, 1) IN(" & psMonedas & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        sSql = sSql & " AND (SUBSTRING(CFI.cCtaCod, 6, 3) IN(" & psProductos & ")) "
    End If
    sSql = sSql & " )  C1FteIng "
    sSql = sSql & " ON C1.cCtaCod=C1FteIng.cCtaCod "

    sSql = sSql & " Inner Join "
    sSql = sSql & " ( "
    sSql = sSql & " SELECT PRP.cCtaCod, PRP.CPERSCOD AS CodAnalista, RHH.CUSER AS Analista, PERSO.cPersNombre AS cNomAnalista "
    sSql = sSql & " FROM PRODUCTOPERSONA PRP "
    sSql = sSql & " INNER JOIN PERSONA PERSO ON PRP.CPERSCOD = PERSO.CPERSCOD "
    sSql = sSql & " INNER JOIN RRHH RHH ON PRP.CPERSCOD = RHH.CPERSCOD "
    sSql = sSql & " Where PRP.NPRDPERSRELAC=" & gColRelPersAnalista & " "
    If Len(Trim(psMonedas)) > 0 Then
        sSql = sSql & " AND (SUBSTRING(PRP.cCtaCod, 9, 1) IN(" & psMonedas & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        sSql = sSql & " AND (SUBSTRING(PRP.cCtaCod, 6, 3) IN(" & psProductos & ")) "
    End If
    sSql = sSql & " AND (PRP.CPERSCOD IN(" & psAnalistas & ")) "
    sSql = sSql & " GROUP BY PRP.cCtaCod, PRP.CPERSCOD, RHH.CUSER, PERSO.cPersNombre "
    sSql = sSql & " ) c1Analista "
    sSql = sSql & " ON C1.cCtaCod=c1Analista.cCtaCod "

    sSql = sSql & " Left Join "
    sSql = sSql & " ( "
    sSql = sSql & " SELECT PRP.cCtaCod, Perso.cpersnombre AS cNomGarante, perso.cpersdireccdomicilio AS cDirGarante, UBI.cUbiGeoDescripcion as cZonaGarante "
    sSql = sSql & " FROM persona PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD "
    sSql = sSql & " INNER JOIN dbo.UbicacionGeografica UBI ON perso.cPersDireccUbiGeo = UBI.cUbiGeoCod "
    sSql = sSql & " Where PRP.NPRDPERSRELAC=" & gColRelPersGarante & " "
    If Len(Trim(psMonedas)) > 0 Then
        sSql = sSql & " AND (SUBSTRING(PRP.cCtaCod, 9, 1) IN(" & psMonedas & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        sSql = sSql & " AND (SUBSTRING(PRP.cCtaCod, 6, 3) IN(" & psProductos & ")) "
    End If
    sSql = sSql & " ) c1Garante "
    sSql = sSql & " ON C1.cCtaCod=c1Garante.cCtaCod "
    sSql = sSql & " ORDER BY c1Analista.cNomAnalista, C1.nMoneda, C1.nDiasAtraso DESC "
 
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaListadoMorosos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaAgencias:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
 
Public Function Recupera_ConsolidadoCarteraxAnalista(ByVal pnTipoCambio As Currency, ByVal pdFecFin As String, ByVal psMoneda As String, ByVal psProductos As String, ByVal psAgencias As String, ByVal psAnalistas) As ADODB.Recordset
   
Dim oConecta As DConecta
Dim sCondicion As String
Dim oTipCambio As nTipoCambio
'Obtencion de la fecha del mes pasado
Dim dFechaPasada As String
Dim nTipoCambioPasado As Currency
Dim lsBaseConsol As String
Dim lsSQL As String
Dim loDataRep As dColPFunciones
Dim lrDataRep As ADODB.Recordset

    lsSQL = "Select nConsSisValor FROM ConstSistema WHERE nConsSisCod = 43"

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
        lsBaseConsol = lrDataRep!nConsSisValor
    Set loDataRep = Nothing
 
On Error GoTo ErrorRecuperaConsolidado
 
    dFechaPasada = DateAdd("d", -1, "01/" & Mid(Format(pdFecFin, "dd/MM/YYYY"), 4, 7))
    
    'Ya tengo la fecha del ultimo dia del mes pasado     'La fecha actual es la que pasan como parametro
    'dFechaPasada = Format(dFechaPasada, "dd/MM/YYYY")
      
    Set oTipCambio = New nTipoCambio
    nTipoCambioPasado = Format(oTipCambio.EmiteTipoCambio(dFechaPasada, TCFijoMes), "0.00")
    Set oTipCambio = Nothing
     
'1) Saldos de Cartera Vigente (Toda la cartera vigente) sin incluir judicial
 
    lsSQL = " SELECT 1 AS Lugar, COUNT(Producto.cCtaCod) AS Cantidad, "
    lsSQL = lsSQL & " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Producto.nSaldo "
    lsSQL = lsSQL & " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Producto.nSaldo * " & pnTipoCambio & " "
    lsSQL = lsSQL & " End) AS Total, RRHH.cUser AS cAnalista, Persona.cPersNombre as cNomAnalista, "
    lsSQL = lsSQL & " SUBSTRING(Producto.cCtaCod, 4, 2) AS cAgencia, Agencias.cAgeDescripcion AS cDesAgencia "
    lsSQL = lsSQL & " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN "
    lsSQL = lsSQL & " PERSONA ON ProductoPersona.cPersCod= Persona.cPersCod "
    lsSQL = lsSQL & " INNER JOIN RRHH ON Persona.cPersCod = RRHH.cPersCod "
    lsSQL = lsSQL & " INNER JOIN Agencias ON SUBSTRING(Producto.cCtaCod, 4, 2) = Agencias.cAgeCod "
    lsSQL = lsSQL & " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') "
    lsSQL = lsSQL & " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) "
    lsSQL = lsSQL & " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) "
    lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 4, 2) IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & psProductos & ")) "
    End If
    
    lsSQL = lsSQL & " GROUP BY RRHH.cUser, Persona.cPersNombre, SUBSTRING(Producto.cCtaCod, 4, 2), Agencias.cAgeDescripcion "
 
 
 lsSQL = lsSQL & " UNION ALL "



'2) Saldos de Cartera Vencida

 'Se consideran: Los Vigentes Vencidos(2021) => Se sacara de Producto
    '               Para los Vigentes Morosos(2022) => Para los productos 3 y 4 Se sacara las cuotas Pendientes con fecha vencida
    '               No se considera los productos  1 y 2 porque no existen con esas condiciones (Se obvia)
    '               * De considerarse los productos 1 y 2 se sacarian de la tabla Producto

    lsSQL = lsSQL & " Select '2' as cLugar,  SUM(T.Cantidad) as Cantidad, SUM(T.Total) as Total, "
    lsSQL = lsSQL & " T.cAnalista as cAnalista, T.cNomAnalista, T.cCodAgencia as cAgencia, T.cDesAgencia "
    lsSQL = lsSQL & " From  ( "
    lsSQL = lsSQL & " Select "
    lsSQL = lsSQL & " COUNT(P.cCtaCod) as Cantidad, "
    lsSQL = lsSQL & " SUM(case when SUBSTRING(p.cCtaCod, 9, 1)='" & gMonedaNacional & "' then P.nSaldo "
    lsSQL = lsSQL & " when SUBSTRING(P.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then P.nSaldo * " & pnTipoCambio & " End)  as Total, "
    lsSQL = lsSQL & " RH.cUser AS cAnalista, Per.cPersNombre as cNomAnalista, "
    lsSQL = lsSQL & " SUBSTRING(PP.cCtaCod, 4, 2) AS cCodAgencia, "
    lsSQL = lsSQL & " A.cAgeDescripcion AS cDesAgencia "
    lsSQL = lsSQL & " From producto P "
    lsSQL = lsSQL & " INNER JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN RRHH RH ON PP.cPersCod = RH.cPersCod "
    lsSQL = lsSQL & " INNER JOIN Persona Per ON PP.cPersCod = Per.cPersCod "
    lsSQL = lsSQL & " INNER JOIN  Agencias A ON SUBSTRING(PP.cCtaCod, 4, 2) = A.cAgeCod "
    lsSQL = lsSQL & " Where PP.nPrdPersRelac =" & gColRelPersAnalista & " "
    lsSQL = lsSQL & " AND PP.cPersCod IN (" & psAnalistas & ") "
    lsSQL = lsSQL & " AND p.nPrdEstado  in ('" & gColocEstVigVenc & "') "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ") "
    End If
    lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 4, 2) IN (" & psAgencias & ") "
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 6, 3) IN (" & psProductos & ") "
    End If
    lsSQL = lsSQL & " Group By RH.cUser, Per.cPersNombre, "
    lsSQL = lsSQL & " SUBSTRING(PP.cCtaCod, 4, 2), A.cAgeDescripcion "
 
    lsSQL = lsSQL & " Union All "

    lsSQL = lsSQL & " Select COUNT(DISTINCT P.cCtaCod) as nCantidad, "
    
    lsSQL = lsSQL & " SUM(case when SUBSTRING(p.cCtaCod, 9, 1)='" & gMonedaNacional & "' then (CCD.nMonto - CCD.nMontoPagado) "
    lsSQL = lsSQL & " when SUBSTRING(P.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then ((CCD.nMonto * " & pnTipoCambio & ") - (CCD.nMontoPagado * " & pnTipoCambio & ")) End)  as Total, "

    lsSQL = lsSQL & " RH.cUser AS cAnalista, Per.cPersNombre as cNomAnalista, "
    lsSQL = lsSQL & " SUBSTRING(PP.cCtaCod, 4, 2) AS cCodAgencia, "
    lsSQL = lsSQL & " A.cAgeDescripcion AS cDesAgencia "
    lsSQL = lsSQL & " FROM Producto P "
    lsSQL = lsSQL & " INNER JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN RRHH RH ON PP.cPersCod = RH.cPersCod "
    lsSQL = lsSQL & " INNER JOIN Persona Per ON PP.cPersCod = Per.cPersCod"
    lsSQL = lsSQL & " INNER JOIN ColocacCred CCR ON P.cCtaCod=CCR.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Coloccalendario CC ON CC.cCtaCod=CCR.cCtaCod AND CC.nNroCalen = CCR.nNroCalen "
    lsSQL = lsSQL & " INNER JOIN ColocCalendDet CCD ON CC.cCtaCod=CCD.cCtaCod  AND CC.nNroCalen = CCD.nNroCalen "
    lsSQL = lsSQL & " AND CC.nColocCalendApl=CCD.nCOlocCalendApl AND CC.nCuota = CCD.nCuota "
    lsSQL = lsSQL & " INNER JOIN  Agencias A ON SUBSTRING(PP.cCtaCod, 4, 2) = A.cAgeCod "
    lsSQL = lsSQL & " Where PP.nPrdPersRelac =" & gColRelPersAnalista & " "
    lsSQL = lsSQL & " AND PP.cPersCod IN (" & psAnalistas & ") "
    lsSQL = lsSQL & " and CC.ncoloccalendapl=" & gColocCalendAplCuota & " " '- -1 Cuota
    lsSQL = lsSQL & " and convert(varchar(8), CC.dvenc, 112)<='" & Format(pdFecFin, "YYYYMMdd") & "' "
    lsSQL = lsSQL & " and CC.nColocCalendEstado=" & gColocCalendEstadoPendiente & " " '-- Vigente 0
    lsSQL = lsSQL & " AND CCD.nPrdConceptoCod=" & gColocConceptoCodCapital & " " '-- 1000 Capital
    lsSQL = lsSQL & " AND P.nPrdEstado in('" & gColocEstVigMor & "') "
    lsSQL = lsSQL & " and substring(P.cCtaCod,6,1) IN ('3','4') "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ") "
    End If
    lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 4, 2) IN (" & psAgencias & ") "
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 6, 3) IN (" & psProductos & ") "
    End If
    lsSQL = lsSQL & " Group By RH.cUser, Per.cPersNombre, "
    lsSQL = lsSQL & " SUBSTRING(PP.cCtaCod, 4, 2), "
    lsSQL = lsSQL & " A.cAgeDescripcion "

    lsSQL = lsSQL & " Union All "

    lsSQL = lsSQL & " select COUNT(DISTINCT P.cCtaCod) as nCantidad, "
    lsSQL = lsSQL & " SUM(case when SUBSTRING(p.cCtaCod, 9, 1)='" & gMonedaNacional & "' then (CCD.nMonto - CCD.nMontoPagado) "
    lsSQL = lsSQL & " when SUBSTRING(P.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then ((CCD.nMonto * " & pnTipoCambio & ") - (CCD.nMontoPagado * " & pnTipoCambio & ")) End)  as Total, "

    lsSQL = lsSQL & " RH.cUser AS cAnalista, Per.cPersNombre as cNomAnalista, "
    lsSQL = lsSQL & " SUBSTRING(PP.cCtaCod, 4, 2) AS cCodAgencia, "
    lsSQL = lsSQL & " A.cAgeDescripcion AS cDesAgencia "
    lsSQL = lsSQL & " FROM Producto P "
    lsSQL = lsSQL & " INNER JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN RRHH RH ON PP.cPersCod = RH.cPersCod "
    lsSQL = lsSQL & " INNER JOIN Persona Per ON PP.cPersCod = Per.cPersCod "
    lsSQL = lsSQL & " INNER JOIN ColocacCred CCR ON P.cCtaCod=CCR.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Coloccalendario CC ON CC.cCtaCod=CCR.cCtaCod AND CC.nNroCalen = CCR.nNroCalPar "
    lsSQL = lsSQL & " INNER JOIN ColocCalendDet CCD ON CC.cCtaCod=CCD.cCtaCod  AND CC.nNroCalen = CCD.nNroCalen "
    lsSQL = lsSQL & " AND CC.nColocCalendApl=CCD.nCOlocCalendApl  AND CC.nCuota = CCD.nCuota "
    lsSQL = lsSQL & " INNER JOIN  Agencias A ON SUBSTRING(PP.cCtaCod, 4, 2) = A.cAgeCod "
    lsSQL = lsSQL & " Where PP.nPrdPersRelac =" & gColRelPersAnalista & " "
    lsSQL = lsSQL & " AND PP.cPersCod IN (" & psAnalistas & ") "
    lsSQL = lsSQL & " and CC.ncoloccalendapl=" & gColocCalendAplCuota & " " '- -1 Cuota
    lsSQL = lsSQL & " and convert(varchar(8), CC.dvenc, 112)<='" & Format(pdFecFin, "YYYYMMdd") & "' "
    lsSQL = lsSQL & " and CC.nColocCalendEstado=" & gColocCalendEstadoPendiente & " " '-- Vigente 0
    lsSQL = lsSQL & " AND CCD.nPrdConceptoCod=" & gColocConceptoCodCapital & " " '-- 1000 Capital
    lsSQL = lsSQL & " AND P.nPrdEstado in('" & gColocEstVigMor & "')  and substring(P.cCtaCod,6,1) IN ('3','4') "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ") "
    End If
    lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 4, 2) IN (" & psAgencias & ") "
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 6, 3) IN (" & psProductos & ") "
    End If
    lsSQL = lsSQL & " Group By RH.cUser, Per.cPersNombre, "
    lsSQL = lsSQL & " SUBSTRING(PP.cCtaCod, 4, 2), A.cAgeDescripcion "
    lsSQL = lsSQL & ") T  Group By T.cAnalista , T.cNomAnalista, T.cCodAgencia, T.cDesAgencia "


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'3) Saldo de Cartera Refinanciada (Todos los Refinanciados(Normal, Vencido, Moroso)

lsSQL = lsSQL & " UNION ALL "

lsSQL = lsSQL & " SELECT 3 AS Lugar, COUNT(Producto.cCtaCod) AS Cantidad, "
    lsSQL = lsSQL & " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Producto.nSaldo "
    lsSQL = lsSQL & " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Producto.nSaldo * " & pnTipoCambio & " "
    lsSQL = lsSQL & " End) AS Total, RRHH.cUser AS cAnalista, Persona.cPersNombre as cNomAnalista, "
    lsSQL = lsSQL & " SUBSTRING(Producto.cCtaCod, 4, 2) AS cAgencia, Agencias.cAgeDescripcion AS cDesAgencia "
    lsSQL = lsSQL & " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN "
    lsSQL = lsSQL & " PERSONA ON ProductoPersona.cPersCod= Persona.cPersCod "
    lsSQL = lsSQL & " INNER JOIN RRHH ON Persona.cPersCod = RRHH.cPersCod "
    lsSQL = lsSQL & " INNER JOIN Agencias ON SUBSTRING(Producto.cCtaCod, 4, 2) = Agencias.cAgeCod "
    lsSQL = lsSQL & " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') "
    lsSQL = lsSQL & " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) "
    lsSQL = lsSQL & " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "')) "
    lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 4, 2) IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser, Persona.cPersNombre, SUBSTRING(Producto.cCtaCod, 4, 2), Agencias.cAgeDescripcion "
 
 
'4) Obtengo los ingresos a recuperaciones
   '--Considero la suma de COlocacEstado.Monto para ColocacEstado.nPrdEstado=2201
  
  ' 4 OK
  
    lsSQL = lsSQL & " UNION ALL "
    
 
    lsSQL = lsSQL & " SELECT 4 AS Lugar, COUNT(Producto.cCtaCod) AS Cantidad, "
    lsSQL = lsSQL & " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Producto.nSaldo "
    lsSQL = lsSQL & " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Producto.nSaldo * " & pnTipoCambio & " "
    lsSQL = lsSQL & " End) AS Total, RRHH.cUser AS cAnalista, Persona.cPersNombre as cNomAnalista, "
    lsSQL = lsSQL & " SUBSTRING(Producto.cCtaCod, 4, 2) AS cAgencia, Agencias.cAgeDescripcion AS cDesAgencia "
    lsSQL = lsSQL & " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN "
    lsSQL = lsSQL & " PERSONA ON ProductoPersona.cPersCod= Persona.cPersCod "
    lsSQL = lsSQL & " INNER JOIN RRHH ON Persona.cPersCod = RRHH.cPersCod "
    lsSQL = lsSQL & " INNER JOIN Agencias ON SUBSTRING(Producto.cCtaCod, 4, 2) = Agencias.cAgeCod "
    lsSQL = lsSQL & " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') "
    lsSQL = lsSQL & " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) "
    lsSQL = lsSQL & " AND (Producto.nPrdEstado IN ('" & gColocEstRecVigJud & "')) "
    lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 4, 2) IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser, Persona.cPersNombre, SUBSTRING(Producto.cCtaCod, 4, 2), Agencias.cAgeDescripcion "


'''' D E S E M B O L S O S


'
'3) Obtengo desembolsos nuevos      La fecha es el rango entre el 31 del mes pasado a la fecha consultada y la fecha consultada
                                   ' Aca busco los coloccondicion=1
' para la cartera vigente

    'ok
    lsSQL = lsSQL & " UNION ALL "
    
    lsSQL = lsSQL & " SELECT 5 AS Lugar, COUNT(Colocaciones.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Colocaciones.nMontoCol " & _
    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Colocaciones.nMontoCol * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS cAnalista, Persona.cPersNombre as cNomAnalista, " & _
    " SUBSTRING(Producto.cCtaCod, 4, 2) AS cAgencia, Agencias.cAgeDescripcion AS cDesAgencia " & _
    " FROM Producto INNER JOIN ProductoPersona ON Producto.cCtaCod = ProductoPersona.cCtaCod INNER JOIN " & _
    " Persona ON ProductoPersona.cPersCod = Persona.cPersCod INNER JOIN RRHH ON Persona.cPersCod = RRHH.cPersCod INNER JOIN " & _
    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod INNER JOIN ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod " & _
    " INNER JOIN Agencias ON SUBSTRING(Producto.cCtaCod, 4, 2) = Agencias.cAgeCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (SUBSTRING(Producto.cCtaCod, 4, 2) IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & psProductos & ")) "
    End If
    
    lsSQL = lsSQL & " AND (ColocacCred.nColocCondicion=" & gColocCredCondNormal & ") "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)>'" & Format(dFechaPasada, "YYYYMMdd") & "') "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)<= '" & Format(pdFecFin, "YYYYmmdd") & "') "
    lsSQL = lsSQL & " GROUP BY RRHH.cUser, Persona.cPersNombre, SUBSTRING(Producto.cCtaCod, 4, 2), Agencias.cAgeDescripcion "
  
' 4) Obtengo desembolosos represtamos                La fecha es el rango entre el 31 del mes pasado a la fecha consultada y la fecha consultada
'                                                    'Aca busco los coloccondicion<>1
' para la cartera vigente

    lsSQL = lsSQL & " UNION ALL "
    
    lsSQL = lsSQL & " SELECT 6 AS Lugar, COUNT(Colocaciones.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Colocaciones.nMontoCol " & _
    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Colocaciones.nMontoCol * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS cAnalista, Persona.cPersNombre as cNomAnalista, " & _
    " SUBSTRING(Producto.cCtaCod, 4, 2) AS cAgencia, Agencias.cAgeDescripcion AS cDesAgencia " & _
    " FROM Producto INNER JOIN ProductoPersona ON Producto.cCtaCod = ProductoPersona.cCtaCod INNER JOIN " & _
    " Persona ON ProductoPersona.cPersCod = Persona.cPersCod INNER JOIN RRHH ON Persona.cPersCod = RRHH.cPersCod INNER JOIN " & _
    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod INNER JOIN ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod " & _
    " INNER JOIN Agencias ON SUBSTRING(Producto.cCtaCod, 4, 2) = Agencias.cAgeCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (SUBSTRING(Producto.cCtaCod, 4, 2) IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & psProductos & ")) "
    End If
    
'    lssql = lssql & " AND (ColocacCred.nColocCondicion<>" & gColocCredCondNormal & ") " & _
 '   " AND Colocaciones.dVigencia >= '" & Format(dFechaPasada, "mm/dd/yyyy") & "' AND Colocaciones.dVigencia <= '" & Format(pdFecFin, "mm/dd/yyyy") & "' " & _
  '
  
    lsSQL = lsSQL & " AND (ColocacCred.nColocCondicion<>" & gColocCredCondNormal & ") "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)>'" & Format(dFechaPasada, "YYYYMMdd") & "') "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)<= '" & Format(pdFecFin, "YYYYmmdd") & "') "
    lsSQL = lsSQL & " GROUP BY RRHH.cUser, Persona.cPersNombre, SUBSTRING(Producto.cCtaCod, 4, 2), Agencias.cAgeDescripcion "

  
'5) Obtengo los desembolsos Refinanciados            La fecha es el rango entre el 31 del mes pasado a la fecha consultada y la fecha consultada
                                                    'Aca busco el estado 2030 a 2032
' para la cartera vigente

    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 7 AS Lugar, COUNT(Colocaciones.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Colocaciones.nMontoCol " & _
    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Colocaciones.nMontoCol * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS cAnalista, Persona.cPersNombre as cNomAnalista, " & _
    " SUBSTRING(Producto.cCtaCod, 4, 2) AS cAgencia, Agencias.cAgeDescripcion AS cDesAgencia " & _
    " FROM Producto INNER JOIN ProductoPersona ON Producto.cCtaCod = ProductoPersona.cCtaCod INNER JOIN " & _
    " Persona ON ProductoPersona.cPersCod = Persona.cPersCod INNER JOIN RRHH ON Persona.cPersCod = RRHH.cPersCod INNER JOIN " & _
    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod " & _
    " INNER JOIN Agencias ON SUBSTRING(Producto.cCtaCod, 4, 2) = Agencias.cAgeCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "')) " & _
    " AND (SUBSTRING(Producto.cCtaCod, 4, 2) IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & psProductos & ")) "
    End If
    
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)>'" & Format(dFechaPasada, "YYYYMMdd") & "') "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)<= '" & Format(pdFecFin, "YYYYmmdd") & "') "
    lsSQL = lsSQL & " GROUP BY RRHH.cUser, Persona.cPersNombre, SUBSTRING(Producto.cCtaCod, 4, 2), Agencias.cAgeDescripcion  "
    
    lsSQL = lsSQL & " ORDER BY SUBSTRING(dbo.Producto.cCtaCod, 4, 2), rrhh.cUser"
    
    'Lugar: 1 si es Cartera Normal, 2 Si es Cartera Vencida, 3 Si es CArtera Refinanciada
    '       4 Si es Cobranza Judicial, 5 Si es Desembolsos Nuevos, 6 Si es DEsembolsos Represtamos
    '       ' 7 Si es Desembolsos Refinanciados
 

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set Recupera_ConsolidadoCarteraxAnalista = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaConsolidado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaDatosCartasWORD(ByVal pcTipoCadena As Byte, ByVal pcCondicion, ByVal pcMoneda As String, ByVal pcProductos As String, ByVal pcAnalistas, ByVal pnDiaIni As Integer, ByVal pnDiaFin As Integer, ByVal pnNota1 As Integer, ByVal pnNota2 As Integer, ByVal pnTipoCuotas As Integer, ByVal pnCuotas As Integer) As ADODB.Recordset
         
Dim sSql As String
Dim oConecta As DConecta

' pctipocadena=0 es el reporte mas simple : AVISO
' pcTipoCadena=1 es el reporte para todas las cartas de morosos menos la primera
' pcTipoCadena=2 es el reporte para paralelo

sSql = "SELECT P.cCtaCod, P.cPersNombre, P.cPersDireccDomicilio, P.cUbiGeoDescripcion, "
If pcTipoCadena = 1 Or pcTipoCadena = 2 Then 'quiere decir que incluye la cuota para el reporte
     sSql = sSql & " Cal.nCuota, P.nDiasAtraso, ISNULL(FIA.cDesFiador,'') AS cDesFiador, ISNULL(FIA.cDireccionFiador,'') AS cDireccionFiador, ISNULL(FIA.cZonaFiador,'') AS cZonaFiador, " & _
       " P.cMoneda, P.cDesMoneda, NEGOCIO.cRazSocDirecc, "
    
    If pcTipoCadena = 2 Then
        sSql = sSql & " Cuot.nCuotasPend, isnull(cNotas.nColocNota,0) nColocNota, "
    End If
 
End If

    sSql = sSql & " An.cAnalista, An.cDesAnalista FROM " & _
    " (select distinct Producto.cCtaCod, ProductoPersona.cPersCod, ColocacCred.nDiasAtraso, Persona.cPersNombre, Persona.cPersDireccDomicilio, " & _
    " SUBSTRING(Producto.cCtaCod, 9, 1) as cMoneda, Constante.cConsDescripcion as cDesMoneda, UbicacionGeografica.cUbiGeoDescripcion, SUBSTRING(Producto.cCtaCod, 6, 3) AS cProducto , ColocacCred.nNroCalen " & _
    " FROM Producto INNER JOIN ProductoPersona ON Producto.cCtaCod = ProductoPersona.cCtaCod INNER JOIN " & _
    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod INNER JOIN " & _
    " ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod INNER JOIN " & _
    " Persona ON ProductoPersona.cPersCod = Persona.cPersCod INNER JOIN " & _
    " UbicacionGeografica ON Persona.cPersDireccUbiGeo = UbicacionGeografica.cUbiGeoCod " & _
    " INNER JOIN CONSTANTE ON SUBSTRING(Producto.cCtaCod, 9, 1) = Constante.nConsValor " & _
    " WHERE (Producto.nPrdEstado IN ('" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "')) " & _
    " AND (ColocacCred.nDiasAtraso >=" & pnDiaIni & " AND ColocacCred.nDiasAtraso<=" & pnDiaFin & ") " & _
    " AND (ProductoPersona.nPrdPersRelac=" & gColRelPersTitular & ") "
    
    If Len(Trim(pcMoneda)) > 0 Then
        sSql = sSql & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & pcMoneda & ")) "
    End If
    If Len(Trim(pcProductos)) > 0 Then
        sSql = sSql & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & pcProductos & ")) "
    End If
    
    sSql = sSql & " AND (ColocacCred.nColocCondicion IN (" & pcCondicion & ")) " & _
     " AND (Constante.nConsCod=" & gMoneda & ")) P Inner Join " & _
     " ( SELECT  distinct PRP.cCtaCod, RHH.CUSER as cAnalista, PERSO.cPersNombre as cDesAnalista " & _
     " FROM RRHH RHH INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD " & _
     " INNER JOIN PERSONA PERSO ON PRP.CPERSCOD = PERSO.CPERSCOD Where PRP.nPrdPersRelac =" & gColRelPersAnalista & _
     " AND PRP.cPersCod IN(" & pcAnalistas & ") " & _
     " ) as An ON P.cCtaCod=An.cCtaCod "
     
If pcTipoCadena = 1 Or pcTipoCadena = 2 Then
     sSql = sSql & " INNER JOIN (Select cCtaCod, nNroCalen, Max(nCuota) as nCuota From ColocCalendario " & _
        " Where nColocCalendEstado=" & gColocCalendEstadoPendiente & " AND nColocCalendApl=" & gColocCalendAplCuota & " Group By cCtaCod, nNroCalen) " & _
        " as Cal ON Cal.cCtaCod = P.cCtaCod And Cal.nNroCalen = P.nNroCalen "

    'Aca obtengo la direccion del negocio
    sSql = sSql & " LEFT JOIN (select PER1.cPersCod, PFIN1.cRazSocDirecc " & _
            " FROM Persona PER1 INNER JOIN PersFteIngreso PFIN1 ON PER1.cPersCod = PFIN1.cPersCod " & _
            " WHERE PFIN1.cUltimaActualizacion = (SELECT MAX(PFIN2.cUltimaActualizacion) As cUltimo " & _
            " FROM Persona PER2 INNER JOIN PersFteIngreso PFIN2 ON PER2.cPersCod = PFIN2.cPersCod " & _
            " Where per1.cPersCod = PER2.cPersCod GROUP BY PER2.cPersCod)) AS NEGOCIO " & _
            " ON P.cPersCod=NEGOCIO.cPersCod "

        If pcTipoCadena = 2 Then
            'Cuotas solo para el reporte paralelo
            sSql = sSql & " Inner Join( SELECT dbo.ColocacCred.cCtaCod, " & _
                " ColocacEstado.nCuotas - ColocacCred.nNroProxCuota + 1 AS nCuotasPend " & _
                " FROM dbo.ColocacCred INNER JOIN dbo.ColocacEstado ON dbo.ColocacCred.cCtaCod = " & _
                " dbo.ColocacEstado.cCtaCod " & _
                " WHERE (dbo.ColocacEstado.nPrdEstado IN ('" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "')) "

                If pnTipoCuotas = 0 Then
                    sSql = sSql & " AND (ColocacEstado.nCuotas - ColocacCred.nNroProxCuota + 1) <= " & CInt(pnCuotas)
                ElseIf pnTipoCuotas = 1 Then
                    sSql = sSql & " AND round( (  100 * (ColocacEstado.nCuotas - ColocacCred.nNroProxCuota + 1) / ColocacEstado.nCuotas ) , 0) <= " & CInt(pnCuotas)
                End If
                
                
            sSql = sSql & " ) Cuot ON P.cCtaCod=Cuot.cCtaCod " & _
                " LEFT OUTER JOIN " & _
                " ( "
                 
            sSql = sSql & " SELECT  cCtaCod, MAX(ColocCalificacionAnalista.dColocNotaFecha) AS dFechaMax, " & _
                " ISNULL((Select Max(nColocNota) FROM ColocCalificacionAnalista CCA " & _
                " Where CCA.cCtaCod = ColocCalificacionAnalista.cCtaCod And CCA.dColocNotaFecha = Max(dbo.ColocCalificacionAnalista.dColocNotaFecha) " & _
                " ),0) as nColocNota FROM dbo.ColocCalificacionAnalista Where (nColocNota >=" & pnNota1 & " and nColocNota<=" & pnNota2 & ") " & _
                " GROUP BY cCtaCod ) cNotas ON P.cCtaCod=cNotas.cCtaCod  "
         
        End If

        sSql = sSql & " LEFT OUTER JOIN (SELECT PRP.cCtaCod, Perso.cPersNombre  as cDesFiador, " & _
        " Perso.cPersDireccDomicilio as cDireccionFiador, UbicacionGeografica.cUbiGeoDescripcion as cZonaFiador " & _
        " FROM PERSONA Perso INNER JOIN ProductoPersona PRP ON PRP.CPersCod = PERSO.cPersCod INNER JOIN " & _
        " UbicacionGeografica ON Perso.cPersDireccUbiGeo = UbicacionGeografica.cUbiGeoCod " & _
        " WHERE PRP.nPrdPersRelac=" & gColRelPersGarante & ") AS FIA ON FIA.cCtaCod = P.cCtaCod "
    
End If
 
sSql = sSql & " ORDER BY P.nDiasAtraso DESC, An.cAnalista, P.cPersNombre "

Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosCartasWORD = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaConsolidado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function GetProductos() As ADODB.Recordset

Dim cn As New DConecta
Dim Rs As New ADODB.Recordset
Dim csql As String
  
'csql = " Select cGrupo=nConsCod, cValor=nconsvalor, cProducto= cconsdescripcion, " & _
'       " cNivel=case when nconsvalor in(select min(nconsvalor) from constante K where K.nconscod=C.nConscod AND substring(convert(varchar(3), K.nconsvalor),1,1) = substring(convert(varchar(3), C.nconsvalor),1,1)) " & _
'       " Then 1 Else 2 End " & _
'       " From constante C where C.nConsCod='1001' and nconsvalor not in(" & Producto.gCapAhorros & ", " & Producto.gCapPlazoFijo & ", " & Producto.gCapCTS & ", " & Producto.gColConsuPrendario & ", " & Producto.gColCFComercial & ", " & Producto.gColCFPYME & ") Order by nconsvalor"
       
csql = "select nAgruCod as cGrupo, nAgruCab as cValor, cAgruDes as cProducto, 1 as cNivel from RepAgruProd where nAgruCod=1 " & _
    " Union " & _
    " select nAgruCod as cGrupo, cProdCod as cValor, cProdDesc as cProducto,2  as cNivel   from  RepAgruProdDet where nAgruCod=1 " & _
    " order by cvalor "
       
cn.AbreConexion

Set Rs = cn.CargaRecordSet(csql)
Rs.ActiveConnection = Nothing
Set GetProductos = Rs
 
cn.CierraConexion

End Function
