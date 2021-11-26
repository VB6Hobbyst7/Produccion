VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DPigContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim sql As String
Dim oConn As DConecta

'***********************************************************
' CMCPL - OBTIENE LAS BOLETAS/FACTURAS POR VENTAS DE JOYAS
' EAFA - 15/10/2002
'***********************************************************
Public Function dObtieneDocumentosVentasJoyas(ByVal pdFecDesde As String, ByVal pdFecHasta As String, _
           Optional ByVal pnTipoDoc As Integer = -1, Optional ByVal pnTipoVta As Integer = -1) As Recordset

Dim Rs As ADODB.Recordset

On Error GoTo dError

sql = " SELECT cDocumento, nMonto, cUltimaActualizacion, Substring(cUltimaActualizacion,1,8) dFechaVta, nCodTipo, MDP.nNroMov, " _
       & " nFlag = (SELECT M.nMovFlag FROM Mov M WHERE M.nMovNro = MDP.nNroMov ), " _
       & " Cliente = (SELECT P.cPersNombre FROM Persona P WHERE P.cPersCod = MDP.cPersCod), " _
       & " DesTipDoc = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigTipoDocumento & " AND nConsValor =nCodTipo), " _
       & " DesTipVta  = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigTipoVentaJoya & " AND nConsValor = nMotivo) " _
       & " FROM MovDocPig MDP" _
       & " WHERE substring(cUltimaActualizacion,1,8) between '" & pdFecDesde & "' and '" & pdFecHasta & "'" _
       & "   AND   nMotivo = Case When " & pnTipoVta & " = 0 Then nMotivo Else " & pnTipoVta & " end " _
       & "   AND nCodTipo = Case When " & pnTipoDoc & " = 0 Then nCodTipo Else " & pnTipoDoc & " end "
   
Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneDocumentosVentasJoyas = Rs
Set Rs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos de Joyas en <<dObtieneDocumentosVentasJoyas>>", Err.Description
 
End Function

Private Sub Class_Initialize()
Dim oIni As ClasIni
Dim sConn As String

    Set oIni = New ClasIni
    sConn = oIni.CadenaConexion
    Set oIni = Nothing
    
Set oConn = New DConecta
If oConn.AbreConexion(sConn) = False Then
    Call RaiseError(MyUnhandledError, "DPigContrato:Initialize. Error en Conexion a Base de datos")
End If

End Sub
Private Sub Class_Terminate()
    oConn.CierraConexion
    Set oConn = Nothing
End Sub

Public Function dObtieneDatosCreditoPignoraticio(ByVal psCuenta As String) As Recordset
'Obtiene Datos de Contrato Pignoraticio
Dim lrs As ADODB.Recordset

On Error GoTo dError

sql = "SELECT P.cCtaCod, P.nSaldo, P.nPrdEstado, P.nTasaInteres, " _
    & " C.nPlazo, C.dVenc, C.nMontoCol, C.dVigencia, " _
    & "  CP.nPiezas,  " _
    & " nPlazoIni = (SELECT nPlazo FROM ColocacEstado WHERE cCtaCod ='" & psCuenta & "' " _
    & "             AND nPrdEstado = " & gPigEstRegis & "  AND nPlazo > 0 ),  " _
    & " dFecVencIni = ( SELECT DateAdd(d, nPlazo, dPrdEstado) FROM ColocacEstado WHERE cCtaCod ='" & psCuenta & "' " _
    & "                 AND nPrdEstado = " & gPigEstRegis & "  AND nPlazo > 0 )   " _
    & " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod " _
    & " INNER JOIN ColocPigno CP ON C.cCtaCod = CP.cCtaCod " _
    & " WHERE P.cCtaCod = '" & psCuenta & "'"

Set lrs = oConn.CargaRecordSet(sql)

Set dObtieneDatosCreditoPignoraticio = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

Public Function dObtieneCalendario(ByVal psCuenta As String) As Recordset
'Obtiene Datos de Calendarios por Conceptos
Dim lrs As ADODB.Recordset

On Error GoTo dError

sql = "SELECT  P.cCtaCod, " _
    & " PigCapital = ( SELECT nMonto FROM ColocCalendDetPig WHERE cCtaCod ='" & psCuenta & "' " _
    & "                 AND nPrdConceptoCod = " & gColPigConceptoCodCapital & "  and nColocCalendApl = 0  ) , " _
    & " Comisiones = ( SELECT nMonto FROM ColocCalendDetPig WHERE cCtaCod ='" & psCuenta & "' " _
    & "                 AND nPrdConceptoCod = " & gColPigConceptoCodTasacion & " and nColocCalendApl = 0  )  " _
    & " FROM  ColocCalendDetPig P " _
    & " WHERE P.cCtaCod = '" & psCuenta & "' and nColocCalendApl = 0"

Set lrs = oConn.CargaRecordSet(sql)

Set dObtieneCalendario = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
End Function

Public Function nObtieneDatosCostosRegistroCredPignoraticio(ByVal psCtaCod As String) As Recordset
Dim lrs As ADODB.Recordset

On Error GoTo dError

    sql = " SELECT P.cCtaCod, P.nPrdEstado, dVigencia, P.nTasaInteres, CP.nNroDuplic, " _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodInteresCompensatorio & "   then nMonto else 0 end )  as nInteres," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodImpuesto & "  then nMonto else 0 end )  as nImpuesto," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodTasacion & "  then nMonto else 0 end )  as nTasacion," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodCustodia & "  then nMonto else 0 end )  as nCustodia " _
         & " FROM Colocaciones c Inner join Producto P on P.cCtaCod = c.cCtaCod " _
         & " Inner Join ColocPignoraticio CP On C.cCtaCod = CP.cCtaCod Inner join ColocCalendario CCal on CCal.cCtaCod = C.cCtaCod " _
         & " Inner join ColoccalendDetPig CCalD on CCal.cCtaCod = CCalD.cCtaCod and CCal.nNroCalen = CCalD.nNroCalen " _
         & " And CCal.nColocCalendApl = CCalD.nColocCalendApl " _
         & " Where P.cCtaCod  = '" & psCtaCod & "' AND CCalD.nColocCalendApl =  " & gColocCalendAplCuota & " " _
         & " GROUP BY P.cCtaCod, P.nPrdEstado, dVigencia, P.nTasaInteres, CP.nNroDuplic "
    
Set lrs = oConn.CargaRecordSet(sql)

Set nObtieneDatosCostosRegistroCredPignoraticio = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
End Function

'**********************************************************************************************
' CAFF  - Modificado  - 04/09/2002 - CMCPL - Obtiene Datos de Credito Pignoraticio PERSONA
'**********************************************************************************************
Public Function dObtieneDatosCreditoPignoraticioPersonas(ByVal psCuenta As String) As Recordset

Dim lrs As ADODB.Recordset

On Error GoTo dError

sql = "SELECT PP.cPersCod, P.cPersNombre, P.cPersDireccDomicilio, P.cPersTelefono, " _
    & " P.cPersDireccUbiGeo, CE.cCalifiCliente, " _
    & " Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),'') , " _
    & " Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ) , " _
    & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) , " _
    & " NroDNI = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdDNI & " ),   " _
    & " NroRUC = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdRUC & " ),   " _
    & " DescCalif = (Select cConsDescripcion From Constante Where nConsCod = 8005 and nConsValor = CE.cCalifiCliente )   " _
    & " FROM ProductoPersona PP INNER JOIN Persona P  ON PP.cPersCod = P.cPersCod " _
    & "                                         LEFT JOIN ColocPigEvalCliente CE ON PP.cPersCod = CE.cPersCod " _
    & " WHERE cCtaCod = '" & psCuenta & "' "
    
Set lrs = oConn.CargaRecordSet(sql)

Set dObtieneDatosCreditoPignoraticioPersonas = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

'***************************************************
' CAFF -   Lista los Creditos de una Persona
'***************************************************
Public Function dObtieneCredPigDePersona(ByVal psPersCod As String, _
        ByVal psEstado As String, Optional ByVal psAgencia As String = "") As Recordset

Dim Rs As ADODB.Recordset
Dim lsAgencia As String

If Trim(psAgencia) = "" Then
    lsAgencia = "__"
Else
    lsAgencia = Trim(psAgencia)
End If

On Error GoTo dError

sql = "SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, " _
        & "T1.cConsDescripcion cRelacion, UPPER(T2.cConsDescripcion) cProducto, UPPER(T3.cConsDescripcion) cMoneda " _
        & "FROM ProductoPersona PP INNER JOIN Producto P ON P.cCtaCod = PP.cCtaCod INNER JOIN " _
        & "Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN " _
        & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN Constante T2 " _
        & "ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2.nConsValor) INNER JOIN " _
        & "Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) " _
        & "WHERE PP.cPersCod = '" & psPersCod & "' " _
        & "AND T1.nConsCod = " & gColocRelacPers & " AND T.nConsCod = " & gColocEstado & " AND " _
        & "T2.nConsCod = " & gProducto & " AND T3.nConsCod = " & gMoneda _
        & "AND P.cCtaCod like '%" & lsAgencia & "305" & "%' " _
        & "AND P.nPrdEstado in (" & psEstado & ") " & " ORDER BY PP.cCtaCod"

Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneCredPigDePersona = Rs
Set Rs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneCredPigDePersona>>", Err.Description
    
End Function

'************************************************************************
' CAFF  -  04/09/2002 -  Obtiene datos generales de un contrato
'************************************************************************
Public Function dObtieneDatosContrato(ByVal psCtaCod As String, ByVal pnTipoTasacion As Integer) As Recordset

Dim Rs As ADODB.Recordset

On Error GoTo dError

sql = "SELECT C.cCtaCod, nPrdEstado, nDiasAtraso, Estado = (SELECT cConsDescripcion FROM Constante WHERE nConsValor = P.nPrdEstado AND nConsCod = " _
    & gColocEstado & "), nPlazo, dVenc, nMontoCol, nTasaInteres, nNroDuplic, nMonto ComiTasacion, P.nSaldo, dVigencia, " _
    & "Count(J.cCtaCod) nPiezas, SUM(nPesoBruto) nPBruto, Sum(nPesoNeto) nPNeto, SUM(nTasacion + nTasacionAdicional) nTasacion, " _
    & "nTasacionAdicional, cPersCod, nTotalBalanza, nTasador = (SELECT cUser FROM RRHH WHERE cPersCod = J.cPersCod) " _
    & "FROM Colocaciones C INNER JOIN Producto P ON C.cCtaCod = P.cCtaCod INNER JOIN ColocPigno CP ON C.cCtaCod = CP.cCtaCod " _
    & "INNER JOIN ColocCalendDetPig CD ON  C.cCtaCod = CD.cCtaCod INNER JOIN ColocPigJoyaTasacion J ON C.cCtaCod = j.cCtaCod " _
    & "WHERE C.cCtaCod = '" & psCtaCod & "' AND nColocCalendApl = 0 AND nPrdConceptoCod = " & gColPigConceptoCodTasacion _
    & "AND J.nTipoTasacion = " & pnTipoTasacion _
    & " Group By C.cCtaCod, nPlazo, dVenc, nMontoCol, nTasaInteres, nNroDuplic, nMonto, cPersCod, nPrdEstado, nSaldo, dVigencia, " _
    & "nDiasAtraso, nTasacionAdicional, nTotalBalanza"

Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneDatosContrato = Rs
Set Rs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

'********************************************************************
' CRSF - 02/09/02 - OBTIENE LOS DATOS DE LOS CREDITOS PIGNORATICIOS
'********************************************************************
Public Function dObtieneCreditoPigno(ByVal psCuenta As String) As Recordset
'Obtiene Datos de Credito Pignoraticio Joyas
Dim lrs As ADODB.Recordset

On Error GoTo dError

sql = "SELECT P.cCtaCod, nPiezas, nPrdEstado, dPrdEstado, nDiasAtraso, nMontoCol, dVenc, nNroDuplic, " _
    & "Descri = (Select cConsDescripcion FROM Constante E WHERE E.nConsValor = P.nPrdEstado " _
    & "AND nConsCod = 3001) FROM PRODUCTO P INNER JOIN ColocPigno CP " _
    & "ON P.cCtaCod = CP.cCtaCod INNER JOIN Colocaciones C ON C.cCtaCod = CP.cCtaCod " _
    & "WHERE p.cCtaCod = '" & psCuenta & "' ORDER BY p.cCtaCod"

Set lrs = oConn.CargaRecordSet(sql)

Set dObtieneCreditoPigno = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function
'***************************************************************
' CRSF - 02/09/02 - OBTIENE DATOS DE EVALUACION DE CLIENTE
' DFAO - 14/07/03 - MODIFICADO
'***************************************************************

Public Function dEvaluacionCli(ByVal psPersona As String) As Recordset
'Obtiene Evaluacion del Cliente
Dim lrs As ADODB.Recordset

On Error GoTo dError

sql = "SELECT A.*, A.cCalifiCliente as Califi, B.cConsDescripcion as CMC, A.cEvalSbsCliente as EvalSbs, C.nConsValor, " _
           & " C.cConsDescripcion as SBS, D.cConsDescripcion as EVALANT " _
           & " from ColocPigEvalCliente A inner join Constante B on a.cCalifiCliente = B.nConsValor" _
           & " inner Join Constante C on a.cEvalSBSCliente = C.nConsValor " _
           & " Left Join constante D on a.cCalifiClienteAnt = D.nConsValor " _
           & " Where b.nConsCod = " & gColocPigCalifCte & " and A.cPersCod = '" & psPersona & "'" _
           & " and C.nConsCod = " & gColocPigCalifSbs & "  and  ( D.nConsCod  = " & gColocPigCalifCte & "or cCalifiClienteAnt is null)"
           
Set lrs = oConn.CargaRecordSet(sql)

Set dEvaluacionCli = lrs
Set lrs = Nothing

Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dEvaluacionCli>>", Err.Description
End Function

'****************************************************************
' CRSF - 02/09/02 - OBTIENE CUALQUIER VALOR DE LAS CONSTANTES
'****************************************************************
Public Function dEvaluacion(ByVal psCodigo As Integer) As Recordset
'Obtiene Evaluacion del Cliente
Dim lrs As ADODB.Recordset

On Error GoTo dError

sql = "SELECT  nConsCod, nConsValor, cConsDescripcion,  Convert(VarChar,nConsValor) + ' - ' + cConsDescripcion as Nom" _
           & " From Constante " _
           & " Where nConsCod = " & psCodigo & " and nConsValor <> " & psCodigo & " Order By nConsValor"
Set lrs = oConn.CargaRecordSet(sql)

Set dEvaluacion = lrs
Set lrs = Nothing

Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dEvaluacionCli>>", Err.Description
End Function
'*******************************************************************
' CAFF - 04/09/02 - OBTIENE EL DETALLE DE LAS JOYAS DE UN CONTRATO
'*******************************************************************
Public Function dObtieneDetalleJoyas(ByVal psCtaCod As String, Optional ByVal pnTipoTasacion = 0) As Recordset
Dim Rs As ADODB.Recordset

On Error GoTo dError

sql = "SELECT nItemPieza Item, C.cConsDescripcion Tipo, C.nConsValor nTipo, cUser, J.nMaterial, J.nConservacion, " _
    & "SubTipo = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = 8002 AND nConsValor = J.nSubTipoJoya), " _
    & "nSubTipo = (SELECT nConsValor FROM Constante WHERE nConsCod = 8002 AND nConsValor = J.nSubTipoJoya), " _
    & "Material = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigMaterial & " AND nConsValor = J.nMaterial), " _
    & "Estado = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigEstConservaJoya & " AND nConsValor = J.nConservacion), " _
    & "Situacion = (SELECT TOP 1 cConsDescripcion FROM Constante INNER JOIN ColocPigProceso CR ON nConsValor = CR.nSituacionPieza " _
    & "AND CR.cCtaCod = J.cCtaCod WHERE nConsCod = " & gColocPigSituacionPieza & " Order By nRemate desc), " _
    & "nPesoBruto PBruto, nPesoNeto PNeto, nTasacion Tasacion, nTasacionAdicional TasAdicion, cDescripcion Observacion, " _
    & "cDescripcionAdic as ObsAdicion, J.nTipoTasacion " _
    & "FROM ColocPigJoyaTasacion J INNER JOIN CONSTANTE C ON C.nConsValor = J.nTipoJoya " _
    & " INNER JOIN RRHH RH ON J.cPersCod = RH.cPersCod "
    
    If pnTipoTasacion = 0 Then
        sql = sql & "INNER JOIN COLOCPIGNO CP ON CP.cCtaCod = J.cCtaCod AND CP.nTipoTasacion = J.nTipoTasacion "
    End If

sql = sql & " WHERE C.nConsCod = " & gColocPigTipoJoya & " AND J.cCtaCod = '" & psCtaCod & "'"

If pnTipoTasacion <> 0 Then
    sql = sql & "AND J.nTipoTasacion = " & pnTipoTasacion
End If

sql = sql & " ORDER BY nItemPieza "
Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneDetalleJoyas = Rs
Set Rs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos de Joyas en <<dObtieneDetalleJoyas>>", Err.Description
 
End Function

'***********************************************
' CRSF - 02/09/02 - LOS VALORES DE LA COMISION
'***********************************************
Public Function dComision(ByVal psCtaCod As String, ByVal pnConst As String) As Recordset
'Obtiene Evaluacion del Cliente
Dim lrs As ADODB.Recordset

On Error GoTo dError

sql = "SELECT  cCtaCod, dPrdEstado,nPrdEstado,nMonto" _
           & " From  ColocacEstado " _
           & " Where cCtaCod = " & psCtaCod & " and  nPrdEstado in ( " & pnConst & ") "

Set lrs = oConn.CargaRecordSet(sql)

Set dComision = lrs
Set lrs = Nothing

Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dEvaluacion>>", Err.Description
End Function

Public Function dObtieneDatosPignoraticioDeuda(ByVal psCuenta As String) As Recordset
'Obtiene Datos de la Deuda de un Credito Pignoraticio
    
Dim lrs As ADODB.Recordset

On Error GoTo dError

sql = "SELECT sum(CASE WHEN nPrdConceptoCod = " & gColPigConceptoCodCapital & "  then nmonto else 0 end) capital, " _
         & "  sum(CASE WHEN nPrdConceptoCod = " & gColPigConceptoCodInteresComp & "  then nmonto else 0 end) intCompensatorio, " _
         & "  sum(CASE WHEN nPrdConceptoCod = " & gcolPigConceptoCodInteresMora & "  then nmonto else 0 end) IntMoratorio, " _
         & "  sum(CASE WHEN nPrdConceptoCod = " & gColPigConceptoCodComiServ & "  then nmonto else 0 end) ComServicio, " _
         & "  sum(CASE WHEN nPrdConceptoCod = " & gColPigConceptoCodComiVencida & "  then nmonto else 0 end) ComVencida, " _
         & "  sum(CASE WHEN nPrdConceptoCod = " & gColPigConceptoCodPreparaRemate & "  then nmonto else 0 end) DerRemate, " _
         & "  CapitalMinimo = (Select nParamValor from ColocParametro Where nParamVar = 8018), " _
         & "  PorcCapitalMinimo = (Select nParamValor from ColocParametro Where nParamVar = 8017) " _
         & "  FROM ColocPigno CP INNER JOIN ColocCalendDetPig CD ON CP.cCtacod = CD.cCtaCod " _
         & "  and CP.nNumCalend = CD.nNroCalen AND nColocCalendApl = " & gColocCalendAplCuota _
         & " WHERE CP.cCtaCod = '" & psCuenta & "'"

Set lrs = oConn.CargaRecordSet(sql)

Set dObtieneDatosPignoraticioDeuda = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

'**************************************************************************************************************************
 'CMCPL - Devuelve los datos de los contratos para la remesa a nivel de totales para la cabecera de la guia
 'CAFF - 18/09/2002     -  Solo es a nivel de Contratos
'**************************************************************************************************************************
Public Function dObtieneContratosColocPigGuia(ByVal pnOrigen As Integer, ByVal psEstado As String, _
        Optional ByVal psFecha As String, Optional ByVal pnDiasCustodia As Integer = -1, _
        Optional ByVal pnMotivo As Integer, Optional ByVal pnDestino As Integer) As Recordset

Dim Rs As Recordset

sql = "SELECT COUNT(P.cCtaCod) nItem, SUM(nSaldo) SaldoCap, SUM(nPiezas) nPiezas, SUM(nPBruto) nPBruto, SUM(nPNeto) nPNeto, " _
        & "SUM(nTasacion) nTasacion " _
        & " FROM " _
        & "     (SELECT C.cCtaCod, P.nSaldo, Count(J.cCtaCod) nPiezas, SUM(nPesoBruto) nPBruto, Sum(nPesoNeto) nPNeto, " _
                & "SUM(nTasacion + ISNULL(nTasacionAdicional,0)) nTasacion " _
                & "FROM Colocaciones C INNER JOIN Producto P ON C.cCtaCod = P.cCtaCod INNER JOIN ColocPigno CP ON " _
                & "C.cCtaCod = CP.cCtaCod INNER JOIN ColocPigJoyaTasacion J ON C.cCtaCod = J.cCtaCod "
   
If pnMotivo = 3 Then
    sql = sql & " INNER JOIN MOVCOL MC ON CP.CCTACOD = MC.CCTACOD INNER JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO "
End If
    
sql = sql & "WHERE J.nTipoTasacion = CP.nTipoTasacion AND nUbicaLote = " & pnOrigen & " And nPrdEstado in (" & psEstado & ") "

If pnMotivo = 3 Then
    sql = sql & "AND SUBSTRING(MC.COPECOD, 1,4) = '1510' AND convert(int, Substring(cMovNro, 18,2)) = " & pnDestino
End If

sql = sql & " AND C.cCtaCod NOT IN " _
          & "(SELECT cCtaCod FROM COLOCPIGGUIADET GD INNER JOIN COLOCPIGGUIAETAPA GE ON GE.cNumDoc = GD.cNumDoc WHERE nEtapa = 1) "
        
If pnDiasCustodia <> -1 Then
    sql = sql & "AND DATEDIFF(dd, dPrdEstado, '" & Format(psFecha, "mm/dd/yyyy") & "') > " & pnDiasCustodia
End If

sql = sql & " GROUP BY C.cCtaCod, nSaldo) P"

Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneContratosColocPigGuia = Rs
Set Rs = Nothing

End Function
'***********************************************************************************************************
 'CMCPL - Devuelve los datos de los contratos de la remesa para el detalle de la guia
 'CAFF - 18/09/2002     - Solo es a nivel de Contratos
'***********************************************************************************************************
Public Function dObtieneContratosColocPigGuiaDet(ByVal pnOrigen As Integer, ByVal psEstado As String, _
        Optional ByVal psFecha As String, Optional ByVal pnDiasCustodia As Integer = -1, _
        Optional ByVal pnMotivo As Integer, Optional ByVal pnDestino As Integer) As Recordset

Dim Rs As Recordset

sql = "SELECT C.cCtaCod, P.nSaldo, Count(J.cCtaCod) nPiezas, SUM(nPesoBruto) nPBruto, Sum(nPesoNeto) nPNeto, " _
        & "SUM(nTasacion + nTasacionAdicional) nTasacion " _
        & "FROM Colocaciones C INNER JOIN Producto P ON C.cCtaCod = P.cCtaCod INNER JOIN ColocPigno CP ON " _
        & "C.cCtaCod = CP.cCtaCod INNER JOIN ColocPigJoyaTasacion J ON C.cCtaCod = J.cCtaCod "
        
If pnMotivo = 3 Then
    sql = sql & " INNER JOIN MOVCOL MC ON CP.CCTACOD = MC.CCTACOD INNER JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO "
End If
        
sql = sql & " WHERE J.nTipoTasacion = CP.nTipoTasacion AND nUbicaLote = " & pnOrigen & " And nPrdEstado in (" & psEstado & ")"
        
If pnMotivo = 3 Then
    sql = sql & "AND SUBSTRING(MC.COPECOD, 1,4) = '1510' AND Convert(int, Substring(cMovNro, 18,2)) = " & pnDestino
End If
        
sql = sql & " AND C.cCtaCod NOT IN " _
          & "(SELECT cCtaCod FROM COLOCPIGGUIADET GD INNER JOIN COLOCPIGGUIAETAPA GE ON GE.cNumDoc = GD.cNumDoc WHERE nEtapa = 1) "
        
If pnDiasCustodia <> -1 Then
    sql = sql & "AND DATEDIFF(dd, dPrdEstado, '" & Format(psFecha, "mm/dd/yyyy") & "') > " & pnDiasCustodia
End If
        
sql = sql & "GROUP BY C.cCtaCod, nSaldo ORDER BY C.cCtaCod"

Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneContratosColocPigGuiaDet = Rs
Set Rs = Nothing

End Function
'**************************************************************************************************************************
 'CMCPL - Devuelve los datos de las Piezas para la remesa a nivel de totales para la cabecera de la guia
 'CAFF - 18/09/2002     -  Solo es a nivel de Piezas
'**************************************************************************************************************************
Public Function dObtienePiezasColocPigGuia(ByVal pnOrigen As Integer, ByVal psEstado As String, _
                Optional ByVal psFecha As String, Optional ByVal pnMotivo As Integer) As Recordset

Dim Rs As Recordset

sql = "SELECT Count(P.cCtaCod) nItem, SUM(nSaldo) SaldoCap, SUM(nPesoBruto) nPBruto, SUM(nPesoNeto) nPNeto, " _
        & "SUM(nTasacion) nTasacion " _
        & "FROM " _
        & "(SELECT C.cCtaCod, J.nItemPieza, P.nSaldo, nPesoBruto, nPesoNeto, nTasacion + nTasacionAdicional nTasacion " _
        & "FROM Colocaciones C INNER JOIN Producto P ON C.cCtaCod = P.cCtaCod INNER JOIN ColocPigno CP ON C.cCtaCod = CP.cCtaCod " _
        & "INNER JOIN ColocPigJoyaTasacion J ON CP.cCtaCod = j.cCtaCod AND J.nTipoTasacion = CP.nTipoTasacion " _
        & "WHERE nUbicaLote = " & pnOrigen & " AND nPrdEstado in (" & psEstado & ") " _
        & "AND C.cCtaCod + CONVERT(varchar(2),nItemPieza) NOT IN " _
        & "(SELECT cCtaCod + CONVERT(varchar(2),nItemPieza) FROM COLOCPIGGUIADET GD INNER JOIN COLOCPIGGUIAETAPA GE " _
        & "ON GD.cNumDoc = GE.cNumDoc INNER JOIN COLOCPIGGUIA G ON G.cNumDoc = GD.cNumDoc " _
        & " WHERE nEtapa = 1 AND nMotivo = " & pnMotivo & " )) P "

Set Rs = oConn.CargaRecordSet(sql)

Set dObtienePiezasColocPigGuia = Rs
Set Rs = Nothing

End Function

'***********************************************************************************************************
 'CMCPL - Devuelve los datos de las piezas de la remesa para el detalle de la guia
 'CAFF - 18/09/2002     - Solo es nivel de Piezas
'***********************************************************************************************************
Public Function dObtienePiezasColocPigGuiaDet(ByVal pnOrigen As Integer, ByVal psEstado As String, Optional ByVal psFecha As String, _
        Optional ByVal pnAgencia As Integer = -1) As Recordset

Dim Rs As Recordset

sql = "SELECT C.cCtaCod, J.nItemPieza, P.nSaldo, nPesoBruto, nPesoNeto, nTasacion + nTasacionAdicional nTasacion " _
        & "FROM Colocaciones C INNER JOIN Producto P ON C.cCtaCod = P.cCtaCod INNER JOIN ColocPigno CP ON C.cCtaCod = CP.cCtaCod " _
        & "INNER JOIN ColocPigJoyaTasacion J ON CP.cCtaCod = J.cCtaCod AND J.nTipoTasacion = CP.nTipoTasacion " _
        & "WHERE nUbicaLote = " & pnOrigen & " AND nPrdEstado in (" & psEstado & ") "
        
Set Rs = oConn.CargaRecordSet(sql)

Set dObtienePiezasColocPigGuiaDet = Rs
Set Rs = Nothing

End Function

'******************************************************************************
 'CMCPL - Selecciona las Guias dependiendo de la Etapa (Despacho, Confirmacion)
 'CAFF - 21/09/2002
'******************************************************************************
Public Function dObtieneGuias(ByVal pnEtapa As Integer, Optional ByVal pnOrigen As Integer = -1, _
            Optional ByVal pnDestino As Integer = -1, Optional ByVal pnMotivo As Integer = -1) As Recordset

Dim Rs As Recordset

sql = "SELECT cNumDoc, Motivo = (SELECT cConsDescripcion FROM CONSTANTE WHERE nConsCod = 8015 AND nConsValor = G.nMotivo), " _
        & "Origen = (SELECT cConsDescripcion FROM CONSTANTE WHERE nConsCod = 8004 AND nConsValor = G.nOrigen), " _
        & "Destino = (SELECT cConsDescripcion FROM CONSTANTE WHERE nConsCod = 8004 AND nConsValor = G.nDestino), " _
        & "nTotItem, TipoGuia = (SELECT cConsDescripcion FROM CONSTANTE WHERE nConsCod = 8010 AND nConsValor = G.nTipoGuia), " _
        & "nMotivo, nPesoNeto, nPesoBruto, nTotTasacion " _
        & "FROM COLOCPIGGUIA G WHERE cFlag = " & pnEtapa _

If pnOrigen <> -1 Then
    sql = sql & "AND nOrigen = " & pnOrigen
End If

If pnDestino <> -1 Then
    sql = sql & "AND nDestino = " & pnDestino
End If

If pnMotivo <> -1 Then
    sql = sql & "AND nMotivo = " & pnMotivo
End If
        
sql = sql & " ORDER BY cNumDoc"
        
Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneGuias = Rs
Set Rs = Nothing

End Function

'**************************************************************************************************************************
 'CMCPL - Devuelve los datos de las Piezas para la remesa y los Inserta en una tabla Temporal
 'CAFF - 23/09/2002     -  Solo es a nivel de Piezas
'**************************************************************************************************************************
Public Sub dInsertaPiezasAdjColocPigGuia(ByVal pnOrigen As Integer, ByVal pnDestino As Integer, ByVal psEstado As String, _
        Optional ByVal pnAgencia As Integer = -1, Optional ByVal pnMaterial As Integer = -1, _
        Optional ByVal pnRemate As Integer = -1, Optional ByVal psPieza As String = "@")


On Error GoTo dError

sql = "Delete TMPCOLOCPIGPIEZAS"
    oConn.Ejecutar sql

sql = "INSERT TMPCOLOCPIGPIEZAS(nOrigen, nDestino, cCodCta, nItemPieza, nPesoBruto, nPesoNeto, nTotTasacion, nValorAdjud) " _
        & "SELECT Origen = " & pnOrigen & ", Destino = " & pnDestino & ", C.cCtaCod, J.nItemPieza, nPesoBruto, nPesoNeto, " _
        & "nTasacion + nTasacionAdicional nTasacion, nValorProceso " _
        & "FROM Colocaciones C INNER JOIN Producto P ON C.cCtaCod = P.cCtaCod INNER JOIN ColocPigno CP ON C.cCtaCod = CP.cCtaCod " _
        & "INNER JOIN ColocPigJoyaTasacion J ON C.cCtaCod = j.cCtaCod INNER JOIN ColocPigProceso R ON J.cCtaCod = R.cCtaCod AND R.nItemPieza = J.nItemPieza " _
        & "WHERE J.nTipoTasacion = CP.nTipoTasacion AND nUbicaPieza = " & pnOrigen & " AND nPrdEstado in (" & psEstado & ") " _
        & "AND nTipoProceso IN (4,5,6) "

        If pnAgencia <> -1 Then
            sql = sql & " AND SUBSTRING(P.cCtaCod, 4,2) = " & pnAgencia
        End If

        If pnMaterial <> -1 Then
            sql = sql & "AND nMaterial = " & pnMaterial
        End If

        If pnRemate <> -1 Then
            sql = sql & "AND nRemate = " & pnRemate
        End If

        If psPieza <> "@" Then
            sql = sql & "AND P.cCtaCod = '" & Mid(psPieza, 1, 18) & "' AND J.nItemPieza = " & Mid(psPieza, 19, Len(psPieza) - 18)
        End If

    oConn.Ejecutar sql

    Exit Sub
    
dError:
    If vbObjectError = -2147221504 Then
        MsgBox "Pieza ya fue seleccionada", vbInformation, "Aviso"
    Else
        Err.Raise Err.Number, "Inserta datos Temporales en TmpColocPigPiezas <<dObtienePiezasAdjColocPigGuia>>", Err.Description
    End If

End Sub

'**************************************************************************************************************************
 'CMCPL - Devuelve los datos de las Piezas para la remesa a nivel de totales para la cabecera de la guia
 'CAFF - 23/09/2002     -  Solo es a nivel de Piezas
'**************************************************************************************************************************
Public Function dObtienePiezasAdjColocPigGuia(ByVal pnOrigen As Integer, ByVal pnDestino As Integer) As Recordset

Dim Rs As Recordset

sql = "SELECT COUNT(cCodCta) nTotItem, SUM(nPesoBruto) nPBruto, SUM(nPesoNeto) nPNeto, SUM(nTotTasacion) nTasacion, " _
        & "SUM(nValorAdjud) nValorAdjud FROM TMPCOLOCPIGPIEZAS WHERE nOrigen = " & pnOrigen & " AND nDestino = " & pnDestino

Set Rs = oConn.CargaRecordSet(sql)

Set dObtienePiezasAdjColocPigGuia = Rs
Set Rs = Nothing

End Function
'**********************************************************************************************************
 'CMCPL - Devuelve los datos de las Piezas para la remesa de la Tabla Temporal
 'CAFF - 24/09/2002     -  Solo es a nivel de Piezas
'**********************************************************************************************************
Public Function dObtienePiezasAdjColocPigGuiaDet(ByVal pnOrigen As Integer, ByVal pnDestino As Integer) As Recordset

Dim Rs As Recordset

sql = "SELECT cCodCta, nItemPieza FROM TMPCOLOCPIGPIEZAS WHERE nOrigen = " & pnOrigen & " AND nDestino = " & pnDestino

Set Rs = oConn.CargaRecordSet(sql)

Set dObtienePiezasAdjColocPigGuiaDet = Rs
Set Rs = Nothing

End Function
'****************************************************************************************
'CAFF - 25/09/2002 - PARA EL DETALLE DE LAS GUIAS EN CONFIRMACION DE CONTENIDO POR LOTES
'****************************************************************************************
Public Function dObtieneColocPigGuiaDetLotes(ByVal psNumDoc As String) As Recordset
Dim Rs As Recordset

sql = "SELECT GD.cCtaCod, GD.nItemPieza, SUM(nPesoNeto) PNeto, COUNT(CP.cCtaCod) nPiezas " _
    & "FROM ColocPigGuiaDet GD " _
            & "INNER JOIN COLOCPIGNO C ON GD.cCtaCod = C.cCtaCod " _
            & "INNER JOIN COLOCPIGJOYATASACION CP ON CP.cCtaCod = GD.cCtaCod " _
            & "WHERE CP.nTipoTasacion = C.nTipoTasacion AND cNumDoc = '" & psNumDoc & "'" _
            & "GROUP BY GD.cCtaCod, GD.nItemPieza "

Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneColocPigGuiaDetLotes = Rs
Set Rs = Nothing

End Function
'*****************************************************************************************
'CAFF - 25/09/2002 - PARA EL DETALLE DE LAS GUIAS EN CONFIRMACION EN CONTENIDO POR PIEZAS
'*****************************************************************************************
Public Function dObtieneColocPigGuiaDetPiezas(ByVal psNumDoc As String) As Recordset
Dim Rs As Recordset

sql = "SELECT GD.cCtaCod, GD.nItemPieza, nPesoNeto, cDescripcion " _
    & "FROM ColocPigGuiaDet GD " _
            & "INNER JOIN COLOCPIGNO C ON GD.cCtaCod = C.cCtaCod " _
            & "INNER JOIN COLOCPIGJOYATASACION CP ON (CP.cCtaCod = GD.cCtaCod AND " _
            & "GD.nItemPieza = CP.nItemPieza) " _
            & "WHERE CP.nTipoTasacion = C.nTipoTasacion AND cNumDoc = '" & psNumDoc & "'"
            
Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneColocPigGuiaDetPiezas = Rs
Set Rs = Nothing

End Function
'****************************************************************
'CAFF - 26/09/2002 - DEVUELVE LOS DATOS DE UNA PIEZA ADJUDICADA
'****************************************************************
Public Function dObtieneDatosPieza(ByVal psCodCta As String, pnItemPieza As Integer) As Recordset
Dim Rs As Recordset

sql = "SELECT nRemate, C.cCtaCod, J.nItemPieza, cConsDescripcion, nPesoBruto, nPesoNeto, cDescripcion, nValorProceso " _
    & "FROM Colocaciones C INNER JOIN ColocPigno CP ON C.cCtaCod = CP.cCtaCod " _
                        & "INNER JOIN ColocPigJoyaTasacion J ON C.cCtaCod = J.cCtaCod " _
                        & "INNER JOIN ColocPigProceso R ON CP.cCtaCod = R.cCtaCod " _
                        & "INNER JOIN Constante CO ON CO.nConsValor = J.nMaterial " _
    & "WHERE J.nTipoTasacion = CP.nTipoTasacion AND nTipoProceso IN (4,5,6) AND nConsCod = 8003 " _
           & "AND C.cCtaCod = '" & psCodCta & "' AND J.nItemPieza = " & pnItemPieza
    
Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneDatosPieza = Rs
Set Rs = Nothing

End Function

'*********************************************************************
' EAFA - 02/09/02 OBTIENE PIEZAS PARA EL CALCULO DE LA LINEA
'*********************************************************************
Public Function dObtieneDatosPignoraticioJoyas(ByVal psCuenta As String) As Recordset
'Obtiene Datos de las Joyas para realizar el recalculo del Credito para el Uso de Linea
    
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = "SELECT CP.cCtaCod, CPJT.nItemPieza, CPJT.nMaterial, CPJT.nConservacion, " _
         & "            CPJT.nPesoBruto, CPJT.nPesoNeto, CPJT.nTasacion, CPJT.nTasacionAdicional, " _
         & "            CPJT.cDescripcion, CPJT.cDescripcionAdic, CPJT.nTipoJoya, CPJT.nSubTipoJoya, " _
         & "            CPJT.nUbicaPieza, CPJT.cPersCod, CP.nUsoLineaNro, CP.nTipoTasacion " _
         & "  FROM ColocPigno CP INNER JOIN ColocPigJoyaTasacion CPJT ON CP.cCtacod = CPJT.cCtaCod " _
         & "   and CP.nTipoTasacion = CPJT.nTipoTasacion " _
         & " WHERE CP.cCtaCod = '" & psCuenta & "'"

Set lrs = oConn.CargaRecordSet(lsSQL)

Set dObtieneDatosPignoraticioJoyas = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

'********************************************************************************
'CRSF -  23/09/2002 - OBTIENE MAXIMO NUMERO DE REMATE MAS 1 - COLOCPIGREMATE
'********************************************************************************
Public Function dObtieneMaxRemate() As Long
Dim lrs As ADODB.Recordset

On Error GoTo dError

sql = "SELECT  max(nRemate) + 1 as nremate" _
         & "  FROM ColocPigRemate"
         
Set lrs = oConn.CargaRecordSet(sql)
         
   dObtieneMaxRemate = IIf(IsNull(lrs!nRemate), 1, lrs!nRemate)
   
 Set lrs = Nothing
Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Remate en <<dObtieneDatosRemate>>", Err.Description
End Function

'**************************************************************************************
' 23/09/2002 - SELECCION DE CONTRATOS PARA EL REMATE - RETASACION - COLOCPIGNO
'**************************************************************************************
Public Function dObtieneSeleccionContratos(ByVal pnDiasVenc As Integer, ByVal pdFecIniRem As Date) As Recordset
Dim Rs As Recordset

    On Error GoTo dError
    
     sql = "SELECT Distinct CP.cCtaCod Contrato, nPiezas Pieza, " _
        & "DiasAtraso = Datediff(dd, dVenc, '" & Format(pdFecIniRem, "mm/dd/yyyy") & "') " & _
        " FROM COLOCPIGJOYATASACION J INNER JOIN COLOCPIGNO CP ON J.cCtaCod = CP.cCtaCod " & _
        " INNER JOIN COLOCACIONES C ON CP.cCtaCod = C.cCtaCod " & _
        " WHERE J.nTipoTasacion = CP.nTipoTasacion " & _
        " AND J.cCtaCod In (" _
    & " SELECT CP.cCtaCod " _
    & "FROM COLOCPIGNO CP INNER JOIN COLOCACIONES C ON CP.CCTACOD = C.CCTACOD " _
    & "INNER JOIN PRODUCTO P ON P.CCTACOD = CP.CCTACOD " _
    & "WHERE DATEDIFF(dd, dVenc, '" & Format(pdFecIniRem, "mm/dd/yyyy") & "') >= " & pnDiasVenc _
    & " AND nPrdEstado IN (" & gPigEstAmortiz & "," & gPigEstDesemb & "," & gPigEstReusoLin & ") AND nTipoTasacion <> 3)"

    Set Rs = oConn.CargaRecordSet(sql)
    Set dObtieneSeleccionContratos = Rs
    Set Rs = Nothing
    Exit Function
    
dError:
    Err.Raise Err.Number, "Obtiene Datos Remate en <<dObtieneDatosRemate>>", Err.Description
End Function

'**********************************************************************************
' CRSF -  30/09/2002 CONTRATOS PARA LA RETASACION MANUAL, FALTA REMATE - COLOCPIGNO
'**********************************************************************************
Public Function dObtieneContratosRetas(ByVal pnTasacion As Integer, ByVal pnDiasVenc As Integer, ByVal pdFecIniRem As Date) As Recordset
'Obtiene los Datos de los contratos para la selecion de remates

    Dim lrs As ADODB.Recordset
    On Error GoTo dError
    sql = "SELECT CP.cCtaCod, nPiezas, P.nPrdEstado,  " _
             & " nEstadoCont, nUbicaLote, nTipoTasacion, nDiasAtraso, " _
             & "(Select b.cConsDescripcion From constante b  where b.nconscod = '8004' and b.nconsvalor = CP.nUbicalote ) as ubica, " _
             & "(Select b.cConsDescripcion From constante b  where b.nconscod = '3001' and b.nconsvalor = P.nPrdEstado) as estado " _
             & " FROM ColocPigno CP INNER JOIN Colocaciones C On CP.cCtaCod = C.cCtaCod INNER JOIN PRODUCTO P ON CP.CCTACOD = P.CCTACOD " _
             & " WHERE nTipoTasacion =  " & pnTasacion & " AND DATEDIFF(dd, dVenc, '" & Format(pdFecIniRem, "mm/dd/yyyy") _
             & "') >= " & pnDiasVenc & " Order By CP.cCtaCod "

    Set lrs = oConn.CargaRecordSet(sql)
    Set dObtieneContratosRetas = lrs
    Set lrs = Nothing
     Exit Function
dError:
        Err.Raise Err.Number, "Obtiene Datos Contratos Reatsacion <<dObtieneContratosRetas>>", Err.Description
End Function

'*******************************************************************
' CRSF -  23/09/2002 - OBTIENE DATOS DEL REMATE - COLOCPIGREMATE
'*******************************************************************
Public Function dObtieneDatosRemate(ByVal pnRemate As Integer) As Recordset
'Obtiene los Datos del Remate
Dim lrs As ADODB.Recordset

On Error GoTo dError
   sql = "SELECT  nRemate, nTipoProceso, cUbicacion, dInicio,dFin, dReferencia, nAdjudicado, nSobrante,  " _
       & " R.cPerscod, cPersNombre, dAdjudica, dAviso, dInicio, dfin, " _
       & " cConsDescripcion = (SELECT cConsDescripcion FROM CONSTANTE C WHERE nConsCod = 8004 AND nConsValor = R.cUbicacion)" _
       & " FROM ColocPigRemate R LEFT JOIN PERSONA P ON R.cPersCod = P.cPersCod " _
       & " WHERE nRemate = " & pnRemate

Set lrs = oConn.CargaRecordSet(sql)
Set dObtieneDatosRemate = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Remate en <<dObtieneDatosRemate>>", Err.Description
    
End Function

Public Function dSeleccionaSobrantes(ByVal pnRemate As Long) As Recordset

On Error GoTo dError

    sql = "SELECT nRemate, PR.cCtaCod, P.cPersCod, SUM(ISNULL(nValorVenta,0) - (nValorDeuda)) nSobrante " _
        & "FROM COLOCPIGPROCESO PR INNER JOIN PRODUCTOPERSONA PP ON PR.CCTACOD = PP.CCTACOD " _
        & "                        INNER JOIN PERSONA P ON PP.CPERSCOD = P.CPERSCOD " _
        & "WHERE (nValorVenta > 0 Or nValorDeudaComp < 0) And nRemate = " & pnRemate _
        & " GROUP BY nRemate, PR.cCtaCod, P.cPersCod " _
        & " HAVING SUM(ISNULL(nValorVenta, 0) - (nValorDeuda)) > 0"

    Set dSeleccionaSobrantes = oConn.CargaRecordSet(sql)

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Contratos con Sobrante de Dinero <<dSeleccionaSobrantes>>", Err.Description

End Function

'*****************************************************************
'31/01/2003 VALORES DEL CONTRATO (LOTES) PARA EL REMATE
'17/06/2003  - MODIFICADO CRSF
'*****************************************************************
Public Function dObtieneContratosLotes(ByVal psCtaCod As String, ByVal pnTipoProceso As Integer) As Recordset

Dim lrs As ADODB.Recordset
    
    On Error GoTo dError

    sql = "SELECT CP.cCtaCod, nPrdEstado, CP.nPiezas, CP.nTipoTasacion, CP.nUbicaLote, SUM(nPesoNeto) PesoNeto, SUM(nTasacion) Retasacion, " _
        & "SUM(nValorDeuda) ValorDeuda, SUM(nValorProceso) ValorProceso " _
        & "FROM COLOCPIGNO CP INNER JOIN PRODUCTO P ON CP.cCtaCod = P.cCtaCod " _
        & "INNER JOIN COLOCPIGJOYATASACION CT ON CP.cCtaCod = CT.cCtaCod " _
        & "INNER JOIN COLOCPIGPROCESO CR ON CR.cCtaCod = CT.cCtaCod AND CT.nItemPieza = CR.nItemPieza " _
        & "WHERE CT.nTipoTasacion = 3 AND nTipoProceso = " & pnTipoProceso & " AND CP.cCtaCod = '" & psCtaCod & "' " _
        & " GROUP BY CP.cCtaCod, CP.nPiezas, CP.nTipoTasacion, nPrdEstado, CP.nUbicaLote"
     '& "AND nSituacionPieza = " & gPigSituacionDisponible
    Set lrs = oConn.CargaRecordSet(sql)
    Set dObtieneContratosLotes = lrs
    Set lrs = Nothing
    
    Exit Function
dError:
        Err.Raise Err.Number, "Obtiene Datos Contratos Lotes <<dObtieneContratosLotes>>", Err.Description
End Function

'*******************************************************************
' CRSF -  14/10/2002 VALORES DEL CONTRATO (PIEZAS) PARA EL REMATE
' CRSF - MODIFICADO CRSF 17/06/2003
'*******************************************************************
Public Function dObtieneContratosPiezas(ByVal psCtaCod As String, ByVal pnTipoProceso As Integer) As Recordset
    
Dim lrs As ADODB.Recordset
    
    On Error GoTo dError
      
    sql = "SELECT CP.cCtaCod, CP.nItemPieza, CP.nMaterial, CP.nTipoJoya, CP.cDescripcion, CP.nPesoNeto,  CR.nValorDeuda, " _
      & "CR.nValorProceso, C.cConsDescripcion, D.cConsDescripcion DescriTipo, CR.nValorVenta, CR.cComprador " _
      & "FROM COLOCPIGJOYATASACION CP INNER JOIN COLOCPIGPROCESO CR ON CP.cCtaCod = CR.cCtaCod " _
      & "AND CP.nItemPieza = CR.nItemPieza " _
      & "INNER JOIN constante C ON c.nConsValor = CP.nmaterial AND c.nconscod = '8003' " _
      & "INNER JOIN constante D ON d.nConsValor = CP.nTipoJoya AND d.nconscod = '8001' " _
      & "WHERE  CP.nTipoTasacion = 3 AND CP.cCtaCod =  '" & psCtaCod & "'" _
      & "AND nTipoProceso = " & pnTipoProceso
       ' AND CR.nSituacionPieza = 1
    Set lrs = oConn.CargaRecordSet(sql)
    Set dObtieneContratosPiezas = lrs
    Set lrs = Nothing
     
     Exit Function

dError:
        Err.Raise Err.Number, "Obtiene Datos Contratos Piezas <<dObtieneContratosPiezas>>", Err.Description
End Function
'******************************************************************************
' CRSF -  17/10/2002 LISTADO DE CLIENTES CON PIEZAS O CONTRATOS POR FACTURAR
'******************************************************************************
Public Function dObtieneClienteRematado(ByVal pnPieza As ColocPigSituacionPieza) As Recordset

Dim lrs As ADODB.Recordset
On Error GoTo dError
    
sql = "SELECT distinct cComprador, count(nitempieza) Piezas, sum(nValorVenta) ValorVenta " _
      & " FROM  ColocPigProceso " _
      & " WHERE  nSituacionPieza = " & pnPieza & " and cComprador is Not Null " _
      & " GROUP BY cComprador"
      
    Set lrs = oConn.CargaRecordSet(sql)
    Set dObtieneClienteRematado = lrs
    Set lrs = Nothing
     
     Exit Function
dError:
        Err.Raise Err.Number, "Obtiene Datos Cliente Rematados <<dObtieneClienteRematado>>", Err.Description

End Function
'******************************************************************************
' CRSF -  17/10/2002LISTADO DE PIEZAS REMATADAS POR CLIENTES(NO FACTURADAS)
'******************************************************************************
Public Function dObtieneListadoPiezaRem(ByVal psCliente As String, ByVal pnPiezas As ColocPigSituacionPieza, _
            ByVal pnTasacion As ColocPigTipoTasacion, Optional ByVal pnRemate As Integer) As Recordset
    
    On Error GoTo dError
    
    sql = "SELECT A.cCtaCod, nRemate, nTipoProceso, nPrdEstado, A.nItemPieza Pieza, C.cConsDescripcion Material, D.cConsDescripcion Tipo," _
        & "A.cDescripcion, A.nPesoNeto, B.nValorProceso ValorRemate, b.nValorVenta ValorVenta, " _
        & "nValorTasacion = (SELECT (nTasacion + nTasacionAdicional) Tasacion FROM ColocPigJoyaTasacion J " _
        & "WHERE J.cCtaCod = A.cCtaCod AND J.nItemPieza = A.nItemPieza AND J.nTipoTasacion = 1) " _
        & "FROM ColocPigJoyaTasacion A INNER JOIN ColocPigno CP ON A.cCtaCod = CP.cCtaCod AND A.nTipoTasacion = CP.nTipoTasacion " _
        & "INNER JOIN ColocPigProceso B ON b.cCtaCod = a.cCtaCod AND B.nItemPieza = A.nItemPieza " _
        & "INNER JOIN PRODUCTO P ON A.cCtaCod = P.cCtaCod " _
        & "INNER JOIN constante c ON c.nConsValor = A.nMaterial and c.nConsCod = '8003'" _
        & "INNER JOIN constante d ON d.nConsValor = A.nTipoJoya and d.nConsCod = '8001'" _
        & "WHERE  a.ntipotasacion = " & pnTasacion & " and B.nSituacionPieza = " & pnPiezas & _
        " and b.cComprador = '" & psCliente & "' AND nRemate = " & pnRemate _
        & "ORDER BY P.CCTACOD "
      
    Set dObtieneListadoPiezaRem = oConn.CargaRecordSet(sql)

Exit Function

dError:
        Err.Raise Err.Number, "Obtiene Datos Listado Piezas Rematados <<dObtieneListadoPiezaRem>>", Err.Description
End Function

'******************************************************************************************
' CRSF -  29/10/2002 LISTADO DE PIEZAS POR CONCEPTOS DE LA TABLA COLOCPIGPROCESODET
'******************************************************************************************
Public Function dObtieneConceptoPieza(ByVal psCuenta As String, ByVal pnPieza As Integer, _
                                     ByVal pnRemate As Long) As Recordset
                                     
Dim lrs As ADODB.Recordset
    
    On Error GoTo dError
    sql = "SELECT cCtaCod,nItemPieza,nCodConcepto,nMonto " _
      & " FROM ColocpigProcesoDet  " _
      & " WHERE  cCtaCod = " & psCuenta & " and nRemate = " & pnRemate & " and nItemPieza = " & pnPieza & ""
      
    Set lrs = oConn.CargaRecordSet(sql)
    Set dObtieneConceptoPieza = lrs
    Set lrs = Nothing

    Exit Function

dError:
        Err.Raise Err.Number, "Obtiene Concepto Pieza <<dObtieneConceptoPieza>>", Err.Description
        
End Function



'**********************************************************************
'CMCPL OBTIENE CONTRATOS PARA CONSULTA
'EAFA - 17/09/02
'**********************************************************************
Public Function dBuscaContratosConsultaCredPignoraticio(ByVal psCliente As String, ByVal psEstados As String, _
                                    Optional ByVal psCtaCod As String) As Recordset

Dim lrDatos As ADODB.Recordset
Dim lsSQL As String

    lsSQL = " SELECT PP.cCtaCod, P.nPrdestado, " _
              & " DescAge = (select cAgeDescripcion from Agencias Where cAgeCod = substring(PP.cCtaCod, 4, 2)) " _
              & " FROM ProductoPersona PP INNER JOIN Producto P ON P.cCtaCod = PP.cCtaCod " _
              & "                                         INNER JOIN ColocPignoraticio CP on CP.cctacod = P.cctacod " _
              & " WHERE PP.cPersCod ='" & psCliente & "' "
              
    If psEstados <> "" Then
         lsSQL = lsSQL & " AND P.nPrdEstado in (" & psEstados & ")" _

    End If
    If psCtaCod <> "" Then
        lsSQL = lsSQL & " AND PP.cCtaCod = '" & psCtaCod & "'"
    End If
    
    Set lrDatos = oConn.CargaRecordSet(lsSQL)
    If (lrDatos.BOF And lrDatos.EOF) Or lrDatos Is Nothing Then
        MsgBox "No se encontro Contratos para el Cliente ", vbInformation, "Aviso"
        Set lrDatos = Nothing
        Exit Function
    End If
    
    
    Set dBuscaContratosConsultaCredPignoraticio = lrDatos
    Set lrDatos = Nothing
    
End Function

'**********************************************************************
'CMCPL OBTIENE DATOS DEL CLIENTE
'EAFA - 17/09/02
'**********************************************************************
Public Function dObtieneDatosClienteCredPignoraticio(ByVal psCliente As String) As Recordset

Dim lrDatos As ADODB.Recordset
Dim lsSQL As String

lsSQL = " SELECT P.cPersCod, P.cPersNombre, P.cPersTelefono, P.cPersDireccDomicilio, " _
          & "             CE.cCalifiCliente, CE.cEvalSBSCliente, nPersPersoneria, " _
          & " NroDNI = (Select ISNULL(cPersIDnro,'') From PersID Where cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdDNI & " ), " _
          & " NroRUC = (Select ISNULL(cPersIDnro,'') From PersID Where cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdRUC & " ), " _
          & " DescCalif = (Select cConsDescripcion From Constante Where nConsCod = 8005 and nConsValor = CE.cCalifiCliente ) ,  " _
          & " DescSBS = (Select cConsDescripcion From Constante Where nConsCod = 3010 and nConsValor = CE.cEvalSBSCliente )   " _
          & " FROM Persona P  LEFT JOIN ColocPigEvalCliente CE ON P.cPersCod = CE.cPersCod " _
          & " WHERE P.cPersCod = '" & psCliente & "'"
    
    Set lrDatos = oConn.CargaRecordSet(lsSQL)
            
    Set dObtieneDatosClienteCredPignoraticio = lrDatos
    Set lrDatos = Nothing
    
End Function

'**********************************************************************
'CMCPL OBTIENE DATOS DEL CREDITO
'EAFA - 17/09/02
'**********************************************************************
Public Function dObtieneDetalleCredPignoraticio(ByVal psCtaCod As String, ByVal psEstados As String) As Recordset

Dim lrDatos As ADODB.Recordset
Dim lsSQL As String

lsSQL = "SELECT C.cCtaCod, nPrdEstado, dPrdEstado, nDiasAtraso, nPlazo, dVenc, nMontoCol, " _
          & "             nTasaInteres, nNroDuplic, nUsoLineaNro, P.nSaldo, dVigencia, " _
          & " nTasaMora = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
          & "                     WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ) , " _
          & " nInteresComp = (Select nMonto - nMontoPagado From ColocCalendDetPig where cCtaCod='" & psCtaCod & "' and nNroCalen=CP.nNumCalend and " _
          & "                         nColocCalendApl='" & gColocCalendAplCuota & "' and nCuota=1 and nPrdConceptoCod='" & gColPigConceptoCodInteresComp & "')," _
          & " nInteresMora = (Select nMonto - nMontoPagado From ColocCalendDetPig where cCtaCod='" & psCtaCod & "' and nNroCalen=CP.nNumCalend and " _
          & "                         nColocCalendApl='" & gColocCalendAplCuota & "' and nCuota=1 and nPrdConceptoCod='" & gcolPigConceptoCodInteresMora & "')," _
          & " nComisionVenc = (Select nMonto- nMontoPagado  From ColocCalendDetPig where cCtaCod='" & psCtaCod & "' and nNroCalen=CP.nNumCalend and " _
          & "                         nColocCalendApl='" & gColocCalendAplCuota & "' and nCuota=1 and nPrdConceptoCod='" & gColPigConceptoCodComiVencida & "')," _
          & " nDerechoRemate = (Select nMonto - nMontoPagado From ColocCalendDetPig where cCtaCod='" & psCtaCod & "' and nNroCalen=CP.nNumCalend and " _
          & "                   nColocCalendApl='" & gColocCalendAplCuota & "' and nCuota=1 and nPrdConceptoCod='" & gColPigConceptoCodPreparaRemate & "')," _
          & " CapitalMinimo = (Select nParamValor from ColocParametro Where nParamVar = 8018), " _
          & "  PorcCapitalMinimo = (Select nParamValor from ColocParametro Where nParamVar = 8017), " _
          & " nPenalidad = (Select nMonto - nMontoPagado FROM ColocCalendDetPig WHERE cCtaCod = '" & psCtaCod & "' AND nNroCalen = CP.nNumCalend AND " _
          & "               nColocCalendApl = '" & gColoCalendAplCuota & "' AND nCuota = 1 AND nPrdConceptoCod = '" & gColPigConceptoCodPenalidad & "')," _
          & " nComisionServicio = (Select nMonto - nMontoPagado FROM ColocCalendDetPig WHERE cCtaCod = '" & psCtaCod & "' AND nNroCalen = CP.nNumCalend AND " _
          & "               nColocCalendApl = '" & gColoCalendAplCuota & "' AND nCuota = 1 AND nPrdConceptoCod = '" & gColPigConceptoCodComiServ & "') " _
          & "  FROM Colocaciones C INNER JOIN Producto P ON C.cCtaCod = P.cCtaCod " _
          & "                                   LEFT JOIN ColocPigno CP ON C.cCtaCod = CP.cCtaCod " _
          & " WHERE C.cCtaCod = '" & psCtaCod & "'"
    
    Set lrDatos = oConn.CargaRecordSet(lsSQL)
    
    If lrDatos Is Nothing Or (lrDatos.BOF And lrDatos.EOF) Then
        Set lrDatos = Nothing
        MsgBox "Datos no encontrados", vbInformation, "Aviso"
        Set lrDatos = Nothing
        Exit Function
    End If
        
    Set dObtieneDetalleCredPignoraticio = lrDatos
    Set lrDatos = Nothing
    
End Function

'*************************************************************************************
' CMCPL - OBTIENE EL DETALLE DE LOS DESEMBOLSOS DE UN CONTRATO
' EAFA - 23/09/02
'*************************************************************************************
Public Function dObtieneDesembolsosContrato(ByVal psCtaCod As String) As Recordset
Dim Rs As ADODB.Recordset

On Error GoTo dError

sql = "SELECT MC.nNroCalen, M.cMovNro, MC.nMovNro, "
sql = sql & " DescAge = (Select cAgeDescripcion From Agencias Where cAgeCod=Substring(cMovNro, 18,2)),"
sql = sql & " Sum(CASE WHEN nPrdConceptoCod =" & gColPConceptoCodCapital & " then MCD.nMonto Else 0 End) Capital,"
sql = sql & " Sum(CASE WHEN nPrdConceptoCod =" & gColPConceptoCodTasacion & "  then MCD.nMonto Else 0 End) ComTasacion "
sql = sql & " FROM MovCol MC INNER JOIN MOv M ON M.nMovNro = MC.nMovNro  and M.nMovFlag = '0' "
sql = sql & " INNER JOIN MovColDet MCD ON MCD.nMovNro = MC.nMovNro "
sql = sql & " AND MCD.cOpeCod = MC.cOpeCod "
sql = sql & " WHERE MC.cCtaCod = '" & psCtaCod & "' and MC.cOpeCod IN ('" & gColPOpeDesembolsoEFE & "')"
sql = sql & " GROUP BY MC.nNroCalen, M.cMovNro, MC.nMovNro "
sql = sql & " ORDER BY MC.nMovNro "


'sql = "  SELECT MC.nNroCalen, M.cMovNro, MC.nMovNro, " _
'       & "  DescAge = (Select cAgeDescripcion From Agencias Where cAgeCod=Substring(cMovNro, 18,2)), " _
'       & "              Sum(CASE WHEN nPrdConceptoCod = " & gColPigConceptoCodCapital & "  then MCD.nMonto Else 0 End) Capital, " _
'       & "              Sum(CASE WHEN nPrdConceptoCod = " & gColPigConceptoCodTasacion & "  then MCD.nMonto Else 0 End) ComTasacion " _
'       & "    FROM MovCol MC INNER JOIN MOv M ON M.nMovNro = MC.nMovNro  and M.nMovFlag = '" & gMovFlagVigente & "'" _
'       & "                               INNER JOIN MovColDet MCD ON MCD.nMovNro = MC.nMovNro " _
'       & "                                                                         AND MCD.cOpeCod = MC.cOpeCod " _
'       & " WHERE MC.cCtaCod = '" & psCtaCod & "' and MC.cOpeCod IN ('" & gPigOpeDesembolsoEFE & "','" & gPigOpeDesembolsoUL & "')" _
'       & " GROUP BY MC.nNroCalen, M.cMovNro, MC.nMovNro " _
'       & " ORDER BY MC.nMovNro "
   
Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneDesembolsosContrato = Rs
Set Rs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos de Desembolsos en <<dObtieneDesembolsosContrato>>", Err.Description
 
End Function

'****************************************************************************
' CMCPL - OBTIENE EL DETALLE DE LOS PAGOS DE UN CONTRATO
' EAFA - 23/09/02
'****************************************************************************
Public Function dObtienePagosContrato(ByVal psCtaCod As String) As Recordset
Dim Rs As ADODB.Recordset

On Error GoTo dError

sql = "  SELECT MC.nNroCalen, MC.nMonto, MC.nMovNro, M.cMovNro, M.cMovDesc, " _
       & "  DescAge = (Select cAgeDescripcion From Agencias Where cAgeCod=Substring(cMovNro, 18,2)), " _
       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPConceptoCodCapital & "  then MCD.nMonto Else 0 End) Capital, " _
       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPConceptoCodInteresCompensatorio & "  then MCD.nMonto else 0 end) intCompensatorio, " _
       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPConceptoCodInteresMoratorio & "  then MCD.nMonto else 0 end) IntMoratorio, " _
       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPConceptoCodCustodia & "  then MCD.nMonto else 0 end) ComServicio, " _
       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPConceptoCodCustodiaVencida & "  then MCD.nMonto else 0 end) ComVencida, " _
       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPConceptoCodImpuesto & "  then MCD.nMonto else 0 end) ComVencida, " _
       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPConceptoCodPreparaRemate & "  then MCD.nMonto else 0 end) DerRemate, " _
       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPConceptoCodCustodiaDiferida & "  then MCD.nMonto else 0 end) CustDiferida " _
       & "  FROM MovCol MC INNER JOIN MOV M ON M.nMovNro = MC.nMovNro  and M.nMovFlag = '" & gMovFlagVigente & "'" _
       & "      INNER JOIN MovColDet MCD ON MCD.nMovNro = MC.nMovNro AND MCD.cOpeCod = MC.cOpeCod " _
       & " WHERE MC.cCtaCod = '" & psCtaCod & "' AND MC.cOpeCod IN ('" & gColPOpeAmortNorEFE & "','" & gColPOpeAmortNorCHQ & "','" & gColPOpeRenovNorEFE & "','" & gColPOpeRenovMorEFE & "'" _
       & ",'" & gColPOpeRenovNorCHQ & "','" & gColPOpeRenovMorCHQ & "','" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCancelNorCHQ & "','" & gColPOpeCancelMorCHQ & "'" _
       & ",'" & gColPOpeRenovNorCHQ & "','" & gPigOpeCancelMorEFE & "','" & gColPOpeCobCusDiferida & "','" & gColPOpeDevJoyas & "')" _
       & " GROUP BY MC.nNroCalen, MC.nMonto, MC.nMovNro, M.cMovNro, M.cMovDesc " _
       & " ORDER BY MC.nMovNro "

'& "  Sum(CASE WHEN nPrdConceptoCod = " & gColPConceptoCodCustodiaDiferida & "  then MCD.nMonto else 0 end) Penalidad, " _


'sql = "  SELECT MC.nNroCalen, MC.nMonto, MC.nMovNro, M.cMovNro, M.cMovDesc, " _
'       & "  DescAge = (Select cAgeDescripcion From Agencias Where cAgeCod=Substring(cMovNro, 18,2)), " _
'       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPigConceptoCodCapital & "  then MCD.nMonto Else 0 End) Capital, " _
'       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPigConceptoCodInteresComp & "  then MCD.nMonto else 0 end) intCompensatorio, " _
'       & "  Sum(CASE WHEN nPrdConceptoCod = " & gcolPigConceptoCodInteresMora & "  then MCD.nMonto else 0 end) IntMoratorio, " _
'       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPigConceptoCodComiServ & "  then MCD.nMonto else 0 end) ComServicio, " _
'       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPigConceptoCodComiVencida & "  then MCD.nMonto else 0 end) ComVencida, " _
'       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPigConceptoCodPreparaRemate & "  then MCD.nMonto else 0 end) DerRemate, " _
'       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPigConceptoCodPenalidad & "  then MCD.nMonto else 0 end) Penalidad, " _
'       & "  Sum(CASE WHEN nPrdConceptoCod = " & gColPigConceptoCodCustodia & "  then MCD.nMonto else 0 end) CustDiferida " _
'       & "  FROM MovCol MC INNER JOIN MOV M ON M.nMovNro = MC.nMovNro  and M.nMovFlag = '" & gMovFlagVigente & "'" _
'       & "      INNER JOIN MovColDet MCD ON MCD.nMovNro = MC.nMovNro AND MCD.cOpeCod = MC.cOpeCod " _
'       & " WHERE MC.cCtaCod = '" & psCtaCod & "' AND MC.cOpeCod IN ('" & gPigOpeAmortNorUL & "','" & gPigOpeAmortNorEFE & "','" & gPigOpeAmortMorEFE _
'       & "','" & gPigOpeCancelNorEFE & "','" & gPigOpeCancelMorEFE & "','" & gPigOpeCobCusDiferida & "','" & gPigOpeDevJoyas & "','" & gPigOpeAmortMorUL & "') " _
'       & " GROUP BY MC.nNroCalen, MC.nMonto, MC.nMovNro, M.cMovNro, M.cMovDesc " _
'       & " ORDER BY MC.nMovNro "

Set Rs = oConn.CargaRecordSet(sql)

Set dObtienePagosContrato = Rs
Set Rs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos de Pagos en <<dObtienePagosContrato>>", Err.Description
 
End Function

'**********************************************************************
'CMCPL OBTIENE DATOS DEL CLIENTE PARA LA VENTA
'EAFA - 22/10/2002
'**********************************************************************
Public Function dObtieneDatosClienteVentaJoyas(ByVal psCliente As String) As Recordset

Dim lrDatos As ADODB.Recordset
Dim lsSQL As String

lsSQL = " SELECT P.cPersCod, P.cPersNombre, P.cPersTelefono, P.cPersDireccDomicilio, " _
          & "                   P.nPersPersoneria, " _
          & " NroDNI = (Select ISNULL(cPersIDnro,'') From PersID Where cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdDNI & " ), " _
          & " NroRUC = (Select ISNULL(cPersIDnro,'') From PersID Where cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdRUC & " ) " _
          & " FROM Persona P   " _
          & " WHERE P.cPersCod = '" & psCliente & "'"
    
Set lrDatos = oConn.CargaRecordSet(lsSQL)
    
Set dObtieneDatosClienteVentaJoyas = lrDatos
Set lrDatos = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos de Cliente para Venta de Joyas en <<dObtieneDatosClienteVentaJoyas>>", Err.Description

    
End Function

'****************************************************************************
' CMCPL - OBTIENE EL DETALLE DE UNA JOYA PARA LA VENTA
' EAFA - 02/10/2002
'****************************************************************************
Public Function dObtieneDetalleJoyaVenta(ByVal psCtaCod As String, ByVal pnPieza As Integer) As Recordset

Dim Rs As ADODB.Recordset

On Error GoTo dError

sql = " SELECT CPJT.nItemPieza Item, CPJT.nUbicaPieza, CPP.nSituacionPieza, CPEJ.nRemate, " _
       & " Tipo = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigTipoJoya & " AND nConsValor = CPJT.nTipoJoya), " _
       & " Material = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigMaterial & " AND nConsValor = CPJT.nMaterial), " _
       & " SubTipo = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigSubTipoJoya & " AND nConsValor = CPJT.nSubTipoJoya), " _
       & " Estado = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigEstConservaJoya & " AND nConsValor = CPJT.nConservacion), " _
       & " nTasador = (SELECT cUser FROM RRHH WHERE cPersCod = CPJT.cPersCod), " _
       & " CPJT.nPesoBruto PBruto, CPJT.nPesoNeto PNeto, CPJT.nTasacion Tasacion, CPJT.nTasacionAdicional TasAdicion, CPJT.cDescripcion Observacion, " _
       & " CPJT.cDescripcionAdic ObservAdic, CPP.nValorDeuda, CPP.nValorDeudaComp " _
       & " FROM ColocPigJoyaTasacion CPJT INNER JOIN ColocPigEstadoJoya CPEJ ON CPEJ.cCtaCod = CPJT.cCtaCod and CPEJ.nItemPieza = CPJT.nItemPieza and CPEJ.cFlag = '0' " _
       & "    INNER JOIN ColocPigProceso CPP ON CPP.nRemate = CPEJ.nRemate and CPP.nTipoProceso = " & gPigTipoVentas _
       & "                                                                                                                                and CPP.cCtaCod = CPJT.cCtaCod and CPP.nItemPieza = CPJT.nItemPieza " _
       & " WHERE CPJT.cCtaCod = '" & psCtaCod & "' AND CPJT.nTipoTasacion = " & gPigTipoTasacNor & " and CPJT.nItemPieza = " & pnPieza

Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneDetalleJoyaVenta = Rs
Set Rs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos de Joyas en <<dObtieneDetalleJoyas>>", Err.Description

End Function

'***********************************************************************
' CMCPL - OBTIENE EL DETALLE DE UNA BOLETA/FACTURA POR VENTAS DE JOYAS
' EAFA - 15/10/2002 ' MODIFICADO DFAO - 16/06/2003
'***********************************************************************
Public Function dObtieneDetDocumentoVentasJoyas(ByVal pnTipoDoc As Integer, ByVal psNroDoc As String) As Recordset
Dim Rs As ADODB.Recordset

On Error GoTo dError

sql = " SELECT t1.nCodTipo, t2.nItemDoc, t2.cDocumento, t2.cCtaCod, t2.nPieza, t3.nMaterial, t3.nTipoJoya, t3.nSubTipoJoya, t3.nConservacion,  " _
        & "    t3.nPesoNeto pNeto, t3.nTasacionAdicional, t3.cDescripcion Observacion, t3.cDescripcionAdic ObservAdic, t1.cPersCod, t1.nMotivo, t1.nNroMov, " _
        & "    Sum(Case When t2.nCodConcepto = " & gColPigConceptoCodPrecioVenta & " then t2.nMonto else 0 end) SubVenta, " _
        & "    Sum(Case When t2.nCodConcepto = " & gColPigConceptoCodIgv & " then t2.nMonto else 0 end) IGV, " _
        & " nFlag = (SELECT t5.nMovFlag FROM Mov t5 WHERE t5.nMovNro = t1.nNroMov ), " _
        & " nRema = (SELECT t4.nRemate FROM ColocPigProceso t4 WHERE t4.cCtaCod=t2.cCtaCod and t4.nItemPieza=t2.nPieza and t4.nTipoProceso= " & gPigTipoVentas & "), " _
        & " Tipo = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigTipoJoya & " AND nConsValor = t3.nTipoJoya), " _
        & " Material = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigMaterial & " AND nConsValor = t3.nMaterial), " _
        & " SubTipo = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigSubTipoJoya & " AND nConsValor = t3.nSubTipoJoya), " _
        & " Estado = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigEstConservaJoya & " AND nConsValor = t3.nConservacion) " _
        & "  FROM MovDocPig t1 Inner Join MovDocPigDet t2 on t2.nNroMov=t1.nNroMov " _
        & "  INNER JOIN ColocPigno t6 on t6.cCtaCod=t2.cCtaCod " _
        & "  INNER JOIN ColocPigJoyaTasacion t3 on t3.cCtaCod=t2.cCtaCod AND t3.nItemPieza=t2.nPieza " _
        & "  AND t3.nTipoTasacion= t6.nTipoTasacion " _
        & " WHERE t1.nCodTipo = " & pnTipoDoc & "AND   t1.cDocumento = '" & psNroDoc & "'" _
        & " GROUP BY t1.nCodTipo, t2.nItemDoc, t2.cDocumento, t2.cCtaCod, t2.nPieza,t3.nMaterial, t3.nTipoJoya, t3.nSubTipoJoya, t3.nConservacion, " _
        & "     t3.nPesoNeto, t3.nTasacionAdicional, t3.cDescripcion, t3.cDescripcionAdic, t1.cPersCod, t1.nMotivo, t1.nNroMov "
   
Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneDetDocumentoVentasJoyas = Rs
Set Rs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos de Joyas en <<dObtieneDetDocumentoVentasJoyas>>", Err.Description
 
End Function

'***********************************************************************
' CMCPL - OBTIENE EL DETALLE DE UNA POLIZA POR VENTAS DE JOYAS EN REMATE
' DFAO - 13/06/2003
'***********************************************************************
Public Function dObtieneDetDocumentoVentasJoyasRemate(ByVal pnTipoDoc As Integer, ByVal psNroDoc As String, ByVal pnestJoya As Integer) As Recordset
Dim Rs As ADODB.Recordset

On Error GoTo dError

sql = " SELECT t1.nCodTipo, t2.nItemDoc, t2.cDocumento, t2.cCtaCod, t2.nPieza, t3.nMaterial, t3.nTipoJoya, t3.nSubTipoJoya, t3.nConservacion,  " _
        & "    t3.nPesoNeto pNeto, t3.nTasacionAdicional, t3.cDescripcion Observacion, t3.cDescripcionAdic ObservAdic, t1.cPersCod, t1.nMotivo, t1.nNroMov, " _
        & "    Sum(Case When t2.nCodConcepto = " & gColPigConceptoCodPrecioVenta & " then t2.nMonto else 0 end) SubVenta, " _
        & "    Sum(Case When t2.nCodConcepto = " & gColPigConceptoCodIgv & " then t2.nMonto else 0 end) IGV, " _
        & " nFlag = (SELECT t5.nMovFlag FROM Mov t5 WHERE t5.nMovNro = t1.nNroMov ), " _
        & " nRema = (SELECT t4.nRemate FROM ColocPigProceso t4 WHERE t4.cCtaCod=t2.cCtaCod and t4.nItemPieza=t2.nPieza and t4.nSituacionPieza = " & pnestJoya & " and t4.cComprador is Not Null  ), " _
        & " nTipoProceso = (SELECT t8.nTipoProceso FROM ColocPigProceso t8 WHERE t8.cCtaCod=t2.cCtaCod and t8.nItemPieza=t2.nPieza and t8.nSituacionPieza = " & pnestJoya & " and t8.cComprador is Not Null  ), " _
        & " Comprador = (SELECT t7.cComprador FROM ColocPigProceso t7 WHERE t7.cCtaCod=t2.cCtaCod and t7.nItemPieza=t2.nPieza and t7.nSituacionPieza = " & pnestJoya & " and t7.cComprador is Not Null  ), " _
        & " Tipo = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigTipoJoya & " AND nConsValor = t3.nTipoJoya), " _
        & " Material = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigMaterial & " AND nConsValor = t3.nMaterial), " _
        & " SubTipo = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigSubTipoJoya & " AND nConsValor = t3.nSubTipoJoya), " _
        & " Estado = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigEstConservaJoya & " AND nConsValor = t3.nConservacion) " _
        & "  FROM MovDocPig t1 Inner Join MovDocPigDet t2 on t2.nNroMov=t1.nNroMov " _
        & "  INNER JOIN ColocPigJoyaTasacion t3 on t3.cCtaCod=t2.cCtaCod AND t3.nItemPieza=t2.nPieza " _
        & "  AND t3.nTipoTasacion= " & gPigTipoTasacRetasac _
        & " WHERE t1.nCodTipo = " & pnTipoDoc & "AND   t1.cDocumento = '" & psNroDoc & "'" _
        & " GROUP BY t1.nCodTipo, t2.nItemDoc, t2.cDocumento, t2.cCtaCod, t2.nPieza,t3.nMaterial, t3.nTipoJoya, t3.nSubTipoJoya, t3.nConservacion, " _
        & "     t3.nPesoNeto, t3.nTasacionAdicional, t3.cDescripcion, t3.cDescripcionAdic, t1.cPersCod, t1.nMotivo, t1.nNroMov "
   
Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneDetDocumentoVentasJoyasRemate = Rs
Set Rs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos de Joyas en <<dObtieneDetDocumentoVentasJoyasRemate>>", Err.Description
 
End Function

'****************************************************************************
' CMCPL - OBTIENE JOYAS PARA LA SELECCION VENTA/FUNDICION
' EAFA - 24/10/2002
'****************************************************************************
Public Function dObtieneJoyasVentaFundicion(ByVal psCondicion As String, ByVal psCondicionAdic As String, ByVal pnTipoProceso As Integer) As Recordset

Dim Rs As ADODB.Recordset

On Error GoTo dError

sql = " SELECT CPJT.cCtaCod, CPJT.nItemPieza Item,  CPEJ.nRemate, CPP.nValorDeuda, CPP.nValorDeudaComp, CPP.nValorProceso, " _
       & " Tipo = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigTipoJoya & " AND nConsValor = CPJT.nTipoJoya), " _
       & " Material = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigMaterial & " AND nConsValor = CPJT.nMaterial), " _
       & " OtroProceso = (SELECT T.nTipoProceso FROM ColocPigProceso T WHERE T.nRemate  = CPEJ.nRemate and T.cCtaCod = CPJT.cCtaCod and T.nItemPieza = CPJT.nItemPieza and T.nTipoProceso = " & pnTipoProceso & "), " _
       & " DescEstado = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigTipoProceso & " AND nConsValor = CPEJ.nEstadoJoya), " _
       & " CPJT.nPesoNeto PNeto, CPEJ.nEstadoJoya nTipoProceso " _
       & " FROM ColocPigJoyaTasacion CPJT INNER JOIN ColocPigEstadoJoya CPEJ ON CPEJ.cCtaCod = CPJT.cCtaCod and CPEJ.nItemPieza = CPJT.nItemPieza " & psCondicionAdic & " and CPEJ.cFlag = '0' " _
       & "      INNER JOIN ColocPigProceso CPP ON CPP.nRemate = CPEJ.nRemate and CPP.nTipoProceso = " & gPigTipoAdjudicacionCaja _
       & "      AND CPP.cCtaCod = CPJT.cCtaCod and CPP.nItemPieza = CPJT.nItemPieza "

sql = sql & psCondicion & " And nUbicaPieza = " & gPigUbicaBovOfPrin

Set Rs = oConn.CargaRecordSet(sql)

Set dObtieneJoyasVentaFundicion = Rs
Set Rs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos de Joyas en <<dObtieneJoyasVentaFundicion>>", Err.Description

End Function

Public Function dObtieneValorTasacion(ByVal psCodCta As String) As Currency
Dim Rs As Recordset

    On Error GoTo dError
    sql = "SELECT SUM(nTasacion + nTasacionAdicional) Tasacion FROM COLOCPIGJOYATASACION WHERE nTipoTasacion = 1 AND " _
            & "cCtaCod = '" & psCodCta & "'"
    
    Set Rs = oConn.CargaRecordSet(sql)
    
    If Not IsNull(Rs!Tasacion) Then
        dObtieneValorTasacion = Rs!Tasacion
    Else
        dObtieneValorTasacion = 0
    End If
    
    Set Rs = Nothing
    
    Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Valor de Tasacion <<dObtieneValorTasacion>>", Err.Description

End Function

Public Function dObtieneValorRemesa(ByVal psNumGuia As String) As Recordset
Dim Rs As Recordset

    On Error GoTo dError
    
    sql = "SELECT GD.cCtaCod, P.nPrdEstado, Sum(nTasacion + nTasacionAdicional) Tasacion FROM COLOCPIGGUIADET GD " _
        & " INNER JOIN COLOCPIGJOYATASACION J ON GD.cCtaCod = J.cCtaCod " _
        & " INNER JOIN COLOCPIGGUIA G ON G.cNumDoc = GD.cNumDoc INNER JOIN COLOCPIGNO C ON C.cCtaCod = J.cCtaCod " _
        & " INNER JOIN PRODUCTO P ON P.cCtaCod = GD.cCtaCod " _
        & " WHERE C.nTipoTasacion = J.nTipoTasacion AND G.cNumDoc = '" & psNumGuia & "'" _
        & " GROUP BY GD.cCtaCod, P.nPrdEstado"
        
    Set Rs = oConn.CargaRecordSet(sql)
    
    If Not (Rs.EOF And Rs.BOF) Then
        Set dObtieneValorRemesa = Rs
    End If
    
    Set Rs = Nothing
    
Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Valor de Remesa <<dObtieneValorRemesa>>", Err.Description

End Function

Public Function dObtieneMovExtornoUsoLinea(ByVal pnMovNro As Long, ByVal psOpeCod As String) As Long
Dim Rs As Recordset

    On Error GoTo dError
    If psOpeCod = gPigOpeAmortizEFE Then
        sql = "SELECT MR.nMovNro, cOpeCod FROM MOVREF MR INNER JOIN MOVCOL MC ON MR.nMovNro = MC.nMovNro " _
            & "WHERE nMovNroRef = " & pnMovNro & "AND cOpeCod like '" & Mid(psOpeCod, 1, 4) & "%'  AND SUBSTRING(cOpeCod,6,1) IN (3,4)"
    Else
        sql = "SELECT MR.nMovNro, cOpeCod FROM MOVREF MR INNER JOIN MOVCOL MC ON MR.nMovNro = MC.nMovNro " _
            & "WHERE nMovNroRef = " & pnMovNro & "AND cOpeCod like '" & Mid(psOpeCod, 1, 4) & "%'  AND SUBSTRING(cOpeCod,6,1) IN (2)"
    End If
    Set Rs = oConn.CargaRecordSet(sql)
    
    If Not (Rs.EOF And Rs.BOF) Then
        dObtieneMovExtornoUsoLinea = Rs!nMovNro
    End If
    
    Set Rs = Nothing
    Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Mov de Extorno <<dObtieneMovExtornoUsoLinea>>", Err.Description
    
End Function

Public Function dObtieneDeudaTotal(ByVal psCodPers As String, ByVal pdFecSis As String) As Currency
Dim Rs As Recordset

    On Error GoTo dError

    sql = "SELECT SUM(DeudaTotal) DeudaTotal FROM (" _
        & "SELECT (nSaldo + dbo.ColocPignoInteres(TI1.nTasaInteres, datediff(dd, dPrdEstado, '" _
        & Format(pdFecSis, "mm/dd/yyyy") & "'), nSaldo) + " _
        & "dbo.ColocPignoInteres(TI2.nTasaInteres, datediff(dd, dVenc, '" & Format(pdFecSis, "mm/dd/yyyy") & "'), nSaldo) + " _
        & "ISNULL((Select nMonto - nMontoPagado  From ColocCalendDetPig where cCtaCod = P.cCtaCod " _
        & "AND nNroCalen = CP.nNumCalend and nColocCalendApl= 1 and nCuota=1 and nPrdConceptoCod= 8205), 0) + " _
        & "ISNULL((Select nMonto - nMontoPagado From ColocCalendDetPig where cCtaCod = P.cCtaCod " _
        & "AND nNroCalen= CP.nNumCalend and nColocCalendApl= 1 and nCuota =1 and nPrdConceptoCod= 8206), 0) + " _
        & "ISNULL((Select nMonto - nMontoPagado FROM ColocCalendDetPig WHERE cCtaCod = P.cCtaCod " _
        & "AND nNroCalen = CP.nNumCalend AND nColocCalendApl = 1 AND nCuota = 1 AND nPrdConceptoCod = 8203), 0)) DeudaTotal " _
        & "FROM PRODUCTOPERSONA PP INNER JOIN PRODUCTO P ON PP.cCtaCod = P.cCtaCod " _
        & "INNER JOIN COLOCACIONES C ON C.cCtaCod = PP.cCtaCod " _
        & "INNER JOIN COLOCPIGNO CP ON C.cCtaCod = CP.cCtaCod " _
        & "INNER JOIN PRODUCTOTASAINTERES TI1 ON P.cCtaCod = TI1.cCtaCod " _
        & "INNER JOIN PRODUCTOTASAINTERES TI2 ON P.cCtaCod = TI2.cCtaCod " _
        & "Where TI1.nPrdTasaInteres = 200 And TI2.nPrdTasaInteres = 210 " _
        & "AND cPersCod = '" & psCodPers & "' AND NPRDESTADO IN (2802,2803,2804, 2807)) P"
        
    Set Rs = oConn.CargaRecordSet(sql)
    
    If Not (Rs.EOF And Rs.BOF) Then
        dObtieneDeudaTotal = IIf(IsNull(Rs!DeudaTotal), 0, Rs!DeudaTotal)
    End If
    
    Set Rs = Nothing
    Exit Function
    
dError:
    Err.Raise Err.Number, "Obtiene Deuda Total <<dObtieneDeudaTotal>>", Err.Description
    
End Function

Public Function dObtieneDeudaTotalDetallada(ByVal psCodPers As String, ByVal pdFecSis As String) As Recordset

    On Error GoTo dError
   
    sql = "SELECT C.cCtaCod, cConsDescripcion Estado, cConsDescripcion Estado, " _
        & "(nSaldo + dbo.ColocPignoInteres(TI1.nTasaInteres, datediff(dd, dPrdEstado, '" _
        & Format(pdFecSis, "mm/dd/yyyy") & "'), nSaldo) + " _
        & "dbo.ColocPignoInteres(TI2.nTasaInteres, datediff(dd, dVenc, '" & Format(pdFecSis, "mm/dd/yyyy") & "'), nSaldo) + " _
        & "ISNULL((Select nMonto - nMontoPagado  From ColocCalendDetPig where cCtaCod = P.cCtaCod " _
        & "AND nNroCalen = CP.nNumCalend and nColocCalendApl= 1 and nCuota=1 and nPrdConceptoCod= 8205), 0) + " _
        & "ISNULL((Select nMonto - nMontoPagado From ColocCalendDetPig where cCtaCod = P.cCtaCod " _
        & "AND nNroCalen= CP.nNumCalend and nColocCalendApl= 1 and nCuota =1 and nPrdConceptoCod= 8206), 0) + " _
        & "ISNULL((Select nMonto - nMontoPagado FROM ColocCalendDetPig WHERE cCtaCod = P.cCtaCod " _
        & "AND nNroCalen = CP.nNumCalend AND nColocCalendApl = 1 AND nCuota = 1 AND nPrdConceptoCod = 8203), 0)) DeudaTotal " _
        & "FROM PRODUCTOPERSONA PP INNER JOIN PRODUCTO P ON PP.cCtaCod = P.cCtaCod " _
        & "INNER JOIN COLOCACIONES C ON C.cCtaCod = PP.cCtaCod " _
        & "INNER JOIN COLOCPIGNO CP ON C.cCtaCod = CP.cCtaCod " _
        & "INNER JOIN PRODUCTOTASAINTERES TI1 ON P.cCtaCod = TI1.cCtaCod " _
        & "INNER JOIN PRODUCTOTASAINTERES TI2 ON P.cCtaCod = TI2.cCtaCod " _
        & "LEFT JOIN CONSTANTE ON nConsValor = nPrdEstado AND nConsCod = 3001 " _
        & "Where TI1.nPrdTasaInteres = 200 And TI2.nPrdTasaInteres = 210 " _
        & "AND cPersCod = '" & psCodPers & "' AND NPRDESTADO IN (2802,2803,2804, 2807) " _
        & "ORDER BY C.cCtaCod"

    Set dObtieneDeudaTotalDetallada = oConn.CargaRecordSet(sql)
    
    Exit Function
    
dError:
    Err.Raise Err.Number, "Obtiene Deuda Total Detallada <<dObtieneDeudaTotalDetallada>>", Err.Description
    
End Function

'********************************************************************
Public Function dObtieneCreditosPignoPersona(ByVal psPersCod As String) As Recordset
'Obtiene Datos de Credito Pignoraticio Joyas
Dim lrs As ADODB.Recordset

On Error GoTo dError

sql = "SELECT P.cCtaCod, nPiezas, nPrdEstado, dPrdEstado, nDiasAtraso, nMontoCol, dVenc, " _
    & "Descri = (Select cConsDescripcion FROM Constante E WHERE E.nConsValor = P.nPrdEstado " _
    & "AND nConsCod = 3001) FROM PRODUCTO P INNER JOIN ColocPigno CP " _
    & "ON P.cCtaCod = CP.cCtaCod INNER JOIN Colocaciones C ON C.cCtaCod = CP.cCtaCod " _
    & "INNER JOIN PRODUCTOPERSONA PP ON P.cCtaCod = PP.cCtaCod " _
    & "WHERE PP.cPersCod = '" & psPersCod & "' ORDER BY P.cCtaCod"

Set lrs = oConn.CargaRecordSet(sql)

Set dObtieneCreditosPignoPersona = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

'****************************************************************************
' CMCPL - PASE A FUNDICION
' DFAO - 10/06/2003

'****************************************************************************
Public Function dObtieneDatosFundicion(ByVal pdTipo, ByVal pdComDesde As String, ByVal pdComHasta As String) As Recordset

Dim Rs As ADODB.Recordset

On Error GoTo dError

If pdTipo = "T" Then
    sql = " select CPEJ.nremate Remate,count(*) Piezas,SUM(CPJT.nPesoBruto)PesoBruto,SUM(CPJT.nPesoNeto) PesoNeto,SUM(CPJT.nTasacion) Tasacion,SUM(CPJT.nTasacionAdicional) TasacionAdicional " _
       & " FROM ColocPigEstadoJoya CPEJ " _
       & " INNER JOIN ColocPigJoyaTasacion CPJT on CPJT.cCtaCod =CPEJ.cCtacod and CPJT.nItemPieza = CPEJ.nItemPieza " _
       & " INNER JOIN ColocPigProceso CPP on CPP.nRemate=CPEJ.nRemate and CPP.nTipoProceso = CPEJ.nEstadoJoya and CPP.cCtaCod =CPEJ.cCtacod and CPP.nItemPieza = CPEJ.nItemPieza " _
       & " WHERE CPP.nSituacionPieza= " & gPigSituacionDisponible & " and CPEJ.nEstadoJoya = " & gPigTipoFundicion & " and CPEJ.nRemate between " & Val(pdComDesde) & " and " & Val(pdComHasta) _
       & " Group by CPEJ.nremate "
    Set Rs = oConn.CargaRecordSet(sql)
Else
    sql = " select CPEJ.nremate Remate, CPEJ.nEstadoJoya, CPEJ.cCtaCod, CPEJ.nItemPieza" _
       & " FROM ColocPigEstadoJoya CPEJ " _
       & " INNER JOIN ColocPigJoyaTasacion CPJT on CPJT.cCtaCod =CPEJ.cCtacod and CPJT.nItemPieza = CPEJ.nItemPieza " _
       & " INNER JOIN ColocPigProceso CPP on CPP.nRemate=CPEJ.nRemate and CPP.nTipoProceso = CPEJ.nEstadoJoya and CPP.cCtaCod =CPEJ.cCtacod and CPP.nItemPieza = CPEJ.nItemPieza " _
       & " WHERE CPP.nSituacionPieza= " & gPigSituacionDisponible & " and CPEJ.nEstadoJoya = " & gPigTipoFundicion & " and CPEJ.nRemate between " & Val(pdComDesde) & " and " & Val(pdComHasta)
    Set Rs = oConn.CargaRecordSet(sql)
End If

Set dObtieneDatosFundicion = Rs
Set Rs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos de Joyas en <<dObtieneJoyasVentaFundicion>>", Err.Description

End Function


'***********************************************************************
' OBTIENE EL DETALLE DE RESCATES DE SOBRANTE DE PIEZAS DE UN CONTRATO
' JAPP - 30/06/03
'***********************************************************************
Public Function dObtieneDatosSobrantePieza(ByVal psCuenta As String) As Recordset

Dim lrs As ADODB.Recordset
Dim oRema As DPigFunciones
Dim lsRemate As Integer

    Set oRema = New DPigFunciones
    lsRemate = oRema.GetRemate(psCuenta)
    Set oRema = Nothing
    
On Error GoTo dError

sql = "SELECT t1.nItemPieza Item, " _
    & "Tipo = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigTipoJoya & " AND nConsValor = t2.nTipoJoya), " _
    & "SubTipo = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigSubTipoJoya & " AND nConsValor = t2.nSubTipoJoya), " _
    & "Material = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigMaterial & " AND nConsValor = t2.nMaterial), " _
    & "Estado = (SELECT cConsDescripcion FROM Constante WHERE nConsCod = " & gColocPigEstConservaJoya & " AND nConsValor = t2.nConservacion), " _
    & "nTipo = (SELECT nConsValor FROM Constante WHERE nConsCod = " & gColocPigTipoJoya & " AND nConsValor = t2.nTipoJoya), " _
    & "nSubTipo = (SELECT nConsValor FROM Constante WHERE nConsCod = " & gColocPigSubTipoJoya & " AND nConsValor = t2.nSubTipoJoya), " _
    & "nMaterial = (SELECT nConsValor FROM Constante WHERE nConsCod = " & gColocPigMaterial & "  AND nConsValor = t2.nMaterial), " _
    & "nEstado = (SELECT nConsValor FROM Constante WHERE nConsCod = " & gColocPigEstConservaJoya & " AND nConsValor = t2.nConservacion), " _
    & "nPesoBruto PBruto, nPesoNeto PNeto, nTasacion Tasacion, nTasacionAdicional TasAdicion, cDescripcion Observacion, " _
    & "cDescripcionAdic as ObsAdicion, t2.nTipoTasacion " _
    & "From ColocPigProceso t1 Inner Join ColocPigJoyaTasacion t2 " _
    & "On t1.cCtaCod = t2.cCtaCod and t1.nItemPieza = t2.nItemPieza " _
    & "Where t1.nTipoProceso = 4 and t2.nTipoTasacion = 3 " _
    & "and t1.nSituacionPieza = 1 and t1.nRemate = " & lsRemate & " " _
    & "and t1.cCtaCod = '" & psCuenta & "'"

Set lrs = oConn.CargaRecordSet(sql)

Set dObtieneDatosSobrantePieza = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description

End Function

Public Function dEvaluacionCliManual(ByVal psPersona As String) As Recordset
'Obtiene Evaluacion del Cliente
Dim lrs As ADODB.Recordset

On Error GoTo dError

sql = " select icalifica from colocpigevalmanual c "
sql = sql & " where cPersCod = " & psPersona & " and fEvalPigno =(select max(fevalpigno) from colocpigevalmanual where cPersCod=c.cPersCod) "

Set lrs = oConn.CargaRecordSet(sql)

Set dEvaluacionCliManual = lrs
Set lrs = Nothing

Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dEvaluacionCliManual>>", Err.Description
End Function

'**********************************************************************
' OBTIENE DATOS DEL CLIENTE DE PIGNORATICIO - BUSCADO POR EL CONTRATO
' CRSF - 02/09/02
'**********************************************************************
Public Function dClientePigno(ByVal psCtaCod As String) As Recordset
'Obtiene Datos de Cliente de Pignoraticio

Dim lrs As ADODB.Recordset

On Error GoTo dError

sql = "Select P.cPersCod, P.cPersnombre, P.cPersDireccDomicilio, PId.cPersIdNro" _
     & " From productopersona a Join Persona P On P.cPersCod  = A.cPersCod Join PersID PID On  PID.cPersCod = P.cPersCod" _
     & " Where PID.cPersIDTpo = 1 and a.cctacod = " & psCtaCod & ""
Set lrs = oConn.CargaRecordSet(sql)

Set dClientePigno = lrs
Set lrs = Nothing

Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Datos Cliente Pigno en <<dClientePigno >>", Err.Description
End Function


