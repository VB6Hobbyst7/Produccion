VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMColPContrato"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim coConex As COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String


Private Sub Class_Initialize()
    Dim loIni As New COMConecta.DCOMClasIni
    Set coConex = New COMConecta.DCOMConecta
    csConexion = loIni.CadenaConexion
    csNegocio = loIni.BaseNegocio
    csCentralPer = loIni.BasePersonas
    csCentralCom = loIni.BaseComunes
    csCentralImg = loIni.BaseImagenes
    csAdminist = loIni.BaseAdministracion
 
    If coConex.AbreConexion(csConexion) = False Then
        Call oError.RaiseError(oError.MyUnhandledError, "DColPContrato:Initialize. Error en Conexion a Base de datos")
    End If
End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

Public Function dObtieneDatosCreditoPignoraticio(ByVal psCuenta As String) As ADODB.Recordset
'Obtiene Datos de Contrato Pignoraticio
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New ADODB.Recordset

On Error GoTo dError


'lsSQL = "SELECT  P.cCtaCod, P.nSaldo, P.nPrdEstado, P.nTasaInteres, " _
'    & " C.nPlazo, C.dVenc, C.nMontoCol, C.dVigencia, " _
'    & " CP.nOroBruto, CP.nOroNeto, CP.nPiezas, CP.nTasacion,CP.cLote, CP.cPrdCtaTpo,  " _
'    & " nPlazoIni = ( SELECT nPlazo FROM ColocacEstado WHERE cCtaCod ='" & psCuenta & "' " _
'    & "                 AND nPrdEstado = " & gColPEstRegis & "  AND nPlazo > 0 ),  " _
'    & " dFecVencIni = ( SELECT DateAdd(d, nPlazo, dPrdEstado) FROM ColocacEstado WHERE cCtaCod ='" & psCuenta & "' " _
'    & "                 AND nPrdEstado = " & gColPEstRegis & "  AND nPlazo > 0 )," _
'    & " nIntDif=(Select nmonto from movcoldet mc join mov m on m.nmovnro=mc.nmovnro   where mc.copecod='120201' and nprdconceptocod=2100 and m.nmovflag=0 and cctacod=p.cctacod ), " _
'    & " cEstado=(Select cconsdescripcion from constante where nconscod='3001' and nconsvalor=P.nprdestado ), " _
'    & " cTipCta=case when CP.cPrdCtaTpo='0' then 'INDIVIDUAL' when CP.cPrdCtaTpo='1' then 'MANCOMUNADA' when CP.cPrdCtaTpo='2' then 'INDISTINTA'  end " _
'    & " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'    & " INNER JOIN ColocPignoraticio CP ON C.cCtaCod = CP.cCtaCod " _
'    & " WHERE P.cCtaCod = '" & psCuenta & "'"

'*** PEAC 20161207
' lsSQL = "  SELECT  P.cCtaCod, P.nSaldo, P.nPrdEstado, P.nTasaInteres,  "
' lsSQL = lsSQL & "  C.nPlazo, C.dVenc, C.nMontoCol, C.dVigencia,  "
' lsSQL = lsSQL & "  CP.nOroBruto, CP.nOroNeto, CP.nPiezas, CP.nTasacion,CP.cLote, CP.cPrdCtaTpo,CP.nNroDuplic , "
' lsSQL = lsSQL & " nPlazoIni, dFecVenIni ,"
' lsSQL = lsSQL & " nIntDif=(Select nmonto from movcoldet mc join mov m on m.nmovnro=mc.nmovnro   where mc.copecod='120201' and nprdconceptocod= " & gColPConceptoCodInteresCompensatorio & " and m.nmovflag=0 and cctacod=p.cctacod ),"
' lsSQL = lsSQL & " cEstado=(Select cconsdescripcion from constante where nconscod='3001' and nconsvalor=P.nprdestado ), "
' lsSQL = lsSQL & " cTipCta=case when CP.cPrdCtaTpo='0' then 'INDIVIDUAL' when CP.cPrdCtaTpo='1' then 'MANCOMUNADA' when CP.cPrdCtaTpo='2' then 'INDISTINTA'  end "
' lsSQL = lsSQL & " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod "
' lsSQL = lsSQL & " LEFT JOIN (SELECT CCTACOD, NPLAZOINI=NPLAZO,DFECVENINI=DateAdd(d, nPlazo, dPrdEstado)  FROM COLOCACESTADO WHERE nPrdEstado = " & gColPEstRegis & " AND nPlazo > 0  ) CE ON CE.CCTACOD=P.CCTACOD "
' lsSQL = lsSQL & " INNER JOIN ColocPignoraticio CP ON C.cCtaCod = CP.cCtaCod "
' lsSQL = lsSQL & " WHERE P.cCtaCod ='" & psCuenta & "'"

lsSQL = "stp_sel_ObtieneDatosCreditoPignoraticio '" & psCuenta & "' "
'*** FIN PEAC

Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneDatosCreditoPignoraticio = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

Public Function dObtieneDatosCreditoPignoraticioJoyas(ByVal psCuenta As String) As ADODB.Recordset
'Obtiene Datos de Credito Pignoraticio Joyas
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New ADODB.Recordset

On Error GoTo dError

lsSQL = "SELECT SUM(Isnull( CASE WHEN cKilataje = '14' THEN nPesoOro END , 0 )) AS nK14 , " _
      & "       SUM(Isnull( CASE WHEN ckilataje = '16' THEN nPesoOro end , 0 )) as nK16 , " _
      & "       SUM(Isnull( CASE WHEN ckilataje = '18' THEN nPesoOro end , 0 )) as nK18 , " _
      & "       SUM(Isnull( CASE WHEN ckilataje = '21' THEN nPesoOro end , 0 )) as nK21 " _
      & " FROM ColocPigJoya  " _
      & " WHERE cCtaCod = '" & psCuenta & "' "


Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneDatosCreditoPignoraticioJoyas = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

Public Function dObtieneTasasCreditoPignoraticio(ByVal psCuenta As String) As ADODB.Recordset
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New ADODB.Recordset

On Error GoTo dError

lsSQL = " Select " _
        & " nComp=(Select ntasainteres from productotasainteres where cctacod='" & psCuenta & "' and nprdtasainteres=1), " _
        & " nMora=(Select ntasainteres from productotasainteres where cctacod='" & psCuenta & "' and nprdtasainteres=3), " _
        & " nGracia=(Select ntasainteres from productotasainteres where cctacod='" & psCuenta & "' and nprdtasainteres=5), " _
        & " nMontoPrestamo=(Select nmonto from movcoldet mc join mov m on m.nmovnro=mc.nmovnro   where mc.copecod='120100' and nprdconceptocod= " & gColPConceptoCodCapital & " and m.nmovflag=0 and cctacod='" & psCuenta & "' )  "


Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneTasasCreditoPignoraticio = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description

End Function

Public Function dObtieneDatosCreditoPignoraticioJoyasDet(ByVal psCuenta As String, Optional ByVal PbTipo As Boolean = False) As ADODB.Recordset
'Obtiene Datos de Credito Pignoraticio Joyas
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New ADODB.Recordset

On Error GoTo dError
'RECO20141201 ERS163-2014**********************************
'If PbTipo = False Then
'    lsSQL = "SELECT cKilataje,nItem, nPiezas, nPesoBruto, nPesoNeto, nValTasac, cDescrip " _
'          & " FROM ColocPigJoyaDet  " _
'          & " WHERE cCtaCod = '" & psCuenta & "' "
'Else
'    lsSQL = "SELECT nPiezas,cKilataje  ,  nPesoBruto, nPesoNeto,nValTasac,cDescrip " _
'          & " FROM ColocPigJoyaDet  " _
'          & " WHERE cCtaCod = '" & psCuenta & "' "
'End If
lsSQL = "stp_sel_ObtieneCredPignoDet '" & psCuenta & "'"
'RECO FIN**************************************************
Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneDatosCreditoPignoraticioJoyasDet = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContratoDet>>", Err.Description
    
End Function

Public Function nObtieneDatosCostosRegistroCredPignoraticio(ByVal psCtaCod As String) As ADODB.Recordset

'************************************
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New ADODB.Recordset

On Error GoTo dError

Dim loRegValida As DCOMColPFunciones
Dim lrValida As ADODB.Recordset

'*** PEAC 20161129 - COMENTADO PARA CREAR UN STORE
    'lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, dVigencia, P.nTasaInteres, CP.nNroDuplic, " _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodInteresCompensatorio & "   then nMonto else 0 end )  as nInteres," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodImpuesto & "  then nMonto else 0 end )  as nImpuesto," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodTasacion & "  then nMonto else 0 end )  as nTasacion," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodCustodia & "  then nMonto else 0 end )  as nCustodia " _
         & " FROM Colocaciones c Inner join Producto P on P.cCtaCod = c.cCtaCod " _
         & " Inner Join ColocPignoraticio CP On C.cCtaCod = CP.cCtaCod Inner join ColocCalendario CCal on CCal.cCtaCod = C.cCtaCod " _
         & " Inner join ColoccalendDetPig CCalD on CCal.cCtaCod = CCalD.cCtaCod and CCal.nNroCalen = CCalD.nNroCalen " _
         & " And CCal.nColocCalendApl = CCalD.nColocCalendApl " _
         & " Where P.cCtaCod  = '" & psCtaCod & "' AND CCalD.nColocCalendApl =  " & gColocCalendAplDesembolso & " " _
         & " GROUP BY P.cCtaCod, P.nPrdEstado, dVigencia, P.nTasaInteres, CP.nNroDuplic "
    
    lsSQL = "exec stp_sel_ObtieneDatosCostosRegistroCredPignoraticio '" & psCtaCod & "' "
    
Set lrs = coConex.CargaRecordSet(lsSQL)

Set nObtieneDatosCostosRegistroCredPignoraticio = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
End Function

'*** PEAC 20161206
Public Function nObtieneDatosObtieneCosNotarialPignoraticio(ByVal psCtaCod As String) As ADODB.Recordset

Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New ADODB.Recordset

On Error GoTo dError

Dim loRegValida As DCOMColPFunciones
Dim lrValida As ADODB.Recordset
    
lsSQL = "exec stp_sel_ObtieneCosNotarialPigno '" & psCtaCod & "' "
    
Set lrs = coConex.CargaRecordSet(lsSQL)

Set nObtieneDatosObtieneCosNotarialPignoraticio = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Costos Contrato en <<nObtieneDatosObtieneCosNotarialPignoraticio>>", Err.Description
End Function

Public Function dObtieneDatosCreditoPignoraticioPersonas(ByVal psCuenta As String) As ADODB.Recordset
'Obtiene Datos de Credito Pignoraticio PERSONA
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New ADODB.Recordset

On Error GoTo dError

'lsSQL = "SELECT PP.cPersCod, P.cPersNombre, P.cPersDireccDomicilio, P.cPersTelefono, " _
'    & " P.cPersDireccUbiGeo, " _
'    & " Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),'') , " _
'    & " Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ) , " _
'    & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) , " _
'    & " NroDNI = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdDNI & " ),   " _
'    & " NroRUC = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdRUC & " )   " _
'    & " FROM ProductoPersona PP INNER JOIN Persona P  ON PP.cPersCod = P.cPersCod " _
'    & " WHERE cCtaCod = '" & psCuenta & "' "


'lsSQL = " SELECT PP.cPersCod,replace(substring(cpersnombre,1,case when PATINDEX ( '%,%', cpersnombre )<>0 then PATINDEX ( '%,%', cpersnombre )-1 else PATINDEX ( '%,%', cpersnombre ) end),'/',' ') as cPersApellido,"
'lsSQL = lsSQL & " substring(cpersnombre,PATINDEX ( '%,%', cpersnombre )+1,len(cpersnombre)) as cPersNombre , P.cPersDireccDomicilio, P.cPersTelefono, "
'lsSQL = lsSQL & " P.cPersDireccUbiGeo, "
'lsSQL = lsSQL & " Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),'') , "
'lsSQL = lsSQL & " Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ) , "
'lsSQL = lsSQL & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) , "
'lsSQL = lsSQL & " NroDNI = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdDNI & " ),   "
'lsSQL = lsSQL & " NroRUC = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdRUC & " )   "
'lsSQL = lsSQL & " FROM ProductoPersona PP INNER JOIN Persona P  ON PP.cPersCod = P.cPersCod "
'lsSQL = lsSQL & " WHERE cCtaCod = '" & psCuenta & "' "

'*** PEAC 20161207
'**********RECO 20130913 INC1309130005****************
'lsSQL = " SELECT PP.cPersCod,replace(substring(cpersnombre,1,case when PATINDEX ( '%,%', cpersnombre )<>0 then PATINDEX ( '%,%', cpersnombre )-1 else PATINDEX ( '%,%', cpersnombre ) end),'/',' ') as cPersApellido,"
'lsSQL = lsSQL & " substring(cpersnombre,PATINDEX ( '%,%', cpersnombre )+1,len(cpersnombre)) as cPersNombre , P.cPersDireccDomicilio, P.cPersTelefono, "
'lsSQL = lsSQL & " P.cPersDireccUbiGeo, "
'lsSQL = lsSQL & " Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),'') , "
'lsSQL = lsSQL & " Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ) , "
'lsSQL = lsSQL & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) , "
''lsSQL = lsSQL & " NroDNI = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdDNI & " ),   "
'lsSQL = lsSQL & " NroDNI = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo in  (" & gPersIdDNI & "," & gPersIdExtranjeria & ") ),   "
'lsSQL = lsSQL & " NroRUC = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdRUC & " )   "
'lsSQL = lsSQL & " FROM ProductoPersona PP INNER JOIN Persona P  ON PP.cPersCod = P.cPersCod "
'lsSQL = lsSQL & " WHERE cCtaCod = '" & psCuenta & "' "
'*******************END RECO*****************

lsSQL = "EXEC stp_sel_ObtieneDatosCreditoPignoraticioPersonas '" & psCuenta & "' "
'*** FIN PEAC

Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneDatosCreditoPignoraticioPersonas = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function


Public Function dObtieneCredPigDePersona(ByVal psPersCod As String, _
        ByVal psEstados As String, Optional ByVal psAgencia As String = "") As ADODB.Recordset
'Obtiene Listado Creditos Pig. de una Persona
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Dim lsAgencia As String

If Trim(psAgencia) = "" Then
    lsAgencia = "__"
Else
    lsAgencia = Trim(psAgencia)
    
End If

On Error GoTo dError

'COMENTADO BY ARLO20190304 ERS042 - 2018 BEGIN
'If psEstados = "2810,2812" Then
'
''Modificado Por RIRO 20130723 SEGUN ERS101-2013
'    'MAVM 20100609
'    lsSQL = "SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, " _
'            & "T1.cConsDescripcion cRelacion, UPPER(T2.cConsDescripcion) cProducto, " _
'            & "UPPER(T3.cConsDescripcion) cMoneda " _
'            & "FROM ProductoPersona PP INNER JOIN Producto P  " _
'            & "ON P.cCtaCod = PP.cCtaCod inner join colocaciones c on c.cctacod = p.cctacod INNER JOIN " _
'            & csCentralCom & "Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN " & csCentralCom & "" _
'            & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN " & csCentralCom & "Constante T2 " _
'            & "ON SUBSTRING(PP.cCtaCod,6  ,3) = CONVERT(Varchar(3),T2.nConsValor) INNER JOIN " & csCentralCom & "" _
'            & "Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) " _
'            & "INNER JOIN COLOCPIGSOBRANTE S ON S.cCtaCod = P.cCtaCod " _
'            & "WHERE PP.cPersCod = '" & psPersCod & "' " _
'            & "AND T1.nConsCod = " & gColocRelacPers & " AND T.nConsCod = " & gColocEstado & " AND " _
'            & "T2.nConsCod = " & gProducto & " AND T3.nConsCod = " & gMoneda _
'            & "AND P.cCtaCod like '___" & lsAgencia & "305" & "%' " _
'            & "AND P.nPrdEstado in ( " & psEstados & " ) " _
'            & "AND S.nEstadoSobrante = 0"
'Else
'    lsSQL = "SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, " _
'            & "T1.cConsDescripcion cRelacion, UPPER(T2.cConsDescripcion) cProducto, " _
'            & "UPPER(T3.cConsDescripcion) cMoneda, c.dVenc, c.nMontoCol nMonto " _
'            & "FROM ProductoPersona PP INNER JOIN Producto P  " _
'            & "ON P.cCtaCod = PP.cCtaCod inner join colocaciones c on c.cctacod = p.cctacod INNER JOIN " _
'            & csCentralCom & "Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN " & csCentralCom & "" _
'            & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor LEFT JOIN " & csCentralCom & "Constante T2 " _
'            & "ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2.nConsValor) and T2.nConsCod = " & gProducto & " LEFT JOIN " & csCentralCom & "Constante T2B " _
'            & "ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2B.nConsValor) and T2B.nConsCod = " & gTpoProducto & " INNER JOIN " & csCentralCom & "" _
'            & "Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) " _
'            & "WHERE PP.cPersCod = '" & psPersCod & "' " _
'            & "AND T1.nConsCod = " & gColocRelacPers & " AND T.nConsCod = " & gColocEstado & " " _
'            & "AND T3.nConsCod = " & gMoneda & " " _
'            & "AND (P.cCtaCod like '___" & lsAgencia & "305" & "%' Or P.cCtaCod like '___" & lsAgencia & "709" & "%') " _
'            & "AND P.nPrdEstado in ( " & psEstados & " ) "
'End If
'
'lsSQL = lsSQL & " ORDER BY PP.cCtaCod"
'END

lsSQL = "EXEC stp_sel_ObtieneCredPigDePersona '" & psEstados & "','" & psPersCod & "' " 'ARLO20190304 ERS042-2018

Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneCredPigDePersona = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

Public Function dObtieneDatosCreditoPignoraticioCostos(ByVal psCtaCod As String)
    Dim sql As String
    Dim lrs As ADODB.Recordset
  On Error GoTo dError
    sql = " select nPrdConceptoCod,nMonto from ColocCalendDetPig where cctacod='" & psCtaCod & "'" & _
          " Order BY nPrdConceptoCod"
    
    Set lrs = coConex.CargaRecordSet(sql)
    Set dObtieneDatosCreditoPignoraticioCostos = lrs
    Set lrs = Nothing
Exit Function

dError:
    Err.Raise Err.Number, "No existe Costos", Err.Description
End Function

Public Function VerificarCreditoxRescatar(ByVal psCtaCod As String) As Boolean
    Dim sql As String
    Dim lrs As New ADODB.Recordset
  On Error GoTo dError
    
    '*** PEAC 20101220
    'sql = "select count(*)NroReg from ColocacEstado where cctacod='" & psCtaCod & "' and nPrdEstado =" & gColPEstCance & " "
    sql = " exec stp_sel_VerificarCreditoxRescatar '" & psCtaCod & "'"
    '*** FIN PEAC
    
    Set lrs = coConex.CargaRecordSet(sql)
    If Not (lrs.EOF And lrs.BOF) Then
        If lrs!NroReg = 0 Then
            VerificarCreditoxRescatar = False
        Else
            VerificarCreditoxRescatar = True
        End If
    End If
    Set lrs = Nothing
    Exit Function
dError:
    Err.Raise Err.Number, "Error en Datos", Err.Description
End Function

Public Function dObtieneDatosCreditoPignoraticioTasas(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sql As String
    Dim lrs As ADODB.Recordset
  On Error GoTo dError
    sql = " select nPrdTasaInteres,nTasaInteres from ProductoTasaInteres where cctacod='" & psCtaCod & "'" & _
          " Order BY nPrdTasaInteres"
    
    Set lrs = coConex.CargaRecordSet(sql)
    Set dObtieneDatosCreditoPignoraticioTasas = lrs
    Set lrs = Nothing
Exit Function

dError:
    Err.Raise Err.Number, "No existe Tasas de Interes", Err.Description
End Function

'*** PEAC 20080715
Public Function RecuperaCredExcluirAdj(ByVal psAge As String, ByVal psNumProceso As String) As ADODB.Recordset
Dim sSql As String

sSql = "exec stp_sel_ListadoParaExcluirAdjudica '" & psAge & "','" & psNumProceso & "'"

Set RecuperaCredExcluirAdj = coConex.CargaRecordSet(sSql)

End Function

'MAVM 20100717
Public Function dObtieneDatosRegJoyasDet(ByVal psCodJoyas As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    Set lrs = New ADODB.Recordset
On Error GoTo dError
    lsSQL = "Select cKilataje, nItem, nPiezas, nPesoBruto, nPesoNeto, nTasacion, cDescripcion " _
              & " FROM JoyaGarantiaDet  " _
              & " WHERE cJoyGarCod = '" & psCodJoyas & "' "
    Set lrs = coConex.CargaRecordSet(lsSQL)
    Set dObtieneDatosRegJoyasDet = lrs
    Set lrs = Nothing
    Exit Function
dError:
Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContratoDet>>", Err.Description
End Function

Public Function GetDatosRegJoyasCodPer(ByVal cPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim lrs As ADODB.Recordset
    Set lrs = New ADODB.Recordset

    sSql = "Select cJoyGarCod cPersNombre, cJoyGarCod cPersCod From JoyaGarantia " _
           & "where cJoyGarCod Not In (Select cNroDoc From Garantias G Where G.cNroDoc = cJoyGarCod and nTpoGarantia = 7) " _
           & "and cPersCod = '" & cPersCod & "'"

    Set lrs = coConex.CargaRecordSet(sSql)
    Set GetDatosRegJoyasCodPer = lrs
    Set lrs = Nothing
    Exit Function
End Function

Public Function GetDatosRegJoyas(ByVal sJoyasCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim lrs As ADODB.Recordset
    Set lrs = New ADODB.Recordset
    
    sSql = "Select cJoyGarCod, cPersCod, nPiezas, nTasacion, nPesoBruto, nPesoNeto, cMovNro " _
        & "From JoyaGarantia " _
        & "Where cJoyGarCod = '" & sJoyasCod & "'"
    
    Set lrs = coConex.CargaRecordSet(sSql)
    Set GetDatosRegJoyas = lrs
    Set lrs = Nothing
End Function

'MAVM 20100723
Public Function dObtieneDatosCredJoyasPersonas(ByVal iTipo As Integer, ByVal psCuenta As String, pscJoyGarCod As String) As ADODB.Recordset
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New ADODB.Recordset
On Error GoTo dError

If iTipo = 1 Then
    lsSQL = "SELECT PP.cPersCod,replace(substring(cpersnombre,1,case when PATINDEX ( '%,%', cpersnombre )<>0 then PATINDEX ( '%,%', cpersnombre )-1 else PATINDEX ( '%,%', cpersnombre ) end),'/',' ') as cPersApellido,"
    lsSQL = lsSQL & " substring(cpersnombre,PATINDEX ( '%,%', cpersnombre )+1,len(cpersnombre)) as cPersNombre , P.cPersDireccDomicilio, P.cPersTelefono, "
    lsSQL = lsSQL & " P.cPersDireccUbiGeo, "
    lsSQL = lsSQL & " Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),'') , "
    lsSQL = lsSQL & " Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ) , "
    lsSQL = lsSQL & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) , "
    lsSQL = lsSQL & " NroDNI = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdDNI & " ),   "
    lsSQL = lsSQL & " NroRUC = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdRUC & " )   "
    lsSQL = lsSQL & " FROM ProductoPersona PP INNER JOIN Persona P  ON PP.cPersCod = P.cPersCod "
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCuenta & "' and PP.nPrdPersRelac = 20"
Else
    lsSQL = "Select P.cPersCod, replace(substring(cpersnombre,1,case when PATINDEX ( '%,%', cpersnombre )<>0 then PATINDEX ( '%,%', cpersnombre )-1 else PATINDEX ( '%,%', cpersnombre ) end),'/',' ') as cPersApellido, "
    lsSQL = lsSQL & " substring(cpersnombre,PATINDEX ( '%,%', cpersnombre )+1,len(cpersnombre)) as cPersNombre , P.cPersDireccDomicilio, P.cPersTelefono,  P.cPersDireccUbiGeo,  Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),'') ,  Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ), "
    lsSQL = lsSQL & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) ,  NroDNI = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo =  1 ),    NroRUC = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo =  2 )  "
    lsSQL = lsSQL & "  From JoyaGarantia JG Inner Join Persona P On JG.cPersCod = P.cPersCod"
    lsSQL = lsSQL & " WHERE JG.cJoyGarCod = '" & pscJoyGarCod & "'"
End If

Set lrs = coConex.CargaRecordSet(lsSQL)
Set dObtieneDatosCredJoyasPersonas = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
End Function

Public Function dObtieneDatosCreditoGarantJoyasDet(ByVal iTipo As Integer, ByVal psCuenta As String, pscJoyGarCod As String) As ADODB.Recordset
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Set lrs = New ADODB.Recordset
If iTipo = 1 Then
lsSQL = "SELECT JD.nItem, JD.nPiezas, JD.nPesoBruto, JD.nPesoNeto, JD.cDescripcion " _
        & "From Producto P INNER JOIN Colocaciones C ON P.cCtaCod = c.cCtaCod and P.nPrdEstado = 2050 " _
        & "Inner Join ProductoPersona PP on C.cCtaCod = PP.cCtaCod " _
        & "Inner Join PersGarantia PG on PP.cPersCod = PG.cPersCod and PG.nRelacion = 1 " _
        & "Inner Join Garantias G on PG.cNumGarant = G.cNumGarant and cTpoDoc = 145 " _
        & "Inner Join JoyaGarantia J on G.cNroDoc = J.cJoyGarCod " _
        & "Inner join JoyaGarantiaDet JD on J.cJoyGarCod = JD.cJoyGarCod " _
        & " WHERE P.cCtaCod = '" & psCuenta & "' "
Else
lsSQL = "SELECT JD.nItem, JD.nPiezas, JD.nPesoBruto, JD.nPesoNeto, JD.cDescripcion " _
        & "From JoyaGarantia J  Inner join JoyaGarantiaDet JD on J.cJoyGarCod = JD.cJoyGarCod " _
        & "WHERE J.cJoyGarCod = '" & pscJoyGarCod & "' And J.nJoyGarEstado in (0, 2)"
End If
Set lrs = coConex.CargaRecordSet(lsSQL)
Set dObtieneDatosCreditoGarantJoyasDet = lrs
Set lrs = Nothing
End Function

Public Function dObtieneListaCredJoyasPersona(ByVal psPersCod As String, ByVal iTipo As Integer) As ADODB.Recordset
Dim lrs As ADODB.Recordset
Dim lsSQL As String

On Error GoTo dError

    lsSQL = "Select IsNull(CG.cCtaCod, '')cCtaCod, JG.cJoyGarCod, T.cConsDescripcion " _
            & "From JoyaGarantia JG Inner Join Constante T on JG.nJoyGarEstado = T.nConsValor AND T.nConsCod = 3038 " _
            & "Left Join Garantias G on G.cNroDoc = JG.cJoyGarCod And G.cTpoDoc = 145 " _
            & "Left join ColocGarantia CG on CG.cNumGarant = G.cNumGarant And CG.nEstado = 3 " _
            & "Left Join Producto P on CG.cCtaCod = P.cCtaCod And P.nPrdEstado = 2050 " _
            & "WHERE cPersCod = '" & psPersCod & "' "
    If iTipo = 1 Then
        lsSQL = lsSQL & " And JG.nJoyGarEstado in( '2') "
    Else
        lsSQL = lsSQL & " And JG.nJoyGarEstado in( '0') "
    End If
            
    lsSQL = lsSQL & " ORDER BY P.cCtaCod"

Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneListaCredJoyasPersona = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
    
End Function
'***************************RECO20131120 ERS158-2013*******************
Public Function dPigObtenerValoresTasacion() As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_PigObtenerValoresTasacion"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set dPigObtenerValoresTasacion = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
        
End Function

Public Sub dPigAcutalizaValorTasacion(ByVal nTasacionCod As Integer, ByVal n14kt As Integer, ByVal n16kt As Integer _
                                    , ByVal n18kt As Integer, ByVal n21kt As Integer, ByVal cultimaActualizacion As String)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_upd_PigAcutalizaValorTasacion " & nTasacionCod & "," & n14kt & "," & n16kt & "," & _
                n18kt & "," & n21kt & ",'" & cultimaActualizacion & "'"
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
End Sub

Public Function dVerificarCredPignoAdjudicado(ByVal psPersCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_VerificarCredPignoAdjudicado '" & psPersCod & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set dVerificarCredPignoAdjudicado = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
        
End Function

'*** PEAC 20161216
Public Function dVerificarCredPignoTipoCliente(ByVal psPersCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_VerificarCredPignoTipoCliente '" & psPersCod & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set dVerificarCredPignoTipoCliente = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Obtiene Tipo de cliente de Pigno", Err.Description
        
End Function


Public Function dVerificarCredPignoDesembolso(ByVal psPersCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_VerificarCredPignoDesembolso '" & psPersCod & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set dVerificarCredPignoDesembolso = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
End Function

Public Function PigObtenerValorTasacionxTpoClienteKt(ByVal nTpoCliente As Integer) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_PigObtenerValorTasacionxTpoClienteKt '" & nTpoCliente & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set PigObtenerValorTasacionxTpoClienteKt = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
End Function
'*****************************END RECO*********************************
'RECO20140116 INC1401130010**************************************************************************
Public Function ObtieneDatosPersonaXCredito(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_ObtieneDatosPersonaXCredito '" & psCtaCod & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set ObtieneDatosPersonaXCredito = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Obtiene Detos Cliente por Nº Credito", Err.Description
End Function
'RECO END********************************************************************************
'RECO20140128 ERS002 RECO-N**************************************************************************
Public Function ObtieneValorKilotajePesoCredPigno(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC STP_SEL_ObtieneValorKilotajePesoCredPigno '" & psCtaCod & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set ObtieneValorKilotajePesoCredPigno = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Obtiene Detos Cliente por Nº Credito", Err.Description
End Function

Public Sub InsertaRegistroCreditoPignoAmpliado(ByVal psCtaCod As String, ByVal psCtaCodAmp As String, ByVal pnEstado As Integer, ByVal psEstadoFec As String, ByVal nMonto As Double, ByVal nMontoAmp As Double, ByVal psUser As String)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_ins_InsertaRegistroCreditoPignoAmpliado '" & psCtaCod & "','" & psCtaCodAmp & "'," & pnEstado & ",'" & psEstadoFec & "'," & nMonto & "," & nMontoAmp & ",'" & psUser & "'"
                
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
End Sub

Public Function ListaCredPignoEstado(ByVal psEstadoCad As String, ByVal psAgeCod As String, ByVal pnTpoBusc As Integer, ByVal psPersCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_ListaCredPignoEstado '" & psEstadoCad & "','" & psAgeCod & "'," & pnTpoBusc & ",'" & psPersCod & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set ListaCredPignoEstado = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Lista los Creditos Pignoraticios según su estado", Err.Description
End Function

Public Function ObtieneCuentaCancPignoAmpliado(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_ObtieneCuentaCancPignoAmpliado '" & psCtaCod & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set ObtieneCuentaCancPignoAmpliado = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Lista los Creditos Pignoraticios según su estado", Err.Description
End Function

Public Function ObtieneEstadoCredPignoAmp(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_ObtieneEstadoCredPignoAmp '" & psCtaCod & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set ObtieneEstadoCredPignoAmp = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Lista los Creditos Pignoraticios según su estado", Err.Description
End Function

'RECO FIN********************************************************************************
'RECO20140311 ERS160-2013****************************************************************
Public Function ObtieneMontoDesembCreDPigMes(ByVal psCtaCod As String, ByVal psPeriodo As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_ObtieneMontoDesembCreDPigMes '" & psCtaCod & "','" & psPeriodo & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set ObtieneMontoDesembCreDPigMes = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Lista los Creditos Pignoraticios según su estado", Err.Description
End Function
'RECO FIN********************************************************************************
'RECO20140326****************************************************************************
Public Function RecuperaEstadoCredPignoAnterior(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_RecuperaEstadoCredPignoAnterior '" & psCtaCod & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set RecuperaEstadoCredPignoAnterior = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Lista los Creditos Pignoraticios según su estado", Err.Description
End Function

'RECO FIN********************************************************************************
'RECO20140707 ERS074-2014*****************************************************************
'Modificación: TORE20180417 ERS054-2017
Public Function ListaCredPrepRetasacion(ByVal psAgeCad As String, ByVal psFecDesde As String _
                                        , ByVal psFecHasta As String, ByVal pnEstadoJoy As Integer _
                                        , ByVal pnTrimestre As Integer, ByVal psAnio As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_ObtieneCredPigPreparaRetasacion '" & psAgeCad & "','" & psFecDesde & "','" & psFecHasta & "'," & pnEstadoJoy & "," & pnTrimestre & ",'" & psAnio & "'"
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ListaCredPrepRetasacion = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
'TORE - ERS054-2017
Public Function ObtenerObservacionRetasacion(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "exec stp_sel_ObtieneCredRetasacion '" & psCtaCod & "'"
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ObtenerObservacionRetasacion = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener las observaciones de la retasacion de joyas", Err.Description
End Function
'Modificación: TORE20180417 ERS054-2017
Public Function DevuelveCtaCodCredPRetasacion(ByVal psAgeCod As String, ByVal psTop As String, ByVal psEstado As String _
                                                , ByVal psFecDesde As String, ByVal psFecHasta As String _
                                                , ByVal pnTrimestre As Integer, ByVal psAnio As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_DevuelveCtaCodCredPRetasacion '" & psAgeCod & "','" & psTop & "','" & psEstado & "','" & psFecDesde & "','" & psFecHasta & "'," & pnTrimestre & ",'" & psAnio & "'"
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set DevuelveCtaCodCredPRetasacion = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
'Modificación: TORE20180417 ERS054-2017
Public Function DevuelveDatosListaCredPrepRetasacion(ByVal psCtaCad As String, ByVal pnEstado As Integer, ByVal pnCodigoID As Integer) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_DevuelveDatosListaCredPrepRetasacion '" & psCtaCad & "'," & pnEstado & "," & pnCodigoID
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set DevuelveDatosListaCredPrepRetasacion = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
'TORE ERS054-2017
Public Function RegistraPreparacionRetasacion(ByVal psAgeCod As String, ByVal pnTpoPrepara As Integer, _
                                                ByVal pnNumCred As Integer, ByVal pdFecha As String, _
                                                ByVal psUser As String, ByVal psCodRetasacion As String, _
                                                ByVal pnTotLotes As Integer, ByVal pnCriterio As Integer, _
                                                ByVal pdFechaIni As String, ByVal pdFechaFin As String, _
                                                ByVal pnTrimestre As Integer, ByVal psAnio As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        'lsSQL = "stp_ins_RegistraPreparacionRetasacion '" & psAgeCod & "'," & pnTpoPrepara & "," & pnNumCred & ",'" & Format(pdFecha, "yyyy/MM/dd") & "','" & psUser & "','" & psCodRetasacion & "'"
        '[TORE RFC1811260001: ADD pnTotLotes]
        lsSQL = "stp_ins_RegistraPreparacionRetasacion '" & psAgeCod & "'," & _
                                                        pnTpoPrepara & "," & _
                                                        pnNumCred & ",'" & _
                                                        Format(pdFecha, "yyyy/MM/dd") & "','" & _
                                                        psUser & "','" & psCodRetasacion & "', " & _
                                                        pnTotLotes & ", " & pnCriterio & ", '" & _
                                                        pdFechaIni & "', '" & pdFechaFin & "', " & pnTrimestre & ", '" & _
                                                        psAnio & "'"
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set RegistraPreparacionRetasacion = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al registrar los datos", Err.Description
End Function
'TORE ERS054-2017
Public Sub RegistraPreparacionRetasacionDet(ByVal pnCodigoID As Integer, ByVal psCtaCod As String, ByVal pnNroRetasacion As String)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_ins_RegistraPreparacionRetasacionDet " & pnCodigoID & ",'" & psCtaCod & "','" & pnNroRetasacion & "'"
        coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
End Sub
'TORE - ERS054-2017
Public Function CargaListaRetasaciones(ByVal pnTpoBusc As Integer, ByVal psNumBusc As String, _
                                        ByVal psDel As String, ByVal psHasta As String, _
                                        ByVal psCtaCad As String, ByVal pnTrimestre As Integer, _
                                        ByVal psAnio As String, ByVal pnEstado As Integer) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_CargaListaRetasaciones " & pnTpoBusc & ",'" & psNumBusc & "','" & psDel _
                & "','" & psHasta & "','" & psCtaCad & "'," & pnTrimestre & ", '" & psAnio & "', " & pnEstado
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set CargaListaRetasaciones = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al registrar los datos", Err.Description
End Function
'TORE ERS054-2017
Public Function ObtenerDetRetasacion(ByVal psCtaCod As String, ByVal pnCodigoID As Integer, ByVal pnItem As Integer) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "stp_sel_ObtenerDatosRetasacion '" & psCtaCod & "'," & pnCodigoID & "," & pnItem
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ObtenerDetRetasacion = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al registrar los datos", Err.Description
End Function
'TORE ERS054-2017
Public Function ObtieneMiembrosRetasacion(ByVal pnCodigoID As Integer) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "stp_sel_ObtenerMiembrosRetasacion " & pnCodigoID
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ObtieneMiembrosRetasacion = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
'TORE ERS054-2017
Public Function ObtieneCredPignoRetasacion(ByVal psAgeCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        'lsSQL = "stp_sel_ObtieneCredPignoRetasacion '" & psAgeCod & "'," & pnTpoBusq
        lsSQL = "stp_sel_ObtenerCredPignoraticioCtaCod '" & psAgeCod & "'"
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ObtieneCredPignoRetasacion = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al registrar los datos", Err.Description
End Function
'TORE ERS054-2017
Public Function ObtieneCodPrepaRetasacion(ByVal psNombreArchivo As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_ObtenerCodPrepRetasacion '" & psNombreArchivo & "'"
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ObtieneCodPrepaRetasacion = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al registrar los datos", Err.Description
End Function

'TORE ERS054-2017
'Public Sub RegistraRetasacion(ByVal pnCodigoID As Integer, ByVal psCtaCod As String, ByVal pnItem As Integer, ByVal psKilataje As String, _
'                              ByVal pnPiezas As Integer, ByVal pnPesoBruto As Double, ByVal pnPesoNeto As Double, ByVal pnValTasac As Double, _
'                              ByVal psNroHolograma As String, ByVal psObservaciones As String)
'    Dim lsSQL As String
'    On Error GoTo dError
'        lsSQL = "exec stp_ins_RegistraRetasacion " & pnCodigoID & ",'" & psCtaCod & "'," & pnItem & ",'" & psKilataje & "'," _
'                              & pnPiezas & "," & pnPesoBruto & "," & pnPesoNeto & "," & pnValTasac & ",'" & psNroHolograma & "','" & psObservaciones & "'"
'
'    coConex.CargaRecordSet (lsSQL)
'    Exit Sub
'dError:
'        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
'End Sub
'TORE - Correccion
Public Function RegistraRetasacion(ByVal pnCodigoID As Long, ByVal psCtaCod As String, ByVal pnItem As Integer, ByVal psKilataje As String, _
                              ByVal pnPiezas As Integer, ByVal pnPesoBruto As Double, ByVal pnPesoNeto As Double, ByVal pnValTasac As Double, _
                              ByVal psNroHolograma As String, ByVal psObservaciones As String, ByVal pdFechaRetas As Date) As Integer
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo dError
    lsSQL = "exec stp_ins_RegistraRetasacion " & pnCodigoID & ",'" & psCtaCod & "'," & pnItem & ",'" & psKilataje & "'," _
                              & pnPiezas & "," & pnPesoBruto & "," & pnPesoNeto & "," & pnValTasac & ",'" & psNroHolograma & "','" & psObservaciones & "','" & Format(pdFechaRetas, "yyyyMMdd") & "'"
                
    Set lrs = coConex.CargaRecordSet(lsSQL)
    If Not (lrs.BOF And lrs.EOF) Then
       RegistraRetasacion = lrs!Exito
    Else
       RegistraRetasacion = 0 'El registro de insersion o actualizacion no fue procesado.
    End If
    Set lrs = Nothing
    Exit Function
dError:
        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
End Function



'TORE ERS054-2017
Public Sub RegistraMiembrosRetasacion(ByVal pnCodigoID As Integer, ByVal pnOrden As Integer, _
                                      ByVal psNombreMiembro As String, ByVal psRolMiembro As String)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "exec stp_ins_RegistraMiembrosRetasacion " & pnCodigoID & "," & pnOrden & ",'" & psNombreMiembro & "','" & psRolMiembro & "'"
                
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Registra los miembros de la retasacion", Err.Description
End Sub
'TORE ERS054-2017
Public Sub RegistraFechaRetasacion(ByVal pnCodigoID As Integer, ByVal psUserRetas As String)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "exec stp_upd_InsertarFechaRetasacion " & pnCodigoID & ",'" & psUserRetas & "'"
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
End Sub
'[TORE ADD 01062019]
Public Function ObtenerCodRetasacion(ByVal pnCodigoID As Integer, ByVal psCredito As String) As ADODB.Recordset
   Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "stp_sel_ObtenerCodRetasacion " & pnCodigoID & ", '" & psCredito & "'"
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ObtenerCodRetasacion = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
'TORE ERS054-2017
Public Function ObtieneCtaCodRetasacionPrep(ByVal pnCodigoID As Integer) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_ObtieneCtaCodRetasacionPrep " & pnCodigoID
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ObtieneCtaCodRetasacionPrep = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
Public Sub FinalizaRetasacionPreparada(ByVal pnCodigoID As Integer, ByVal psFecRetas As String, ByVal psUserRetas As String)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_upd_FinalizaRetasacionPreparada " & pnCodigoID & ",'" & psFecRetas & "','" & psUserRetas & "'"
                
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Al finalizar retasacion", Err.Description
End Sub
Public Function DevuelveHistorialRetasacion(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_DevuelveHistorialRetasacion '" & psCtaCod & "'"
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set DevuelveHistorialRetasacion = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
Public Function ObtieneHistorialCredRetasacion(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_ObtieneCredRetasacion '" & psCtaCod & "'"
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ObtieneHistorialCredRetasacion = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
'RECO FIN*********************************************************************************

'RECO20140827 ERS032-2014*****************************************************************
Public Sub RegistroConfigTpoPrenda(ByVal psTipoDesc As String, ByVal pnPiezas As Integer)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_ins_RegistroConfigTpoPrenda '" & psTipoDesc & "'," & pnPiezas
                
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Registra Tipo", Err.Description
End Sub
Public Sub RegistroConfigSubTpoPrenda(ByVal pnTipoCod As Integer, ByVal psSubTipoDesc As String)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_ins_RegistraConfigSupTpoPrenda " & pnTipoCod & ",'" & psSubTipoDesc & "'"
                
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Registra SubTipo", Err.Description
End Sub
Public Function ObtieneTipoPrenda() As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_ObtieneTipoPrenda "
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ObtieneTipoPrenda = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description

End Function

Public Sub ActualizaConfigTpoPrenda(ByVal pnTipoCod As Integer, ByVal psTipoDesc As String, ByVal pnPiezas As Integer)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_upd_ActualizaConfigTipoPrenda " & pnTipoCod & ",'" & psTipoDesc & "'," & pnPiezas
                
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Actualiza Tipo", Err.Description
End Sub

Public Sub QuitaConfigTipoPrenda(ByVal pnTipoCod As Integer)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_upd_QuitaConfigTipoPrenda " & pnTipoCod
                
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Quita tipo", Err.Description
End Sub

Public Function ObtieneSubTipoPrendaXTipo(ByVal pnTipoCod As Integer) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_ObtieneConfigSubTipoPrendaXTipo " & pnTipoCod
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ObtieneSubTipoPrendaXTipo = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description

End Function

Public Sub ActualizaConfigSubTpoPrenda(ByVal pnSubTipoCod As Integer, ByVal psSubTipoDesc As String)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_upd_ActualizaConfigSubTipoPrenda " & pnSubTipoCod & ",'" & psSubTipoDesc & "'"
                
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Actualiza SubTipo", Err.Description
End Sub

Public Sub QuitaConfigSubTipoPrenda(ByVal pnSubTipoCod As Integer)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_upd_QuitaConfigSubTipoPrenda " & pnSubTipoCod
                
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Quita SubTipo", Err.Description
End Sub
Public Function TasaXCalificSBS() As ADODB.Recordset
 Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_TasaCalificSBS"
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set TasaXCalificSBS = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
Public Sub ActualizaConfiguracionTasaXCalificSBS(ByVal pnCodigo As Integer, ByVal pnTasa As Integer)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_upd_ActualizaConfiguracionTasaXCalificSBS " & pnCodigo & "," & pnTasa
                
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Actualizar tasa por calificación SBS", Err.Description
End Sub
Public Function ListaCriteriosEvaluacionClientePignoraticio(ByVal pnCalificacion As Integer)
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_ListaCriteriosEvaluacionClientePignoraticio " & pnCalificacion
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ListaCriteriosEvaluacionClientePignoraticio = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
Public Sub ActualizaValoresCalificClientePigno(ByVal pnCalificacion As Integer, ByVal pnCriterio As Integer, ByVal pnValor As Integer, ByVal pnValorPrestTas As Integer)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_upd_ActualizaValoresCalificClientePigno " & pnCalificacion & "," & pnCriterio & "," & pnValor & "," & pnValorPrestTas
                
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Actualizar calificacion de clientes pignoraticio", Err.Description
End Sub
Public Function ListaTipoPrendaPigno() As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_ListaTipoPrendaPigno"
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ListaTipoPrendaPigno = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
Public Function ListaSubTipoPrendaPigno(ByVal pnTipoCod As Integer) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_ListaSubTipoPrendaPigno " & pnTipoCod
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ListaSubTipoPrendaPigno = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
Public Function ObtieneValorCriterioCalificSBSPigno(ByVal psPersCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_ObtieneValorCriterioCalificSBSPigno '" & psPersCod & "'"
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ObtieneValorCriterioCalificSBSPigno = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
Public Function ValorCriterioInactividadPigno(ByVal psPersCod As String, ByVal pnInactividad As Integer, ByVal psFecSis As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_ValorCriterioInactividadPigno '" & psPersCod & "'," & pnInactividad & ",'" & psFecSis & "'"
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ValorCriterioInactividadPigno = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
Public Function ObtieneDatosCalificClientePignoraticio() As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_ObtieneDatosCalificClientePignoraticio "
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ObtieneDatosCalificClientePignoraticio = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
Public Function ObtieneTasaPignoCalificacionSBS(ByVal nCodigo As Integer) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
        lsSQL = "stp_sel_ObtieneTasaPignoCalificacionSBS " & nCodigo
        Set lrs = coConex.CargaRecordSet(lsSQL)
        Set ObtieneTasaPignoCalificacionSBS = lrs
        Set lrs = Nothing
        Exit Function
dError:
        Err.Raise Err.Number, "Error al obtener los datos", Err.Description
End Function
Public Function DevulveDetalleJoya(ByVal psCuenta As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    Set lrs = New ADODB.Recordset
    
    On Error GoTo dError
    lsSQL = "stp_sel_DevulveDetalleJoya '" & psCuenta & "'"
    Set lrs = coConex.CargaRecordSet(lsSQL)
    Set DevulveDetalleJoya = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Obtiene detalle de joyas en <<DevulveDetalleJoya>>", Err.Description
End Function
'RECO FIN*********************************************************************************
'RECO20141201 ERS163-2014*******************************
Public Function ObtieneValoresAdjudicacion(ByVal psCuenta As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    Set lrs = New ADODB.Recordset
    
    On Error GoTo dError
    lsSQL = "stp_sel_ObtieneValoresAdjudicacion '" & psCuenta & "'"
    Set lrs = coConex.CargaRecordSet(lsSQL)
    Set ObtieneValoresAdjudicacion = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Obtiene detalle de adjudicación <<ObtieneValoresAdjudicacion>>", Err.Description
End Function
Public Sub InsertaRegistroCreditoCampAdj(ByVal pnMovNro As Long, ByVal psCtaCodAdj As String, ByVal psCtaCod As String, ByVal pnMontoInicial As Double)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_ins_RegistraCredCampAdj " & pnMovNro & ",'" & psCtaCodAdj & "','" & psCtaCod & "'," & pnMontoInicial
                
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Registra Crédito Adjudicado", Err.Description
End Sub
Public Function DevuelveValorDesemCampAdjudicado(ByVal psCuenta As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    Set lrs = New ADODB.Recordset
    
    On Error GoTo dError
    lsSQL = "stp_sel_DevuelveValorDesemCampAdjudicado '" & psCuenta & "'"
    Set lrs = coConex.CargaRecordSet(lsSQL)
    Set DevuelveValorDesemCampAdjudicado = lrs
    Set lrs = Nothing
    Exit Function
dError:
        Err.Raise Err.Number, "Obtiene detalle de desembolso <<DevuelveValorDesemCampAdjudicado>>", Err.Description
End Function
Public Function ObtieneValorPesoNetoJoya(ByVal psCuenta As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    Set lrs = New ADODB.Recordset
    
    On Error GoTo dError
    lsSQL = "stp_sel_ObtieneValorPesoNetoJoya '" & psCuenta & "'"
    Set lrs = coConex.CargaRecordSet(lsSQL)
    Set ObtieneValorPesoNetoJoya = lrs
    Set lrs = Nothing
    Exit Function
dError:
        Err.Raise Err.Number, "Obtiene detalle de peso de oro<<ObtieneValorPesoNetoJoya>>", Err.Description
End Function
'RECO FIN***********************************************
'RECO20160225 ERS040-2015*********************************************
Public Function PignoDevuelveNroComprobanteAdj(ByVal psAgeCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    Set lrs = New ADODB.Recordset
    
    On Error GoTo dError
    lsSQL = "stp_sel_PignoDevuelveNroComprobanteAdj'" & psAgeCod & "'"
    Set lrs = coConex.CargaRecordSet(lsSQL)
    Set PignoDevuelveNroComprobanteAdj = lrs
    Set lrs = Nothing
    Exit Function
dError:
        Err.Raise Err.Number, "Devulve número comprobante<<PignoDevuelveNroComprobanteAdj>>", Err.Description
End Function
Public Sub PignoRegistraNroCompAdj(ByVal psComprobanteNro As String, ByVal psNumero As String, ByVal psCtaCod As String, ByVal pnMovNro As Long)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_ins_PignoRegistraNroCompAdj '" & psComprobanteNro & "','" & psNumero & "','" & psCtaCod & "'," & pnMovNro
                
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Resgitra Nro Comp Adj.", Err.Description
End Sub
Public Function PignoReimpresionCompAdj(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    Set lrs = New ADODB.Recordset
    
    On Error GoTo dError
    lsSQL = "stp_sel_PignoReimpresionCompAdj'" & psCtaCod & "'"
    Set lrs = coConex.CargaRecordSet(lsSQL)
    Set PignoReimpresionCompAdj = lrs
    Set lrs = Nothing
    Exit Function
dError:
        Err.Raise Err.Number, "Devulve datos comprobante reimpresion<<PignoDevuelveNroComprobanteAdj>>", Err.Description
End Function
'RECO FIN*************************************************************
'NAGL ERS012-2017 20170710****************************************************
Public Function PignoDescripcionJoyaAdj(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    Set lrs = New ADODB.Recordset
    On Error GoTo dError
    lsSQL = "stp_sel_PignoDescripcionJoyaAdj'" & psCtaCod & "'"
    Set lrs = coConex.CargaRecordSet(lsSQL)
    Set PignoDescripcionJoyaAdj = lrs
    Set lrs = Nothing
    Exit Function
dError:
        Err.Raise Err.Number, "Adquiere Descripción Joyas.", Err.Description
End Function

Public Sub RegistroVentaContratoPigno(ByVal psOpeTpo As String, ByVal pnDocTpo As Long, ByVal psDocNro As String, ByVal pdDocFecha As Date, ByVal psPersCod As String, _
                                      ByVal psCtaCod As String, ByVal psDescrip As String, ByVal pnVVenta As Currency, ByVal pnIGV As Currency, ByVal pnPVentas As Currency, ByVal pnMovNro As Long, ByVal pnMoneda As Integer, _
                                      ByVal psDocNroRefe As String, ByVal pnDocTpoRefe As Integer, ByVal pnTipoCambio As Currency, ByVal psAreaAgecod As String)
 Dim lsSQL As String
 On Error GoTo dError
    lsSQL = "EXEC stp_ins_PignoRegistroVenta '" & psOpeTpo & "', " & pnDocTpo & ", '" & psDocNro & "', '" & Format(pdDocFecha & " " & coConex.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "', '" & psPersCod & "', '" & psCtaCod & "', '" & psDescrip & "', " & pnVVenta & ", " & pnIGV & ", " & pnPVentas & ", '" & psDocNroRefe & "', " & pnDocTpoRefe & ", " & pnMovNro & ", " & pnTipoCambio & "," & pnMoneda & ", '" & psAreaAgecod & "'"
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Registro de Venta Nro Comp Adj.", Err.Description
End Sub
'NAGL ERS 012-2017 **************************************************************

'JOEP ERS047 20170904
Public Function VerificaLimiteZonaGeo(ByVal pncCtacod As String, ByVal pnMonto As Currency, ByVal pnMoneda As Integer) As Boolean
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
    
On Error GoTo dError
    
sSql = "exec stp_sel_ERS047_VerificaSuperaUmbralZonaGeog '" & pncCtacod & "'," & pnMonto & "," & pnMoneda

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set rs = oConecta.CargaRecordSet(sSql)
If Not (rs.EOF And rs.BOF) Then
    If rs!Resultado <> "" Then
       VerificaLimiteZonaGeo = True
    End If
End If
oConecta.CierraConexion
Set oConecta = Nothing

Exit Function

dError:
        Err.Raise Err.Number, "Hubo un Error Comunicarse con TI-Desarrollo.", Err.Description

End Function

Public Sub RegistroLimiteZonaGeog(ByVal pcCtaCod As String, ByVal pnMonto As Currency)
 Dim lsSQL As String
 On Error GoTo dError
    lsSQL = "EXEC stp_ins_ERS047_SolicitudAutorizacionZonaGeog '" & pcCtaCod & "', " & pnMonto & ""
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Hubo un Error Comunicarse con TI-Desarrollo.", Err.Description
End Sub

Public Function ValidaLimiteZonaGeog(ByVal pcCtaCod As String, ByVal pnTpProducto As String) As ADODB.Recordset
 Dim lsSQL As String
 Dim rs As ADODB.Recordset
 On Error GoTo dError
    lsSQL = "EXEC stp_sel_ERS047_VerificaSolicitudAutorizacionZonaGeog '" & pcCtaCod & "'," & pnTpProducto & ""
    Set rs = coConex.CargaRecordSet(lsSQL)
    Set ValidaLimiteZonaGeog = rs
    Set rs = Nothing
    Exit Function
dError:
        Err.Raise Err.Number, "Hubo un Error Comunicarse con TI-Desarrollo.", Err.Description
End Function

'Limite Producto
Public Function VerificaLimiteTpProducto(ByVal pncCtacod As String, ByVal pnMonto As Currency, ByVal pnMoneda As Integer) As Boolean
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
    
On Error GoTo dError
    
sSql = "exec stp_sel_ERS047_VerificaSuperaUmbralTpProducto '" & pncCtacod & "'," & pnMonto & "," & pnMoneda

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set rs = oConecta.CargaRecordSet(sSql)
If Not (rs.EOF And rs.BOF) Then
    If rs!Resultado <> "" Then
       VerificaLimiteTpProducto = True
    End If
End If
oConecta.CierraConexion
Set oConecta = Nothing

Exit Function

dError:
        Err.Raise Err.Number, "Hubo un Error Comunicarse con TI-Desarrollo.", Err.Description

End Function

Public Sub RegistroLimiteTpProducto(ByVal pcCtaCod As String, ByVal pnTpProducto As Integer, ByVal pnMonto As Currency)
 Dim lsSQL As String
 On Error GoTo dError
    lsSQL = "EXEC stp_ins_ERS047_SolicitudAutorizacionTpProducto '" & pcCtaCod & "', " & pnTpProducto & "," & pnMonto & ""
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Hubo un Error Comunicarse con TI-Desarrollo.", Err.Description
End Sub

Public Sub DeleteLimiteTpProducto(ByVal pcCtaCod As String, ByVal pnTpProducto As String, ByVal nValor As Integer)
 Dim lsSQL As String
 On Error GoTo dError
    lsSQL = "EXEC stp_del_ERS047_SolicitudAutorizacionZonaxProductoxGarant '" & pcCtaCod & "', '" & pnTpProducto & "'," & nValor & ""
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Hubo un Error Comunicarse con TI-Desarrollo.", Err.Description
End Sub

Public Function ValidaLimiteTpProducto(ByVal pcCtaCod As String, ByVal pnTpProducto As String) As ADODB.Recordset
 Dim lsSQL As String
 Dim rs As ADODB.Recordset
 On Error GoTo dError
    lsSQL = "EXEC stp_sel_ERS047_VerificaSolicitudAutorizacionProducto '" & pcCtaCod & "', '" & pnTpProducto & "'"
    Set rs = coConex.CargaRecordSet(lsSQL)
    Set ValidaLimiteTpProducto = rs
    Set rs = Nothing
    Exit Function
dError:
        Err.Raise Err.Number, "Hubo un Error Comunicarse con TI-Desarrollo.", Err.Description
End Function
'JOEP ERS047 20170904

'Inicio JOEP - ARLO ERS082-2017 - 20171213 ----MODIFICADO DESDE LA 60
    'JOEP
Public Function ObtieneDatosFlexEdit(ByVal pnFlexValor As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
On Error GoTo dError
    sSql = "exec stp_sel_ERS082_2017_ObtenerDatosFlexParametroPig " & pnFlexValor & ""
    Set rs = coConex.CargaRecordSet(sSql)
    Set ObtieneDatosFlexEdit = rs
    Set rs = Nothing
    Exit Function
dError:
        Err.Raise Err.Number, "Hubo un Error Comunicarse con TI-Desarrollo.", Err.Description
End Function

Public Function ObtieneDatosCondicionFlex() As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
On Error GoTo dError
    sSql = "exec stp_sel_ERS082_2017_CargaDatosCondicion "
    Set rs = coConex.CargaRecordSet(sSql)
    Set ObtieneDatosCondicionFlex = rs
    Set rs = Nothing
    Exit Function
dError:
        Err.Raise Err.Number, "Hubo un Error Comunicarse con TI-Desarrollo.", Err.Description
End Function
'------FIN MODIFICADO DESDE LA 60

'Public Sub RegistraDatosSegmentacion(ByVal nTpFlex As Integer, ByVal nId As Integer, Optional ByVal pcSeg As String, Optional ByVal pcSubSeg As String, Optional ByVal pnMontoTasaDesde As Currency, Optional ByVal pnMontoTasaHasta As Currency, Optional ByVal pnDiasDesde As Currency, Optional ByVal pnDiasHasta As Currency, Optional ByVal pnCond As Integer, Optional ByVal pnCantAdj As Integer, Optional ByVal pnDiasAdj As Integer, Optional ByVal pnTipoCli As Integer)
Public Sub RegistraDatosSegmentacion(ByVal nTpFlex As Integer, ByVal nId As Integer, Optional ByVal pcSeg As String = "", Optional ByVal pcSubSeg As String = "", _
Optional ByVal pnMontoTasaDesde As Currency = -1, Optional ByVal pnMontoTasaHasta As Currency = -1, Optional ByVal pnDiasDesde As Currency = -1, Optional ByVal pnDiasHasta As Currency = -1, _
Optional ByVal pnCond As Integer = -1, Optional ByVal pnCantAdj As Integer = -1, Optional ByVal pnDiasAdj As Integer = -1, Optional ByVal pnTipoCli As Integer = -1, Optional ByVal pnSegExt As Integer = -1, Optional ByVal pcMovNro As String = "") 'JOEP20210812 Segmentacion Externa
 Dim lsSQL As String
 On Error GoTo dError
    'lsSQL = "EXEC stp_ins_ERS082_2017_RegistraDatosCliRecu " & nTpFlex & "," & nId & ",'" & pcSeg & "','" & pcSubSeg & "'," & pnMontoTasaDesde & "," & pnMontoTasaHasta & "," & pnDiasDesde & "," & pnDiasHasta & "," & pnCond & "," & pnCantAdj & "," & pnDiasAdj & "," & pnTipoCli & ""
    lsSQL = "EXEC stp_ins_ERS082_2017_RegistraDatosCliRecu " & nTpFlex & "," & nId & ",'" & pcSeg & "','" & pcSubSeg & "'," & pnMontoTasaDesde & "," & pnMontoTasaHasta & "," & pnDiasDesde & "," & pnDiasHasta & "," & pnCond & "," & pnCantAdj & "," & pnDiasAdj & "," & pnTipoCli & "," & pnSegExt & ",'" & pcMovNro & "'" 'JOEP20210812 Segmentacion Externa
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
    Err.Raise Err.Number, "Hubo un Error Comunicarse con TI-Desarrollo.", Err.Description
End Sub
'JOEP

'ARLO
Public Function dVerificarSegmentacion(ByVal psPersCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_ERS082_2017_dVerificarSegmentacion '" & psPersCod & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set dVerificarSegmentacion = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
        
End Function
Public Function dVerificarConfSegmentacion(ByVal psTpoCliente As Integer) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_ERS082_2017_dVerificarConfSegmentacion '" & psTpoCliente & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set dVerificarConfSegmentacion = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
        
End Function
Public Function dVerificaTpoCliente(ByVal psPersCod As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_ERS082_2017_dVerificaTpoCliente '" & psPersCod & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set dVerificaTpoCliente = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
        
End Function
Public Function dObtieneValorSegmentacion(ByVal psSegmento As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo dError
    
        lsSQL = "EXEC stp_sel_ERS082_2017_dObtieneValorSegmentacion '" & psSegmento & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    
    Set dObtieneValorSegmentacion = lrs
    Set lrs = Nothing
    
    Exit Function
    
dError:
        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
        
End Function
'Comento JOEP20210513
'Public Sub RegistroFormatoEvalPig(ByVal psCctaCod As String, ByVal pnIngreso As Double, ByVal pnEgreso As Double, ByVal psMovNro As String, ByVal pnTipCli As Integer, ByVal pnCantAdjuSeg As Integer, ByVal pnDiasSeg As Integer, ByVal pcSegmento As String, ByVal pnDiasSinAdj As Integer, ByVal pnSegmentoAnt As String)  'Agrego JOEP20180220 pnTipCli,pnCantAdjuSeg,pnDiasSeg,pcSegmento: JOEP Pig Pase pnDiasSinAdj,pnSegmentoAnt
'    Dim lsSQL As String
'    On Error GoTo dError
'        lsSQL = "EXEC stp_ins_ERS082_2017_RegistroFormatoEvalPig '" & psCctaCod & "'," & pnIngreso & " ," & pnEgreso & ",'" & psMovNro & "'," & pnTipCli & "," & pnCantAdjuSeg & "," & pnDiasSeg & ",'" & pcSegmento & "'," & pnDiasSinAdj & ",'" & pnSegmentoAnt & "'" 'JOEP Pig Pase pnDiasSinAdj,pnSegmentoAnt
'
'    coConex.CargaRecordSet (lsSQL)
'    Exit Sub
'dError:
'        Err.Raise Err.Number, "Registra Tipo", Err.Description
'End Sub
'Comento JOEP20210513

'Add JOEP20210513 Seg. Prendario Externo
Public Sub RegistroFormatoEvalPig(ByVal psCctaCod As String, ByVal pnIngreso As Double, ByVal pnEgreso As Double, ByVal psMovNro As String, ByVal pnTipCli As Integer, ByVal pnCantAdjuSeg As Integer, ByVal pnDiasSeg As Integer, ByVal pcSegmento As String, ByVal pnDiasSinAdj As Integer, ByVal pnSegmentoAnt As String, Optional ByRef pArrayDatosSegPredExt As Variant = Nothing) 'Add JOEP20210513 Seg. Prendario Externo
    Dim lsSQL As String
    On Error GoTo dError
    
    Dim CredVenc As Integer, ObtEmpresa As Integer
    Dim nCalCPP As Double, RC18 As Double, RC17 As Double, RC16 As Double, RC15 As Double, RC14 As Double, RC13 As Double, RC12 As Double, RC11 As Double, RC10 As Double, RC09 As Double, RC08 As Double, RC07 As Double
    Dim dFechaRCC As String
    Dim SegPrenExter As String
    Dim iMat As Integer
    CredVenc = -1
    ObtEmpresa = -1
    SegPrenExter = ""
    nCalCPP = -1
    RC18 = -1
    RC17 = -1
    RC16 = -1
    RC15 = -1
    RC14 = -1
    RC13 = -1
    RC12 = -1
    RC11 = -1
    RC10 = -1
    RC09 = -1
    RC08 = -1
    RC07 = -1
    dFechaRCC = ""
    
    If IsArray(pArrayDatosSegPredExt) Then
        For iMat = 0 To UBound(pArrayDatosSegPredExt) - 1
            Select Case iMat
                Case 0: CredVenc = pArrayDatosSegPredExt(iMat)
                Case 1: ObtEmpresa = pArrayDatosSegPredExt(iMat)
                Case 2: nCalCPP = pArrayDatosSegPredExt(iMat)
                Case 3: RC18 = pArrayDatosSegPredExt(iMat)
                Case 4: RC17 = pArrayDatosSegPredExt(iMat)
                Case 5: RC16 = pArrayDatosSegPredExt(iMat)
                Case 6: RC15 = pArrayDatosSegPredExt(iMat)
                Case 7: RC14 = pArrayDatosSegPredExt(iMat)
                Case 8: RC13 = pArrayDatosSegPredExt(iMat)
                Case 9: RC12 = pArrayDatosSegPredExt(iMat)
                Case 10: RC11 = pArrayDatosSegPredExt(iMat)
                Case 11: RC10 = pArrayDatosSegPredExt(iMat)
                Case 12: RC09 = pArrayDatosSegPredExt(iMat)
                Case 13: RC08 = pArrayDatosSegPredExt(iMat)
                Case 14: RC07 = pArrayDatosSegPredExt(iMat)
                Case 15: dFechaRCC = pArrayDatosSegPredExt(iMat)
                Case 16: SegPrenExter = pArrayDatosSegPredExt(iMat)
            End Select
        Next
    End If
            
        lsSQL = "EXEC stp_ins_ERS082_2017_RegistroFormatoEvalPig '" & psCctaCod & "'," & pnIngreso & " ," & pnEgreso & ",'" & psMovNro & "'," & pnTipCli & "," & pnCantAdjuSeg & "," & pnDiasSeg & ",'" & pcSegmento & "'," & pnDiasSinAdj & ",'" & pnSegmentoAnt & "'," _
        & CredVenc & "," & ObtEmpresa & "," & nCalCPP & "," & RC18 & "," & RC17 & "," & RC16 & "," & RC15 & "," & RC14 & "," & RC13 & "," & RC12 & "," & RC11 & "," & RC10 & "," & RC09 & "," & RC08 & "," & RC07 & ",'" & Format(dFechaRCC, "yyyymmdd") & "','" & SegPrenExter & "'"
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Registra Tipo", Err.Description
End Sub
'Add JOEP20210513 Seg. Prendario Externo

'ARLO
'Inicio JOEP - ARLO ERS082-2017 - 20171213

'JOEP20180412 Pig Pase
Public Sub ObtieneSaldoCartSaldoCap(ByVal pcCtaCod As String)
    Dim lsSQL As String
On Error GoTo dError
        lsSQL = "EXEC stp_sel_ObtieneSaldoCarteraSaldoCapPigno '" & pcCtaCod & "'"
    coConex.CargaRecordSet (lsSQL)
Exit Sub
    
dError:
        Err.Raise Err.Number, "Obtiene saldo de cartera, saldo capital titular", Err.Description
End Sub
'JOEP20180412 Pig Pase
'JUCS TI-ERS063-2017
Public Function ListaHistHologramas(ByVal Agencia As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
   ' On Error GoTo ErrListaHistHologramas
    sSql = "exec stp_sel_HistHologramas '" & Agencia & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ListaHistHologramas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    End Function
'JUCS TI ERS063-2017 InsertHolograma
Public Sub GuardarHologramas(ByVal psMovNro As String, ByVal nHologIni As Long, ByVal nHologFin As Long)
 Dim lsSQL As String
 On Error GoTo dError
      lsSQL = "EXEC stp_ins_Hologramas '" & psMovNro & "', " & nHologIni & "," & nHologFin
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Hubo un Error al registrar el Holograma Comunicarse con TI-Desarrollo.", Err.Description
End Sub
'END insert JUCS

'JUCS TI ERS063-2017 ActualizaHolograma
Public Sub ActualizaMontosHologramas(ByVal piHologramaID As Integer, ByVal nHologIni As Long, ByVal nHologFin As Long)
    Dim lsSQL As String
    On Error GoTo dError
        lsSQL = "EXEC stp_upd_ActualizarHolograma " & piHologramaID & "," & nHologIni & "," & nHologFin
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dError:
        Err.Raise Err.Number, "Obtiene Lista de Credito", Err.Description
End Sub
'End Update

'JUCS TI 063 VERIFICAR RANGO HOLOGRAMAS
  Public Function VerificaRangoUtilizado(ByVal nHologIni As Long, ByVal nHologFin As Long, ByVal pnTipo As Integer, ByVal pcAgeCod As String) As Boolean
  'APRI20181018 MEJORA ADD pcAgeCod
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
       On Error GoTo ErrVerificaRangoUtilizado
    'lsSQL = "SELECT  VerificaRangHolog=dbo.ExisteRangoHolograma ( '" & nHologIni & "','" & nHologFin & "'," & pnTipo & ")"
    lsSQL = "SELECT  VerificaRangHolog=dbo.ExisteRangoHolograma ( '" & nHologIni & "','" & nHologFin & "'," & pnTipo & ",'" & pcAgeCod & "')" 'APRI20181018 MEJORA
    Set rs = coConex.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        VerificaRangoUtilizado = rs!VerificaRangHolog
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrVerificaRangoUtilizado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'END JUCS FUNCION RANGO HOLOGRAMAS

'JUCS FUNCION VERIFICA SI EXISTE HOLOGRAMA ACTIVO
  Public Function VerificaExisteHologramaActivo(ByVal Agencia As Integer) As Boolean
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
       On Error GoTo ErrVerificaExisteHologramaActivo
    lsSQL = "SELECT  VerificaHologActivo=dbo.ExisteHologramaActivo ('" & Agencia & "')"
    Set rs = coConex.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        VerificaExisteHologramaActivo = rs!VerificaHologActivo
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrVerificaExisteHologramaActivo:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'END FUNCIÓN JUCS

'Modificación JUCS TI-ERS063-2017 para frmColPRegContratoDet
Public Function ObtieneTasador(ByVal Agencia As Integer) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrObtieneTasador
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
             sSql = "exec stp_sel_ObtieneUserRegContrato '" & Agencia & "' "
    Set ObtieneTasador = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrObtieneTasador:
    Err.Raise Err.Number, "Error En Proceso ObtieneTasador", Err.Description
End Function
Public Function ObtenerHologramaActivo(ByVal Agencia As String) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrObtenerHologramaActivo
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSql = "exec ObtenerHologramaActivo '" & Agencia & "' "
    Set ObtenerHologramaActivo = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrObtenerHologramaActivo:
    Err.Raise Err.Number, "Error En Proceso: ObtenerHologramaActivo", Err.Description
End Function
'JUCS FUNCION VERIFICA SI EXISTE HOLOGRAMA IGUAL EN ColocPigJoyaDet
  Public Function HologDuplicado(ByVal Holograma As String, ByVal cCodAge As String) As Boolean
  'APRI20190515 ERS005-209 ADD cCodAge
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrHologDuplicado
    lsSQL = "SELECT  VerificaHolograma=dbo.ExisteHologramasIguales (" & Holograma & ",'" & cCodAge & "')"
   oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    If Not rs.EOF Then
        HologDuplicado = rs!VerificaHolograma
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrHologDuplicado:
    Err.Raise Err.Number, "Error En Proceso HologDuplicado", Err.Description
End Function
'END FUNCIÓN JUCS
Public Function ObtieneContador(ByVal Agencia As Integer) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrObtieneContador
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
             sSql = "exec stp_sel_ObtieneContadorHolog " & Agencia & " "
    Set ObtieneContador = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrObtieneContador:
    Err.Raise Err.Number, "Error En Proceso ObtieneContador", Err.Description
End Function
'JUCS Actualizar contador (incrementa el monto por cada contrato)
Public Function ActualizaContador(ByVal Contador As Integer, ByVal cMovNro As String)
Dim lsSQL As String
    lsSQL = " exec stp_upd_ContadorHolog  '" & Contador & "', '" & cMovNro & "'"
    coConex.ConexionActiva.Execute lsSQL
    coConex.CargaRecordSet (lsSQL)
End Function
Public Function ObtieneTasadorHolograma(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrObtieneTasador
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
             sSql = "exec stp_sel_ObtenerTasadorHolograma '" & pcCtaCod & "' "
    Set ObtieneTasadorHolograma = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrObtieneTasador:
    Err.Raise Err.Number, "Error En Proceso ObtieneTasador", Err.Description
End Function
'END JUCS
'APRI20190515 ERS005-2019
Public Function ObtenerMaxNroHolograma(ByVal pnCod As String) As Long
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrObtenerMaxNroHolograma
    lsSQL = "SELECT  MaxNroHolograma=dbo.ObtenerMaxNroHolograma (" & pnCod & ")"
   oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    If Not rs.EOF Then
        ObtenerMaxNroHolograma = rs!MaxNroHolograma
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrObtenerMaxNroHolograma:
    Err.Raise Err.Number, "Error En Proceso ObtenerMaxNroHolograma", Err.Description
End Function
Public Function ObtieneHologramasPendientes(ByVal pnCod As Integer) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrObtieneHologramasPendientes
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSql = "exec stp_sel_ObtenerHologramaPendiente " & pnCod
    Set ObtieneHologramasPendientes = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrObtieneHologramasPendientes:
    Err.Raise Err.Number, "Error En Proceso HologramasPendientes", Err.Description
End Function
'END APRI

'CROB 20180524 BEGIN
Public Function ExisteAprobacionCredNivelesPendientesPigno(ByVal psCtaCod As String) As Boolean
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    ExisteAprobacionCredNivelesPendientesPigno = False
    lsSQL = "exec stp_sel_VerificaAprobacionCredNivelesPendientesPigno '" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    If Not lrDatos.EOF Then
        ExisteAprobacionCredNivelesPendientesPigno = True
    End If
    Set lrDatos = Nothing
End Function

Public Sub RegistrarSolicitudNivAprobPigno(ByVal psCctaCod As String, ByVal pnCuota As Integer, ByVal pnTasa As Currency, ByVal psCPersCod As String, ByVal pnTasacion As Currency, _
                                            ByVal pnPrestamo As Currency, ByVal pnTotPiezas As Integer, ByVal pnTotOroBruto As Currency, ByVal pnTotOroNeto As Currency, _
                                            ByVal pcMovNroSol As String, ByVal psFechaHora As String)
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    
    lsSQL = "exec stp_ins_RegistrarSolicitudNivAprobPigno '" & psCctaCod & "'," & pnCuota & "," & pnTasa & ",'" & psCPersCod & "'," & pnTasacion & "," & pnPrestamo & "," & pnTotPiezas & "," _
                                                           & pnTotOroBruto & "," & pnTotOroNeto & ",'" & pcMovNroSol & "'," & "'" & psFechaHora & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub

Public Function ValidarNivelAprobacionPigno(ByVal psCtaCod As String) As Boolean
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    ValidarNivelAprobacionPigno = False
    lsSQL = "exec stp_sel_ValidarNivelAprobacionPigno '" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    If Not lrDatos.EOF Then
        ValidarNivelAprobacionPigno = True
    End If
    Set lrDatos = Nothing
End Function

Public Function RecuperaCreditosPorAprobarPigno(ByVal psCodUser As String, ByVal psAgeCod As String, ByVal pdFecha As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrCargar
    lsSQL = "Exec stp_sel_RecuperaCreditosPorAprobarPigno '" & psCodUser & "','" & psAgeCod & "'" & ",'" & Format(pdFecha, "yyyy-mm-d") & "'"
    Set RecuperaCreditosPorAprobarPigno = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrCargar:
    Err.Raise Err.Number, "Error en proceso (DCOMNivelAprobacion: RecuperaCreditosPorAprobarPigno)", Err.Description
End Function

Public Sub AprobarCreditoNivelAprobacionPigno(ByVal psCtaCod As String, ByVal psMovNroApr As String)
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "exec stp_ins_AprobarCreditoNivelAprobacionPigno '" & psCtaCod & "','" & psMovNroApr & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub

Public Sub RechazarCreditoNivelAprobacionPigno(ByVal psCtaCod As String, ByVal psMovNroRechazo As String)
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "exec stp_upd_RechazarCreditoNivelAprobacionPigno '" & psCtaCod & "','" & psMovNroRechazo & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub

Public Function ObtenerUITdelAnio(ByVal psValor As String) As Currency
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    ObtenerUITdelAnio = False
    lsSQL = "exec sp_sel_ParametroCredito '" & psValor & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    If Not lrDatos.EOF Then
        ObtenerUITdelAnio = lrDatos!nParamValor
    Else
        ObtenerUITdelAnio = 0
    End If
    Set lrDatos = Nothing
End Function
'END CROB 20180524

'JOEP20190614 Mejora en la evaluacion de segmentacion acta 084-2019
Public Function ObtieneResulSeg(ByVal pcPersCod As String, ByVal nMontoTasa As Currency, ByVal nCalfCPP As Integer, ByVal nDiasPerma As Integer, ByVal nCantAdj As Integer, ByVal nDiasSinAdj As Integer, ByVal nVerVencidos As Integer, ByVal nTipoCliente As Integer, Optional ByVal nObtEmpresa As Integer = 0, Optional ByVal nRCC18 As Integer = 0, _
Optional ByVal nRCC17 As Integer = 0, Optional ByVal nRCC16 As Integer = 0, Optional ByVal nRCC15 As Integer = 0, Optional ByVal nRCC14 As Integer = 0, Optional ByVal nRCC13 As Integer = 0, Optional ByVal nRCC12 As Integer = 0, Optional ByVal nRCC11 As Integer = 0, Optional ByVal nRCC10 As Integer = 0, Optional ByVal nRCC09 As Integer = 0, _
Optional ByVal nRCC08 As Integer = 0, Optional ByVal nRCC07 As Integer = 0) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrObtieneResulSeg
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
             sSql = "exec stp_sel_ObtenerCalificacion '" & pcPersCod & "'," & nMontoTasa & "," & nCalfCPP & "," & nDiasPerma & "," & nCantAdj & "," & nDiasSinAdj & "," & nVerVencidos & "," & nTipoCliente & "," & nObtEmpresa & "," & _
nRCC18 & "," & nRCC17 & "," & nRCC16 & "," & nRCC15 & "," & nRCC14 & "," & nRCC13 & "," & nRCC12 & "," & nRCC11 & "," & nRCC10 & "," & nRCC09 & "," & nRCC08 & "," & nRCC07 & ""
    Set ObtieneResulSeg = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrObtieneResulSeg:
    Err.Raise Err.Number, "Error En Proceso Evaluacion de Segmentacion", Err.Description
End Function
'JOEP20190614 Mejora en la evaluacion de segmentacion acta 084-2019
'APRI20181025 MEJORA
Public Sub DeshabilitarRangoHolograma(ByVal piHologramaID As Integer, ByVal pcMovNroDes As String)
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "exec stp_upd_DesactivarRangoHolograma " & piHologramaID & ",'" & pcMovNroDes & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub
'END APRI


'JOEP20210925 campana prendario
Public Function CampPrendarioConfigDatos() As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrCampPrendarioConfigDatos
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
             sSql = "exec stp_sel_CampanaPrendarioObtieneConiguracionTasas "
    Set CampPrendarioConfigDatos = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrCampPrendarioConfigDatos:
    Err.Raise Err.Number, "Error Configuracion de tasa Retencion Prendario", Err.Description
End Function

Public Sub CampPrendarioRegConf(ByVal pnMontoIni As Double, ByVal pnMontoFin As Double, ByVal pnTEAIni As Double, ByVal pnTEAFin As Double, ByVal pnId As Integer, ByVal pcMovNro As String, ByVal pCont As Integer)
    Dim lsSQL As String
    On Error GoTo dErrorCampPrendarioRegConf
    lsSQL = "EXEC stp_sel_CampanaPrendarioRegistraConiguracionTasas " & pnMontoIni & "," & pnMontoFin & " ," & pnTEAIni & "," & pnTEAFin & "," & pnId & ",'" & pcMovNro & "'," & pCont & ""
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dErrorCampPrendarioRegConf:
        Err.Raise Err.Number, "Registra Configuracion de tasas Retencion", Err.Description
End Sub

Public Function CampPrendarioObtDatosPers(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrCampPrendarioObtDatosPers
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSql = "exec stp_sel_CampanaPrendarioObtieneDatosPersonaRetencion '" & pcCtaCod & "'"
    Set CampPrendarioObtDatosPers = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrCampPrendarioObtDatosPers:
    Err.Raise Err.Number, "Error Obtener datos Cliente Retencion Prendario", Err.Description
End Function

Public Sub CampPrenRegRetencion(ByVal pcCtaCod As String, ByVal pnTEAMin As Double, ByVal pnTEAMax As Double, ByVal pnTEANew As Double, ByVal pcObser As String, ByVal pcMovNro As String, ByVal pcMovNroVB As String, ByVal pnEstado As Integer)
    Dim lsSQL As String
    On Error GoTo dErrorCampPrenRegRetencion
    lsSQL = "EXEC stp_ins_CampanaPrendarioRegistroRetencion '" & pcCtaCod & "'," & pnTEAMin & " ," & pnTEAMax & "," & pnTEANew & ",'" & pcObser & "','" & pcMovNro & "','" & pcMovNroVB & "'," & pnEstado & ""
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dErrorCampPrenRegRetencion:
        Err.Raise Err.Number, "Registra de Retencion de credito", Err.Description
End Sub

Public Function CampPrendarioDescCampa(ByVal pnIdcampana As Integer) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrCampPrendarioDescCampa
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSql = "exec stp_sel_CampanaPrendarioDescripcionCampa " & pnIdcampana & ""
    Set CampPrendarioDescCampa = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrCampPrendarioDescCampa:
    Err.Raise Err.Number, "Error Obtener datos Cliente Retencion Prendario", Err.Description
End Function

Public Function CampPrendarioActivaCampana(ByVal pcPersCod As String) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrCampPrendarioActivaCampana
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSql = "exec stp_sel_CampanaPrendarioActivaCampanaPrendario '" & pcPersCod & "'"
    Set CampPrendarioActivaCampana = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrCampPrendarioActivaCampana:
    Err.Raise Err.Number, "Error Obtener datos Cliente Retencion Prendario", Err.Description
End Function

Public Sub CampPrenRegCampCred(ByVal pcCtaCod As String, ByVal pnCampana As Integer, ByVal pcGlosa As String, ByVal pnTEN As Double, ByVal pnRenovacion As Integer, ByVal pnDesct As Double, ByVal pnTEMNew As Double, ByVal pcMovNro As String, ByVal pnEstado As Integer, ByVal pnModulo As Integer)
    Dim lsSQL As String
    On Error GoTo dErrorCampPrenRegCampCred
    lsSQL = "EXEC stp_ins_CampanaPrendarioRegistroCredCamp '" & pcCtaCod & "'," & pnCampana & " ,'" & pcGlosa & "'," & pnTEN & "," & pnRenovacion & "," & pnDesct & "," & pnTEMNew & ",'" & pcMovNro & "'," & pnEstado & "," & pnModulo & ""
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dErrorCampPrenRegCampCred:
        Err.Raise Err.Number, "Registra de Retencion de credito", Err.Description
End Sub

Public Function CampPrendarioDesVerificaCampReten(ByVal pcCtaCod As String, ByVal pdFechaSis As String) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrCampPrendarioDesVerificaCampReten
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSql = "exec stp_sel_CampanaPrendarioRenovacionVerificaRetencionCampana '" & pcCtaCod & "','" & Format(pdFechaSis, "YYYYMMDD") & "'"
    Set CampPrendarioDesVerificaCampReten = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrCampPrendarioDesVerificaCampReten:
    Err.Raise Err.Number, "Error Obtener datos Cliente Retencion Prendario", Err.Description
End Function

Public Sub CampPrenRegRenovacionRetencion(ByVal pcCtaCod As String, ByVal pnTEM As Double, ByVal pcMovNro As String, ByVal pnEstado As Integer)
    Dim lsSQL As String
    On Error GoTo dErrorCampPrenRegRenovacionRetencion
    lsSQL = "EXEC stp_ins_CampanaPrendarioRegRenovacionRetencion '" & pcCtaCod & "'," & pnTEM & ",'" & pcMovNro & "'," & pnEstado & ""
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dErrorCampPrenRegRenovacionRetencion:
        Err.Raise Err.Number, "Registra de Retencion de credito", Err.Description
End Sub

Public Function CampPrendarioDesbCampa(ByVal pcCtaCod As String, ByVal dfechaSis As String, ByVal pnModulo As Integer, Optional ByVal pnRenov As Integer = -1) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrCampPrendarioDesbCampa
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSql = "exec stp_sel_CampanaPrendarioDesembolsoVerificacampana '" & pcCtaCod & "','" & Format(dfechaSis, "YYYYMMDD") & "'," & pnModulo & "," & pnRenov & ""
    Set CampPrendarioDesbCampa = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrCampPrendarioDesbCampa:
    Err.Raise Err.Number, "Error Obtener datos Cliente Retencion Prendario", Err.Description
End Function

Public Function CampPrendarioDesbVerfCampaReferido(ByVal pcPersCod As String, ByVal pcCodUser As String) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrCampPrendarioDesbVerfCampaReferido
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSql = "exec stp_sel_CampanaPrendarioDesembolsoVerfCampanaReferido '" & pcPersCod & "','" & pcCodUser & "'"
    Set CampPrendarioDesbVerfCampaReferido = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrCampPrendarioDesbVerfCampaReferido:
    Err.Raise Err.Number, "Error Obtener datos Cliente Retencion Prendario", Err.Description
End Function

Public Sub CampPrenDesbRegReferido(ByVal pcCtaCod As String, ByVal pcPersCodReferido As String, ByVal pcMovNro As String, ByVal pnEstado As Integer)
    Dim lsSQL As String
    On Error GoTo dErrorCampPrenDesbRegReferido
    lsSQL = "EXEC stp_ins_CampanaPrendarioDesembolsoRegReferido '" & pcCtaCod & "','" & pcPersCodReferido & "','" & pcMovNro & "'," & pnEstado & ""
    coConex.CargaRecordSet (lsSQL)
    Exit Sub
dErrorCampPrenDesbRegReferido:
        Err.Raise Err.Number, "Registra de Retencion de credito", Err.Description
End Sub

Public Function CampPrendarioConsulta(ByVal pcCtaCod As String, ByVal pnModulo As Integer) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrCampPrendarioConsulta
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSql = "exec stp_sel_CampanaPrendarioViewConsultar '" & pcCtaCod & "'," & pnModulo & ""
    Set CampPrendarioConsulta = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrCampPrendarioConsulta:
    Err.Raise Err.Number, "Error Obtener datos Cliente Retencion Prendario", Err.Description
End Function

Public Function CampPrendarioHabilitaReferido(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrCampPrendarioHabilitaReferido
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSql = "exec stp_sel_CampanaPrendarioHabilitaReferidos '" & pcCtaCod & "'"
    Set CampPrendarioHabilitaReferido = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrCampPrendarioHabilitaReferido:
    Err.Raise Err.Number, "Error Obtener datos Cliente Retencion Prendario", Err.Description
End Function

Public Function CampPrendarioValidadReferido(ByVal pcCtaCod As String, ByVal pnTEANew As Double) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrCampPrendarioValidadReferido
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSql = "exec stp_sel_CampanaPrendarioValidaRetencion '" & pcCtaCod & "'," & pnTEANew & ""
    Set CampPrendarioValidadReferido = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrCampPrendarioValidadReferido:
    Err.Raise Err.Number, "Error Obtener datos Cliente Retencion Prendario", Err.Description
End Function
'JOEP20210925 campana prendario
'RIRO 20210921
Public Function CampDevuelveCampana(ByVal psCtaCod As String) As Integer

Dim lrs As ADODB.Recordset
Dim lsSQL As String

CampDevuelveCampana = 0
Set lrs = New ADODB.Recordset

On Error GoTo dError

Dim loRegValida As DCOMColPFunciones
Dim lrValida As ADODB.Recordset
    
lsSQL = "exec stp_sel_CampanaPrendarioViewConsultar '" & psCtaCod & "', 1"
    
Set lrs = coConex.CargaRecordSet(lsSQL)

If Not lrs Is Nothing Then
    If Not lrs.EOF And Not lrs.BOF Then
        If lrs.RecordCount > 0 Then
            CampDevuelveCampana = lrs!nCampana
        End If
    End If
End If

Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Costos Contrato en <<nObtieneDatosObtieneCosNotarialPignoraticio>>", Err.Description
End Function
'END RIRO ****************


'JOEP20210812 Segmentacion Externa
Public Function ConfiSegTitulo(ByVal nOpTitulo As Integer) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrConfiSegTitulo
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSql = "exec stp_sel_SegmentacionTitulos " & nOpTitulo & ""
    Set ConfiSegTitulo = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrConfiSegTitulo:
    Err.Raise Err.Number, "Error Configuracion Segmentacion Prendario Externo", Err.Description
End Function

Public Function ConfiSegExtDatos(ByVal nOpcion As Integer) As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrConfiSegExtDatos
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSql = "exec stp_sel_SegmentacionCargaDatos " & nOpcion & ""
    Set ConfiSegExtDatos = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function
ErrConfiSegExtDatos:
    Err.Raise Err.Number, "Error Configuracion Segmentacion Prendario Externo Datos", Err.Description
End Function
'JOEP20210812 Segmentacion Externa
