VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DColPContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String

Dim coConex As DConecta
Private Sub Class_Initialize()
    Dim loIni As ClasIni
    
    Set loIni = New ClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing

Set coConex = New DConecta
If coConex.AbreConexion(csConexion) = False Then
    Call RaiseError(MyUnhandledError, "DColPContrato:Initialize. Error en Conexion a Base de datos")
End If
End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

Public Function dObtieneDatosCreditoPignoraticio(ByVal psCuenta As String) As Recordset
'Obtiene Datos de Contrato Pignoraticio
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = "SELECT  P.cCtaCod, P.nSaldo, P.nPrdEstado, P.nTasaInteres, " _
    & " C.nPlazo, C.dVenc, C.nMontoCol, C.dVigencia, " _
    & " CP.nOroBruto, CP.nOroNeto, CP.nPiezas, CP.nTasacion,CP.cLote, CP.cPrdCtaTpo,  " _
    & " nPlazoIni = ( SELECT nPlazo FROM ColocacEstado WHERE cCtaCod ='" & psCuenta & "' " _
    & "                 AND nPrdEstado = " & gColPEstRegis & "  AND nPlazo > 0 ),  " _
    & " dFecVencIni = ( SELECT DateAdd(d, nPlazo, dPrdEstado) FROM ColocacEstado WHERE cCtaCod ='" & psCuenta & "' " _
    & "                 AND nPrdEstado = " & gColPEstRegis & "  AND nPlazo > 0 )   " _
    & " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod " _
    & " INNER JOIN ColocPignoraticio CP ON C.cCtaCod = CP.cCtaCod " _
    & " WHERE P.cCtaCod = '" & psCuenta & "'"

Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneDatosCreditoPignoraticio = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

Public Function dObtieneDatosCreditoPignoraticioJoyas(ByVal psCuenta As String) As Recordset
'Obtiene Datos de Credito Pignoraticio Joyas
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

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

Public Function nObtieneDatosCostosRegistroCredPignoraticio(ByVal psCtaCod As String) As Recordset

'************************************
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, dVigencia, P.nTasaInteres, CP.nNroDuplic, " _
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
    
Set lrs = coConex.CargaRecordSet(lsSQL)

Set nObtieneDatosCostosRegistroCredPignoraticio = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
End Function

Public Function dObtieneDatosCreditoPignoraticioPersonas(ByVal psCuenta As String) As Recordset
'Obtiene Datos de Credito Pignoraticio PERSONA
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = "SELECT PP.cPersCod, P.cPersNombre, P.cPersDireccDomicilio, P.cPersTelefono, " _
    & " P.cPersDireccUbiGeo, " _
    & " Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),'') , " _
    & " Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ) , " _
    & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) , " _
    & " NroDNI = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdDNI & " ),   " _
    & " NroRUC = (Select ISNULL(cPersIDnro,'') From PersID WHERE cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdRUC & " )   " _
    & " FROM ProductoPersona PP INNER JOIN Persona P  ON PP.cPersCod = P.cPersCod " _
    & " WHERE cCtaCod = '" & psCuenta & "' "

Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneDatosCreditoPignoraticioPersonas = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function


Public Function dObtieneCredPigDePersona(ByVal psPersCod As String, _
        ByVal psEstados As String, Optional ByVal psAgencia As String = "") As Recordset
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

If psEstados = "2810,2812" Then

    lsSQL = "SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, " _
            & "T1.cConsDescripcion cRelacion, UPPER(T2.cConsDescripcion) cProducto, " _
            & "UPPER(T3.cConsDescripcion) cMoneda " _
            & "FROM ProductoPersona PP INNER JOIN Producto P  " _
            & "ON P.cCtaCod = PP.cCtaCod INNER JOIN " _
            & csCentralCom & "Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN " & csCentralCom & "" _
            & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN " & csCentralCom & "Constante T2 " _
            & "ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2.nConsValor) INNER JOIN " & csCentralCom & "" _
            & "Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) " _
            & "INNER JOIN COLOCPIGSOBRANTE S ON S.cCtaCod = P.cCtaCod " _
            & "WHERE PP.cPersCod = '" & psPersCod & "' " _
            & "AND T1.nConsCod = " & gColocRelacPers & " AND T.nConsCod = " & gColocEstado & " AND " _
            & "T2.nConsCod = " & gProducto & " AND T3.nConsCod = " & gMoneda _
            & "AND P.cCtaCod like '___" & lsAgencia & "305" & "%' " _
            & "AND P.nPrdEstado in ( " & psEstados & " ) " _
            & "AND S.nEstadoSobrante = 0"
Else
    lsSQL = "SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, " _
            & "T1.cConsDescripcion cRelacion, UPPER(T2.cConsDescripcion) cProducto, " _
            & "UPPER(T3.cConsDescripcion) cMoneda " _
            & "FROM ProductoPersona PP INNER JOIN Producto P  " _
            & "ON P.cCtaCod = PP.cCtaCod INNER JOIN " _
            & csCentralCom & "Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN " & csCentralCom & "" _
            & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN " & csCentralCom & "Constante T2 " _
            & "ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2.nConsValor) INNER JOIN " & csCentralCom & "" _
            & "Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) " _
            & "WHERE PP.cPersCod = '" & psPersCod & "' " _
            & "AND T1.nConsCod = " & gColocRelacPers & " AND T.nConsCod = " & gColocEstado & " AND " _
            & "T2.nConsCod = " & gProducto & " AND T3.nConsCod = " & gMoneda _
            & "AND P.cCtaCod like '___" & lsAgencia & "305" & "%' " _
            & "AND P.nPrdEstado in ( " & psEstados & " ) "
End If

lsSQL = lsSQL & " ORDER BY PP.cCtaCod"


Set lrs = coConex.CargaRecordSet(lsSQL)

Set dObtieneCredPigDePersona = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

