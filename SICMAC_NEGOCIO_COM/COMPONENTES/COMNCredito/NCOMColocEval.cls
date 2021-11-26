VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMColocEval"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************
'* Clase que contiene las Consultas para Calificacion de Cartera
'* LAYG - 01/11/2001
'****************************
Option Explicit
Dim mbTrans As Boolean
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018

Public Function ObtieneTablaCalif(ByVal psServer As String) As ADODB.Recordset
Dim Sql As String
Dim Co As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Set Co = New COMConecta.DCOMConecta
Sql = "Select nCalCodTab, cCalDes, nRangIni, nRangfin, cCalif, nProvision, cRefinan " & _
      "From ColocCalificaTabla Order by nCalCodTab"
Co.AbreConexion
Set rs = Co.CargaRecordSet(Sql)
Co.CierraConexion
Set ObtieneTablaCalif = rs
End Function

Public Sub InsertaColocCalifTabla(ByVal psServer As String, ByVal psCalDes As String, _
 ByVal pnRangIni As Integer, ByVal pnRangFin As Integer, ByVal psCalif As String, ByVal pnProvision As Double)
Dim Co As COMConecta.DCOMConecta
Dim SQL1 As String
Dim Sql2 As String
Dim rs As New ADODB.Recordset
Dim nNro As Integer
Set Co = New COMConecta.DCOMConecta
SQL1 = "Select Isnull(max(nCalCodTab),0) Nro From ColocCalificatabla "
Co.AbreConexion
Set rs = Co.CargaRecordSet(SQL1)
nNro = nNro + 1
Sql2 = " Insert ColocCalificaTabla (nCalCodTab, cCalDes, nRangIni, nRangfin, cCalif, nProvision) " & _
       "  values (" & nNro & ",'" & psCalDes & "'," & pnRangIni & "," & _
       " " & pnRangFin & ",'" & psCalif & "'," & pnProvision & ")"
Co.Ejecutar (Sql2)
Co.CierraConexion
Set rs = Nothing
Set Co = Nothing
End Sub

Public Sub UpdateColocCalifTabla(ByVal psServer As String, ByVal psCalDes As String, _
 ByVal pnRangIni As Integer, ByVal pnRangFin As Integer, ByVal psCalif As String, _
 ByVal pnProvision As Double, ByVal pnCod As Integer)
Dim Co As COMConecta.DCOMConecta
Set Co = New COMConecta.DCOMConecta
Dim Sql As String
Sql = " Update ColocCalificaTabla  " & _
      " Set  cCalDes='" & psCalDes & "', nRangIni=" & pnRangIni & ", " & _
      " nRangfin=" & pnRangFin & ", cCalif='" & psCalif & "', nProvision=" & pnProvision & "  " & _
      " Where  nCalCodTab = " & pnCod
Co.AbreConexion
Co.Ejecutar (Sql)
Co.CierraConexion
Set Co = Nothing
End Sub

Public Function nObtieneCreditosCalificados(Optional ByRef psMensaje As String = "") As Recordset

'************************************

Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset


    lsSQL = "SELECT   Per.CPersNombre Persona, CEval.cEvalCalif Cal, " _
        & " DocNat = (Select ISNULL(PerID.cPersIDnro,'') From PersId PerID " _
        & "          Where PerId.cPersCod = Per.cPersCod AND PerId.cPersIDTpo ='" & gPersIdDNI & "') ,  " _
        & " DocTri =(Select ISNULL(PerID.cPersIDnro,'') From PersId PerID " _
        & "          Where PerId.cPersCod = Per.cPersCod AND PerId.cPersIDTpo ='" & gPersIdRUC & "') ,  " _
        & " Per.cPersCod CodPers, ISNULL(CEval.cEvalObs,'') AS Obs " _
        & " FROM ColocEvalCalif CEval JOIN Persona Per ON Per.cPersCod = CEval.cPersCod " _
        & " ORDER BY Per.cPersNombre "

    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        psMensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "ERROR: No existen datos "
        Exit Function
    End If
    
    Set nObtieneCreditosCalificados = lrDatos
    Set lrDatos = Nothing
End Function
Public Function nobtineEvalconsulta(ByVal psCodPers As String, _
                                ByVal opt As Integer, ByVal pdFecha As Date, ByVal psServerConsol As String, _
                                Optional ByRef psMensaje As String = "") As ADODB.Recordset
Dim lsSQL As String
Dim lsCad As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
lsCad = IIf(opt = 0, " and CED.cPersCod='" & psCodPers & "' ", "  ")

lsSQL = "Select CED.cPersCod,"
lsSQL = lsSQL & " (Select cPersNombre From Persona Where cPersCod=CED.cPersCod) Nombre,"
lsSQL = lsSQL & " C.cCalGen , CED.cCtaCod, C.nSaldoCap, C.nDiasAtraso"
lsSQL = lsSQL & " From ColocEvalCalifDetalle CED"
lsSQL = lsSQL & " Inner Join ColocCalifProv C on C.cCtaCod=CED.cCtaCod "
lsSQL = lsSQL & " Where CED.dEval='" & Format(pdFecha, "mm/dd/yyyy") & "' "
lsSQL = lsSQL & lsCad
lsSQL = lsSQL & " Order by CED.cPersCod"
    

    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        psMensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    Set nobtineEvalconsulta = lrDatos
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "ERROR: No existen datos "
        Set lrDatos = Nothing
        Exit Function
    End If
    Set lrDatos = Nothing

End Function

Public Function nVerifExisteEvaluacion(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
                                    Optional ByRef psMensaje As String = "") As Boolean
Dim lsSQL As String
Dim loReg As COMDColocPig.DCOMColPFunciones
Dim lrExiste As ADODB.Recordset

    lsSQL = "SELECT cPersCod FROM ColocEvalCalif " _
          & "WHERE cPersCod ='" & psPersCod & "' And nEvalTipo = " & pnEvalTipo & " "
    
    Set loReg = New COMDColocPig.DCOMColPFunciones
    Set lrExiste = loReg.dObtieneRecordSet(lsSQL)
    If lrExiste Is Nothing Then
        psMensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    
    If lrExiste.BOF And lrExiste.EOF Then
        nVerifExisteEvaluacion = False
    Else
        nVerifExisteEvaluacion = True
    End If
    
    Set lrExiste = Nothing
End Function

Public Function nObtieneDatosClienteCreditos(ByVal psCodPers As String, ByVal pdFechaFinMes As Date, _
                                            Optional ByRef psMensaje As String = "") As Recordset
'************************************

Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset

'lsSQL = "SELECT P.cCtaCod, Per.cPersCod, PP.nPrdPersRelac, P.nPrdEstado, C.nMontoCol, " _
        & "     C.cLineaCred, C.dVigencia, P.nSaldo,  " _
        & "     cCodAnalista = (SELECT ISNULL(rh.cUser,'') From ProductoPersona PPer " _
        & "                       JOIN RRHH rh ON PPer.cPersCod = rh.cPersCod " _
        & "                       Where PPer.cCtaCod = P.cCtaCod And PPer.nPrdPersRelac = " & 28 & " ), " _
        & "     cNota1 = (Select ISNULL(ccAna.nColocNota,'') FROM ColocCalificacionAnalista ccAna " _
        & "                Where ccAna.cCtaCod = P.cCtaCod And " _
        & "                ccAna.dColocNotaFecha = (Select Max(dColocNotaFecha) From ColocCalificacionAnalista " _
        & "                                         Where cCtaCod = P.cCtaCod ) ) , " _
        & "     cNomAgencia = (Select ISNULL(cAgeDescripcion,'') FROM Agencias Where cAgeCod = substring(P.cCtaCod,4,2) ),  " _
        & "     (  SELECT CS.nDiasAtraso  " _
        & "        FROM ColocacSaldo CS " _
        & "        WHERE CS.dFecha = (Select  Max(dFecha) From ColocacSaldo CS1 WHERE CS1.CCtaCod =P.CCtaCod AND CS1.dFecha <='" & Format(pdFechaFinMes, "mm/dd/yyyy") & "') " _
        & "        AND CS.cCtaCod=P.CCtaCod ) AS  AtrasoFM , " _
        & "     (  SELECT CS.nSaldoCap  " _
        & "        FROM ColocacSaldo CS " _
        & "        WHERE CS.dFecha = (Select  Max(dFecha) From ColocacSaldo CS1 WHERE CS1.CCtaCod =P.CCtaCod AND CS1.dFecha <='" & Format(pdFechaFinMes, "mm/dd/yyyy") & "') " _
        & "        AND CS.cCtaCod=P.CCtaCod ) AS  SaldoFM , "
'lsSQL = lsSQL & " " _
        & "     (  SELECT CS.cNota1  " _
        & "        FROM ColocacSaldo CS " _
        & "        WHERE CS.dFecha = (Select  Max(dFecha) From ColocacSaldo CS1 WHERE CS1.CCtaCod =P.CCtaCod AND CS1.dFecha <='" & Format(pdFechaFinMes, "mm/dd/yyyy") & "') " _
        & "        AND CS.cCtaCod=P.CCtaCod ) AS  Nota1FM , " _
        & "     (  SELECT CS.dFecha  " _
        & "        FROM ColocacSaldo CS " _
        & "        WHERE CS.dFecha = (Select  Max(dFecha) From ColocacSaldo CS1 WHERE CS1.CCtaCod =P.CCtaCod AND CS1.dFecha <='" & Format(pdFechaFinMes, "mm/dd/yyyy") & "') " _
        & "        AND CS.cCtaCod=P.CCtaCod ) AS dFecha  " _
        & " FROM ProductoPersona PP " _
        & "     JOIN Producto P ON PP.cCtaCod = P.cCtaCod " _
        & "     JOIN Colocaciones C ON C.cCtaCod = P.cCtaCod " _
        & "     JOIN Persona Per ON Per.cPersCod = PP.cPersCod " _
        & " WHERE Per.cPersCod = '" & Trim(psCodPers) & "' " _
        & "  "
'ICA layg 2004/10/15
lsSQL = "SELECT P.cCtaCod, Per.cPersCod, PP.nPrdPersRelac, P.nPrdEstado, C.nMontoCol, " _
        & "     C.cLineaCred, C.dVigencia, P.nSaldo,  " _
        & "     cCodAnalista = (SELECT ISNULL(rh.cUser,'') From ProductoPersona PPer " _
        & "                       JOIN RRHH rh ON PPer.cPersCod = rh.cPersCod " _
        & "                       Where PPer.cCtaCod = P.cCtaCod And PPer.nPrdPersRelac = " & 28 & " ), " _
        & "     cNota1 = (Select ISNULL(ccAna.nColocNota,'') FROM ColocCalificacionAnalista ccAna " _
        & "                Where ccAna.cCtaCod = P.cCtaCod And " _
        & "                ccAna.dColocNotaFecha = (Select Max(dColocNotaFecha) From ColocCalificacionAnalista " _
        & "                                         Where cCtaCod = P.cCtaCod ) ) , " _
        & "     cNomAgencia = (Select ISNULL(cAgeDescripcion,'') FROM Agencias Where cAgeCod = substring(P.cCtaCod,4,2) ),  " _
        & "     CS.nDiasAtraso as AtrasoFM, CS.nSaldoCap as SaldoFM, CS.cNota1 as cNota1FM ,CS.dFecha as dFecha  " _
        & " FROM ProductoPersona PP " _
        & "     JOIN Producto P ON PP.cCtaCod = P.cCtaCod " _
        & "     JOIN Colocaciones C ON C.cCtaCod = P.cCtaCod " _
        & "     JOIN Persona Per ON Per.cPersCod = PP.cPersCod " _
        & "     LEFT JOIN " _
        & "       (SELECT  CS.CCTACOD, CS.nSaldoCap, CS.CNOTA1,CS.DFECHA, CS.NDIASATRASO " _
        & "        FROM ColocacSaldo CS " _
        & "        WHERE CS.dFecha = (Select  Max(dFecha) From ColocacSaldo CS1 WHERE CS1.dFecha <='" & Format(pdFechaFinMes, "mm/dd/yyyy") & "') " _
        & "        Group BY CS.CCTACOD, CS.nSaldoCap, CS.CNOTA1,CS.DFECHA, CS.NDIASATRASO )    CS " _
        & "     ON CS.CCTACOD = P.CCTACOD " _
        & " WHERE Per.cPersCod = '" & Trim(psCodPers) & "' ORDER BY C.dVigencia "
        

    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        psMensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "ERROR: No existen datos "
        Exit Function
    End If
    
    Set nObtieneDatosClienteCreditos = lrDatos
    Set lrDatos = Nothing
End Function

Public Function nObtieneDatosClienteGarantias(ByVal psCodCta As String, _
                                              Optional ByRef psMensaje As String = "") As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset

lsSQL = "SELECT   GC.cCtaCod, G.cTpoDoc, G.cNroDoc, P.cPersCod , G.cNumGarant,  " _
    & " P.cPersNombre, PE1.nPrdPersRelac,  TC.cConsDescripcion as TipoGarantia,  " _
    & " G.nTpoGarantia, Convert(VarChar(30),ISNULL(G.cDescripcion,'')) as  Descripcion, " _
    & " G.cNroDoc,G.nMoneda as Moneda, ISNULL(G.nTasacion,0) AS TASACION, " _
    & " ISNULL(G.nRealizacion,0) AS REALIZACION , ISNULL(G.nPorGravar,0) AS PORGRAVAR, " _
    & " G.nEstado,  ISNULL(GC.nGravado,0) as Gravado,  GC.nMoneda as MonedaGC " _
    & " FROM  ColocGarantia GC    JOIN Garantias G ON G.cNumGarant  =GC.cNumGarant  " _
    & " JOIN PersGarantia PG        ON PG.cNumGarant  = GC.cNumGarant " _
    & " JOIN Persona P        ON P.cPersCod = PG.cPersCod   " _
    & " JOIN ProductoPersona PE1  ON (PE1.cPersCod=PG.cPersCod AND PE1.cCtaCod =GC.cCtaCod) " _
    & " JOIN Constante TC      ON TC.nConsValor = G.nTpoGarantia  " _
    & " WHERE   TC.nConsCod = " & gPersGarantia & " " _
    & " AND GC.cCtaCod='" & psCodCta & "' ORDER BY GC.cCtaCod "
    
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        psMensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "ERROR: No existen datos de Garantia de Credito " & psCodCta
        Exit Function
    End If
    
    Set nObtieneDatosClienteGarantias = lrDatos
    Set lrDatos = Nothing

End Function

Public Function nObtieneDatosFuentesIngreso(ByVal psCodCta As String, _
                                           Optional ByRef psMensaje As String = "") As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset

lsSQL = "SELECT  'TipoFte'= Case When F.nPersTipFte = 1 THEN 'Dependiente' ELSE 'Independiente' END , " _
    & " F.cNumFuente, ISNULL(F.cRazSocDescrip,'') AS cRazonSocial, ISNULL(F.cRazSocDirecc,'') AS cDireccion, " _
    & " ISNULL(Z.cUbiGeoDescripcion,'NI') AS ZONA, ISNULL(F.cRazSocUbiGeo,'') AS cCodZon,  " _
    & " ISNULL(F.cPersFICargo,'') as Cargo  " _
    & "  " _
    & "  " _
    & " FROM  PersFteIngreso F JOIN Persona P  ON P.cPersCod = F.cPersCod  " _
    & " LEFT JOIN PersFIDependiente FD ON FD.cNumFuente  = F.cNumFuente " _
    & " LEFT JOIN PersFIindependiente FI ON FI.cNumFuente = F.cNumFuente  " _
    & " LEFT JOIN UbicacionGeografica Z ON Z.cUbiGeoCod = F.cRazSocUbiGeo  " _
    & "  " _
    & " WHERE P.cPersCod ='" & psCodCta & "'  " _
    & " ORDER BY F.cNumFuente "
    
    
'SQL = "  SELECT  'TIPOFTE'=   CASE " _
'        & "                     WHEN F.cTipoFuente='D' THEN 'DEPENDIENTE' " _
'        & "                     Else 'INDEPENDIENTE' " _
'        & "                 END, " _
'        & "         ISNULL(F.cRazonSocial,'') AS cRazonSocial, ISNULL(F.cdireccion,'') AS cdireccion, ISNULL(Z.cDesZon,'NI') AS ZONA, ISNULL(F.ccodzon,'') AS cCodZon, " _
'        & "         ISNULL(S.CNOMTAB,'NI') as Sector, ISNULL(F.cSector,'') AS cSector, ISNULL(A.CNOMTAB,'NI') AS ACTIVIDAD,  ISNULL(F.cActeCon,'') AS cActEcon, ISNULL(F.ccargo,'') as Cargo, " _
'        & "         FD.dFecEval AS FDFECHA, ISNULL(FD.nIngClt,0) AS INGRESO , ISNULL(FD.nGasFam,0) AS GASTOS, " _
'        & "         B.dFecBalanc AS BALFECHA, ISNULL(B.ningfam,0) AS BALINGFAM ,ISNULL(B.ngasfam,0) AS BALGASFAM,F.CNUMFUENTE " _
'        & " FROM    FUENTEINGRESO F  JOIN " & gcCentralPers & "Persona P ON F.CCODPERS=P.CCODPERS " _
'        & "         LEFT JOIN FDEPENDIENTE FD   ON FD.CNUMFUENTE=F.CNUMFUENTE " _
'        & "         LEFT JOIN BALANCE B     ON B.CNUMFUENTE=F.CNUMFUENTE " _
'        & "         LEFT JOIN " & gcCentralCom & "Zonas Z       ON Z.CCODZON=F.CCODZON " _
'        & "         LEFT JOIN " & gcCentralCom & "TablaCod S         ON S.CVALOR = F.CSECTOR " _
'        & "         LEFT JOIN " & gcCentralCom & "TablaCod A         ON A.CVALOR = F.CACTECON " _
'        & " WHERE   P.CCODPERS='" & lsCodPers & "' AND S.CCODTAB LIKE '20__' " _
'        & "         AND A.CCODTAB LIKE '35__' ORDER BY  F.CNUMFUENTE "
    
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        psMensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "ERROR: No existen datos "
        Exit Function
    End If
    
    Set nObtieneDatosFuentesIngreso = lrDatos
    Set lrDatos = Nothing

End Function

'************************************
Public Function nObtieneCreditosEvaluadosPersDetalles(ByVal psCodPers As String, ByVal pnEvalTipo As Integer) As Recordset

Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset


    lsSQL = "SELECT cCtacod as Credito, dEval as Fecha,  " _
        & " IsNull(cEvalCalifDet,'') as Cal, IsNull(cEvalObsDet,'') Obs , " _
        & " Isnull(nSaldoCap,0) as nSaldoCap, Isnull(nDiasAtraso,0) nDiasAtraso " _
        & " FROM ColocEvalCalifDetalle  " _
        & " WHERE cPersCod ='" & psCodPers & "' and nEvalTipo = " & pnEvalTipo _
        & " ORDER BY dEval, cCtaCod "
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        Exit Function
    End If
    
    Set nObtieneCreditosEvaluadosPersDetalles = lrDatos
    Set lrDatos = Nothing
End Function

Public Function nCalificaCabecera(psPersCod As String, Optional psFecCal As String = "@") As String
Dim lsSQL As String
Dim loReg As ADODB.Recordset
Dim loConecta As COMConecta.DCOMConecta

    lsSQL = "SELECT ISNULL(MAX(cEvalCalifDet),0) AS Maxima FROM ColocEvalCalifDetalle " _
          & "WHERE cPersCod  ='" & Trim(psPersCod) & "' "
    If psFecCal <> "@" Then
        lsSQL = lsSQL & " AND DATEDIFF(D,dEval, '" & Format(psFecCal, "mm/dd/yyyy") & "' ) = 0 "
    End If
    
    Set loConecta = New COMConecta.DCOMConecta
        loConecta.AbreConexion
        Set loReg = loConecta.CargaRecordSet(lsSQL)
        loConecta.CierraConexion
    Set loConecta = Nothing
    
    nCalificaCabecera = IIf(IsNull(loReg!Maxima), 0, Trim(loReg!Maxima))
    loReg.Close
    Set loReg = Nothing

End Function

Public Sub nCalifPersonaNuevo(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        ByVal psEvalCalif As String, ByVal psUltimaAct As String, ByVal psEvalObs As String, _
        pmCalifDetalle As Variant, Optional pbEjecBatch As Boolean = False)

'** Insert ColocEvalCalif
'** Insert ColocEvalCalifDetalle
'************************************

Dim lsSQL As String
Dim loReg As COMDCredito.DCOMColocEval
Dim lnMovNro As Long
Dim lsCalifGen As String

'On Error GoTo ErrorCalif
Set loReg = New COMDCredito.DCOMColocEval
    'loReg.dBeginTrans
    mbTrans = True
    
    '** Inserta ColocEvalCalif
    Call loReg.dInsertColocEvalCalif(psPersCod, pnEvalTipo, psEvalCalif, psUltimaAct, psEvalObs, False)
    
    '** Obtiene la Maxima Calificacion
    lsCalifGen = nCalificaCabecera(psPersCod)
    '** Update ColocEvalCalif ( con la Maxima Calificacion)
    Call loReg.dUpdateColocEvalCalif(psPersCod, pnEvalTipo, lsCalifGen, psUltimaAct, , pbEjecBatch)
    
    'mnEjecutaBatch = loReg.dEjecutaBatch
    'loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ErrorCalif:
    If mbTrans Then
        'loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCredPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nCalifPersonaModifica(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        ByVal psEvalCalif As String, ByVal psUltimaAct As String, ByVal psEvalObs As String, _
        Optional pbEjecBatch As Boolean = False)

'** Insert ColocEvalCalif
'** Insert ColocEvalCalifDetalle
'************************************

Dim lsSQL As String
Dim loReg As COMDCredito.DCOMColocEval
Dim lnMovNro As Long
Dim lsCalifGen As String

'On Error GoTo ErrorCalif
Set loReg = New COMDCredito.DCOMColocEval
    'loReg.dBeginTrans
    mbTrans = True
    
    '** Obtiene la Maxima Calificacion
    lsCalifGen = nCalificaCabecera(psPersCod)
    '** Update ColocEvalCalif
    Call loReg.dUpdateColocEvalCalif(psPersCod, pnEvalTipo, lsCalifGen, psUltimaAct, psEvalObs, False)
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    'loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ErrorCalif:
    If mbTrans Then
        'loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCredPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nCalifPersonaElimina(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        ByVal psUltimaAct As String, Optional pbEjecBatch As Boolean = False)

'** Insert ColocEvalCalif
'** Insert ColocEvalCalifDetalle
'************************************

Dim lsSQL As String
Dim loReg As COMDCredito.DCOMColocEval
Dim lnMovNro As Long
Dim lsCalifGen As String

'On Error GoTo ErrorCalif
Set loReg = New COMDCredito.DCOMColocEval
    'loReg.dBeginTrans
    mbTrans = True
    
    '** Delete ColocEvalCalif
    Call loReg.dDeleteColocEvalCalif(psPersCod, pnEvalTipo, False)
    
    '** Obtiene la Maxima Calificacion
    lsCalifGen = nCalificaCabecera(psPersCod)
    '** Update ColocEvalCalif ( con la Maxima Calificacion)
    Call loReg.dUpdateColocEvalCalif(psPersCod, pnEvalTipo, lsCalifGen, psUltimaAct, , pbEjecBatch)
    
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    'loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ErrorCalif:
    If mbTrans Then
        'loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCredPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nCalifDetalleNuevo(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        ByVal psCtaCod As String, ByVal pdEval As Date, ByVal psEvalEstado As String, _
        ByVal psEvalCalifDet As String, ByVal pnSaldoCap As Double, ByVal pnDiasAtraso As Integer, _
        ByVal psUltimaAct As String, _
        ByVal psEvalObsDet As String, Optional pbEjecBatch As Boolean = False)

'** Insert ColocEvalCalifDetalle
'** Update ColocEvalCalif
'************************************

Dim loReg As COMDCredito.DCOMColocEval
Dim lnMovNro As Long
Dim lsCalifGen As String
'On Error GoTo ErrorCalif
Set loReg = New COMDCredito.DCOMColocEval
    'loReg.dBeginTrans
    mbTrans = True
    
    '** Inserta ColocEvalCalifDetalle
    Call loReg.dInsertColocEvalCalifDetalle(psPersCod, pnEvalTipo, psCtaCod, pdEval, _
            "", psEvalCalifDet, pnSaldoCap, pnDiasAtraso, psUltimaAct, psEvalObsDet, False)
    '** Obtiene la Maxima Calificacion
    lsCalifGen = nCalificaCabecera(psPersCod)
    '** Update ColocEvalCalif ( con la Maxima Calificacion)
    Call loReg.dUpdateColocEvalCalif(psPersCod, pnEvalTipo, lsCalifGen, psUltimaAct, , pbEjecBatch)
            
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    'loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ErrorCalif:
    If mbTrans Then
        'loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCredPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Sub nCalifDetalleModifica(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        ByVal psCtaCod As String, ByVal pdEval As Date, ByVal psEvalEstado As String, _
        ByVal psEvalCalifDet As String, ByVal pnSaldoCap As Double, ByVal pnDiasAtraso As Integer, _
        ByVal psUltimaAct As String, ByVal psEvalObsDet As String, Optional pbEjecBatch As Boolean = False)

'** Update ColocEvalCalifDetalle
'** Update ColocEvalCalif
'************************************

Dim lsSQL As String
Dim loReg As COMDCredito.DCOMColocEval
Dim lnMovNro As Long
Dim lsCalifGen As String
On Error GoTo ErrorCalif
Set loReg = New COMDCredito.DCOMColocEval
    'loReg.dBeginTrans
    mbTrans = True
    
    '** Update ColocEvalCalifDetalle
    Call loReg.dUpdateColocEvalCalifDetalle(psPersCod, pnEvalTipo, psCtaCod, Format(pdEval, "yyyy/mm/dd"), _
             "", psEvalCalifDet, pnSaldoCap, pnDiasAtraso, psUltimaAct, psEvalObsDet, False)
    '** Obtiene la Maxima Calificacion
    lsCalifGen = nCalificaCabecera(psPersCod)
    '** Update ColocEvalCalif ( con la Maxima Calificacion)
    Call loReg.dUpdateColocEvalCalif(psPersCod, pnEvalTipo, lsCalifGen, psUltimaAct, , pbEjecBatch)
            
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    'loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ErrorCalif:
    If mbTrans Then
        'loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCredPignoraticio", "Error en Funcion Modifica Calificacion Detalle "

End Sub

Public Sub nCalifDetalleElimina(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        ByVal psCtaCod As String, ByVal pdEval As Date, ByVal psEvalEstado As String, _
        ByVal psUltimaAct As String, Optional pbEjecBatch As Boolean = False)

'** Update ColocEvalCalifDetalle
'** Update ColocEvalCalif
'************************************

Dim lsSQL As String
Dim loReg As COMDCredito.DCOMColocEval
Dim lnMovNro As Long
Dim lsCalifGen As String
On Error GoTo ErrorCalif
Set loReg = New COMDCredito.DCOMColocEval
    'loReg.dBeginTrans
    mbTrans = True
    
    '** Update ColocEvalCalifDetalle
    Call loReg.dDeleteColocEvalCalifDetalle(psPersCod, pnEvalTipo, psCtaCod, pdEval, _
              psEvalEstado, False)
    '** Obtiene la Maxima Calificacion
    lsCalifGen = nCalificaCabecera(psPersCod)
    '** Update ColocEvalCalif ( con la Maxima Calificacion)
    Call loReg.dUpdateColocEvalCalif(psPersCod, pnEvalTipo, lsCalifGen, psUltimaAct, , pbEjecBatch)
            
    'mnEjecutaBatch = loRegPig.dEjecutaBatch
    'loReg.dCommitTrans
    mbTrans = False
Set loReg = Nothing

Exit Sub
ErrorCalif:
    If mbTrans Then
        'loReg.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCredPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub


Public Function nObtieneDatosEvaluacionAutomatica(ByVal psFechaCalif As Date, ByVal psFechaData As Date, _
                                                Optional ByRef psMensaje As String = "") As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset

lsSQL = " Select Apd.cPersCod, Apd.cCtaCod, Apd.nSaldoCap,  Apd.nDiasAtraso, APD.cEvalCalifDet " _
    & " From ColocEvalCalifDetalle Apd " _
    & " Where datediff(d, dEval,'" & Format(psFechaCalif, "yyyy/mm/dd") & "')= 0 " _
    & " And apd.cCtaCod not in  " _
    & " (Select cCtaCod From ColocEvalCalifDetalle " _
    & "     Where datediff(d, dEval,'" & Format(psFechaData, "yyyy/mm/dd") & "')= 0 ) "
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos para Calificacion Automatica"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos para Calificacion Automatica"
        Exit Function
    End If
    
    Set nObtieneDatosEvaluacionAutomatica = lrDatos
    Set lrDatos = Nothing

End Function

' Obtiene la calificación automatica del Credito
Public Function nGeneraCalificacionAutomatica(ByVal psCuenta As String, ByVal psRutaConsol As String, _
                                            Optional ByRef psMensaje As String = "") As String
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
Dim lsCalifxDA As String
'Dim loCalif As nColocEvalCalProceso
   
lsCalifxDA = ""
   

    'Obtengo Datos del Credito
    lsSQL = " Select c.cCtaCod, c.cRefinan, c.nDiasAtraso, c.nPrdEstado, " _
        & " (SELECT t.cCalif cAbrev FROM ColocCalificaTabla t " _
        & "    WHERE substring(ltrim(str(t.nCalCodTab)),1,3) = Substring(c.cctacod,6,1) + '0' " _
        & "             + Case When C.cRefinan ='R' then '1' else '0' end " _
        & "    AND nRangIni <= C.nDiasAtraso And  nRangFin >= C.nDiasAtraso ) CalifxTabla  " _
        & " From " & psRutaConsol & " Creditoconsol c  " _
        & " Where c.cCtaCod ='" & psCuenta & "' " _
        & " And c.nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc _
        & "," & gColocEstRefNorm & "," & gColocEstRefMor & "," & gColocEstRefVenc & "," & gColPEstRegis _
        & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & "," & gColocEstRecVigJud & ",2050,2060) "
        
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        psMensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        'MsgBox "ERROR: No existen datos En Tabla de Parametros ", vbInformation, "Aviso"
        Exit Function
    End If
    
    'Obtengo la Calificacion
    lsCalifxDA = Trim(lrDatos!CalifxTabla)
    '*** Aplico las Reglas de Negocio
    'Set loCalif = New nColocEvalCalProceso
    '    lsCalifxDA = nCalificaCorrigeCalifxDiaAtraso(lsCalifxDA, lrDatos!nPrdEstado ,)
    'Set loCalif = Nothing

    nGeneraCalificacionAutomatica = lsCalifxDA
    Set lrDatos = Nothing

End Function


'**************************************************************************
'******** DE ACA EN ADELANTE SON FUNCIONES DE LA CLASE nColocEvalCalProceso
'**************************************************************************

Public Sub nCalificaPreparaArchivoCalificacion(ByVal psServerConsol As String, _
                                    ByVal psFecha As Date, _
                                    ByVal pnTipoCambio As Double, _
                                    ByVal psMensaje As String, _
                                    ByVal pMatProducto As Variant)

'** Llena tabla ColocCalifProv (Contiene las Calificaciones de los creditos Vigentes)
Dim loConec As COMConecta.DCOMConecta
Dim loCalif As COMNCredito.NCOMColocEval
Dim lrDatos As ADODB.Recordset
Dim lsSQL As String

Dim lsCalificacion As String
Dim lsGarPrefProc As String

Dim lbGarPrefe As Boolean
'On Error GoTo ErrorConexion

'Obtener parametro
'Dim nPorcentaje As Double
'Dim oParam As New COMDCredito.DCOMParametro
'    nPorcentaje = oParam.RecuperaValorParametro(9005)
'Set opram = Nothing
'*****************************************************

'Set loCalif = New COMNCredito.NCOMColocEval
'Set lrDatos = loCalif.nCalificaLLenaTablaCalifProv(psServerConsol, psFecha, pnTipoCambio, psMensaje)
'Set lrDatos = loCalif.nCalificaLLenaTablaCalifProv(psServerConsol, psFecha, pnTipoCambio, psMensaje)
'Set loCalif = Nothing

Set lrDatos = nCalificaLLenaTablaCalifProv(psServerConsol, psFecha, pnTipoCambio, psMensaje)

If psMensaje <> "" Then
    Exit Sub
End If

If lrDatos Is Nothing Then
    psMensaje = "No existen datos"
    Exit Sub
Else
    Set loConec = New COMConecta.DCOMConecta
    loConec.AbreConexion
    'Borra los datos de ColocCalifProv
    lsSQL = "Delete ColocCalifProv "
    loConec.Ejecutar lsSQL
    
    Do While Not lrDatos.EOF
        'Aplica Reglas de Negocio
        'Set loCalif = New COMNCredito.NCOMColocEval
        'lsCalificacion = loCalif.nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado)
        'Set loCalif = Nothing
        '--
'        lsCalificacion = lrDatos!cCalDias
'        If lrDatos!cTpoProdCod = "514" Then
'            lsCalificacion = "0"
'        ElseIf lrDatos!nPrdEstado = 2030 Or lrDatos!nPrdEstado = 2031 Or lrDatos!nPrdEstado = 2032 Then 'Aplica a todos los Refinanciados
'            lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
'        'ElseIf UBound(pMatProducto) = 0 Then 'Todos los Producto
'        '    lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
'        ElseIf pMatProducto(0) = 1 And Mid(lrDatos!cCtaCod, 6, 1) = "1" Then
'            lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
'        ElseIf pMatProducto(1) = 1 And Mid(lrDatos!cCtaCod, 6, 1) = "2" Then
'            lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
'        ElseIf pMatProducto(4) = 1 And Mid(lrDatos!cCtaCod, 6, 3) = "305" Then
'            lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
'        ElseIf pMatProducto(2) = 1 And Mid(lrDatos!cCtaCod, 6, 1) = "3" Then
'            lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
'        ElseIf pMatProducto(3) = 1 And Mid(lrDatos!cCtaCod, 6, 1) = "4" Then
'            lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
'        End If
        lsCalificacion = lrDatos!cCalDias
        '**ARLO20180712 ERS042 - 2018
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("N0000098", lrDatos!cTpoProdCod) Then
        'If lrDatos!cTpoProdCod = "514" Then
        '**ARLO20180712 ERS042 - 2018
            lsCalificacion = "0"
        ElseIf lrDatos!nPrdEstado = 2030 Or lrDatos!nPrdEstado = 2031 Or lrDatos!nPrdEstado = 2032 Then
            lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
        Else
            '**ARLO20180712 ERS042 - 2018
            If objProducto.GetResultadoCondicionCatalogo("N0000099", lrDatos!cTpoProdCod) Then
            'If lrDatos!cTpoProdCod = "705" Then
            '**ARLO20180712 ERS042 - 2018
                If pMatProducto(8) = 1 Then
                    lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
                End If
            Else
                If pMatProducto(0) = 1 And Mid(lrDatos!cTpoCredCod, 1, 1) = "1" Then
                    lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
                ElseIf pMatProducto(1) = 1 And Mid(lrDatos!cTpoCredCod, 1, 1) = "2" Then
                    lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
                ElseIf pMatProducto(2) = 1 And Mid(lrDatos!cTpoCredCod, 1, 1) = "3" Then
                    lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
                ElseIf pMatProducto(3) = 1 And Mid(lrDatos!cTpoCredCod, 1, 1) = "4" Then
                    lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
                ElseIf pMatProducto(4) = 1 And Mid(lrDatos!cTpoCredCod, 1, 1) = "5" Then
                    lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
                ElseIf pMatProducto(5) = 1 And Mid(lrDatos!cTpoCredCod, 1, 1) = "6" Then
                    lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
                ElseIf pMatProducto(6) = 1 And Mid(lrDatos!cTpoCredCod, 1, 1) = "7" Then
                    lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
                ElseIf pMatProducto(7) = 1 And Mid(lrDatos!cTpoCredCod, 1, 1) = "8" Then
                    lsCalificacion = nCalificaCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado, psFecha, loConec, lrDatos!cCtaCod)
                End If
            End If
        End If

        
'        lsSQL = "INSERT INTO ColocCalifProv (cPersCod,cCtaCod, nPrdEstado, cCalNor, nSaldoCap, cRefinan, nDiasAtraso,nProvision," _
'          & " nMontoApr,cLineaCred, cCodAnalista, dFecVig, nGarPref,nGarMuyRR, nGarAutoL,nGarNOPref) " _
'          & "VALUES('" & lrDatos!cPersCod & "','" & lrDatos!cCtaCod & "'," & lrDatos!nPrdEstado & ",'" _
'          & lsCalificacion & "'," & IIf(IsNull(lrDatos!nSaldoCap), 0, lrDatos!nSaldoCap) & ",'" & lrDatos!cRefinan & "'," & lrDatos!nDiasAtraso & ",0," _
'          & IIf(IsNull(lrDatos!nMontoApr), 0, lrDatos!nMontoApr) & ",'" & lrDatos!cLineaCred & "','" & lrDatos!cCodAnalista & "','" _
'          & Format(lrDatos!dFecVig, "mm/dd/yyyy") & "'," & Format(lrDatos!nGarPref, "#.00") & "," _
'          & Format(lrDatos!nGarMuyRR, "#.00") & "," & Format(lrDatos!nGarAutoL, "#.00") & "," & Format(lrDatos!nGarNOPref, "#.00") & " )"
'ALPA 20100605 B2**************************************************************************
'      lsSQL = "INSERT INTO ColocCalifProv (cPersCod,cCtaCod, nPrdEstado, cCalNor, nSaldoCap, cRefinan, nDiasAtraso,nProvision," _
'          & " nMontoApr,cLineaCred, cCodAnalista, dFecVig, nGarPref,nGarMuyRR, nGarAutoL,nGarNOPref,nIntSusp,nIntDev,nCapVencido,cAgeCodAct,cTpoProdCod,cTpoCredCod) " _
'          & "VALUES('" & lrDatos!cPersCod & "','" & lrDatos!cCtaCod & "'," & lrDatos!nPrdEstado & ",'" _
'          & lsCalificacion & "'," & IIf(IsNull(lrDatos!nSaldoCap), 0, lrDatos!nSaldoCap) & ",'" & lrDatos!cRefinan & "'," & lrDatos!nDiasAtraso & ",0," _
'          & IIf(IsNull(lrDatos!nMontoApr), 0, lrDatos!nMontoApr) & ",'" & lrDatos!cLineaCred & "','" & lrDatos!cCodAnalista & "','" _
'          & Format(lrDatos!dFecVig, "mm/dd/yyyy") & "'," & Format(lrDatos!nGarPref, "#.00") & "," _
'          & Format(lrDatos!nGarMuyRR, "#.00") & "," & Format(lrDatos!nGarAutoL, "#.00") & "," & Format(lrDatos!nGarNOPref, "#.00") & "," & lrDatos!nIntSusp & "," & lrDatos!nIntDev & "," & lrDatos!nCapVencido & ",'" & lrDatos!cAgeCodAct & "','" & lrDatos!cTpoProdCod & "','" & lrDatos!cTpoCredCod & "')"
'       lsSQL = "INSERT INTO ColocCalifProv (cPersCod,cCtaCod, nPrdEstado, cCalNor, nSaldoCap, cRefinan, nDiasAtraso,nProvision," _
'          & " nMontoApr,cLineaCred, cCodAnalista, dFecVig, nGarPref,nGarMuyRR, nGarAutoL,nGarNOPref,nIntSusp,nIntDev,nCapVencido,cAgeCodAct,cTpoProdCod,cTpoCredCod) " _
'          & "VALUES('" & lrDatos!cPersCod & "','" & lrDatos!cCtaCod & "'," & lrDatos!nPrdEstado & ",'" _
'          & lsCalificacion & "'," & IIf(IsNull(lrDatos!nSaldoCap), 0, lrDatos!nSaldoCap) & ",'" & lrDatos!cRefinan & "'," & lrDatos!nDiasAtraso & ",0," _
'          & IIf(IsNull(lrDatos!nMontoApr), 0, lrDatos!nMontoApr) & ",'" & lrDatos!cLineaCred & "','" & lrDatos!cCodAnalista & "','" _
'          & Format(lrDatos!dFecVig, "mm/dd/yyyy") & "'," & Format(lrDatos!nGarPref, "#.00") & "," _
'          & Format(lrDatos!nGarMuyRR, "#.00") & "," & Format(lrDatos!nGarAutoL, "#.00") & "," & Format(lrDatos!nGarNOPref, "#.00") & "," & lrDatos!nIntSusp & "," & lrDatos!nIntDev & "," & lrDatos!nCapVencido & ",'" & lrDatos!cAgeCodAct & "','" & lrDatos!cTpoProdCod & "','" & lrDatos!cTpoCredCod & "')"
'lsSQL = "INSERT INTO ColocCalifProv (cPersCod,cCtaCod, nPrdEstado, cCalNor, nSaldoCap, cRefinan, nDiasAtraso,nProvision," _
'          & " nMontoApr,cLineaCred, cCodAnalista, dFecVig, nGarPref,nGarMuyRR, nGarAutoL,nGarNOPref,cTpoCredCod,nIntSusp,nIntDev,nCapVencido,cAgeCodAct) " _
'          & "VALUES('" & lrDatos!cPersCod & "','" & lrDatos!cCtaCod & "'," & lrDatos!nPrdEstado & ",'" _
'          & lsCalificacion & "'," & IIf(IsNull(lrDatos!nSaldoCap), 0, lrDatos!nSaldoCap) & ",'" & lrDatos!cRefinan & "'," & lrDatos!nDiasAtraso & ",0," _
'          & IIf(IsNull(lrDatos!nMontoApr), 0, lrDatos!nMontoApr) & ",'" & lrDatos!cLineaCred & "','" & lrDatos!cCodAnalista & "','" _
'          & Format(lrDatos!dFecVig, "mm/dd/yyyy") & "'," & Format(lrDatos!nGarPref, "#.00") & "," _
'          & Format(lrDatos!nGarMuyRR, "#.00") & "," & Format(lrDatos!nGarAutoL, "#.00") & "," & Format(lrDatos!nGarNOPref, "#.00") & ",'" & lrDatos!cTpoCredCod & "'," & lrDatos!nIntSusp & "," & lrDatos!nIntDev & "," & lrDatos!nCapVencido & ",'" & lrDatos!cAgeCodAct & "')"
'ALPA 20101231*******************************************************************************
lsSQL = "INSERT INTO ColocCalifProv (cPersCod,cCtaCod, nPrdEstado, cCalNor, nSaldoCap, cRefinan, nDiasAtraso,nProvision," _
          & " nMontoApr,cLineaCred, cCodAnalista, dFecVig, nGarPref,nGarMuyRR, nGarAutoL,nGarNOPref,cTpoCredCod,nIntSusp,nIntDev,nCapVencido,cAgeCodAct,cPersCIIU) " _
          & "VALUES('" & lrDatos!cPersCod & "','" & lrDatos!cCtaCod & "'," & lrDatos!nPrdEstado & ",'" _
          & lsCalificacion & "'," & IIf(IsNull(lrDatos!nSaldoCap), 0, lrDatos!nSaldoCap) & ",'" & lrDatos!cRefinan & "'," & lrDatos!nDiasAtraso & ",0," _
          & IIf(IsNull(lrDatos!nMontoApr), 0, lrDatos!nMontoApr) & ",'" & lrDatos!cLineaCred & "','" & lrDatos!cCodAnalista & "','" _
          & Format(lrDatos!dFecVig, "mm/dd/yyyy") & "'," & Format(lrDatos!nGarPref, "#.00") & "," _
          & Format(lrDatos!nGarMuyRR, "#.00") & "," & Format(lrDatos!nGarAutoL, "#.00") & "," & Format(lrDatos!nGarNOPref, "#.00") & ",'" & lrDatos!cTpoCredCod & "'," & lrDatos!nIntSusp & "," & lrDatos!nIntDev & "," & lrDatos!nCapVencido & ",'" & lrDatos!cAgeCodAct & "','" & lrDatos!cPersCIIU & "')"

'********************************************************************************************
        loConec.Ejecutar lsSQL
        lrDatos.MoveNext
    Loop
    Set loConec = Nothing
End If
lrDatos.Close
Set lrDatos = Nothing
Set loCalif = Nothing

Exit Sub

ErrorPreparaArchivoCalificacion:
    Err.Raise Err.Number, "Prepara Archivo Calificacion", Err.Description
    
End Sub

Public Function nCalificaLLenaTablaCalifProv(ByVal psRutaServer As String, _
        Optional ByVal psFechaData As String = "", Optional ByVal pnTipoCambio As Double = 0, _
        Optional ByRef psMensaje As String = "") As Recordset

'************************************

Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
Dim lsRFAEstadoVigentes As String
Dim lsRFAEstadoJudicial As String
lsRFAEstadoVigentes = "2060,2061,2062"
lsRFAEstadoJudicial = "2261,2265"
'ALPA20160118**********************************
'lsSQL = "exec stp_sel_ObtenerListadoCreditosGarantias " & pnTipoCambio
'Set loReg = New COMConecta.DCOMConecta
'loReg.AbreConexion
'loReg.Ejecutar (lsSQL)

''ALPA20160118**********************************
''#TGarantias
'
'    '*********  Creditos
''    lsSQL = "SELECT * FROM (    SELECT  C.cCtaCod, C.nSaldoCap, C.nPrdEstado, C.cRefinan, C.nDiasAtraso,  C.nMontoApr, C.cLineaCred , C.cCodAnalista, C.dFecVig, " _
''         & " cPersCod = (Select Max(cPersCod) From " & psRutaServer & "ProductoPersonaConsol PE " _
''         & "             Where PE.cCtaCod=C.cCtaCod And PE.nPrdPersRelac = " & gColRelPersTitular & "), " _
''         & "cCalDias = (SELECT t.cCalif cAbrev FROM ColocCalificaTabla t    " _
''         & "            WHERE substring(ltrim(str(t.nCalCodTab)),1,3) = Substring(c.cctacod,6,1) + '0' " _
''         & "             + Case When C.cRefinan ='R' then '1' else '0' end " _
''         & "            AND nRangIni <= C.nDiasAtraso And  nRangFin >= C.nDiasAtraso), " _
''         & "nGarPref =( Select Isnull(Sum(case when nMoneda =1 then nHipoteca + nVehicular  + nIndustrial + nAgricola " _
''         & "                                   when nMoneda = 2 then (nHipoteca + nVehicular + nIndustrial + nAgricola ) * " & pnTipoCambio & " " _
''         & "             end),0) " _
''         & "            From " & psRutaServer & "CredGarantias cg " _
''         & "            Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "' ), " _
''         & "nGarMuyRR = 0, " _
''         & "nGarAutoL = ( Select Isnull(sum(case when nMoneda = 1 then nDeposito  " _
''         & "                                     when nMoneda = 2 then (nDeposito )  * " & pnTipoCambio & " " _
''         & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "'), " _
''         & "nGarNOPref = ( Select Isnull(sum(case when nMoneda = 1 then nOtrosG + nJoyas " _
''         & "                                     when nMoneda = 2 then (nOtrosG + nJoyas )  * " & pnTipoCambio & " " _
''         & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "') " _
''         & " FROM  " & psRutaServer & "CreditoConsol C   JOIN " & psRutaServer & "CREDITOSALDOCONSOL CS ON CS.cCtaCod =C.cCtaCod " _
''         & " Where C.nPrdEstado in ( " & gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & "," _
''         & gColocEstRefNorm & "," & gColocEstRefMor & "," & gColocEstRefVenc & "," _
''         & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & "," & lsRFAEstadoVigentes & ") " _
''         & " And Substring(C.cCtaCod,7,2) <> '21' " _
''         & " AND DateDiff(dd,CS.dFecha,'" & Format(psFechaData, "mm/dd/yyyy") & "') = 0 "
'        'ALPA 20100605 B2****************************************************************
'        'lsSQL = "SELECT * FROM (    SELECT  C.cCtaCod, C.nSaldoCap, C.nPrdEstado, C.cRefinan, C.nDiasAtraso,  C.nMontoApr, C.cLineaCred , C.cCodAnalista, C.dFecVig, "
'        'lsSQL = "SELECT * FROM (    "
'        lsSQL = "SELECT X.*,P.cPersCIIU FROM (    " 'EJVG20140908
'        lsSQL = lsSQL & " SELECT  CO.cAgeCodAct,CO.cTpoCredCod,CO.cTpoProdCod,C.cCtaCod, C.nSaldoCap, C.nPrdEstado, C.cRefinan, C.nDiasAtraso,  C.nMontoApr, C.cLineaCred , C.cCodAnalista, C.dFecVig, "
'        '********************************************************************************
'        lsSQL = lsSQL & " cPersCod = (Select Max(cPersCod) From " & psRutaServer & "ProductoPersonaConsol PE "
'        lsSQL = lsSQL & "             Where PE.cCtaCod=C.cCtaCod And PE.nPrdPersRelac = " & gColRelPersTitular & "), "
'        lsSQL = lsSQL & "cCalDias = (SELECT t.cCalif cAbrev FROM ColocCalificaTabla t    "
'        lsSQL = lsSQL & "            WHERE substring(ltrim(str(t.nCalCodTab)),1,3) = Substring(CO.cTpoCredCod,1,1) + '0' "
'        lsSQL = lsSQL & "             + Case When C.cRefinan ='R' then '1' else '0' end "
'        lsSQL = lsSQL & "            AND nRangIni <= C.nDiasAtraso And  nRangFin >= C.nDiasAtraso), "
''        lsSQL = lsSQL & "nGarPref =( Select Isnull(Sum(case when nMoneda =1 then nHipoteca + nVehicular  + nIndustrial + nAgricola "
''        lsSQL = lsSQL & "                                   when nMoneda = 2 then (nHipoteca + nVehicular + nIndustrial + nAgricola ) * " & pnTipoCambio & " "
''        lsSQL = lsSQL & "             end),0) "
''        lsSQL = lsSQL & "            From " & psRutaServer & "CredGarantias cg "
''        lsSQL = lsSQL & "            Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "' ), "
''        lsSQL = lsSQL & "nGarMuyRR = 0, "
''        lsSQL = lsSQL & "nGarAutoL = ( Select Isnull(sum(case when nMoneda = 1 then nDeposito  "
''        lsSQL = lsSQL & "                                     when nMoneda = 2 then (nDeposito )  * " & pnTipoCambio & " "
''        lsSQL = lsSQL & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "'), "
''        lsSQL = lsSQL & "nGarNOPref = ( Select Isnull(sum(case when nMoneda = 1 then nOtrosG + nJoyas "
''        lsSQL = lsSQL & "                                     when nMoneda = 2 then (nOtrosG + nJoyas )  * " & pnTipoCambio & " "
''        lsSQL = lsSQL & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "') "
'        lsSQL = lsSQL & "nGarPref = IsNull(TE.nSuma_1,0), "
'        lsSQL = lsSQL & "nGarMuyRR = 0, "
'        lsSQL = lsSQL & "nGarAutoL =IsNull(TE.nSuma_2,0), "
'        lsSQL = lsSQL & "nGarNOPref =IsNull(TE.nSuma_3,0) "
'        'ALPA 20100614 B2****************************************************************************
'        'ALPA 20100614 B2****************************************************************************
'        lsSQL = lsSQL & " ,C.nIntSusp,C.nIntDev,isnull(C.nCapVencido,0) nCapVencido"
'        '********************************************************************************************
'        lsSQL = lsSQL & " FROM  " & psRutaServer & "CreditoConsol C   JOIN " & psRutaServer & "CREDITOSALDOCONSOL CS ON CS.cCtaCod =C.cCtaCod "
'        'ALPA 20100605 B2****************************************************************************
'        lsSQL = lsSQL & " INNER JOIN Colocaciones  CO ON C.cCtaCod =CO.cCtaCod "
'        lsSQL = lsSQL & " LEFT JOIN ##TGarantias  TE ON TE.cCtaCod =CO.cCtaCod "
'        '********************************************************************************************
'        lsSQL = lsSQL & " Where C.nPrdEstado in ( " & gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & ","
'        lsSQL = lsSQL & gColocEstRefNorm & "," & gColocEstRefMor & "," & gColocEstRefVenc & ","
'        lsSQL = lsSQL & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & "," & lsRFAEstadoVigentes & ") "
'        lsSQL = lsSQL & " And Substring(CO.cTpoCredCod,2,2) <> '81' "
'        lsSQL = lsSQL & " AND DateDiff(dd,CS.dFecha,'" & Format(psFechaData, "mm/dd/yyyy") & "') = 0 "
'
'        'Recuperaciones
''        lsSQL = lsSQL & " Union " _
''         & "SELECT  C.cCtaCod, C.nSaldoCap, C.nPrdEstado, C.cRefinan, C.nDiasAtraso, " _
''         & " C.nMontoApr, C.cLineaCred , C.cCodAnalista, C.dFecVig, " _
''         & " cPersCod = (Select Max(cPersCod) From " & psRutaServer & "ProductoPersonaConsol PE " _
''         & "             Where PE.cCtaCod=C.cCtaCod And PE.nPrdPersRelac = " & gColRelPersTitular & "), " _
''         & "cCalDias = (SELECT t.cCalif cAbrev FROM ColocCalificaTabla t        " _
''         & "            WHERE substring(ltrim(str(t.nCalCodTab)),1,3) = Substring(c.cctacod,6,1) + '0' " _
''         & "             + Case When C.cRefinan ='R' then '1' else '0' end " _
''         & "            AND nRangIni <= C.nDiasAtraso And  nRangFin >= C.nDiasAtraso), " _
''         & "nGarPref =( Select Isnull(Sum(case when nMoneda =1 then nHipoteca + nVehicular + + nIndustrial + nAgricola " _
''         & "                                   when nMoneda = 2 then (nHipoteca + nVehicular + nIndustrial + nAgricola ) * " & pnTipoCambio & " " _
''         & "             end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "'), " _
''         & "nGarMuyRR = 0, " _
''         & "nGarAutoL = ( Select Isnull(sum(case when nmoneda = 1 then nCartaFianza + nDeposito " _
''         & "                                     when nmoneda = 2 then (nCartaFianza + nDeposito )  * " & pnTipoCambio & " " _
''         & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "' ) ," _
''         & "nGarNOPref = ( Select Isnull(sum(case when nMoneda = 1 then nOtrosG + nJoyas " _
''         & "                                     when nMoneda = 2 then (nOtrosG + nJoyas)  * " & pnTipoCambio & " " _
''         & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "') " _
''         & " FROM  " & psRutaServer & "CreditoConsol C  " _
''         & " JOIN " & psRutaServer & "CREDITOSALDOCONSOL CS ON CS.cCtaCod =C.cCtaCod " _
''         & " Where C.nPrdEstado in ( " & gColocEstRecVigJud & "," & gColocEstRecVigCast & ",2205,2206," & lsRFAEstadoJudicial & " )" _
''         & " AND DateDiff(dd,CS.dFecha,'" & Format(psFechaData, "mm/dd/yyyy") & "') = 0 "
'        lsSQL = lsSQL & " Union "
'        'ALPA 20100605 B2*****************************************************************************
'        lsSQL = lsSQL & " SELECT  CO.cAgeCodAct,CO.cTpoCredCod,CO.cTpoProdCod,C.cCtaCod, C.nSaldoCap, C.nPrdEstado, C.cRefinan, C.nDiasAtraso,  "
'        'lsSQL = lsSQL & "SELECT C.cCtaCod, C.nSaldoCap, C.nPrdEstado, C.cRefinan, C.nDiasAtraso, "
'        '*********************************************************************************************
'        lsSQL = lsSQL & " C.nMontoApr, C.cLineaCred , C.cCodAnalista, C.dFecVig, "
'        lsSQL = lsSQL & " cPersCod = (Select Max(cPersCod) From " & psRutaServer & "ProductoPersonaConsol PE "
'        lsSQL = lsSQL & "             Where PE.cCtaCod=C.cCtaCod And PE.nPrdPersRelac = " & gColRelPersTitular & "), "
'        lsSQL = lsSQL & "cCalDias = (SELECT t.cCalif cAbrev FROM ColocCalificaTabla t        "
'        lsSQL = lsSQL & "            WHERE substring(ltrim(str(t.nCalCodTab)),1,3) = Substring(CO.cTpoCredCod,1,1) + '0' "
'        lsSQL = lsSQL & "             + Case When C.cRefinan ='R' then '1' else '0' end "
'        lsSQL = lsSQL & "            AND nRangIni <= C.nDiasAtraso And  nRangFin >= C.nDiasAtraso), "
'        'ALPA 20151219*******************************************************************************
''        lsSQL = lsSQL & "nGarPref =( Select Isnull(Sum(case when nMoneda =1 then nHipoteca + nVehicular + + nIndustrial + nAgricola "
''        lsSQL = lsSQL & "                                   when nMoneda = 2 then (nHipoteca + nVehicular + nIndustrial + nAgricola ) * " & pnTipoCambio & " "
''        lsSQL = lsSQL & "             end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "'), "
''        lsSQL = lsSQL & "nGarMuyRR = 0, "
''        lsSQL = lsSQL & "nGarAutoL = ( Select Isnull(sum(case when nmoneda = 1 then nCartaFianza + nDeposito "
''        lsSQL = lsSQL & "                                     when nmoneda = 2 then (nCartaFianza + nDeposito )  * " & pnTipoCambio & " "
''        lsSQL = lsSQL & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "' ) ,"
''        lsSQL = lsSQL & "nGarNOPref = ( Select Isnull(sum(case when nMoneda = 1 then nOtrosG + nJoyas "
''        lsSQL = lsSQL & "                                     when nMoneda = 2 then (nOtrosG + nJoyas)  * " & pnTipoCambio & " "
''        lsSQL = lsSQL & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "') "
'        lsSQL = lsSQL & "nGarPref = IsNull(TE.nSuma_1,0), "
'        lsSQL = lsSQL & "nGarMuyRR = 0, "
'        lsSQL = lsSQL & "nGarAutoL =IsNull(TE.nSuma_2,0), "
'        lsSQL = lsSQL & "nGarNOPref =IsNull(TE.nSuma_3,0) "
'        'ALPA 20100614 B2****************************************************************************
'        lsSQL = lsSQL & " ,C.nIntSusp,C.nIntDev,isnull(C.nCapVencido,0) nCapVencido"
'        '********************************************************************************************
'        lsSQL = lsSQL & " FROM  " & psRutaServer & "CreditoConsol C  "
'        lsSQL = lsSQL & " JOIN " & psRutaServer & "CREDITOSALDOCONSOL CS ON CS.cCtaCod =C.cCtaCod "
'        'ALPA 20100605 B2****************************************************************************
'        lsSQL = lsSQL & " INNER JOIN Colocaciones  CO ON C.cCtaCod =CO.cCtaCod "
'        '********************************************************************************************
'        lsSQL = lsSQL & " LEFT JOIN ##TGarantias  TE ON TE.cCtaCod =CO.cCtaCod "
'        lsSQL = lsSQL & " Where C.nPrdEstado in ( " & gColocEstRecVigJud & "," & gColocEstRecVigCast & ",2205,2206," & lsRFAEstadoJudicial & " )"
'        lsSQL = lsSQL & " AND DateDiff(dd,CS.dFecha,'" & Format(psFechaData, "mm/dd/yyyy") & "') = 0 "
'    'Carta Fianza
''    lsSQL = lsSQL & " Union " _
''         & " SELECT  C.cCtaCod, C.nMontoApr as nSaldoCap, nPrdEstado, 'N' as cRefinan, 0 as nDiasAtraso, " _
''         & " C.nMontoApr, '0101'+ Substring(C.cCtaCod,9,1) + '1' + Substring(C.cCtaCod,6,3)+'01' as cLineaCred , C.cCodAnalista, C.dFecVig, " _
''         & " cPersCod = (Select Max(cPersCod) From " & psRutaServer & "ProductoPersonaConsol PE " _
''         & "             Where PE.cCtaCod=C.cCtaCod And PE.nPrdPersRelac = " & gColRelPersTitular & " ) , " _
''         & "cCalDias = '0', " _
''         & "nGarPref =( Select Isnull(Sum(case when nMoneda =1 then nHipoteca + nVehicular + + nIndustrial + nAgricola " _
''         & "                                   when nMoneda = 2 then (nHipoteca + nVehicular + nIndustrial + nAgricola ) * " & pnTipoCambio & " " _
''         & "             end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "' ), " _
''         & "nGarMuyRR = 0, " _
''         & "nGarAutoL = ( Select Isnull(sum(case when nmoneda = 1 then nCartaFianza + nDeposito " _
''         & "                                     when nMoneda = 2 then (nCartaFianza + nDeposito )  * " & pnTipoCambio & " " _
''         & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "' ), " _
''         & " nGarNOPref = ( Select Isnull(sum(case when nMoneda = 1 then nOtrosG + nJoyas " _
''         & "                                     when nMoneda = 2 then (nOtrosG + nJoyas )  * " & pnTipoCambio & " " _
''         & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "') " _
''         & " FROM  " & psRutaServer & "CartaFianzaConsol C  " _
''         & " Where C.nPrdEstado in ( " & gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & "," & gColocEstRenovada & ") " _
''         & " )X ORDER BY cPersCod,cCtaCod"
'lsSQL = lsSQL & " Union "
''ALPA 20100605 B2****************************************************************************
''lsSQL = lsSQL & " SELECT  C.cCtaCod, C.nMontoApr as nSaldoCap, nPrdEstado, 'N' as cRefinan, 0 as nDiasAtraso, "
'lsSQL = lsSQL & " SELECT  CO.cAgeCodAct,CO.cTpoCredCod,CO.cTpoProdCod,C.cCtaCod, C.nMontoApr as nSaldoCap, nPrdEstado, 'N' as cRefinan, 0 as nDiasAtraso, "
''********************************************************************************************
'lsSQL = lsSQL & " C.nMontoApr, '0101'+ Substring(C.cCtaCod,9,1) + '1' + CO.cTpoCredCod +'01' as cLineaCred , C.cCodAnalista, C.dFecVig, "
'lsSQL = lsSQL & " cPersCod = (Select Max(cPersCod) From " & psRutaServer & "ProductoPersonaConsol PE "
'lsSQL = lsSQL & "             Where PE.cCtaCod=C.cCtaCod And PE.nPrdPersRelac = " & gColRelPersTitular & " ) , "
'lsSQL = lsSQL & "cCalDias = '0', "
''ALPA 20151219*********************************************************************************
''lsSQL = lsSQL & "nGarPref =( Select Isnull(Sum(case when nMoneda =1 then nHipoteca + nVehicular + + nIndustrial + nAgricola "
''lsSQL = lsSQL & "                                   when nMoneda = 2 then (nHipoteca + nVehicular + nIndustrial + nAgricola ) * " & pnTipoCambio & " "
''lsSQL = lsSQL & "             end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "' ), "
''lsSQL = lsSQL & "nGarMuyRR = 0, "
''lsSQL = lsSQL & "nGarAutoL = ( Select Isnull(sum(case when nmoneda = 1 then nCartaFianza + nDeposito "
''lsSQL = lsSQL & "                                     when nMoneda = 2 then (nCartaFianza + nDeposito )  * " & pnTipoCambio & " "
''lsSQL = lsSQL & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "' ), "
''lsSQL = lsSQL & " nGarNOPref = ( Select Isnull(sum(case when nMoneda = 1 then nOtrosG + nJoyas "
''lsSQL = lsSQL & "                                     when nMoneda = 2 then (nOtrosG + nJoyas )  * " & pnTipoCambio & " "
''lsSQL = lsSQL & "              end),0) From " & psRutaServer & "CredGarantias cg Where cg.cCtaCod = C.cCtaCod and convert(char(10),dfecha,112)='" & Format(psFechaData, "yyyymmdd") & "') "
'lsSQL = lsSQL & "nGarPref = IsNull(TE.nSuma_1,0), "
'lsSQL = lsSQL & "nGarMuyRR = 0, "
'lsSQL = lsSQL & "nGarAutoL =IsNull(TE.nSuma_2,0), "
'lsSQL = lsSQL & "nGarNOPref =IsNull(TE.nSuma_3,0) "
''ALPA 20100614 B2****************************************************************************
'lsSQL = lsSQL & " ,0 nIntSusp,0 nIntDev,0 nCapVencido "
''********************************************************************************************
'lsSQL = lsSQL & " FROM  " & psRutaServer & "CartaFianzaConsol C  "
''ALPA 20100605 B2****************************************************************************
'lsSQL = lsSQL & " INNER JOIN Colocaciones  CO ON C.cCtaCod =CO.cCtaCod "
''********************************************************************************************
'lsSQL = lsSQL & " LEFT JOIN ##TGarantias  TE ON TE.cCtaCod =CO.cCtaCod "
'lsSQL = lsSQL & " Where C.nPrdEstado in ( " & gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & "," & gColocEstRenovada & ") "
''lsSQL = lsSQL & " )X ORDER BY cPersCod,cCtaCod"
'lsSQL = lsSQL & " )X LEFT JOIN DBConsolidada..Persona P ON P.cPersCod = X.cPersCod " 'EJVG20140908
'lsSQL = lsSQL & " ORDER BY X.cPersCod,X.cCtaCod"
         
'         & " )X where cPersCod='1090100256650' ORDER BY cPersCod,cCtaCod"
    
    'ARCV ordenamos para que alinee correctamente
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    lsSQL = "exec stp_sel_obtenercartera '" & Format(psFechaData, "YYYY/MM/DD") & "'," & pnTipoCambio
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos para Calificacion"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos para Calificacion"
        Exit Function
    End If
     
    Set nCalificaLLenaTablaCalifProv = lrDatos
    Set lrDatos = Nothing
End Function

'ARCV 08-06-2007
'Chequear los dias de atraso
Public Function nCalificaCorrigeCalifxDiaAtraso(ByVal psCalif As String, ByVal pnPrdEstado As Integer, _
                                                ByVal psFecha As String, ByVal loConec As COMConecta.DCOMConecta, _
                                                ByVal psCtaCod As String) As String
Dim lsCalifCorrige As String
Dim lSql As String
Dim rs As ADODB.Recordset


    'Creditos en Recuperaciones = 4
If pnPrdEstado = gColocEstRecVigJud Or pnPrdEstado = gColocEstRecVigCast Or _
       pnPrdEstado = 2205 Or pnPrdEstado = 2206 Or pnPrdEstado = 2261 Or _
       pnPrdEstado = 2262 Or pnPrdEstado = 2265 Or pnPrdEstado = 2266 Then
        
        lsCalifCorrige = "4"
Else

'0: Normal, 1: CPP, 2: Deficiente, 3: Dudoso, 4: Perdido

'No refinanciados pagados en un mes
lSql = "SELECT CASE WHEN C.nPrdEstado IN (2020,2021,2022,2101,2104,2106,2107) THEN"
lSql = lSql & " ("
'ALPA20160509***************************************************************************************
'lSql = lSql & " CASE WHEN SUBSTRING(C.cCtaCod,6,1) ='2' OR SUBSTRING(C.cCtaCod,6,1) ='3' THEN"
lSql = lSql & " CASE WHEN SUBSTRING(CCT.cTpoCredCod,1,1) in ('4','5','6','7') THEN"
'***************************************************************************************************
lSql = lSql & "     (CASE WHEN DATEDIFF(d,P.dFecVenc,P.dFecPag)<=8 THEN '0'"
lSql = lSql & "  WHEN DATEDIFF(d,P.dFecVenc,P.dFecPag)>=9 AND DATEDIFF(d,P.dFecVenc,P.dFecPag)<=30 THEN '1'"
lSql = lSql & "  WHEN DATEDIFF(d,P.dFecVenc,P.dFecPag)>=31 AND DATEDIFF(d,P.dFecVenc,P.dFecPag)<=60 THEN '2'"
lSql = lSql & "  WHEN DATEDIFF(d,P.dFecVenc,P.dFecPag)>=61 AND DATEDIFF(d,P.dFecVenc,P.dFecPag)<=120 THEN '3'"
lSql = lSql & "  Else '4'"
lSql = lSql & "    End"
lSql = lSql & "    )"
'ALPA20160509***************************************************************************************
'lSql = lSql & "    WHEN SUBSTRING(C.cCtaCod,6,1) ='1' OR SUBSTRING(C.cCtaCod,6,1) ='4' THEN"
lSql = lSql & "    WHEN SUBSTRING(CCT.cTpoCredCod,1,1) in ('1','2','3','8') THEN"
'***************************************************************************************************
lSql = lSql & "     (CASE WHEN DATEDIFF(d,P.dFecVenc,P.dFecPag)<=30 THEN '0'"
lSql = lSql & "         WHEN DATEDIFF(d,P.dFecVenc,P.dFecPag)>=31 AND DATEDIFF(d,P.dFecVenc,P.dFecPag)<=90 THEN '1'"
lSql = lSql & "         WHEN DATEDIFF(d,P.dFecVenc,P.dFecPag)>=91 AND DATEDIFF(d,P.dFecVenc,P.dFecPag)<=120 THEN '2'"
lSql = lSql & "         WHEN DATEDIFF(d,P.dFecVenc,P.dFecPag)>=121 AND DATEDIFF(d,P.dFecVenc,P.dFecPag)<=365 THEN '3'"
lSql = lSql & "         Else '4'"
lSql = lSql & "     End"
lSql = lSql & "     )"
lSql = lSql & " End"
lSql = lSql & " )"
lSql = lSql & " ELSE "
'Refinanciados pagados en un mes


'ALPA 20130702************************************************************************************
lSql = lSql & " ( "
lSql = lSql & " Case WHEN "
lSql = lSql & "     (CASE WHEN C.nDiasAtraso<=8 THEN '1' "
lSql = lSql & "         WHEN C.nDiasAtraso>=9 AND C.nDiasAtraso<=30 THEN '2' "
lSql = lSql & "         WHEN C.nDiasAtraso>=31 AND C.nDiasAtraso<=90 THEN '3' "
lSql = lSql & "         Else '4' "
lSql = lSql & "     END ) < ISNULL(CCP.cCalGen,'0') THEN ISNULL(CCP.cCalGen,'0') "
lSql = lSql & " Else "
lSql = lSql & "             ( "
lSql = lSql & " Case WHEN "
lSql = lSql & "     (CASE WHEN C.nDiasAtraso<=8 THEN '1' "
lSql = lSql & "         WHEN C.nDiasAtraso>=9 AND C.nDiasAtraso<=30 THEN '2' "
lSql = lSql & "         WHEN C.nDiasAtraso>=31 AND C.nDiasAtraso<=90 THEN '3' "
lSql = lSql & "         Else '4' "
lSql = lSql & "     END ) < ISNULL(CCP.cCalGen,'0') THEN ISNULL(CCP.cCalGen,'0') "
lSql = lSql & "    Else "
lSql = lSql & "     (CASE WHEN C.nDiasAtraso<=8 THEN '1' "
lSql = lSql & "         WHEN C.nDiasAtraso>=9 AND C.nDiasAtraso<=30 THEN '2' "
lSql = lSql & "         WHEN C.nDiasAtraso>=31 AND C.nDiasAtraso<=90 THEN '3' "
lSql = lSql & "         Else '4'"
lSql = lSql & "       END ) "
lSql = lSql & "    End ) "
lSql = lSql & " End "
lSql = lSql & " ) "
lSql = lSql & " END As cCalifica"
lSql = lSql & " FROM DBConsolidada..CreditoConsol C INNER JOIN DBConsolidada..PlanDesPagConsol P"
lSql = lSql & " ON C.cCtaCod = P.cCtaCod AND nTipo= 1  AND nEstado= 1 AND"
lSql = lSql & " cNroCuo = (SELECT MIN(cNroCuo) FROM DBConsolidada..PlanDesPagConsol WHERE cCtaCod=C.cCtaCod AND nEstado= 1 AND nTipo=1"
lSql = lSql & "         AND MONTH(dFecPag)=MONTH('" & Format(psFecha, "yyyymmdd") & "') AND YEAR(dFecPag)=YEAR('" & Format(psFecha, "yyyymmdd") & "'))"
lSql = lSql & " AND MONTH(dFecPag)=MONTH('" & Format(psFecha, "yyyymmdd") & "') AND YEAR(dFecPag)=YEAR('" & Format(psFecha, "yyyymmdd") & "')"
lSql = lSql & " INNER JOIN DBConsolidada..CreditoConsolTotal CCT on C.cCtaCod=CCT.cCtaCod "
lSql = lSql & " LEFT JOIN ColocacRefinanc CR ON CR.cCtaCod = C.cCtaCod AND CR.nEstado= 3 "
lSql = lSql & " LEFT JOIN DBConsolidada..ColocCalifProvTotal CCP ON CCP.cCtaCod = CR.cCtaCodRef AND DATEDIFF(d,CCP.dFecha,'" & Format(psFecha, "yyyymmdd") & "')=0 "
lSql = lSql & " WHERE C.cCtaCod='" & psCtaCod & "' AND C.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107)"
'******************************************************************************************************************
    Set rs = loConec.CargaRecordSet(lSql)
    
    If Not rs.EOF Then
        lsCalifCorrige = rs!cCalifica
    Else
       'No se hicieron pagos en el MEs de Calificacion
       lSql = " SELECT CASE WHEN C.nPrdEstado IN (2020,2021,2022,2101,2104,2106,2107) THEN "
       lSql = lSql & "  ( "
       'ALPA20160509***************************************************************************************
       'lSql = lSql & " CASE WHEN SUBSTRING(C.cCtaCod,6,1) ='2' OR SUBSTRING(C.cCtaCod,6,1) ='3' THEN "
       lSql = lSql & " CASE WHEN SUBSTRING(CCT.cTpoCredCod,1,1) in ('4','5','6','7')  THEN "
       '***************************************************************************************************
       lSql = lSql & " (CASE WHEN C.nDiasAtraso<=8 THEN '0' "
       lSql = lSql & "    WHEN C.nDiasAtraso>=9 AND C.nDiasAtraso<=30 THEN '1' "
       lSql = lSql & "    WHEN C.nDiasAtraso>=31 AND C.nDiasAtraso<=60 THEN '2' "
       lSql = lSql & "    WHEN C.nDiasAtraso>=61 AND C.nDiasAtraso<=90 THEN '3' "
       lSql = lSql & "    Else '4' "
       lSql = lSql & " End "
       lSql = lSql & ") "
       'ALPA20160509***************************************************************************************
       'lSql = lSql & " WHEN SUBSTRING(C.cCtaCod,6,1) ='1' OR SUBSTRING(C.cCtaCod,6,1) ='4' THEN "
       lSql = lSql & " WHEN SUBSTRING(CCT.cTpoCredCod,1,1) in ('1','2','3','8') THEN "
       '***************************************************************************************************
       lSql = lSql & " (CASE WHEN C.nDiasAtraso<=30 THEN '0' "
       lSql = lSql & "    WHEN C.nDiasAtraso>=31 AND C.nDiasAtraso<=90 THEN '1' "
       lSql = lSql & "    WHEN C.nDiasAtraso>=91 AND C.nDiasAtraso<=120 THEN '2' "
       lSql = lSql & "    WHEN C.nDiasAtraso>=121 AND C.nDiasAtraso<=365 THEN '3' "
       lSql = lSql & "    Else '4' "
       lSql = lSql & " End "
       lSql = lSql & " ) "
       lSql = lSql & " End "
       lSql = lSql & " ) "
       lSql = lSql & " ELSE "
        '**DAOR 20070903, según Memo JTI 061-2007-CMACM, 073-UR-2007-CMACM
        lSql = lSql & " ( "
        lSql = lSql & " CASE WHEN C.nDiasAtraso<=8 THEN '1' "
        lSql = lSql & "     WHEN C.nDiasAtraso>=9 AND C.nDiasAtraso<=30 THEN '2' "
        lSql = lSql & "     WHEN C.nDiasAtraso>=31 AND C.nDiasAtraso<=90 THEN '3' "
        lSql = lSql & "     Else '4'"
        lSql = lSql & " End "
        lSql = lSql & " ) "
        lSql = lSql & " END As cCalifica "
        lSql = lSql & " FROM DBConsolidada..CreditoConsol C "
        lSql = lSql & " INNER JOIN DBConsolidada..CreditoConsolTotal CCT on C.cCtaCod=CCT.cCtaCod "
        lSql = lSql & " LEFT JOIN ColocacRefinanc CR ON CR.cCtaCod = C.cCtaCod AND CR.nEstado= 3 "
        lSql = lSql & " LEFT JOIN DBConsolidada..ColocCalifProvTotal CCP ON CCP.cCtaCod = CR.cCtaCodRef AND DATEDIFF(d,CCP.dFecha,'" & Format(psFecha, "yyyymmdd") & "')=0 "
        lSql = lSql & " WHERE C.cCtaCod='" & psCtaCod & "' AND C.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107)"

        Set rs = loConec.CargaRecordSet(lSql)
        If rs.EOF Then
            lsCalifCorrige = ""
        Else
            lsCalifCorrige = rs!cCalifica
        End If
    End If
'---

End If

nCalificaCorrigeCalifxDiaAtraso = lsCalifCorrige
End Function

'** Asigna calificacion al credito segun los dias de atraso y las tablas de codigos
'** Utiliza ColocCalificaTabla

Public Function nCalifxDiasAtraso(ByVal psCodCta As String, ByVal pnPrdEstado As Integer, _
        ByVal pbGarantPref As Boolean, ByVal pnDiasAtraso As Integer, _
        ByVal psRutaServer As String, _
        Optional ByRef psMensaje As String = "") As String


Dim rs As ADODB.Recordset
Dim loConecta As COMConecta.DCOMConecta
Dim lsSQL As String
Dim lsCodTab As String
Dim lsCondicProv As String
Dim lsRefinanciado As String
'If Mid(psCodCta, 6, 1) = "1" Or Mid(psCodCta, 6, 1) = "2" Or Mid(psCodCta, 6, 1) = "4" Then
'    lsCondicProv = Mid(psCodCta, 6, 1)
'    lsCondicProv = lsCondicProv & IIf(pbGarantPref = True, "1", "0")  ' 1 Con GarPref / 0 Sin GarPref
'ElseIf Mid(psCodCta, 6, 1) = "3" And Mid(psCodCta, 6, 3) <> "305" Then
'    lsCondicProv = Mid(psCodCta, 6, 1) & "0"
'ElseIf Mid(psCodCta, 6, 3) = "305" Then
'    lsCondicProv = "70"
'End If

'lsCodTab = lsCondicProv
If pnPrdEstado = "2030" Or pnPrdEstado = "2031" Or pnPrdEstado = "2032" _
                        Or pnPrdEstado = "2205" Then
    lsRefinanciado = "1"
Else
    lsRefinanciado = "0"
End If
lsCodTab = Mid(psCodCta, 6, 1) + IIf(pbGarantPref = True, "1", "0") + lsRefinanciado

If pnPrdEstado = gColocEstRecVigJud Or pnPrdEstado = 2205 Then ' Cobranza Judicial
        '// Creditos Judiciales con menos de 120 dias de mora acumulada
        '// Se califican Dudosos - 11/10/2002- Luis Lopez
'    If fgJudicialMora120(lsCodCta) = True Then  ' Si tiene Menos de 120 dias de Mora
'        fgAsinaCalificacionCredito = "3"
'    Else
        nCalifxDiasAtraso = "4"
'    endif
Else
    lsSQL = "Select cCalif From ColocCalificaTabla " _
          & "Where nCalCodTab like '" & lsCodTab & "%'  " _
          & "AND nRangIni <= " & pnDiasAtraso & " AND nRangFin >= " & pnDiasAtraso & " "
    Set loConecta = New COMConecta.DCOMConecta
        loConecta.AbreConexion
        Set rs = loConecta.CargaRecordSet(lsSQL)
    Set loConecta = Nothing
    If (rs.BOF And rs.EOF) Or rs Is Nothing Then
        psMensaje = " No encuentro Tabla de Calificaciones"
    Else
        nCalifxDiasAtraso = Trim(rs!cCalif)
    End If
    rs.Close
    Set rs = Nothing
End If
End Function


'*******************************************************************************
'** DE ACA PARA ADELANTE SON FUNCIONES PARA EVITAR LAS CONEXIONES INNECESARIAS
'*******************************************************************************
Public Sub nCalificaActualizarGarantias(ByVal pdFecSis As Date, _
                                ByVal pnTipoCambio As Double)

Dim loConec As COMConecta.DCOMConecta
Dim lsSQL As String
Dim rs As New ADODB.Recordset

Dim lnGarant As Integer

On Error GoTo ErrorActualizarGarantias

Set loConec = New COMConecta.DCOMConecta
    loConec.AbreConexion

'ARCV 09-06-2007
'    lsSQL = "UPDATE ColocCalifProv SET nGarant =  " _
'        & "Case When Substring(cCtaCod, 6, 1) = '4' then 3 " _
'        & "     When nGarAutoL >= (Case when substring(cCtaCod,9,1) ='1' then nSaldoCap else nSaldoCap * " & pnTipoCambio & " end ) then 1 " _
'        & "     When nGarMuyRR >= (Case when substring(cCtaCod,9,1) ='1' then nSaldoCap else nSaldoCap * " & pnTipoCambio & " end ) " _
'        & "          And Substring(cCtaCod, 6, 1) <> '3'  then 2   " _
'        & "     When nGarPref >= (Case when substring(cCtaCod,9,1) ='1' then nSaldoCap else nSaldoCap * " & pnTipoCambio & " end ) " _
'        & "           then 3   " _
'        & "     else 4 end "
'    loConec.Ejecutar lsSQL
'ALPA 20100605 B2***************************************************
'    lsSQL = "UPDATE ColocCalifProv SET nGarant =  " _
'        & "Case When Substring(cCtaCod, 6, 1) = '4' then 3 " _
'        & "     When nGarAutoL >= (Case when substring(cCtaCod,9,1) ='1' then nSaldoCap else nSaldoCap * " & pnTipoCambio & " end ) " _
'        & " OR (nGarAutoL> nGarPref AND nGarAutoL>nGarNoPref) then 1 " _
'        & "     When nGarMuyRR >= (Case when substring(cCtaCod,9,1) ='1' then nSaldoCap else nSaldoCap * " & pnTipoCambio & " end ) " _
'        & "  then 2  " _
'        & "     When nGarPref >= (Case when substring(cCtaCod,9,1) ='1' then nSaldoCap else nSaldoCap * " & pnTipoCambio & " end ) " _
'        & "     OR (nGarPref> nGarNoPref)      then 3   " _
'        & "     else 4 end "
'        lsSQL = "UPDATE ColocCalifProv SET nGarant =  " _
'        & "Case When Substring(cTpoCredCod, 1, 2) = '85' then 3 " _
'        & "     When nGarAutoL >= (Case when substring(cCtaCod,9,1) ='1' then nSaldoCap else nSaldoCap * " & pnTipoCambio & " end ) " _
'        & " OR (nGarAutoL> nGarPref AND nGarAutoL>nGarNoPref) then 1 " _
'        & "     When nGarMuyRR >= (Case when substring(cCtaCod,9,1) ='1' then nSaldoCap else nSaldoCap * " & pnTipoCambio & " end ) " _
'        & "  then 2  " _
'        & "     When nGarPref >= (Case when substring(cCtaCod,9,1) ='1' then nSaldoCap else nSaldoCap * " & pnTipoCambio & " end ) " _
'        & "     OR (nGarPref> nGarNoPref)      then 3   " _
'        & "     else 4 end "
        lsSQL = "UPDATE ColocCalifProv SET nGarant =  " _
        & "Case When nGarAutoL >= (Case when substring(cCtaCod,9,1) ='1' then nSaldoCap else nSaldoCap * " & pnTipoCambio & " end ) " _
        & " AND nGarAutoL> 0 then 1 " _
        & "     When nGarMuyRR >= (Case when substring(cCtaCod,9,1) ='1' then nSaldoCap else nSaldoCap * " & pnTipoCambio & " end ) and nGarMuyRR>0 " _
        & "  then 2  " _
        & "     When (nGarPref >= (Case when substring(cCtaCod,9,1) ='1' then nSaldoCap else nSaldoCap * " & pnTipoCambio & " end ) " _
        & "     AND (nGarPref> nGarNoPref)) and nGarPref>0     then 3   " _
        & "     else 4 end "
    loConec.Ejecutar lsSQL
'***********************************************************************
    'And Substring(cCtaCod, 6, 1) <> '3' (NO RESTRINGIR PARA LOS CONSUMOS)
    
    'No usar GarPreferida ==> Si tienen garantia preferida y son judiciales
    ' ==>> Es Perdida y  los dias de atraso > 24 meses =  720 dias de Ingreso a Judicial
    ' ==>> Es Dudoso y  los dias de atraso > 36 meses =  1080 dias de Ingreso a Judicial
    lsSQL = "UPDATE ColocCalifProv SET nGarant = 4 " _
        & "Where nPrdEstado in(" & gColocEstRecVigJud & "," & gColocEstRecVigCast & ")  " _
        & "And nGarant <> 4 and nDiasAtraso > 360 "
    loConec.Ejecutar lsSQL
    
    lsSQL = "UPDATE ColocCalifProv SET nGarant = 4 " _
        & "Where nGarant <> 4 and nDiasAtraso >= 360 "
    loConec.Ejecutar lsSQL
    
    'Actualiza Garantias Preferidas (No usa) ' Cusco 2006-12
    'No considera por el momento garantias autoliquidables Efrain 06/02/2007 Where nGarant = 3
    'lsSQL = "UPDATE ColocCalifProv SET nGarant = 4 Where nGarant = 3 "'
    
    'Se quito para MAYNAS 05-06-2007
    'lsSQL = "UPDATE ColocCalifProv SET nGarant = 4  "
    'loConec.Ejecutar lsSQL
    
    
    
    '** Actualiza la Garantias Preferidas Seleccionada - 2006-12  layg
    'Comentado by JACA 20110630*************************************************
'            lsSQL = " Update ColocCalifProv set nGarant = 3 Where cCtaCod in " _
'                  & "   (Select GCC.cCtaCod from ColocCalifGarantia GPref " _
'                  & "    join ColocGarantia GCC on GPref.cNumGarant = GCC.cNumGarant " _
'                  & "     Where nEstadoGar = 0 )"
'
'
'            loConec.Ejecutar lsSQL
'            'Quita las desmarcadas
'            lsSQL = " Update ColocCalifProv set nGarant = 4 Where cCtaCod in " _
'                  & "   (Select GCC.cCtaCod from ColocCalifGarantia GPref " _
'                  & "    join ColocGarantia GCC on GPref.cNumGarant = GCC.cNumGarant " _
'                  & "     Where nEstadoGar = 1 )"
'
'            loConec.Ejecutar lsSQL
    'JACA END**********************************************************************
    
    'By Capi Segun Memo 083-UR pa que no considere a los creditos de consumo con Garantia Preferida
        'lsSQL = " Update ColocCalifProv set nGarant = 4 Where Substring(cCtaCod,6,1)='3' And nGarant=3 "
        lsSQL = " Update ColocCalifProv set nGarant = 4 Where Substring(cTpoCredCod,1,2) in ('65','75') And nGarant=3 "
        loConec.Ejecutar lsSQL
        'ALPA 20100727****************************************************************
        lsSQL = "update colocCalifProv set nGarant = 4  where cTpoCredCod='755'"
        loConec.Ejecutar lsSQL
        '*****************************************************************************
    'End By
    Set loConec = Nothing
    
    Exit Sub
    
ErrorActualizarGarantias:
    Err.Raise Err.Number, "Actualizar Garantias", Err.Description
End Sub

Public Sub nCalificaCMACT(ByVal pnTipoCambio As Double)
' Calificacion mayor de los productos que tenga en la Cmact
' 2007-01 (1494) LAYG no alinear menores a 100 soles
Dim lsSQL As String
Dim rs As New ADODB.Recordset
Dim loConec As COMConecta.DCOMConecta
Dim lrDat As ADODB.Recordset
Dim lsCadConexion As String

On Error GoTo ErrorCalificaCMACT

Set loConec = New COMConecta.DCOMConecta
    loConec.AbreConexion 'lsCadConexion
    '*** para utilizar cCalRUnico - ICA
    'RFA - Actualiza Calif del RFA
    lsSQL = "UPDATE ColocCalifProv SET cCalRUnico = cCalRFA WHERE cCalRFA is not Null "
    loConec.Ejecutar lsSQL
    
    'Actualiza Calif Riesgo Unico con la Calificacion por dias de atraso en caso no tenga una asignada
    lsSQL = "UPDATE ColocCalifProv SET cCalRUnico = cCalNor WHERE (cCalRUnico is Null or cCalRUnico in (' ')) "
    loConec.Ejecutar lsSQL
    '****
    'Actualiza la Calificacion CMAC 'Para los otros que son menores a 100
    lsSQL = "Update ColocCalifProv set cCalCmac = y.CalCaja "
    lsSQL = lsSQL & "From ColocCalifProv x "
    lsSQL = lsSQL & " Join "
    lsSQL = lsSQL & "(Select cPersCod, "
    lsSQL = lsSQL & " CalCaja =( Select Max(ca1.cCalRUnico) "
    lsSQL = lsSQL & "          From ColocCalifProv ca1 "
    lsSQL = lsSQL & "          Where ca1.cPersCod = ca.cPersCod "
    'lssql = lssql & "          And ca1.cEstado in ('F','1','4','6','7') "
    lsSQL = lsSQL & "          And ca1.nSaldoCap > (case when substring(ca1.cCtaCod,9,1)='1' then 100 else 100 / " & pnTipoCambio & " end) "
    lsSQL = lsSQL & "          AND ca1.nPrdEstado IN(2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107) )  "
    lsSQL = lsSQL & " From ColocCalifProv ca "
    lsSQL = lsSQL & ") y "
    lsSQL = lsSQL & "On y.cPersCod = x.cPersCod  "
    
    loConec.Ejecutar lsSQL
    
    'Para los otros que son menores a 100
    lsSQL = "Update ColocCalifProv set cCalCmac = y.CalCaja "
    lsSQL = lsSQL & "From ColocCalifProv x "
    lsSQL = lsSQL & " Join "
    lsSQL = lsSQL & "(Select cPersCod, "
    lsSQL = lsSQL & " CalCaja =( Select Max(ca1.cCalRUnico) "
    lsSQL = lsSQL & "          From ColocCalifProv ca1 "
    lsSQL = lsSQL & "          Where ca1.cPersCod = ca.cPersCod "
    lsSQL = lsSQL & "          )  "
    lsSQL = lsSQL & " From ColocCalifProv ca "
    lsSQL = lsSQL & " WHERE nPrdEstado IN (2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107)) y "
    lsSQL = lsSQL & "On y.cPersCod = x.cPersCod  "
    lsSQL = lsSQL & "Where x.cCalCmac is null  "
    
    loConec.Ejecutar lsSQL
    
    '**DAOR 20070910, Modificaciones segùn Informe 0005-2007-CD-CMAC-M**************
    'La calificación CMACM para clientes con créditos castigados en Perdido(4)
    lsSQL = "Update ColocCalifProv set cCalCmac = '4' "
    lsSQL = lsSQL & "From ColocCalifProv x "
    lsSQL = lsSQL & " Join "
    lsSQL = lsSQL & "(Select cPersCod "
    lsSQL = lsSQL & " From ColocCalifProv ca "
    lsSQL = lsSQL & " WHERE nPrdEstado IN(2202,2206)  ) y "
    lsSQL = lsSQL & "On y.cPersCod = x.cPersCod  "
    lsSQL = lsSQL & "Where x.cCalCmac is null  "
    
    loConec.Ejecutar lsSQL
    
    '*******************************************************************************
    'ALPA 20100723******************************************************************
lsSQL = " Update ColocCalifProv "
lsSQL = lsSQL & "set cCalCmac=case when cCalRUnico>(case Calif when '0' then '1' else Calif end) then cCalRUnico else (case Calif when '0' then '1' else Calif end) end, "
lsSQL = lsSQL & "cCalRUnico=case when cCalRUnico>(case Calif when '0' then '1' else Calif end) then cCalRUnico else (case Calif when '0' then '1' else Calif end) end "
lsSQL = lsSQL & "From "
lsSQL = lsSQL & "( "
lsSQL = lsSQL & "select x.cCtaCod,case isnull(cCalNuev,'') when '' then cCalAnte else cCalNuev end Calif,x.cPersCod "
lsSQL = lsSQL & "From "
lsSQL = lsSQL & "( "
lsSQL = lsSQL & " select distinct C.cCtaCod,CT.cCalGen cCalNuev, cCalAnte,cCtaCodRef,C.cPersCod "
lsSQL = lsSQL & "from ColocCalifProv C "
lsSQL = lsSQL & " left join DBConsolidada..ColocCalifProvTotal CT on C.cCtaCod=CT.cCtaCod "
lsSQL = lsSQL & " and datediff(d,CT.dFecha,(select Max(dFecha) from DBConsolidada..ColocCalifProvTotal where cCtaCod=CT.cCtaCod))=0 "
lsSQL = lsSQL & " left join ( "
lsSQL = lsSQL & " select CR.cCtaCod,cCtaCodRef,cCalGen cCalAnte "
lsSQL = lsSQL & " from ColocacRefinanc CR "
lsSQL = lsSQL & " inner join DBConsolidada..ColocCalifProvTotal CT2 on CR.cCtaCodRef=CT2.cCtaCod "
lsSQL = lsSQL & "    and datediff(d,CT2.dFecha,(select Max(dFecha) from DBConsolidada..ColocCalifProvTotal where cCtaCod=CT2.cCtaCod))=0 "
lsSQL = lsSQL & " ) CR on C.cCtaCod=CR.cCtaCod "
lsSQL = lsSQL & " where C.cRefinan='R' "
lsSQL = lsSQL & " ) x "
lsSQL = lsSQL & " ) y inner join ColocCalifProv P on y.cPersCod=P.cPersCod "
loConec.Ejecutar lsSQL

lsSQL = "exec B2_stp_upd_CambiarCalificacionAlineamientoInterno " & pnTipoCambio & ""
loConec.Ejecutar lsSQL

'ALPA 20110905
lsSQL = lsSQL & " Update ColocCalifProv "
lsSQL = lsSQL & " Set cCalCMAC = x.cCalCMACM "
lsSQL = lsSQL & " From "
lsSQL = lsSQL & " ( "
lsSQL = lsSQL & " select max(cCalCMAC) cCalCMACM,cPersCod "
lsSQL = lsSQL & " From ColocCalifProv "
lsSQL = lsSQL & " group by cPersCod "
lsSQL = lsSQL & " ) x inner join ColocCalifProv c on x.cPersCod=c.cPersCod "
loConec.Ejecutar lsSQL

Set loConec = Nothing
Exit Sub

ErrorCalificaCMACT:
    Err.Raise Err.Number, "Califica CMACT", Err.Description
End Sub

Public Sub nCalificaEvaluacionRiesgos(ByVal psFecha As String)

Dim lsSQL As String
Dim lsCalEvalua As String
Dim loConec As COMConecta.DCOMConecta
Dim lrDat As ADODB.Recordset
Dim lrDatEval As ADODB.Recordset

'*** Obtiene Calificacion en ultima Evaluacion de Colocaciones
'*** Asignada en la Evaluacion de la Calificacion
'*** de la tabla ColocEvalCalif

On Error GoTo ErrorCalculaEvaluacion
Set loConec = New COMConecta.DCOMConecta
    loConec.AbreConexion
    
'ALPA 20110901*****************

lsSQL = " Update ColocEvalCalifDetalle"
lsSQL = lsSQL & " Set cEvalCalifDet = cCalGen"
lsSQL = lsSQL & " from DBConsolidada..ColocCalifProvTotal CT"
lsSQL = lsSQL & "         inner join ColocEvalCalifDetalle CD on CT.cCtaCod=CD.cCtaCod"
lsSQL = lsSQL & " where DateDiff(d,DATEADD(M,-1,'" & Format(psFecha, "yyyy/mm/dd") & "'),dFecha)=0"
lsSQL = lsSQL & "        and DATEDIFF(d,dEval,'" & Format(psFecha, "yyyy/mm/dd") & "')=0"
lsSQL = lsSQL & "        and cEvalObsDet not like '%comportamiento de pago%'"
loConec.Ejecutar lsSQL

lsSQL = " Update ColocEvalCalifDetalle "
lsSQL = lsSQL & " set cEvalCalifDet ="
lsSQL = lsSQL & " case when   nPrdEstado in ('2202','2206') then 4"
lsSQL = lsSQL & " Else"
lsSQL = lsSQL & " case when nCalDA>x.cEvalCalifDet"
lsSQL = lsSQL & " then nCalDA"
lsSQL = lsSQL & " End"
lsSQL = lsSQL & " end,"
lsSQL = lsSQL & "    cEvalObsDet = 'Actual.Sistem.'"
lsSQL = lsSQL & " From"
lsSQL = lsSQL & " ("
lsSQL = lsSQL & " select x.cPerscod,x.nLogico,x.cEvalCalifDet,max(C.nDiasAtraso) nDiasAtraso,"
lsSQL = lsSQL & " case when c.nDiasAtraso<=0 then 0"
lsSQL = lsSQL & " when c.nDiasAtraso between 1  and  8 then 1"
lsSQL = lsSQL & "   when c.nDiasAtraso between 9  and 30 then 2"
lsSQL = lsSQL & "   when c.nDiasAtraso between 31 and 90 then 3"
lsSQL = lsSQL & "  else 4"
lsSQL = lsSQL & "  end nCalDA,cRefinan,nPrdEstado"
lsSQL = lsSQL & " From"
lsSQL = lsSQL & " ("
lsSQL = lsSQL & " select cPerscod,SUM(nLogico) nLogico,max(cEvalCalifDet) cEvalCalifDet"
lsSQL = lsSQL & " From"
lsSQL = lsSQL & " ("
lsSQL = lsSQL & " select cPerscod,case DateDiff(d,dEval,'" & Format(psFecha, "yyyy/mm/dd") & "')"
lsSQL = lsSQL & "            when 0 then cEvalCalifDet else -1 end cEvalCalifDet,"
lsSQL = lsSQL & "        cEvalObsDet , nDiasAtraso"
lsSQL = lsSQL & " ,case cEvalObsDet when 'Eval.Automat' then 1 else 0 end nLogico"
lsSQL = lsSQL & " From ColocEvalCalifDetalle"
lsSQL = lsSQL & " where "
lsSQL = lsSQL & "   DateDiff(d,dEval,'" & Format(psFecha, "yyyy/mm/dd") & "') between 0 and 170"
lsSQL = lsSQL & " group by cPerscod,dEval,"
lsSQL = lsSQL & " cEvalCalifDet,cEvalObsDet,nDiasAtraso) x"
lsSQL = lsSQL & " group by cPerscod"
lsSQL = lsSQL & " ) X"
lsSQL = lsSQL & " inner join ColocCalifProv C on X.cPersCod=C.cPersCod"
lsSQL = lsSQL & " where (C.cRefinan='R' or C.nPrdEstado in ('2202','2206'))"
lsSQL = lsSQL & " group by x.cPerscod,x.nLogico,x.cEvalCalifDet,case when c.nDiasAtraso<=0 then 0"
lsSQL = lsSQL & " when c.nDiasAtraso between 1  and  8 then 1"
lsSQL = lsSQL & "   when c.nDiasAtraso between 9  and 30 then 2"
lsSQL = lsSQL & "   when c.nDiasAtraso between 31 and 90 then 3"
lsSQL = lsSQL & "  else 4"
lsSQL = lsSQL & "  end,cRefinan,nPrdEstado"
lsSQL = lsSQL & " ) x"
lsSQL = lsSQL & " inner join ColocEvalCalifDetalle CD on x.cPersCod=CD.cPersCod "
lsSQL = lsSQL & " where DateDiff(d,dEval,'" & Format(psFecha, "yyyy/mm/dd") & "') = 0"
lsSQL = lsSQL & "   and cEvalObsDet not like '%comportamiento de pago%'"
lsSQL = lsSQL & "   and ( (nPrdEstado in ('2202','2206')) or (nCalDA>x.cEvalCalifDet))"
loConec.Ejecutar lsSQL
'******************************


lsSQL = " SELECT  distinct(cPersCod) cPerscod, ca.cRefinan " _
      & ",cCalEval = (SELECT Max(a.cEvalCalifDet) cCalEval FROM ColocEvalCalifDetalle a " _
      & "             WHERE a.cPersCod = ca.cPersCod " _
      & "             AND nEvalTipo = 0 AND DateDiff(dd,dEval,'" & Format(psFecha, "mm/dd/yyyy") & "') = 0 ) " _
      & ",cObserva = (SELECT Max(a.cEvalObsDet) FROM ColocEvalCalifDetalle a " _
      & "             WHERE a.cPersCod = ca.cPersCod " _
      & "             AND nEvalTipo = 0 AND DateDiff(dd,dEval,'" & Format(psFecha, "mm/dd/yyyy") & "') = 0 ) " _
      & " FROM ColocCalifProv ca "


    Set lrDat = loConec.CargaRecordSet(lsSQL)

    If Not (lrDat.BOF And lrDat.EOF) Then
        Do While Not lrDat.EOF
            lsCalEvalua = IIf(IsNull(lrDat!cCalEval), "", lrDat!cCalEval)
            ' ******************
            If lsCalEvalua <> "" Then
                '*** Actualiza ColocCalifProv
                'By Capi Memo 083-2007
                If lrDat!cRefinan = "R" Then
                    lsSQL = "UPDATE ColocCalifProv SET cCalEval ='" & lsCalEvalua & "' " _
                      & "WHERE cPersCod ='" & Trim(lrDat!cPersCod) & "' "
                      'By Capi Segun Memo 083-UR pa que solo cambie a los refinanciados
                      loConec.Ejecutar lsSQL
                Else
                    If lrDat!cObserva <> "Eval.Automat" Then
                        lsSQL = "UPDATE ColocCalifProv SET cCalEval ='" & lsCalEvalua & "' " _
                        & "WHERE cPersCod ='" & Trim(lrDat!cPersCod) & "' "
                        loConec.Ejecutar lsSQL
                    End If
                      'By Capi Segun Memo 083-UR pa que solo cambie a los refinanciados
                End If
                
                
                'End By
            End If
            '************
        
            lrDat.MoveNext
        Loop
    End If
    Set lrDat = Nothing
    'ALPA 20100408******************************************
    lsSQL = "update ColocCalifProv set  "
    lsSQL = lsSQL & " nIntSusp=isnull(nIntSusp,0)+isnull(nIntDev,0),"
    lsSQL = lsSQL & " nIntDev=0 "
    lsSQL = lsSQL & " where (cCalGen in (3,4) or cRefinan in ('R'))"
    lsSQL = lsSQL & "   and nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107,2092)"
    loConec.Ejecutar lsSQL
    
    lsSQL = "update ColocCalifProv set  "
    lsSQL = lsSQL & " nIntDev=isnull(nIntSusp,0)+isnull(nIntDev,0),"
    lsSQL = lsSQL & " nIntSusp=0 "
    lsSQL = lsSQL & " where cCalGen in (0,1,2) and cRefinan in ('N')"
    lsSQL = lsSQL & "   and nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107,2092)"
    loConec.Ejecutar lsSQL
    
    lsSQL = "update DbConsolidada..CreditoConsol"
    lsSQL = lsSQL & " set nIntSusp=isnull(DC.nIntSusp,0)+isnull(DC.nIntDev,0),"
    lsSQL = lsSQL & " nIntDev = 0"
    lsSQL = lsSQL & " from ColocCalifProv C inner join DbConsolidada..CreditoConsol DC on (C.cCtaCod=DC.cCtaCod)"
    lsSQL = lsSQL & " where (C.cCalGen in (3,4) or C.cRefinan in ('R'))"
    lsSQL = lsSQL & "   and C.nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107,2092)"
    loConec.Ejecutar lsSQL
    
    lsSQL = " update DbConsolidada..CreditoConsol"
    lsSQL = lsSQL & " set nIntDev=isnull(DC.nIntSusp,0)+isnull(DC.nIntDev,0),"
    lsSQL = lsSQL & " nIntSusp = 0"
    lsSQL = lsSQL & " from ColocCalifProv C inner join DbConsolidada..CreditoConsol DC on (C.cCtaCod=DC.cCtaCod)"
    'ALPA 201010**************
    'lsSQL = lsSQL & " where C.cCalGen in (0,1,2) or C.cRefinan in ('N')"
    lsSQL = lsSQL & " where C.cCalGen in (0,1,2) and C.cRefinan in ('N')"
    '*************************
    lsSQL = lsSQL & "   and C.nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107,2092)"
    loConec.Ejecutar lsSQL
    '*******************************************************
    Set loConec = Nothing
    Exit Sub
ErrorCalculaEvaluacion:
    Err.Raise Err.Number, "Califica Evaluacion Riesgos ", Err.Description
End Sub


Public Sub nCalificaSistemaFinanciero(ByVal psBDRCC As String, _
                                    ByVal pdFecAlin As Date, _
                                    ByRef psMensaje As String, ByVal psServConsol As String)
' Obtiene la Calificac del Sistema Financiero a CreditoAudi
Dim lsSQL As String
Dim lrDat As ADODB.Recordset
Dim loConec As COMConecta.DCOMConecta

On Error GoTo ErrorCalificaSistemaFinanciero

    Set loConec = New COMConecta.DCOMConecta
    loConec.AbreConexion

    'ALPA 20101124 *************************
    If nExisteTabla("##tmpRCC", psServConsol) = True Then
        lsSQL = "drop table ##tmpRCC"
        loConec.Ejecutar lsSQL
    End If
    
    lsSQL = "select distinct D.Cod_Emp,R2.Cod_Edu,D.dFecha into ##tmpRCC"
    lsSQL = lsSQL & " From " & psBDRCC & "RCCTotalDet D"
    lsSQL = lsSQL & "    inner join " & psBDRCC & "RCCTotal R2 on D.Cod_Edu=R2.Cod_Edu"
    lsSQL = lsSQL & "       and CONVERT(CHAR(10),D.dFecha,112)=CONVERT(CHAR(10),R2.FEC_REP,112)"
    lsSQL = lsSQL & " Where D.Cod_Edu = R2.Cod_Edu"
    lsSQL = lsSQL & "   and R2.Can_Ents=1"
    lsSQL = lsSQL & "   and CONVERT(CHAR(10),D.dFecha,112)='" & Format(pdFecAlin, "yyyymmdd") & "' and D.Cod_Emp<>'00109'"
    loConec.Ejecutar lsSQL
    '**************************************
    lsSQL = "Update ColocCalifProv Set cCalSistF = null "
    loConec.Ejecutar lsSQL

    '*********************************************
    'Personas - Documento Civil
    lsSQL = " Select CA.cPersCod, Can_Ents, "
    lsSQL = lsSQL & " Case when Calif_4 > 20 then 4 "
    lsSQL = lsSQL & "      when Calif_3 > 20 then 3 "
    lsSQL = lsSQL & "      when Calif_2 > 20 then 2 "
    lsSQL = lsSQL & "      when Calif_1 > 20 then 1 "
    lsSQL = lsSQL & "      else 0 end  as nCalSF "
    lsSQL = lsSQL & " FROM ColocCalifProv CA "
    lsSQL = lsSQL & " INNER join Persona P on CA.cPersCod = P.cPersCod "
    lsSQL = lsSQL & " INNER join PersId PI on P.cPersCod = PI.cPersCod "
    lsSQL = lsSQL & " INNER join " & psBDRCC & "RCCTotal R on LTRIM(RTRIM(R.cod_Doc_Id)) = LTRIM(RTRIM(PI.cPersIDnro)) and CONVERT(CHAR(10),R.FEC_REP,112)='" & Format(pdFecAlin, "yyyymmdd") & "' "
    'ALPA 20101124********************************************************
    lsSQL = lsSQL & " LEFT JOIN ##tmpRCC T on R.Cod_Edu=T.Cod_Edu"
    lsSQL = lsSQL & "    and CONVERT(CHAR(10),R.FEC_REP,112)=CONVERT(CHAR(10),T.dFecha,112)"
    '*********************************************************************
    lsSQL = lsSQL & " WHERE PI.cPersIDTpo = " & gPersIdDNI & " "
    lsSQL = lsSQL & " AND ( LTRIM(RTRIM(PI.cPersIDnro)) <> '' or PI.cPersIDnro <> null ) "
    lsSQL = lsSQL & " And ( r.Can_Ents > 1     "
    'ALPA 20101124********************************************************
    lsSQL = lsSQL & " or (T.Cod_Emp is Not NULL)"
    '*********************************************************************
    lsSQL = lsSQL & "     ) "
    lsSQL = lsSQL & " AND (case R.Tip_Doc_Id when 1 then 1 "
    lsSQL = lsSQL & "                          when 4 then 2 "
    lsSQL = lsSQL & "    when 5 then 11 end)=PI.cPersIdTpo "
    lsSQL = lsSQL & " Group by CA.cPersCod, Can_Ents, "
    lsSQL = lsSQL & " Case when Calif_4 > 20 then 4 "
    lsSQL = lsSQL & "      when Calif_3 > 20 then 3 "
    lsSQL = lsSQL & "      when Calif_2 > 20 then 2 "
    lsSQL = lsSQL & "      when Calif_1 > 20 then 1 "
    lsSQL = lsSQL & "      Else 0 end "

        'Se saco la condicion
        '& "        or (r.Can_Ents = 1 And  ca.dFecVig >= '" & Format(DateAdd("d", -1 * Day(gdFecData), gdFecData), "yyyy/mm/dd") & "' )  " _

    Set lrDat = loConec.CargaRecordSet(lsSQL)
        If lrDat.BOF And lrDat.EOF Then
            psMensaje = "No existe Data para Calificar Personas Natural"
        Else
        
            Do While Not lrDat.EOF
                
                lsSQL = " UPDATE ColocCalifProv Set cCalSistF = '" & Trim(Str(lrDat!nCalSF)) & "' " _
                    & " WHERE cPersCod ='" & lrDat!cPersCod & "'  "
                loConec.Ejecutar lsSQL
                
                lrDat.MoveNext
            Loop
        End If
    Set lrDat = Nothing
    
    '*********************************************
    'Personas - Documento Tributario
    lsSQL = " Select CA.cPersCod, Can_Ents,  "
    lsSQL = lsSQL & " Case when Calif_4 > 20 then 4 "
    lsSQL = lsSQL & "      when Calif_3 > 20 then 3 "
    lsSQL = lsSQL & "      when Calif_2 > 20 then 2 "
    lsSQL = lsSQL & "      when Calif_1 > 20 then 1 "
    lsSQL = lsSQL & "      else 0 end  as nCalSF "
    lsSQL = lsSQL & " FROM ColocCalifProv CA "
    lsSQL = lsSQL & " INNER join Persona P on CA.cPersCod = P.cPersCod "
    lsSQL = lsSQL & " INNER join PersId PI on P.cPersCod = PI.cPersCod "
    lsSQL = lsSQL & " INNER join " & psBDRCC & "RCCTotal R on LTRIM(RTRIM(R.Cod_Doc_Trib)) = LTRIM(RTRIM(PI.cPersIDnro)) "
     'ALPA 20101124********************************************************
    lsSQL = lsSQL & " LEFT JOIN ##tmpRCC T on R.Cod_Edu=T.Cod_Edu"
    lsSQL = lsSQL & "    and CONVERT(CHAR(10),R.FEC_REP,112)=CONVERT(CHAR(10),T.dFecha,112)"
    '*********************************************************************
    lsSQL = lsSQL & " WHERE PI.cPersIDTpo = " & gPersIdRUC & " "
    lsSQL = lsSQL & " AND ( LTRIM(RTRIM(PI.cPersIDnro)) <> '' or PI.cPersIDnro <> null ) "
    lsSQL = lsSQL & " And ( r.Can_Ents > 1     "
    'ALPA 20101124********************************************************
    lsSQL = lsSQL & " or (T.Cod_Emp is Not NULL)"
    '*********************************************************************
    lsSQL = lsSQL & "     ) "
'    lsSQL = lsSQL & "AND (case R.Tip_Doc_Id when 1 then 1 "
'    lsSQL = lsSQL & "                          when 4 then 2 "
'    lsSQL = lsSQL & "    when 5 then 11 end)=PI.cPersIdTpo "
    lsSQL = lsSQL & " Group by CA.cPersCod, Can_Ents,  "
    lsSQL = lsSQL & " Case when Calif_4 > 20 then 4 "
    lsSQL = lsSQL & "      when Calif_3 > 20 then 3 "
    lsSQL = lsSQL & "      when Calif_2 > 20 then 2 "
    lsSQL = lsSQL & "      when Calif_1 > 20 then 1 "
    lsSQL = lsSQL & "      Else 0 end "
        
        'Se saco condicion
        '& "        or (r.Can_Ents = 1 And  ca.dFecVig >= '" & Format(DateAdd("d", -1 * Day(gdFecData), gdFecData), "yyyy/mm/dd") & "' ) " _

    Set lrDat = loConec.CargaRecordSet(lsSQL)
        If lrDat.BOF And lrDat.EOF Then
            psMensaje = IIf(psMensaje <> "", psMensaje & vbCrLf, "") & "No existe Data para Calificar Persona Juridica "
        Else
            Do While Not lrDat.EOF
        
                lsSQL = " UPDATE ColocCalifProv Set cCalSistF = '" & Trim(Str(lrDat!nCalSF)) & "' " _
                    & " WHERE cPersCod ='" & lrDat!cPersCod & "'  "
        
                loConec.Ejecutar lsSQL
                
                lrDat.MoveNext
            Loop
        End If
    'ALPA 20111010
    lsSQL = "exec stp_upd_ColocCalifProvCan_EntsYnSumSalRCC "
    loConec.Ejecutar lsSQL
        
    Set lrDat = Nothing
    Set loConec = Nothing

Exit Sub

ErrorCalificaSistemaFinanciero:
    Err.Raise Err.Number, "Califica Sistema Financiero", Err.Description
End Sub

Public Sub nCalificaGeneral(ByRef psMensaje As String, Optional dFecha As Date)
Dim lsSQL As String
Dim lrs As ADODB.Recordset
Dim lsCalifGen As String
Dim lsCalCMAC As String
Dim lsCalSistF As String, lsCalHist As String
Dim lsCalEval As String
Dim rsVerif As ADODB.Recordset
Dim loConec As COMConecta.DCOMConecta

Dim lsCalObser As String ' Calific Observada ***

On Error GoTo ErrorCalificacionGeneral

Set loConec = New COMConecta.DCOMConecta
loConec.AbreConexion

'lsSQL = " SELECT  distinct(cPersCod), cCalCMAC, cCalSistF, cCalEval, cFlagProv, " _
'      & " Isnull(Sum(case when substring(cTpoCredCod,1,2) in ('15','25','35') then 0 else 1 end),0) as Minorista  " _
'      & " FROM ColocCalifProv  " _
'      & " Group by cPersCod, cCalCMAC, cCalSistF, cCalEval, cFlagProv "
lsSQL = " SELECT  distinct(Per.cPersCod), Per.cCalCMAC, Per.cCalSistF, Per.cCalEval, Per.cFlagProv, "
lsSQL = lsSQL & " Isnull(Sum(case when substring(Per.cTpoCredCod,1,2) in ('15','25','35') then 0 else 1 end),0) as Minorista,Per.nDiasAtraso,Per.nPrdEstado,Per.cRefinan,nDiasAtrasoSF=IsNull(tDASistF.nDiasAtrasoSF,0)  "
lsSQL = lsSQL & " FROM ColocCalifProv  Per "
lsSQL = lsSQL & "  LEFT join ("
lsSQL = lsSQL & "    select max(distinct RD.Condicion) nDiasAtrasoSF,C.cPersCod"
lsSQL = lsSQL & "    from DBConsolidada..RCCTotal RC"
lsSQL = lsSQL & "    inner join DBConsolidada..RCCTotalDet RD on RC.Cod_Edu=RD.Cod_Edu,"
lsSQL = lsSQL & "    Persid P,ColocCalifProv C"
lsSQL = lsSQL & "    Where (RC.Cod_Doc_Trib = P.cPersIDNro Or RC.Cod_Doc_Id = P.cPersIDNro)"
lsSQL = lsSQL & "      and P.cPersCod=C.cPersCod and right(RD.Cod_Emp,3) <>'109'"
lsSQL = lsSQL & "    group by C.cPersCod"
lsSQL = lsSQL & "   ) tDASistF on Per.cPersCod = tDASistF.cPersCod"
'lsSQL = lsSQL & " WHERE Per.cPersCod='1090100198897'"
lsSQL = lsSQL & " Group by Per.cRefinan,Per.nPrdEstado,Per.nDiasAtraso,Per.cPersCod, Per.cCalCMAC, Per.cCalSistF, Per.cCalEval, Per.cFlagProv, IsNull(tDASistF.nDiasAtrasoSF,0) "
      
Set loConec = New COMConecta.DCOMConecta
loConec.AbreConexion
Set lrs = loConec.CargaRecordSet(lsSQL)
    If lrs.BOF And lrs.EOF Then
        psMensaje = "No existen Datos para Calificar "
    Else
        Do While Not lrs.EOF
        
            '********************************************************************************
            '** En caso la Ofic. Riesgos le haya asignado una calificacion, se asigna esta calificacion
            '** Si el Cliente no tiene calificacion en el Sist. Financiero se le asigna la Calif de la CMAC
            '** Si la Calificacion de la CMAC es mayor o igual que la del Sist. Financiero se le asigna la Calif CMAC
            '** En caso de Calif Sist Financ sea mayor se le asigna la Calif del Sist.Financiero menos una categoria
            If lrs!Minorista = 1 And lrs!nDiasAtraso >= 1800 And (lrs!nPrdEstado = 2202 Or lrs!nPrdEstado = 2204 Or lrs!nPrdEstado = 2206) Then
                lsCalifGen = 4
            Else
                lsCalifGen = ""
                lsCalCMAC = IIf(IsNull(lrs!cCalCMAC), "", lrs!cCalCMAC)
                lsCalSistF = IIf(IsNull(lrs!cCalSistF), "", lrs!cCalSistF)
                lsCalHist = IIf(IsNull(lrs!cCalEval), "", lrs!cCalEval)
    
                'If IsNull(lrs!cFlagProv) Then ' ***
                '    lsCalObser = "-1"
                'Else
                '    lsCalObser = Trim(Val(lrs!cFlagProv))
                'End If
                If lrs!cRefinan = "R" Then
                    'lsCalifGen = lsCalSistF
                    If Trim(lsCalSistF) = "" Or Trim(lsCalSistF) <= "2" Then
                        lsCalifGen = lsCalSistF
                    ElseIf Not (lsCalSistF = "4" And lrs!nDiasAtrasoSF >= 1800) Then   'ALPA20160301
                        lsCalifGen = lsCalSistF - 1  ' Una Calificacion menor
                    End If
                ElseIf Trim(lsCalHist) <> "" Then
                    lsCalifGen = lsCalHist
                'ElseIf Trim(lsCalObser) <> -1 And lsCalObser >= lsCalCMAC Then
                '    lsCalifGen = lsCalObser  '  Calif Observado
                ElseIf Trim(lsCalSistF) = "" Then
                    lsCalifGen = lsCalCMAC  '  No tiene cred en Sist Financ
                ElseIf lsCalCMAC >= lsCalSistF Then
                    lsCalifGen = lsCalCMAC  ' Si la calif CMAC es mayor
                Else  ' Si la CalSistF es mayor
                    If Val(lsCalSistF) < 3 And lrs!Minorista > 0 Then ' No se alinean
                        lsCalifGen = lsCalCMAC
                    Else ' Se alinean  con un grado de discrepancia
                        'ARCV 19-06-2007
                        'lsCalifGen = lsCalSistF - 1  ' Una Calificacion menor
                        If lsCalSistF >= 3 Then
                            If Not (lsCalSistF = "4" And lrs!nDiasAtrasoSF >= 1800) Then 'ALPA20160301
                                lsCalifGen = lsCalSistF - 1  ' Una Calificacion menor
                            End If
                        Else
                            lsCalifGen = lsCalCMAC  'ARCV 19-06-2007 : Si no cumple poner de la CMACT
                        End If
                        '-------
                    End If
                End If
            End If
            If lsCalifGen < lsCalCMAC And Trim(lsCalHist) <> "" Then
                lsCalifGen = lsCalCMAC
            End If
            '*** Actualiza ColocCalifProv
            If Val(lsCalifGen) <= 0 Then lsCalifGen = "0"
            lsSQL = "UPDATE ColocCalifProv SET cCalGen = '" & lsCalifGen & "'" _
                  & "WHERE cPersCod ='" & Trim(lrs!cPersCod) & "'"
            loConec.Ejecutar lsSQL
            
            'ALPA20131127******************************************************************
            lsSQL = "UPDATE ColocCalifProv SET cCalCMAC = '" & lsCalifGen & "'" _
                  & "WHERE cPersCod ='" & Trim(lrs!cPersCod) & "' and cRefinan='R'"
            loConec.Ejecutar lsSQL
            '******************************************************************************
        
            lrs.MoveNext
        Loop
        Set lrs = Nothing
        'ALPA 20100730***************************
        lsSQL = "exec B2_stp_sel_CalificacionCambiarACastigo '" & Format(dFecha, "YYYY/MM/DD") & "'"
        loConec.Ejecutar lsSQL
        
        lsSQL = "Update ColocCalifProv "
        lsSQL = lsSQL & "set cCalGen"
        lsSQL = lsSQL & "=case "
        lsSQL = lsSQL & "    when cCalGen>(case Calif when '0' then '1' else Calif end) then cCalGen"
        lsSQL = lsSQL & "    else (case Calif when '0' then '1' else Calif end)"
        lsSQL = lsSQL & " end "
        lsSQL = lsSQL & "From "
        lsSQL = lsSQL & " ( "
        lsSQL = lsSQL & "select x.cCtaCod,case isnull(cCalNuev,'') when '' then cCalAnte else cCalNuev end Calif,cPersCod "
        lsSQL = lsSQL & "From "
        lsSQL = lsSQL & "( "
        lsSQL = lsSQL & " select distinct C.cCtaCod,CT.cCalGen cCalNuev, cCalAnte,cCtaCodRef,C.cPersCod "
        lsSQL = lsSQL & "from ColocCalifProv C "
        lsSQL = lsSQL & " left join DBConsolidada..ColocCalifProvTotal CT on C.cCtaCod=CT.cCtaCod "
        lsSQL = lsSQL & " and datediff(d,CT.dFecha,(select Max(dFecha) from DBConsolidada..ColocCalifProvTotal where cCtaCod=CT.cCtaCod))=0"
        lsSQL = lsSQL & " left join ("
        lsSQL = lsSQL & " select CR.cCtaCod,cCtaCodRef,cCalGen cCalAnte"
        lsSQL = lsSQL & " from ColocacRefinanc CR"
        lsSQL = lsSQL & "    inner join DBConsolidada..ColocCalifProvTotal CT2 on CR.cCtaCodRef=CT2.cCtaCod"
        lsSQL = lsSQL & "    and datediff(d,CT2.dFecha,(select Max(dFecha) from DBConsolidada..ColocCalifProvTotal where cCtaCod=CT2.cCtaCod))=0"
        lsSQL = lsSQL & " ) CR on C.cCtaCod=CR.cCtaCod"
        lsSQL = lsSQL & " where C.cRefinan='R'"
        lsSQL = lsSQL & " ) x"
        lsSQL = lsSQL & " ) y inner join ColocCalifProv P on y.cPersCod=P.cPersCod"
        loConec.Ejecutar lsSQL
        
        lsSQL = "Update ColocCalifProv "
        lsSQL = lsSQL & "Set cCalGen = x.cEvalCamb "
        lsSQL = lsSQL & "From"
        lsSQL = lsSQL & "("
        lsSQL = lsSQL & "    select distinct C.cPersCod, C.cCalEval cEvalCamb"
        lsSQL = lsSQL & "    From ColocCalifProv C inner join ColocEvalCalifDetalle CD on C.cCtaCod=CD.cCtaCod"
        lsSQL = lsSQL & "    Where C.cCalEval Is Not Null and C.cRefinan in ('R','N') and datediff(d,dEval,'" & Format(dFecha, "YYYY/MM/DD") & "')=0"
        lsSQL = lsSQL & "       and CD.cEvalObsDet not in ('Eval.Automat') "
        lsSQL = lsSQL & ") x inner join ColocCalifProv P on x.cPersCod=P.cPersCod"
        loConec.Ejecutar lsSQL
        '****************************************
        '*ALPA 20100910 Permanencia en la calificación
        lsSQL = "Update ColocCalifProv "
        lsSQL = lsSQL & "set nCalGenPerCon=case C.cCalgen when isnull(X.cCalGenPer,0) then isnull(X.nCalGenPerCon,0)+1 else 1 end, "
        lsSQL = lsSQL & "cCalGenPer = C.cCalgen "
        lsSQL = lsSQL & "From "
        lsSQL = lsSQL & "( "
        lsSQL = lsSQL & "select distinct C.cPersCod,isnull(CC.cCalGenPer,'0') cCalGenPer, "
        lsSQL = lsSQL & "isnull(CC.nCalGenPerCon,0) nCalGenPerCon "
        lsSQL = lsSQL & "from ColocCalifProv C "
        lsSQL = lsSQL & "LEFT join DBConsolidada..ColocCalifProvTotal CC on "
        lsSQL = lsSQL & "       C.cPersCod=CC.cPersCod and datediff(d,CC.dFecha,(select max(dFecha) from DBConsolidada..ColocCalifProvTotal))=0 "
        lsSQL = lsSQL & ") X inner join ColocCalifProv C on X.cPersCod=C.cPersCod "
        'End Permanencia en la calificación
        loConec.Ejecutar lsSQL
        
        'ALPA 20110827
        lsSQL = " Update colocCalifProv "
        lsSQL = lsSQL & " Set cCalgen = x.cCalgen "
        lsSQL = lsSQL & " From "
        lsSQL = lsSQL & " ( "
        lsSQL = lsSQL & " select cPersCod,max(cCalGen) cCalGen "
        lsSQL = lsSQL & " From colocCalifProv "
        lsSQL = lsSQL & " where nprdEstado in (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107,2201,2202,2203,2204,2205,2206) "
        lsSQL = lsSQL & " group by cPersCod "
        lsSQL = lsSQL & " ) X "
        lsSQL = lsSQL & " inner join colocCalifProv C on C.cPersCod=X.cPersCod and C.cCalGen<>X.cCalgen "
        loConec.Ejecutar lsSQL
        
        'JUEZ 20130205 Personas con creditos en otras entidades con dias atraso >1800 ya no alinean, se quedan con calCMAC ******
        lsSQL = " Select distinct (c.cPersCod) cPersCod,cCalCMAC,cCalSistF "
        lsSQL = lsSQL & "From ColocCalifProv c "
        lsSQL = lsSQL & "Inner Join Persona P on c.cPersCod=P.cPersCod "
        lsSQL = lsSQL & "Inner Join (Select distinct Cod_Edu "
        lsSQL = lsSQL & "            From DBConsolidada..RCCTotalDet R "
        lsSQL = lsSQL & "            Where Cod_Cuenta like '14[12][1456]%' and condicion>=1800 and Cod_Emp<>'00109' "
        lsSQL = lsSQL & "               and Cod_Edu in (Select distinct R.Cod_Edu "
        lsSQL = lsSQL & "                               From DBConsolidada..RCCTotal R "
        lsSQL = lsSQL & "                               Inner Join DBConsolidada..RCCTotalDet RD on R.Cod_Edu=RD.Cod_Edu and Cod_Emp='00109') "
        lsSQL = lsSQL & "            ) r on p.cPersCodSbs=r.Cod_Edu "
        lsSQL = lsSQL & "Where nPrdEstado NOT IN (2202,2206)"

        Set lrs = loConec.CargaRecordSet(lsSQL)
        If Not lrs.BOF And Not lrs.EOF Then
            Do While Not lrs.EOF
                lsSQL = "UPDATE ColocCalifProv SET cCalGen = '" & lrs!cCalCMAC & "' " _
                  & "WHERE cPersCod ='" & Trim(lrs!cPersCod) & "'"
                loConec.Ejecutar lsSQL
                lrs.MoveNext
            Loop
            Set lrs = Nothing
        End If
        'END JUEZ ***********************************************************************
        
        'ALPA20131127******************************************************************
        lsSQL = "UPDATE ColocCalifProv SET cCalCMAC = cCalGen " _
                & "WHERE cRefinan='R'"
        loConec.Ejecutar lsSQL
        '******************************************************************************
        'ALPA20151228******************************************************************
        lsSQL = "UPDATE ColocCalifProv SET nGarant=4 " _
                & "WHERE ((nCalGenPerCon>=36 and cCalGen=3 and nGarant=3) or (nCalGenPerCon>=24 and cCalGen=4 and (nGarant= 1 or nGarant=3))) "
        loConec.Ejecutar lsSQL
        '******************************************************************************
        Set loConec = Nothing
            
    End If
    
    Exit Sub
    
ErrorCalificacionGeneral:
    Err.Raise Err.Number, "Calificacion General", Err.Description
End Sub

Public Sub nCalificaCalculaProvision(ByVal psFecha As String)

Dim lsSQL As String
Dim lrs As New ADODB.Recordset
Dim lrs2 As New ADODB.Recordset 'ALPA20131226
Dim loConec As COMConecta.DCOMConecta
Dim loCalif As COMNCredito.NCOMColocEval
Dim lnProvision As Currency
Dim lmProGarAutoL(5)
Dim lmProGarMuyRR(5) As Double
Dim lmProGarPref(5) As Double
Dim lmProGarSinG(5) As Double
Dim lnSaldoCapitalSinDiferido As Currency 'ALPA20131226
Dim lnSaldoInteresSinDiferido As Currency 'ALPA20131226
Dim lnPorceCapitalSinDiferido As Currency 'ALPA20131226
Dim lnPorceInteresSinDiferido As Currency 'ALPA20131226
Dim lnCapitalSinDiferido As Currency 'ALPA20131226
On Error GoTo ErrorCalculaProvision

lmProGarAutoL(0) = 1
lmProGarMuyRR(0) = 1: lmProGarMuyRR(1) = 1.25: lmProGarMuyRR(2) = 6.25: lmProGarMuyRR(3) = 15: lmProGarMuyRR(4) = 30
lmProGarPref(0) = 1: lmProGarPref(1) = 2.5: lmProGarPref(2) = 12.5: lmProGarPref(3) = 30: lmProGarPref(4) = 60
lmProGarSinG(0) = 1: lmProGarSinG(1) = 5:   lmProGarSinG(2) = 25:   lmProGarSinG(3) = 60: lmProGarSinG(4) = 100

Set loConec = New COMConecta.DCOMConecta
loConec.AbreConexion

' Perdida en los ultimos 24 meses ==> Utiliza 100 % (Lopez) 2005/12/01
'lsSQL = "Update ColocCalifProv set nGarant = 4  " _
'    & "Where nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & "," _
'    & gColocEstRefNorm & "," & gColocEstRefMor & "," & gColocEstRefVenc & "," _
'    & "2201,2205)  And cCalGen ='4' " _
'    & "And nSaldoCap > 0 and nGarant <> 4 " _
'    & "And cCtaCod in ( " _
'    & "    Select cCtaCod from ( " _
'    & "            Select cCtaCod, count(cCtaCod) veces " _
'    & "            From ColocCalifProvTotal Where dFecha >= dateadd(m,-24,'" & Format(psFecha, "yyyy/mm/dd") & "')  " _
'    & "            And cCalGen ='4' " _
'    & "    Group by cCtaCod " _
'    & "    Having Count(cCtaCod) >= 24 " _
'    & "    ) x " _
'    & " ) "
'loConec.Ejecutar lsSQL
'ALPA 20160308*********************
lsSQL = " Update ColocCalifProv"
lsSQL = lsSQL & " set nMontoCapRefinan=0,nMontoIntRefinan=0"
loConec.Ejecutar lsSQL
'**********************************
'ALPA 20081209**********************************************************************
''lsSQL = "Select CA.cPersCod, CA.cCtaCod,  CA.nPrdEstado, CA.cRefinan, CA.nSaldoCap, " _
''      & "CA.cCalGen, CA.nMontoApr, CA.nGarPref, CA.nGarant  " _
''      & " From ColocCalifProv CA  " _
''      & " ORDER BY CA.cPersCod "
'JUEZ 20160130 *****************************************
lsSQL = "exec stp_sel_ReporteInteresDiferidoAmpliados"
loConec.Ejecutar lsSQL
'END JUEZ **********************************************

lsSQL = "exec stp_sel_ListaCalculaProvision"
Set lrs = loConec.CargaRecordSet(lsSQL)

Do While Not lrs.EOF
    'If Trim(lrs!cCalGen) = "" Then lrs!cCalGen = "4"
    'ALPA201312********************
    lnSaldoCapitalSinDiferido = 0
    lnPorceInteresSinDiferido = 0
    lnSaldoCapitalSinDiferido = 0
    lnSaldoInteresSinDiferido = 0
    lnCapitalSinDiferido = lrs!nSaldoCap
'    If lrs!cRefinan = "R" Then
        lsSQL = "exec stp_sel_InteresDiferido '" & Trim(lrs!cCtaCod) & "','" & Format(psFecha, "YYYY/MM/DD") & "'"
        Set lrs2 = loConec.CargaRecordSet(lsSQL)
        If Not (lrs2.BOF Or lrs2.EOF) Then
            Do While Not lrs2.EOF
                lnSaldoCapitalSinDiferido = lrs2!nSaldo
                lnSaldoInteresSinDiferido = lrs2!nInteres
                lrs2.MoveNext
            Loop
            lnCapitalSinDiferido = lnSaldoCapitalSinDiferido
            lsSQL = "UPDATE ColocCalifProv SET nMontoCapRefinan= " & lnSaldoCapitalSinDiferido & ","
            lsSQL = lsSQL & " nMontoIntRefinan=" & lnSaldoInteresSinDiferido
            lsSQL = lsSQL & " WHERE cPersCod ='" & Trim(lrs!cPersCod) & "' And cCtaCod ='" & Trim(lrs!cCtaCod) & "' "
            loConec.Ejecutar lsSQL
        Else
            lnCapitalSinDiferido = lrs!nSaldoCap
        End If
'        lsSQL = "UPDATE ColocCalifProv SET nMontoCapRefinan= " & lnSaldoCapitalSinDiferido & ","
'        lsSQL = lsSQL & " nMontoIntRefinan=" & lnSaldoInteresSinDiferido
'        lsSQL = lsSQL & " WHERE cPersCod ='" & Trim(lrs!cPersCod) & "' And cCtaCod ='" & Trim(lrs!cCtaCod) & "' "
'        loConec.Ejecutar lsSQL
'    End If
    '******************************
    lnProvision = 0
    'ALPA 20081205*************************************************************
    If lrs!cCalGen = 0 Then
        If lrs!nGarant = 1 Then
            'lnProvision = lrs!nSaldoCap * lrs!nProvisionA
            lnProvision = lnCapitalSinDiferido * lrs!nProvisionA
            
        Else
            'lnProvision = lrs!nSaldoCap * lrs!nProvision
            lnProvision = lnCapitalSinDiferido * lrs!nProvision
        End If
    Else
    '**************************************************************************
        If (((lrs!nCalGenPerCon >= 24 And lrs!nCalGenPerCon <= 35) And lrs!cCalGen = 4) Or ((lrs!nCalGenPerCon >= 36) And lrs!cCalGen = 3)) And (lrs!nGarant = 1 Or lrs!nGarant = 3) And CInt(lrs!cCalSistF) <> "5" And (CInt(lrs!cCalSistF) - 1) = CInt(lrs!cCalCMAC) Then
            'lnProvision = lrs!nSaldoCap * lmProGarSinG(IIf(Trim(lrs!cCalGen) = "", 4, lrs!cCalGen)) ' Porcentaje de Provision
            lnProvision = lnCapitalSinDiferido * lmProGarSinG(IIf(Trim(lrs!cCalGen) = "", 4, lrs!cCalGen)) ' Porcentaje de Provision
        Else
            If lrs!nGarant = 1 Then
                'lnProvision = lrs!nSaldoCap * lmProGarAutoL(0)
                lnProvision = lnCapitalSinDiferido * lmProGarAutoL(0)
            ElseIf lrs!nGarant = 2 Then
                'lnProvision = lrs!nSaldoCap * lmProGarMuyRR(IIf(Trim(lrs!cCalGen) = "", 4, lrs!cCalGen)) ' Porcentaje de Provision
                lnProvision = lnCapitalSinDiferido * lmProGarMuyRR(IIf(Trim(lrs!cCalGen) = "", 4, lrs!cCalGen)) ' Porcentaje de Provision
            ElseIf lrs!nGarant = 3 Then
                'lnProvision = lrs!nSaldoCap * lmProGarPref(IIf(Trim(lrs!cCalGen) = "", 4, lrs!cCalGen)) ' Porcentaje de Provision
                lnProvision = lnCapitalSinDiferido * lmProGarPref(IIf(Trim(lrs!cCalGen) = "", 4, lrs!cCalGen)) ' Porcentaje de Provision
            ElseIf lrs!nGarant = 4 Then
                'lnProvision = lrs!nSaldoCap * lmProGarSinG(IIf(Trim(lrs!cCalGen) = "", 4, lrs!cCalGen)) ' Porcentaje de Provision
                lnProvision = lnCapitalSinDiferido * lmProGarSinG(IIf(Trim(lrs!cCalGen) = "", 4, lrs!cCalGen)) ' Porcentaje de Provision
            End If
        End If
'        'If Mid(lrs!cCtaCod, 6, 3) = "423" Then 'Si es MiVivienda
'        If lrs!cTpoCredCod = "853" Or lrs!cTpoCredCod = "854" Then 'Si es MiVivienda
'            lnProvision = (2 / 3) * lnProvision
'        End If
    'ALPA 20081205*******************************************************
    End If
    '********************************************************************
    'ALPA20131230***********************************************************************
    If lrs!cTpoCredCod = "853" Or lrs!cTpoCredCod = "854" Then 'Si es MiVivienda
        lnProvision = (2 / 3) * lnProvision
    End If
    '********************************************************************
    'Redondeo
    lnProvision = Format(lnProvision / 100, "#0.00")
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000100", lrs!cTpoProdCod) Then
    'If lrs!cTpoProdCod = "514" Then 'Si es Carta Fianza
    '**ARLO20180712 ERS042 - 2018
        'lnProvision = Round((lnProvision / 2), 2) 'Comento JOEP20180622 Acta 122-2018
        'Agrego JOEP20180622 Acta 122-2018
        If lrs!nTpRiesgo = 1 Then
            lnProvision = Round((lnProvision / 2), 2) 'Hacer
        Else
            lnProvision = Round(lnProvision, 2) 'Mixto y Dar
        End If
        'Agrego JOEP20180622 Acta 122-2018
    End If
    lsSQL = "UPDATE ColocCalifProv SET nProvision= " & lnProvision _
          & " WHERE cPersCod ='" & Trim(lrs!cPersCod) & "' And cCtaCod ='" & Trim(lrs!cCtaCod) & "' "
    loConec.Ejecutar lsSQL
    lrs.MoveNext
Loop
lrs.Close
Set lrs = Nothing
Exit Sub

ErrorCalculaProvision:
    Err.Raise Err.Number, "Calcula Provision", Err.Description
End Sub
'ALPA20131130
'Public Sub CalculaProvisionRCC(ByVal psRutaServer As String, ByVal pnTipoCambio As Double, ByVal pdFecha As Date)
Public Sub CalculaProvisionRCC(ByVal psRutaServer As String, ByVal pnTipoCambio As Double, ByVal pdFecha As Date)
Dim lsSQL As String
Dim lrs As New ADODB.Recordset
Dim loConec As COMConecta.DCOMConecta
Dim lnProvisionSinDiferido As Currency 'ALPA20131226
Dim lrs2 As New ADODB.Recordset 'ALPA20131226

On Error GoTo ErrorCalculaProvision

'ALPA20131127*****************************
Set loConec = New COMConecta.DCOMConecta
loConec.AbreConexion
'Para actualizar el monto a "0-Cero" de la provision RCC
lsSQL = "UPDATE ColocCalifProv SET nProvisionRCC=0"
loConec.Ejecutar lsSQL
'*****************************************

'lsSQL = "SELECT CC.cCtaCod,nProvision=ROUND(" & _
'        " SUM( " & _
'        " (CASE WHEN GCC.nMoneda= 2 " & _
'        " THEN (GCC.nMontoGrava/CC.nMontoDesemb)*CC.nSaldoCap ELSE " & _
'        " ((GCC.nMontoGrava/CC.nMontoDesemb)*CC.nSaldoCap)/" & CStr(pnTipoCambio) & "  END)* (nParamValor/100) " & _
'        ") " & _
'        " ,2),CP.cRefinan,CP.cTpoCredCod " & _
'        " From " & _
'        " ColocCalifProv CP INNER JOIN " & psRutaServer & "CreditoConsol CC ON CP.cCtaCod=CC.cCtaCod " & _
'        " INNER JOIN ColocacCred CCC ON CCC.cCtaCod = CC.cCtaCod" & _
'        " INNER JOIN " & psRutaServer & "GarantCredConsol GCC ON CC.cCtaCod=GCC.cCtaCod " & _
'        " INNER JOIN Garantias GC ON GC.cNumGarant=GCC.cNumGarant Inner Join PersGarantia PG ON GC.cNumGarant =   PG.cNumgarant AND PG.nRelacion = 1" & _
'        " INNER JOIN ColocParametro ON nParamVar = ( " & _
'        " CASE WHEN GC.nGarTpoRealiz=1  THEN 100 WHEN CP.nGarant= 3 THEN 101 ELSE 102 END) " & _
'        " WHERE CP.cCalGen=0 AND ISNULL(CCC.nCodRCC,0)=2 AND SUBSTRING(CC.cCtaCod,9,1)='2' " & _
'        " GROUP BY CC.cCtaCod,CP.cRefinan,CP.cTpoCredCod ORDER BY CC.cCtaCod"
lsSQL = "exec stp_sel_ProvisionRCC '" & Format(pdFecha, "YYYY/MM/DD") & "'," & pnTipoCambio
'ALPA20131127
'Set loConec = New COMConecta.DCOMConecta
'loConec.AbreConexion

Set lrs = loConec.CargaRecordSet(lsSQL)

Do While Not lrs.EOF
    'ALPA201312********************
    lnProvisionSinDiferido = lrs!nProvision
'    If lrs!cRefinan = "R" Then
        lsSQL = "exec stp_sel_InteresDiferido '" & Trim(lrs!cCtaCod) & "','" & Format(pdFecha, "YYYY/MM/DD") & "'"
        Set lrs2 = loConec.CargaRecordSet(lsSQL)
        If Not (lrs2.BOF Or lrs2.EOF) Then
            Do While Not lrs2.EOF
                lnProvisionSinDiferido = lrs2!nProvisionN
                lrs2.MoveNext
            Loop
        End If
'    Else
'        lnProvisionSinDiferido = lrs!nProvision
'    End If
    '******************************
    'Actualiza ProvisionRCC
    
    If lrs!cTpoCredCod = "853" Or lrs!cTpoCredCod = "854" Then 'Si es MiVivienda
        lnProvisionSinDiferido = (2 / 3) * lnProvisionSinDiferido
    End If
    
    lsSQL = "UPDATE ColocCalifProv SET nProvisionRCC= " & lnProvisionSinDiferido
    lsSQL = lsSQL & " WHERE cCtaCod ='" & Trim(lrs!cCtaCod) & "' "
    loConec.Ejecutar lsSQL
    
    lrs.MoveNext
Loop
lrs.Close
Set lrs = Nothing
Exit Sub

ErrorCalculaProvision:
    Err.Raise Err.Number, "Calcula Provision", Err.Description
End Sub



Public Sub VerificaTablaTemporal(ByVal psCodUser As String, _
                                    ByRef psTablaTMP As String)
Dim rs As ADODB.Recordset
Dim Sql As String
Dim oCon As COMConecta.DCOMConecta

psTablaTMP = "TMPPAGORFA" & psCodUser

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion

Set rs = oCon.CargaRecordSet("select * from sysobjects where name like '%" & psTablaTMP & "%'")
If Not rs.EOF And Not rs.BOF Then
    Sql = "DROP TABLE " & psTablaTMP
    oCon.Ejecutar Sql
End If
rs.Close
oCon.CierraConexion
Set rs = Nothing
Set oCon = Nothing

End Sub

Public Sub ObtieneTablaTemporal(ByVal psCodUser As String, _
                                    ByRef psTablaTMP As String, _
                                    ByRef pbExiste As Boolean)
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

psTablaTMP = "TMPPAGORFA" & psCodUser

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion

pbExiste = False
Set rs = oCon.CargaRecordSet("select * from sysobjects where name like '%" & psTablaTMP & "%'")
If Not rs.EOF And Not rs.BOF Then
    pbExiste = True
End If
rs.Close
Set rs = Nothing
Set oCon = Nothing
End Sub


Public Sub CalificaRFANew(ByVal psCodUser As String, _
                            ByVal psServerConsol As String, _
                            ByVal pdFecha As Date, _
                            ByRef psTablaTMP As String)

Dim lsSQL As String
Dim rs As New ADODB.Recordset
Dim loConec As COMConecta.DCOMConecta
Dim lrDat As ADODB.Recordset
Dim lrDatCalif As ADODB.Recordset
Dim lsCalMaxActual As String, lsCalAnterior As String
Dim lsCalRFA As String
'Dim lsTablaTMP As String
Dim bExisteTablaTMP As Boolean

On Error GoTo ErrorCalificaRFANew

Set loConec = New COMConecta.DCOMConecta
    
'    VerificaTablaTemporal psCodUser, psTablaTMP
    Call ObtieneTablaTemporal(psCodUser, psTablaTMP, bExisteTablaTMP)
    
    loConec.AbreConexion
    loConec.ConexionActiva.CommandTimeout = 300
    If bExisteTablaTMP Then
        lsSQL = "DROP TABLE " & psTablaTMP
    Else
        lsSQL = ""
    End If
    lsSQL = lsSQL & "         SELECT    C.CCTACOD, P.CPERSCOD, P.CPERSNOMBRE, CC.CRFA, C.NDIASATRASO, "
    lsSQL = lsSQL & "           ISNULL(DATOS.nDiasAtrPag,C.NDIASATRASO) AS nDiasAtrPag, datos.dFecPag, datos.dFecVenc "
    lsSQL = lsSQL & "  Into " & psTablaTMP
    lsSQL = lsSQL & "  FROM    " & psServerConsol & "CREDITOCONSOL C"
    lsSQL = lsSQL & "           FULL OUTER JOIN"
    lsSQL = lsSQL & "               (SELECT P.cCtaCod,  P.dFecVenc, convert(datetime,Pago.dFecPag) AS dFecPag,"
    lsSQL = lsSQL & "                       datediff(Day, P.dFecVenc, convert(DateTime, Pago.dFecPag)) As nDiasAtrPag"
    lsSQL = lsSQL & "                From"
    lsSQL = lsSQL & "                   ("
    lsSQL = lsSQL & "                       SELECT  DISTINCT MD.cCtaCod,M.CMOVNRO,MD.nNroCuota,"
    lsSQL = lsSQL & "                               CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,5,2) + '/' + SUBSTRING(M.CMOVNRO,7,2) + '/' + SUBSTRING(M.CMOVNRO,1,4))  AS DFECPAG"
    lsSQL = lsSQL & "                       FROM    MOV M "
    lsSQL = lsSQL & "                               JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO"
    lsSQL = lsSQL & "                               JOIN MOVCOLDET MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.COPECOD = MC.COPECOD AND MD.CCTACOD = MC.CCTACOD"
    lsSQL = lsSQL & "                       WHERE   (MC.COPECOD LIKE '100[234567]%' OR MC.COPECOD LIKE '12[16][012]%') AND NMOVFLAG = 0"
    lsSQL = lsSQL & "                               AND MD.CCTACOD IN (SELECT CCTACOD FROM COLOCACCRED WHERE CRFA IN ('RFA','RFC','DIF'))"
    lsSQL = lsSQL & "                               AND LEFT(M.CMOVNRO,6)<= '" & Format(pdFecha, "yyyymm") & "'"
    lsSQL = lsSQL & "                               AND M.CMOVNRO IN (  SELECT MAX(M1.CMOVNRO)"
    lsSQL = lsSQL & "                                                   FROM    MOV M1"
    lsSQL = lsSQL & "                                                           JOIN MOVCOL MC1 ON MC1.NMOVNRO = M1.NMOVNRO"
    lsSQL = lsSQL & "                                                   WHERE   M1.NMOVFLAG=0 AND MC1.CCTACOD=MC.CCTACOD AND LEFT(M1.CMOVNRO,6)<= '" & Format(pdFecha, "yyyymm") & "'"
    lsSQL = lsSQL & "                                                   AND (MC1.COPECOD LIKE '100[234567]%' OR MC1.COPECOD LIKE '12[16][012]%'))"
    lsSQL = lsSQL & "                       ) AS PAGO"
    lsSQL = lsSQL & "               JOIN " & psServerConsol & "PLANDESPAGCONSOL P ON P.CCTACOD = PAGO.CCTACOD AND P.cNroCuo =  PAGO.NNROCUOTA  AND P.nTipo = 1 and P.nEstado = 1"
    lsSQL = lsSQL & "               ) AS DATOS ON DATOS.CCTACOD = C.CCTACOD"
    lsSQL = lsSQL & "           JOIN COLOCACCRED CC ON CC.CCTACOD = C.CCTACOD"
    lsSQL = lsSQL & "           JOIN PRODUCTOPERSONA R ON R.CCTACOD = C.CCTACOD"
    lsSQL = lsSQL & "           JOIN PERSONA P ON P.CPERSCOD = R.CPERSCOD AND R.NPRDPERSRELAC=20"
    lsSQL = lsSQL & "    WHERE  CC.CRFA IN ('RFA','RFC','DIF') AND"
    lsSQL = lsSQL & "           C.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107,2061,2060) "

    loConec.ConexionActiva.Execute lsSQL ' Ejecutar (lsSQL)
    
    lsSQL = "           SELECT  CPERSCOD, CPERSNOMBRE, AVG(nDiasAtrPag) AS nDiasAtrPag "
    lsSQL = lsSQL & "   From "
    lsSQL = lsSQL & "       (   SELECT *"
    lsSQL = lsSQL & "           FROM   " & psTablaTMP & " P"
    lsSQL = lsSQL & "           WHERE   NOT EXISTS (SELECT P1.CPERSCOD"
    lsSQL = lsSQL & "                               FROM " & psTablaTMP & " P1"
    lsSQL = lsSQL & "                               WHERE P1.CPERSCOD = P.CPERSCOD AND P1.nDiasAtrPag>0  )"
    lsSQL = lsSQL & "                   AND NOT EXISTS (SELECT P1.CPERSCOD"
    lsSQL = lsSQL & "                                   FROM " & psTablaTMP & " P1"
    lsSQL = lsSQL & "                                   WHERE P1.CPERSCOD = P.CPERSCOD AND P1.nDiasAtraso>0 )"
    lsSQL = lsSQL & "                   AND NOT EXISTS (SELECT P1.CPERSCOD"
    lsSQL = lsSQL & "                                   FROM " & psTablaTMP & " P1"
    lsSQL = lsSQL & "                                   WHERE P1.CPERSCOD = P.CPERSCOD AND P1.dFecPag is null )"
    lsSQL = lsSQL & "                               ) AS X "
    lsSQL = lsSQL & "   GROUP BY CPERSCOD, CPERSNOMBRE "
    lsSQL = lsSQL & "   ORDER BY CPERSNOMBRE "

    Set lrDat = loConec.CargaRecordSet(lsSQL)
    If Not (lrDat.BOF And lrDat.EOF) Then
        Do While Not lrDat.EOF
            
            'If lrDat!cPersCod = "1080100751979" Then Stop
            
            lsSQL = "UPDATE ColocCalifProv " _
                & " SET cCalRFA = '0' " _
                & "WHERE cPersCod ='" & Trim(lrDat!cPersCod) & "'"
            
            loConec.Ejecutar lsSQL

            lrDat.MoveNext
        Loop
    End If
    Set lrDat = Nothing
    
Set loConec = Nothing

Exit Sub

ErrorCalificaRFANew:
    Err.Raise Err.Number, "Califica RFANew", Err.Description

End Sub

Public Sub EndeudamientoSistFinanc(ByVal pnTipoCambio As Double, _
                                ByVal psBDRCC As String)

' Reclasifica los Creditos MES
Dim lsSQL As String
Dim rs As ADODB.Recordset
Dim loConec As COMConecta.DCOMConecta

On Error GoTo ErrorEndeudamientoSistFinanc

lsSQL = " Select CA.cPersCod, " _
    & " Isnull( Sum(case when substring(cod_cuenta,3,1) = '2' And rd.cod_cuenta like '14_[1456]%' then Val_saldo / " & pnTipoCambio _
    & "        End),0) nEndeudaDol " _
    & " FROM (select C.cPersCod From ColocCalifProv c where c.nPrdEstado in (2020,2021,2022,2030,2031,2032) group by c.cPersCod) ca " _
    & " INNER join Persona p on CA.cPersCod = P.cPersCod " _
    & " INNER join PersId PI on P.cPersCod = PI.cPersCod " _
    & " INNER join " & psBDRCC & "rccTotal r on LTRIM(RTRIM(r.Cod_Doc_Id)) = LTRIM(RTRIM(PI.cPersIDnro)) " _
    & " INNER join " & psBDRCC & "rccTotalDet rd on LTRIM(RTRIM(r.Cod_Sbs)) = LTRIM(RTRIM(rd.Cod_Sbs)) " _
    & " WHERE  PI.cPersIDTpo = " & gPersIdDNI & " " _
    & " AND( LTRIM(RTRIM(pi.cPersIDnro)) <>'' or pi.cPersIDnro <>null ) " _
    & " And rd.cod_cuenta like '14_[1456]%' " _
    & " Group by CA.cPersCod "

 Set loConec = New COMConecta.DCOMConecta
 loConec.AbreConexion
 Set rs = loConec.CargaRecordSet(lsSQL)

Do While Not rs.EOF

    lsSQL = "UPDATE ColocCalifProv SET nMESReclas =  " & Format(rs!nEndeudaDol, "#0.00") _
        & "WHERE cPersCod='" & Trim(rs!cPersCod) & "'"
    
    loConec.Ejecutar lsSQL
    
    rs.MoveNext
Loop
rs.Close

'*** Personas Juridicas
lsSQL = " Select CA.cPersCod, " _
    & " Isnull( Sum(case when substring(cod_cuenta,3,1) = '2' And rd.cod_cuenta like '14_[1456]%' then Val_saldo / " & pnTipoCambio _
    & "        End),0) nEndeudaDol " _
    & " FROM (select C.cPersCod From ColocCalifProv c where c.nPrdEstado in (2020,2021,2022,2030,2031,2032) group by c.cPersCod) ca " _
    & " INNER join Persona p on CA.cPersCod = P.cPersCod " _
    & " INNER join PersId PI on P.cPersCod = PI.cPersCod " _
    & " INNER join " & psBDRCC & "rccTotal r on LTRIM(RTRIM(r.Cod_Doc_Trib)) = LTRIM(RTRIM(PI.cPersIDnro)) " _
    & " INNER join " & psBDRCC & "rccTotalDet rd on LTRIM(RTRIM(r.Cod_Sbs)) = LTRIM(RTRIM(rd.Cod_Sbs)) " _
    & " WHERE PI.cPersIDTpo = " & gPersIdRUC & " " _
    & " AND ( LTRIM(RTRIM(pi.cPersIDnro)) <>'' or pi.cPersIDnro <>null ) " _
    & " And rd.cod_cuenta like '14_[1456]%' " _
    & " Group by CA.cPersCod "

Set rs = loConec.CargaRecordSet(lsSQL)

Do While Not rs.EOF

    lsSQL = "UPDATE ColocCalifProv SET nMESReclas =  " & Format(rs!nEndeudaDol, "#0.00") _
        & "WHERE cPersCod='" & Trim(rs!cPersCod) & "'"
    
    loConec.Ejecutar (lsSQL)
    
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
Set loConec = Nothing
Exit Sub

ErrorEndeudamientoSistFinanc:
    Err.Raise Err.Number, "Endeudamiento Sistema Financiero", Err.Description
End Sub

Public Function GetFechaAlin(ByVal psBDRCC As String) As String
Dim Sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

GetFechaAlin = ""
oCon.AbreConexion
Sql = "SELECT MAX(FEC_REP) as dFecAlin FROM " & psBDRCC & "RCCTOTAL"
Set rs = oCon.CargaRecordSet(Sql)
If Not rs.EOF And Not rs.BOF Then
    GetFechaAlin = Format(rs!dFecAlin, "dd/mm/yyyy")
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing

End Function

Public Sub nCalificaRiesgoCredCambiario()
' Calificacion x Riesgo crediticio Cambiario
'0=No identificado, 1= No Expuesto, 2=Expuesto
' 2007-01
Dim lsSQL As String
Dim rs As New ADODB.Recordset
Dim loConec As COMConecta.DCOMConecta
Dim lrDat As ADODB.Recordset

On Error GoTo ErrorCalificaRCambiario

Set loConec = New COMConecta.DCOMConecta
    loConec.AbreConexion
    
    '*** Metodologia CUSCO
    '---******** Los valores que le han asignado en el Negocio
    lsSQL = "Update ColocCalifProv set nRCambiario = y.nCodRCC "
    'lssql= lssql &   "select y.nCodRCC,  isnull(y.nCodRCC,0), * "
    lsSQL = lsSQL & "From ColocCalifProv c join "
    lsSQL = lsSQL & "ColocacCred y on c.cctacod = y.cCtacod "
    loConec.Ejecutar lsSQL
    
    '-- Prendario ==> Todos son No Expuestos (1)
    lsSQL = "Update coloccalifProv set nRCambiario = 1 where cTpoCredCod='755' "
    loConec.Ejecutar lsSQL
    
    '-- Soles y no identificados ==> No expuesto(1)
    lsSQL = "Update coloccalifProv set nRCambiario = 1 "
    lsSQL = lsSQL & "Where Substring(cctacod,9,1)='1' and nRCambiario is null "
    loConec.Ejecutar lsSQL
    
    '-- Dolares Comerciales y no identificados ==> Expuesto (2)
    lsSQL = "Update ColoccalifProv set nRCambiario = 2 "
    lsSQL = lsSQL & "Where substring(cctacod,9,1)='2' and substring(cctacod,6,1)='1' and nRCambiario is null "
    loConec.Ejecutar lsSQL
    
    '-- Resto ==>  0(No Identificado)
    lsSQL = "Update coloccalifProv set nRCambiario = 0 Where nRCambiario is null "
    loConec.Ejecutar lsSQL

    '-- Calificacion por Cliente
    '--- Si tien un expuesto >> todos son expuestos (2)
    lsSQL = "Update ColocCalifProv set nRCambiario = 2 "
    lsSQL = lsSQL & "Where cperscod in (Select cPersCod from ColocCalifProv where nRCambiario = 2) "
    loConec.Ejecutar lsSQL
    
    '--- Si tiene algun No identificado Todos los que tenga son No Identificados
    lsSQL = "Update ColocCalifProv set nRCambiario = 0 "
    lsSQL = lsSQL & "Where cperscod in (Select cPersCod from ColocCalifProv where nRCambiario = 0) "
    loConec.Ejecutar lsSQL
    
    '--- Si son comerciales y Tienen 0 >> 2 Expuesto
    lsSQL = "Update coloccalifProv set nRCambiario = 2 "
    lsSQL = lsSQL & "Where cPerscod in(Select cPersCod from ColocCalifProv where nRCambiario =0  and substring(cctacod,6,1)='1' ) "
    loConec.Ejecutar lsSQL
    
    '--- Verifica que el cliente tenga un solo indicador pro persona
'    select cperscod, nrcambiario, count(nrCambiario)
'    From coloccalifProv
'    group by cperscod, nrcambiario
'    Having Count(Distinct(nrcambiario)) > 1


'    ---Verifica - x Tipo de RCambiario
'    select substring(cctacod,9,1) Moneda,substring(cctacod,6,1) Producto ,nrcambiario, count(cctacod), sum(nsaldocap)
'    from coloccalifprov where cCalGen = '0'
'    group by substring(cctacod,9,1), substring(cctacod,6,1) , nrcambiario
'    order by substring(cctacod,9,1), substring(cctacod,6,1) , nrcambiario


    '---**** REporte RCambiario
    '---Verifica - x Tipo de RCambiario
    'select substring(cctacod,9,1) Moneda, substring(cctacod,6,1) Producto,
    'substring(cctacod,4,2) Agencia, count(cctacod), sum(nsaldocap), sum(nProvision)
    'from coloccalifprov where cCalGen = '0'  and nrcambiario= 0
    'group by substring(cctacod,9,1), substring(cctacod,6,1) , substring(cctacod,4,2)
    'order by substring(cctacod,9,1), substring(cctacod,6,1) , substring(cctacod,4,2)
    '---** Report hoja trabajo
    'select substring(cctacod,9,1) Moneda, substring(cctacod,6,1) Producto, cCalGen Calific,
    'substring(cctacod,4,2) Agencia, count(cctacod) NroCredito, sum(nsaldocap), sum(nProvision)
    'from coloccalifprov where nprdestado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201)
    'group by substring(cctacod,9,1), substring(cctacod,6,1) , cCalGen, substring(cctacod,4,2)
    'order by substring(cctacod,9,1), substring(cctacod,6,1) , cCalGen, substring(cctacod,4,2)
    Set loConec = Nothing
    Exit Sub

ErrorCalificaRCambiario:
    Err.Raise Err.Number, "Califica Riesgo Cambiario", Err.Description
End Sub


Public Sub nCalificaMetodologiaInterna(ByVal psServerConsol As String, ByVal psFecha As String)
' Calificacion Metodologia Interna
' 2007-02-24
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim loConec As COMConecta.DCOMConecta
Dim lrDat As ADODB.Recordset
 
On Error GoTo Error
 
Set loConec = New COMConecta.DCOMConecta
    loConec.AbreConexion
    
    '*************************** metodologia
    
    'ARCV 05-06-2007: Calificacion Interna CMAC-Maynas
    
    'Paso 1 : Cargar la asignada por el sistema
    Sql = "Update colocCalifProv set cCalifInterna = cCalNor"
    loConec.Ejecutar Sql
    
    'Paso 2 : Califica los refinanciados
    'Sql = "Update colocCalifProv set cCalifInterna = "
    'Sql = Sql & " case when y.nDiasAtraso <= 0 then '1' "
    'Sql = Sql & "     when y.nDiasAtraso >= 1 and y.nDiasAtraso <=9 then '2' "
    'Sql = Sql & "     when y.nDiasAtraso >= 10 and y.nDiasAtraso <=30 then '3' "
    'Sql = Sql & "     when y.nDiasAtraso > 30  then '4' "
    'Sql = Sql & " End "
    'Sql = Sql & "from ColocCalifProv x "
    'Sql = Sql & " Join "
    'Sql = Sql & "(select c.cCtaCod,c.nDiasAtraso "
    'Sql = Sql & "from dbconsolidada.dbo.CreditoConsol c "
    'Sql = Sql & "Where c.nPrdEstado in (2030,2031,2032) "
    'Sql = Sql & ") y "
    'Sql = Sql & " on x.cCtaCod = y.cCtaCod "
    'loConec.Ejecutar Sql

    '--** Paso 3  La mayor por cliente
    Sql = "update colocCalifProv set cCalifInterna = y.Mayor "
    Sql = Sql & "from colocCalifProv x "
    Sql = Sql & "Join "
    Sql = Sql & "(select cPersCod, count(*) Nro, max(cCalifInterna) Mayor  "
    Sql = Sql & "from ColocCalifProv group by cPersCod) y "
    Sql = Sql & "on x.cPersCod=y.cPersCod  "
    loConec.Ejecutar Sql
    
'    '---**** Paso 1
'    Sql = "Update colocCalifProv set cCalifInterna = "
'    Sql = Sql & "case when y.AtrasoProm >= 0 and y.AtrasoProm <=3 then '1' "
'    Sql = Sql & "     when y.AtrasoProm >= 4 and y.AtrasoProm <=8 then '2' "
'    Sql = Sql & "     when y.AtrasoProm >= 9 and y.AtrasoProm <=30 then '3' "
'    Sql = Sql & "     when y.AtrasoProm >= 31 and y.AtrasoProm <=60 then '4' "
'    Sql = Sql & "     when y.AtrasoProm >= 61 and y.AtrasoProm <=120 then '5' "
'    Sql = Sql & "     when y.AtrasoProm >= 120 and y.AtrasoProm <=99999 then '6' "
'    Sql = Sql & " End "
'    'sql = sql & "select * "
'    Sql = Sql & "from ColocCalifProv x "
'    Sql = Sql & " Join "
'    Sql = Sql & "(select c.cCtaCod, round(c.nDiasAtrasoAcum / CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' then 1 else "
'    Sql = Sql & "                                                  case when count(case when p.dFecVenc <='" & Format(psFecha, "yyyy/mm/dd") & "' then p.cCtaCod end) = 0 then 1 "
'    Sql = Sql & "                                                       else count(case when p.dFecVenc <='" & Format(psFecha, "yyyy/mm/dd") & "' then p.cCtaCod end) end  END ,0) AtrasoProm "
'    'sql = sql & "c.nDiasAtrasoAcum , count(case when p.dFecVenc <='" & Format(psFecha, "yyyy/mm/dd") & "' then p.cCtaCod end) Nro "
'    Sql = Sql & "from dbconsolidada.dbo.CreditoConsol c "
'    Sql = Sql & "LEFT join dbconsolidada.dbo.PlandespagConsol p "
'    Sql = Sql & "on c.cCtaCod = p.cCtaCod "
'    Sql = Sql & "Where c.nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) "
'    Sql = Sql & "and isnull(p.nTipo, 1) = 1 "
'    Sql = Sql & "group by c.cCtaCod, c.nDiasAtrasoAcum "
'    Sql = Sql & ") y "
'    Sql = Sql & " on x.cCtaCod = y.cCtaCod "
'    loConec.Ejecutar Sql
'    '--** Paso 2
'    Sql = "update colocCalifProv set cCalifInterna = '6' where nPrdEstado in (2201,2205,2202,2206) "
'    loConec.Ejecutar Sql
'    '--** Paso 3  La mayor por cliente
'    Sql = "update colocCalifProv set cCalifInterna = y.Mayor "
'    Sql = Sql & "from colocCalifProv x "
'    Sql = Sql & "Join "
'    Sql = Sql & "(select cPersCod, count(*) Nro, max(cCalifInterna) Mayor  "
'    Sql = Sql & "from ColocCalifProv group by cPersCod) y "
'    Sql = Sql & "on x.cPersCod=y.cPersCod  "
'    loConec.Ejecutar Sql
    
    Set loConec = Nothing
    Exit Sub
 
Error:
    Err.Raise Err.Number, "Califica Riesgo Cambiario", Err.Description

End Sub


Public Sub nCalificaDiasAtrasoSBS(ByVal psServerConsol As String, ByVal psFecha As String)

Dim Sql As String
Dim ldFecha As Date
Dim loConec As COMConecta.DCOMConecta
Dim lr As New ADODB.Recordset
ldFecha = psFecha
   
On Error GoTo Error
 
Set loConec = New COMConecta.DCOMConecta
    loConec.AbreConexion
   
    loConec.CommadTimeOut = 12000
    Sql = "Update ColocCalifProv set cIndicadorAtraso = null, cIndicadorAtrasoPers = null "
    loConec.Ejecutar Sql
   
    '--- ** 1 Crea la tabla temporal
   Sql = "Create table #tmpIndicador (orden int identity,  cCtaCod char(18), cNroCuo int, " _
      & "                            DAtraso int, SaldoK money, PorcAcum money, Dias Int) "
   loConec.Ejecutar Sql
   
    '--- *** 2 Insertar Datos
    Sql = " Insert into #tmpIndicador (cCtaCod, cNroCuo, DAtraso, SaldoK) " _
      & " Select p.cCtaCod, p.cnrocuo, isnull(datediff(d, p.dfecvenc, dfecpag),0) DAtraso, " _
      & " SaldoK =c.nMontoApr - " _
      & "        (Select IsNull(sum(nCapital),0) from " & psServerConsol & "plandespagConsol pd " _
      & "         Where pd.cCtaCod  = p.cCtaCod and ntipo =1 and pd.cnrocuo < p.cNrocuo) " _
      & " From " & psServerConsol & "plandespagConsol p join " & psServerConsol & "CreditoConsol c on p.cCtaCod = c.cCtaCod " _
      & " where p.cCtaCod in (Select ca.cCtaCod from ColocCalifProv ca where ca.nPrdEstado in (2020,2021,2022,2030,2031,2032) and substring(ca.cCtaCod,6,1) in ('1','2'))  " _
      & " And p.ntipo =1 And p.dfecVenc  >= dateadd(year,-1,'" & Format(ldFecha, "YYYYMMDD") & "')  and p.dfecVenc  <= '" & Format(ldFecha, "YYYYMMDD") & "' " _
      & " Order by  p.cCtaCod, DAtraso Desc "
    loConec.Ejecutar Sql
    
    '--- *** 3 actualiza los porcentajes acumulados
    Sql = "Update x set x.PorcAcum = y.nPorc " _
        & " from #tmpIndicador x inner join " _
        & "(select Orden,(Select sum(SaldoK) from #tmpIndicador td where td.cCtaCod = t.cCtaCod and td.orden <= t.Orden) / " _
        & "(Select Isnull(sum(SaldoK),0) from #tmpIndicador ts where ts.cCtaCod = t.cCtaCod ) * 100 as nPorc " _
        & "From #tmpIndicador t) y on x.Orden = y.Orden     "
    loConec.Ejecutar Sql
 
    '--- *** 4 Actualiza Dias Atraso del Menor que cumpla con 10% acumulado
    Sql = "Update x set x.Dias = y.DiasIndica " _
        & "From #tmpIndicador x inner join " _
        & " (select t.cCtaCod, " _
        & "      (select top 1 datraso  from #tmpIndicador ti where ti.cCtaCod = t.cCtaCod and ti.PorcAcum >=10 order by PorcAcum ) DiasIndica " _
        & " From #tmpIndicador t " _
        & " Group by t.cCtaCod) y on x.cCtaCod = y.cCtaCod "
    loConec.Ejecutar Sql
    
    '--- *** 5 Actualiza Indicador
    Sql = " select t.cCtaCod, " _
        & " Case When Dias <= 0 then 'A'  " _
        & "      When Dias > 0  and Dias <= 8  then 'B' " _
        & "      When Dias > 8  and Dias <= 15 then 'C' " _
        & "      When Dias > 15 and Dias <= 30 then 'D' " _
        & "      When Dias > 30 and Dias <= 60 then 'E' " _
        & "      When Dias > 60 and Dias <= 90 then 'F' " _
        & "      When Dias > 90 and Dias <= 120 then 'G' " _
        & "      When Dias > 120  then 'H'  End Indicador " _
        & "From #tmpIndicador t  group by t.cCtaCod,Dias"
    
    Set lr = loConec.CargaRecordSet(Sql)
    Do While Not lr.EOF
        Sql = "Update ColocCalifProv set cIndicadorAtraso = '" & lr!Indicador & "' Where cCtaCod = '" & lr!cCtaCod & "' "
        loConec.Ejecutar Sql
        lr.MoveNext
    Loop
    Set lr = Nothing
    '--- *** 6 Elimina Temporal
    Sql = "Drop table #tmpIndicador "
    loConec.Ejecutar Sql
   
    '--- *** 7 Parches
    '** Creditos en recuperaciones
    Sql = " Update ColocCalifProv set cIndicadorAtraso = 'H' where nPrdEstado in (2201,2205,2202,2206) "
    loConec.Ejecutar Sql
    
    '** Cartas Fianza
    Sql = " Update ColocCalifProv set cIndicadorAtraso = 'A' where Substring(cCtaCod,7,2) ='21' "
    loConec.Ejecutar Sql
    '**  Nulos
    Sql = "Update ColocCalifProv set cIndicadorAtraso = 'A' where cCalGen ='0' and cIndicadorAtraso is null "
    loConec.Ejecutar Sql
    Sql = "Update ColocCalifProv set cIndicadorAtraso = 'B' where cCalGen ='1' and cIndicadorAtraso is null "
    loConec.Ejecutar Sql
    Sql = "Update ColocCalifProv set cIndicadorAtraso = 'D' where cCalGen ='2' and cIndicadorAtraso is null "
    loConec.Ejecutar Sql
    Sql = "Update ColocCalifProv set cIndicadorAtraso = 'E' where cCalGen ='3' and cIndicadorAtraso is null "
    loConec.Ejecutar Sql
    Sql = "Update ColocCalifProv set cIndicadorAtraso = 'G' where cCalGen ='4' and cIndicadorAtraso is null "
    loConec.Ejecutar Sql
    
    '--- *** 8 Iguala la mejor posible por cliente
    
    Sql = "UPDATE ColocCalifProv set cIndicadorAtrasoPers = y.cIndicadorMayor " _
        & "From ColocCalifProv x join " _
        & " (Select cPersCod, Isnull(MAX(CA1.cIndicadorAtraso),'X') as cIndicadorMayor " _
        & "  From ColocCalifProv CA1 " _
        & "  Group by cPersCod) y " _
        & "  On x.cPersCod = y.cPersCod "
    loConec.Ejecutar Sql
    
 ' --- Solo los que tienen COMERCIALEs
    Sql = "Update ColocCalifProv Set cIndicadorAtrasoPers = 'X' " _
        & "Where cPersCod not in(Select cPersCod from ColocCalifProv where substring(cCtaCod,3,1) = '1' ) "
    'loConec.Ejecutar sql

    
    Set loConec = Nothing
    Exit Sub
    
Error:
    Err.Raise Err.Number, "Califica Riesgo Cambiario", Err.Description

End Sub

'**DAOR 20071124
'**Devuelve los registros de la calificacion de cartera
Public Function darCalificacionCartera(ByVal pdFechaFinMes As Date, ByVal pnTipCamb As Double, ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ReporteCarteraCreditos '" & Format(pdFechaFinMes, "yyyymmdd") & "'," & pnTipCamb
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
        loReg.Ejecutar ("stp_sel_ObtenerActEconomica") 'JUEZ 20121015
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos para Calificacion"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos para Calificacion"
        Exit Function
    End If
    Set darCalificacionCartera = lrDatos
    Set lrDatos = Nothing
End Function


'**DAOR 20071227, Método que genera la calificación sin alineamiento
Public Sub generarCalificacionSinAlineamiento(ByRef psMensaje As String)
Dim lsSQL As String
Dim lrs As ADODB.Recordset
Dim rsVerif As ADODB.Recordset
Dim loConec As COMConecta.DCOMConecta

On Error GoTo ErrorCalificacionSinAlineamiento

    Set loConec = New COMConecta.DCOMConecta
    loConec.AbreConexion
    
    lsSQL = " exec stp_upd_GeneraCalificacionSinAlineamiento"
          
    Set loConec = New COMConecta.DCOMConecta
        loConec.AbreConexion
    Set lrs = loConec.CargaRecordSet(lsSQL)
    If lrs!nRegistros = 0 Then
        psMensaje = "No existen Datos Para Empezar a Generar la Calificación Sin Alineamiento "
    End If
    Set lrs = Nothing
    Set loConec = Nothing
    Exit Sub
    
ErrorCalificacionSinAlineamiento:
    Err.Raise Err.Number, "Calificacion Sin Alineamiento", Err.Description
End Sub

'**DAOR 20071227, Método que realiza el calculo de la provisión
'**tomando como criterio la calificación sin alineamiento
Public Sub calculaProvisionSinAlineamiento(ByVal psFecha As String)
Dim lsSQL As String
Dim lrs As New ADODB.Recordset
Dim lrs2 As New ADODB.Recordset
Dim loConec As COMConecta.DCOMConecta
Dim loCalif As COMNCredito.NCOMColocEval
Dim lnProvision As Currency
Dim lmProGarAutoL(5)
Dim lmProGarMuyRR(5) As Double
Dim lmProGarPref(5) As Double
Dim lmProGarSinG(5) As Double
Dim lnSaldoCapitalSinDiferido As Currency 'ALPA20131226
Dim lnSaldoInteresSinDiferido As Currency 'ALPA20131226
Dim lnPorceCapitalSinDiferido As Currency 'ALPA20131226
Dim lnPorceInteresSinDiferido As Currency 'ALPA20131226
Dim lnCapitalSinDiferido As Currency 'ALPA20131226

On Error GoTo ErrorCalculaProvisionSinAlineamiento
    lmProGarAutoL(0) = 1
    lmProGarMuyRR(0) = 1: lmProGarMuyRR(1) = 1.25: lmProGarMuyRR(2) = 6.25: lmProGarMuyRR(3) = 15: lmProGarMuyRR(4) = 30
    lmProGarPref(0) = 1: lmProGarPref(1) = 2.5: lmProGarPref(2) = 12.5: lmProGarPref(3) = 30: lmProGarPref(4) = 60
    lmProGarSinG(0) = 1: lmProGarSinG(1) = 5:   lmProGarSinG(2) = 25:   lmProGarSinG(3) = 60: lmProGarSinG(4) = 100
    
    Set loConec = New COMConecta.DCOMConecta
        loConec.AbreConexion
    lsSQL = " exec stp_sel_ColocCalifProvParaCalculoProvision "
    Set lrs = loConec.CargaRecordSet(lsSQL)
    
    Do While Not lrs.EOF
        lnProvision = 0
        'ALPA201312********************
        lnSaldoCapitalSinDiferido = 0
        lnPorceInteresSinDiferido = 0
        lnCapitalSinDiferido = lrs!nSaldoCap
'        If lrs!cRefinan = "R" Then
            lsSQL = "exec stp_sel_InteresDiferido '" & Trim(lrs!cCtaCod) & "','" & Format(psFecha, "YYYY/MM/DD") & "'"
            Set lrs2 = loConec.CargaRecordSet(lsSQL)
            If Not (lrs2.BOF Or lrs2.EOF) Then
                Do While Not lrs2.EOF
                    lnSaldoCapitalSinDiferido = lrs2!nSaldo
                    lnSaldoInteresSinDiferido = lrs2!nInteres
                    lrs2.MoveNext
                Loop
                lnCapitalSinDiferido = lnSaldoCapitalSinDiferido
            End If
'            loConec.Ejecutar lsSQL
'        End If
        '******************************
        
        If lrs!nGarant = 1 Then
            'lnProvision = lrs!nSaldoCap * lmProGarAutoL(0)
            lnProvision = lnCapitalSinDiferido * lmProGarAutoL(0)
        ElseIf lrs!nGarant = 2 Then
            'lnProvision = lrs!nSaldoCap * lmProGarMuyRR(IIf(Trim(lrs!cCalSinAli) = "", 4, lrs!cCalSinAli)) ' Porcentaje de Provision
            lnProvision = lnCapitalSinDiferido * lmProGarMuyRR(IIf(Trim(lrs!cCalSinAli) = "", 4, lrs!cCalSinAli)) ' Porcentaje de Provision
        ElseIf lrs!nGarant = 3 Then
            'lnProvision = lrs!nSaldoCap * lmProGarPref(IIf(Trim(lrs!cCalSinAli) = "", 4, lrs!cCalSinAli)) ' Porcentaje de Provision
            lnProvision = lnCapitalSinDiferido * lmProGarPref(IIf(Trim(lrs!cCalSinAli) = "", 4, lrs!cCalSinAli)) ' Porcentaje de Provision
        ElseIf lrs!nGarant = 4 Then
            'lnProvision = lrs!nSaldoCap * lmProGarSinG(IIf(Trim(lrs!cCalSinAli) = "", 4, lrs!cCalSinAli)) ' Porcentaje de Provision
            lnProvision = lnCapitalSinDiferido * lmProGarSinG(IIf(Trim(lrs!cCalSinAli) = "", 4, lrs!cCalSinAli)) ' Porcentaje de Provision
        End If
    
        If lrs!cTpoCredCod = "853" Or lrs!cTpoCredCod = "854" Then 'Si es MiVivienda
            lnProvision = (2 / 3) * lnProvision
        End If
        'Redondeo
        lnProvision = Format(lnProvision / 100, "#0.00")
        
        lsSQL = " exec stp_upd_ColocCalifProvProvisionSinAli '" & Trim(lrs!cPersCod) & "','" & Trim(lrs!cCtaCod) & "'," & lnProvision
        loConec.Ejecutar lsSQL
        
        lrs.MoveNext
    Loop
    lrs.Close
    Set lrs = Nothing
    Exit Sub

ErrorCalculaProvisionSinAlineamiento:
    Err.Raise Err.Number, "Calcula Provision Sin Alineamiento", Err.Description
End Sub


'**DAOR 20071227, Método que realiza el calculo de la provisión
'**tomando como criterio la calificación del sistema financiero
Public Sub calculaProvisionSistemaFinanciero(ByVal psFecha As String)
Dim lsSQL As String
Dim lrs As New ADODB.Recordset
Dim loConec As COMConecta.DCOMConecta
Dim loCalif As COMNCredito.NCOMColocEval
Dim lnProvision As Currency
Dim lmProGarAutoL(5)
Dim lmProGarMuyRR(5) As Double
Dim lmProGarPref(5) As Double
Dim lmProGarSinG(5) As Double
Dim lnSaldoCapitalSinDiferido As Currency 'ALPA20131226
Dim lnSaldoInteresSinDiferido As Currency 'ALPA20131226
Dim lnPorceCapitalSinDiferido As Currency 'ALPA20131226
Dim lnPorceInteresSinDiferido As Currency 'ALPA20131226
Dim lnCapitalSinDiferido As Currency 'ALPA20131226
Dim lrs2 As New ADODB.Recordset 'ALPA20131226

On Error GoTo ErrorCalculaProvisionSistemaFinanciero
    lmProGarAutoL(0) = 1
    lmProGarMuyRR(0) = 1: lmProGarMuyRR(1) = 1.25: lmProGarMuyRR(2) = 6.25: lmProGarMuyRR(3) = 15: lmProGarMuyRR(4) = 30
    lmProGarPref(0) = 1: lmProGarPref(1) = 2.5: lmProGarPref(2) = 12.5: lmProGarPref(3) = 30: lmProGarPref(4) = 60
    lmProGarSinG(0) = 1: lmProGarSinG(1) = 5:   lmProGarSinG(2) = 25:   lmProGarSinG(3) = 60: lmProGarSinG(4) = 100
    
    Set loConec = New COMConecta.DCOMConecta
        loConec.AbreConexion
    lsSQL = " exec stp_sel_ColocCalifProvParaCalculoProvision "
    Set lrs = loConec.CargaRecordSet(lsSQL)
    
    Do While Not lrs.EOF
        'ALPA201312********************
        lnSaldoCapitalSinDiferido = 0
        lnPorceInteresSinDiferido = 0
        lnCapitalSinDiferido = lrs!nSaldoCap
'        If lrs!cRefinan = "R" Then
            lsSQL = "exec stp_sel_InteresDiferido '" & Trim(lrs!cCtaCod) & "','" & Format(psFecha, "YYYY/MM/DD") & "'"
            Set lrs2 = loConec.CargaRecordSet(lsSQL)
            If Not (lrs2.BOF Or lrs2.EOF) Then
                Do While Not lrs2.EOF
                    lnSaldoCapitalSinDiferido = lrs2!nSaldo
                    lnSaldoInteresSinDiferido = lrs2!nInteres
                    lrs2.MoveNext
                Loop
                lnCapitalSinDiferido = lnSaldoCapitalSinDiferido
            End If
'            loConec.Ejecutar lsSQL
'        End If
        '******************************
        If Trim(lrs!cCalSistF) <> "" Then
            lnProvision = 0
            If lrs!nGarant = 1 Then
                'lnProvision = lrs!nSaldoCap * lmProGarAutoL(0)
                lnProvision = lnCapitalSinDiferido * lmProGarAutoL(0)
            ElseIf lrs!nGarant = 2 Then
                'lnProvision = lrs!nSaldoCap * lmProGarMuyRR(IIf(Trim(lrs!cCalSistF) = "", 4, lrs!cCalSistF)) ' Porcentaje de Provision
                lnProvision = lnCapitalSinDiferido * lmProGarMuyRR(IIf(Trim(lrs!cCalSistF) = "", 4, lrs!cCalSistF)) ' Porcentaje de Provision
            ElseIf lrs!nGarant = 3 Then
                'lnProvision = lrs!nSaldoCap * lmProGarPref(IIf(Trim(lrs!cCalSistF) = "", 4, lrs!cCalSistF)) ' Porcentaje de Provision
                lnProvision = lnCapitalSinDiferido * lmProGarPref(IIf(Trim(lrs!cCalSistF) = "", 4, lrs!cCalSistF)) ' Porcentaje de Provision
            ElseIf lrs!nGarant = 4 Then
                'lnProvision = lrs!nSaldoCap * lmProGarSinG(IIf(Trim(lrs!cCalSistF) = "", 4, lrs!cCalSistF)) ' Porcentaje de Provision
                lnProvision = lnCapitalSinDiferido * lmProGarSinG(IIf(Trim(lrs!cCalSistF) = "", 4, lrs!cCalSistF)) ' Porcentaje de Provision
            End If
        
            If lrs!cTpoCredCod = "853" Or lrs!cTpoCredCod = "854" Then 'Si es MiVivienda
                lnProvision = (2 / 3) * lnProvision
            End If
            'Redondeo
            lnProvision = Format(lnProvision / 100, "#0.00")
            
            lsSQL = " exec stp_upd_ColocCalifProvProvisionSistF '" & Trim(lrs!cPersCod) & "','" & Trim(lrs!cCtaCod) & "'," & lnProvision
            loConec.Ejecutar lsSQL
        End If
        lrs.MoveNext
    Loop
    lrs.Close
    Set lrs = Nothing
    Exit Sub

ErrorCalculaProvisionSistemaFinanciero:
    Err.Raise Err.Number, "Calcula Provision Sistema Financiero", Err.Description
End Sub
'ALPA 20081209****************************************************************
Public Sub nCalificaCalculaProvisionProciclica(ByVal psFecha As String)

Dim lsSQL As String
Dim lrs As New ADODB.Recordset
Dim loConec As COMConecta.DCOMConecta
Dim loCalif As COMNCredito.NCOMColocEval
Dim lnProvisionProciclica As Currency
Dim lnSaldoCapitalSinDiferido As Currency 'ALPA20131226
Dim lnSaldoInteresSinDiferido As Currency 'ALPA20131226
Dim lnPorceCapitalSinDiferido As Currency 'ALPA20131226
Dim lnPorceInteresSinDiferido As Currency 'ALPA20131226
Dim lnCapitalSinDiferido As Currency 'ALPA20131226
Dim lrs2 As New ADODB.Recordset 'ALPA20131226
Dim lmProGarAutoL(5)
Dim lmProGarMuyRR(5) As Double
Dim lmProGarPref(5) As Double
Dim lmProGarSinG(5) As Double

On Error GoTo ErrorCalculaProvision

Set loConec = New COMConecta.DCOMConecta
loConec.AbreConexion

lsSQL = "exec stp_sel_ListaCalculaProvision"
Set lrs = loConec.CargaRecordSet(lsSQL)

Do While Not lrs.EOF
    lnProvisionProciclica = 0
    'ALPA201312********************
    lnSaldoCapitalSinDiferido = 0
    lnSaldoInteresSinDiferido = 0
    lnPorceCapitalSinDiferido = 0
    lnPorceInteresSinDiferido = 0
    lnCapitalSinDiferido = lrs!nSaldoCap
'    If lrs!cRefinan = "R" Then
        lsSQL = "exec stp_sel_InteresDiferido '" & Trim(lrs!cCtaCod) & "','" & Format(psFecha, "YYYY/MM/DD") & "'"
        Set lrs2 = loConec.CargaRecordSet(lsSQL)
        If Not (lrs2.BOF Or lrs2.EOF) Then
            Do While Not lrs2.EOF
                lnSaldoCapitalSinDiferido = lrs2!nSaldo
                lnSaldoInteresSinDiferido = lrs2!nInteres
                lrs2.MoveNext
            Loop
            lnCapitalSinDiferido = lnSaldoCapitalSinDiferido
        End If
'        loConec.Ejecutar lsSQL
'    End If
    '******************************

    If lrs!cCalGen = 0 Then
        If lrs!nGarant = 1 Then
            'lnProvisionProciclica = lrs!nSaldoCap * lrs!nProvisionProciclicaA
            lnProvisionProciclica = lnCapitalSinDiferido * lrs!nProvisionProciclicaA
        Else
            'lnProvisionProciclica = lrs!nSaldoCap * lrs!nProvisionProciclica
            lnProvisionProciclica = lnCapitalSinDiferido * lrs!nProvisionProciclica
        End If
    End If
    lnProvisionProciclica = Format(lnProvisionProciclica / 100, "#0.00")
    
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000101", lrs!cTpoProdCod) Then
    'If lrs!cTpoProdCod = "514" Then 'Si es Carta Fianza
    '**ARLO20180712 ERS042 - 2018
        lnProvisionProciclica = Round((lnProvisionProciclica / 2), 2)
    End If
    If lrs!cTpoCredCod = "853" Or lrs!cTpoCredCod = "854" Then 'Si es MiVivienda
        lnProvisionProciclica = (2 / 3) * lnProvisionProciclica
    End If
    lsSQL = "UPDATE ColocCalifProv SET nProvisionProciclica= " & lnProvisionProciclica _
          & " WHERE cPersCod ='" & Trim(lrs!cPersCod) & "' And cCtaCod ='" & Trim(lrs!cCtaCod) & "' "
    loConec.Ejecutar lsSQL

    lrs.MoveNext
Loop
lrs.Close
Set lrs = Nothing
Exit Sub

ErrorCalculaProvision:
    Err.Raise Err.Number, "Calcula Provision", Err.Description
End Sub
'*****************************************************************************
'**ALPA 20080605
Public Function darAnexo5DCOFIDE(ByVal pdFechaFinMes As Date, ByVal pnTipCamb As Double, ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_Reporte5DCofide '" & Format(pdFechaFinMes, "yyyymmdd") & "'," & pnTipCamb
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    Set darAnexo5DCOFIDE = lrDatos
    Set lrDatos = Nothing
End Function

'*** PEAC 20100514
Public Function RecuperaDatosCartasWORD_NEW(ByVal pcTipoCadena As Byte, ByVal pcCondicion, ByVal pcMoneda As String, ByVal pcProductos As String, ByVal pcAnalistas As String, _
                                      ByVal pnDiaIni As Integer, ByVal pnDiaFin As Integer, _
                                    Optional ByVal psUbicacionGeo As String = "", _
                                    Optional ByVal psCodUser As String = "", _
                                    Optional ByVal psMaqUser As String, Optional ByVal psAgencias As String, Optional ByVal psCodAge As String) As ADODB.Recordset
         
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperaDatosCartasWORD_NEW

pcCondicion = Replace(pcCondicion, " ", "")
pcMoneda = Replace(Replace(pcMoneda, "'", ""), " ", "")
pcProductos = Replace(Replace(pcProductos, "'", ""), " ", "")
pcAnalistas = Replace(Replace(pcAnalistas, "'", ""), " ", "")
psAgencias = Replace(Replace(psAgencias, "'", ""), " ", "")

If Len(psUbicacionGeo) > 0 Then
    sSql = " exec stp_sel_RecuperaDatosCartasWORD_NEW '" & psCodAge & "'," & pcTipoCadena & ",'" & pcAnalistas & "','" & psAgencias & "'," & pnDiaIni & "," & pnDiaFin & ",'" & pcCondicion & "','" & pcMoneda & "','" & pcProductos & "','" & psUbicacionGeo & "'"
Else
    sSql = " exec stp_sel_RecuperaDatosCartasWORD_NEW '" & psCodAge & "'," & pcTipoCadena & ",'" & pcAnalistas & "','" & psAgencias & "'," & pnDiaIni & "," & pnDiaFin & ",'" & pcCondicion & "','" & pcMoneda & "','" & pcProductos & "'"
End If

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaDatosCartasWORD_NEW = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaDatosCartasWORD_NEW:
    Err.Raise Err.Number, "Error En Proceso COMNCredito.NCOMColEval.RecuperaDatosCartasWORD_NEW", Err.Description
End Function

'ALPA 20100914*********************
Public Function darReclasificacionCartera(ByVal psFechaFinMes As String, ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec B2_stp_sel_ReporteReclasificados '" & psFechaFinMes & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos para ReClasificacion"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos para ReClasificacion"
        Exit Function
    End If
    Set darReclasificacionCartera = lrDatos
    Set lrDatos = Nothing
End Function

'ALPA 20101124
Public Function nExisteTabla(ByVal lsNombreTabla As String, ByVal psServConsol As String) As Boolean
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String

nExisteTabla = False
Set oCom = New COMConecta.DCOMConecta
        oCom.AbreConexion
lsSQL = "Select * From " & psServConsol & "SysObjects Where Name = '" & lsNombreTabla & "' and type ='U'"
Set rs = oCom.CargaRecordSet(lsSQL)
oCom.CierraConexion
    Set oCom = Nothing
If Not (rs.BOF And rs.EOF) Then
    nExisteTabla = True
End If
rs.Close
Set rs = Nothing
End Function

Public Function ObtenerCarteraComparativa(ByVal pdFecha As Date, ByVal pnTipoCambio As Currency) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String

'ObtenerCarteraComparativa = False
Set oCom = New COMConecta.DCOMConecta
oCom.AbreConexion

lsSQL = "exec stp_ins_ComparacionCartera '" & Format(pdFecha, "YYYYMMDD") & "','MAC'"
oCom.CargaRecordSet (lsSQL)
lsSQL = "exec stp_ins_ComparacionCartera '" & Format(pdFecha, "YYYYMMDD") & "','MAN'"
oCom.CargaRecordSet (lsSQL)
lsSQL = "exec stp_sel_ComparacionCartera '" & Format(pdFecha, "YYYYMMDD") & "'," & pnTipoCambio
Set rs = oCom.CargaRecordSet(lsSQL)
Set ObtenerCarteraComparativa = rs
oCom.CierraConexion
Set oCom = Nothing
'rs.Close
'Set rs = Nothing
End Function
'ALPA 20120214***********************************************************************
Public Function ObtenerIntDevTpoCred(ByVal pdFecha As Date) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    lsSQL = "exec stp_sel_IntDevTpoCred '" & Format(pdFecha, "YYYYMMDD") & "'"
    Set rs = oCom.CargaRecordSet(lsSQL)
    Set ObtenerIntDevTpoCred = rs
    oCom.CierraConexion
Set oCom = Nothing
End Function

Public Sub InsertarIntDevTpoCred(ByVal psTpoCredCod As String, ByVal psOrden As String, ByVal pdFecha As Date, ByVal pnIntDif As Currency)
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    lsSQL = "exec stp_ins_IntDevTpoCred '" & psTpoCredCod & "','" & psOrden & "','" & Format(pdFecha, "YYYYMMDD") & "'," & pnIntDif
    oCom.CargaRecordSet (lsSQL)
    oCom.CierraConexion
Set oCom = Nothing
End Sub

Public Sub EliminarIntDevTpoCred(ByVal psTpoCredCod As String, ByVal psOrden As String, ByVal pdFecha As Date)
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    lsSQL = "exec stp_del_IntDevTpoCred '" & psTpoCredCod & "','" & psOrden & "','" & Format(pdFecha, "YYYYMMDD") & "'"
    oCom.CargaRecordSet (lsSQL)
    oCom.CierraConexion
Set oCom = Nothing
End Sub
'************************************************************************************

Public Function InsertaExposiciones2A1(ByVal pdFecha As Date, _
                                       ByVal pnOrdenTCredito As Integer, _
                                       ByVal pnValor As Currency)
Dim lsSQL As String

On Error GoTo ErrorInsertaExposiciones2A1

Dim oCom As COMConecta.DCOMConecta
Set oCom = New COMConecta.DCOMConecta
 oCom.AbreConexion
    lsSQL = "exec stp_ins_Exposiciones2A1Verifica '" & Format(pdFecha, "YYYYMMDD") & "','" & pnOrdenTCredito & "','" & pnValor & "'"
    oCom.CargaRecordSet (lsSQL)
    lsSQL = "exec stp_ins_Exposiciones2A1Verifica '" & Format(pdFecha, "YYYYMMDD") & "',17,'" & pnValor & "'"
    oCom.CargaRecordSet (lsSQL)
    oCom.CierraConexion
Set oCom = Nothing
Exit Function

ErrorInsertaExposiciones2A1:
    Err.Raise Err.Number, "Error InsertaExposiciones2A1", Err.Description
End Function
'ALPA 20120628
Public Function CargarExposicionesAjustadas2A1(ByVal dFecha As Date) As ADODB.Recordset
On Error GoTo ErrorCargarExposicionesAjustadas2A1
Dim oCom As COMConecta.DCOMConecta
Set oCom = New COMConecta.DCOMConecta
        Dim sSql As String
        oCom.AbreConexion
        sSql = "exec stp_sel_ObtenerExposicionesAjustadas2A1 '" & Format(dFecha, "YYYY/MM/DD") & "'"
        Set CargarExposicionesAjustadas2A1 = oCom.CargaRecordSet(sSql)
        oCom.CierraConexion
        Exit Function
ErrorCargarExposicionesAjustadas2A1:
        Set CargarExposicionesAjustadas2A1 = Nothing
        Err.Raise Err.Number, "Negocio Recibido", Err.Description
End Function
Public Function CargaBalanceGenxCodigo(ByVal psFecha As String) As ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Set oCom = New COMConecta.DCOMConecta
Dim sSql As String
    oCom.AbreConexion
    sSql = "exec stp_sel_BalanceGenxCodigo '" & Format(psFecha, "YYYYMMDD") & "'"
    Set CargaBalanceGenxCodigo = oCom.CargaRecordSet(sSql)
    oCom.CierraConexion
End Function
Public Function CargaDatosPatrimonio(ByVal lsAnio As String, ByVal lsMes As String, ByVal lsTipoBalance As String) As ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Set oCom = New COMConecta.DCOMConecta
Dim sSql As String
    oCom.AbreConexion
    sSql = "exec stp_sel_Reporte3Patrimonio '" & lsAnio & "', '" & lsMes & "', '" & lsTipoBalance & "'"
    Set CargaDatosPatrimonio = oCom.CargaRecordSet(sSql)
    oCom.CierraConexion
End Function

Public Function ReporteRiesgoCambiario(ByVal sBalAnio As String, ByVal sBalMes As String) As ADODB.Recordset
On Error GoTo ReporteRiesgoCambiario
    Dim oCom As COMConecta.DCOMConecta
    Set oCom = New COMConecta.DCOMConecta
        Dim sSql As String
        oCom.AbreConexion
        sSql = "exec stp_sel_ReporteRiesgoCambiario '" & sBalAnio & "', '" & sBalMes & "'"
        Set ReporteRiesgoCambiario = oCom.CargaRecordSet(sSql)
        oCom.CierraConexion
        Exit Function
ReporteRiesgoCambiario:
        Set ReporteRiesgoCambiario = Nothing
        Err.Raise Err.Number, "Financiamiento Recibido", Err.Description
End Function

Public Function ObtenerExposiciones2A1(ByVal pdFecha As Date) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    lsSQL = "exec stp_sel_Exposiciones2A1 '" & Format(pdFecha, "YYYYMMDD") & "'"
    Set rs = oCom.CargaRecordSet(lsSQL)
    Set ObtenerExposiciones2A1 = rs
    oCom.CierraConexion
Set oCom = Nothing
End Function
'ALPA 20120628
Public Function FactorAjusteRiesgoOperac(ByVal sBalAnio As String, ByVal sBalMes As String, ByVal nTipo As Integer) As ADODB.Recordset
On Error GoTo FactorAjusteRiesgoOperac
        Dim oCom As COMConecta.DCOMConecta
        Dim lsSQL As String
        Set oCom = New COMConecta.DCOMConecta
        oCom.AbreConexion
        lsSQL = "exec stp_sel_FactorAjusteRiesgoOperac '" & sBalAnio & "', '" & sBalMes & "', '" & nTipo & "'"
        Set FactorAjusteRiesgoOperac = oCom.CargaRecordSet(lsSQL)
        oCom.CierraConexion
        Exit Function
FactorAjusteRiesgoOperac:
        Set FactorAjusteRiesgoOperac = Nothing
        Err.Raise Err.Number, "Financiamiento Recibido", Err.Description
End Function
'*************************************************

'MIOL 20120801, SEGUN RQ12077 ******************************************************
Public Function obtenerCreditosSugeridosHipotecarios(Optional ByVal pnTipoRegCtrl As TipoRegControl = gTpoRegCtrlInformeLegal, Optional psCtaCod As String = "", _
                                                 Optional ByVal psTitular As String = "", Optional ByVal psAgeCod As String = "") As ADODB.Recordset 'OK
    Dim oDCOMColocEval As COMDCredito.DCOMColocEval
    Set oDCOMColocEval = New COMDCredito.DCOMColocEval
    
    Set obtenerCreditosSugeridosHipotecarios = oDCOMColocEval.getCreditosSugeridosHipotecarios(pnTipoRegCtrl, psCtaCod, psTitular, psAgeCod)
    Set oDCOMColocEval = Nothing
End Function

Public Function obtenerMontoVinculados(ByVal pscCtaCod As String, ByVal dFecSis As String) As ADODB.Recordset 'OK
    Dim oDCOMColocEval As COMDCredito.DCOMColocEval
    Set oDCOMColocEval = New COMDCredito.DCOMColocEval
    
    Set obtenerMontoVinculados = oDCOMColocEval.getMontoVinculados(pscCtaCod, dFecSis)
    Set oDCOMColocEval = Nothing
End Function

Public Function obtenerCreditoAprobadosTotalesSC() As ADODB.Recordset 'OK
    Dim oDCOMColocEval As COMDCredito.DCOMColocEval
    Set oDCOMColocEval = New COMDCredito.DCOMColocEval
    
    Set obtenerCreditoAprobadosTotalesSC = oDCOMColocEval.getCreditoAprobadosTotalesSC()
    Set oDCOMColocEval = Nothing
End Function

Public Function obtenerParametroRevision() As ADODB.Recordset 'OK
    Dim oDCOMColocEval As COMDCredito.DCOMColocEval
    Set oDCOMColocEval = New COMDCredito.DCOMColocEval
    
    Set obtenerParametroRevision = oDCOMColocEval.getParametrosRevision()
    Set oDCOMColocEval = Nothing
End Function

Public Sub insEstadosExpediente(ByVal pscCtaCod As String, ByVal pscEstUbicacion As String, _
                                                 ByVal pscIngExpediente As String, ByVal pscSalObsExpediente As String, _
                                                 ByVal pscReIngObsExpediente As String, ByVal pscSalExpediente As String, _
                                                 ByVal psnEstRevExpediente As Integer, ByVal psnPrdEstado As Integer, _
                                                 Optional ByVal pnTipoRegControl As TipoRegControl = gTpoRegCtrlRiesgos)  'OK
    Dim oDCOMColocEval As COMDCredito.DCOMColocEval
    Set oDCOMColocEval = New COMDCredito.DCOMColocEval
    
    Call oDCOMColocEval.insEstadosExpediente(pscCtaCod, pscEstUbicacion, pscIngExpediente, pscSalObsExpediente, _
                                                                                                pscReIngObsExpediente, pscSalExpediente, psnEstRevExpediente, psnPrdEstado, pnTipoRegControl)
    Set oDCOMColocEval = Nothing

End Sub

Public Sub updateEstadoExpediente(ByVal pscCtaCod As String, Optional ByVal pnTipoRegControl As TipoRegControl = gTpoRegCtrlRiesgos) 'OK 'BY ARLO MODIFY 20171027
    Dim oDCOMColocEval As COMDCredito.DCOMColocEval
    Set oDCOMColocEval = New COMDCredito.DCOMColocEval
    
    Call oDCOMColocEval.updateEstadoExpediente(pscCtaCod, pnTipoRegControl) 'BY ARLO MODIFY 20171027
    Set oDCOMColocEval = Nothing

End Sub

Public Sub updateParametroRevision(ByVal psnParRevExp As Integer) 'OK
    Dim oDCOMColocEval As COMDCredito.DCOMColocEval
    Set oDCOMColocEval = New COMDCredito.DCOMColocEval
    
    Call oDCOMColocEval.updateParametroRevision(psnParRevExp)
    Set oDCOMColocEval = Nothing

End Sub
'END MIOL **************************************************************************

'MIOL 20120920, SEGUN RQ12257 ******************************************************
'OK
Public Function ValidaCuentaCliente(ByVal psPersNroDoc As String) As Boolean
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
Dim rs As ADODB.Recordset

    ValidaCuentaCliente = False
    sSql = "exec stp_val_ClientexNroDocSunat '" & psPersNroDoc & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        ValidaCuentaCliente = True
    End If
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
End Function

'OK
Public Function ValidaClienteBloqueado(ByVal psPersNroDoc As String) As Boolean
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
Dim rs As ADODB.Recordset

    ValidaClienteBloqueado = False
    sSql = "stp_val_ClienteBloqueadoxNroDoc '" & psPersNroDoc & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        ValidaClienteBloqueado = True
    End If
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
End Function

'OK
Public Function obtenerSaldoClienteSunat(ByVal pscNroDoc As String, ByVal psnTipo As Integer) As ADODB.Recordset
    Dim oDCOMColocEval As COMDCredito.DCOMColocEval
    Set oDCOMColocEval = New COMDCredito.DCOMColocEval
    
    Set obtenerSaldoClienteSunat = oDCOMColocEval.getSaldoClienteSunat(pscNroDoc, psnTipo)
    Set oDCOMColocEval = Nothing
End Function

'OK
Public Function obtenerGaranSaldoClienteSunat(ByVal pscNroDoc As String) As ADODB.Recordset
    Dim oDCOMColocEval As COMDCredito.DCOMColocEval
    Set oDCOMColocEval = New COMDCredito.DCOMColocEval
    
    Set obtenerGaranSaldoClienteSunat = oDCOMColocEval.getGaranSaldoClienteSunat(pscNroDoc)
    Set oDCOMColocEval = Nothing
End Function
'END MIOL **************************************************************************


'WIOR 20140620 ********************************************************************
Public Function RecuperaCreditoCargoDebito(ByVal pdFecha As Date, ByVal psAnalistas As String, ByVal psAge As String) As ADODB.Recordset
        
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperaCreditoCargoDebito

psAnalistas = Replace(Replace(psAnalistas, "'", ""), " ", "")

sSql = " EXEC stp_sel_CreditosPagadosAutomatico '" & Format(pdFecha, "yyyymmdd") & "','" & psAnalistas & "','" & psAge & "'"


Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaCreditoCargoDebito = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaCreditoCargoDebito:
    Err.Raise Err.Number, "Error En Proceso COMNCredito.NCOMColEval.RecuperaCreditoCargoDebito", Err.Description
End Function
'WIOR FIN *************************************************************************

'WIOR 20160623 ***
Public Function SobreEndVerificaCartera() As Boolean
Dim oDCOMColocEval As COMDCredito.DCOMColocEval
Dim rsDatos As ADODB.Recordset

Set oDCOMColocEval = New COMDCredito.DCOMColocEval
Set rsDatos = oDCOMColocEval.SobreEndVerificaCartera()

If Not (rsDatos.EOF And rsDatos.BOF) Then
    SobreEndVerificaCartera = True
Else
    SobreEndVerificaCartera = False
End If

Set oDCOMColocEval = Nothing
Set rsDatos = Nothing
End Function

Public Function SobreEndEjecutarCartera() As Boolean
Dim oDCOMColocEval As COMDCredito.DCOMColocEval
On Error GoTo ErrorSobreEndEjecutarCartera

Set oDCOMColocEval = New COMDCredito.DCOMColocEval
Call oDCOMColocEval.SobreEndEjecutarCartera
SobreEndEjecutarCartera = True

Exit Function

ErrorSobreEndEjecutarCartera:
SobreEndEjecutarCartera = False
End Function
'WIOR FIN ********
'RECO20161020 ERS060-2016***********************************************
Public Function ValidaExisteRegProceso(ByVal psCtaCod As String, ByVal pnTipoRegControl As TipoRegControl) As Boolean
    Dim oCredEval As New COMDCredito.DCOMColocEval
    Dim oRs As New ADODB.Recordset
    
    Set oRs = oCredEval.ValidaExisteRegProceso(psCtaCod, pnTipoRegControl)
    If Not (oRs.EOF And oRs.BOF) Then
        If oRs!bExiste = 1 Then
            ValidaExisteRegProceso = True
        Else
            ValidaExisteRegProceso = False
        End If
    Else
        ValidaExisteRegProceso = False
    End If
    Set oCredEval = Nothing
    Set oRs = Nothing
End Function

Public Function ListaCreditosSeguimiento(ByVal psCtaCod As String, ByVal psPersNombre As String, ByVal psAgeCod As String) As ADODB.Recordset
    Dim oEvalD As New COMDCredito.DCOMColocEval
    Set ListaCreditosSeguimiento = oEvalD.ListaCreditosSeguimiento(psCtaCod, psPersNombre, psAgeCod)
    Set oEvalD = Nothing
End Function

Public Function CredDetalleSeguimiento(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oEvalD As New COMDCredito.DCOMColocEval
    Set CredDetalleSeguimiento = oEvalD.CredDetalleSeguimiento(psCtaCod)
    Set oEvalD = Nothing
End Function
'RECO FIN***************************************************************
'ARLO 20170919, SEGUN ERS060-2016 ******************************************************
Public Function obtenerCreditosAprobadosMinuta(Optional ByVal pnTipoRegCtrl As TipoRegControl = gTpoRegCtrlInformeLegal, Optional psCtaCod As String = "", _
                                                 Optional ByVal psTitular As String = "", Optional ByVal psAgeCod As String = "") As ADODB.Recordset 'OK
    Dim oDCOMColocEval As COMDCredito.DCOMColocEval
    Set oDCOMColocEval = New COMDCredito.DCOMColocEval
    
    Set obtenerCreditosAprobadosMinuta = oDCOMColocEval.getCreditosAprobadosMinuta(pnTipoRegCtrl, psCtaCod, psTitular, psAgeCod)
    Set oDCOMColocEval = Nothing
End Function
'ARLO 20170919, SEGUN ERS060-2016 ******************************************************
Public Function obtenerCreditosAsistenteAgencia(Optional ByVal pnTipoRegCtrl As TipoRegControl = gTpoRegCtrlInformeLegal, Optional psCtaCod As String = "", _
                                                 Optional ByVal psTitular As String = "", Optional ByVal psAgeCod As String = "") As ADODB.Recordset 'OK
    Dim oDCOMColocEval As COMDCredito.DCOMColocEval
    Set oDCOMColocEval = New COMDCredito.DCOMColocEval
    
    Set obtenerCreditosAsistenteAgencia = oDCOMColocEval.getCreditosAsistenteAgencia(pnTipoRegCtrl, psCtaCod, psTitular, psAgeCod)
    Set oDCOMColocEval = Nothing
End Function

'CTI1 20180706 ***
Public Function CalcularCapPagSobreEndEjecutarCartera(ByVal pdFecha As Date) As Boolean
Dim oDCOMColocEval As COMDCredito.DCOMColocEval
On Error GoTo ErrorCalcularCapPagSobreEndEjecutarCartera

Set oDCOMColocEval = New COMDCredito.DCOMColocEval

Call oDCOMColocEval.CalculoCapPagoSobreEndEjecutarCartera(pdFecha)

CalcularCapPagSobreEndEjecutarCartera = True

Exit Function

ErrorCalcularCapPagSobreEndEjecutarCartera:
CalcularCapPagSobreEndEjecutarCartera = False
End Function
'CTI1 FIN ********

'LUCV20190601
Public Sub nCalificaPreparaArchivoCalificacionNuevo(ByVal psFecha As String, _
                                                    ByVal pnTipoCambio As Double, _
                                                    ByVal pstipocredito As Variant, _
                                                    ByRef psMensaje As String)

    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    Set oDCOMConecta = New COMConecta.DCOMConecta
    
On Error GoTo ErrorPreparaArchivoCalificacionNuevo

    oDCOMConecta.AbreConexion
    lsSQL = "Exec stp_sel_11356_CalificacionPrepararArchivo '" & psFecha & "','" & pnTipoCambio & "', '" & pstipocredito & "'"
    Set lrDatos = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion
    
    If Not (lrDatos.EOF And lrDatos.BOF) Then
        If lrDatos!bResultado = 1 Then
            psMensaje = ""
        Else
            psMensaje = "Ha ocurrido un error. Por favor, Comunicarse con el Dpto. De TI"
        End If
    Else
        psMensaje = "No existen datos para la generación de la Calificacion"
        Exit Sub
    End If
    
    Set oDCOMConecta = Nothing
    Set lrDatos = Nothing
    Exit Sub
    
ErrorPreparaArchivoCalificacionNuevo:
    Err.Raise Err.Number, "Calificacion Preparar Archivo", Err.Description
End Sub

Public Sub nCalificaActualizarGarantiasNuevo(ByVal pnTipoCambio As Double, ByRef psMensaje As String)
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lrDatosGarantias As ADODB.Recordset
    Set oDCOMConecta = New COMConecta.DCOMConecta
    
On Error GoTo ErrorActualizarGarantias

    oDCOMConecta.AbreConexion
    lsSQL = "Exec stp_sel_11356_CalificacionGarantias '" & pnTipoCambio & "'"
    Set lrDatosGarantias = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion

    
    If Not (lrDatosGarantias.EOF And lrDatosGarantias.BOF) Then
        If lrDatosGarantias!bResultado = 1 Then
            psMensaje = ""
        Else
            psMensaje = "Ha ocurrido un error. Por favor, Comunicarse con el Dpto. De TI"
        End If
    Else
        psMensaje = "No existen datos para la generación de la Calificacion"
        Exit Sub
    End If
    
    Set oDCOMConecta = Nothing
    Set lrDatosGarantias = Nothing
    
    Exit Sub
ErrorActualizarGarantias:
    Err.Raise Err.Number, "Actualizar Garantias", Err.Description
End Sub

Public Sub nCalificaCMACTNuevo(ByVal pnTipoCambio As Double, ByRef psMensaje As String)
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lrDatosCMACM As ADODB.Recordset
    Dim lsCadConexion As String
    Set oDCOMConecta = New COMConecta.DCOMConecta

On Error GoTo ErrorCalificaCMACT

    oDCOMConecta.AbreConexion
    lsSQL = "Exec stp_sel_11356_CalificacionCMACM '" & pnTipoCambio & "'"
    Set lrDatosCMACM = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion
    

    If Not (lrDatosCMACM.EOF And lrDatosCMACM.BOF) Then
        If lrDatosCMACM!bResultado = 1 Then
            psMensaje = ""
        Else
            psMensaje = "Ha ocurrido un error. Por favor, Comunicarse con el Dpto. De TI"
        End If
    Else
        psMensaje = "No existen datos para la generación de la Calificacion"
        Exit Sub
    End If

    Set oDCOMConecta = Nothing
    Set lrDatosCMACM = Nothing
    
    Exit Sub
ErrorCalificaCMACT:
    Err.Raise Err.Number, "Califica CMACT", Err.Description
End Sub


Public Sub nCalificaEvaluacionRiesgosNuevo(ByVal psFecha As String, ByRef psMensaje As String)
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lrDatosEvalRiesgo As ADODB.Recordset
    Dim lsCadConexion As String
    Set oDCOMConecta = New COMConecta.DCOMConecta
    
On Error GoTo ErrorCalculaEvaluacion

    oDCOMConecta.AbreConexion
    lsSQL = "Exec stp_sel_11356_CalificacionEvaluacionRiesgos '" & psFecha & "'"
    Set lrDatosEvalRiesgo = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion


    If Not (lrDatosEvalRiesgo.EOF And lrDatosEvalRiesgo.BOF) Then
        If lrDatosEvalRiesgo!bResultado = 1 Then
            psMensaje = ""
        Else
            psMensaje = "Ha ocurrido un error. Por favor, Comunicarse con el Dpto. De TI"
        End If
    Else
        psMensaje = "No existen datos para la generación de la Calificacion"
        Exit Sub
    End If

    Set oDCOMConecta = Nothing
    Set lrDatosEvalRiesgo = Nothing
    
    Exit Sub
ErrorCalculaEvaluacion:
    Err.Raise Err.Number, "Califica Evaluacion Riesgos ", Err.Description
End Sub

Public Sub nCalificaSistemaFinancieroNuevo(ByVal pdFecAlin As String, _
                                           ByRef psMensaje As String)
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lrDatosSistemaFinaciero As ADODB.Recordset
    Dim lsCadConexion As String
    Set oDCOMConecta = New COMConecta.DCOMConecta
On Error GoTo ErrorCalificaSistemaFinanciero

    oDCOMConecta.AbreConexion
    lsSQL = "Exec stp_sel_11356_CalificacionSistemaFinanciero '" & pdFecAlin & "'"
    Set lrDatosSistemaFinaciero = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion


    If Not (lrDatosSistemaFinaciero.EOF And lrDatosSistemaFinaciero.BOF) Then
        If lrDatosSistemaFinaciero!bResultado = 1 Then
            psMensaje = ""
        Else
            psMensaje = "Ha ocurrido un error. Por favor, Comunicarse con el Dpto. De TI"
        End If
    Else
        psMensaje = "No existen datos para la generación de la Calificacion"
        Exit Sub
    End If

    Set oDCOMConecta = Nothing
    Set lrDatosSistemaFinaciero = Nothing
    
    Exit Sub
ErrorCalificaSistemaFinanciero:
    Err.Raise Err.Number, "Califica Sistema Financiero", Err.Description
End Sub

Public Sub nCalificaGeneralNuevo(ByVal psFecha As String, ByRef psMensaje As String)
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lrDatosGeneral As ADODB.Recordset
    Dim lsCadConexion As String
    Set oDCOMConecta = New COMConecta.DCOMConecta

On Error GoTo ErrorCalificacionGeneral

    oDCOMConecta.AbreConexion
    lsSQL = "Exec stp_sel_11356_CalificacionGeneral '" & psFecha & "'"
    Set lrDatosGeneral = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion

    If Not (lrDatosGeneral.EOF And lrDatosGeneral.BOF) Then
        If lrDatosGeneral!bResultado = 1 Then
            psMensaje = ""
        Else
            psMensaje = "Ha ocurrido un error. Por favor, Comunicarse con el Dpto. De TI"
        End If
    Else
        psMensaje = "No existen datos para la generación de la Calificacion"
        Exit Sub
    End If

    Set oDCOMConecta = Nothing
    Set lrDatosGeneral = Nothing
    
    Exit Sub
ErrorCalificacionGeneral:
    Err.Raise Err.Number, "Calificacion General", Err.Description
End Sub

Public Sub nCalificaCalculaProvisionNuevo(ByVal psFecha As String, ByVal psTipoCambio As String, ByRef psMensaje As String)
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lrDatosProvision As ADODB.Recordset
    Dim lsCadConexion As String
    Set oDCOMConecta = New COMConecta.DCOMConecta
    
On Error GoTo ErrorCalculaProvision

    oDCOMConecta.AbreConexion
    lsSQL = "Exec stp_sel_11356_CalculaProvision '" & psFecha & "','" & psTipoCambio & "'"
    Set lrDatosProvision = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion

    If Not (lrDatosProvision.EOF And lrDatosProvision.BOF) Then
        If lrDatosProvision!bResultado = 1 Then
            psMensaje = ""
        Else
            psMensaje = "Ha ocurrido un error. Por favor, Comunicarse con el Dpto. De TI"
        End If
    Else
        psMensaje = "No existen datos para la generación de la Calificacion"
        Exit Sub
    End If

    Set oDCOMConecta = Nothing
    Set lrDatosProvision = Nothing
    
    Exit Sub
ErrorCalculaProvision:
    Err.Raise Err.Number, "Calcula Provision", Err.Description
End Sub

Public Sub generarCalificacionSinAlineamientoNuevo(ByRef psMensaje As String)
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lrDatosSinAlineamiento As ADODB.Recordset
    Dim lsCadConexion As String
    Set oDCOMConecta = New COMConecta.DCOMConecta

On Error GoTo ErrorCalificacionSinAlineamiento

    oDCOMConecta.AbreConexion
    lsSQL = "Exec stp_sel_11356_GeneraCalificacionSinAlineamiento"
    
    Set lrDatosSinAlineamiento = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion

    If Not (lrDatosSinAlineamiento.EOF And lrDatosSinAlineamiento.BOF) Then
        If lrDatosSinAlineamiento!bResultado = 1 Then
            psMensaje = ""
        Else
            psMensaje = "Ha ocurrido un error. Por favor, Comunicarse con el Dpto. De TI"
        End If
    Else
        If lrDatosSinAlineamiento!nRegistros = 0 Then
            psMensaje = "No existen Datos Para Empezar a Generar la Calificación Sin Alineamiento "
            Exit Sub
        Else
            psMensaje = ""
        End If
    End If

    Set oDCOMConecta = Nothing
    Set lrDatosSinAlineamiento = Nothing
    
    Exit Sub
ErrorCalificacionSinAlineamiento:
    Err.Raise Err.Number, "Calificacion Sin Alineamiento", Err.Description
End Sub

Public Sub calculaProvisionSinAlineamientoNuevo(ByVal psFecha As String, ByRef psMensaje As String)
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lrDatosProvisionSinAli As ADODB.Recordset
    Dim lsCadConexion As String
    Set oDCOMConecta = New COMConecta.DCOMConecta

On Error GoTo ErrorCalculaProvisionSinAlineamiento
  
    oDCOMConecta.AbreConexion
    lsSQL = "EXEC stp_sel_11356_CalculaProvicionSinLineamiento '" & psFecha & "'"
    Set lrDatosProvisionSinAli = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion

    If Not (lrDatosProvisionSinAli.EOF And lrDatosProvisionSinAli.BOF) Then
        If lrDatosProvisionSinAli!bResultado = 1 Then
            psMensaje = ""
        Else
            psMensaje = "Ha ocurrido un error. Por favor, Comunicarse con el Dpto. De TI"
        End If
    Else
        psMensaje = "No existen datos para la generación de la Calificacion"
        Exit Sub
    End If

    Set oDCOMConecta = Nothing
    Set lrDatosProvisionSinAli = Nothing
  
    Exit Sub
ErrorCalculaProvisionSinAlineamiento:
    Err.Raise Err.Number, "Calcula Provision Sin Alineamiento", Err.Description
End Sub

Public Sub calculaProvisionSistemaFinancieroNuevo(ByVal psFecha As String, ByRef psMensaje As String)
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lrDatosProvisionSistFinanc As ADODB.Recordset
    Dim lsCadConexion As String
    Set oDCOMConecta = New COMConecta.DCOMConecta
    
On Error GoTo ErrorCalculaProvisionSistemaFinanciero
   
    oDCOMConecta.AbreConexion
    lsSQL = "Exec stp_sel_11356_CalculaProvisionSistemaFinanciero '" & psFecha & "'"
    Set lrDatosProvisionSistFinanc = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion

    If Not (lrDatosProvisionSistFinanc.EOF And lrDatosProvisionSistFinanc.BOF) Then
        If lrDatosProvisionSistFinanc!bResultado = 1 Then
            psMensaje = ""
        Else
            psMensaje = "Ha ocurrido un error. Por favor, Comunicarse con el Dpto. De TI"
        End If
    Else
        psMensaje = "No existen datos para la generación de la Calificacion"
        Exit Sub
    End If

    Set oDCOMConecta = Nothing
    Set lrDatosProvisionSistFinanc = Nothing
  
    Exit Sub
ErrorCalculaProvisionSistemaFinanciero:
    Err.Raise Err.Number, "Calcula Provision Sistema Financiero", Err.Description
End Sub

Public Sub nCalificaCalculaProvisionProciclicaNuevo(ByVal psFecha As String, ByRef psMensaje As String)
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lrDatosProvisionProciclica As ADODB.Recordset
    Dim lsCadConexion As String
    Set oDCOMConecta = New COMConecta.DCOMConecta
    
On Error GoTo ErrorCalculaProvisionProciclica

    oDCOMConecta.AbreConexion
    lsSQL = "Exec stp_sel_11356_CalculaProvisionProciclica '" & psFecha & "'"
    Set lrDatosProvisionProciclica = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion

    If Not (lrDatosProvisionProciclica.EOF And lrDatosProvisionProciclica.BOF) Then
        If lrDatosProvisionProciclica!bResultado = 1 Then
            psMensaje = ""
        Else
            psMensaje = "Ha ocurrido un error. Por favor, Comunicarse con el Dpto. De TI"
        End If
    Else
        psMensaje = "No existen datos para la generación de la Calificacion"
        Exit Sub
    End If

    Set oDCOMConecta = Nothing
    Set lrDatosProvisionProciclica = Nothing
Exit Sub

ErrorCalculaProvisionProciclica:
    Err.Raise Err.Number, "Calcula Provision Prociclica", Err.Description
End Sub

Public Sub nCalificaRiesgoCredCambiarioNuevo(ByRef psMensaje As String)
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lrDatosRiesgoCambiario As ADODB.Recordset
    Dim lsCadConexion As String
    Set oDCOMConecta = New COMConecta.DCOMConecta

On Error GoTo ErrorCalificaRCambiario

    oDCOMConecta.AbreConexion
    lsSQL = "stp_sel_11356_CalificacionRiesgoCrediticioCambiario"
    Set lrDatosRiesgoCambiario = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion

    If Not (lrDatosRiesgoCambiario.EOF And lrDatosRiesgoCambiario.BOF) Then
        If lrDatosRiesgoCambiario!bResultado = 1 Then
            psMensaje = ""
        Else
            psMensaje = "Ha ocurrido un error. Por favor, Comunicarse con el Dpto. De TI"
        End If
    Else
        psMensaje = "No existen datos para la generación de la Calificacion"
        Exit Sub
    End If

    Set oDCOMConecta = Nothing
    Set lrDatosRiesgoCambiario = Nothing
    
Exit Sub
ErrorCalificaRCambiario:
    Err.Raise Err.Number, "Califica Riesgo Cambiario", Err.Description
End Sub

Public Sub nCalificaMetodologiaInternaNuevo(ByVal psFecha As String, ByRef psMensaje As String)
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim lrDatosMetodologiaInterna As ADODB.Recordset
    Dim lsCadConexion As String
    Set oDCOMConecta = New COMConecta.DCOMConecta
 
On Error GoTo ErrorMetodologiaInterna
 
    oDCOMConecta.AbreConexion
    lsSQL = "sp_sel_11356_CalificacionMetodologiaInterna '" & psFecha & "'"
    Set lrDatosMetodologiaInterna = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion

    If Not (lrDatosMetodologiaInterna.EOF And lrDatosMetodologiaInterna.BOF) Then
        If lrDatosMetodologiaInterna!bResultado = 1 Then
            psMensaje = ""
        Else
            psMensaje = "Ha ocurrido un error. Por favor, Comunicarse con el Dpto. De TI"
        End If
    Else
        psMensaje = "No existen datos para la generación de la Calificacion"
        Exit Sub
    End If

    Set oDCOMConecta = Nothing
    Set lrDatosMetodologiaInterna = Nothing
Exit Sub
 
ErrorMetodologiaInterna:
    Err.Raise Err.Number, "Califica Riesgo Cambiario", Err.Description

End Sub
'Fin LUCV20190601
'*CTI3 ERS0032020**********************************************************************
Public Sub nCalificaCalculaDatosSeguimiento(ByVal psFecha As String, ByRef psMensaje As String)
Dim lsSQL As String
Dim lrs As ADODB.Recordset
Dim rsVerif As ADODB.Recordset
Dim loConec As COMConecta.DCOMConecta

On Error GoTo ErrorCalificaCalculaDatosSeguimiento

    Set loConec = New COMConecta.DCOMConecta
    loConec.AbreConexion
    
    lsSQL = "stp_sel_11356_RSE_Seguimiento '" & psFecha & "'"


     Set lrs = loConec.CargaRecordSet(lsSQL)
     'loConec.CierraConexion
     If Not (lrs.EOF And lrs.BOF) Then
        If lrs!bResultado = 1 Then
            psMensaje = ""
        Else
            psMensaje = "Ha ocurrido un error. Por favor, Comunicarse con el Dpto. De TI"
        End If
    Else
        psMensaje = "No existen datos para el proceso de actualizar datos de seguimiento"
        Exit Sub
    End If

    Set loConec = Nothing
    Set lrs = Nothing
Exit Sub
    
ErrorCalificaCalculaDatosSeguimiento:
    Err.Raise Err.Number, "Calificacion Sin Alineamiento", Err.Description
End Sub
'**********************************************************************************//



