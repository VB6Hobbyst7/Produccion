VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nColocEvalCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************
'* Clase que contiene las Consultas para Calificacion de Cartera
'* LAYG - 01/11/2001
'****************************
Option Explicit
Dim mbTrans As Boolean

Public Function nObtieneCreditosCalificados() As Recordset

'************************************

Dim lsSQL As String
Dim loReg As DConecta
Dim lrDatos As ADODB.Recordset


    lsSQL = "SELECT   Per.CPersNombre Persona, CEval.cEvalCalif Cal, " _
        & " DocNat = (Select ISNULL(PerID.cPersIDnro,'') From PersId PerID " _
        & "          Where PerId.cPersCod = Per.cPersCod AND PerId.cPersIDTpo ='" & gPersIdDNI & "') ,  " _
        & " DocTri =(Select ISNULL(PerID.cPersIDnro,'') From PersId PerID " _
        & "          Where PerId.cPersCod = Per.cPersCod AND PerId.cPersIDTpo ='" & gPersIdRUC & "') ,  " _
        & " Per.cPersCod CodPers, ISNULL(CEval.cEvalObs,'') AS Obs " _
        & " FROM ColocEvalCalif CEval JOIN Persona Per ON Per.cPersCod = CEval.cPersCod " _
        & " ORDER BY Per.cPersNombre "

    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        MsgBox "ERROR: No existen datos ", vbInformation, "Aviso"
        Exit Function
    End If
    
    Set nObtieneCreditosCalificados = lrDatos
    Set lrDatos = Nothing
End Function
Public Function nobtineEvalconsulta(ByVal psCodPers As String, _
ByVal opt As Integer, ByVal pdFecha As Date, ByVal psServerConsol As String) As ADODB.Recordset
Dim lsSQL As String
Dim lsCad As String
Dim loReg As DConecta
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
    

    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    Set nobtineEvalconsulta = lrDatos
    If lrDatos.BOF And lrDatos.EOF Then
        MsgBox "ERROR: No existen datos ", vbInformation, "Aviso"
        Set lrDatos = Nothing
        Exit Function
    End If
    Set lrDatos = Nothing

End Function

Public Function nVerifExisteEvaluacion(ByVal psPersCod As String, ByVal pnEvalTipo As Integer) As Boolean
Dim lsSQL As String
Dim loReg As dColPFunciones
Dim lrExiste As ADODB.Recordset

    lsSQL = "SELECT cPersCod FROM ColocEvalCalif " _
          & "WHERE cPersCod ='" & psPersCod & "' And nEvalTipo = " & pnEvalTipo & " "
    
    Set loReg = New dColPFunciones
    Set lrExiste = loReg.dObtieneRecordSet(lsSQL)
    If lrExiste Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    
    If lrExiste.BOF And lrExiste.EOF Then
        nVerifExisteEvaluacion = False
    Else
        nVerifExisteEvaluacion = True
    End If
    
    Set lrExiste = Nothing
End Function

Public Function nObtieneDatosClienteCreditos(ByVal psCodPers As String, ByVal pdFechaFinMes As Date) As Recordset
'************************************

Dim lsSQL As String
Dim loReg As DConecta
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
        

    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        MsgBox "ERROR: No existen datos ", vbInformation, "Aviso"
        Exit Function
    End If
    
    Set nObtieneDatosClienteCreditos = lrDatos
    Set lrDatos = Nothing
End Function

Public Function nObtieneDatosClienteGarantias(ByVal psCodCta As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As DConecta
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
    
    
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        MsgBox "ERROR: No existen datos de Garantia de Credito " & psCodCta, vbInformation, "Aviso"
        Exit Function
    End If
    
    Set nObtieneDatosClienteGarantias = lrDatos
    Set lrDatos = Nothing

End Function




Public Function nObtieneDatosFuentesIngreso(ByVal psCodCta As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As DConecta
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
    
    
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        MsgBox "ERROR: No existen datos ", vbInformation, "Aviso"
        Exit Function
    End If
    
    Set nObtieneDatosFuentesIngreso = lrDatos
    Set lrDatos = Nothing

End Function

'************************************
Public Function nObtieneCreditosEvaluadosPersDetalles(ByVal psCodPers As String, ByVal pnEvalTipo As Integer) As Recordset

Dim lsSQL As String
Dim loReg As DConecta
Dim lrDatos As ADODB.Recordset


    lsSQL = "SELECT cCtacod as Credito, dEval as Fecha,  " _
        & " IsNull(cEvalCalifDet,'') as Cal, IsNull(cEvalObsDet,'') Obs , " _
        & " Isnull(nSaldoCap,0) as nSaldoCap, Isnull(nDiasAtraso,0) nDiasAtraso " _
        & " FROM ColocEvalCalifDetalle  " _
        & " WHERE cPersCod ='" & psCodPers & "' and nEvalTipo = " & pnEvalTipo _
        & " ORDER BY dEval, cCtaCod "
    
    Set loReg = New DConecta
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
Dim loConecta As DConecta

    lsSQL = "SELECT ISNULL(MAX(cEvalCalifDet),0) AS Maxima FROM ColocEvalCalifDetalle " _
          & "WHERE cPersCod  ='" & Trim(psPersCod) & "' "
    If psFecCal <> "@" Then
        lsSQL = lsSQL & " AND DATEDIFF(D,dEval, '" & Format(psFecCal, "mm/dd/yyyy") & "' ) = 0 "
    End If
    
    Set loConecta = New DConecta
        loConecta.AbreConexion
        Set loReg = loConecta.CargaRecordSet(lsSQL, adLockReadOnly)
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
Dim loReg As dColocEvalCal
Dim lnMovNro As Long
Dim lsCalifGen As String

'On Error GoTo ErrorCalif
Set loReg = New dColocEvalCal
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
Dim loReg As dColocEvalCal
Dim lnMovNro As Long
Dim lsCalifGen As String

'On Error GoTo ErrorCalif
Set loReg = New dColocEvalCal
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
Dim loReg As dColocEvalCal
Dim lnMovNro As Long
Dim lsCalifGen As String

'On Error GoTo ErrorCalif
Set loReg = New dColocEvalCal
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

Dim loReg As dColocEvalCal
Dim lnMovNro As Long
Dim lsCalifGen As String
'On Error GoTo ErrorCalif
Set loReg = New dColocEvalCal
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
Dim loReg As dColocEvalCal
Dim lnMovNro As Long
Dim lsCalifGen As String
On Error GoTo ErrorCalif
Set loReg = New dColocEvalCal
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
Dim loReg As dColocEvalCal
Dim lnMovNro As Long
Dim lsCalifGen As String
On Error GoTo ErrorCalif
Set loReg = New dColocEvalCal
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


Public Function nObtieneDatosEvaluacionAutomatica(ByVal psFechaCalif As Date, ByVal psFechaData As Date) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As DConecta
Dim lrDatos As ADODB.Recordset

lsSQL = " Select Apd.cPersCod, Apd.cCtaCod, Apd.nSaldoCap,  Apd.nDiasAtraso, APD.cEvalCalifDet " _
    & " From ColocEvalCalifDetalle Apd " _
    & " Where datediff(d, dEval,'" & Format(psFechaCalif, "yyyy/mm/dd") & "')= 0 " _
    & " And apd.cCtaCod not in  " _
    & " (Select cCtaCod From ColocEvalCalifDetalle " _
    & "     Where datediff(d, dEval,'" & Format(psFechaData, "yyyy/mm/dd") & "')= 0 ) "
    
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        MsgBox "No existen datos para Calificacion Automatica", vbInformation, "Aviso"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        MsgBox "No existen datos para Calificacion Automatica", vbInformation, "Aviso"
        Exit Function
    End If
    
    Set nObtieneDatosEvaluacionAutomatica = lrDatos
    Set lrDatos = Nothing

End Function

' Obtiene la calificación automatica del Credito
Public Function nGeneraCalificacionAutomatica(ByVal psCuenta As String, ByVal psRutaConsol As String) As String
Dim lsSQL As String
Dim loReg As DConecta
Dim lrDatos As ADODB.Recordset
Dim lsCalifxDA As String
Dim loCalif As nColocEvalCalProceso
   
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
        
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        'MsgBox "ERROR: No existen datos En Tabla de Parametros ", vbInformation, "Aviso"
        Exit Function
    End If
    
    'Obtengo la Calificacion
    lsCalifxDA = Trim(lrDatos!CalifxTabla)
    '*** Aplico las Reglas de Negocio
    Set loCalif = New nColocEvalCalProceso
        lsCalifxDA = loCalif.nCorrigeCalifxDiaAtraso(lsCalifxDA, lrDatos!nPrdEstado)
    Set loCalif = Nothing

    nGeneraCalificacionAutomatica = lsCalifxDA
    Set lrDatos = Nothing

End Function




