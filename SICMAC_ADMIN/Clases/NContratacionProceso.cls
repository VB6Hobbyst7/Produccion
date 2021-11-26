VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NContratacionProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Obtener_RHFichaPersonal_Anexos(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_RHFichaPersonal_Anexos
    sSQL = "SELECT cPersCod,Ubicacion= (select rtrim(isnull((select cdeszon from dbcomunes..zonas where ccodzon like '1'+substring(z.ccodzon,2,2)+'%'),''))+ ' - ' + " & _
          " rtrim(isnull((select cdeszon from dbcomunes..zonas where ccodzon like '2'+substring(z.ccodzon,2,4)+'%'),''))+ ' - ' + rtrim(isnull((select cdeszon from " & _
          " dbcomunes..zonas where ccodzon like '3'+substring(z.ccodzon,2,6)+'%'),''))+ ' - ' + RTrim (cdeszon) from dbcomunes..zonas z where ccodzon=cpersdireccUbiGeo) ," & _
          " cPersDireccUbiGeo,cPersDireccDomicilio,cPersTelfCelular,cPersDocOtros,cPersNomEmer, nParentescoEmer,Parentesco=C.cConsDescripcion, cPersTelefEmer, nTipoSangreEmer, " & _
          " TipoSangre=TS.cConsDescripcion,cPersAlergiaEmer,nTipoCategoriaBrevete,CategoriaBrevete=isnull(TB.cConsDescripcion,''),nNacionalidad=ISNULL(nNacionalidad,0),cPersLugarNacim=isnull(cPersLugarNacim,''),Nacionalidad=ISNULL(NAC.cConsDescripcion,'') " & _
          " FROM RHFichaPersonal_Anexos FPA INNER JOIN Constante C ON FPA.nParentescoEmer= C.nConsValor AND C.nConsCod=9066 INNER JOIN Constante TS ON FPA.nTipoSangreEmer = " & _
          " TS.nConsValor And TS.nConsCod = 1039 LEFT JOIN Constante TB ON FPA.nTipoCategoriaBrevete=TB.nConsValor AND TB.nConsCod=9068 LEFT JOIN Constante NAC ON NAC.nConsValor=FPA.nNacionalidad AND NAC.nConsCod=9004 " & _
          " WHERE FPA.cPersCod='" & pcPersCod & "'"

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_RHFichaPersonal_Anexos = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_RHFichaPersonal_Anexos:
    Err.Raise Err.Number, "Obtener_RHFichaPersonal_Anexos", Err.Description
End Function

Public Function Obtener_RHFichaPersonal_ExpLaboral(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_RHFichaPersonal_ExpLaboral
    sSQL = "SELECT cPersCod, nExpItem=ISNULL(nExpItem,0), cExpEmpresa=ISNULL(cExpEmpresa,''),cExpJefe=ISNULL(cExpJefe,''),cExpUltPuesto=ISNULL(cExpUltPuesto,''),cExpUltSueldo=ISNULL(cExpUltSueldo,0) ,cExpFecIngreso =ISNULL(cExpFecIngreso,''), " & _
            "cExpFecSalida=ISNULL(cExpFecSalida,''),cExpMotivoCese=ISNULL(cExpMotivoCese,'') From RHFichaPersonal_ExpLaboral " & _
            "WHERE cPersCod='" & pcPersCod & "'"

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_RHFichaPersonal_ExpLaboral = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_RHFichaPersonal_ExpLaboral:
    Err.Raise Err.Number, "Obtener_RHFichaPersonal_ExpLaboral", Err.Description
End Function

Public Function Obtener_RHFichaPersonal_Curriculum(ByVal pcPersCod As String, ByVal pnTipoCurriculum As Long) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_RHFichaPersonal_Curriculum
    
    Select Case pnTipoCurriculum
   
        Case 1
            sSQL = "SELECT cPersCod, FPC.cRHCurriculumCod, cRHCurriculumDescripcion=CASE WHEN FPC.cRHCurriculumCod=7 THEN cRHCurriComent ELSE cRHCurriculumDescripcion END," & _
                    " nRHCurriItem, dRHCurriInicio,dRHCurriFin, cRHCurriLugar, nRHCurriSituacion,Situacion=SIT.cConsDescripcion,cRHCurriGrado,nRHCurriEspec,Especialidad=ESP.cConsDescripcion," & _
                    " nRHCurriTipo, cRHCurriComent,bRHCurriOtorgadoCMACT=CONVERT(int,bRHCurriOtorgadoCMACT), FPC.cUltimaActualizacion FROM RHFichaPersonal_Curriculum FPC " & _
                    " INNER JOIN RHCurriculumTabla CT ON FPC.cRHCurriculumCod=CT.cRHCurriculumCod INNER JOIN Constante SIT ON SIT.nConsValor=FPC.nRHCurriSituacion AND SIT.nConsCod=9067" & _
                    " INNER JOIN Constante ESP ON ESP.nConsValor=FPC.nRHCurriEspec AND ESP.nConsCod=9050 " & _
                    " WHERE cPersCod='" & pcPersCod & "' AND nRHCurriTipo=1"
        Case 2
            sSQL = "SELECT cPersCod,nRHCurriItem,nCurriExtenxion=nRHCurriEspec,Extension=CASE WHEN (nRHCurriEspec= 20 AND EXT.nConsCod=6047) OR(nRHCurriEspec= 18 AND EXT.nConsCod=6050) " & _
                  " THEN EXT.cConsDescripcion ELSE '' END, nRHCurriNivel=nRHCurriSituacion,Nivel=NIV.cConsDescripcion,cRHCurriComent," & _
                  " bRHCurriOtorgadoCMACT=CONVERT(int,bRHCurriOtorgadoCMACT), FPC.cUltimaActualizacion FROM RHFichaPersonal_Curriculum FPC INNER JOIN Constante NIV ON " & _
                  " NIV.nConsValor=FPC.nRHCurriSituacion AND NIV.nConsCod=6051 INNER JOIN Constante EXT ON EXT.nConsValor=FPC.cRHCurriculumCod AND EXT.nConsCod IN(6050,6047) " & _
                  " WHERE cPersCod='" & pcPersCod & "' AND nRHCurriTipo=2"
        Case 3
            sSQL = "SELECT cPersCod, nRHCurriItem,Tema=cRHCurriComent, dRHCurriInicio, dRHCurriFin, nRHCurriSituacion, Situacion=SIT.cConsDescripcion,cRHCurriLugar, " & _
                   " bRHCurriOtorgadoCMACT=CONVERT(int,bRHCurriOtorgadoCMACT), FPC.cUltimaActualizacion FROM RHFichaPersonal_Curriculum FPC INNER JOIN Constante SIT " & _
                   " ON SIT.nConsValor=FPC.nRHCurriSituacion AND SIT.nConsCod=9067 " & _
                   " WHERE cPersCod='" & pcPersCod & "' AND nRHCurriTipo=3"
    End Select
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_RHFichaPersonal_Curriculum = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_RHFichaPersonal_Curriculum:
    Err.Raise Err.Number, "Obtener_RHFichaPersonal_Curriculum", Err.Description
End Function

Public Function Obtener_RHFichaPersonal_Respuesta(ByVal pcPersCod As String, _
                                                Optional ByVal pnTipoRespuesta As Long = 0, _
                                                Optional ByVal pnPreguntaCod As Long = 0) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_RHFichaPersonal_Respuesta
    
    If pnTipoRespuesta = 0 And pnPreguntaCod = 0 Then
        sSQL = "SELECT cPersCod, nPreguntaCod, nRespuesta ,Respuesta=C.cConsDescripcion FROM RHFichaPersonal_Respuesta FPR" & _
               " INNER JOIN Constante C ON FPR.nRespuesta=C.nConsValor " & _
               " WHERE cPersCod='" & pcPersCod & "'"
    Else
        sSQL = "SELECT cPersCod, nPreguntaCod, nRespuesta ,Respuesta=C.cConsDescripcion FROM RHFichaPersonal_Respuesta FPR" & _
               " INNER JOIN Constante C ON FPR.nRespuesta=C.nConsValor AND C.nConsCod=" & pnTipoRespuesta & _
               " WHERE cPersCod='" & pcPersCod & "' AND nPreguntaCod=" & pnPreguntaCod
    End If
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_RHFichaPersonal_Respuesta = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_RHFichaPersonal_Respuesta:
    Err.Raise Err.Number, "Obtener_RHFichaPersonal_Respuesta", Err.Description
End Function

Public Function Obtener_RHFichaPersonal_AnexoRespuesta(ByVal pcPersCod As String, _
                                                ByVal pnTipoRespuesta As Long, _
                                                ByVal pnPreguntaCod As Long, _
                                                ByVal pnItem As Long) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_RHFichaPersonal_AnexoRespuesta
    sSQL = "SELECT cPersCod, nPreguntaCod, nRespuesta ,cAnexo From RHFichaPersonal_AnexoRespuesta" & _
            " WHERE cPersCod='" & pcPersCod & "' AND nPreguntaCod=" & pnPreguntaCod & " AND nRespuesta=" & pnTipoRespuesta & " AND nItem=" & pnItem

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_RHFichaPersonal_AnexoRespuesta = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_RHFichaPersonal_AnexoRespuesta:
    Err.Raise Err.Number, "Obtener_RHFichaPersonal_AnexoRespuesta", Err.Description
End Function

Public Function Obtener_RHFichaPersonal_ExperienciaCaja(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_RHFichaPersonal_ExperienciaCaja
    sSQL = "SELECT car.dRHCargoFecha,CargoFecTermino=ISNULL(convert(varchar(10), (select top 1 drhCargoFecha from rhcargos where cperscod=car.cperscod " & _
           " and car.drhcargofecha<drhCargoFecha order by dRHCargoFecha asc),103),''), cat.crhcargodescripcion , age.cAgeDescripcion, crhcargocomentario " & _
           " FROM RHCargos car INNER JOIN RHCargosTabla cat ON car.crhCargoCod = cat.crhCargoCod INNER JOIN agencias age ON car.crhagenciacod=age.cagecod " & _
           " WHERE car.cPersCod='" & pcPersCod & "' order by car.dRHCargoFecha asc "
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_RHFichaPersonal_ExperienciaCaja = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_RHFichaPersonal_ExperienciaCaja:
    Err.Raise Err.Number, "Obtener_RHFichaPersonal_ExperienciaCaja", Err.Description
End Function

Public Function Obtener_Datos_RRHH(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_RRHH
    
    sSQL = "select cagedescripcion,careadescripcion,dingreso,dcese=isnull(CONVERT(varchar(10),dcese,103),'')," & _
           " NivelRemun=isnull(NIVEL.cNomTab,''),SubNivelRemun=isnull(SUBNIVEL.cNomTab,'')," & _
           " cCUSPP=isnull(Emp.cCUSPP,''),cNumAFP=isnull(Emp.cNumAFP,''),cRHCod=ISNULL(cRHCod,''),Cat.cRHCargoDescripcion " & _
           " FROM rrhh rh inner join " & _
           " RHEmpleado Emp on rh.cPersCod=Emp.cPersCod inner join agencias a on rh.cagenciaactual=a.cagecod " & _
           " inner join areas ar on ar.careacod=rh.careacodactual inner join RHCargos Car on rh.cPersCod=Car.cPersCod " & _
           " and drhcargofecha=(select top 1 dRHCargoFecha from rhcargos where cperscod=car.cperscod order by dRHCargoFecha desc)" & _
           " inner join RHCargosTabla Cat on Car.cRHCargoCod=Cat.cRHCargoCod inner join DBComunes..TablaCod SUBNIVEL ON " & _
           " Cat.crhniv_subNiv=SUBNIVEL.cValor inner join DBComunes..TablaCod NIVEL ON LEFT(SUBNIVEL.cValor,2)=NIVEL.cValor " & _
           " AND LEFT(NIVEL.cCodTab,2)= LEFT(SUBNIVEL.cCodTab,2)" & _
           " WHERE car.cperscod='" & pcPersCod & "'"
             
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_RRHH = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function
ErrorObtener_Datos_RRHH:
    Err.Raise Err.Number, "Obtener_Datos_RRHH", Err.Description
End Function

Public Function Obtener_Datos_Persona(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_Persona
    sSQL = " SELECT dpersnaccreac," & _
           " (select RTrim(isnull((select cdeszon from dbcomunes..zonas where ccodzon like '1'+substring(z.ccodzon,2,2)+'%'),''))+ ' - ' +" & _
           " RTrim(isnull((select cdeszon from dbcomunes..zonas where ccodzon like '2'+substring(z.ccodzon,2,4)+'%'),''))+ ' - ' +" & _
           " RTrim(isnull((select cdeszon from dbcomunes..zonas where ccodzon like '3'+substring(z.ccodzon,2,6)+'%'),''))+ ' - ' +" & _
           " RTrim (cdeszon)" & _
           " FROM dbcomunes..zonas z where ccodzon=cpersdireccubigeo)Ubicacion" & _
           " ,cpersdireccdomicilio,cpersemail,cperstelefono,cperstelefono2,cPersDireccUbigeo,cPersNombre from persona  where cPersCod='" & pcPersCod & "'"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_Persona = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function
ErrorObtener_Datos_Persona:
    Err.Raise Err.Number, "Obtener_Datos_Persona", Err.Description
End Function

Public Function Obtener_Datos_PersonaNat(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_PersonaNat
    sSQL = " select cpersnatsexo,estado=isnull((select cconsdescripcion from constante where nconscod=1020 and nconsvalor=npersnatestciv),'')," & _
           " npersnathijos,tposangre=isnull((select cconsdescripcion from constante where nconscod=1039 and nconsvalor=cperstposangre),'')," & _
           " nacionalidad=isnull((select cconsdescripcion from constante where nconscod=9004 and nconsvalor=cnacionalidad),'')" & _
           " from personanat  where cperscod='" & pcPersCod & "'"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_PersonaNat = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function
ErrorObtener_Datos_PersonaNat:
    Err.Raise Err.Number, "Obtener_Datos_PersonaNat", Err.Description
End Function

Public Function Obtener_Datos_Familiares(ByVal pcPersCod As String, Optional ByVal pnTipoGrado As Long = 0) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_Familiares
    
    If pnTipoGrado = 0 Then
        sSQL = " select p.cpersnombre,dpersnaccreac,cpersdireccubigeo,cpersdireccdomicilio," & _
           " isnull(cperstelefono,'')cperstelefono , isnull(cperstelefono2,'')cperstelefono2, PR.cpersrelacperscod, cconsdescripcion, npersrelac,cPersNatSexo,cPersIDnro=ISNULL(cPersIDnro,'')" & _
           " from dbo.PersRelaciones pr inner join persona p on pr.cpersrelacperscod=p.cperscod" & _
           " inner join PersonaNat Nat on Nat.cPersCod= P.cPersCod inner join constante c on c.nconsvalor=pr.npersrelac  and nconscod=1006" & _
           " left join PersId on PersId.cPersCod=P.cPersCod" & _
           " where pr.cperscod='" & pcPersCod & "'"
    Else
        Select Case pnTipoGrado
            Case 1
                sSQL = " select p.cpersnombre,dpersnaccreac,cpersdireccubigeo,cpersdireccdomicilio," & _
                       " isnull(cperstelefono,'')cperstelefono , isnull(cperstelefono2,'')cperstelefono2, PR.cpersrelacperscod, cconsdescripcion, npersrelac,cPersNatSexo,cPersIDnro=ISNULL(cPersIDnro,'')" & _
                       " from dbo.PersRelaciones pr inner join persona p on pr.cpersrelacperscod=p.cperscod" & _
                       " inner join PersonaNat Nat on Nat.cPersCod= P.cPersCod inner join constante c on c.nconsvalor=pr.npersrelac  and nconscod=1006 and nconsvalor IN(1,2)" & _
                       " left join PersId on PersId.cPersCod=P.cPersCod" & _
                       " where pr.cperscod='" & pcPersCod & "'"
            Case 2
                sSQL = " select p.cpersnombre,dpersnaccreac,cpersdireccubigeo,cpersdireccdomicilio," & _
                       " isnull(cperstelefono,'')cperstelefono , isnull(cperstelefono2,'')cperstelefono2, PR.cpersrelacperscod, cconsdescripcion, npersrelac,cPersNatSexo,cPersIDnro=ISNULL(cPersIDnro,'')" & _
                       " from dbo.PersRelaciones pr inner join persona p on pr.cpersrelacperscod=p.cperscod" & _
                       " inner join PersonaNat Nat on Nat.cPersCod= P.cPersCod inner join constante c on c.nconsvalor=pr.npersrelac  and nconscod=1006 and nconsvalor IN(3,4,11)" & _
                       " left join PersId on PersId.cPersCod=P.cPersCod" & _
                       " where pr.cperscod='" & pcPersCod & "'"
            Case 3
                sSQL = " select p.cpersnombre,dpersnaccreac,cpersdireccubigeo,cpersdireccdomicilio," & _
                       " isnull(cperstelefono,'')cperstelefono , isnull(cperstelefono2,'')cperstelefono2, PR.cpersrelacperscod, cconsdescripcion, npersrelac,cPersNatSexo,cPersIDnro=ISNULL(cPersIDnro,'')" & _
                       " from dbo.PersRelaciones pr inner join persona p on pr.cpersrelacperscod=p.cperscod" & _
                       " inner join PersonaNat Nat on Nat.cPersCod= P.cPersCod inner join constante c on c.nconsvalor=pr.npersrelac  and nconscod=1006 and nconsvalor IN(14,7)" & _
                       " left join PersId on PersId.cPersCod=P.cPersCod" & _
                       " where pr.cperscod='" & pcPersCod & "'"
            Case 4
                sSQL = " select p.cpersnombre,dpersnaccreac,cpersdireccubigeo,cpersdireccdomicilio," & _
                       " isnull(cperstelefono,'')cperstelefono , isnull(cperstelefono2,'')cperstelefono2, PR.cpersrelacperscod, cconsdescripcion, npersrelac,cPersNatSexo,cPersIDnro=ISNULL(cPersIDnro,'')" & _
                       " from dbo.PersRelaciones pr inner join persona p on pr.cpersrelacperscod=p.cperscod" & _
                       " inner join PersonaNat Nat on Nat.cPersCod= P.cPersCod inner join constante c on c.nconsvalor=pr.npersrelac  and nconscod=1006 and nconsvalor IN(0,12)" & _
                       " left join PersId on PersId.cPersCod=P.cPersCod" & _
                       " where pr.cperscod='" & pcPersCod & "'"
            Case 5
                sSQL = " select p.cpersnombre,dpersnaccreac,cpersdireccubigeo,cpersdireccdomicilio," & _
                       " isnull(cperstelefono,'')cperstelefono , isnull(cperstelefono2,'')cperstelefono2, PR.cpersrelacperscod, cconsdescripcion, npersrelac,cPersNatSexo,cPersIDnro=ISNULL(cPersIDnro,'')" & _
                       " from dbo.PersRelaciones pr inner join persona p on pr.cpersrelacperscod=p.cperscod" & _
                       " inner join PersonaNat Nat on Nat.cPersCod= P.cPersCod inner join constante c on c.nconsvalor=pr.npersrelac  and nconscod=1006 and nconsvalor =13" & _
                       " left join PersId on PersId.cPersCod=P.cPersCod" & _
                       " where pr.cperscod='" & pcPersCod & "'"
        End Select
    End If
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_Familiares = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function
ErrorObtener_Datos_Familiares:
    Err.Raise Err.Number, "Obtener_Datos_Familiares", Err.Description
End Function

Public Function Obtener_Documentos_Persona(ByVal pcPersCod As String, Optional ByVal pnTipoDoc As Integer = 0) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Documentos_Persona
    If pnTipoDoc = 0 Then
        sSQL = " select nconsvalor,cconsdescripcion,doc=isnull((select cpersidnro from persid where cperscod='" & pcPersCod & "' and cpersidtpo=nconsvalor ),'')" & _
               " from constante where nconscod=1003" & _
               " and nconsvalor in (1,2,9,7,13,6,4)"
    Else
        sSQL = " select nconsvalor,cconsdescripcion,doc=isnull((select cpersidnro from persid where cperscod='" & pcPersCod & "' and cpersidtpo=nconsvalor ),'')" & _
               " from constante where nconscod=1003" & _
               " and nconsvalor =" & pnTipoDoc
    End If
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Documentos_Persona = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Documentos_Persona:
    Err.Raise Err.Number, "Obtener_Documentos_Persona", Err.Description
End Function
