VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DContratacionProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InsertaRHFichaPersonal_Curriculum(ByVal pcPersCod As String, _
                    ByVal pcRHCurriculumCod As String, _
                    ByVal pnRHCurriItem As Long, _
                    ByVal pdRHCurriInicio As Date, _
                    ByVal pdRHCurriFin As Date, _
                    ByVal pcRHCurriLugar As String, _
                    ByVal pnRHCurriSituacion As Long, _
                    ByVal pcRHCurriGrado As String, _
                    ByVal pnRHCurriEspec As Long, _
                    ByVal pnRHCurriTipo As Long, _
                    ByVal pcRHCurriComent As String, _
                    ByVal pbRHCurriOtorgadoCMACT As Integer, _
                    ByVal pcUltimaActualizacion As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHFichaPersonal_Curriculum
    sSQL = "INSERT INTO RHFichaPersonal_Curriculum(cPersCod,cRHCurriculumCod,nRHCurriItem,dRHCurriInicio,dRHCurriFin,cRHCurriLugar,nRHCurriSituacion,cRHCurriGrado,nRHCurriEspec,nRHCurriTipo,cRHCurriComent,bRHCurriOtorgadoCMACT, cUltimaActualizacion)"
    sSQL = sSQL & " VALUES('" & pcPersCod & "','" & pcRHCurriculumCod & "'," & pnRHCurriItem & ",'" & Format(pdRHCurriInicio, gsFormatoFecha) & "','" & Format(pdRHCurriFin, gsFormatoFecha) & "','" & pcRHCurriLugar & "'," & pnRHCurriSituacion & ",'" & pcRHCurriGrado & "'," & pnRHCurriEspec & "," & pnRHCurriTipo & ",'" & pcRHCurriComent & "'," & pbRHCurriOtorgadoCMACT & ",'" & pcUltimaActualizacion & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHFichaPersonal_Curriculum:
    Err.Raise Err.Number, "InsertaRHFichaPersonal_Curriculum", Err.Description
End Sub

Public Sub ModificaRHFichaPersonal_Curriculum(ByVal pcPersCod As String, _
                    ByVal pcRHCurriculumCod As String, _
                    ByVal pnRHCurriItem As Long, _
                    ByVal pdRHCurriInicio As Date, _
                    ByVal pdRHCurriFin As Date, _
                    ByVal pcRHCurriLugar As String, _
                    ByVal pnRHCurriSituacion As Long, _
                    ByVal pcRHCurriGrado As String, _
                    ByVal pnRHCurriEspec As Long, _
                    ByVal pnRHCurriTipo As Long, _
                    ByVal pcRHCurriComent As String, _
                    ByVal pbRHCurriOtorgadoCMACT As Integer, _
                    ByVal pcUltimaActualizacion As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHFichaPersonal_Curriculum
    sSQL = "UPDATE RHFichaPersonal_Curriculum" '
    sSQL = sSQL & " SET cPersCod='" & pcPersCod & _
                    "',cRHCurriculumCod='" & pcRHCurriculumCod & _
                    "',nRHCurriItem=" & pnRHCurriItem & _
                    ",dRHCurriInicio='" & Format(pdRHCurriInicio, gsFormatoFecha) & _
                    "',dRHCurriFin='" & Format(pdRHCurriFin, gsFormatoFecha) & _
                    "',cRHCurriLugar='" & pcRHCurriLugar & _
                    "',nRHCurriSituacion=" & pnRHCurriSituacion & _
                    ",cRHCurriGrado='" & pcRHCurriGrado & _
                    "',nRHCurriEspec=" & pnRHCurriEspec & _
                    ",nRHCurriTipo=" & pnRHCurriTipo & _
                    ",cRHCurriComent='" & pcRHCurriComent & _
                    "',bRHCurriOtorgadoCMACT=" & pbRHCurriOtorgadoCMACT & _
                    ",cUltimaActualizacion='" & pcUltimaActualizacion & _
                    "' WHERE cPersCod='" & pcPersCod & "' AND cRHCurriculumCod='" & pcRHCurriculumCod & "' AND nRHCurriItem=" & pnRHCurriItem
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHFichaPersonal_Curriculum:
    Err.Raise Err.Number, "ModificaRHFichaPersonal_Curriculum", Err.Description
End Sub

Public Function EliminaRHFichaPersonal_Curriculum(ByVal pcPersCod As String, _
                                                ByVal pcRHCurriculumCod As String, _
                                                ByVal pnRHCurriItem As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHFichaPersonal_Curriculum
    sSQL = "DELETE FROM RHFichaPersonal_Curriculum" '
    sSQL = sSQL & " WHERE cPersCod='" & pcPersCod & "' AND cRHCurriculumCod='" & pcRHCurriculumCod & "' AND nRHCurriItem=" & pnRHCurriItem
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHFichaPersonal_Curriculum:
    Err.Raise Err.Number, "EliminaRHFichaPersonal_Curriculum", Err.Description
End Function


Public Sub InsertaRHFichaPersonal_ExpLaboral(ByVal pcPersCod As String, _
                    ByVal pnExpItem As Long, _
                    ByVal pcExpEmpresa As String, _
                    ByVal pcExpJefe As String, _
                    ByVal pcExpUltPuesto As String, _
                    ByVal pcExpUltSueldo As Double, _
                    ByVal pdExpFecIngreso As Date, _
                    ByVal pdExpFecSalida As Date, _
                    ByVal pcExpMotivoCese As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHFichaPersonal_ExpLaboral
    sSQL = "INSERT INTO RHFichaPersonal_ExpLaboral(cPersCod, nExpItem,cExpEmpresa,cExpJefe,cExpUltPuesto,cExpUltSueldo,cExpFecIngreso,cExpFecSalida,cExpMotivoCese)"
    sSQL = sSQL & " VALUES('" & pcPersCod & "'," & pnExpItem & ",'" & pcExpEmpresa & "','" & pcExpJefe & "','" & pcExpUltPuesto & "'," & pcExpUltSueldo & ",'" & Format(pdExpFecIngreso, gsFormatoFecha) & "','" & Format(pdExpFecSalida, gsFormatoFecha) & "','" & pcExpMotivoCese & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHFichaPersonal_ExpLaboral:
    Err.Raise Err.Number, "InsertaRHFichaPersonal_ExpLaboral", Err.Description
End Sub

Public Sub ModificaRHFichaPersonal_ExpLaboral(ByVal pcPersCod As String, _
                    ByVal pnExpItem As Long, _
                    ByVal pcExpEmpresa As String, _
                    ByVal pcExpJefe As String, _
                    ByVal pcExpUltPuesto As String, _
                    ByVal pcExpUltSueldo As Double, _
                    ByVal pdExpFecIngreso As Date, _
                    ByVal pdExpFecSalida As Date, _
                    ByVal pcExpMotivoCese As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHFichaPersonal_ExpLaboral
    sSQL = "UPDATE RHFichaPersonal_ExpLaboral" '
    sSQL = sSQL & " SET cPersCod='" & pcPersCod & _
                    "',nExpItem=" & pnExpItem & _
                    ",cExpEmpresa='" & pcExpEmpresa & _
                    "',cExpJefe='" & pcExpJefe & _
                    "',cExpUltPuesto='" & pcExpUltPuesto & _
                    "',cExpUltSueldo=" & pcExpUltSueldo & _
                    ",dExpFecIngreso='" & Format(pdExpFecIngreso, gsFormatoFecha) & _
                    "',dExpFecSalida='" & Format(pdExpFecSalida, gsFormatoFecha) & _
                    "',cExpMotivoCese='" & pcExpMotivoCese & _
                    "' WHERE cPersCod='" & pcPersCod & "' AND nExpItem=" & pnExpItem
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHFichaPersonal_ExpLaboral:
    Err.Raise Err.Number, "ModificaRHFichaPersonal_ExpLaboral", Err.Description
End Sub

Public Function EliminaRHFichaPersonal_ExpLaboral(ByVal pcPersCod As String, _
                                                ByVal pnExpItem As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHFichaPersonal_ExpLaboral
    sSQL = "DELETE FROM RHFichaPersonal_ExpLaboral" '
    sSQL = sSQL & " WHERE cPersCod='" & pcPersCod & "' AND nExpItem=" & pnExpItem
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHFichaPersonal_ExpLaboral:
    Err.Raise Err.Number, "EliminaRHFichaPersonal_ExpLaboral", Err.Description
End Function

Public Sub InsertaRHFichaPersonal_Respuesta(ByVal pcPersCod As String, _
                                            ByVal pnPreguntaCod As Long, _
                                            ByVal pnRespuesta As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHFichaPersonal_Respuesta
    sSQL = "INSERT INTO RHFichaPersonal_Respuesta(cPersCod,nPreguntaCod,nRespuesta)"
    sSQL = sSQL & " VALUES('" & pcPersCod & "'," & pnPreguntaCod & "," & pnRespuesta & ")"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHFichaPersonal_Respuesta:
    Err.Raise Err.Number, "InsertaRHFichaPersonal_Respuesta", Err.Description
End Sub

Public Sub ModificaRHFichaPersonal_Respuesta(ByVal pcPersCod As String, _
                                            ByVal pnPreguntaCod As Long, _
                                            ByVal pnRespuesta As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHFichaPersonal_Respuesta
    sSQL = "UPDATE RHFichaPersonal_Respuesta" '
    sSQL = sSQL & " SET cPersCod='" & pcPersCod & _
                    "',nPreguntaCod=" & pnPreguntaCod & _
                    ",nRespuesta=" & pnRespuesta & _
                    " WHERE cPersCod='" & pcPersCod & "' AND nPreguntaCod=" & pnPreguntaCod & " AND nRespuesta=" & pnRespuesta
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHFichaPersonal_Respuesta:
    Err.Raise Err.Number, "ModificaRHFichaPersonal_Respuesta", Err.Description
End Sub

Public Function EliminaRHFichaPersonal_Respuesta(ByVal pcPersCod As String, _
                                                ByVal pnPreguntaCod As Long, _
                                                ByVal pnRespuesta As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHFichaPersonal_Respuesta
    sSQL = "DELETE FROM RHFichaPersonal_Respuesta" '
    sSQL = sSQL & " WHERE cPersCod='" & pcPersCod & "' AND nPreguntaCod=" & pnPreguntaCod & " AND nRespuesta=" & pnRespuesta
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHFichaPersonal_Respuesta:
    Err.Raise Err.Number, "EliminaRHFichaPersonal_Respuesta", Err.Description
End Function

Public Sub InsertaRHFichaPersonal_AnexoRespuesta(ByVal pcPersCod As String, _
                                            ByVal pnPreguntaCod As Long, _
                                            ByVal pnRespuesta As Long, _
                                            ByVal pcAnexo As String, _
                                            ByVal pnItem As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHFichaPersonal_AnexoRespuesta
    sSQL = "INSERT INTO RHFichaPersonal_AnexoRespuesta(cPersCod,nPreguntaCod,nRespuesta,cAnexo,nItem)"
    sSQL = sSQL & " VALUES('" & pcPersCod & "'," & pnPreguntaCod & "," & pnRespuesta & ",'" & pcAnexo & "'," & pnItem & ")"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHFichaPersonal_AnexoRespuesta:
    Err.Raise Err.Number, "InsertaRHFichaPersonal_AnexoRespuesta", Err.Description
End Sub

Public Sub ModificaRHFichaPersonal_AnexoRespuesta(ByVal pcPersCod As String, _
                                            ByVal pnPreguntaCod As Long, _
                                            ByVal pnRespuesta As Long, _
                                            ByVal pcAnexo As String, _
                                            ByVal pnItem As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHFichaPersonal_AnexoRespuesta
    sSQL = "UPDATE RHFichaPersonal_AnexoRespuesta" '
    sSQL = sSQL & " SET cPersCod='" & pcPersCod & _
                    "',nPreguntaCod=" & pnPreguntaCod & _
                    ",nRespuesta=" & pnRespuesta & _
                    ",cAnexo='" & pcAnexo & _
                    "', nItem=" & pnItem & _
                    " WHERE cPersCod='" & pcPersCod & "' AND nPreguntaCod=" & pnPreguntaCod & " AND nRespuesta=" & pnRespuesta
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHFichaPersonal_AnexoRespuesta:
    Err.Raise Err.Number, "ModificaRHFichaPersonal_AnexoRespuesta", Err.Description
End Sub

Public Function EliminaRHFichaPersonal_AnexoRespuesta(ByVal pcPersCod As String, _
                                                ByVal pnPreguntaCod As Long, _
                                                ByVal pnRespuesta As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHFichaPersonal_AnexoRespuesta
    sSQL = "DELETE FROM RHFichaPersonal_AnexoRespuesta" '
    sSQL = sSQL & " WHERE cPersCod='" & pcPersCod & "' AND nPreguntaCod=" & pnPreguntaCod & " AND nRespuesta=" & pnRespuesta
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHFichaPersonal_AnexoRespuesta:
    Err.Raise Err.Number, "EliminaRHFichaPersonal_AnexoRespuesta", Err.Description
End Function

Public Sub InsertaRHFichaPersonal_Anexos(ByVal pcPersCod As String, _
                                            ByVal pcPersDireccUbigeo As String, _
                                            ByVal pcPersDireccDomicilio As String, _
                                            ByVal pcPersTelfCelular As String, _
                                            ByVal pcPersDocOtros As String, _
                                            ByVal pcPersNomEmer As String, _
                                            ByVal pnParentescoEmer As Long, _
                                            ByVal pcPersTelefEmer As String, _
                                            ByVal pnTipoSangreEmer As Long, _
                                            ByVal pcPersAlergiaEmer As String, _
                                            ByVal pnTipoCategoriaBrevete As Long, _
                                            ByVal pcPersLugarNacim As String, _
                                            ByVal pnNacionalidad As Integer)
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorInsertaRHFichaPersonal_Anexos
        
    sSQL = "INSERT INTO RHFichaPersonal_Anexos(cPersCod,cPersDireccUbiGeo,cPersDireccDomicilio,cPersTelfCelular,cPersDocOtros,cPersNomEmer,nParentescoEmer,cPersTelefEmer,nTipoSangreEmer,cPersAlergiaEmer,nTipoCategoriaBrevete,cPersLugarNacim,nNacionalidad)"
    sSQL = sSQL & " VALUES('" & pcPersCod & "','" & pcPersDireccUbigeo & "','" & pcPersDireccDomicilio & "','" & pcPersTelfCelular & "','" & pcPersDocOtros & "','" & pcPersNomEmer & "'," & pnParentescoEmer & ",'" & pcPersTelefEmer & "'," & pnTipoSangreEmer & ",'" & pcPersAlergiaEmer & "'," & pnTipoCategoriaBrevete & ",'" & pcPersLugarNacim & "'," & pnNacionalidad & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHFichaPersonal_Anexos:
    Err.Raise Err.Number, "InsertaRHFichaPersonal_Anexos", Err.Description
End Sub

Public Sub ModificaRHFichaPersonal_Anexos(ByVal pcPersCod As String, _
                                            ByVal pcPersDireccUbigeo As String, _
                                            ByVal pcPersDireccDomicilio As String, _
                                            ByVal pcPersTelfCelular As String, _
                                            ByVal pcPersDocOtros As String, _
                                            ByVal pcPersNomEmer As String, _
                                            ByVal pnParentescoEmer As Long, _
                                            ByVal pcPersTelefEmer As String, _
                                            ByVal pnTipoSangreEmer As Long, _
                                            ByVal pcPersAlergiaEmer As String, _
                                            ByVal pnTipoCategoriaBrevete As Long, _
                                            ByVal pcPersLugarNacim As String, _
                                            ByVal pnNacionalidad As Integer)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHFichaPersonal_Anexos
    sSQL = "UPDATE RHFichaPersonal_Anexos" '
    sSQL = sSQL & " SET cPersCod='" & pcPersCod & _
                    "',cPersDireccUbiGeo='" & pcPersDireccUbigeo & _
                    "',cPersDireccDomicilio='" & pcPersDireccDomicilio & _
                    "',cPersTelfCelular='" & pcPersTelfCelular & _
                    "',cPersDocOtros='" & pcPersDocOtros & _
                    "',cPersNomEmer='" & pcPersNomEmer & _
                    "',nParentescoEmer=" & pnParentescoEmer & _
                    ",cPersTelefEmer='" & pcPersTelefEmer & _
                    "',nTipoSangreEmer=" & pnTipoSangreEmer & _
                    ",cPersAlergiaEmer='" & pcPersAlergiaEmer & _
                    "',nTipoCategoriaBrevete=" & pnTipoCategoriaBrevete & _
                    ",cPersLugarNacim='" & pcPersLugarNacim & _
                    "',nNacionalidad=" & pnNacionalidad & _
                    " WHERE cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHFichaPersonal_Anexos:
    Err.Raise Err.Number, "ModificaRHFichaPersonal_Anexos", Err.Description
End Sub

Public Function EliminaRHFichaPersonal_Anexos(ByVal pcPersCod As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHFichaPersonal_Anexos
    sSQL = "DELETE FROM RHFichaPersonal_Anexos" '
    sSQL = sSQL & " WHERE cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHFichaPersonal_Anexos:
    Err.Raise Err.Number, "EliminaRHFichaPersonal_Anexos", Err.Description
End Function

'
'Public Sub InsertaRHFichaPersonal_Persona(ByVal pcPersCod As String, _
'                                            ByVal pcPersDireccUbigeo As String, _
'                                            ByVal pcPersDireccDomicilio As String, _
'                                            ByVal pcPersTelefono As String, _
'                                            ByVal pcPersEmail As String, _
'                                            ByVal pcPersNatHijos As String, _
'                                            ByVal pcPersNatEstCiv As String, _
'                                            ByVal pnPersTpoSangre As String)
'Dim sSQL As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorInsertaRHFichaPersonal_Persona
'
'    sSQL = "INSERT INTO RHFichaPersonal_Persona(cPersCod,cPersDireccUbiGeo,cPersDireccDomicilio,cPersTelfCelular,cPersDocOtros,cPersNomEmer,nParentescoEmer,cPersTelefEmer,nTipoSangreEmer,cPersAlergiaEmer,nTipoCategoriaBrevete,cPersLugarNacim)"
'    sSQL = sSQL & " VALUES('" & pcPersCod & "','" & pcPersDireccUbigeo & "','" & pcPersDireccDomicilio & "','" & pcPersTelfCelular & "','" & pcPersDocOtros & "','" & pcPersNomEmer & "," & pnParentescoEmer & ",'" & pcPersTelefEmer & "'," & pnTipoSangreEmer & ",'" & pcPersAlergiaEmer & "'," & pnTipoCategoriaBrevete & ",'" & pcPersLugarNacim & "')"
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSQL
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Sub
'
'ErrorInsertaRHFichaPersonal_Persona:
'    Err.Raise Err.Number, "InsertaRHFichaPersonal_Persona", Err.Description
'End Sub

Public Sub ModificaRHFichaPersonal_Persona(ByVal pcPersCod As String, _
                                            ByVal pcPersDireccUbigeo As String, _
                                            ByVal pcPersDireccDomicilio As String, _
                                            ByVal pcPersTelefono As String, _
                                            ByVal pcPersEmail As String, _
                                            ByVal pcPersNatHijos As String, _
                                            ByVal pnPersNatEstCiv As Long, _
                                            ByVal pnPersTpoSangre As Long)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHFichaPersonal_Persona
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    
    sSQL = "UPDATE Persona" '
    sSQL = sSQL & " SET cPersDireccUbiGeo='" & pcPersDireccUbigeo & _
                    "',cPersDireccDomicilio='" & pcPersDireccDomicilio & _
                    "',cPersTelefono='" & pcPersTelefono & _
                    "',cPersEmail='" & pcPersEmail & _
                    "' WHERE cPersCod='" & pcPersCod & "'"
                        
    oConecta.ConexionActiva.Execute sSQL
    
    sSQL = "UPDATE PersonaNat" '
    sSQL = sSQL & " SET nPersNatHijos='" & pcPersNatHijos & _
                    "',nPersNatEstCiv=" & pnPersNatEstCiv & _
                    ",nPersTpoSangre=" & pnPersTpoSangre & _
                    " WHERE cPersCod='" & pcPersCod & "'"
                        
    oConecta.ConexionActiva.Execute sSQL
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHFichaPersonal_Persona:
    Err.Raise Err.Number, "ModificaRHFichaPersonal_Persona", Err.Description
End Sub

Public Sub InsertaRHEmpleadoDocumento(ByVal pcPersCod As String, _
                    ByVal pnTipo As Long, _
                    ByVal pnEstado As Long, _
                    ByVal pdFechaEntrega As Date, _
                    ByVal pcUltimaActualizacion As String)
                    
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEmpleadoDocumento
    
    sSQL = "INSERT INTO RHEmpleadoDocumento(cPersCod,nTipo,nEstado,dFechaEntrega,cUltimaActualizacion)"
    sSQL = sSQL & " VALUES('" & pcPersCod & "'," & pnTipo & "," & pnEstado & ",'" & Format(pdFechaEntrega, gsFormatoFecha) & "','" & pcUltimaActualizacion & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEmpleadoDocumento:
    Err.Raise Err.Number, "InsertaRHEmpleadoDocumento", Err.Description
End Sub

Public Sub ModificaRHEmpleadoDocumento(ByVal pcPersCod As String, _
                                        ByVal pnTipo As Long, _
                                        ByVal pnEstado As Long, _
                                        ByVal pdFechaEntrega As Date, _
                                        ByVal pcUltimaActualizacion As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHEmpleadoDocumento
    sSQL = "UPDATE RHEmpleadoDocumento"
    sSQL = sSQL & " SET nEstado=" & pnEstado & _
                    ",dFechaEntrega='" & Format(pdFechaEntrega, gsFormatoFecha) & _
                    "',cUltimaActualizacion='" & pcUltimaActualizacion & _
                    "' WHERE cPersCod='" & pcPersCod & "' AND nTipo=" & pnTipo
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEmpleadoDocumento:
    Err.Raise Err.Number, "ModificaRHInduccion", Err.Description
End Sub

Public Function EliminaRHEmpleadoDocumento(ByVal pcPersCod As String, _
                                        ByVal pnTipo As Long)

Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEmpleadoDocumento
    sSQL = "DELETE FROM RHEmpleadoDocumento" '
    sSQL = sSQL & "' WHERE cPersCod='" & pcPersCod & "' AND nTipo=" & pnTipo
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEmpleadoDocumento:
    Err.Raise Err.Number, "EliminaRHEmpleadoDocumento", Err.Description
End Function

Public Function Obtener_DctosEntregados_x_Persona(ByVal pcPersCod As String, Optional ByVal pnTipoDoc As Long = 0) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_DctosEntregados_x_Persona
    If pnTipoDoc = 0 Then
        sSQL = "SELECT nTipo,Documento=C.cConsDescripcion ,nEstado FROM RHEmpleadoDocumento " & _
               " INNER JOIN CONSTANTE C ON C.nConsValor=nTipo AND nConsCod=9070" & _
               " WHERE cPersCod='" & pcPersCod & "'" & IIf(pnTipoDoc = 0, "", " AND nTipo=" & pnTipoDoc)
    End If
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_DctosEntregados_x_Persona = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_DctosEntregados_x_Persona:
    Err.Raise Err.Number, "Obtener_DctosEntregados_x_Persona", Err.Description
End Function

Public Function Obtener_Personas_Faltantes_EntregaDocs(Optional ByVal pnTipoDoc As Long = 0) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Personas_Faltantes_EntregaDocs
    
    sSQL = "SELECT DISTINCT P.cPersCod,P.cPersNombre FROM RHEmpleadoDocumento EmpDoc INNER JOIN RHEmpleado Emp" & _
           " ON EmpDoc.cPersCod= Emp.cPersCod INNER JOIN Persona P ON P.cPersCod= Emp.cPersCod " & _
           " WHERE EmpDoc.nEstado = 0 " & IIf(pnTipoDoc > 0, " And EmpDoc.nTipo = " & pnTipoDoc, "")
     
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Personas_Faltantes_EntregaDocs = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Personas_Faltantes_EntregaDocs:
    Err.Raise Err.Number, "Obtener_Personas_Faltantes_EntregaDocs", Err.Description
End Function

Public Function Obtener_Documentos_x_Entregar() As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Documentos_x_Entregar
    
    sSQL = "SELECT nConsValor,'0',cConsDescripcion FROM Constante WHERE nConsCod=9070 AND nConsValor<>nConsCod"
     
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Documentos_x_Entregar = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Documentos_x_Entregar:
    Err.Raise Err.Number, "Obtener_Documentos_x_Entregar", Err.Description
End Function

Public Function Obtener_Contratados_x_Proceso(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Contratados_x_Proceso
    
    'Buscar si ya se han contratado(es decir se Asigno un cPersCod al Postulante)
    sSQL = "SELECT SP.cPersCod,cPostuCod,cPersIDnro,cPersNombre FROM RHSeleccionPostulante SP " & _
           " INNER JOIN Persona P ON SP.cPersCod=P.cPersCod INNER JOIN PersID ON P.cPersCod=PersID.cPersCod " & _
           " AND PersID.cPersIDTpo='1' WHERE ISNULL(SP.cPersCod,'')<>'' AND cCodProceso='" & pcCodProceso & "'"
     
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Contratados_x_Proceso = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Contratados_x_Proceso:
    Err.Raise Err.Number, "Obtener_Contratados_x_Proceso", Err.Description
End Function

Public Function Obtener_Datos_Para_Contrato(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_Para_Contrato
    
    sSQL = "SELECT RH.cRHCod,P.cPersNombre,RH.dIngreso,CD.dRHContratoInicio,CD.dRHContratoFin," & _
           " P.CPersDireccDomicilio,CT.cRHCargoDescripcion,c.cConsDescripcion , s.nRHSueldoMonto, PID.cPersIDnro " & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod INNER JOIN RHContrato CON ON P.cPersCod=CON.cPersCod" & _
           " AND CON.cRHContratoNro=(SELECT MAX(CONVERT(int,cRHContratoNro))FROM RHContrato WHERE cPersCod=P.cPersCod)" & _
           " INNER JOIN RHContratoDet CD ON CON.cPersCod=CD.cPersCod AND CON.cRHContratoNro=CD.cRHContratoNro AND " & _
           " CD.dRHContratoFecha=(SELECT MAX(dRHContratoFecha) FROM RHContratoDet WHERE cPersCod=CON.cPersCod AND " & _
           " cRHContratoNro=CON.cRHContratoNro)INNER JOIN RHSueldo S ON S.cPersCod=P.cPersCod AND S.dRHSueldoFecha=" & _
           "(SELECT MAX(dRHSueldoFecha) FROM RHSueldo WHERE cPersCod=P.cPersCod) INNER JOIN Constante C ON " & _
           " CON.nRHContratoTpo=C.nConsValor AND C.nConsCod=6012 INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND " & _
           " CAR.dRHCargoFecha= (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod)" & _
           " INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod INNER JOIN PersID PID ON " & _
           " P.cPersCod=PID.cPersCod AND PID.cPersIDTpo=1 " & _
           " WHERE P.cPersCod='" & pcPersCod & "'"
     
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_Para_Contrato = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Datos_Para_Contrato:
    Err.Raise Err.Number, "Obtener_Datos_Para_Contrato", Err.Description
End Function

Public Function Obtener_Personal_Ingresante(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Personal_Ingresante
    
    sSQL = " SELECT P.cPersCod,P.cPersNombre,dRHContratoInicio=CONVERT(varchar(10),CD.dRHContratoInicio,103),dRHContratoFin=CONVERT(varchar(10),CD.dRHContratoFin,103),s.nRHSueldoMonto , CT.cRHCargoDescripcion, dRHContratoFecha=CONVERT(varchar(10),CD.dRHContratoFecha,103),CT.cRHCargoCod,CD.cRHContratoNro " & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod INNER JOIN RHContrato CON ON P.cPersCod=CON.cPersCod " & _
           " AND CON.cRHContratoNro=(SELECT MAX(CONVERT(int,cRHContratoNro))FROM RHContrato WHERE cPersCod=P.cPersCod) " & _
           " INNER JOIN RHContratoDet CD ON CON.cPersCod=CD.cPersCod AND CON.cRHContratoNro=CD.cRHContratoNro AND " & _
           " CD.dRHContratoFecha=(SELECT MAX(dRHContratoFecha) FROM RHContratoDet WHERE cPersCod=CON.cPersCod AND " & _
           " cRHContratoNro=CON.cRHContratoNro)INNER JOIN RHSueldo S ON S.cPersCod=P.cPersCod AND S.dRHSueldoFecha= " & _
           " (SELECT MAX(dRHSueldoFecha) FROM RHSueldo WHERE cPersCod=P.cPersCod) INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND " & _
           " CAR.dRHCargoFecha= (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) " & _
           " INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod INNER JOIN RHSeleccionPostulante SP ON P.cPersCod = SP.cPersCod " & _
           " AND P.cPersCod NOT IN( SELECT cPersCod FROM RHContratoDet_Impresion WHERE cRHContratoNro=CON.cRHContratoNro)" & _
           " WHERE SP.cCodProceso='" & pcCodProceso & "'"
     
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Personal_Ingresante = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Personal_Ingresante:
    Err.Raise Err.Number, "Obtener_Personal_Ingresante", Err.Description
End Function

Public Sub InsertaRHCargo_Funcion(ByVal pcRHCargoCod As String, _
                                    ByVal pbFuncionBasica As Integer, _
                                    ByVal pcFuncionDescripcion As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHCargo_Funcion
    sSQL = "INSERT INTO RHCargo_Funcion(cRHCargoCod,bFuncionBasica,cFuncionDescripcion)"
    sSQL = sSQL & " VALUES('" & pcRHCargoCod & "'," & pbFuncionBasica & ",'" & pcFuncionDescripcion & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHCargo_Funcion:
    Err.Raise Err.Number, "InsertaRHCargo_Funcion", Err.Description
End Sub

Public Sub ModificaRHCargo_Funcion(ByVal pnFuncionCod As Integer, _
                                    ByVal pcRHCargoCod As String, _
                                    ByVal pbFuncionBasica As Integer, _
                                    ByVal pcFuncionDescripcion As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHCargo_Funcion
    sSQL = "UPDATE RHCargo_Funcion" '
    sSQL = sSQL & " SET cRHCargoCod='" & pcRHCargoCod & _
                    "',bFuncionBasica=" & pbFuncionBasica & _
                    ",cFuncionDescripcion='" & pcFuncionDescripcion & _
                    "' WHERE nFuncionCod=" & pnFuncionCod
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHCargo_Funcion:
    Err.Raise Err.Number, "ModificaRHCargo_Funcion", Err.Description
End Sub

Public Function EliminaRHCargo_Funcion(ByVal pnFuncionCod As Integer)

Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHCargo_Funcion
    sSQL = "DELETE FROM RHCargo_Funcion"
    sSQL = sSQL & " WHERE nFuncionCod=" & pnFuncionCod
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHCargo_Funcion:
    Err.Raise Err.Number, "EliminaRHCargo_Funcion", Err.Description
End Function

Public Sub InsertaRHCargo_SubFuncion(ByVal pnFuncionCod As Integer, _
                                    ByVal pcSubFuncionDescripcion As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHCargo_SubFuncion
    sSQL = "INSERT INTO RHCargo_SubFuncion(nFuncionCod,cSubFuncionDescripcion)"
    sSQL = sSQL & " VALUES(" & pnFuncionCod & ",'" & pcSubFuncionDescripcion & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHCargo_SubFuncion:
    Err.Raise Err.Number, "InsertaRHCargo_SubFuncion", Err.Description
End Sub

Public Sub ModificaRHCargo_SubFuncion(ByVal pnSubFuncionCod As Integer, _
                                    ByVal pcSubFuncionDescripcion As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHCargo_SubFuncion
    sSQL = "UPDATE RHCargo_SubFuncion" '
    sSQL = sSQL & " SET cSubFuncionDescripcion='" & pcSubFuncionDescripcion & _
                    "' WHERE nSubFuncionCod=" & pnSubFuncionCod
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHCargo_SubFuncion:
    Err.Raise Err.Number, "ModificaRHCargo_SubFuncion", Err.Description
End Sub

Public Function EliminaRHCargo_SubFuncion(ByVal pnSubFuncionCod As Integer)

Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHCargo_SubFuncion
    sSQL = "DELETE FROM RHCargo_SubFuncion"
    sSQL = sSQL & " WHERE nSubFuncionCod=" & pnSubFuncionCod
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHCargo_SubFuncion:
    Err.Raise Err.Number, "EliminaRHCargo_SubFuncion", Err.Description
End Function

Public Function Obtener_Funciones_x_Cargo(ByVal pcRHCargoCod As String, Optional ByVal pbFuncionBasica As Integer = -1) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Funciones_x_Cargo
    
    sSQL = "SELECT nFuncionCod,cFuncionDescripcion,bFuncionBasica=CONVERT(int,bFuncionBasica) FROM RHCargo_Funcion WHERE cRHCargoCod='" & pcRHCargoCod & "'" & _
            IIf(pbFuncionBasica >= 0, " AND bFuncionBasica=" & pbFuncionBasica, "")
     
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Funciones_x_Cargo = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Funciones_x_Cargo:
    Err.Raise Err.Number, "Obtener_Funciones_x_Cargo", Err.Description
End Function

Public Function Obtener_SubFunciones_x_Funcion(ByVal pnFuncionCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_SubFunciones_x_Funcion
    
    sSQL = "SELECT nSubFuncionCod,cSubFuncionDescripcion FROM RHCargo_SubFuncion WHERE nFuncionCod=" & pnFuncionCod
     
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_SubFunciones_x_Funcion = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_SubFunciones_x_Funcion:
    Err.Raise Err.Number, "Obtener_SubFunciones_x_Funcion", Err.Description
End Function

Public Function Obtener_Personal_ProcesosSeleccion() As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Personal_ProcesosSeleccion
    
    sSQL = "SELECT DISTINCT P.cPersCod,P.cPersNombre,CT.cRHCargoDescripcion,CT.cRHCargoCod FROM Persona P INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND " & _
           " CAR.dRHCargoFecha= (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) INNER JOIN RHCargosTabla CT ON " & _
           " CAR.cRHCargoCod=CT.cRHCargoCod INNER JOIN RHSeleccionPostulante SP ON P.cPersCod = SP.cPersCod ORDER BY P.cPersNombre"
     
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Personal_ProcesosSeleccion = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Personal_ProcesosSeleccion:
    Err.Raise Err.Number, "Obtener_Personal_ProcesosSeleccion", Err.Description
End Function

Public Function Obtener_Contratos_x_Empleado(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Contratos_x_Empleado
    
    sSQL = "SELECT P.cPersCod,CD.cRHContratoNro,dRHContratoFecha=CONVERT(varchar(10),CD.dRHContratoFecha,103)," & _
           " dRHContratoInicio=CONVERT(varchar(10),CD.dRHContratoInicio,103),dRHContratoFin = CONVERT(VarChar(10), CD.dRHContratoFin, 103),CI.cPersCodResp1,CI.cPersCodResp2,CI.cNumero" & _
           " FROM Persona P INNER JOIN RHContrato CON ON P.cPersCod=CON.cPersCod " & _
           " INNER JOIN RHContratoDet CD ON CON.cPersCod=CD.cPersCod AND CON.cRHContratoNro=CD.cRHContratoNro " & _
           " INNER JOIN RHContratoDet_Impresion CI ON CD.cPersCod=CI.cPersCod AND CD.cRHContratoNro=CI.cRHContratoNro " & _
           " AND CD.dRHContratoFecha=CI.dRHContratoFecha " & _
           " WHERE P.cPerscod='" & pcPersCod & "' ORDER BY CD.dRHContratoFecha"
     
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Contratos_x_Empleado = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Contratos_x_Empleado:
    Err.Raise Err.Number, "Obtener_Contratos_x_Empleado", Err.Description
End Function

Public Sub InsertaRHContratoDet_Impresion(ByVal pcPersCod As String, _
                                    ByVal pcRHContratoNro As String, _
                                    ByVal pdRHContratoFecha As Date, _
                                    ByVal pcPersCodResp1 As String, _
                                    ByVal pcPersCodResp2 As String, _
                                    ByVal pcNumero As String)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHContratoDet_Impresion
    sSQL = "INSERT INTO RHContratoDet_Impresion(cPersCod,cRHContratoNro,dRHContratoFecha,cPersCodResp1,cPersCodResp2,cNumero)"
    sSQL = sSQL & " VALUES('" & pcPersCod & "','" & pcRHContratoNro & "','" & Format(pdRHContratoFecha, gsFormatoFecha) & "','" & pcPersCodResp1 & "','" & pcPersCodResp2 & "','" & pcNumero & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHContratoDet_Impresion:
    Err.Raise Err.Number, "InsertaRHContratoDet_Impresion", Err.Description
End Sub

Public Function Obtener_Numero_Impresion() As String

Dim sSQL As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrorObtener_Numero_Impresion
    sSQL = "SELECT  Codigo= CONVERT(varchar(4), ISNULL(MAX(CONVERT(int,SUBSTRING(cNumero,1,PATINDEX('%-%',cNumero)-1))),0)+1 )" & _
           " +'-'+CONVERT(varchar(4),YEAR(getdate()))FROM RHContratoDet_Impresion WHERE Right(cNumero, 4) = Year(getdate())"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSQL)
        If Not rs.EOF Then
            Obtener_Numero_Impresion = CStr(rs("Codigo"))
        End If
    End If
    
    oConecta.CierraConexion
    
    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Numero_Impresion:
    Err.Raise Err.Number, "Obtener_Numero_Impresion", Err.Description
End Function

Public Function Obtener_Datos_Para_ImpresionContrato(ByVal pcPersCod As String, _
                                                    Optional ByVal pcRHContratoNro As String = "", _
                                                    Optional ByVal pdRHContratoFecha As Date = 0) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_Para_ImpresionContrato
    
    If pcRHContratoNro = "" Then
        sSQL = "SELECT Empleado=P.cPersNombre,Documento=PID.cPersIDnro,Domicilio=P.cPersDireccDomicilio+','+ RTRIM(Z.cDesZon),Trabajador=CASE WHEN PN.cPersNatSexo='M' THEN 'EL TRABAJADOR' ELSE 'LA TRABAJADORA'END," & _
           " Cargo=CT.cRHCargoDescripcion,Inicio=CONVERT(varchar(2),DAY(CD.dRHContratoInicio))+' de '+ RTRIM(M.cNomTab)+ ' del ' + CONVERT(varchar(4),YEAR(CD.dRHContratoInicio)),Fin=CONVERT(varchar(2),DAY(CD.dRHContratoFin))+' de '+ RTRIM(M.cNomTab)+ ' del ' + CONVERT(varchar(4),YEAR(CD.dRHContratoFin)), Sueldo = s.nRHSueldoMonto ,Fecha= CONVERT(varchar(2),DAY(getdate()))+' de '+ RTRIM(M.cNomTab)+ ' del ' + CONVERT(varchar(4),YEAR(getdate())) " & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod INNER JOIN RHContrato CON ON P.cPersCod=CON.cPersCod AND CON.cRHContratoNro=(SELECT MAX(CONVERT(int,cRHContratoNro))FROM RHContrato WHERE cPersCod=P.cPersCod)" & _
           " INNER JOIN RHContratoDet CD ON CON.cPersCod=CD.cPersCod AND CON.cRHContratoNro=CD.cRHContratoNro AND CD.dRHContratoFecha=(SELECT MAX(dRHContratoFecha) FROM RHContratoDet WHERE cPersCod=CON.cPersCod AND " & _
           " cRHContratoNro=CON.cRHContratoNro)INNER JOIN RHSueldo S ON S.cPersCod=P.cPersCod AND S.dRHSueldoFecha=(SELECT MAX(dRHSueldoFecha) FROM RHSueldo WHERE cPersCod=P.cPersCod)" & _
           " INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND CAR.dRHCargoFecha= (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod)INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod " & _
           " INNER JOIN PersID PID ON P.cPersCod = PID.cPersCod And PID.cPersIDTpo = 1 INNER JOIN PersonaNat PN ON P.cPersCod=PN.cPersCod INNER JOIN DBComunes..Zonas Z ON Z.cCodZon=P.cPersDireccUbiGeo" & _
           " INNER JOIN DBComunes..TablaCod M ON LEFT(M.cCodTab,2)='EZ' AND MONTH(getdate()) = CONVERT(int,M.cValor)" & _
           " INNER JOIN DBComunes..TablaCod I ON LEFT(I.cCodTab,2)='EZ' AND MONTH(CD.dRHContratoInicio) = CONVERT(int,I.cValor)" & _
           " INNER JOIN DBComunes..TablaCod F ON LEFT(F.cCodTab,2)='EZ' AND MONTH(CD.dRHContratoFin) = CONVERT(int,F.cValor)" & _
           " WHERE P.cPersCod='" & pcPersCod & "'"
    Else
        sSQL = "SELECT Empleado=P.cPersNombre,Documento=PID.cPersIDnro,Domicilio=P.cPersDireccDomicilio+','+ Z.cDesZon,Trabajador=CASE WHEN PN.cPersNatSexo='M' THEN 'EL TRABAJADOR' ELSE 'LA TRABAJADORA'END," & _
           " Cargo=CT.cRHCargoDescripcion,Inicio=CONVERT(varchar(2),DAY(CD.dRHContratoInicio))+' de '+ RTRIM(M.cNomTab)+ ' del ' + CONVERT(varchar(4),YEAR(CD.dRHContratoInicio)),Fin=CONVERT(varchar(2),DAY(CD.dRHContratoFin))+' de '+ RTRIM(M.cNomTab)+ ' del ' + CONVERT(varchar(4),YEAR(CD.dRHContratoFin)), Sueldo = s.nRHSueldoMonto ,Fecha= CONVERT(varchar(2),DAY(getdate()))+' de '+ RTRIM(M.cNomTab)+ ' del ' + CONVERT(varchar(4),YEAR(getdate())) " & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod INNER JOIN RHContrato CON ON P.cPersCod=CON.cPersCod AND CON.cRHContratoNro='" & pcRHContratoNro & "'" & _
           " INNER JOIN RHContratoDet CD ON CON.cPersCod=CD.cPersCod AND CON.cRHContratoNro=CD.cRHContratoNro AND CONVERT(varchar(10),CD.dRHContratoFecha,103)= CONVERT(varchar(10),'" & pdRHContratoFecha & "',103) INNER JOIN RHSueldo S ON S.cPersCod=P.cPersCod AND S.dRHSueldoFecha= CD.dRHContratoFecha " & _
           " INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND CAR.dRHCargoFecha= (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod)INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod " & _
           " INNER JOIN PersID PID ON P.cPersCod = PID.cPersCod And PID.cPersIDTpo = 1 INNER JOIN PersonaNat PN ON P.cPersCod=PN.cPersCod INNER JOIN DBComunes..Zonas Z ON Z.cCodZon=P.cPersDireccUbiGeo" & _
           " INNER JOIN DBComunes..TablaCod M ON LEFT(M.cCodTab,2)='EZ' AND MONTH(getdate()) = CONVERT(int,M.cValor)" & _
           " INNER JOIN DBComunes..TablaCod I ON LEFT(I.cCodTab,2)='EZ' AND MONTH(CD.dRHContratoInicio) = CONVERT(int,I.cValor)" & _
           " INNER JOIN DBComunes..TablaCod F ON LEFT(F.cCodTab,2)='EZ' AND MONTH(CD.dRHContratoFin) = CONVERT(int,F.cValor)" & _
           " WHERE P.cPersCod='" & pcPersCod & "'"
    End If
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_Para_ImpresionContrato = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Datos_Para_ImpresionContrato:
    Err.Raise Err.Number, "Obtener_Datos_Para_ImpresionContrato", Err.Description
End Function

Public Function Obtener_Datos_Representante_Contrato(ByVal pcPersCod As String) As String

Dim sSQL As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
Dim cNomPersona As String

    On Error GoTo ErrorObtener_Datos_Representante_Contrato
    
    sSQL = " SELECT Cargo= CT.cRHCargoDescripcion,Empleado= P.cPersNombre,Documento=PID.cPersIDnro " & _
           " FROM Persona P INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND " & _
           " CAR.dRHCargoFecha= (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) " & _
           " INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod INNER JOIN PersID PID ON " & _
           " P.cPersCod = PID.cPersCod And PID.cPersIDTpo = 1 " & _
           " WHERE P.cPersCod='" & pcPersCod & "'"
     
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSQL)
        If Not rs.EOF Then
            Obtener_Datos_Representante_Contrato = rs("Cargo") & ", " & PstaNombre(rs("Empleado"), True) & ", con DNI Nº " & rs("Documento")
        Else
            Obtener_Datos_Representante_Contrato = ""
        End If
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Datos_Representante_Contrato:
    Err.Raise Err.Number, "Obtener_Datos_Representante_Contrato", Err.Description
End Function

Public Function Obtener_Documentos_NoEntregados(ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Documentos_NoEntregados
    sSQL = "sp_RHObtener_Documentos_NoEntregados " & pnTipo
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Documentos_NoEntregados = oConecta.Ejecutar(sSQL)
        
    End If
    'oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Documentos_NoEntregados:
    Err.Raise Err.Number, "Obtener_Documentos_NoEntregados", Err.Description
End Function
