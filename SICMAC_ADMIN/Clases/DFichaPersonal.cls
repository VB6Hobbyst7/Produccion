VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DFichaPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InsertaRHEmpleadoDatosComplementarios(ByVal pcPersCod As String, _
                    ByVal pcDomicilioLegal As String, _
                    ByVal pcDomicilioActual As String, _
                    ByVal pnTipoVivienda As Integer, _
                    ByVal pnCostoAlquiler As Double, _
                    ByVal pnTipoConvivencia As Integer, _
                    ByVal pcOtrosConvive As String, _
                    ByVal pcLugarNacimiento As String, _
                    ByVal pnCategoriaBrevete As Integer, _
                    ByVal pcNombreEmergencia As String, _
                    ByVal pcParentescoEmergencia As String, _
                    ByVal pcTelefonoEmergencia As String, _
                    ByVal pbSeguroPrivado As Integer, _
                    ByVal pnNumeroDependientes As Integer, _
                    Optional ByVal pcUltimaActualizacion As String = "")

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEmpleadoDatosComplementarios
    sSql = "INSERT INTO RHEmpleadoDatosComplementarios(cPersCod, cDomicilioLegal,cDomicilioActual, nTipoVivienda, nCostoAlquiler, nTipoConvivencia, cOtrosConvive, cLugarNacimiento, nCategoriaBrevete, cNombreEmergencia, cParentescoEmergencia, cTelefonoEmergencia,bSeguroPrivado,nNumeroDependientes,cUltimaActualizacion)"
    sSql = sSql & " VALUES('" & pcPersCod & "','" & pcDomicilioLegal & "','" & pcDomicilioActual & "'," & pnTipoVivienda & "," & pnCostoAlquiler & "," & pnTipoConvivencia & ",'" & pcOtrosConvive & "','" & pcLugarNacimiento & "'," & pnCategoriaBrevete & ",'" & pcNombreEmergencia & "','" & pcParentescoEmergencia & "','" & pcTelefonoEmergencia & "'," & pbSeguroPrivado & "," & pnNumeroDependientes & ",'" & pcUltimaActualizacion & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEmpleadoDatosComplementarios:
    Err.Raise Err.Number, "InsertaRHEmpleadoDatosComplementarios", Err.Description
End Sub

Public Sub ModificaRHEmpleadoDatosComplementarios(ByVal pcPersCod As String, _
                    ByVal pcDomicilioLegal As String, _
                    ByVal pcDomicilioActual As String, _
                    ByVal pnTipoVivienda As Integer, _
                    ByVal pnCostoAlquiler As Double, _
                    ByVal pnTipoConvivencia As Integer, _
                    ByVal pcOtrosConvive As String, _
                    ByVal pcLugarNacimiento As String, _
                    ByVal pnCategoriaBrevete As Integer, _
                    ByVal pcNombreEmergencia As String, _
                    ByVal pcParentescoEmergencia As String, _
                    ByVal pcTelefonoEmergencia As String, _
                    ByVal pbSeguroPrivado As Integer, _
                    ByVal pnNumeroDependientes As Integer, _
                    Optional ByVal pcUltimaActualizacion As String = "")

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHEmpleadoDatosComplementarios
    sSql = "UPDATE RHEmpleadoDatosComplementarios" '
    sSql = sSql & " SET cDomicilioLegal='" & pcDomicilioLegal & _
                    "',cDomicilioActual='" & pcDomicilioActual & _
                    "',nTipoVivienda =" & pnTipoVivienda & _
                    ",cOtrosConvive='" & pcOtrosConvive & _
                    "',cLugarNacimiento='" & pcLugarNacimiento & _
                    "',nCategoriaBrevete=" & pnCategoriaBrevete & _
                    ",cNombreEmergencia='" & pcNombreEmergencia & _
                    "',cParentescoEmergencia='" & pcParentescoEmergencia & _
                    "',cTelefonoEmergencia='" & pcTelefonoEmergencia & _
                    "',bSeguroPrivado=" & pbSeguroPrivado & _
                    ",nNumeroDependientes=" & pnNumeroDependientes & _
                    ",cUltimaActualizacion='" & pcUltimaActualizacion & _
                    "' WHERE cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEmpleadoDatosComplementarios:
    Err.Raise Err.Number, "ModificaRHEmpleadoDatosComplementarios", Err.Description
End Sub

Public Sub InsertaRHEmpleadoDatosComp_General(ByVal pcPersCod As String, _
                    ByVal pcDomicilioLegal As String, _
                    ByVal pcDomicilioActual As String, _
                    ByVal pnTipoVivienda As Integer, _
                    ByVal pnCostoAlquiler As Double, _
                    ByVal pnTipoConvivencia As Integer, _
                    ByVal pcOtrosConvive As String, _
                    ByVal pcLugarNacimiento As String, _
                    ByVal pnCategoriaBrevete As Integer, _
                    ByVal pbSeguroPrivado As Integer, _
                    ByVal pnNumeroDependientes As Integer, _
                    ByVal pnLugarVivienda As Integer, _
                    ByVal pcUbigeoDomLegal As String, _
                    ByVal pcUbigeoDomActual As String, _
                    Optional ByVal pcUltimaActualizacion As String = "")

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEmpleadoDatosComp_General
    sSql = "INSERT INTO RHEmpleadoDatosComplementarios(cPersCod, cDomicilioLegal,cDomicilioActual, nTipoVivienda, nCostoAlquiler, nTipoConvivencia, cOtrosConvive, cLugarNacimiento, nCategoriaBrevete, bSeguroPrivado,nNumeroDependientes,nLugarVivienda,cUbigeoDomLegal,cUbigeoDomActual,cUltimaActualizacion)"
    sSql = sSql & " VALUES('" & pcPersCod & "','" & pcDomicilioLegal & "','" & pcDomicilioActual & "'," & pnTipoVivienda & "," & pnCostoAlquiler & "," & pnTipoConvivencia & ",'" & pcOtrosConvive & "','" & pcLugarNacimiento & "'," & pnCategoriaBrevete & "," & pbSeguroPrivado & "," & pnNumeroDependientes & "," & pnLugarVivienda & ",'" & pcUbigeoDomLegal & "','" & pcUbigeoDomActual & "','" & pcUltimaActualizacion & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEmpleadoDatosComp_General:
    Err.Raise Err.Number, "InsertaRHEmpleadoDatosComp_General", Err.Description
End Sub

Public Sub InsertaRHEmpleadoDatosComp_Salud(ByVal pcPersCod As String, _
                    ByVal pcNombreEmergencia As String, _
                    ByVal pcParentescoEmergencia As String, _
                    ByVal pcTelefonoEmergencia As String, _
                    Optional ByVal pcUltimaActualizacion As String = "")

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEmpleadoDatosComp_Salud
    sSql = "INSERT INTO RHEmpleadoDatosComplementarios(cPersCod,cNombreEmergencia, cParentescoEmergencia, cTelefonoEmergencia,cUltimaActualizacion)"
    sSql = sSql & " VALUES('" & pcPersCod & "','" & pcNombreEmergencia & "','" & pcParentescoEmergencia & "','" & pcTelefonoEmergencia & "','" & pcUltimaActualizacion & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEmpleadoDatosComp_Salud:
    Err.Raise Err.Number, "InsertaRHEmpleadoDatosComp_Salud", Err.Description
End Sub

Public Sub ModificaRHEmpleadoDatosComp_General(ByVal pcPersCod As String, _
                    ByVal pcDomicilioLegal As String, _
                    ByVal pcDomicilioActual As String, _
                    ByVal pnTipoVivienda As Integer, _
                    ByVal pnCostoAlquiler As Double, _
                    ByVal pnTipoConvivencia As Integer, _
                    ByVal pcOtrosConvive As String, _
                    ByVal pcLugarNacimiento As String, _
                    ByVal pnCategoriaBrevete As Integer, _
                    ByVal pbSeguroPrivado As Integer, _
                    ByVal pnNumeroDependientes As Integer, _
                    ByVal pnLugarVivienda As Integer, _
                    ByVal pcUbigeoDomLegal As String, _
                    ByVal pcUbigeoDomActual As String, _
                    Optional ByVal pcUltimaActualizacion As String = "")

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHEmpleadoDatosComp_General
    sSql = "UPDATE RHEmpleadoDatosComplementarios" '
    sSql = sSql & " SET cDomicilioLegal='" & pcDomicilioLegal & _
                    "',cDomicilioActual='" & pcDomicilioActual & _
                    "',nTipoVivienda =" & pnTipoVivienda & _
                    ",cOtrosConvive='" & pcOtrosConvive & _
                    "',cLugarNacimiento='" & pcLugarNacimiento & _
                    "',nCategoriaBrevete=" & pnCategoriaBrevete & _
                    ",bSeguroPrivado=" & pbSeguroPrivado & _
                    ",nNumeroDependientes=" & pnNumeroDependientes & _
                    ",nLugarVivienda=" & pnLugarVivienda & _
                    ",cUbigeoDomActual='" & pcUbigeoDomActual & _
                    "',cUbigeoDomLegal='" & pcUbigeoDomLegal & _
                    "',cUltimaActualizacion='" & pcUltimaActualizacion & _
                    "' WHERE cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEmpleadoDatosComp_General:
    Err.Raise Err.Number, "ModificaRHEmpleadoDatosComp_General", Err.Description
End Sub

Public Sub ModificaRHEmpleadoDatosComp_Salud(ByVal pcPersCod As String, _
                    ByVal pcNombreEmergencia As String, _
                    ByVal pcParentescoEmergencia As String, _
                    ByVal pcTelefonoEmergencia As String, _
                    Optional ByVal pcUltimaActualizacion As String = "")
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHEmpleadoDatosComp_Salud
    sSql = "UPDATE RHEmpleadoDatosComplementarios" '
    sSql = sSql & " SET cNombreEmergencia='" & pcNombreEmergencia & _
                    "',cParentescoEmergencia='" & pcParentescoEmergencia & _
                    "',cTelefonoEmergencia='" & pcTelefonoEmergencia & _
                    "',cUltimaActualizacion='" & pcUltimaActualizacion & _
                    "' WHERE cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEmpleadoDatosComp_Salud:
    Err.Raise Err.Number, "ModificaRHEmpleadoDatosComp_Salud", Err.Description
End Sub


Public Function EliminaRHEmpleadoDatosComplementarios(ByVal pcPersCod As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEmpleadoDatosComplementarios
    sSql = "DELETE FROM RHEmpleadoDatosComplementarios"
    sSql = sSql & " WHERE cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEmpleadoDatosComplementarios:
    Err.Raise Err.Number, "EliminaRHEmpleadoDatosComplementarios", Err.Description
End Function

Public Sub InsertaRHEmpleadoSaludEnfermedad(ByVal pcPersCod As String, _
                                    ByVal pnItem As Integer, _
                                    ByVal pnTipoSaludEnfermedad As Integer, _
                                    ByVal pcDescripcion As String, _
                                    ByVal pcTratamiento As String, _
                                    ByVal pcMedico As String, _
                                    ByVal pnEstado As Integer, _
                    Optional ByVal pcUltimaActualizacion As String = "")

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEmpleadoSaludEnfermedad
    sSql = "INSERT INTO RHEmpleadoSaludEnfermedad(cPersCod, nItem, nTipoSaludEnfermedad, cDescripcion, cTratamiento, cMedico, nEstado,cUltimaActualizacion)"
    sSql = sSql & " VALUES('" & pcPersCod & "'," & pnItem & "," & pnTipoSaludEnfermedad & ",'" & pcDescripcion & "','" & pcTratamiento & "','" & pcMedico & "'," & pnEstado & ",'" & pcUltimaActualizacion & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEmpleadoSaludEnfermedad:
    Err.Raise Err.Number, "InsertaRHEmpleadoSaludEnfermedad", Err.Description
End Sub

Public Function EliminaRHEmpleadoSaludEnfermedad(ByVal pcPersCod As String, _
                                                ByVal pnItem As Integer)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEmpleadoSaludEnfermedad
    sSql = "DELETE FROM RHEmpleadoSaludEnfermedad"
    sSql = sSql & " WHERE cPersCod='" & pcPersCod & "' AND nItem=" & pnItem
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEmpleadoSaludEnfermedad:
    Err.Raise Err.Number, "EliminaRHEmpleadoSaludEnfermedad", Err.Description
End Function

Public Sub InsertaRHEmpleadoHabilidades(ByVal pcPersCod As String, _
                                    ByVal pnItem As String, _
                                    ByVal pnTipoHabilidad As Integer, _
                                    ByVal pnNivelHabilidad As Integer, _
                                    ByVal pcComentario As String, _
                    Optional ByVal pcUltimaActualizacion As String = "")

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEmpleadoHabilidades
    sSql = "INSERT INTO RHEmpleadoHabilidades(cPersCod, nItem, nTipoHabilidad, nNivelHabilidad, cComentario,cUltimaActualizacion)"
    sSql = sSql & " VALUES('" & pcPersCod & "'," & pnItem & "," & pnTipoHabilidad & "," & pnNivelHabilidad & ",'" & pcComentario & "','" & pcUltimaActualizacion & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEmpleadoHabilidades:
    Err.Raise Err.Number, "InsertaRHEmpleadoHabilidades", Err.Description
End Sub

Public Function EliminaRHEmpleadoHabilidades(ByVal pcPersCod As String, _
                                                ByVal pnItem As Integer)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEmpleadoHabilidades
    sSql = "DELETE FROM RHEmpleadoHabilidades"
    sSql = sSql & " WHERE cPersCod='" & pcPersCod & "' AND nItem=" & pnItem
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEmpleadoHabilidades:
    Err.Raise Err.Number, "EliminaRHEmpleadoHabilidades", Err.Description
End Function

Public Sub InsertaRHEmpleadoCurriculumExperiencia(ByVal pcPersCod As String, _
                    ByVal pnItem As Integer, _
                    ByVal pnTipoExperiencia As Integer, _
                    ByVal pcLugar As String, _
                    ByVal pdInicio As Date, _
                    ByVal pdFin As Date, _
                    ByVal pcCargo As String, _
                    ByVal pcJefeSuperior As String, _
                    ByVal pnSueldo As Double, _
                    ByVal pcMotivoSalida As String, _
                    ByVal pcObservacion As String, _
                    Optional ByVal pcUltimaActualizacion As String = "")

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEmpleadoCurriculumExperiencia
    sSql = "INSERT INTO RHEmpleadoCurriculumExperiencia(cPersCod, nItem, nTipoExperiencia, cLugar, dInicio, dFin, cCargo, cJefeSuperior, nSueldo, cMotivoSalida, cObservacion,cUltimaActualizacion)"
    sSql = sSql & " VALUES('" & pcPersCod & "'," & pnItem & "," & pnTipoExperiencia & ",'" & pcLugar & "','" & Format(pdInicio, gsFormatoFecha) & "','" & Format(pdFin, gsFormatoFecha) & "','" & pcCargo & "','" & pcJefeSuperior & "'," & pnSueldo & ",'" & pcMotivoSalida & "','" & pcObservacion & "','" & pcUltimaActualizacion & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEmpleadoCurriculumExperiencia:
    Err.Raise Err.Number, "InsertaRHEmpleadoCurriculumExperiencia", Err.Description
End Sub


Public Function EliminaRHEmpleadoCurriculumExperiencia(ByVal pcPersCod As String, _
                                                       ByVal pnItem As Integer)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEmpleadoCurriculumExperiencia
    sSql = "DELETE FROM RHEmpleadoCurriculumExperiencia"
    sSql = sSql & " WHERE cPersCod='" & pcPersCod & "' AND nItem=" & pnItem
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEmpleadoCurriculumExperiencia:
    Err.Raise Err.Number, "EliminaRHEmpleadoCurriculumExperiencia", Err.Description
End Function

Public Sub InsertaRHEmpleadoCurriculum(ByVal pcPersCod As String, _
                    ByVal pnItem As Integer, _
                    ByVal pnTipoCurriculum As Integer, _
                    ByVal pnCurriculumCod As Integer, _
                    ByVal pcLugar As String, _
                    ByVal pdInicio As Date, _
                    ByVal pdFin As Date, _
                    ByVal pcComentario As String, _
                    ByVal pnProfesionCod As Integer, _
                    ByVal pnCosto As Double, _
                    ByVal pnNivel As Integer, _
                    ByVal pnGradoObtenido As Integer, _
                    ByVal pbCmactOtorgado As Integer, _
                    Optional ByVal pcUltimaActualizacion As String = "")

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEmpleadoCurriculum
    sSql = "INSERT INTO RHEmpleadoCurriculum(cPersCod, nItem,nTipoCurriculum ,nCurriculumCod, cLugar, dInicio, dFin, cComentario, nProfesionCod, nCosto, nNivel, nGradoObtenido, bCmactOtorgado,cUltimaActualizacion)"
    sSql = sSql & " VALUES('" & pcPersCod & "'," & pnItem & "," & pnTipoCurriculum & ",'" & pnCurriculumCod & "','" & pcLugar & "','" & Format(pdInicio, gsFormatoFecha) & "','" & Format(pdFin, gsFormatoFecha) & "','" & pcComentario & "'," & pnProfesionCod & "," & pnCosto & "," & pnNivel & "," & pnGradoObtenido & "," & pbCmactOtorgado & ",'" & pcUltimaActualizacion & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEmpleadoCurriculum:
    Err.Raise Err.Number, "InsertaRHEmpleadoCurriculum", Err.Description
End Sub

Public Function EliminaRHEmpleadoCurriculum(ByVal pcPersCod As String, _
                                            ByVal pnItem As Integer, _
                                            ByVal pnTipoCurriculum As Integer)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEmpleadoCurriculum
    sSql = "DELETE FROM RHEmpleadoCurriculum"
    sSql = sSql & " WHERE cPersCod='" & pcPersCod & "' AND nItem=" & pnItem & " AND nTipoCurriculum=" & pnTipoCurriculum
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEmpleadoCurriculum:
    Err.Raise Err.Number, "EliminaRHEmpleadoCurriculum", Err.Description
End Function

Public Function Obtener_Empleado_SaludEnfermedad(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Empleado_SaludEnfermedad
    
    sSql = "SELECT nItem, nTipoSaludEnfermedad,T.cConsDescripcion, cDescripcion, cTratamiento, cMedico, nEstado,E.cConsDescripcion" & _
           " FROM RHEmpleadoSaludEnfermedad SE INNER JOIN Constante T ON T.nConsValor=nTipoSaludEnfermedad AND T.nConsCod=9056 INNER JOIN Constante E " & _
           " ON E.nConsValor=nEstado AND E.nConsCod=9057 " & _
           " WHERE cPersCod='" & pcPersCod & "'"

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Empleado_SaludEnfermedad = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Empleado_SaludEnfermedad:
    Err.Raise Err.Number, "Obtener_Empleado_SaludEnfermedad", Err.Description
End Function

Public Function Obtener_Empleado_Habilidades(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Empleado_Habilidades
    
    sSql = "SELECT nItem, nTipoHabilidad,T.cConsDescripcion, nNivelHabilidad,N.cConsDescripcion, cComentario" & _
           " FROM RHEmpleadoHabilidades SH INNER JOIN Constante T ON  nTipoHabilidad=T.nConsValor AND T.nConsCod=9095 " & _
           " INNER JOIN Constante N ON nNivelHabilidad=N.nConsValor AND N.nConsCod=6048 " & _
           " WHERE cPersCod='" & pcPersCod & "'"

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Empleado_Habilidades = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Empleado_Habilidades:
    Err.Raise Err.Number, "Obtener_Empleado_Habilidades", Err.Description
End Function

Public Function Obtener_Empleado_CurriculumExperiencia(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Empleado_CurriculumExperiencia
    
    sSql = "SELECT nItem, nTipoExperiencia,T.cConsDescripcion, cLugar,dInicio=CONVERT(varchar(10),dInicio,103),dFin=CONVERT(varchar(10),dFin,103), cCargo, cJefeSuperior, nSueldo, cMotivoSalida, cObservacion" & _
           " FROM RHEmpleadoCurriculumExperiencia CE INNER JOIN Constante T ON T.nConsValor=nTipoExperiencia AND T.nConsCod=9055 " & _
           " WHERE cPersCod='" & pcPersCod & "'"

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Empleado_CurriculumExperiencia = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Empleado_CurriculumExperiencia:
    Err.Raise Err.Number, "Obtener_Empleado_CurriculumExperiencia", Err.Description
End Function

Public Function Obtener_Empleado_Curriculum(ByVal pcPersCod As String, _
                                        ByVal pnTipoCurriculum As Integer) As ADODB.Recordset

'1 Formacion Academica, 2 Actualizacion
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Empleado_Curriculum
    
    If pnTipoCurriculum = 1 Then
        sSql = "SELECT nItem, nCurriculumCod,C.cConsDescripcion,dInicio=CONVERT(varchar(10),dInicio,103), dFin=CONVERT(varchar(10),dFin,103), nNivel,N.cConsDescripcion, nGradoObtenido, " & _
               " G.cConsDescripcion, nProfesionCod, P.cConsDescripcion, cLugar, cComentario,nCosto, bCmactOtorgado " & _
               " FROM RHEmpleadoCurriculum EC INNER JOIN Constante C ON nCurriculumCod=C.nConsValor AND C.nConsCod=9051 " & _
               " INNER JOIN Constante P ON nProfesionCod=P.nConsValor AND P.nConsCod=6044 INNER JOIN Constante N ON nNivel=N.nConsValor AND N.nConsCod=6045 " & _
               " INNER JOIN Constante G ON nGradoObtenido=G.nConsValor AND G.nConsCod=9052 " & _
               " WHERE cPersCod='" & pcPersCod & "' AND nTipoCurriculum=" & pnTipoCurriculum
    Else
        sSql = "Select nItem,nCurriculumCod, C.cConsDescripcion, nProfesionCod, " & _
               " Case nCurriculumCod   When 0 Then (Select cConsDescripcion From Constante COI Where COI.nConsCod = 6047 And COI.nConsValor = nProfesionCod) " & _
               " When 1 Then (Select cConsDescripcion From Constante COI Where COI.nConsCod = 6048 And COI.nConsValor = nProfesionCod)" & _
               " When 2 Then (Select cConsDescripcion From Constante COI Where COI.nConsCod = 6049 And COI.nConsValor = nProfesionCod)" & _
               " When 3 Then (Select cConsDescripcion From Constante COI Where COI.nConsCod = 6050 And COI.nConsValor = nProfesionCod) End Especialidad, " & _
               " nNivel , N.cConsDescripcion From RHEmpleadoCurriculum INNER JOIN Constante C ON nCurriculumCod=C.nConsValor AND C.nConsCod=6046 " & _
               " INNER JOIN Constante N ON nNivel=N.nConsValor AND N.nConsCod=6051" & _
               " WHERE cPersCod='" & pcPersCod & "' AND nTipoCurriculum=" & pnTipoCurriculum
    End If
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Empleado_Curriculum = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Empleado_Curriculum:
    Err.Raise Err.Number, "Obtener_Empleado_Curriculum", Err.Description
End Function

Public Function Obtener_Empleado_DatosComplementarios(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Empleado_DatosComplementarios
    
    sSql = "SELECT  cDomicilioLegal=ISNULL(cDomicilioLegal,''),cDomicilioActual=ISNULL(cDomicilioActual,''), nTipoVivienda=ISNULL(nTipoVivienda,0),Vivienda=ISNULL(V.cConsDescripcion,''),nCostoAlquiler=ISNULL(nCostoAlquiler,0), nTipoConvivencia=ISNULL(nTipoConvivencia,0),Convivencia=ISNULL(C.cConsDescripcion,''), " & _
           " cOtrosConvive=ISNULL(cOtrosConvive,''),cLugarNacimiento=ISNULL(cLugarNacimiento,''), nCategoriaBrevete=ISNULL(nCategoriaBrevete,0),Brevete=ISNULL(B.cConsDescripcion,''),cNombreEmergencia=ISNULL(cNombreEmergencia,''), cParentescoEmergencia =ISNULL(cParentescoEmergencia,''), " & _
           " cTelefonoEmergencia=ISNULL(cTelefonoEmergencia,''),bSeguroPrivado=CONVERT(int,ISNULL(bSeguroPrivado,0)),nNumeroDependientes=ISNULL(nNumeroDependientes,0),nLugarVivienda=ISNULL(nLugarVivienda,0),LugarVivienda=ISNULL(T.cConsDescripcion,''), " & _
           " cUbiGeoL=ISNULL(cUbiGeoDomLegal,''),UbiLegal=ISNULL(Legal.cUbiGeoDescripcion,''),cUbiGeoA=ISNULL(cUbiGeoDomActual,''),UbiActual=ISNULL(Actual.cUbiGeoDescripcion,'')" & _
           " FROM RHEmpleadoDatosComplementarios LEFT JOIN Constante C " & _
           " ON nTipoConvivencia=C.nConsValor AND C.nConsCod=9063 LEFT JOIN Constante V ON nTipoVivienda=V.nConsValor AND V.nConsCod=9062 " & _
           " LEFT JOIN Constante B ON nCategoriaBrevete=B.nConsValor AND B.nConsCod=9068 " & _
           " LEFT JOIN Constante T ON nLugarVivienda=T.nConsValor AND T.nConsCod=9078 " & _
           " LEFT JOIN UbicacionGeografica Legal ON cUbiGeoDomLegal= Legal.cUbiGeoCod" & _
           " LEFT JOIN UbicacionGeografica Actual ON cUbiGeoDomActual= Actual.cUbiGeoCod" & _
           " WHERE cPersCod='" & pcPersCod & "'"

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Empleado_DatosComplementarios = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Empleado_DatosComplementarios:
    Err.Raise Err.Number, "Obtener_Empleado_DatosComplementarios", Err.Description
End Function

Public Function Obtener_Item_Curriculum(ByVal pcPersCod As String, ByVal pnTipoCurriculum As Integer) As Integer
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorObtener_Item_Curriculum
    
    sSql = "SELECT  Item=ISNULL(MAX(nItem),0)+1 FROM RHEmpleadoCurriculum " & _
           " WHERE cPersCod='" & pcPersCod & "' AND nTipoCurriculum=" & pnTipoCurriculum

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        Obtener_Item_Curriculum = rs("Item")
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Item_Curriculum:
    Err.Raise Err.Number, "Obtener_Item_Curriculum", Err.Description
End Function


Public Function Obtener_Item_Experiencia(ByVal pcPersCod As String) As Integer
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorObtener_Item_Experiencia
    
    sSql = "SELECT  Item=ISNULL(MAX(nItem),0)+1 FROM RHEmpleadoCurriculumExperiencia " & _
           " WHERE cPersCod='" & pcPersCod & "'"

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        Obtener_Item_Experiencia = rs("Item")
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Item_Experiencia:
    Err.Raise Err.Number, "Obtener_Item_Experiencia", Err.Description
End Function

Public Function Obtener_Item_Salud(ByVal pcPersCod As String) As Integer
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorObtener_Item_Salud
    
    sSql = "SELECT  Item=ISNULL(MAX(nItem),0)+1 FROM RHEmpleadoSaludEnfermedad " & _
           " WHERE cPersCod='" & pcPersCod & "'"

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        Obtener_Item_Salud = rs("Item")
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Item_Salud:
    Err.Raise Err.Number, "Obtener_Item_Salud", Err.Description
End Function

Public Function Obtener_Item_Habilidad(ByVal pcPersCod As String) As Integer
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorObtener_Item_Habilidad
    
    sSql = "SELECT  Item=ISNULL(MAX(nItem),0)+1 FROM RHEmpleadoHabilidades " & _
           " WHERE cPersCod='" & pcPersCod & "'"

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        Obtener_Item_Habilidad = rs("Item")
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Item_Habilidad:
    Err.Raise Err.Number, "Obtener_Item_Habilidad", Err.Description
End Function

Public Function Obtener_Especialidades_Curriculum(Optional ByVal pbTecnico As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Especialidades_Curriculum
    
    sSql = "SELECT nConsValor,cConsDescripcion,1 FROM Constante WHERE nConsCod =6044 AND nConsCod<>nConsValor " & IIf(pbTecnico = True, " AND cConsDescripcion LIKE '%CARRERA%TECNICA%' ", " AND cConsDescripcion NOT LIKE '%CARRERA%TECNICA%'") & " ORDER BY cConsDescripcion"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Especialidades_Curriculum = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Especialidades_Curriculum:
    Err.Raise Err.Number, "Obtener_Especialidades_Curriculum", Err.Description
End Function

Public Function Obtener_Especialidades_PostGrados() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Especialidades_PostGrados
    
    sSql = "SELECT nConsValor,cConsDescripcion,1 FROM Constante WHERE nConsCod =9050 AND nConsCod<>nConsValor ORDER BY cConsDescripcion"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Especialidades_PostGrados = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Especialidades_PostGrados:
    Err.Raise Err.Number, "Obtener_Especialidades_PostGrados", Err.Description
End Function

Public Function Obtener_Datos_Generales(ByVal pcPersCod As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_Generales
    
    sSql = "select cagedescripcion,careadescripcion,cRHCargoDescripcion,dingreso,dcese=isnull(CONVERT(varchar(10),dcese,103),'')," & _
           " NivelRemun=isnull(NIVEL.cNomTab,''),SubNivelRemun=isnull(SUBNIVEL.cNomTab,'')" & _
           " FROM rrhh rh inner join " & _
           " RHEmpleado Emp on rh.cPersCod=Emp.cPersCod inner join agencias a on rh.cagenciaactual=a.cagecod " & _
           " inner join areas ar on ar.careacod=rh.careacodactual inner join RHCargos Car on rh.cPersCod=Car.cPersCod " & _
           " and drhcargofecha=(select top 1 dRHCargoFecha from rhcargos where cperscod=car.cperscod order by dRHCargoFecha desc)" & _
           " inner join RHCargosTabla Cat on Car.cRHCargoCod=Cat.cRHCargoCod LEFT join DBComunes..TablaCod SUBNIVEL ON " & _
           " Cat.crhniv_subNiv=SUBNIVEL.cValor LEFT join DBComunes..TablaCod NIVEL ON LEFT(SUBNIVEL.cValor,2)=NIVEL.cValor " & _
           " AND LEFT(NIVEL.cCodTab,2)= LEFT(SUBNIVEL.cCodTab,2)" & _
            " where car.cperscod='" & pcPersCod & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_Generales = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Datos_Generales:
    Err.Raise Err.Number, "Obtener_Datos_Generales", Err.Description
End Function

'Public Sub InsertaRHEmpleadoUbicacion(ByVal pcPersCod As String, _
'                                ByVal pcRHAgenciaCod As String, _
'                                ByVal pcAnexo As String, _
'                                ByVal pcRPM As String, _
'                                ByVal pcCorreo As String)
'
'Dim sSQL As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorInsertaRHEmpleadoUbicacion
'    sSQL = "INSERT INTO RHEmpleadoUbicacion(cPersCod, cRHAgenciaCod,cAnexo, cRPM, cCorreo)"
'    sSQL = sSQL & " VALUES('" & pcPersCod & "','" & pcRHAgenciaCod & "','" & pcAnexo & "','" & pcRPM & "','" & pcCorreo & "')"
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSQL
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Sub
'
'ErrorInsertaRHEmpleadoUbicacion:
'    Err.Raise Err.Number, "InsertaRHEmpleadoUbicacion", Err.Description
'End Sub
'
'Public Sub ModificaRHEmpleadoUbicacion(ByVal pcPersCod As String, _
'                                ByVal pcRHAgenciaCod As String, _
'                                ByVal pcAnexo As String, _
'                                ByVal pcRPM As String, _
'                                ByVal pcCorreo As String)
'Dim sSQL As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorModificaRHEmpleadoUbicacion
'    sSQL = "UPDATE RHEmpleadoUbicacion" '
'    sSQL = sSQL & " SET cPersCod='" & pcPersCod & _
'                    "',cRHAgenciaCod='" & pcRHAgenciaCod & _
'                    "',cAnexo ='" & pcAnexo & _
'                    "',cRPM='" & pcRPM & _
'                    "',cCorreo='" & pcCorreo & _
'                    "' WHERE cPersCod='" & pcPersCod & "'"
'
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSQL
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Sub
'
'ErrorModificaRHEmpleadoUbicacion:
'    Err.Raise Err.Number, "ModificaRHEmpleadoUbicacion", Err.Description
'End Sub

'Public Function EliminaRHEmpleadoUbicacion(ByVal pcPersCod As String)
'Dim sSQL As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorEliminaRHEmpleadoUbicacion
'    sSQL = "DELETE FROM RHEmpleadoUbicacion"
'    sSQL = sSQL & " WHERE cPersCod='" & pcPersCod & "'"
'
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSQL
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorEliminaRHEmpleadoUbicacion:
'    Err.Raise Err.Number, "EliminaRHEmpleadoUbicacion", Err.Description
'End Function

Public Sub InsertaRHEmpleadoUbicacion(ByVal pcPersCod As String, _
                                ByVal pcRHAgenciaCod As String, _
                                ByVal pcContacto As String, _
                                ByVal pnTipoContacto As Integer, _
                    Optional ByVal pcUltimaActualizacion As String = "")

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEmpleadoUbicacion
    sSql = "INSERT INTO RHEmpleadoUbicacion(cPersCod, cRHAgenciaCod,cContacto, nTipoContacto,cUltimaActualizacion)"
    sSql = sSql & " VALUES('" & pcPersCod & "','" & pcRHAgenciaCod & "','" & pcContacto & "'," & pnTipoContacto & ",'" & pcUltimaActualizacion & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEmpleadoUbicacion:
    Err.Raise Err.Number, "InsertaRHEmpleadoUbicacion", Err.Description
End Sub

Public Sub ModificaRHEmpleadoUbicacion(ByVal pcPersCod As String, _
                                ByVal pcRHAgenciaCod As String, _
                                ByVal pcContacto As String, _
                                ByVal pnTipoContacto As Integer, _
                    Optional ByVal pcUltimaActualizacion As String = "")
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHEmpleadoUbicacion
    sSql = "UPDATE RHEmpleadoUbicacion" '
    sSql = sSql & " SET cContacto ='" & pcContacto & _
                    "',nTipoContacto=" & pnTipoContacto & _
                    ",cUltimaActualizacion='" & pcUltimaActualizacion & _
                    "' WHERE cPersCod='" & pcPersCod & "' AND cRHAgenciaCod='" & pcRHAgenciaCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEmpleadoUbicacion:
    Err.Raise Err.Number, "ModificaRHEmpleadoUbicacion", Err.Description
End Sub

Public Function EliminaRHEmpleadoUbicacion(ByVal pcPersCod As String, _
                                        ByVal pcRHAgenciaCod As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEmpleadoUbicacion
    sSql = "DELETE FROM RHEmpleadoUbicacion"
    sSql = sSql & " WHERE cPersCod='" & pcPersCod & "' AND cRHAgenciaCod='" & pcRHAgenciaCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEmpleadoUbicacion:
    Err.Raise Err.Number, "EliminaRHEmpleadoUbicacion", Err.Description
End Function

Public Function Obtener_Empleado_Ubicacion(ByVal pcPersCod As String, Optional ByVal pbLista As Boolean = True) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Empleado_Ubicacion
    
    'sSQL = " SELECT cPersCod,cRHAgenciaCod,cAgeDescripcion,cAnexo,cRPM,cCorreo" & _
           " FROM RHEmpleadoUbicacion EU INNER JOIN Agencias A On A.cAgeCod=EU.cRHAgenciaCod" & _
           " WHERE cPersCod='" & pcPersCod & "'"
    If pbLista = True Then
        sSql = " SELECT cPersCod,cRHAgenciaCod,cAgeDescripcion,nTipoContacto,cConsDescripcion,cContacto" & _
           " FROM RHEmpleadoUbicacion EU INNER JOIN Agencias A On A.cAgeCod=EU.cRHAgenciaCod " & _
           " INNER JOIN Constante C ON C.nConsValor=nTipoContacto AND nConsCod=9094" & _
           " WHERE cPersCod='" & pcPersCod & "'"
    Else
         sSql = " SELECT P.cPersCod,cAreaCodActual,cAreaDescripcion,cAgenciaActual,cAgeDescripcion " & _
           " FROM Persona P INNER JOIN RRHH ON P.cPersCod = RRHH.cPersCod  AND nRHEstado<700 " & _
           " INNER JOIN Areas A ON RRHH.cAreaCodActual= A.cAreaCod " & _
           " INNER JOIN Agencias Ag ON RRHH.cAgenciaActual=Ag.cAgeCod " & _
           " WHERE P.cPersCod='" & pcPersCod & "'"
    End If
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Empleado_Ubicacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Empleado_Ubicacion:
    Err.Raise Err.Number, "Obtener_Empleado_Ubicacion", Err.Description
End Function

Public Function Obtener_Datos_Consulta(ByVal pcPersCod As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_Consulta
    
    sSql = "SELECT cagedescripcion,careadescripcion,cRHCargoDescripcion,dPersNacCreac=ISNULL(CONVERT(varchar(10),dPersNacCreac,103),'')" & _
           "" & _
           " FROM rrhh rh inner join  RHEmpleado Emp on rh.cPersCod=Emp.cPersCod inner join agencias a on rh.cagenciaactual=a.cagecod " & _
           " inner join Persona P on P.cPersCod= Emp.cPersCod inner join areas ar on ar.careacod=rh.careacodactual " & _
           " inner join RHCargos Car on rh.cPersCod=Car.cPersCod and drhcargofecha=(select top 1 dRHCargoFecha from rhcargos where cperscod=car.cperscod order by dRHCargoFecha desc)" & _
           " inner join RHCargosTabla Cat on Car.cRHCargoCod=Cat.cRHCargoCod left join RHEmpleadoUbicacion UBI ON UBI.cPersCod=P.cPersCod" & _
            " where car.cperscod='" & pcPersCod & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    
        Set Obtener_Datos_Consulta = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Datos_Consulta:
    Err.Raise Err.Number, "Obtener_Datos_Consulta", Err.Description
End Function

Public Sub InsertaRHEmpleadoDatosFicha(ByVal pcPersCod As String, _
                                ByVal pnTipoDatos As Integer, _
                                ByVal pnEstado As Integer)

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrorInsertaRHEmpleadoDatosFicha
    Set rs = Obtener_EmpleadoDatosFicha(pcPersCod, pnTipoDatos)
    If rs.EOF Then
        sSql = "INSERT INTO RHEmpleadoDatosFicha(cPersCod, nTipoDatos,nEstado)"
        sSql = sSql & " VALUES('" & pcPersCod & "'," & pnTipoDatos & "," & pnEstado & ")"
    Else
        'Si Existe lo colocamos por Confirmar
        sSql = "UPDATE RHEmpleadoDatosFicha"
        sSql = sSql & " SET nEstado=1 " & _
                    " WHERE cPersCod='" & pcPersCod & "' AND nTipoDatos=" & pnTipoDatos
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Set rs = Nothing
    Exit Sub
    
ErrorInsertaRHEmpleadoDatosFicha:
    Err.Raise Err.Number, "InsertaRHEmpleadoDatosFicha", Err.Description
End Sub

Public Sub ModificaRHEmpleadoDatosFicha(ByVal pcPersCod As String, _
                                ByVal pnTipoDatos As Integer, _
                                ByVal pnEstado As Integer)
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
        
    On Error GoTo ErrorModificaRHEmpleadoDatosFicha
    
    'Si enviamos -1 en nTipoDatos modifica todos
    If pnTipoDatos = 0 Then
        sSql = "UPDATE RHEmpleadoDatosFicha SET nEstado=" & pnEstado & " WHERE cPersCod='" & pcPersCod & "'"
    Else

        Set rs = Obtener_EmpleadoDatosFicha(pcPersCod, pnTipoDatos)
        If rs.EOF Then
            sSql = "INSERT INTO RHEmpleadoDatosFicha(cPersCod, nTipoDatos,nEstado)"
            sSql = sSql & " VALUES('" & pcPersCod & "'," & pnTipoDatos & "," & pnEstado & ")"
        Else
            sSql = "UPDATE RHEmpleadoDatosFicha" '
            sSql = sSql & " SET nEstado=" & pnEstado & _
                    " WHERE cPersCod='" & pcPersCod & "' AND nTipoDatos=" & pnTipoDatos
        End If
    End If
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEmpleadoDatosFicha:
    Err.Raise Err.Number, "ModificaRHEmpleadoDatosFicha", Err.Description
End Sub

Public Function Obtener_EmpleadoDatosFicha(ByVal pcPersCod As String, ByVal pnTipoDatos As Integer) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_EmpleadoDatosFicha
    
    sSql = " SELECT nTipoDatos,nEstado FROM RHEmpleadoDatosFicha" & _
           " WHERE cPersCod='" & pcPersCod & "' AND nTipoDatos=" & pnTipoDatos
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    
        Set Obtener_EmpleadoDatosFicha = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_EmpleadoDatosFicha:
    Err.Raise Err.Number, "Obtener_EmpleadoDatosFicha", Err.Description
End Function

Public Sub InsertaPersRelaciones(ByVal pcPersCod As String, _
                    ByVal pcPersRelacPersCod As String, _
                    ByVal pnPersRelac As Integer, _
                    ByVal pcMovNro As String, _
                    ByVal pcGradoInstruccion As String, _
                    ByVal pbCondicionLaboral As Integer, _
                    ByVal pcPersRelacLugarNcto As String, _
                    ByVal pcPersRelacCentroLaboral As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaPersRelaciones
    sSql = "INSERT INTO PersRelaciones(cPersCod, cPersRelacPersCod, nPersRelac, cMovNro, cGradoInstruccion, bCondicionLaboral,cPersRelacLugarNcto,cPersRelacCentroLaboral)"
    sSql = sSql & " VALUES('" & pcPersCod & "','" & pcPersRelacPersCod & "'," & pnPersRelac & ",'" & pcMovNro & "','" & pcGradoInstruccion & "'," & pbCondicionLaboral & ",'" & pcPersRelacLugarNcto & "','" & pcPersRelacCentroLaboral & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaPersRelaciones:
    Err.Raise Err.Number, "InsertaPersRelaciones", Err.Description
End Sub

Public Function EliminaPersRelaciones(ByVal pcPersCod As String, _
                                      ByVal pcPersRelacPersCod As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaPersRelaciones
    sSql = "DELETE FROM PersRelaciones"
    sSql = sSql & " WHERE cPersCod='" & pcPersCod & "' AND cPersRelacPersCod ='" & pcPersRelacPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaPersRelaciones:
    Err.Raise Err.Number, "EliminaPersRelaciones", Err.Description
End Function

Public Sub InsertaPersID(ByVal pcPersCod As String, _
                    ByVal pcPersIDTipo As String, _
                    ByVal pcPersIDnro As String, _
                    ByVal pcUltimaActualizacion As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaPersID
    sSql = "INSERT INTO PersID(cPersCod, cPersIDTpo, cPersIDnro, cUltimaActualizacion)"
    sSql = sSql & " VALUES('" & pcPersCod & "','" & pcPersIDTipo & "','" & pcPersIDnro & "','" & pcUltimaActualizacion & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaPersID:
    Err.Raise Err.Number, "InsertaPersID", Err.Description
End Sub

Public Sub ModificaPersID(ByVal pcPersCod As String, _
                    ByVal pcPersIDTipo As String, _
                    ByVal pcPersIDnro As String, _
                    ByVal pcUltimaActualizacion As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaPersID
    sSql = "UPDATE PersID SET cPersIDnro='" & pcPersIDnro & "',cUltimaActualizacion='" & pcUltimaActualizacion & _
            "' WHERE cPersCod='" & pcPersCod & "' AND cPersIDTpo='" & pcPersIDTipo & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaPersID:
    Err.Raise Err.Number, "ModificaPersID", Err.Description
End Sub

Public Function EliminaPersID(ByVal pcPersCod As String, _
                              ByVal pcPersIDTpo As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaPersID
    sSql = "DELETE FROM PersID"
    sSql = sSql & " WHERE cPersCod='" & pcPersCod & "' AND cPersIDTpo ='" & pcPersIDTpo & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaPersID:
    Err.Raise Err.Number, "EliminaPersID", Err.Description
End Function

Public Sub ModificaRHFichaPersonal_Persona(ByVal pcPersCod As String, _
                                            ByVal pcPersNombre As String, _
                                            ByVal pdFechaNacimiento As Date, _
                                            ByVal pcNacionalidad As String, _
                                            ByVal pcSexo As String, _
                                            ByVal pcPersTelefono As String, _
                                            ByVal pcPersTelefono2 As String, _
                                            ByVal pcPersEmail As String, _
                                            ByVal pcPersNatHijos As String, _
                                            ByVal pnPersNatEstCiv As Long, _
                                Optional ByVal pcUltimaActualizacion As String = "")
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHFichaPersonal_Persona
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    
    sSql = "UPDATE Persona" '
    sSql = sSql & " SET cPersNombre='" & pcPersNombre & _
                    "',dPersNacCreac='" & Format(pdFechaNacimiento, gsFormatoFecha) & _
                    "',cPersTelefono='" & pcPersTelefono & _
                    "', cPersTelefono2='" & pcPersTelefono2 & _
                    "',cPersEmail='" & pcPersEmail & _
                    "',cUltimaActualizacion='" & pcUltimaActualizacion & _
                    "' WHERE cPersCod='" & pcPersCod & "'"
                        
    oConecta.ConexionActiva.Execute sSql
    
    sSql = "UPDATE PersonaNat"
    sSql = sSql & " SET nPersNatHijos='" & pcPersNatHijos & _
                    "',nPersNatEstCiv=" & pnPersNatEstCiv & _
                    ",cNacionalidad='" & pcNacionalidad & "',cPersNatSexo='" & pcSexo & _
                    "' WHERE cPersCod='" & pcPersCod & "'"
                        
    oConecta.ConexionActiva.Execute sSql
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHFichaPersonal_Persona:
    Err.Raise Err.Number, "ModificaRHFichaPersonal_Persona", Err.Description
End Sub

'Public Function Obtener_Ficha_DatosPersonales(ByVal pcPersCod As String) As adodb.Recordset
'Dim sSQL As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorObtener_Ficha_DatosPersonales
'
'    sSQL = "" & _
'           " WHERE cPersCod='" & pcPersCod & "'"
'
'    Set oConecta = New DConecta
'    If oConecta.AbreConexion Then
'    'oConecta.ConexionActiva.Execute sSql
'        Set Obtener_Ficha_DatosPersonales = oConecta.CargaRecordSet(sSQL)
'    End If
'    oConecta.CierraConexion
'
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorObtener_Ficha_DatosPersonales:
'    Err.Raise Err.Number, "Obtener_Ficha_DatosPersonales", Err.Description
'End Function

Public Function Obtener_Empleado_Relaciones(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Empleado_Relaciones
    
    sSql = "SELECT  P.cPersCod,P.cPersNombre,Edad=DATEDIFF(YEAR,dPersNacCreac,getdate()) ," & _
           " Nat.cPersNatSexo,cPersIDNro,Relacion=C.cConsDescripcion, PR.cGradoInstruccion," & _
           " bCondicionLaboral=CONVERT(int,PR.bCondicionLaboral),Centro=ISNULL(cPersRelacCentroLaboral,''),LugarNacim=ISNULL(cPersRelacLugarNcto,'') " & _
           " FROM PersRelaciones PR  INNER JOIN Persona P   ON PR.cPersRelacPersCod=P.cPersCod " & _
           " INNER JOIN PersonaNat Nat ON P.cPersCod=Nat.cPersCod LEFT JOIN PersID ON " & _
           " P.cPersCod=PersID.cPersCod AND cPersIDTpo IN('1','10') INNER JOIN Constante C " & _
           " ON PR.nPersRelac=C.nConsValor AND nConsCod=1006 " & _
           " WHERE PR.cPersCod='" & pcPersCod & "'"

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Empleado_Relaciones = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Empleado_Relaciones:
    Err.Raise Err.Number, "Obtener_Empleado_Relaciones", Err.Description
End Function

Public Function Habilitar_Deshabilitar_Actualizacion(ByVal pbActualiza As Integer, ByVal pcUltimaActualizacion As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorHabilitar_Deshabilitar_Actualizacion
    sSql = "UPDATE ConstSistema SET cUltimaActuzalizacion='" & pcUltimaActualizacion & "',nConsSisValor=" & pbActualiza & _
           " WHERE nConsSisCod =" & gnActualizarFichaPersonal
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorHabilitar_Deshabilitar_Actualizacion:
    Err.Raise Err.Number, "Habilitar_Deshabilitar_Actualizacion", Err.Description
End Function



'******************* Funciones para la Impresion de Ficha de Personal ***********

Public Function DatosCabecera(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosEmpleadoErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = "select rh.cRHCod,cagedescripcion,careadescripcion,cRHCargoDescripcion,dingreso,dcese=isnull(CONVERT(varchar(10),dcese,103),''), " & _
           " NivelRemun=isnull(NIVEL.cNomTab,''),SubNivelRemun=isnull(SUBNIVEL.cNomTab,''), " & _
           " AFP=ISNULL(AFP.cPersNombre,''),cCUSPP=ISNULL(cCUSPP,'') FROM rrhh rh inner join " & _
           " RHEmpleado Emp on rh.cPersCod=Emp.cPersCod inner join agencias a on rh.cagenciaactual=a.cagecod " & _
           " inner join areas ar on ar.careacod=rh.careacodactual inner join RHCargos Car on rh.cPersCod=Car.cPersCod " & _
           " and drhcargofecha=(select top 1 dRHCargoFecha from rhcargos where cperscod=car.cperscod order by dRHCargoFecha desc) " & _
           " inner join RHCargosTabla Cat on Car.cRHCargoCod=Cat.cRHCargoCod LEFT join DBComunes..TablaCod SUBNIVEL ON " & _
           " Cat.crhniv_subNiv=SUBNIVEL.cValor LEFT join DBComunes..TablaCod NIVEL ON LEFT(SUBNIVEL.cValor,2)=NIVEL.cValor " & _
           " AND LEFT(NIVEL.cCodTab,2)= LEFT(SUBNIVEL.cCodTab,2) LEFT JOIN Persona AFP ON cRHEmplAFPPersCod=AFP.cPersCod WHERE car.cperscod='" & pcPersCod & "'"
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosCabecera = rs
'    End If
    Exit Function
DatosEmpleadoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosEmpleado(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosEmpleadoErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = "Select  P.cPersEmail cEmail,P.cPersTelefono2,P.cPersCod,P.cPersNombre,P.dPersNacCreac, " & _
            " P.cPersTelefono, PN.cPersNatSexo,PN.nPersNatEstCiv,PN.nPersNatHijos,PN.nPersTpoSangre,Nacionalidad='PERU',  " & _
            " Edad=DATEDIFF(YEAR,P.dPersNacCreac,getdate()), " & _
            " estado=isnull((select cconsdescripcion from constante where nconscod=1020 and nconsvalor=npersnatestciv),''), " & _
            " tposangre=isnull((select cconsdescripcion from constante where nconscod=1039 and nconsvalor=cperstposangre),'') " & _
            " FROM Persona P left join PersonaNat PN ON P.cPersCod=PN.cPersCod " & _
            " where p.cperscod='" & pcPersCod & "'"
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosEmpleado = rs
'    End If
    Exit Function
DatosEmpleadoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosDocIdentificacion(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosDocIdentificacionErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = " select nconsvalor,cconsdescripcion,doc=isnull((select cpersidnro from persid where cperscod='" & pcPersCod & "' and cpersidtpo=nconsvalor ),'') " & _
            " from constante where nconscod=1003 " & _
            " and nconsvalor in (1,7) "
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosDocIdentificacion = rs
'    End If
    Exit Function
DatosDocIdentificacionErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosOtros(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosOtrosErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = " SELECT  cDomicilioLegal,cDomicilioActual=ISNULL(cDomicilioActual,''), nTipoVivienda,Vivienda=V.cConsDescripcion, nCostoAlquiler, nTipoConvivencia,Convivencia=C.cConsDescripcion, " & _
            " cOtrosConvive, cLugarNacimiento, nCategoriaBrevete,Brevete=B.cConsDescripcion, cNombreEmergencia, cParentescoEmergencia , " & _
            " cTelefonoEmergencia,bSeguroPrivado=CONVERT(int,bSeguroPrivado), nNumeroDependientes, " & _
            " DepL=rtrim(isnull((select cUbiGeoDescripcion from UbicacionGeografica where cUbiGeoCod like '1'+substring(cUbiGeoDomLegal,2,2)+'%'),''))," & _
            " ProvL=rtrim(isnull((select cUbiGeoDescripcion from UbicacionGeografica where cUbiGeoCod like '2'+substring(cUbiGeoDomLegal,2,4)+'%'),'')), " & _
            " DistL=rtrim(isnull((select cUbiGeoDescripcion from UbicacionGeografica where cUbiGeoCod like '3'+substring(cUbiGeoDomLegal,2,6)+'%'),'')), " & _
            " UbiL=isnull(Legal.cUbiGeoDescripcion,''), " & _
            " DepA=rtrim(isnull((select cUbiGeoDescripcion from UbicacionGeografica where cUbiGeoCod like '1'+substring(cUbiGeoDomActual,2,2)+'%'),'')), " & _
            " ProvA=rtrim(isnull((select cUbiGeoDescripcion from UbicacionGeografica where cUbiGeoCod like '2'+substring(cUbiGeoDomActual,2,4)+'%'),'')), " & _
            " DistA=rtrim(isnull((select cUbiGeoDescripcion from UbicacionGeografica where cUbiGeoCod like '3'+substring(cUbiGeoDomActual,2,6)+'%'),'')), " & _
            " UbiA=isnull(Actual.cUbiGeoDescripcion,'') " & _
            " FROM RHEmpleadoDatosComplementarios INNER JOIN Constante C " & _
            " ON nTipoConvivencia=C.nConsValor AND C.nConsCod=9063 INNER JOIN Constante V ON nTipoVivienda=V.nConsValor AND V.nConsCod=9062 " & _
            " LEFT JOIN Constante B ON nCategoriaBrevete=B.nConsValor AND B.nConsCod=9068 " & _
            " LEFT JOIN UbicacionGeografica Legal ON cUbiGeoDomLegal= Legal.cUbiGeoCod " & _
            " LEFT JOIN UbicacionGeografica Actual ON cUbiGeoDomActual= Actual.cUbiGeoCod " & _
            " WHERE cPersCod='" & pcPersCod & "'"
            
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosOtros = rs
'    End If
    Exit Function
DatosOtrosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosPrimerGradoCon(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosPrimerGradoConErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = " select p.cpersnombre,dpersnaccreac,cpersdireccubigeo,cpersdireccdomicilio, " & _
           " isnull(cperstelefono,'')cperstelefono , isnull(cperstelefono2,'')cperstelefono2, PR.cpersrelacperscod, " & _
           " cconsdescripcion, npersrelac,cPersNatSexo, " & _
           " Edad=DATEDIFF(YEAR,P.dPersNacCreac,getdate()), " & _
           " DNI=(SELECT cPersIDnro FROM PersId where cPerscod=p.cperscod and cpersidtpo=1), " & _
           " Partida=(SELECT cPersIDnro FROM PersId where cPerscod=p.cperscod and cpersidtpo=10) " & _
           " from dbo.PersRelaciones pr inner join persona p on pr.cpersrelacperscod=p.cperscod " & _
           " inner join PersonaNat Nat on Nat.cPersCod= P.cPersCod inner join constante c on c.nconsvalor=pr.npersrelac  and nconscod=1006 and nconsvalor IN(1,2) " & _
           " WHERE pr.cPersCod='" & pcPersCod & "'"
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosPrimerGradoCon = rs
'    End If
    Exit Function
DatosPrimerGradoConErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosSegundoGradoCon(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosSegundoGradoConErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = "  select p.cpersnombre,dpersnaccreac,cpersdireccubigeo,cpersdireccdomicilio, " & _
           " isnull(cperstelefono,'')cperstelefono , isnull(cperstelefono2,'')cperstelefono2, " & _
           " PR.cpersrelacperscod, cconsdescripcion, npersrelac,cPersNatSexo, " & _
           " Edad=DATEDIFF(YEAR,P.dPersNacCreac,getdate()), " & _
           " DNI=(SELECT cPersIDnro FROM PersId where cPerscod=p.cperscod and cpersidtpo=1), " & _
           " Partida=(SELECT cPersIDnro FROM PersId where cPerscod=p.cperscod and cpersidtpo=10) " & _
           " from dbo.PersRelaciones pr inner join persona p on pr.cpersrelacperscod=p.cperscod " & _
           " inner join PersonaNat Nat on Nat.cPersCod= P.cPersCod inner join constante c on c.nconsvalor=pr.npersrelac  and nconscod=1006 and nconsvalor IN(3,4,11) " & _
           " WHERE pr.cPersCod='" & pcPersCod & "'"
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosSegundoGradoCon = rs
'    End If
    Exit Function
DatosSegundoGradoConErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosTercerGradoCon(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosTercerGradoConErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = " select p.cpersnombre,dpersnaccreac,cpersdireccubigeo,cpersdireccdomicilio, " & _
           " isnull(cperstelefono,'')cperstelefono , isnull(cperstelefono2,'')cperstelefono2, " & _
           " PR.cpersrelacperscod, cconsdescripcion, npersrelac,cPersNatSexo, " & _
           " Edad=DATEDIFF(YEAR,P.dPersNacCreac,getdate()), " & _
           " DNI=(SELECT cPersIDnro FROM PersId where cPerscod=p.cperscod and cpersidtpo=1), " & _
           " Partida=(SELECT cPersIDnro FROM PersId where cPerscod=p.cperscod and cpersidtpo=10) " & _
           " from dbo.PersRelaciones pr inner join persona p on pr.cpersrelacperscod=p.cperscod " & _
           " inner join PersonaNat Nat on Nat.cPersCod= P.cPersCod inner join constante c on c.nconsvalor=pr.npersrelac  and nconscod=1006 and nconsvalor IN(14,7) " & _
           " WHERE pr.cPersCod='" & pcPersCod & "'"
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    Set DatosTercerGradoCon = rs
    Exit Function
DatosTercerGradoConErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosPrimerGradoAfi(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosPrimerGradoAfiErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = "  select p.cpersnombre,dpersnaccreac,cpersdireccubigeo,cpersdireccdomicilio, " & _
           " isnull(cperstelefono,'')cperstelefono , isnull(cperstelefono2,'')cperstelefono2, PR.cpersrelacperscod, cconsdescripcion, npersrelac,cPersNatSexo, " & _
           " Edad=DATEDIFF(YEAR,P.dPersNacCreac,getdate()), " & _
           " DNI=(SELECT cPersIDnro FROM PersId where cPerscod=p.cperscod and cpersidtpo=1), " & _
           " Partida=(SELECT cPersIDnro FROM PersId where cPerscod=p.cperscod and cpersidtpo=10) " & _
           " from dbo.PersRelaciones pr inner join persona p on pr.cpersrelacperscod=p.cperscod " & _
           " inner join PersonaNat Nat on Nat.cPersCod= P.cPersCod inner join constante c on c.nconsvalor=pr.npersrelac  and nconscod=1006 and nconsvalor IN(0,12) " & _
           " left join PersId on PersId.cPersCod=P.cPersCod " & _
           " WHERE pr.cPersCod='" & pcPersCod & "'"
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosPrimerGradoAfi = rs
'    End If
    Exit Function
DatosPrimerGradoAfiErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosSegundoGradoAfi(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosSegundoGradoAfiErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = " select p.cpersnombre,dpersnaccreac,cpersdireccubigeo,cpersdireccdomicilio, " & _
           " isnull(cperstelefono,'')cperstelefono , isnull(cperstelefono2,'')cperstelefono2, PR.cpersrelacperscod, cconsdescripcion, npersrelac,cPersNatSexo, " & _
           " Edad=DATEDIFF(YEAR,P.dPersNacCreac,getdate()), " & _
           " DNI=(SELECT cPersIDnro FROM PersId where cPerscod=p.cperscod and cpersidtpo=1), " & _
           " Partida=(SELECT cPersIDnro FROM PersId where cPerscod=p.cperscod and cpersidtpo=10) " & _
           " from dbo.PersRelaciones pr inner join persona p on pr.cpersrelacperscod=p.cperscod " & _
           " inner join PersonaNat Nat on Nat.cPersCod= P.cPersCod inner join constante c on c.nconsvalor=pr.npersrelac  and nconscod=1006 and nconsvalor =13" & _
           " WHERE pr.cPersCod='" & pcPersCod & "'"
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosSegundoGradoAfi = rs
'    End If
    Exit Function
DatosSegundoGradoAfiErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosGrados(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosGradosErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = " SELECT nItem, nCurriculumCod,C.cConsDescripcion cConsDescripcion1,dInicio=CONVERT(varchar(10),dInicio,103), dFin=CONVERT(varchar(10),dFin,103), nNivel,N.cConsDescripcion cConsDescripcion2, nGradoObtenido, " & _
            "   G.cConsDescripcion cConsDescripcion3, nProfesionCod, P.cConsDescripcion cConsDescripcion4, cLugar, cComentario, nCosto, bCmactOtorgado " & _
            "   FROM RHEmpleadoCurriculum EC INNER JOIN Constante C ON nCurriculumCod=C.nConsValor AND C.nConsCod=9051 " & _
            "   INNER JOIN Constante P ON nProfesionCod=P.nConsValor AND P.nConsCod=6044 INNER JOIN Constante N ON nNivel=N.nConsValor AND N.nConsCod=6045 " & _
            "   INNER JOIN Constante G ON nGradoObtenido=G.nConsValor AND G.nConsCod=9052 "
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosGrados = rs
'    End If
    Exit Function
DatosGradosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosEstudiosExternos(ByVal pcPersCod As String, _
                                    Optional ByVal pnTipoActividad As Integer = 0) As ADODB.Recordset

On Error GoTo DatosEstudiosExternosErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    'sSql = " Select nItem,nCurriculumCod, C.cConsDescripcion, nProfesionCod, " & _
           " Case nCurriculumCod   When 0 Then (Select cConsDescripcion From Constante COI Where COI.nConsCod = 6047 And COI.nConsValor = nProfesionCod) " & _
           " When 1 Then (Select cConsDescripcion From Constante COI Where COI.nConsCod = 6048 And COI.nConsValor = nProfesionCod) " & _
           " When 2 Then (Select cConsDescripcion From Constante COI Where COI.nConsCod = 6049 And COI.nConsValor = nProfesionCod) " & _
           " When 3 Then (Select cConsDescripcion From Constante COI Where COI.nConsCod = 6050 And COI.nConsValor = nProfesionCod) End Especialidad, " & _
           " nNivel , N.cConsDescripcion From RHEmpleadoCurriculum INNER JOIN Constante C ON nCurriculumCod=C.nConsValor AND C.nConsCod=6046 " & _
           " INNER JOIN Constante N ON nNivel=N.nConsValor AND N.nConsCod=6051 " & _
           " WHERE cPersCod='" & pcPersCod & "' AND nTipoCurriculum=2"
    
    sSql = " Select nItem,nCurriculumCod, C.cConsDescripcion, nProfesionCod, " & _
           " Case nCurriculumCod   When 0 Then (Select cConsDescripcion From Constante COI Where COI.nConsCod = 6047 And COI.nConsValor = nProfesionCod) " & _
           " When 3 Then (Select cConsDescripcion From Constante COI Where COI.nConsCod = 6050 And COI.nConsValor = nProfesionCod) End Especialidad, " & _
           " nNivel , N.cConsDescripcion From RHEmpleadoCurriculum INNER JOIN Constante C ON nCurriculumCod=C.nConsValor AND C.nConsCod=6046 " & _
           " INNER JOIN Constante N ON nNivel=N.nConsValor AND N.nConsCod=6051 " & _
           " WHERE cPersCod='" & pcPersCod & "' AND nTipoCurriculum=2 AND nCurriculumCod=" & pnTipoActividad
           
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosEstudiosExternos = rs
'    End If
    Exit Function
DatosEstudiosExternosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosDesempeñoProfesional(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosDesempeñoProfesionalErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = " SELECT car.dRHCargoFecha,CargoFecTermino=ISNULL(convert(varchar(10), (select top 1 drhCargoFecha from rhcargos where cperscod=car.cperscod " & _
           " and car.drhcargofecha<drhCargoFecha order by dRHCargoFecha asc),103),''), cat.crhcargodescripcion , age.cAgeDescripcion, crhcargocomentario " & _
           " FROM RHCargos car INNER JOIN RHCargosTabla cat ON car.crhCargoCod = cat.crhCargoCod INNER JOIN agencias age ON car.crhagenciacod=age.cagecod " & _
           " WHERE car.cPersCod='" & pcPersCod & "' order by car.dRHCargoFecha asc "
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosDesempeñoProfesional = rs
'    End If
    Exit Function
DatosDesempeñoProfesionalErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosExperienciaLaboral(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosExperienciaLaboralErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = " SELECT nItem, nTipoExperiencia,T.cConsDescripcion, cLugar,dInicio=CONVERT(varchar(10),dInicio,103),dFin=CONVERT(varchar(10),dFin,103), cCargo, cJefeSuperior, nSueldo, cMotivoSalida, cObservacion " & _
           " FROM RHEmpleadoCurriculumExperiencia CE INNER JOIN Constante T ON T.nConsValor=nTipoExperiencia AND T.nConsCod=9055 " & _
           " WHERE cPersCod='" & pcPersCod & "'"
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosExperienciaLaboral = rs
'    End If
    Exit Function
DatosExperienciaLaboralErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosVivienda(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosViviendaErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = " SELECT  cDomicilioLegal,cDomicilioActual=ISNULL(cDomicilioActual,''), nTipoVivienda,Vivienda=V.cConsDescripcion, nCostoAlquiler, nTipoConvivencia,Convivencia=C.cConsDescripcion, " & _
           " cOtrosConvive, cLugarNacimiento, nCategoriaBrevete,Brevete=B.cConsDescripcion, cNombreEmergencia, cParentescoEmergencia , " & _
           " cTelefonoEmergencia,bSeguroPrivado=CONVERT(int,bSeguroPrivado), nNumeroDependientes FROM RHEmpleadoDatosComplementarios INNER JOIN Constante C " & _
           " ON nTipoConvivencia=C.nConsValor AND C.nConsCod=9063 INNER JOIN Constante V ON nTipoVivienda=V.nConsValor AND V.nConsCod=9062 " & _
           " LEFT JOIN Constante B ON nCategoriaBrevete=B.nConsValor AND B.nConsCod=9068 " & _
           " WHERE cPersCod='" & pcPersCod & "'"
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosVivienda = rs
'    End If
    Exit Function
DatosViviendaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosEnfermedades(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosEnfermedadesErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = " SELECT nItem, nTipoSaludEnfermedad,T.cConsDescripcion, cDescripcion, cTratamiento, cMedico, nEstado,E.cConsDescripcion " & _
           " FROM RHEmpleadoSaludEnfermedad SE INNER JOIN Constante T ON T.nConsValor=nTipoSaludEnfermedad AND T.nConsCod=9056 INNER JOIN Constante E " & _
           " ON E.nConsValor=nEstado AND E.nConsCod=9057 " & _
           " WHERE cPersCod='" & pcPersCod & "'"
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosEnfermedades = rs
'    End If
    Exit Function
DatosEnfermedadesErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosHabilidades(ByVal pcPersCod As String) As ADODB.Recordset
On Error GoTo DatosHabilidadesErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = " SELECT nItem, nTipoHabilidad,T.cConsDescripcion cConsDescripcion1, nNivelHabilidad,N.cConsDescripcion cConsDescripcion2, cComentario " & _
           " FROM RHEmpleadoHabilidades SH INNER JOIN Constante T ON  nTipoHabilidad=T.nConsValor AND T.nConsCod=6048 " & _
           " INNER JOIN Constante N ON nNivelHabilidad=N.nConsValor AND N.nConsCod=6051 " & _
           " WHERE cPersCod='" & pcPersCod & "'"
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosHabilidades = rs
'    End If
    Exit Function
DatosHabilidadesErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function DatosCapacitaciones(ByVal pcPersCod As String) As ADODB.Recordset

On Error GoTo DatosCapacitacionesErr
    Dim OCON As DConecta, sSql As String, rs As ADODB.Recordset
    sSql = " SELECT cCursoDescripcion,nProgCursoAnio,Inicio=CONVERT(varchar(10),dEjeCursoFechaInicio,103), " & _
           " Fin=CONVERT(varchar(10),dEjeCursoFechaTermino,103), " & _
           " Institucion = i.cPersNombre " & _
           " FROM RHCapacitacionParticipantes Part " & _
           " INNER JOIN RHCapacitacionEjecucion Eje ON Part.nProgNro=Eje.nProgNro AND Part.nItem=Eje.nItem " & _
           " INNER JOIN Persona I ON cEjeCursoPersCodInstitucion=I.cPersCod " & _
           " INNER JOIN RHCapacitacionPrograma Prog ON Prog.nProgNro=Eje.nProgNro " & _
           " INNER JOIN RHCapacitacionCursos Cur ON Cur.nCursoCod=Prog.nCursoCod " & _
           " WHERE Part.cPersCod='" & pcPersCod & "' AND Part.nConfirmaAsistencia=1 " & _
           " ORDER BY nProgCursoAnio,dEjeCursoFechaInicio DESC"
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
'    If Not Rs.EOF Then
        Set DatosCapacitaciones = rs
'    End If
    Exit Function
DatosCapacitacionesErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Sub ModificaPersona_TipoSangre(ByVal pcPersCod As String, _
                    ByVal pnPersTpoSangre As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaPersona_TipoSangre
    sSql = "UPDATE PersonaNat" '
    sSql = sSql & " SET nPersTpoSangre='" & pnPersTpoSangre & _
                    "' WHERE cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaPersona_TipoSangre:
    Err.Raise Err.Number, "ModificaPersona_TipoSangre", Err.Description
End Sub

Public Function ListaFichasFaltantes() As ADODB.Recordset

On Error GoTo ListaFichasFaltantesErr
    Dim OCON As DConecta, sSql As String
    
    sSql = " SELECT DISTINCT RH.cRHCod, P.cPersNombre FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod =P.cPersCod " & _
           " AND RH.nRHEstado =201 AND RH.cPersCod NOT IN (SELECT cPersCod FROM RHEmpleadoDatosFicha WHERE nTipoDatos IN(1,2,3)) " & _
           " ORDER BY P.cPersNombre --Solo los Obligatorios"
    
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set ListaFichasFaltantes = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
    Exit Function
ListaFichasFaltantesErr:
    Err.Raise Err.Number, "Error", Err.Description
End Function

Public Function ObtieneFechaActualizacionFicha(ByVal psPersCod As String) As Date

On Error GoTo ObtieneFechaActualizacionFichaErr
    Dim OCON As DConecta, sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = " EXEC RHGetFechaActualizacionFicha '" & psPersCod & "'"
    
    Set OCON = New DConecta
    If OCON.AbreConexion Then
        Set rs = OCON.CargaRecordSet(sSql)
        
        ObtieneFechaActualizacionFicha = rs!Fecha
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
    Exit Function
ObtieneFechaActualizacionFichaErr:
    Err.Raise Err.Number, "Error", Err.Description
End Function

