VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dEvaluacionDesempeño"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InsertaRHEvaluacionCriterios(ByVal pcRHEvalCriterioDescri As String, _
                                    ByVal pcRHEvalCriterioComentario As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEvaluacionCriterios
    
    sSql = "INSERT INTO RHEvaluacionCriterios(cRHEvalCriterioDescri,cRHEvalCriterioComentario)"
    sSql = sSql & " VALUES('" & pcRHEvalCriterioDescri & "','" & pcRHEvalCriterioComentario & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEvaluacionCriterios:
    Err.Raise Err.Number, "InsertaRHEvaluacionCriterios", Err.Description
End Sub

Public Sub ModificaRHEvaluacionCriterios(ByVal pnRHEvalCriterioCod As Integer, _
                                    ByVal pcRHEvalCriterioDescri As String, _
                                    ByVal pcRHEvalCriterioComentario As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHEvaluacionCriterios
    sSql = "UPDATE RHEvaluacionCriterios"
    sSql = sSql & " SET cRHEvalCriterioDescri='" & pcRHEvalCriterioDescri & _
                        "',cRHEvalCriterioComentario='" & pcRHEvalCriterioComentario & _
                        "' WHERE  nRHEvalCriterioCod=" & pnRHEvalCriterioCod

                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEvaluacionCriterios:
    Err.Raise Err.Number, "ModificaRHEvaluacionCriterios", Err.Description
End Sub

Public Function EliminaRHEvaluacionCriterios(ByVal pnRHEvalCriterioCod As Integer)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEvaluacionCriterios
    sSql = "DELETE FROM RHEvaluacionCriterios" '
    sSql = sSql & " WHERE nRHEvalCriterioCod=" & pnRHEvalCriterioCod
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEvaluacionCriterios:
    Err.Raise Err.Number, "EliminaRHEvaluacionCriterios", Err.Description
End Function

Public Function Obtener_Criterios_asignados(pnRHEvalCriterioCod As Integer) As Boolean
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorObtener_Criterios_asignados
    
    sSql = "SELECT nRHEvalCriterioCod, cRHCargoCod  " & _
           " FROM RHEvaluacionCriteriosCargo WHERE nRHEvalCriterioCod=" & pnRHEvalCriterioCod
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        If Not rs.EOF Then
            Obtener_Criterios_asignados = True  'Si existen Criterios asignados
        Else
           sSql = "SELECT nRHEvalCriterioCod, cRHCod  " & _
                " FROM RHEvaluacionDesempeñoCriterios WHERE nRHEvalCriterioCod=" & pnRHEvalCriterioCod
           Set rs = oConecta.CargaRecordSet(sSql)
           Obtener_Criterios_asignados = Not rs.EOF
        End If
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Set rs = Nothing
    Exit Function
ErrorObtener_Criterios_asignados:
    Err.Raise Err.Number, "Obtener_Criterios_asignados", Err.Description
End Function

Public Function Obtener_Criterios_Evaluacion() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Criterios_Evaluacion
    
    sSql = "SELECT nRHEvalCriterioCod, cRHEvalCriterioDescri, cRHEvalCriterioComentario " & _
           " FROM RHEvaluacionCriterios "
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Criterios_Evaluacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Criterios_Evaluacion:
    Err.Raise Err.Number, "Obtener_Criterios_Evaluacion", Err.Description
End Function

Public Sub InsertaRHEvaluacionCriteriosCargo(ByVal pnRHEvalCriterioCod As Integer, _
                                    ByVal pcRHCargoCod As String, _
                                    ByVal pnPonderacion As Double)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEvaluacionCriteriosCargo
    
    sSql = "INSERT INTO RHEvaluacionCriteriosCargo(nRHEvalCriterioCod,cRHCargoCod,nPonderacion)"
    sSql = sSql & " VALUES(" & pnRHEvalCriterioCod & ",'" & pcRHCargoCod & "'," & pnPonderacion & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEvaluacionCriteriosCargo:
    Err.Raise Err.Number, "InsertaRHEvaluacionCriteriosCargo", Err.Description
End Sub

'Public Sub ModificaRHEvaluacionCriteriosCargo(ByVal pnRHEvalCriterioCod As Integer, _
'                                    ByVal pcRHCargoCod As String, _
'                                    ByVal pnPonderacion As Double)
'
'Dim sSQL As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorModificaRHEvaluacionCriteriosCargo
'    sSQL = "UPDATE RHEvaluacionCriteriosCargo"
'    sSQL = sSQL & " SET cRHEvalCriterioDescri='" & pcRHEvalCriterioDescri & _
'                        "',cRHEvalCriterioComentario='" & pcRHEvalCriterioComentario & _
'                        "' WHERE  nRHEvalCriterioCod=" & pnRHEvalCriterioCod
'
'
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSQL
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Sub
'
'ErrorModificaRHEvaluacionCriteriosCargo:
'    Err.Raise Err.Number, "ModificaRHEvaluacionCriteriosCargo", Err.Description
'End Sub

Public Function EliminaRHEvaluacionCriteriosCargo(ByVal pnRHEvalCriterioCod As Integer, _
                                                ByVal pcRHCargoCod As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEvaluacionCriteriosCargo
    sSql = "DELETE FROM RHEvaluacionCriteriosCargo" '
    sSql = sSql & " WHERE nRHEvalCriterioCod=" & pnRHEvalCriterioCod & " AND cRHCargoCod='" & pcRHCargoCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEvaluacionCriteriosCargo:
    Err.Raise Err.Number, "EliminaRHEvaluacionCriteriosCargo", Err.Description
End Function

Public Function Obtener_Cargos_x_Criterio(pnRHEvalCriterioCod As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Cargos_x_Criterio
    
    sSql = "SELECT nRHEvalCriterioCod, ECC.cRHCargoCod,Cargo=Car.cRHCargoDescripcion ,nPonderacion FROM RHEvaluacionCriteriosCargo ECC" & _
           " INNER JOIN RHCargosTabla Car ON Car.cRHCargoCod = ECC.cRHCargoCod " & _
           " Where nRHEvalCriterioCod=" & pnRHEvalCriterioCod
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Cargos_x_Criterio = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Cargos_x_Criterio:
    Err.Raise Err.Number, "Obtener_Cargos_x_Criterio", Err.Description
End Function

Public Function Obtener_Todos_Cargos() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Todos_Cargos
    
    sSql = "SELECT cRHCargoCod,cRHCargoDescripcion FROM RHCargosTabla WHERE bRHCargoEstado=1 ORDER BY cRHCargoDescripcion"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Todos_Cargos = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Todos_Cargos:
    Err.Raise Err.Number, "Obtener_Todos_Cargos", Err.Description
End Function

Public Function Obtener_Evaluadores_x_CargoEvaluado(pcRHCargoCodEvaluado As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Evaluadores_x_CargoEvaluado
    
    sSql = "SELECT cRHCargoCodEvaluador,cRHCargoDescripcion,nOrden,bTodasAgencias=CONVERT(int,bTodasAgencias) from " & _
           " RHEvaluacionNivelAprobacion ENA INNER JOIN RHCargosTabla Cat " & _
           " ON Cat.cRHCargoCod= ENA.cRHCargoCodEvaluador " & _
           " WHERE cRHCargoCodEvaluado ='" & pcRHCargoCodEvaluado & "' ORDER BY nOrden"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Evaluadores_x_CargoEvaluado = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Evaluadores_x_CargoEvaluado:
    Err.Raise Err.Number, "Obtener_Evaluadores_x_CargoEvaluado", Err.Description
End Function

Public Sub InsertaRHEvaluacionNivelAprobacion(ByVal pcRHCargoCodEvaluado As String, _
                                    ByVal pcRHCargoCodEvaluador As String, _
                                    ByVal pnOrden As Integer, _
                                    ByVal pbTodasAgencias As Integer, _
                                    Optional ByVal pbPeriodica As Integer = 0)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEvaluacionNivelAprobacion
    
    sSql = "INSERT INTO RHEvaluacionNivelAprobacion(cRHCargoCodEvaluado,cRHCargoCodEvaluador,nOrden,bTodasAgencias)"
    sSql = sSql & " VALUES('" & pcRHCargoCodEvaluado & "','" & pcRHCargoCodEvaluador & "'," & pnOrden & "," & pbTodasAgencias & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEvaluacionNivelAprobacion:
    Err.Raise Err.Number, "InsertaRHEvaluacionNivelAprobacion", Err.Description
End Sub

Public Sub ModificaRHEvaluacionNivelAprobacion(ByVal pcRHCargoCodEvaluado As String, _
                                            ByVal pcRHCargoCodEvaluador As String, _
                                            ByVal pnOrden As Integer, _
                                            ByVal pbTodasAgencias As Integer, _
                                            Optional ByVal pbPeriodica As Integer = 0)

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorModificaRHEvaluacionNivelAprobacion
    sSql = "UPDATE RHEvaluacionNivelAprobacion"
    sSql = sSql & " SET nOrden=" & pnOrden & _
                    ",bTodasAgencias=" & pbTodasAgencias & _
                    ",bPeriodica=" & pbPeriodica & _
                    " WHERE  cRHCargoCodEvaluado='" & pcRHCargoCodEvaluado & "' AND cRHCargoCodEvaluador='" & pcRHCargoCodEvaluador & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorModificaRHEvaluacionNivelAprobacion:
    Err.Raise Err.Number, "ModificaRHEvaluacionNivelAprobacion", Err.Description
End Sub

Public Function EliminaRHEvaluacionNivelAprobacion(ByVal pcRHCargoCodEvaluado As String, _
                                                ByVal pcRHCargoCodEvaluador As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEvaluacionNivelAprobacion
    sSql = "DELETE FROM RHEvaluacionNivelAprobacion" '
    sSql = sSql & " WHERE  cRHCargoCodEvaluado='" & pcRHCargoCodEvaluado & "' AND cRHCargoCodEvaluador='" & pcRHCargoCodEvaluador & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEvaluacionNivelAprobacion:
    Err.Raise Err.Number, "EliminaRHEvaluacionNivelAprobacion", Err.Description
End Function

Public Function Obtener_Siguiente_CodCriterio() As Integer

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrorObtener_Siguiente_CodCriterio
    sSql = "SELECT Siguiente=ISNULL(MAX(nRHEvalCriterioCod),0)+1 FROM RHEvaluacionCriterios "
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_Siguiente_CodCriterio = 1
        Else
            Obtener_Siguiente_CodCriterio = rs("Siguiente") '+ 1
        End If
    End If
    
    oConecta.CierraConexion
    
    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Siguiente_CodCriterio:
    Err.Raise Err.Number, "Obtener_Siguiente_CodCriterio", Err.Description
End Function

Public Function Obtener_Criterios_x_Cargo(pcRHCargoCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Criterios_x_Cargo
    
    sSql = " SELECT ECC.nRHEvalCriterioCod,Factor=EC.cRHEvalCriterioDescri,'','','','','',Ponderacion=ECC.nPonderacion" & _
           " FROM RHEvaluacionCriteriosCargo ECC INNER JOIN RHEvaluacionCriterios EC ON " & _
           " ECC.nRHEvalCriterioCod = EC.nRHEvalCriterioCod WHERE ECC.cRHCargoCod='" & pcRHCargoCod & "'"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Criterios_x_Cargo = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Criterios_x_Cargo:
    Err.Raise Err.Number, "Obtener_Criterios_x_Cargo", Err.Description
End Function

Public Function Obtener_Datos_Empleado(pcPersCod As String, Optional ByVal bTodos As Boolean = True) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_Empleado
    
    If bTodos = False Then
        sSql = " SELECT Car.cRHCargoCod ,Car.cRHAgenciaCod,Cat.cRHCargoDescripcion,P.cPersNombre" & _
               " FROM Persona P INNER JOIN  RHCargos Car ON Car.cPersCod=P.cPersCod AND " & _
               " Car.dRHCargoFecha = (SELECT MAX(dRHCargoFecha) FROM RHCargos WHERE cPersCod=P.cPersCod) " & _
               " INNER JOIN RHCargosTabla Cat ON Cat.cRHCargoCod=Car.cRHCargoCod" & _
               " WHERE Car.cPersCod ='" & pcPersCod & "'"
    Else
        'sSQL = " SELECT RH.cRHCod,P.cPersNombre, RH.dIngreso,CD.dRHContratoInicio,CD.dRHContratoFin,CT.cRHCargoDescripcion,CAR.cRHCargoCod, " & _
               " s.nRHSueldoMonto,A.cAgeDescripcion,Ar.cAreaDescripcion,Meses = DateDiff(Month, RH.dIngreso, CD.dRHContratoFin)" & _
               " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod INNER JOIN RHContrato CON ON P.cPersCod=CON.cPersCod " & _
               " AND CON.cRHContratoNro=(SELECT MAX(CONVERT(int,cRHContratoNro))FROM RHContrato WHERE cPersCod=P.cPersCod) " & _
               " INNER JOIN RHContratoDet CD ON CON.cPersCod=CD.cPersCod AND CON.cRHContratoNro=CD.cRHContratoNro AND " & _
               " CD.dRHContratoFecha=(SELECT MAX(dRHContratoFecha) FROM RHContratoDet WHERE cPersCod=CON.cPersCod AND " & _
               " cRHContratoNro=CON.cRHContratoNro)INNER JOIN RHSueldo S ON S.cPersCod=P.cPersCod AND S.dRHSueldoFecha= " & _
               " (SELECT MAX(dRHSueldoFecha) FROM RHSueldo WHERE cPersCod=P.cPersCod) INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod " & _
               " AND CAR.dRHCargoFecha= (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod)" & _
               " INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod INNER JOIN Agencias A ON A.cAgeCod=CAR.cRHAgenciaCod " & _
               " INNER JOIN Areas Ar ON Ar.cAreaCod = CAR.cRHAreaCod WHERE P.cPersCod='" & pcPersCod & "'"
        
        sSql = " SELECT RH.cRHCod,P.cPersNombre, RH.dIngreso,CD.dRHContratoInicio,CD.dRHContratoFin,CT.cRHCargoDescripcion,CAR.cRHCargoCod, " & _
               " s.nRHSueldoMonto,A.cAgeDescripcion,Ar.cAreaDescripcion,Meses = DateDiff(Month, RH.dIngreso, CD.dRHContratoFin)," & _
               " UltimoInc= ISNULL( CONVERT(varchar(10),(SELECT MAX(dRHSueldoFecha) FROM RHSueldo WHERE cPersCod='" & pcPersCod & "' AND nRHSueldoMonto<s.nRHSueldoMonto) ,103) ,'') ," & _
               " SueldoAntesInc= ISNULL( (SELECT MAX(nRHSueldoMonto) FROM RHSueldo WHERE cPersCod='" & pcPersCod & "' AND nRHSueldoMonto<s.nRHSueldoMonto),0)" & _
               " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod INNER JOIN RHContrato CON ON P.cPersCod=CON.cPersCod " & _
               " AND CON.cRHContratoNro=(SELECT MAX(CONVERT(int,cRHContratoNro))FROM RHContrato WHERE cPersCod=P.cPersCod) " & _
               " INNER JOIN RHContratoDet CD ON CON.cPersCod=CD.cPersCod AND CON.cRHContratoNro=CD.cRHContratoNro AND " & _
               " CD.dRHContratoFecha=(SELECT MAX(dRHContratoFecha) FROM RHContratoDet WHERE cPersCod=CON.cPersCod AND " & _
               " cRHContratoNro=CON.cRHContratoNro)INNER JOIN RHSueldo S ON S.cPersCod=P.cPersCod AND S.dRHSueldoFecha= " & _
               " (SELECT MAX(dRHSueldoFecha) FROM RHSueldo WHERE cPersCod=P.cPersCod) INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod " & _
               " AND CAR.dRHCargoFecha= (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod)" & _
               " INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod INNER JOIN Agencias A ON A.cAgeCod=CAR.cRHAgenciaCod " & _
               " INNER JOIN Areas Ar ON Ar.cAreaCod = CAR.cRHAreaCod WHERE P.cPersCod='" & pcPersCod & "'"
               
    End If
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_Empleado = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Datos_Empleado:
    Err.Raise Err.Number, "Obtener_Datos_Empleado", Err.Description
End Function

Public Function Obtener_Datos_Calificacion(pcRHEvaCalificacionCod As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_Calificacion
    
    sSql = "SELECT Inicio=nRHEvalCalificacionRangoInicio, Fin = nRHEvalCalificacionRangoFin ,Calificacion=cRHEvalCalificacionDescripcion,Factor= nRHEvalCalificacionFactor" & _
           " FROM RHEvaluacionCalificacion Where cRHEvaCalificacionCod ='" & pcRHEvaCalificacionCod & "'"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_Calificacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Datos_Calificacion:
    Err.Raise Err.Number, "Obtener_Datos_Calificacion", Err.Description
End Function

Public Function Obtener_Nota_x_Calificacion(pnNota As Double) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Nota_x_Calificacion
    
    sSql = " SELECT cRHEvaCalificacionCod,cRHEvalCalificacionDescripcion FROM RHEvaluacionCalificacion " & _
           " WHERE " & pnNota & " BETWEEN nRHEvalCalificacionRangoInicio AND nRHEvalCalificacionRangoFin"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Nota_x_Calificacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Nota_x_Calificacion:
    Err.Raise Err.Number, "Obtener_Nota_x_Calificacion", Err.Description
End Function

Public Sub InsertaRHEvaluacionDesempeño(ByVal pcRHCod As String, ByVal pcPersCodEvaluador As String, ByVal pdRHEvalInicio As Date, _
                                        ByVal pdRHEvalFin As Date, ByVal pcRHEvalCalificacionCod As String, ByVal pnRHEvalNota As Double, _
                                        ByVal pcRHEvalFortalezas As String, ByVal pcRHEvalDebilidades As String, ByVal pbRHEvalPromocionar As Integer, _
                                        ByVal pcRHCargoCodPropuesto As String, ByVal pbRHEvalIncremento As Integer, ByVal pnRHEvalSueldoPropuesto As Double, _
                                        ByVal pbRHEvalRenovacion As Integer, ByVal pcRHEvalFechaPropuesta As String, ByVal pcRHEvalComentarioRRHH As String, _
                                        ByVal pcUltimaActualizacion As String, ByVal pnEstado As Integer)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEvaluacionDesempeño
    
    sSql = "INSERT INTO RHEvaluacionDesempeño(cRHCod, cPersCodEvaluador, dRHEvalInicio, dRHEvalFin, cRHEvalCalificacionCod, nRHEvalNota, cRHEvalFortalezas," & _
            "cRHEvalDebilidades, bRHEvalPromocionar, cRHCargoCodPropuesto,bRHEvalIncremento, nRHEvalSueldoPropuesto, bRHEvalRenovacion, cRHEvalFechaPropuesta, " & _
            "cRHEvalComentarioRRHH, cUltimaActualizacion, nEstado)"
    sSql = sSql & " VALUES('" & pcRHCod & "','" & pcPersCodEvaluador & "','" & Format(pdRHEvalInicio, gsFormatoFecha) & "','" & Format(pdRHEvalFin, gsFormatoFecha) & "','" & pcRHEvalCalificacionCod & "'," & pnRHEvalNota & ",'" & pcRHEvalFortalezas & "','" & _
             pcRHEvalDebilidades & "'," & pbRHEvalPromocionar & ",'" & pcRHCargoCodPropuesto & "'," & pbRHEvalIncremento & "," & pnRHEvalSueldoPropuesto & "," & pbRHEvalRenovacion & ",'" & pcRHEvalFechaPropuesta & "','" & _
             pcRHEvalComentarioRRHH & "','" & pcUltimaActualizacion & "'," & pnEstado & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEvaluacionDesempeño:
    Err.Raise Err.Number, "InsertaRHEvaluacionDesempeño", Err.Description
End Sub

Public Sub ModificaRHEvaluacionDesempeño(ByVal pcRHCod As String, ByVal pcRHEvalCalificacionCod As String, ByVal pnRHEvalNota As Double, _
                                        ByVal pcRHEvalFortalezas As String, ByVal pcRHEvalDebilidades As String, ByVal pbRHEvalPromocionar As Integer, _
                                        ByVal pcRHCargoCodPropuesto As String, ByVal pbRHEvalIncremento As Integer, ByVal pnRHEvalSueldoPropuesto As Double, _
                                        ByVal pbRHEvalRenovacion As Integer, ByVal pdRHEvalFechaPropuesta As Date, _
                                        ByVal pcUltimaActualizacion As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHEvaluacionDesempeño
    sSql = "UPDATE RHEvaluacionDesempeño"
    sSql = sSql & " SET cRHEvalCalificacionCod='" & pcRHEvalCalificacionCod & _
                        "',nRHEvalNota=" & pnRHEvalNota & _
                        ",cRHEvalFortalezas='" & pcRHEvalFortalezas & _
                        "',cRHEvalDebilidades='" & pcRHEvalDebilidades & _
                        "',bRHEvalPromocionar=" & pbRHEvalPromocionar & _
                        ",cRHCargoCodPropuesto='" & pcRHCargoCodPropuesto & _
                        "',bRHEvalIncremento=" & pbRHEvalIncremento & _
                        ",nRHEvalSueldoPropuesto=" & pnRHEvalSueldoPropuesto & _
                        ",bRHEvalRenovacion=" & pbRHEvalRenovacion & _
                        ",dRHEvalFechaPropuesta='" & pdRHEvalFechaPropuesta & _
                        "',cUltimaActualizacion='" & pcUltimaActualizacion & _
                        "' WHERE cRHCod='" & pcRHCod & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEvaluacionDesempeño:
    Err.Raise Err.Number, "ModificaRHEvaluacionDesempeño", Err.Description
End Sub

Public Sub ModificaRHEvaluacionDesempeño_x_RRHH(ByVal pcRHCod As String, _
                                            ByVal pcRHEvalComentarioRRHH As String, _
                                            ByVal pcUltimaActualizacion As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHEvaluacionDesempeño_x_RRHH
    sSql = "UPDATE RHEvaluacionDesempeño"
    sSql = sSql & " SET cRHEvalComentarioRRHH='" & pcRHEvalComentarioRRHH & _
                        "',cUltimaActualizacion='" & pcUltimaActualizacion & _
                        "' ,nEstado=2 WHERE cRHCod='" & pcRHCod & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEvaluacionDesempeño_x_RRHH:
    Err.Raise Err.Number, "ModificaRHEvaluacionDesempeño_x_RRHH", Err.Description
End Sub

Public Sub ModificaRHEvaluacionDesempeño_x_Decision(ByVal pcRHCod As String, _
                                            ByVal pbRHEvalRenovacionDecision As Integer, _
                                            ByVal pcRHEvalFechaDecision As String, _
                                            ByVal pcRHEvalCargoDecision As String, _
                                            ByVal pnRHEvalSueldoDecision As Double, _
                                            ByVal pcRHEvalObservacionesDecision As String, _
                                            ByVal pcUltimaActualizacion As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHEvaluacionDesempeño_x_Decision
    sSql = "UPDATE RHEvaluacionDesempeño"
    sSql = sSql & " SET bRHEvalRenovacionDecision=" & pbRHEvalRenovacionDecision & _
                        ",cRHEvalFechaDecision='" & pcRHEvalFechaDecision & _
                        "',cRHEvalCargoDecision='" & pcRHEvalCargoDecision & _
                        "',nRHEvalSueldoDecision=" & pnRHEvalSueldoDecision & _
                        ",cRHEvalObservacionesDecision='" & pcRHEvalObservacionesDecision & _
                        "',cUltimaActualizacion='" & pcUltimaActualizacion & _
                        "',nEstado=3 WHERE cRHCod='" & pcRHCod & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEvaluacionDesempeño_x_Decision:
    Err.Raise Err.Number, "ModificaRHEvaluacionDesempeño_x_Decision", Err.Description
End Sub

Public Function EliminaRHEvaluacionDesempeño(ByVal pcRHCod As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEvaluacionDesempeño
    sSql = "DELETE FROM RHEvaluacionDesempeño" '
    sSql = sSql & " WHERE cRHCod='" & pcRHCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEvaluacionDesempeño:
    Err.Raise Err.Number, "EliminaRHEvaluacionDesempeño", Err.Description
End Function

Public Sub InsertaRHEvaluacionDesempeñoAprobaciones(ByVal pcRHCod As String, ByVal pcRHCargoCodEvaluado As String, ByVal pcRHCargoCodEvaluador As String, _
                                                    ByVal pbAprobado As Integer, ByVal pnOrden As Integer, _
                                                    Optional ByVal pcPersCodEvaluador As String = "", Optional ByVal pdFecha As Date = 0)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEvaluacionDesempeñoAprobaciones
    
    If pdFecha <> 0 And pcPersCodEvaluador <> "" Then
        sSql = "INSERT INTO RHEvaluacionDesempeñoAprobaciones(cRHCod, cRHCargoCodEvaluado, cRHCargoCodEvaluador, cPersCodEvaluador, bAprobado, nOrden, dFecha)" & _
         " VALUES('" & pcRHCod & "','" & pcRHCargoCodEvaluado & "','" & pcRHCargoCodEvaluador & "','" & pcPersCodEvaluador & "'," & pbAprobado & "," & pnOrden & ",'" & pdFecha & "')"
    Else    'No Indicar Fecha ni el que Aprueba
        sSql = "INSERT INTO RHEvaluacionDesempeñoAprobaciones(cRHCod, cRHCargoCodEvaluado, cRHCargoCodEvaluador, bAprobado, nOrden)" & _
        " VALUES('" & pcRHCod & "','" & pcRHCargoCodEvaluado & "','" & pcRHCargoCodEvaluador & "'," & pbAprobado & "," & pnOrden & ")"
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEvaluacionDesempeñoAprobaciones:
    Err.Raise Err.Number, "InsertaRHEvaluacionDesempeñoAprobaciones", Err.Description
End Sub

Public Sub ModificaRHEvaluacionDesempeñoAprobaciones(ByVal pcRHCod As String, ByVal pcPersCodEvaluador As String, _
                                                    ByVal pcRHCargoCodEvaluador As String, ByVal pbAprobado As Integer, ByVal pdFecha As Date)

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrorModificaRHEvaluacionDesempeñoAprobaciones
    sSql = "UPDATE RHEvaluacionDesempeñoAprobaciones"
    sSql = sSql & " SET cPersCodEvaluador='" & pcPersCodEvaluador & _
                        "',bAprobado=" & pbAprobado & _
                        ",dFecha='" & Format(pdFecha, gsFormatoFecha) & _
                        "' WHERE cRHCod='" & pcRHCod & "' AND cRHCargoCodEvaluador='" & pcRHCargoCodEvaluador & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    
    'Actualizar el Estado de la Evaluacion
    '11-19  (HAY UN NIVEL MENOS DENTRO DE LOS REGISTROS)
    sSql = " SELECT * FROM RHEvaluacionDesempeñoAprobaciones WHERE cRHCod='" & pcRHCod & "' AND nOrden=2 AND bAprobado=1"
    'sSql = " SELECT * FROM RHEvaluacionDesempeñoAprobaciones WHERE cRHCod='" & pcRHCod & "' AND nOrden=3 AND bAprobado=1"
    Set rs = oConecta.CargaRecordSet(sSql)
    
    If Not rs.EOF Then  'Hay una Aprobacion de 1 Miembro del Comite + el Gerente de Area
        
        'sSQL = "UPDATE RHEvaluacionDesempeño SET nEstado=4 WHERE cRHCod='" & pcRHCod & "'"
        sSql = "UPDATE RHEvaluacionDesempeñoFinal SET nEstado=4 WHERE cRHCod='" & pcRHCod & "'"
        oConecta.ConexionActiva.Execute sSql
    End If
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEvaluacionDesempeñoAprobaciones:
    Err.Raise Err.Number, "ModificaRHEvaluacionDesempeñoAprobaciones", Err.Description
End Sub

Public Function EliminaRHEvaluacionDesempeñoAprobaciones(ByVal pcRHCod As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEvaluacionDesempeñoAprobaciones
    sSql = "DELETE FROM RHEvaluacionDesempeñoAprobaciones" '
    sSql = sSql & " WHERE cRHCod='" & pcRHCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEvaluacionDesempeñoAprobaciones:
    Err.Raise Err.Number, "EliminaRHEvaluacionDesempeñoAprobaciones", Err.Description
End Function

Public Function Obtener_Aprobaciones_Gerencia(pcRHCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Aprobaciones_Gerencia
    
    sSql = "SELECT P.cPersCod,P.cPersNombre,Cat.cRHCargoDescripcion,EDA.nOrden,dFecha=convert(varchar(10),EDA.dFecha,103) " & _
           " FROM RHEvaluacionDesempeñoAprobaciones EDA INNER JOIN Persona P ON P.cPersCod = EDA.cPersCodEvaluador " & _
           " INNER JOIN RHCargosTabla Cat ON Cat.cRHCargoCod=EDA.cRHCargoCodEvaluador " & _
           " WHERE cRHCod='" & pcRHCod & "' ORDER BY nOrden "
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Aprobaciones_Gerencia = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Aprobaciones_Gerencia:
    Err.Raise Err.Number, "Obtener_Aprobaciones_Gerencia", Err.Description
End Function

Public Function Buscar_PeriodoEvaluacion_Iniciado(ByVal pnPeriodo As Integer, _
                                        ByVal pnAnio As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorBuscar_PeriodoEvaluacion_Iniciado
    
    sSql = "SELECT*FROM RHEvaluacionDesempeñoPeriodo WHERE nRHEvalPeriodo=" & pnPeriodo & " AND nRHEvalAño=" & pnAnio & " AND bEstado=1 "
       
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Buscar_PeriodoEvaluacion_Iniciado = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorBuscar_PeriodoEvaluacion_Iniciado:
    Err.Raise Err.Number, "Buscar_PeriodoEvaluacion_Iniciado", Err.Description
End Function

Public Function Obtener_Empleados_Evaluacion(pcRHCargoCod As String, pcAgeCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Empleados_Evaluacion
    
    'sSQL = "SELECT P.cPersCod, RH.cRHCod, P.cPersNombre, CT.cRHCargoDescripcion, EDP.nRHEvalPeriodo,EDP.nRHEvalAño " & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod " & _
           " AND CAR.dRHCargoFecha= (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) INNER JOIN RHCargosTabla CT " & _
           " ON CAR.cRHCargoCod=CT.cRHCargoCod INNER JOIN RHEvaluacionNivelAprobacion ENA ON ENA.cRHCargoCodEvaluado=CAR.cRHCargoCod " & _
           " AND (Car.cRHAgenciaCod ='" & pcAgeCod & "' OR ENA.bTodasAgencias=1 ) AND ENA.nOrden=(SELECT MIN(nOrden)FROM RHEvaluacionNivelAprobacion WHERE cRHCargoCodEvaluado= CAR.cRHCargoCod)" & _
           " INNER JOIN RHEvaluacionDesempeñoPeriodica EDP ON EDP.cRHCod=RH.cRHCod AND ISNULL(EDP.cPersCodEvaluador,'')='' INNER JOIN RHEvaluacionDesempeñoPeriodo PERIODO " & _
           " ON EDP.nRHEvalPeriodo=PERIODO.nRHEvalPeriodo AND EDP.nRHEvalAño=PERIODO.nRHEvalAño AND PERIODO.bEstado=1 " & _
           " WHERE ENA.cRHCargoCodEvaluador='" & pcRHCargoCod & "' ORDER BY P.cPersNombre"
           
    sSql = " SELECT P.cPersCod, RH.cRHCod, P.cPersNombre, CT.cRHCargoDescripcion, EDP.nRHEvalPeriodo,EDP.nRHEvalAño " & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod " & _
           " AND CAR.dRHCargoFecha= (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) INNER JOIN RHCargosTabla CT " & _
           " ON CAR.cRHCargoCod=CT.cRHCargoCod INNER JOIN RHEvaluacionPeriodicaNivelAprobacion ENA ON ENA.cRHCargoCodEvaluado=CAR.cRHCargoCod " & _
           " AND (Car.cRHAgenciaCod ='" & pcAgeCod & "' OR ENA.bTodasAgencias=1 ) AND ENA.nOrden=(SELECT MIN(nOrden)FROM RHEvaluacionPeriodicaNivelAprobacion WHERE cRHCargoCodEvaluado= CAR.cRHCargoCod)" & _
           " INNER JOIN RHEvaluacionDesempeñoPeriodica EDP ON EDP.cRHCod=RH.cRHCod AND ISNULL(EDP.cPersCodEvaluador,'')='' INNER JOIN RHEvaluacionDesempeñoPeriodo PERIODO " & _
           " ON EDP.nRHEvalPeriodo=PERIODO.nRHEvalPeriodo AND EDP.nRHEvalAño=PERIODO.nRHEvalAño AND PERIODO.bEstado=1 " & _
           " WHERE ENA.cRHCargoCodEvaluador='" & pcRHCargoCod & "' ORDER BY P.cPersNombre"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Empleados_Evaluacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Empleados_Evaluacion:
    Err.Raise Err.Number, "Obtener_Empleados_Evaluacion", Err.Description
End Function

Public Sub ModificaRHEvaluacionDesempeñoPeriodica(ByVal pcRHCod As String, ByVal pnRHEvalPeriodo As Integer, ByVal pnRHEvalAnio As Integer, _
                                        ByVal pcPersCodEvaluador As String, ByVal pdRHEvalInicio As Date, ByVal pdRHEvalFin As Date, _
                                        ByVal pcRHEvalCalificacionCod As String, ByVal pnRHEvalNota As Double, _
                                        ByVal pcRHEvalFortalezas As String, ByVal pcRHEvalDebilidades As String, _
                                        ByVal pcUltimaActualizacion As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHEvaluacionDesempeñoPeriodica
    sSql = "UPDATE RHEvaluacionDesempeñoPeriodica"
    sSql = sSql & " SET cPersCodEvaluador='" & pcPersCodEvaluador & _
                        "',dRHEvalInicio='" & Format(pdRHEvalInicio, gsFormatoFecha) & _
                        "',dRHEvalFin='" & Format(pdRHEvalFin, gsFormatoFecha) & _
                        "',cRHEvalCalificacionCod='" & pcRHEvalCalificacionCod & _
                        "',nRHEvalNota=" & pnRHEvalNota & _
                        ",cRHEvalFortalezas='" & pcRHEvalFortalezas & _
                        "',cRHEvalDebilidades='" & pcRHEvalDebilidades & _
                        "',cUltimaActualizacion='" & pcUltimaActualizacion & _
                        "' WHERE cRHCod='" & pcRHCod & "' AND nRHEvalPeriodo=" & pnRHEvalPeriodo & _
                        " AND nRHEvalAño=" & pnRHEvalAnio

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEvaluacionDesempeñoPeriodica:
    Err.Raise Err.Number, "ModificaRHEvaluacionDesempeñoPeriodica", Err.Description
End Sub

Public Function Obtener_Datos_EvaluacionPeriodica(pcPersCod As String, pnAnio As Integer, pnPeriodo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_EvaluacionPeriodica
    
    sSql = " SELECT RH.cRHCod,P.cPersNombre, RH.dIngreso,CT.cRHCargoDescripcion,CAR.cRHCargoCod, s.nRHSueldoMonto,A.cAgeDescripcion, " & _
           " Ar.cAreaDescripcion,Meses = DateDiff(Month, RH.dIngreso, getdate()),Inicio= CASE WHEN EDP.cRHCod IS NULL THEN CONVERT(varchar(10),CD.dRHContratoInicio,103) " & _
           " ELSE dbo.getFechamov(EDP.cUltimaActualizacion,103) END,Fin= CONVERT(varchar(10),getdate(),103) " & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod INNER JOIN RHContrato CON ON P.cPersCod=CON.cPersCod AND " & _
           " CON.cRHContratoNro=(SELECT MAX(CONVERT(int,cRHContratoNro))FROM RHContrato WHERE cPersCod=P.cPersCod) INNER JOIN RHContratoDet CD ON " & _
           " CON.cPersCod=CD.cPersCod AND CON.cRHContratoNro=CD.cRHContratoNro AND CD.dRHContratoFecha=(SELECT MAX(dRHContratoFecha) FROM RHContratoDet " & _
           " WHERE cPersCod=CON.cPersCod AND cRHContratoNro=CON.cRHContratoNro)INNER JOIN RHSueldo S ON S.cPersCod=P.cPersCod AND S.dRHSueldoFecha= " & _
           " (SELECT MAX(dRHSueldoFecha) FROM RHSueldo WHERE cPersCod=P.cPersCod) INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND CAR.dRHCargoFecha= " & _
           " (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod INNER JOIN Agencias A " & _
           " ON A.cAgeCod=CAR.cRHAgenciaCod INNER JOIN Areas Ar ON Ar.cAreaCod = CAR.cRHAreaCod LEFT JOIN RHEvaluacionDesempeñoPeriodica EDP ON RH.cRHCod=EDP.cRHCod " & _
           " AND EDP.nRHEvalAño=" & pnAnio & " AND EDP.nRHEvalPeriodo=(SELECT MAX(nRHEvalPeriodo)FROM RHEvaluacionDesempeñoPeriodica WHERE cRHCod=RH.cRHCod " & _
           " AND nRHEvalPeriodo<" & pnPeriodo & " AND nRHEvalAño=" & pnAnio & " AND cPersCodEvaluador IS NOT NULL) WHERE P.cPersCod='" & pcPersCod & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_EvaluacionPeriodica = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Datos_EvaluacionPeriodica:
    Err.Raise Err.Number, "Obtener_Datos_EvaluacionPeriodica", Err.Description
End Function

Public Sub InsertaRHEvaluacionPeriodicaCriterios(ByVal pnRHEvalCriterioCod As Integer, _
                                                ByVal pcRHCod As String, _
                                                ByVal pnRHEvalPeriodo As Integer, _
                                                ByVal pnRHEvalAnio As Integer, _
                                                ByVal pcRHEvalCalificacionCod As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEvaluacionPeriodicaCriterios
    
    sSql = "INSERT INTO RHEvaluacionPeriodicaCriterios(nRHEvalCriterioCod,cRHCod,nRHEvalPeriodo,nRHEvalAño,cRHEvalCalificacionCod)"
    sSql = sSql & " VALUES(" & pnRHEvalCriterioCod & ",'" & pcRHCod & "'," & pnRHEvalPeriodo & "," & pnRHEvalAnio & ",'" & pcRHEvalCalificacionCod & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEvaluacionPeriodicaCriterios:
    Err.Raise Err.Number, "InsertaRHEvaluacionPeriodicaCriterios", Err.Description
End Sub

Public Function EliminaRHEvaluacionPeriodicaCriterios(ByVal pcRHCod As String, _
                                                    ByVal pnRHEvalPeriodo As Integer, _
                                                    ByVal pnRHEvalAnio As Integer)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEvaluacionPeriodicaCriterios
    sSql = "DELETE FROM RHEvaluacionPeriodicaCriterios" '
    sSql = sSql & " WHERE cRHCod='" & pcRHCod & "' AND nRHEvalPeriodo=" & pnRHEvalPeriodo & _
            " AND nRHEvalAño=" & pnRHEvalAnio
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEvaluacionPeriodicaCriterios:
    Err.Raise Err.Number, "EliminaRHEvaluacionPeriodicaCriterios", Err.Description
End Function

Public Sub InsertaRHEvaluacionDesempeñoCriterios(ByVal pnRHEvalCriterioCod As Integer, _
                                                ByVal pcRHCod As String, _
                                                ByVal pcRHEvalCalificacionCod As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEvaluacionDesempeñoCriterios
    
    sSql = "INSERT INTO RHEvaluacionDesempeñoCriterios(nRHEvalCriterioCod,cRHCod,cRHEvalCalificacionCod)"
    sSql = sSql & " VALUES(" & pnRHEvalCriterioCod & ",'" & pcRHCod & "','" & pcRHEvalCalificacionCod & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEvaluacionDesempeñoCriterios:
    Err.Raise Err.Number, "InsertaRHEvaluacionDesempeñoCriterios", Err.Description
End Sub

Public Function EliminaRHEvaluacionDesempeñoCriterios(ByVal pcRHCod As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEvaluacionDesempeñoCriterios
    sSql = "DELETE FROM RHEvaluacionDesempeñoCriterios" '
    sSql = sSql & " WHERE cRHCod='" & pcRHCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEvaluacionDesempeñoCriterios:
    Err.Raise Err.Number, "EliminaRHEvaluacionDesempeñoCriterios", Err.Description
End Function

Public Sub InsertaRHEvaluacionPeriodicaCapacitacion(ByVal pcRHCod As String, _
                                                ByVal pnRHEvalPeriodo As Integer, _
                                                ByVal pnRHEvalAnio As Integer, _
                                                ByVal pnCapacitacionTemaCod As Integer, _
                                                ByVal pcComentario As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEvaluacionPeriodicaCapacitacion
    
    sSql = "INSERT INTO RHEvaluacionPeriodicaCapacitacion(cRHCod,nRHEvalPeriodo,nRHEvalAño,nCapacitacionTemaCod,cComentario)"
    sSql = sSql & " VALUES('" & pcRHCod & "'," & pnRHEvalPeriodo & "," & pnRHEvalAnio & "," & pnCapacitacionTemaCod & ",'" & pcComentario & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEvaluacionPeriodicaCapacitacion:
    Err.Raise Err.Number, "InsertaRHEvaluacionPeriodicaCapacitacion", Err.Description
End Sub

Public Function EliminaRHEvaluacionPeriodicaCapacitacion(ByVal pcRHCod As String, _
                                                    ByVal pnRHEvalPeriodo As Integer, _
                                                    ByVal pnRHEvalAnio As Integer)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEvaluacionPeriodicaCapacitacion
    sSql = "DELETE FROM RHEvaluacionPeriodicaCapacitacion" '
    sSql = sSql & " WHERE cRHCod='" & pcRHCod & "' AND nRHEvalPeriodo=" & pnRHEvalPeriodo & _
            " AND nRHEvalAño=" & pnRHEvalAnio
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEvaluacionPeriodicaCapacitacion:
    Err.Raise Err.Number, "EliminaRHEvaluacionPeriodicaCapacitacion", Err.Description
End Function

Public Sub InsertaRHEvaluacionCapacitacion(ByVal pcRHCod As String, _
                                          ByVal pnCapacitacionTemaCod As Integer, _
                                          ByVal pcComentario As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEvaluacionCapacitacion
    
    sSql = "INSERT INTO RHEvaluacionCapacitacion(cRHCod,nCapacitacionTemaCod,cComentario)"
    sSql = sSql & " VALUES('" & pcRHCod & "'," & pnCapacitacionTemaCod & ",'" & pcComentario & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEvaluacionCapacitacion:
    Err.Raise Err.Number, "InsertaRHEvaluacionCapacitacion", Err.Description
End Sub

Public Function EliminaRHEvaluacionCapacitacion(ByVal pcRHCod As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEvaluacionCapacitacion
    sSql = "DELETE FROM RHEvaluacionCapacitacion" '
    sSql = sSql & " WHERE cRHCod='" & pcRHCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEvaluacionCapacitacion:
    Err.Raise Err.Number, "EliminaRHEvaluacionCapacitacion", Err.Description
End Function

Public Function Obtener_Empleados_EvaluacionFinal(pcRHCargoCod As String, pcAgeCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Empleados_EvaluacionFinal
    
    'sSQL = "SELECT P.cPersCod, RH.cRHCod, P.cPersNombre, CT.cRHCargoDescripcion,CAR.cRHCargoCod" & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod AND RH.nRHEstado=201 AND RH.cRHCod NOT IN (SELECT cRHCod FROM RHEvaluacionDesempeño)" & _
           " INNER JOIN RHContrato CON ON P.cPersCod=CON.cPersCod AND CON.cRHContratoNro= " & _
           " (SELECT MAX(CONVERT(int,cRHContratoNro))FROM RHContrato WHERE cPersCod=P.cPersCod) " & _
           " INNER JOIN RHContratoDet CD ON CON.cPersCod=CD.cPersCod AND CON.cRHContratoNro=CD.cRHContratoNro AND " & _
           " CD.dRHContratoFecha=(SELECT MAX(dRHContratoFecha) FROM RHContratoDet WHERE cPersCod=CON.cPersCod AND " & _
           " cRHContratoNro=CON.cRHContratoNro)INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND CAR.dRHCargoFecha= " & _
           " (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod " & _
           " INNER JOIN RHEvaluacionNivelAprobacion ENA ON ENA.cRHCargoCodEvaluado=CAR.cRHCargoCod AND " & _
           " (Car.cRHAgenciaCod ='" & pcAgeCod & "' OR ENA.bTodasAgencias=1 ) AND ENA.nOrden= (SELECT MIN(nOrden)FROM RHEvaluacionNivelAprobacion WHERE cRHCargoCodEvaluado= CAR.cRHCargoCod)" & _
           " WHERE ENA.cRHCargoCodEvaluador='" & pcRHCargoCod & "' AND MONTH(CD.dRHContratoFin)=MONTH(getdate()) AND YEAR(CD.dRHContratoFin)=YEAR(getdate()) " & _
           " ORDER BY P.cPersNombre "
     sSql = "SELECT P.cPersCod, RH.cRHCod, P.cPersNombre, CT.cRHCargoDescripcion,CAR.cRHCargoCod" & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod AND RH.nRHEstado=201 AND RH.cRHCod NOT IN (SELECT cRHCod FROM RHEvaluacionDesempeñoFinal)" & _
           " INNER JOIN RHContrato CON ON P.cPersCod=CON.cPersCod AND CON.cRHContratoNro= " & _
           " (SELECT MAX(CONVERT(int,cRHContratoNro))FROM RHContrato WHERE cPersCod=P.cPersCod) " & _
           " INNER JOIN RHContratoDet CD ON CON.cPersCod=CD.cPersCod AND CON.cRHContratoNro=CD.cRHContratoNro AND " & _
           " CD.dRHContratoFecha=(SELECT MAX(dRHContratoFecha) FROM RHContratoDet WHERE cPersCod=CON.cPersCod AND " & _
           " cRHContratoNro=CON.cRHContratoNro)INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND CAR.dRHCargoFecha= " & _
           " (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod " & _
           " INNER JOIN RHEvaluacionNivelAprobacion ENA ON ENA.cRHCargoCodEvaluado=CAR.cRHCargoCod AND " & _
           " (Car.cRHAgenciaCod ='" & pcAgeCod & "' OR ENA.bTodasAgencias=1 ) AND ENA.nOrden= (SELECT MIN(nOrden)FROM RHEvaluacionNivelAprobacion WHERE cRHCargoCodEvaluado= CAR.cRHCargoCod)" & _
           " WHERE ENA.cRHCargoCodEvaluador='" & pcRHCargoCod & "' AND MONTH(CD.dRHContratoFin)=MONTH(getdate()) AND YEAR(CD.dRHContratoFin)=YEAR(getdate()) " & _
           " ORDER BY P.cPersNombre "
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Empleados_EvaluacionFinal = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Empleados_EvaluacionFinal:
    Err.Raise Err.Number, "Obtener_Empleados_EvaluacionFinal", Err.Description
End Function

Public Function Obtener_Empleados_EvaluacionFinal_ComentarioRRHH() As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Empleados_EvaluacionFinal_ComentarioRRHH
    
    'sSQL = " SELECT P.cPersCod, RH.cRHCod, P.cPersNombre, CT.cRHCargoDescripcion,CAR.cRHCargoCod " & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod AND RH.nRHEstado=201" & _
           " INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND CAR.dRHCargoFecha=" & _
           " (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) " & _
           " INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod " & _
           " INNER JOIN RHEvaluacionDesempeño EVA ON RH.cRHCod=EVA.cRHCod AND nEstado=1 " & _
           " ORDER BY P.cPersNombre"
    sSql = " SELECT P.cPersCod, RH.cRHCod, P.cPersNombre, CT.cRHCargoDescripcion,CAR.cRHCargoCod " & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod AND RH.nRHEstado=201" & _
           " INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND CAR.dRHCargoFecha=" & _
           " (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) " & _
           " INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod " & _
           " INNER JOIN RHEvaluacionDesempeñoFinal EVA ON RH.cRHCod=EVA.cRHCod AND nEstado=3 AND ISNULL(cRHEvalComentarioRRHH,'')=''" & _
           " ORDER BY P.cPersNombre"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Empleados_EvaluacionFinal_ComentarioRRHH = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Empleados_EvaluacionFinal_ComentarioRRHH:
    Err.Raise Err.Number, "Obtener_Empleados_EvaluacionFinal_ComentarioRRHH", Err.Description
End Function

Public Function Obtener_Empleados_EvaluacionFinal_Decision(pcRHCargoCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Empleados_EvaluacionFinal_Decision
    
    'sSQL = "SELECT P.cPersCod, RH.cRHCod, P.cPersNombre, CT.cRHCargoDescripcion,CAR.cRHCargoCod,Estado=EVA.nEstado" & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod AND RH.nRHEstado=201 " & _
           " INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND CAR.dRHCargoFecha= " & _
           " (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod)" & _
           " INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod " & _
           " INNER JOIN RHEvaluacionDesempeño EVA ON RH.cRHCod=EVA.cRHCod " & _
           " INNER JOIN RHEvaluacionNivelAprobacion ENA ON ENA.cRHCargoCodEvaluado=CAR.cRHCargoCod " & _
           " INNER JOIN RHEvaluacionDesempeñoAprobaciones EDA ON EDA.cRHCod = EVA.cRHCod " & _
           " AND EDA.cRHCargoCodEvaluador = ENA.cRHCargoCodEvaluador AND EDA.bAprobado=0 AND(( EDA.nOrden=2 AND EVA.nEstado=2)OR(EDA.nOrden=3 AND EVA.nEstado IN (3,4)))" & _
           " WHERE ENA.cRHCargoCodEvaluador='" & pcRHCargoCod & "'" & _
           " ORDER BY P.cPersNombre"
    sSql = "SELECT P.cPersCod, RH.cRHCod, P.cPersNombre, CT.cRHCargoDescripcion,CAR.cRHCargoCod,Estado=EVA.nEstado" & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod AND RH.nRHEstado=201 " & _
           " INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND CAR.dRHCargoFecha= " & _
           " (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod)" & _
           " INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod " & _
           " INNER JOIN RHEvaluacionDesempeñoFinal EVA ON RH.cRHCod=EVA.cRHCod " & _
           " INNER JOIN RHEvaluacionNivelAprobacion ENA ON ENA.cRHCargoCodEvaluado=CAR.cRHCargoCod " & _
           " INNER JOIN RHEvaluacionDesempeñoAprobaciones EDA ON EDA.cRHCod = EVA.cRHCod " & _
           " AND EDA.cRHCargoCodEvaluador = ENA.cRHCargoCodEvaluador AND EDA.bAprobado=0 AND(( EDA.nOrden=2 AND EVA.nEstado=2)OR(EDA.nOrden=2 AND EVA.nEstado IN (3,4)))" & _
           " WHERE ENA.cRHCargoCodEvaluador='" & pcRHCargoCod & "'" & _
           " ORDER BY P.cPersNombre"
        'EDA.Orden=2        'Ya no va, el Estado 2 ya no existe     19-11
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Empleados_EvaluacionFinal_Decision = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Empleados_EvaluacionFinal_Decision:
    Err.Raise Err.Number, "Obtener_Empleados_EvaluacionFinal_Decision", Err.Description
End Function

Public Sub InsertaRHEvaluacionDesempeñoAprobaciones_x_Evaluacion(ByVal pcRHCod As String, ByVal pcRHCargoCod As String)

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorInsertaRHEvaluacionDesempeñoAprobaciones_x_Evaluacion
    
    sSql = " SELECT cRHCargoCodEvaluador, nOrden " & _
           " FROM RHEvaluacionNivelAprobacion WHERE cRHCargoCodEvaluado='" & pcRHCargoCod & "'"
    
    Set oConecta = New DConecta
    
    oConecta.AbreConexion
    
    Set rs = oConecta.CargaRecordSet(sSql)
    While Not rs.EOF
        Call InsertaRHEvaluacionDesempeñoAprobaciones(pcRHCod, pcRHCargoCod, rs("cRHCargoCodEvaluador"), 0, rs("nOrden"))
        rs.MoveNext
    Wend

    'oConecta.ConexionActiva.Execute sSQL
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEvaluacionDesempeñoAprobaciones_x_Evaluacion:
    Err.Raise Err.Number, "InsertaRHEvaluacionDesempeñoAprobaciones_x_Evaluacion", Err.Description
End Sub

Public Function Obtener_CriteriosEvaluados_x_Evaluacion(ByVal pcRHCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_CriteriosEvaluados_x_Evaluacion
    
    sSql = "SELECT EC.nRHEvalCriterioCod,EC.cRHEvalCriterioDescri, cRHEvalCalificacionCod " & _
           " FROM RHEvaluacionDesempeñoCriterios EDC INNER JOIN RHEvaluacionCriterios EC " & _
           " ON EDC.nRHEvalCriterioCod= EC.nRHEvalCriterioCod WHERE EDC.cRHCod='" & pcRHCod & "'" & _
           " ORDER BY EC.nRHEvalCriterioCod"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_CriteriosEvaluados_x_Evaluacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_CriteriosEvaluados_x_Evaluacion:
    Err.Raise Err.Number, "Obtener_CriteriosEvaluados_x_Evaluacion", Err.Description
End Function

Public Function Obtener_Capacitaciones_x_Evaluacion(ByVal pcRHCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Capacitaciones_x_Evaluacion
    
    sSql = " SELECT EC.nCapacitacionTemaCod,C.cConsDescripcion,EC.cComentario " & _
           " FROM RHEvaluacionCapacitacion EC INNER JOIN Constante C ON EC.nCapacitacionTemaCod = c.nConsValor And c.nConsCod = 9074 " & _
           " WHERE EC.cRHCod='" & pcRHCod & "'"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Capacitaciones_x_Evaluacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Capacitaciones_x_Evaluacion:
    Err.Raise Err.Number, "Obtener_Capacitaciones_x_Evaluacion", Err.Description
End Function

Public Function Obtener_Datos_x_EvaluacionDesempeño(ByVal pcRHCod As String, ByVal pnTipoOperacion As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_x_EvaluacionDesempeño
    
    Select Case pnTipoOperacion
        Case 0  'Solo Datos Evaluacion
            'sSQL = "SELECT cRHCod, cPersCodEvaluador, dRHEvalInicio, dRHEvalFin, cRHEvalCalificacionCod, nRHEvalNota, cRHEvalFortalezas, " & _
                   " cRHEvalDebilidades,bRHEvalPromocionar=CONVERT(int,bRHEvalPromocionar),cRHCargoCodPropuesto,Cargo=ISNULL(Car.cRHCargoDescripcion,''), bRHEvalIncremento=CONVERT(int,bRHEvalIncremento), nRHEvalSueldoPropuesto,bRHEvalRenovacion=CONVERT(int,bRHEvalRenovacion), cRHEvalFechaPropuesta " & _
                   " FROM RHEvaluacionDesempeño LEFT JOIN RHCargosTabla Car ON Car.cRHCargoCod=cRHCargoCodPropuesto " & _
                   " WHERE cRHCod='" & pcRHCod & "'"
            sSql = "SELECT cRHCod, cPersCodEvaluador, dRHEvalInicio, dRHEvalFin, " & _
                   " bRHEvalPromocionar=CONVERT(int,bRHEvalPromocionar),cRHCargoCodPropuesto,Cargo=ISNULL(Car.cRHCargoDescripcion,''), bRHEvalIncremento=CONVERT(int,bRHEvalIncremento), nRHEvalSueldoPropuesto,bRHEvalRenovacion=CONVERT(int,bRHEvalRenovacion), cRHEvalFechaPropuesta " & _
                   " FROM RHEvaluacionDesempeñoFinal LEFT JOIN RHCargosTabla Car ON Car.cRHCargoCod=cRHCargoCodPropuesto " & _
                   " WHERE cRHCod='" & pcRHCod & "'"
        Case 1  'Comentario de Recursos
            'sSQL = "SELECT cRHCod, cRHEvalComentarioRRHH " & _
                   " FROM RHEvaluacionDesempeño WHERE cRHCod='" & pcRHCod & "'"
            sSql = "SELECT cRHCod, cRHEvalComentarioRRHH " & _
                   " FROM RHEvaluacionDesempeñoFinal WHERE cRHCod='" & pcRHCod & "'"
        Case 2  'Recomendaciones
            'sSQL = "SELECT cRHCod, bRHEvalRenovacionDecision=CONVERT(int,bRHEvalRenovacionDecision), cRHEvalFechaDecision, cRHEvalCargoDecision,Cargo=ISNULL(Car.cRHCargoDescripcion,''), nRHEvalSueldoDecision , cRHEvalObservacionesDecision, ED.cUltimaActualizacion " & _
                   " FROM RHEvaluacionDesempeño ED INNER JOIN RHCargosTabla Car ON Car.cRHCargoCod=cRHEvalCargoDecision WHERE cRHCod='" & pcRHCod & "'"
            sSql = "SELECT cRHCod, bRHEvalRenovacionDecision=CONVERT(int,bRHEvalRenovacionDecision), cRHEvalFechaDecision, cRHEvalCargoDecision,Cargo=ISNULL(Car.cRHCargoDescripcion,''), nRHEvalSueldoDecision , cRHEvalObservacionesDecision, ED.cUltimaActualizacion " & _
                   " FROM RHEvaluacionDesempeñoFinal ED LEFT JOIN RHCargosTabla Car ON Car.cRHCargoCod=cRHEvalCargoDecision WHERE cRHCod='" & pcRHCod & "'"
    End Select
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_x_EvaluacionDesempeño = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Datos_x_EvaluacionDesempeño:
    Err.Raise Err.Number, "Obtener_Datos_x_EvaluacionDesempeño", Err.Description
End Function

Public Function Obtener_DatosEvaluador_x_Evaluacion(ByVal pcRHCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_DatosEvaluador_x_Evaluacion
    
    'sSQL = " SELECT P.cPersNombre,Cat.cRHCargoDescripcion,Fecha=dbo.getFechaMov(ED.cUltimaActualizacion,103)" & _
           " FROM RHEvaluacionDesempeño ED INNER JOIN Persona P ON ED.cPersCodEvaluador= P.cPersCod " & _
           " INNER JOIN RRHH RH ON RH.cPersCod = P.cPersCod INNER JOIN RHCargos Car ON Car.cPersCod=P.cPersCod AND " & _
           " Car.dRHCargoFecha = (SELECT MAX(dRHCargoFecha) FROM RHCargos WHERE cPersCod=P.cPersCod) " & _
           " INNER JOIN RHCargosTabla Cat ON Cat.cRHCargoCod=Car.cRHCargoCod " & _
           " WHERE ED.cRHCod='" & pcRHCod & "'"
    sSql = " SELECT P.cPersNombre,Cat.cRHCargoDescripcion,Fecha=dbo.getFechaMov(ED.cUltimaActualizacion,103)" & _
           " FROM RHEvaluacionDesempeñoFinal ED INNER JOIN Persona P ON ED.cPersCodEvaluador= P.cPersCod " & _
           " INNER JOIN RRHH RH ON RH.cPersCod = P.cPersCod INNER JOIN RHCargos Car ON Car.cPersCod=P.cPersCod AND " & _
           " Car.dRHCargoFecha = (SELECT MAX(dRHCargoFecha) FROM RHCargos WHERE cPersCod=P.cPersCod) " & _
           " INNER JOIN RHCargosTabla Cat ON Cat.cRHCargoCod=Car.cRHCargoCod " & _
           " WHERE ED.cRHCod='" & pcRHCod & "'"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_DatosEvaluador_x_Evaluacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_DatosEvaluador_x_Evaluacion:
    Err.Raise Err.Number, "Obtener_DatosEvaluador_x_Evaluacion", Err.Description
End Function

Public Function Obtener_Empleados_EvaluacionHistorico() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Empleados_EvaluacionHistorico
    
    sSql = " SELECT P.cPersCod, RH.cRHCod, P.cPersNombre, CT.cRHCargoDescripcion " & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod AND nRHEstado=201 AND cRHCod LIKE 'E%' " & _
           " INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND CAR.dRHCargoFecha= " & _
           " (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod " & _
           " ORDER BY P.cPersNombre"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Empleados_EvaluacionHistorico = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Empleados_EvaluacionHistorico:
    Err.Raise Err.Number, "Obtener_Empleados_EvaluacionHistorico", Err.Description
End Function

Public Function Obtener_EvaluacionPeriodicas_x_Empleado(ByVal pcRHCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_EvaluacionPeriodicas_x_Empleado
    
    sSql = " SELECT EDP.cRHCod, EDP.nRHEvalPeriodo,EDP.nRHEvalAño,EDP.nRHEvalNota,P.cPersNombre,CT.cRHCargoDescripcion ,E.cPersCod" & _
           " FROM RHEvaluacionDesempeñoPeriodica EDP INNER JOIN Persona P ON EDP.cPersCodEvaluador=P.cPersCod " & _
           " INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND CAR.dRHCargoFecha= " & _
           " (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) " & _
           " INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod " & _
           " INNER JOIN RRHH E ON EDP.cRHCod=E.cRHCod" & _
           " WHERE EDP.cRHCod='" & pcRHCod & "'"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set Obtener_EvaluacionPeriodicas_x_Empleado = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_EvaluacionPeriodicas_x_Empleado:
    Err.Raise Err.Number, "Obtener_EvaluacionPeriodicas_x_Empleado", Err.Description
End Function

Public Function Obtener_EvaluacionPeriodica(pcRHCod As String, pnAnio As Integer, pnPeriodo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_EvaluacionPeriodica
    
    sSql = "SELECT nRHEvalNota,cRHEvalFortalezas,cRHEvalDebilidades " & _
           "FROM RHEvaluacionDesempeñoPeriodica WHERE nRHEvalPeriodo=" & pnPeriodo & " AND nRHEvalAño=" & pnAnio & " AND cRHCod='" & pcRHCod & " '"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_EvaluacionPeriodica = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_EvaluacionPeriodica:
    Err.Raise Err.Number, "Obtener_EvaluacionPeriodica", Err.Description
End Function

Public Function Obtener_Capacitaciones_x_EvaluacionPeriodica(ByVal pcRHCod As String, pnAnio As Integer, pnPeriodo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Capacitaciones_x_EvaluacionPeriodica
    
    'sSQL = " SELECT EC.nCapacitacionTemaCod,C.cConsDescripcion,EC.cComentario " & _
           " FROM RHEvaluacionPeriodicaCapacitacion EC INNER JOIN Constante C ON EC.nCapacitacionTemaCod = c.nConsValor And c.nConsCod = 9074 " & _
           " WHERE EC.cRHCod='" & pcRHCod & "' AND EC.nRHEvalPeriodo=" & pnPeriodo & " AND EC.nRHEvalAño=" & pnAnio
    
    sSql = " SELECT EC.nCapacitacionTemaCod,C.cCursoDescripcion,EC.cComentario " & _
           " FROM RHEvaluacionPeriodicaCapacitacion EC INNER JOIN RHCapacitacionCursos C ON EC.nCapacitacionTemaCod = c.nCursoCod" & _
           " WHERE EC.cRHCod='" & pcRHCod & "' AND EC.nRHEvalPeriodo=" & pnPeriodo & " AND EC.nRHEvalAño=" & pnAnio
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Capacitaciones_x_EvaluacionPeriodica = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Capacitaciones_x_EvaluacionPeriodica:
    Err.Raise Err.Number, "Obtener_Capacitaciones_x_EvaluacionPeriodica", Err.Description
End Function

Public Function Obtener_CriteriosEvaluados_x_EvaluacionPeriodica(ByVal pcRHCod As String, pnAnio As Integer, pnPeriodo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_CriteriosEvaluados_x_EvaluacionPeriodica
    
    sSql = " SELECT EC.nRHEvalCriterioCod,EC.cRHEvalCriterioDescri, cRHEvalCalificacionCod " & _
           " FROM RHEvaluacionPeriodicaCriterios EDC INNER JOIN RHEvaluacionCriterios EC " & _
           " ON EDC.nRHEvalCriterioCod= EC.nRHEvalCriterioCod " & _
           " WHERE EDC.cRHCod='" & pcRHCod & "' AND EDC.nRHEvalPeriodo=" & pnPeriodo & " AND EDC.nRHEvalAño=" & pnAnio & _
           " ORDER BY EC.nRHEvalCriterioCod "
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_CriteriosEvaluados_x_EvaluacionPeriodica = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_CriteriosEvaluados_x_EvaluacionPeriodica:
    Err.Raise Err.Number, "Obtener_CriteriosEvaluados_x_EvaluacionPeriodica", Err.Description
End Function

Public Function Obtener_Empleados_EvaluacionFinal_Consulta() As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Empleados_EvaluacionFinal_Consulta
    
    'sSQL = " SELECT P.cPersCod, RH.cRHCod, P.cPersNombre, CT.cRHCargoDescripcion,CAR.cRHCargoCod,Estado=EVA.nEstado" & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod AND RH.nRHEstado=201 " & _
           " INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND CAR.dRHCargoFecha= " & _
           " (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) " & _
           " INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod " & _
           " INNER JOIN RHEvaluacionDesempeño EVA ON RH.cRHCod=EVA.cRHCod AND EVA.nEstado IN(3,4) " & _
           " ORDER BY P.cPersNombre "
    sSql = " SELECT P.cPersCod, RH.cRHCod, P.cPersNombre, CT.cRHCargoDescripcion,CAR.cRHCargoCod,Estado=EVA.nEstado" & _
           " FROM RRHH RH INNER JOIN Persona P ON RH.cPersCod=P.cPersCod AND RH.nRHEstado=201 " & _
           " INNER JOIN RHCargos CAR ON CAR.cPersCod=P.cPersCod AND CAR.dRHCargoFecha= " & _
           " (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) " & _
           " INNER JOIN RHCargosTabla CT ON CAR.cRHCargoCod=CT.cRHCargoCod " & _
           " INNER JOIN RHEvaluacionDesempeñoFinal EVA ON RH.cRHCod=EVA.cRHCod AND EVA.nEstado IN(3,4) " & _
           " ORDER BY P.cPersNombre "
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Empleados_EvaluacionFinal_Consulta = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Empleados_EvaluacionFinal_Consulta:
    Err.Raise Err.Number, "Obtener_Empleados_EvaluacionFinal_Consulta", Err.Description
End Function

Public Sub InsertaRHEvaluacionPeriodicaNivelAprobacion(ByVal pcRHCargoCodEvaluado As String, _
                                    ByVal pcRHCargoCodEvaluador As String, _
                                    ByVal pnOrden As Integer, _
                                    ByVal pbTodasAgencias As Integer)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEvaluacionPeriodicaNivelAprobacion
    
    sSql = "INSERT INTO RHEvaluacionPeriodicaNivelAprobacion(cRHCargoCodEvaluado,cRHCargoCodEvaluador,nOrden,bTodasAgencias)"
    sSql = sSql & " VALUES('" & pcRHCargoCodEvaluado & "','" & pcRHCargoCodEvaluador & "'," & pnOrden & "," & pbTodasAgencias & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEvaluacionPeriodicaNivelAprobacion:
    Err.Raise Err.Number, "InsertaRHEvaluacionPeriodicaNivelAprobacion", Err.Description
End Sub


Public Function EliminaRHEvaluacionPeriodicaNivelAprobacion(ByVal pcRHCargoCodEvaluado As String, _
                                                ByVal pcRHCargoCodEvaluador As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEvaluacionPeriodicaNivelAprobacion
    sSql = "DELETE FROM RHEvaluacionPeriodicaNivelAprobacion" '
    sSql = sSql & " WHERE  cRHCargoCodEvaluado='" & pcRHCargoCodEvaluado & "' AND cRHCargoCodEvaluador='" & pcRHCargoCodEvaluador & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEvaluacionPeriodicaNivelAprobacion:
    Err.Raise Err.Number, "EliminaRHEvaluacionPeriodicaNivelAprobacion", Err.Description
End Function

Public Sub InsertaRHEvaluacionDesempeñoFinal(ByVal pcRHCod As String, ByVal pcPersCodEvaluador As String, ByVal pdRHEvalInicio As Date, _
                                        ByVal pdRHEvalFin As Date, ByVal pbRHEvalPromocionar As Integer, _
                                        ByVal pcRHCargoCodPropuesto As String, ByVal pbRHEvalIncremento As Integer, ByVal pnRHEvalSueldoPropuesto As Double, _
                                        ByVal pbRHEvalRenovacion As Integer, ByVal pcRHEvalFechaPropuesta As String, ByVal pcRHEvalComentarioRRHH As String, _
                                        ByVal pcUltimaActualizacion As String, ByVal pnEstado As Integer)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHEvaluacionDesempeñoFinal
    
    sSql = "INSERT INTO RHEvaluacionDesempeñoFinal(cRHCod, cPersCodEvaluador, dRHEvalInicio, dRHEvalFin, " & _
            "bRHEvalPromocionar, cRHCargoCodPropuesto,bRHEvalIncremento, nRHEvalSueldoPropuesto, bRHEvalRenovacion, cRHEvalFechaPropuesta, " & _
            "cRHEvalComentarioRRHH, cUltimaActualizacion, nEstado)"
    sSql = sSql & " VALUES('" & pcRHCod & "','" & pcPersCodEvaluador & "','" & Format(pdRHEvalInicio, gsFormatoFecha) & "','" & Format(pdRHEvalFin, gsFormatoFecha) & "'," & _
             pbRHEvalPromocionar & ",'" & pcRHCargoCodPropuesto & "'," & pbRHEvalIncremento & "," & pnRHEvalSueldoPropuesto & "," & pbRHEvalRenovacion & ",'" & pcRHEvalFechaPropuesta & "','" & _
             pcRHEvalComentarioRRHH & "','" & pcUltimaActualizacion & "'," & pnEstado & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHEvaluacionDesempeñoFinal:
    Err.Raise Err.Number, "InsertaRHEvaluacionDesempeñoFinal", Err.Description
End Sub

Public Sub ModificaRHEvaluacionDesempeñoFinal(ByVal pcRHCod As String, ByVal pbRHEvalPromocionar As Integer, _
                                        ByVal pcRHCargoCodPropuesto As String, ByVal pbRHEvalIncremento As Integer, ByVal pnRHEvalSueldoPropuesto As Double, _
                                        ByVal pbRHEvalRenovacion As Integer, ByVal pdRHEvalFechaPropuesta As Date, _
                                        ByVal pcUltimaActualizacion As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHEvaluacionDesempeñoFinal
    sSql = "UPDATE RHEvaluacionDesempeñoFinal"
    sSql = sSql & " SET bRHEvalPromocionar=" & pbRHEvalPromocionar & _
                        ",cRHCargoCodPropuesto='" & pcRHCargoCodPropuesto & _
                        "',bRHEvalIncremento=" & pbRHEvalIncremento & _
                        ",nRHEvalSueldoPropuesto=" & pnRHEvalSueldoPropuesto & _
                        ",bRHEvalRenovacion=" & pbRHEvalRenovacion & _
                        ",dRHEvalFechaPropuesta='" & pdRHEvalFechaPropuesta & _
                        "',cUltimaActualizacion='" & pcUltimaActualizacion & _
                        "' WHERE cRHCod='" & pcRHCod & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEvaluacionDesempeñoFinal:
    Err.Raise Err.Number, "ModificaRHEvaluacionDesempeñoFinal", Err.Description
End Sub

Public Sub ModificaRHEvaluacionDesempeñoFinal_x_RRHH(ByVal pcRHCod As String, _
                                            ByVal pcRHEvalComentarioRRHH As String, _
                                            ByVal pcUltimaActualizacion As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHEvaluacionDesempeñoFinal_x_RRHH
    sSql = "UPDATE RHEvaluacionDesempeñoFinal"
    sSql = sSql & " SET cRHEvalComentarioRRHH='" & pcRHEvalComentarioRRHH & _
                        "',cUltimaActualizacion='" & pcUltimaActualizacion & _
                        "' WHERE cRHCod='" & pcRHCod & "'"
                        
                        '"' ,nEstado=2 WHERE cRHCod='" & pcRHCod & "'"  'No modificamos el ESTADO
                        ' 19-11
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEvaluacionDesempeñoFinal_x_RRHH:
    Err.Raise Err.Number, "ModificaRHEvaluacionDesempeñoFinal_x_RRHH", Err.Description
End Sub

Public Sub ModificaRHEvaluacionDesempeñoFinal_x_Decision(ByVal pcRHCod As String, _
                                            ByVal pbRHEvalRenovacionDecision As Integer, _
                                            ByVal pcRHEvalFechaDecision As String, _
                                            ByVal pcRHEvalCargoDecision As String, _
                                            ByVal pnRHEvalSueldoDecision As Double, _
                                            ByVal pcRHEvalObservacionesDecision As String, _
                                            ByVal pcUltimaActualizacion As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHEvaluacionDesempeñoFinal_x_Decision
    sSql = "UPDATE RHEvaluacionDesempeñoFinal"
    sSql = sSql & " SET bRHEvalRenovacionDecision=" & pbRHEvalRenovacionDecision & _
                        ",cRHEvalFechaDecision='" & pcRHEvalFechaDecision & _
                        "',cRHEvalCargoDecision='" & pcRHEvalCargoDecision & _
                        "',nRHEvalSueldoDecision=" & pnRHEvalSueldoDecision & _
                        ",cRHEvalObservacionesDecision='" & pcRHEvalObservacionesDecision & _
                        "',cUltimaActualizacion='" & pcUltimaActualizacion & _
                        "',nEstado=3 WHERE cRHCod='" & pcRHCod & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHEvaluacionDesempeñoFinal_x_Decision:
    Err.Raise Err.Number, "ModificaRHEvaluacionDesempeñoFinal_x_Decision", Err.Description
End Sub

Public Function EliminaRHEvaluacionDesempeñoFinal(ByVal pcRHCod As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHEvaluacionDesempeñoFinal
    sSql = "DELETE FROM RHEvaluacionDesempeñoFinal" '
    sSql = sSql & " WHERE cRHCod='" & pcRHCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHEvaluacionDesempeñoFinal:
    Err.Raise Err.Number, "EliminaRHEvaluacionDesempeñoFinal", Err.Description
End Function

Public Function Obtener_Evaluadores_x_Cargo_EvaluacionPeriodica(pcRHCargoCodEvaluado As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Evaluadores_x_Cargo_EvaluacionPeriodica
    
    sSql = "SELECT cRHCargoCodEvaluador,cRHCargoDescripcion,nOrden,bTodasAgencias=CONVERT(int,bTodasAgencias) from " & _
           " RHEvaluacionPeriodicaNivelAprobacion ENA INNER JOIN RHCargosTabla Cat " & _
           " ON Cat.cRHCargoCod= ENA.cRHCargoCodEvaluador " & _
           " WHERE cRHCargoCodEvaluado ='" & pcRHCargoCodEvaluado & "' ORDER BY nOrden"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Evaluadores_x_Cargo_EvaluacionPeriodica = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Evaluadores_x_Cargo_EvaluacionPeriodica:
    Err.Raise Err.Number, "Obtener_Evaluadores_x_Cargo_EvaluacionPeriodica", Err.Description
End Function


