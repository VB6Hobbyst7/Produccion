VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSeleccionProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InsertaRHSeleccionProceso(ByVal pcCodProceso As String, _
                    ByVal pcDescProceso As String, _
                    ByVal pnTipoProceso As Integer, _
                    ByVal pnTipoContrato As Integer, _
                    ByVal pdFechaInicio As Date, _
                    ByVal pdFechaFin As Date, _
                    ByVal pnEstado As Long, _
                    Optional ByVal pcCodProcesoPublicacion As String = "")
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionProceso
    
    sSql = "INSERT INTO RHSeleccionProceso(cCodProceso,cDescProceso,nTipoProceso,nTipoContrato,dFechaIni,dFechaFin,nEstado,cCodProcesoPublicacion)"
    sSql = sSql & " VALUES('" & pcCodProceso & "','" & pcDescProceso & "'," & pnTipoProceso & "," & pnTipoContrato & ",'" & Format(pdFechaInicio, gsFormatoFecha) & "','" & Format(pdFechaFin, gsFormatoFecha) & "'," & pnEstado & ",'" & pcCodProcesoPublicacion & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionProceso:
    Err.Raise Err.Number, "InsertaRHSeleccionProceso", Err.Description
End Sub

Public Sub ModificaRHSeleccionProceso(ByVal pcCodProceso As String, _
                    ByVal pcDescProceso As String, _
                    ByVal pnTipoProceso As Integer, _
                    ByVal pnTipoContrato As Integer, _
                    ByVal pdFechaInicio As Date, _
                    ByVal pdFechaFin As Date, _
                    ByVal pnEstado As Long, _
                    Optional ByVal pcCodProcesoPublicacion As String = "")

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHSeleccionProceso
    sSql = "UPDATE RHSeleccionProceso"
    sSql = sSql & " SET cCodProceso='" & pcCodProceso & _
                        "',cDescProceso='" & pcDescProceso & _
                        "',nTipoProceso=" & pnTipoProceso & _
                        ",nTipoContrato=" & pnTipoContrato & _
                        ",cCodProcesoPublicacion='" & pcCodProcesoPublicacion & _
                        "' WHERE cCodProceso='" & pcCodProceso & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHSeleccionProceso:
    Err.Raise Err.Number, "ModificaRHSeleccionProceso", Err.Description
End Sub

Public Function EliminaRHSeleccionProceso(ByVal pcCodProceso As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionProceso
    sSql = "DELETE FROM RHSeleccionProceso" '
    sSql = sSql & " WHERE cCodProceso ='" & pcCodProceso & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionProceso:
    Err.Raise Err.Number, "EliminaRHSeleccionProceso", Err.Description
End Function


Public Sub InsertaRHSeleccionProcesoComite(ByVal pcCodProceso As String, _
                                            ByVal pcPersCod As String, _
                                            ByVal pnCargo As Long, _
                                            ByVal pcUltimaActualizacion As String, _
                                            Optional ByVal pcPersCodTitu As String = "", _
                                            Optional ByVal pnOrdenRevision As Integer = 0)
                                            
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionProcesoComite
    
    If pcPersCodTitu = "" Then  'Miembro Suplente
        sSql = "INSERT INTO RHSeleccionProcesoComite(cCodProceso,cPersCod,nCargo,cUltimaActualizacion,nOrdenRevision)"
        sSql = sSql & " VALUES('" & pcCodProceso & "','" & pcPersCod & "'," & pnCargo & ",'" & pcUltimaActualizacion & "'," & pnOrdenRevision & ")"
    Else    'Miembro Titular
        sSql = "INSERT INTO RHSeleccionProcesoComite(cCodProceso,cPersCod,nCargo,cUltimaActualizacion,cPersCodTitu,nOrdenRevision)"
        sSql = sSql & " VALUES('" & pcCodProceso & "','" & pcPersCod & "'," & pnCargo & ",'" & pcUltimaActualizacion & "','" & pcPersCodTitu & "'," & pnOrdenRevision & ")"
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionProcesoComite:
    Err.Raise Err.Number, "InsertaRHSeleccionProcesoComite", Err.Description
End Sub

'Public Sub ModificaRHSeleccionProcesoComite(ByVal pcCodProceso As String, _
'                                            ByVal pcPersCod As String, _
'                                            ByVal pnCargo As Long, _
'                                            ByVal pcUltimaActualizacion As String)
'Dim sSql As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorModificaRHSeleccionProcesoComite
'    sSql = "UPDATE RHSeleccionProcesoComite" '
'    sSql = sSql & " SET cCodProceso='" & pcCodProceso & "',cPersCod='" & pcPersCod & "'" & ",nEstado=" & pnEstado & " WHERE nCodSeleccionProcesoComite=" & pnCodSeleccionProcesoComite
'
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSql
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Sub
'
'ErrorModificaRHSeleccionProcesoComite:
'    Err.Raise Err.Number, "ModificaRHSeleccionProcesoComite", Err.Description
'End Sub

Public Function EliminaRHSeleccionProcesoComite(ByVal pcCodProceso As String, _
                                            ByVal pcPersCod As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionProcesoComite
    sSql = "DELETE FROM RHSeleccionProcesoComite " '
    sSql = sSql & "WHERE cCodProceso='" & pcCodProceso & _
                    "' AND cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionProcesoComite:
    Err.Raise Err.Number, "EliminaRHSeleccionProcesoComite", Err.Description
End Function

Public Sub InsertaRHSeleccionFase(ByVal pnCodFase As Long, _
                    ByVal pcCodProceso As String, _
                    ByVal pnNotaMax As Double, _
                    ByVal pnNotaMin As Double, _
                    ByVal pnNotaAproba As Double, _
                    ByVal pdFechaInicio As Date, _
                    ByVal pdFechaFin As Date, _
                    ByVal pnOrden As Long, _
                    ByVal pnPonderacion As Double)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionFase
    
    sSql = "INSERT INTO RHSeleccionFase(nCodFase, cCodProceso, nNotaMax, nNotaMin, nNotaAprobatoria, dFechaIni, dFechaFin, nOrden, nPonderacion)"
    sSql = sSql & " VALUES(" & pnCodFase & ",'" & pcCodProceso & "'," & pnNotaMax & "," & pnNotaMin & "," & pnNotaAproba & ",'" & pdFechaInicio & "','" & pdFechaFin & "'," & pnOrden & "," & pnPonderacion & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    
    'Cambiar la Fecha Inicial del Proceso Seleccion(para la Primera Fase)
    If pnOrden = 1 Then
        sSql = "UPDATE RHSeleccionProceso SET dFechaIni='" & Format(pdFechaInicio, gsFormatoFecha) & "' WHERE cCodProceso='" & pcCodProceso & "'"
        oConecta.ConexionActiva.Execute sSql
    End If
    'Cambiar la Fecha Final Proceso de Seleccion
    sSql = "UPDATE RHSeleccionProceso SET dFechaFin='" & Format(pdFechaFin, gsFormatoFecha) & "' WHERE cCodProceso='" & pcCodProceso & "'"
    oConecta.ConexionActiva.Execute sSql
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionFase:
    Err.Raise Err.Number, "InsertaRHSeleccionFase", Err.Description
End Sub

Public Sub ModificaRHSeleccionFase(ByVal pnCodFase As Long, _
                    ByVal pcCodProceso As String, _
                    ByVal pnNotaMax As Double, _
                    ByVal pnNotaMin As Double, _
                    ByVal pnNotaAproba As Double, _
                    ByVal pdFechaInicio As Date, _
                    ByVal pdFechaFin As Date, _
                    ByVal pnOrden As Long, _
                    ByVal pnPonderacion As Double, _
                    Optional ByVal pbActualizaFechas As Boolean = True)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHSeleccionFase
    
    If pbActualizaFechas = True Then
        sSql = "UPDATE RHSeleccionFase"
        sSql = sSql & " SET nCodFase=" & pnCodFase & _
                        ",cCodProceso='" & pcCodProceso & _
                        "',nNotaMax=" & pnNotaMax & _
                        ",nNotaMin=" & pnNotaMin & _
                        ",nNotaAprobatoria=" & pnNotaAproba & _
                        ",dFechaIni='" & pdFechaInicio & _
                        "',dFechaFin='" & pdFechaFin & _
                        "',nOrden=" & pnOrden & _
                        ",nPonderacion=" & pnPonderacion & _
                        " WHERE  nCodFase=" & pnCodFase & _
                        " AND cCodProceso='" & pcCodProceso & "'"
    Else    'Solo Actualizar las Ponderaciones mas no los demas datos
        sSql = "UPDATE RHSeleccionFase"
        sSql = sSql & " SET nCodFase=" & pnCodFase & _
                        ",cCodProceso='" & pcCodProceso & _
                        "',nNotaMax=" & pnNotaMax & _
                        ",nNotaMin=" & pnNotaMin & _
                        ",nNotaAprobatoria=" & pnNotaAproba & _
                        ",nPonderacion=" & pnPonderacion & _
                        " WHERE  nCodFase=" & pnCodFase & _
                        " AND cCodProceso='" & pcCodProceso & "'"
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHSeleccionFase:
    Err.Raise Err.Number, "ModificaRHSeleccionFase", Err.Description
End Sub

Public Function EliminaRHSeleccionFase(ByVal pnCodFase As Long, _
                                    ByVal pcCodProceso As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionFase
    sSql = "DELETE FROM RHSeleccionFase" '
    sSql = sSql & " WHERE nCodFase=" & pnCodFase & _
            " AND cCodProceso ='" & pcCodProceso & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionFase:
    Err.Raise Err.Number, "EliminaRHSeleccionFase", Err.Description
End Function

Public Sub InsertaRHSeleccionRequisitos(ByVal pnCodRequisito As Long, _
                        ByVal pcDescRequisito As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionRequisitos
    
    sSql = "INSERT INTO RHSeleccionRequisitos(nCodRequisito, cDescRequisito)"
    sSql = sSql & " VALUES(" & pnCodRequisito & ",'" & pcDescRequisito & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionRequisitos:
    Err.Raise Err.Number, "InsertaRHSeleccionRequisitos", Err.Description
End Sub

Public Sub ModificaRHSeleccionRequisitos(ByVal pnCodRequisito As Long, _
                        ByVal pcDescRequisito As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHSeleccionRequisitos
    sSql = "UPDATE RHSeleccionRequisitos"
    sSql = sSql & " SET nCodRequisito=" & pnCodRequisito & _
                        ",cDescRequisito='" & pcDescRequisito & "'" & _
                        " WHERE  nCodRequisito=" & pnCodRequisito
 
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHSeleccionRequisitos:
    Err.Raise Err.Number, "ModificaRHSeleccionRequisitos", Err.Description
End Sub

Public Function EliminaRHSeleccionRequisitos(ByVal pnCodRequisito As Long)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionRequisitos
    sSql = "DELETE FROM RHSeleccionRequisitos" '
    sSql = sSql & " WHERE nCodRequisito=" & pnCodRequisito
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionRequisitos:
    Err.Raise Err.Number, "EliminaRHSeleccionRequisitos", Err.Description
End Function

Public Sub InsertaRHSeleccionRequisito(ByVal pcCodProceso As String, _
                        ByVal pnCodRequisito As Long, _
                        ByVal pnTipoRequisito As Long)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionRequisito
    
    sSql = "INSERT INTO RHSeleccionRequisito(cCodProceso,nCodRequisito, nTipoRequisito)"
    sSql = sSql & " VALUES('" & pcCodProceso & "'," & pnCodRequisito & "," & pnTipoRequisito & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionRequisito:
    Err.Raise Err.Number, "InsertaRHSeleccionRequisito", Err.Description
End Sub

Public Sub ModificaRHSeleccionRequisito(ByVal pcCodProceso As String, _
                        ByVal pnCodRequisito As Long, _
                        ByVal pnTipoRequisito As Long)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHSeleccionRequisito
    sSql = "UPDATE RHSeleccionRequisito"
    sSql = sSql & " SET pcCodProceso='" & pcCodProceso & _
                        "',nCodRequisito=" & pnCodRequisito & _
                        ",nTipoRequisito=" & pnTipoRequisito & _
                        " WHERE  nCodRequisito=" & pnCodRequisito & _
                            " AND pcCodProceso='" & pcCodProceso & "'"
 
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHSeleccionRequisito:
    Err.Raise Err.Number, "ModificaRHSeleccionRequisito", Err.Description
End Sub

Public Function EliminaRHSeleccionRequisito(ByVal pcCodProceso As String, _
                        ByVal pnCodRequisito As Long)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionRequisito
    sSql = "DELETE FROM RHSeleccionRequisito" '
    sSql = sSql & " WHERE nCodRequisito=" & pnCodRequisito & _
                " AND cCodProceso='" & pcCodProceso & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionRequisito:
    Err.Raise Err.Number, "EliminaRHSeleccionRequisito", Err.Description
End Function

Public Sub InsertaRHSeleccionEvaluacion(ByVal pnCodFase As Long, _
                                    ByVal pcCodProceso As String, _
                                    ByVal pcPersCod As String, _
                                    ByVal pcPersCodResp As String, _
                                    ByVal pnNota As Double, _
                                    ByVal pcComentario As String, _
                                    ByVal pbGanador As Integer, _
                                    Optional ByVal pcUltimaActualizacion As String = "")
                    
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
Dim nLista As Long
    
    On Error GoTo ErrorInsertaRHSeleccionEvaluacion
    
        sSql = "INSERT INTO RHSeleccionEvaluacion(nCodFase,cCodProceso,cPersCod,cPersCodResp,nNota,cComentario,bGanador)"
        sSql = sSql & " VALUES(" & pnCodFase & ",'" & pcCodProceso & "','" & pcPersCod & "','" & pcPersCodResp & "'," & pnNota & ",'" & pcComentario & "'," & pbGanador & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    
    'Para el control del Historico de Listado de Postulantes
    If pcUltimaActualizacion <> "" Then
        sSql = "SELECT NroLista=ISNULL(nLista,0) FROM RHSeleccionEvaluacionHistorico WHERE nCodFase=" & pnCodFase & " AND cCodProceso='" & pcCodProceso & "' AND cPersCodResp='" & pcPersCodResp & "'"
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            nLista = 1
        Else
            nLista = rs("NroLista")
        End If
    
        sSql = "INSERT INTO RHSeleccionEvaluacionHistorico(nLista,nCodFase,cCodProceso,cPersCod,cPersCodResp,nNota,cComentario,bGanador,bEstado,cUltimaActualizacion)"
        sSql = sSql & " VALUES(" & nLista & "," & pnCodFase & ",'" & pcCodProceso & "','" & pcPersCod & "','" & pcPersCodResp & "'," & pnNota & ",'" & pcComentario & "'," & pbGanador & ",1,'" & pcUltimaActualizacion & "')"
        oConecta.ConexionActiva.Execute sSql
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Set rs = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionEvaluacion:
    Err.Raise Err.Number, "InsertaRHSeleccionEvaluacion", Err.Description
End Sub

Public Sub ModificaRHSeleccionEvaluacion(ByVal pnCodFase As Long, _
                                        ByVal pcCodProceso As String, _
                                        ByVal pcPersCod As String, _
                                        ByVal pcPersCodResp As String, _
                                        ByVal pnNota As Double, _
                                        ByVal pcComentario As String, _
                                        ByVal pbGanador As Integer, _
                                        Optional ByVal pcUltimaActualizacion As String = "")

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
Dim nLista As Long
    
    On Error GoTo ErrorModificaRHSeleccionEvaluacion
    
        sSql = "UPDATE RHSeleccionEvaluacion"
        sSql = sSql & " SET nCodFase=" & pnCodFase & _
                        " ,cCodProceso='" & pcCodProceso & _
                        "',cPersCod='" & pcPersCod & _
                        "',cPersCodResp='" & pcPersCodResp & _
                        "',nNota=" & pnNota & _
                        ",cComentario='" & pcComentario & _
                        "',bGanador=" & pbGanador & _
               " WHERE cCodProceso='" & pcCodProceso & "'" & _
               " AND nCodFase=" & pnCodFase & " AND cPersCod='" & pcPersCod & "' AND cPersCodResp='" & pcPersCodResp & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    
    'Para el control del Historico de Listado de Postulantes
    If pcUltimaActualizacion <> "" Then
    
        sSql = "UPDATE RHSeleccionEvaluacionHistorico SET bEstado=0 WHERE cCodProceso='" & pcCodProceso & "' AND nCodFase=" & pnCodFase & " AND cPersCod='" & pcPersCod & "' AND cPersCodResp='" & pcPersCodResp & "'"
        oConecta.ConexionActiva.Execute sSql
    
        sSql = "SELECT NroLista=ISNULL(MAX(nLista),0) FROM RHSeleccionEvaluacionHistorico WHERE nCodFase=" & pnCodFase & " AND cCodProceso='" & pcCodProceso & "' AND cPersCodResp='" & pcPersCodResp & "' AND bEstado=0"
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            nLista = 1
        Else
            nLista = rs("NroLista") + 1
        End If
    
        sSql = "INSERT INTO RHSeleccionEvaluacionHistorico(nLista,nCodFase,cCodProceso,cPersCod,cPersCodResp,nNota,cComentario,bGanador,bEstado,cUltimaActualizacion)"
        sSql = sSql & " VALUES(" & nLista & "," & pnCodFase & ",'" & pcCodProceso & "','" & pcPersCod & "','" & pcPersCodResp & "'," & pnNota & ",'" & pcComentario & "'," & pbGanador & ",1,'" & pcUltimaActualizacion & "')"
        oConecta.ConexionActiva.Execute sSql
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHSeleccionEvaluacion:
    Err.Raise Err.Number, "ModificaRHSeleccionEvaluacion", Err.Description
End Sub

Public Function EliminaRHSeleccionEvaluacion(ByVal pnCodFase As Long, _
                                    ByVal pcCodProceso As String, _
                                    ByVal pcPersCod As String, _
                                    ByVal pcPersCodResp As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionEvaluacion
    sSql = "DELETE FROM RHSeleccionEvaluacion" '
    sSql = sSql & " WHERE cCodProceso ='" & pcCodProceso & "'" & _
            " AND nCodFase=" & pnCodFase & " AND cPersCod='" & pcPersCod & "' AND cPersCodResp='" & pcPersCodResp & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionEvaluacion:
    Err.Raise Err.Number, "EliminaRHSeleccionEvaluacion", Err.Description
End Function

Public Sub InsertaRHSeleccionEvaluacionCurricular(ByVal pnCodFase As Long, _
                                    ByVal pcCodProceso As String, _
                                    ByVal pcPersCod As String, _
                                    ByVal pcPersCodResp As String, _
                                    ByVal pbTipoEvalua As Integer, _
                                    ByVal pnNota As Double)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionEvaluacionCurricular
    
    sSql = "INSERT INTO RHSeleccionEvaluacionCurricular(nCodFase,cCodProceso,cPersCod,cPersCodResp, bTipoEvalua,nNota)"
    sSql = sSql & " VALUES(" & pnCodFase & ",'" & pcCodProceso & "','" & pcPersCod & "','" & pcPersCodResp & "'," & pbTipoEvalua & "," & pnNota & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionEvaluacionCurricular:
    Err.Raise Err.Number, "InsertaRHSeleccionEvaluacionCurricular", Err.Description
End Sub

Public Sub ModificaRHSeleccionEvaluacionCurricular(ByVal pnCodFase As Long, _
                                    ByVal pcCodProceso As String, _
                                    ByVal pcPersCod As String, _
                                    ByVal pcPersCodResp As String, _
                                    ByVal pbTipoEvalua As Integer, _
                                    ByVal pnNota As Double)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHSeleccionEvaluacionCurricular
    
    sSql = "UPDATE RHSeleccionEvaluacionCurricular"
    sSql = sSql & " SET nCodFase=" & pnCodFase & _
                        " ,cCodProceso='" & pcCodProceso & _
                        "',cPersCod='" & pcPersCod & _
                        "',bTipoEvalua=" & pbTipoEvalua & _
                        " WHERE cCodProceso='" & pcCodProceso & "'" & _
                        " AND nCodFase=" & pnCodFase & " AND cPersCod='" & pcPersCod & "'" & _
                        " AND cPersCodResp='" & pcPersCodResp & "' AND bTipoEvalua=" & pbTipoEvalua
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHSeleccionEvaluacionCurricular:
    Err.Raise Err.Number, "ModificaRHSeleccionEvaluacionCurricular", Err.Description
End Sub

Public Function EliminaRHSeleccionEvaluacionCurricular(ByVal pnCodFase As Long, _
                                    ByVal pcCodProceso As String, _
                                    ByVal pcPersCod As String, _
                                    ByVal pcPersCodResp As String, _
                                    ByVal pbTipoEvalua As Integer)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionEvaluacionCurricular
    sSql = "DELETE FROM RHSeleccionEvaluacionCurricular" & _
                " WHERE cCodProceso='" & pcCodProceso & "'" & _
                " AND nCodFase=" & pnCodFase & " AND cPersCod='" & pcPersCod & "'" & _
                " AND cPersCodResp='" & pcPersCodResp & "' AND bTipoEvalua=" & pbTipoEvalua
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionEvaluacionCurricular:
    Err.Raise Err.Number, "EliminaRHSeleccionEvaluacionCurricular", Err.Description
End Function

'Public Sub InsertaRHSeleccionEvaluacionEntrevista(ByVal pnCodFase As Long, _
'                                    ByVal pcCodProceso As String, _
'                                    ByVal pcPersCod As String, _
'                                    ByVal pcPersCodEntre As String, _
'                                    ByVal pnNota As Double)
'
'Dim sSql As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorInsertaRHSeleccionEvaluacionEntrevista
'
'    sSql = "INSERT INTO RHSeleccionEvaluacionEntrevista(nCodFase,cCodProceso,cPersCod,cPersCodEntre,nNota)"
'    sSql = sSql & " VALUES(" & pnCodFase & ",'" & pcCodProceso & "','" & pcPersCod & "','" & pcPersCodEntre & "'," & pnNota & "')"
'
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSql
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Sub
'
'ErrorInsertaRHSeleccionEvaluacionEntrevista:
'    Err.Raise Err.Number, "InsertaRHSeleccionEvaluacionEntrevista", Err.Description
'End Sub
'
'Public Sub ModificaRHSeleccionEvaluacionEntrevista(ByVal pnCodFase As Long, _
'                                    ByVal pcCodProceso As String, _
'                                    ByVal pcPersCod As String, _
'                                    ByVal pcPersCodEntre As String, _
'                                    ByVal pnNota As Double)
'
'Dim sSql As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorModificaRHSeleccionEvaluacionEntrevista
'
'    sSql = "UPDATE RHSeleccionEvaluacionEntrevista"
'    sSql = sSql & " SET nNota=" & pnNota & _
'                        " WHERE cCodProceso='" & pcCodProceso & "'" & _
'                        " AND nCodFase=" & pnCodFase & " AND cPersCod='" & pcPersCod & "'" & _
'                        " AND cPersCodEntre=" & pcPersCodEntre & "'"
'
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSql
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Sub
'
'ErrorModificaRHSeleccionEvaluacionEntrevista:
'    Err.Raise Err.Number, "ModificaRHSeleccionEvaluacionEntrevista", Err.Description
'End Sub
'
'Public Function EliminaRHSeleccionEvaluacionEntrevista(ByVal pnCodFase As Long, _
'                                    ByVal pcCodProceso As String, _
'                                    ByVal pcPersCod As String, _
'                                    ByVal pcPersCodEntre As String)
'
'Dim sSql As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorEliminaRHSeleccionEvaluacionEntrevista
'    sSql = "DELETE FROM RHSeleccionEvaluacionEntrevista" & _
'                " WHERE cCodProceso='" & pcCodProceso & "'" & _
'                " AND nCodFase=" & pnCodFase & " AND cPersCod='" & pcPersCod & "'" & _
'                " AND cPersCodEntre='" & pcPersCodEntre & "'"
'
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSql
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorEliminaRHSeleccionEvaluacionEntrevista:
'    Err.Raise Err.Number, "EliminaRHSeleccionEvaluacionEntrevista", Err.Description
'End Function


Public Sub InsertaRHSeleccionFaseAprobacion(ByVal pnCodFase As Long, _
                                            ByVal pcCodProceso As String, _
                                            ByVal pcPersCod As String, _
                                            ByVal pnLista As Long, _
                                            ByVal pnAceptado As Long, _
                                            ByVal pcComentario As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionFaseAprobacion
    sSql = "INSERT INTO RHSeleccionFaseAprobacion(nCodFase,cCodProceso,cPersCod,nLista,nAceptado ,cComentario,dFecha)"
    sSql = sSql & " VALUES(" & pnCodFase & ",'" & pcCodProceso & "','" & pcPersCod & "'," & pnLista & "," & pnAceptado & ",'" & pcComentario & "',getdate())"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionFaseAprobacion:
    Err.Raise Err.Number, "InsertaRHSeleccionFaseAprobacion", Err.Description
End Sub

Public Sub ModificaRHSeleccionFaseAprobacion(ByVal pnCodFase As Long, _
                                            ByVal pcCodProceso As String, _
                                            ByVal pcPersCod As String, _
                                            ByVal pnItem As Long, _
                                            ByVal pnAceptado As Long, _
                                            ByVal pcComentario As String)
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorModificaRHSeleccionFaseAprobacion
    sSql = "UPDATE RHSeleccionFaseAprobacion" '
    sSql = sSql & " SET nCodFase=" & pnCodFase & ", cCodProceso='" & pcCodProceso & "',cPersCod='" & pcPersCod & "'" & ",nItem=" & pnItem & ",nAceptado=" & pnAceptado & ",cComentario='" & pcComentario & "'" & _
            " WHERE nCodFase=" & pnCodFase & " AND cCodProceso='" & pcCodProceso & "' AND cPersCod='" & pcPersCod & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorModificaRHSeleccionFaseAprobacion:
    Err.Raise Err.Number, "ModificaRHSeleccionFaseAprobacion", Err.Description
End Sub

Public Function EliminaRHSeleccionFaseAprobacion(ByVal pnCodFase As Long, _
                                            ByVal pcCodProceso As String, _
                                            ByVal pcPersCod As String, _
                                            ByVal pnItem As Long)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionFaseAprobacion
    sSql = "DELETE FROM RHSeleccionFaseAprobacion " '
    sSql = sSql & " WHERE nCodFase=" & pnCodFase & " AND cCodProceso='" & pcCodProceso & "' AND cPersCod='" & pcPersCod & "' AND nItem=" & pnItem
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionFaseAprobacion:
    Err.Raise Err.Number, "EliminaRHSeleccionFaseAprobacion", Err.Description
End Function

Public Sub InsertaRHPersonaFamiliar(ByVal psPersCod As String, ByVal psDNI As String, ByVal psNombre As String, ByVal psFechaNac As Date, ByVal psZona As String, _
                                      ByVal psDireccion As String, ByVal psTelefono As String, ByVal psOTelefono As String, ByVal psEmail As String, _
                                      ByVal psSexo As String, ByVal psEstadoCivil As Integer, ByVal pshijos As Integer)
                    
Dim sSql1 As String
Dim sSql2 As String
Dim sSql3 As String

Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHPersonaFamiliar
    
   
    sSql1 = "insert into Persona(cPersCod,cPersNombre,dPersNacCreac,cPersDireccUbiGeo,cPersDireccDomicilio,cPersTelefono,cPersTelefono2,cPersEmail,nPersPersoneria,cPersEstado)" & _
            " values('" & psPersCod & "','" & psNombre & "','" & Format(psFechaNac, gsFormatoFecha) & "','" & psZona & "','" & psDireccion & "', '" & psTelefono & "','" & psOTelefono & " ','" & psEmail & "','1','10')"
    
    sSql2 = "insert into PersID(cPersCod,cPersIDnro,cPersIDTpo)" & _
            " values ('" & psPersCod & "','" & psDNI & "','1') "
    
    sSql3 = "insert into PersonaNat(cPersCod,cPersNatSexo,nPersNatEstCiv,nPersNatHijos)" & _
            " values('" & psPersCod & "','" & psSexo & "','" & psEstadoCivil & "','" & pshijos & "') "
    
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql1
    oConecta.ConexionActiva.Execute sSql2
    oConecta.ConexionActiva.Execute sSql3
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHPersonaFamiliar:
    Err.Raise Err.Number, "InsertaRHPersonaFamiliar", Err.Description
End Sub

Public Sub InsertaRHSeleccionNivelAprobacion(ByVal pcCodCargoNivel As String, _
                                        ByVal pcCodCargoAprobac As String, _
                                        ByVal pnOrden As Long, _
                                        ByVal pnNumAprobac As Long, _
                                        ByVal pcCodCargoSol As String, _
                                        ByVal pbTodasAge As Integer, _
                                        ByVal pbVerificaUltNivel As Integer)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionNivelAprobacion
    
    sSql = "INSERT INTO RHSeleccionNivelAprobacion(cCodCargoReq,cCodCargoAprobac,nOrden,nNumAprobac,cCodCargoSol,bTodasAge,bVerificaUltNivel)"
    sSql = sSql & " VALUES('" & pcCodCargoNivel & "','" & pcCodCargoAprobac & "'," & pnOrden & "," & pnNumAprobac & ",'" & pcCodCargoSol & "'," & pbTodasAge & "," & pbVerificaUltNivel & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionNivelAprobacion:
    Err.Raise Err.Number, "InsertaRHSeleccionNivelAprobacion", Err.Description
End Sub

Public Sub ModificaRHSeleccionNivelAprobacion(ByVal pnCodNivelAprobac As Long, _
                                        ByVal pcCodCargoNivel As String, _
                                        ByVal pcCodCargoAprobac As String, _
                                        ByVal pnOrden As Long, _
                                        ByVal pnNumAprobac As Long, _
                                        ByVal pcCodCargoSol As String, _
                                        ByVal pbTodasAge As Integer, _
                                        ByVal pbVerificaUltNivel As Integer)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHSeleccionNivelAprobacion
    sSql = "UPDATE RHSeleccionNivelAprobacion"
    sSql = sSql & " SET cCodCargoReq='" & pcCodCargoNivel & _
                        "',cCodCargoAprobac='" & pcCodCargoAprobac & _
                        "',nOrden=" & pnOrden & _
                        ",nNumAprobac=" & pnNumAprobac & _
                        ",cCodCargoSol='" & pcCodCargoSol & _
                        "',bTodasAge=" & pbTodasAge & _
                        ",bVerificaUltNivel=" & pbVerificaUltNivel & _
                        " WHERE nCodNivelAprobac=" & pnCodNivelAprobac
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHSeleccionNivelAprobacion:
    Err.Raise Err.Number, "ModificaRHSeleccionNivelAprobacion", Err.Description
End Sub

Public Function EliminaRHSeleccionNivelAprobacion(ByVal pnCodNivelAprobac As Long)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionNivelAprobacion
    sSql = "DELETE FROM RHSeleccionNivelAprobacion" '
    sSql = sSql & " WHERE nCodNivelAprobac =" & pnCodNivelAprobac
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionNivelAprobacion:
    Err.Raise Err.Number, "EliminaRHSeleccionNivelAprobacion", Err.Description
End Function

Public Sub InsertaRHSeleccionRequerimientoAprobaciones(ByVal pcCodRequerimiento As String, _
                                        ByVal pcCodCargoReq As String, _
                                        ByVal pcPersCodReq As String, _
                                        ByVal pcCodCargoAprobac As String, _
                                        ByVal pcPersCodAprobac As String, _
                                        ByVal pnOrden As Long, _
                                        ByVal pbAprobado As Integer, _
                                        ByVal pbTodasAge As Integer)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionRequerimientoAprobaciones
    
    sSql = "INSERT INTO RHSeleccionRequerimientoAprobaciones(cCodRequerimiento,cCodCargoReq,cPersCodReq,cCodCargoAprobac,cPersCodAprobac,nOrden,bAprobado,bTodasAge)"
    sSql = sSql & " VALUES('" & pcCodRequerimiento & "','" & pcCodCargoReq & "','" & pcPersCodReq & "','" & pcCodCargoAprobac & "','" & pcPersCodAprobac & "'," & pnOrden & "," & pbAprobado & "," & pbTodasAge & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionRequerimientoAprobaciones:
    Err.Raise Err.Number, "InsertaRHSeleccionRequerimientoAprobaciones", Err.Description
End Sub

Public Sub ModificaRHSeleccionRequerimientoAprobaciones(ByVal pcCodRequerimiento As String, _
                                        ByVal pcCodCargoReq As String, _
                                        ByVal pcPersCodReq As String, _
                                        ByVal pcCodCargoAprobac As String, _
                                        ByVal pcPersCodAprobac As String, _
                                        ByVal pnOrden As Long, _
                                        ByVal pbAprobado As Integer)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHSeleccionRequerimientoAprobaciones
    sSql = "UPDATE RHSeleccionRequerimientoAprobaciones"
    sSql = sSql & " SET cCodRequerimiento='" & pcCodRequerimiento & _
                        "',cCodCargoReq='" & pcCodCargoReq & _
                        "',cPersCodReq=" & pcPersCodReq & _
                        "',cCodCargoAprobac='" & pcCodCargoAprobac & _
                        "',cPersCodAprobac=" & pcPersCodAprobac & _
                        "',nOrden=" & pnOrden & _
                        ",bAprobado=" & pbAprobado & _
                        " WHERE cCodRequerimiento='" & pcCodRequerimiento & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHSeleccionRequerimientoAprobaciones:
    Err.Raise Err.Number, "ModificaRHSeleccionRequerimientoAprobaciones", Err.Description
End Sub

Public Function EliminaRHSeleccionRequerimientoAprobaciones(ByVal pcCodRequerimiento As String, _
                                                        ByVal pcCodCargoActual As String, _
                                                        ByVal pcPersCodReq As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionRequerimientoAprobaciones
    sSql = "DELETE FROM RHSeleccionRequerimientoAprobaciones" '
    sSql = sSql & " WHERE cCodRequerimiento='" & pcCodRequerimiento & "' AND cCodCargoActual='" & pcCodCargoActual & "' AND cPersCodReq='" & pcPersCodReq & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionRequerimientoAprobaciones:
    Err.Raise Err.Number, "EliminaRHSeleccionRequerimientoAprobaciones", Err.Description
End Function


Public Sub InsertaRHSeleccionComiteAprobacion(ByVal pcCodProceso As String, _
                                            ByVal pcPersCod As String, _
                                            ByVal pnTipoAprob As Integer, _
                                            ByVal pcComentario As String, _
                                            ByVal pnAceptado As Integer)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionComiteAprobacion
    sSql = "INSERT INTO RHSeleccionComiteAprobacion(cCodProceso,cPersCod,nTipoAprob,cComentario,nAceptado)"
    sSql = sSql & " VALUES('" & pcCodProceso & "','" & pcPersCod & "'," & pnTipoAprob & ",'" & pcComentario & "'," & pnAceptado & ")"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionComiteAprobacion:
    Err.Raise Err.Number, "InsertaRHSeleccionComiteAprobacion", Err.Description
End Sub

Public Sub ModificaRHSeleccionComiteAprobacion(ByVal pcCodProceso As String, _
                                            ByVal pcPersCod As String, _
                                            ByVal pnTipoAprob As Integer, _
                                            ByVal pcComentario As String, _
                                            ByVal pnAceptado As Integer)
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorModificaRHSeleccionComiteAprobacion
    sSql = "UPDATE RHSeleccionComiteAprobacion" '
    sSql = sSql & " SET cCodProceso='" & pcCodProceso & "',cPersCod='" & pcPersCod & "'" & ",nTipoAprob=" & pnTipoAprob & _
            ", cComentario='" & pcComentario & "', nAceptado=" & pnAceptado & _
            "WHERE cCodProceso='" & pcCodProceso & _
                    "' AND cPersCod='" & pcPersCod & "'" & _
                    " AND nTipoAprob=" & pnTipoAprob

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorModificaRHSeleccionComiteAprobacion:
    Err.Raise Err.Number, "ModificaRHSeleccionComiteAprobacion", Err.Description
End Sub

Public Function EliminaRHSeleccionComiteAprobacion(ByVal pcCodProceso As String, _
                                            ByVal pcPersCod As String, _
                                            ByVal pnTipoAprob As Integer)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionComiteAprobacion
    sSql = "DELETE FROM RHSeleccionComiteAprobacion " '
    sSql = sSql & "WHERE cCodProceso='" & pcCodProceso & _
                    "' AND cPersCod='" & pcPersCod & "'" & _
                    " AND nTipoAprob=" & pnTipoAprob
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionComiteAprobacion:
    Err.Raise Err.Number, "EliminaRHSeleccionComiteAprobacion", Err.Description
End Function

Public Sub InsertaRHSeleccionGerenciaAprobacion(ByVal pcCodProceso As String, _
                                            ByVal pcPersCod As String, _
                                            ByVal pcComentario As String, _
                                            ByVal pnAceptado As Integer)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionGerenciaAprobacion
    sSql = "INSERT INTO RHSeleccionGerenciaAprobacion(cCodProceso,cPersCod,cComentario,nAceptado)"
    sSql = sSql & " VALUES('" & pcCodProceso & "','" & pcPersCod & "','" & pcComentario & "'," & pnAceptado & ")"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionGerenciaAprobacion:
    Err.Raise Err.Number, "InsertaRHSeleccionGerenciaAprobacion", Err.Description
End Sub

Public Sub ModificaRHSeleccionGerenciaAprobacion(ByVal pcCodProceso As String, _
                                            ByVal pcPersCod As String, _
                                            ByVal pcComentario As String, _
                                            ByVal pnAceptado As Integer)
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorModificaRHSeleccionGerenciaAprobacion
    sSql = "UPDATE RHSeleccionGerenciaAprobacion" '
    sSql = sSql & " SET cCodProceso='" & pcCodProceso & "',cPersCod='" & pcPersCod & "'" & _
            ", cComentario='" & pcComentario & "', nAceptado=" & pnAceptado & _
            "WHERE cCodProceso='" & pcCodProceso & _
                    "' AND cPersCod='" & pcPersCod & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorModificaRHSeleccionGerenciaAprobacion:
    Err.Raise Err.Number, "ModificaRHSeleccionGerenciaAprobacion", Err.Description
End Sub

Public Function EliminaRHSeleccionGerenciaAprobacion(ByVal pcCodProceso As String, _
                                            ByVal pcPersCod As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionGerenciaAprobacion
    sSql = "DELETE FROM RHSeleccionGerenciaAprobacion " '
    sSql = sSql & "WHERE cCodProceso='" & pcCodProceso & _
                    "' AND cPersCod='" & pcPersCod & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionGerenciaAprobacion:
    Err.Raise Err.Number, "EliminaRHSeleccionGerenciaAprobacion", Err.Description
End Function

Public Function Obtener_Numero_IP() As String
Dim OConsSistema As New NConstSistemas
Dim cValorConsSist As String

'142: Bit de Control de Pruebas de Pagina Web
cValorConsSist = OConsSistema.LeeConstSistema(142)  '(gnModoPruebaRRHH)
If cValorConsSist = "1" Then
    Obtener_Numero_IP = "[128.107.2.10]"
Else
    Obtener_Numero_IP = "[128.107.2.24]"
End If

Set OConsSistema = Nothing
End Function

Public Function BuscarRHSeleccionNivelesCurriculum(ByVal pbMostrarBachiller As Boolean) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
    
    If pbMostrarBachiller = True Then
        sSql = "SELECT nConsValor,cConsDescripcion FROM CONSTANTE WHERE  nConsCod='9052' and nConsValor <> '9052' AND nConsValor>0 ORDER BY nconsvalor"
    Else
        sSql = "SELECT nConsValor,cConsDescripcion FROM CONSTANTE WHERE  nConsCod='9052' and nConsValor <> '9052' AND nConsValor NOT IN (0,2) ORDER BY nconsvalor"
    End If
    
    If OCON.AbreConexion Then
        Set BuscarRHSeleccionNivelesCurriculum = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

'**************** FUNCIONES DE ANITA: FILTROS DE PROCESO ************

Public Function BuscarRHSeleccionValores(ByVal psValor As String) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
    
    sSql = "SELECT nConsValor,cConsDescripcion FROM CONSTANTE WHERE  nConsCOD='" & psValor & "' and nConsValor <> '" & psValor & "'AND nConsValor>0 ORDER BY nconsvalor"

    If OCON.AbreConexion Then
        Set BuscarRHSeleccionValores = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSeleccionFiltro(Optional ByVal pnTipoFiltro As Integer = 0) As ADODB.Recordset
'0: Todos los Filtros (de Evaluacion y de Admisibilidad)
'1: Solo los de Evaluacion
'2: Solo los de Admisibilidad

    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
    
    Select Case pnTipoFiltro
        Case 0
            sSql = "SELECT cCodigo,cCampoFiltro FROM " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionFiltros"
        Case 1
            sSql = "SELECT cCodigo,cCampoFiltro FROM " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionFiltros WHERE cCodigo IN(2,7)"
        Case 2
            sSql = "SELECT cCodigo,cCampoFiltro FROM " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionFiltros" 'Por Ahora son Todos
    End Select
    If OCON.AbreConexion Then
        Set BuscarRHSeleccionFiltro = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSeleccionFilCondicionN() As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = "SELECT * FROM " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionFilCondicion  where valor <> 'IN' "

    If OCON.AbreConexion Then
        Set BuscarRHSeleccionFilCondicionN = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSeleccionFilCondicionT() As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = "SELECT * FROM " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionFilCondicion where valor='='or valor='in'"

    If OCON.AbreConexion Then
        Set BuscarRHSeleccionFilCondicionT = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSPostulantes(ByVal pscadena As String, pcCodProceso As String) As ADODB.Recordset
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = " select  distinct p.cUsuario ,cNombre,cApellidoPat +' '+ cApellidoMat,cDesZon,cDireccion,p.cTelefono,cEmail,cFechaNac,cDNI,c1.cConsDescripcion,c2.cConsDescripcion,nhijos from " & Obtener_Numero_IP & ".DBWeb.dbo.postulante p " & _
           " inner join " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionCurriculumDPersonales cp on p.cUsuario=cp.cPersCod" & _
           " LEFT join " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionCurriculumDLaborales cl on p.cUsuario=cl.cPersCod" & _
           " inner join " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProcesoPostulante spp on spp.cUsuario=p.cUsuario COLLATE SQL_Latin1_General_CP437_BIN " & _
           " inner join " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProceso spro on spro.cCodProceso=spp.cCodProceso and spro.nEstado=1 " & _
           " inner join constante c1 on p.csexo= c1.nconsvalor and c1.nconscod='9054'" & _
           " inner join constante c2 on p.cEstadoCivil= c2.nconsvalor and c2.nconscod='1020'" & _
           " INNER join DBComunes..Zonas z  on p.czona= z.cCodZon COLLATE SQL_Latin1_General_CP437_BIN " & _
           " WHERE " & pscadena & " AND spro.cCodProceso='" & pcCodProceso & "'"

    If OCON.AbreConexion Then
        Set BuscarRHSPostulantes = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
End Function

Public Sub InsertaRHSeleccionProcesoRequisito(ByVal pscCodProceso As String, _
                        ByVal pscFiltro As String, ByVal pscCondicon As String, ByVal psValor1 As String, ByVal psValor2 As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionProcesoRequisito
    
    sSql = "INSERT INTO " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProcesoRequisitos(cCodProceso,cFiltro,cCondicon,Valor1,Valor2)"
    sSql = sSql & " VALUES('" & pscCodProceso & "','" & pscFiltro & "','" & pscCondicon & "','" & psValor1 & "','" & psValor2 & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionProcesoRequisito:
    Err.Raise Err.Number, "InsertaRHSeleccionProcesoRequisito", Err.Description
End Sub

Public Function BuscarRHProcesoRequisito(ByVal psproceso As String) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = "SELECT NumReq=count(ccodproceso) FROM " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProcesoRequisitos " & _
           "WHERE ccodproceso = '" & psproceso & "' "

    If OCON.AbreConexion Then
        Set BuscarRHProcesoRequisito = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function
    
Public Sub EliminarRHSeleccionProcesoRequisito(ByVal psproceso As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminarRHSeleccionProcesoRequisito
    
    sSql = "DELETE " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProcesoRequisitos WHERE ccodproceso = '" & psproceso & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorEliminarRHSeleccionProcesoRequisito:
    Err.Raise Err.Number, "EliminarRHSeleccionProcesoRequisito", Err.Description
End Sub

Public Sub InsertaRHSeleccionPostulante(ByVal pscodproceso As String, ByVal psUsuario As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionPostulante
    
    sSql = "insert into RHSeleccionPostulante(cCodProceso,cPostuCod)"
    sSql = sSql & " values('" & pscodproceso & "','" & psUsuario & "') "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionPostulante:
    Err.Raise Err.Number, "InsertaRHSeleccionPostulante", Err.Description
End Sub

Public Sub EliminarRHSeleccionPostulante(ByVal pscodproceso As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminarRHSeleccionPostulante
    
    sSql = "delete RHSeleccionPostulante where cCodProceso='" & pscodproceso & "'"
    
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorEliminarRHSeleccionPostulante:
    Err.Raise Err.Number, "EliminarRHSeleccionPostulante", Err.Description
End Sub
'**** nuevas
Public Function BuscarRHZonas() As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = "select cCodZon,cDesZon from  DBComunes..Zonas" & _
           " where substring (ccodzon,4,9) = '000000000'" & _
           " order by cCodZon"

    If OCON.AbreConexion Then
        Set BuscarRHZonas = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSPostulantesAprobadosporProceso(ByVal pscodproceso As String) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = "select cpostucod from DBCmactAux..RHSeleccionPostulante" & _
           " where ccodproceso='" & pscodproceso & "' and nEstado='1'"

    If OCON.AbreConexion Then
        Set BuscarRHSPostulantesAprobadosporProceso = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSPostulanteAprobado(ByVal pscodpostulante As String) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = " select * from DBCmactAux..Postulante " & _
           " where cPostuCod = '" & pscodpostulante & "'"

    If OCON.AbreConexion Then
        Set BuscarRHSPostulanteAprobado = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSPersonaDNI(ByVal psDNI As String) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = " select count(cPersIDnro) from DBCmactAux..PersID" & _
           " where cPersIDnro= '" & psDNI & "'"

    If OCON.AbreConexion Then
        Set BuscarRHSPersonaDNI = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function generarcodigo(ByVal psCodAge As String, ByVal psCodCmact As String) As String
    Dim objpersona As New DPersona
    Dim nobjpersona As New NPersona
    Dim lsUltCorr As String
    Dim lnCorr As Long
    Dim lnCodchequeo As String
    Dim lsCadFinal As String
    Dim sCodPersDistrib As String
    
    'Es temporal miemtras se logea el Usuario
    
    Dim sCodAgencia As String
    
    If psCodAge = "" Then
        sCodAgencia = "07"  'Por Defecto..AG ESPAÑA
    Else
        sCodAgencia = psCodAge
    End If
    
    lsUltCorr = objpersona.UltimoCorrelativo(sCodAgencia) '(psCodAge)
    lnCorr = Val(lsUltCorr) + 1
    lsUltCorr = FillNum(Trim(Str(lnCorr)), 7, "0")
    lnCodchequeo = nobjpersona.DigitoChequeo(lsUltCorr)
    lsCadFinal = lsUltCorr & Trim(Str(lnCodchequeo))
    sCodPersDistrib = Right(Trim(sCodAgencia), 2) & Trim(lsCadFinal) 'Right(Trim(psCodAge), 2) & Trim(lsCadFinal)
    generarcodigo = Trim(psCodCmact) & Trim(sCodAgencia) & Trim(lsCadFinal) 'Trim(psCodCmact) & Trim(psCodAge) & Trim(lsCadFinal)
    
    Set nobjpersona = Nothing
    Set objpersona = Nothing

End Function


Public Sub InsertaRHSPostulanteAprobado(ByVal cUsuario As String, ByVal psNombre As String, ByVal psApellidoPat As String, ByVal psApellidoMat As String, _
                                        ByVal psZona As String, ByVal psDireccion As String, ByVal psTelefono As String, ByVal psOTelefono As String, ByVal psEmail As String, _
                                        ByVal dFechaNac As Date, ByVal psDNI As String, ByVal psLM As String, ByVal psSexo As String, ByVal psEstadoCivil As String, ByVal pshijos As String)
    Dim sSql As String
    Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSPostulanteAprobado
    
    sSql = " insert into postulante(cPostuCod,cNombre,cApellidoPat,cApellidoMat,cZona," & _
           " cDireccion,cTelefono,cOTelefono,cEmail,cFechaNac,cDNI,cLM,cSexo,cEstadoCivil,nhijos )" & _
           " values('" & cUsuario & "','" & psNombre & "', '" & psApellidoPat & "'," & _
           " '" & psApellidoMat & "','" & psZona & "','" & psDireccion & "','" & psTelefono & "'," & _
           " '" & psOTelefono & "','" & psEmail & "','" & Format(dFechaNac, gsFormatoFecha) & "','" & psDNI & "', '" & psLM & "','" & psSexo & "','" & psEstadoCivil & "','" & pshijos & "' )"
             
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSPostulanteAprobado:
    Err.Raise Err.Number, "InsertaRHSPostulanteAprobado", Err.Description

End Sub

Public Sub InsertaRHSPostulanteADpersonales(ByVal psUsuario As String, ByVal pnNivel As Long, ByVal psGrado As String, ByVal psAreaEstudio As String, ByVal psUniversidad As String, ByVal psSituacion As String, ByVal psOtros As String)
    
    Dim sSql As String
    Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSPostulanteADpersonales
    'cPersCod,nNivelAlcanzado,cDescAreaEstudio,cUniversidad,cSituacion
    sSql = " insert into DBCmactAux..RHSeleccionCurriculumDPersonales(cPostuCod,nNivelAlcanzado,nGrado,cDescAreaEstudio,cUniversidad,cSituacion,cOtros)" & _
            " values('" & psUsuario & "'," & pnNivel & ",'" & psGrado & "','" & psAreaEstudio & "','" & psUniversidad & "','" & psSituacion & "','" & psOtros & "')"
                     
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSPostulanteADpersonales:
    Err.Raise Err.Number, "InsertaRHSPostulanteADpersonales", Err.Description

End Sub

Public Sub InsertaRHSPostulanteADLaborales(ByVal psUsuario As String, ByVal psEmpresa As String, ByVal psDescCargo As String, ByVal psTipoContrato As String, _
                                          ByVal psAreaTrabajo As String, ByVal psFunciones As String, ByVal psjefe As String, ByVal psCargoJefe, _
                                          ByVal dfechai As Date, ByVal dfechaf As Date, ByVal psTelefono As String)
                                          
    
    Dim sSql As String
    

    Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSPostulanteADLaborales
    
    
    sSql = " insert into DBCmactAux..RHSeleccionCurriculumDLaborales(cPostuCod,cEmpresa,cDescCargo,cTipoContrato,cAreaTrabajo,cFunciones,cJefe,cCargoJefe,dFechaIni,dFechaFin,cTelefono)" & _
            " values('" & psUsuario & "','" & psEmpresa & "','" & psDescCargo & "','" & psTipoContrato & "','" & psAreaTrabajo & "','" & psFunciones & "','" & psjefe & "','" & psCargoJefe & "','" & Format(dfechai, gsFormatoFecha) & "','" & Format(dfechaf, gsFormatoFecha) & "','" & psTelefono & "')"
            
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
    
ErrorInsertaRHSPostulanteADLaborales:
    Err.Raise Err.Number, "InsertaRHSPostulanteADLaborales", Err.Description

End Sub


Public Sub InsertaRHPostulanteAPersona(ByVal psPersCod As String, ByVal psDNI As String, ByVal psNombre As String, ByVal psFechaNac As Date, ByVal psZona As String, _
                                      ByVal psDireccion As String, ByVal psTelefono As String, ByVal psOTelefono As String, ByVal psEmail As String, _
                                      ByVal psSexo As String, ByVal psEstadoCivil As Integer, ByVal pshijos As Integer, ByRef psPersCodOpe As String, _
                                      Optional ByVal psUltimaActualizacion As String = "")
                    
Dim sSql1 As String
Dim sSql2 As String
Dim sSql3 As String

Dim oConecta As DConecta

Dim rs As New ADODB.Recordset
Dim sSql As String
Dim cCodPersonaReg As String
    
    On Error GoTo ErrorInsertaRHPostulanteAPersona
    
   sSql = "SELECT cPersCod,cPersIDnro FROM PersID WHERE cPersIDnro='" & psDNI & "'"
   
   Set oConecta = New DConecta
   oConecta.AbreConexion
   oConecta.BeginTrans
   
   Set rs = oConecta.CargaRecordSet(sSql)
   
   If rs.EOF Then   'Si no Existe la Persona, lo insertamos
        sSql1 = "insert into Persona(cPersCod,cPersNombre,dPersNacCreac,cPersDireccUbiGeo,cPersDireccDomicilio,cPersTelefono,cPersTelefono2,cPersEmail,nPersPersoneria,cPersEstado,cUltimaActualizacion)" & _
            " values('" & psPersCod & "','" & psNombre & "','" & Format(psFechaNac, gsFormatoFecha) & "','" & psZona & "','" & psDireccion & "', '" & psTelefono & "','" & psOTelefono & " ','" & psEmail & "','1','10','" & psUltimaActualizacion & "')"
    
        sSql2 = "insert into PersID(cPersCod,cPersIDnro,cPersIDTpo)" & _
            " values ('" & psPersCod & "','" & psDNI & "','1') "
    
        sSql3 = "insert into PersonaNat(cPersCod,cPersNatSexo,nPersNatEstCiv,nPersNatHijos)" & _
            " values('" & psPersCod & "','" & psSexo & "','" & psEstadoCivil & "','" & pshijos & "') "
    
   Else 'No insertar, solo actualizar
        cCodPersonaReg = rs!cPersCod
        
        sSql1 = "UPDATE Persona SET cPersNombre='" & psNombre & _
                "',dPersNacCreac='" & Format(psFechaNac, gsFormatoFecha) & _
                "',cPersDireccUbiGeo='" & psZona & _
                "',cPersDireccDomicilio='" & psDireccion & _
                "',cPersTelefono='" & psTelefono & _
                "',cPersTelefono2='" & psOTelefono & _
                "',cPersEmail='" & psEmail & _
                "',nPersPersoneria=1,cPersEstado='10'" & _
                ",cUltimaActualizacion='" & psUltimaActualizacion & _
                "' WHERE cPersCod='" & cCodPersonaReg & "'"
    
        'sSql2 = "UPDATE PersID SET cPersIDnro='" & psDNI & "',cPersIDTpo=1" & _
                " WHERE cPersCod='" & cCodPersonaReg & "'"
    
        sSql3 = "UPDATE PersonaNat SET cPersNatSexo='" & psSexo & _
                "',nPersNatEstCiv=" & psEstadoCivil & _
                ",nPersNatHijos=" & pshijos & _
                " WHERE cPersCod='" & cCodPersonaReg & "'"
   End If
   
   'Limpiamos en caso exista
   If Not rs.EOF Then
        sSql = "DELETE FROM RHEmpleadoDocumento WHERE cPersCod='" & cCodPersonaReg & "'"
        oConecta.ConexionActiva.Execute sSql
   Else
        cCodPersonaReg = psPersCod
   End If
   
   sSql = "INSERT INTO RHEmpleadoDocumento(cPersCod,nTipo,nEstado)" & _
            " SELECT '" & cCodPersonaReg & "', nConsValor, 0 FROM Constante WHERE nConsCod=9070 AND nConsValor<>nConsCod"
    
   'Insertamos para el Control de Documentos
   oConecta.ConexionActiva.Execute sSql
    
   'Set oConecta = New DConecta
   'oConecta.AbreConexion
   oConecta.ConexionActiva.Execute sSql1
   If rs.EOF Then  'Solo insertamos el Documento mas no se debe actualizar
        oConecta.ConexionActiva.Execute sSql2
   End If
   oConecta.ConexionActiva.Execute sSql3
     
   'Devolvemos la Persona que se inserto o actualizo
   psPersCodOpe = cCodPersonaReg
     
   oConecta.CommitTrans
   oConecta.CierraConexion
   Set rs = Nothing
   Set oConecta = Nothing
   Exit Sub
    
ErrorInsertaRHPostulanteAPersona:
    Err.Raise Err.Number, "InsertaRHPostulanteAPersona", Err.Description
    'oConecta.RollBackTrans
End Sub

Public Function BuscarRHSPostulanteporAprobar(ByVal pscodpostulante As String) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = " SELECT cUsuario, cNombre, cApellidoPat, cApellidoMat, cZona, cDireccion, cTelefono, cOTelefono, cEmail, cFechaNac, cDNI, cLM=ISNULL(cLM,''), cSexo, cEstadoCivil, nhijos FROM " & Obtener_Numero_IP & ".DBWeb.dbo.postulante " & _
           " where cusuario = '" & pscodpostulante & "'"

    If OCON.AbreConexion Then
        Set BuscarRHSPostulanteporAprobar = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSPostulanteporAprobarPers(ByVal pscodpostulante As String) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = " select cPersCod,nNivelAlcanzado,nGrado,cDescAreaEstudio,cUniversidad,cSituacion,cOtros=ISNULL(cOtros,'') from " & Obtener_Numero_IP & ".DBWeb.dbo.RhSeleccionCurriculumDPersonales " & _
           " where cperscod='" & pscodpostulante & "'"

    If OCON.AbreConexion Then
        Set BuscarRHSPostulanteporAprobarPers = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSPostulanteporAprobarLab(ByVal pscodpostulante As String) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = " select cPersCod,cEmpresa,cDescCargo,cTipoContrato,cAreaTrabajo,cFunciones,cJefe,cCargoJefe,dFechaIni,dFechaFin,cTelefono from " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionCurriculumDLaborales " & _
           " where cperscod='" & pscodpostulante & "'"

    If OCON.AbreConexion Then
        Set BuscarRHSPostulanteporAprobarLab = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSPersonaCodigo(ByVal psDNI As String) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = " select cPersCod from DBCmactAux..PersID where cPersIDnro= '" & psDNI & "'"
           

    If OCON.AbreConexion Then
        Set BuscarRHSPersonaCodigo = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function


Public Sub ActualizaRHSeleccionPostulante(ByVal psPersCod As String, ByVal psUsuario As String, ByVal psproceso As String)
    
    Dim sSql As String

    Dim oConecta As DConecta
    
    On Error GoTo ErrorActualizaRHSeleccionPostulante
               
    sSql = " update RHSeleccionPostulante set cPersCod ='" & psPersCod & "' " & _
           " where cPostuCod='" & psUsuario & "'  and cCodProceso='" & psproceso & "'"
          
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorActualizaRHSeleccionPostulante:
    Err.Raise Err.Number, "ActualizaRHSeleccionPostulante", Err.Description

End Sub

Public Function BuscarRHSPersonaTrabajoCaja(ByVal psDNI As String) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = "select careadescripcion,cagedescripcion,crhcargodescripcion from persid  pei" & _
           " inner join rrhh rh on pei.cperscod=rh.cperscod" & _
           " inner join areas a on a.careacod=rh.careacod" & _
           " inner join agencias ag on ag.cagecod=rh.cagenciaasig" & _
           " inner join rhcargos rc on rh.cperscod=rc.cperscod and rc.drhcargofecha=(select top 1 drhcargofecha from rhcargos  where cperscod=rc.cperscod order by drhcargofecha desc )" & _
           " inner join rhcargostabla rct on rc.crhcargocod=rct.crhcargocod " & _
           " where cpersidtpo=1 and cpersidnro='" & psDNI & "'"
    
    If OCON.AbreConexion Then
        Set BuscarRHSPersonaTrabajoCaja = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSPersonaCredJudicial(ByVal psDNI As String) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = " select cj.ccodcta,nsaldcap from [128.107.2.9].dbcmact07.dbo.credcjudi cj " & _
           " inner join [128.107.2.9].dbcmact07.dbo.perscreditoconsol pc on pc.ccodcta=cj.ccodcta " & _
           " inner join dbpersona..persona p on pc.ccodpers=p.ccodpers" & _
           " where  pc.crelacta='TI' and p.cnudoci='" & Trim(psDNI) & "'"
    
    If OCON.AbreConexion Then
        Set BuscarRHSPersonaCredJudicial = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSPersonaParCaja(ByVal psDNI As String) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = " select p.cpersnombre,careadescripcion,cagedescripcion,crhcargodescripcion " & _
           " from persrelaciones pr inner join persid pei " & _
           " on pr.cpersrelacperscod=pei.cperscod " & _
           " left join persona p on p.cperscod=pr.cperscod " & _
           " left join rrhh rh on rh.cperscod=p.cperscod" & _
           " left join areas a on a.careacod=rh.careacod " & _
           " left join agencias ag on ag.cagecod=rh.cagenciaasig" & _
           " left join rhcargos rc on rh.cperscod=rc.cperscod and rc.drhcargofecha=(select top 1 drhcargofecha from rhcargos  where cperscod=rc.cperscod order by drhcargofecha desc )" & _
           " left join rhcargostabla rct on rc.crhcargocod=rct.crhcargocod" & _
           " where cpersidtpo=1 and cpersidnro='" & psDNI & "'"
    
    If OCON.AbreConexion Then
        Set BuscarRHSPersonaParCaja = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSPersonaDespedidos(ByVal psDNI As String) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = " SELECT TOP 1 P.cpersnombre,careadescripcion,cagedescripcion,crhcargodescripcion,Motivo=ISNULL(C.cConsDescripcion,''),Comentario=ISNULL(rhc.cRHRescisionComent,'') " & _
           " FROM RRHH rh INNER JOIN Persona P ON P.cPersCod= rh.cPersCod AND rh.nRHEstado=701 INNER JOIN PersId Pei ON Pei.cPersCod=P.cPersCod AND Pei.cPersidtpo=1 " & _
           " inner join areas a on a.careacod=rh.careacod inner join agencias ag on ag.cagecod=rh.cagenciaasig inner join rhcargos rc on rh.cperscod=rc.cperscod and " & _
           " rc.drhcargofecha=(select top 1 drhcargofecha from rhcargos where cperscod=rc.cperscod order by drhcargofecha desc ) " & _
           " inner join rhcargostabla rct on rc.crhcargocod=rct.crhcargocod inner join RHContrato rhc ON rhc.cPersCod=P.cPersCod " & _
           " left join Constante C ON C.nConsValor=rhc.nRHRescisionTpo AND nConsCod=6006 " & _
           " where cPersidNro='" & psDNI & "'"
    
    If OCON.AbreConexion Then
        Set BuscarRHSPersonaDespedidos = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    
    Set OCON = Nothing
End Function

Public Function BuscarRHSPostulanteVetado(ByVal psDNI As String) As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    sSql = "SELECT P.cPostuCod,Pv.cCodProceso,P.cNombre,Pv.cMotivo FROM Postulante P INNER JOIN RHPostulanteVetado Pv " & _
           " ON P.cPostuCod= Pv.cPostuCod " & _
           " WHERE P.cDNI='" & psDNI & "'"
    
    If OCON.AbreConexion Then
        Set BuscarRHSPostulanteVetado = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
End Function


Public Function BuscarRHSPostulanteEliminar(ByVal psproceso As String) As ADODB.Recordset

    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
    
    sSql = "select cPostuCod from RHSeleccionPostulante where ccodproceso ='" & psproceso & "'"
    
    If OCON.AbreConexion Then
        Set BuscarRHSPostulanteEliminar = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function EliminarRHSPostulanteProceso(ByVal pspostucod As String)
     Dim sSql1 As String
     Dim sSql2 As String
     Dim sSql3 As String
     Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminarRHSPostulanteProceso
               
    sSql1 = " delete RHSeleccionCurriculumDPersonales where cpostucod = '" & pspostucod & "' "
    sSql2 = " delete RHSeleccionCurriculumDLaborales where cpostucod='" & pspostucod & " '"
    sSql3 = " delete Postulante where cpostucod = '" & pspostucod & "'"
          
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql1
    oConecta.ConexionActiva.Execute sSql2
    oConecta.ConexionActiva.Execute sSql3
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminarRHSPostulanteProceso:
    Err.Raise Err.Number, "EliminarRHSPostulanteProceso", Err.Description

End Function

Public Function BuscarRHSCurriculumPostulanteG(ByVal pscodproceso As String) As ADODB.Recordset
'select cPostuCod from RHSeleccionPostulante where cpostucod='" &  &  " ' and ccodproceso ='"& &"'
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
    
    sSql = " select  distinct P.cPostuCod, cNombre +' '+ cApellidoPat +' '+ cApellidoMat NombreApellido ,cDesZon,cDireccion,p.cTelefono," & _
           " cOTelefono,cEmail,cFechaNac,cDNI,c1.cConsDescripcion sexo,c2.cConsDescripcion estadocivil,nHijos,cNombre, cApellidoPat,cApellidoMat, " & _
           " Departamento= rtrim(isnull((select cdeszon from dbcomunes..zonas where ccodzon like '1'+substring(z.ccodzon,2,2)+'%'),''))," & _
           " Provincia= rtrim(isnull((select cdeszon from dbcomunes..zonas where ccodzon like '2'+substring(z.ccodzon,2,4)+'%'),''))," & _
           " Distrito= rtrim(isnull((select cdeszon from dbcomunes..zonas where ccodzon like '3'+substring(z.ccodzon,2,6)+'%'),''))" & _
           " FROM postulante p " & _
           " inner join DbComunes..Zonas z on p.czona=z.cCodZon" & _
           " inner join Constante c1 on p.cSexo=c1.nconsvalor and c1.nconscod= '9054'" & _
           " inner join RHSeleccionPostulante rp on rp.cPostuCod=P.cPostuCod" & _
           " inner join Constante c2 on p.cEstadoCivil=c2.nconsvalor and c2.nconscod= '1020' " & _
           " WHERE cCodProceso ='" & Trim(pscodproceso) & "'"
    
    If OCON.AbreConexion Then
        Set BuscarRHSCurriculumPostulanteG = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSCurriculumPostulanteAprobado(ByVal pscodproceso As String) As ADODB.Recordset

    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
    
    sSql = "SELECT DISTINCT P.cPostuCod, cNombre +' '+ cApellidoPat +' '+ cApellidoMat NombreApellido ,cDesZon,cDireccion,p.cTelefono," & _
           " cOTelefono,cEmail,cFechaNac,cDNI,c1.cConsDescripcion sexo,c2.cConsDescripcion estadocivil,nHijos,cNombre, cApellidoPat,cApellidoMat, " & _
           " Departamento= rtrim(isnull((select cdeszon from dbcomunes..zonas where ccodzon like '1'+substring(z.ccodzon,2,2)+'%'),''))," & _
           " Provincia= rtrim(isnull((select cdeszon from dbcomunes..zonas where ccodzon like '2'+substring(z.ccodzon,2,4)+'%'),'')), " & _
           " Distrito= rtrim(isnull((select cdeszon from dbcomunes..zonas where ccodzon like '3'+substring(z.ccodzon,2,6)+'%'),'')) " & _
           " FROM postulante p inner join DbComunes..Zonas z on p.czona=z.cCodZon inner join Constante c1 on p.cSexo=c1.nconsvalor and c1.nconscod= '9054' " & _
           " inner join RHSeleccionPostulante rp on rp.cPostuCod=P.cPostuCod inner join Constante c2 on p.cEstadoCivil=c2.nconsvalor and c2.nconscod= '1020' " & _
           " INNER JOIN RHSeleccionEvaluacion SE ON SE.cPersCod=RP.cPostuCod AND SE.cCodProceso=RP.cCodProceso AND SE.nCodFase=6 AND SE.bGanador=1 " & _
           " WHERE RP.cCodProceso ='" & pscodproceso & "'"
    
    If OCON.AbreConexion Then
        Set BuscarRHSCurriculumPostulanteAprobado = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSCurriculumPostulanteP(ByVal cpostucod As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
    
    sSql = " select c1.cConsDescripcion,c2.cConsDescripcion,cUniversidad,cSituacion,Grado=ISNULL(c3.cConsDescripcion,''), " & _
           " cOtros=ISNULL(cOtros,'')  from dbo.RHSeleccionCurriculumDPersonales rp" & _
           " inner join Constante c1 on rp.nnivelAlcanzado=c1.nconsvalor and c1.nconscod= '9051'" & _
           " inner join Constante c2 on rp.cDescAreaEstudio=c2.nconsvalor and c2.nconscod= '9050'" & _
           " left join Constante c3 on rp.nGrado=c3.nconsvalor and c3.nconscod= '9052'" & _
           " where cpostucod='" & cpostucod & "'"
    
    If OCON.AbreConexion Then
        Set BuscarRHSCurriculumPostulanteP = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSCurriculumPostulanteL(ByVal cpostucod As String) As ADODB.Recordset
'select cPostuCod from RHSeleccionPostulante where cpostucod='" &  &  " ' and ccodproceso ='"& &"'
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
    
    sSql = " select cEmpresa,cDescCargo,cTipoContrato,c.cConsDescripcion,cFunciones,cJefe,cCargoJefe,dFechaIni,dFechaFin,cTelefono" & _
           " from RHSeleccionCurriculumDLaborales rl" & _
           " inner join Constante c  on rl.cAreaTrabajo=c.nconsvalor and c.nconscod= '9053'" & _
           " where cpostucod='" & cpostucod & "'"
    
    If OCON.AbreConexion Then
        Set BuscarRHSCurriculumPostulanteL = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function BuscarRHSCurriculumReferencias(ByVal cpostucod As String) As ADODB.Recordset
'select cPostuCod from RHSeleccionPostulante where cpostucod='" &  &  " ' and ccodproceso ='"& &"'
    
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
    
    sSql = " SELECT cPersCod, cLaboroCaja, cResidirProvincia, cBrevete, cNomPariente, cParentesco, cCargoPariente, " & _
           " cLogros, cExCorto, nPreteneconomicas, cArea " & _
           " FROM " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionCurriculumReferencias WHERE cPersCod='" & cpostucod & "'"
    
    If OCON.AbreConexion Then
        Set BuscarRHSCurriculumReferencias = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Sub InsertaRHSeleccionCriteriosEvaCurri(ByVal pscCodProceso As String, _
                        ByVal pscFiltro As String, ByVal pscCondicon As String, ByVal psValor1 As String, ByVal psValor2 As String, ByVal pnPuntaje As Long)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionCriteriosEvaCurri
    
    sSql = "INSERT INTO RHSeleccionCriteriosEvaCurri(cCodProceso,cFiltro,cCondicon,Valor1,Valor2,nPuntaje)"
    sSql = sSql & " VALUES('" & pscCodProceso & "','" & pscFiltro & "','" & pscCondicon & "','" & psValor1 & "','" & psValor2 & "'," & pnPuntaje & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionCriteriosEvaCurri:
    Err.Raise Err.Number, "InsertaRHSeleccionCriteriosEvaCurri", Err.Description
End Sub

Public Sub EliminarRHSeleccionCriteriosEvaCurri(ByVal pscCodProceso As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminarRHSeleccionCriteriosEvaCurri
    
    sSql = "DELETE RHSeleccionCriteriosEvaCurri WHERE ccodproceso = '" & pscCodProceso & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorEliminarRHSeleccionCriteriosEvaCurri:
    Err.Raise Err.Number, "EliminarRHSeleccionCriteriosEvaCurri", Err.Description
End Sub

Public Sub InsertaRHSeleccionFases(ByVal pcNomFase As String, _
                                    ByVal pcComentarioFase As String, _
                                    ByVal pbEjecuta As Integer, _
                                    ByVal pcResponsable As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionFases
    
    sSql = "INSERT INTO RHSeleccionFases(cNomFase, cComentarioFase, bEjecuta, cResponsable)"
    sSql = sSql & " VALUES('" & pcNomFase & "','" & pcComentarioFase & "'," & pbEjecuta & ",'" & pcResponsable & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionFases:
    Err.Raise Err.Number, "InsertaRHSeleccionFases", Err.Description
End Sub

Public Sub ModificaRHSeleccionFases(ByVal pnCodFase As Long, _
                                    ByVal pcNomFase As String, _
                                    ByVal pcComentarioFase As String, _
                                    ByVal pbEjecuta As Integer, _
                                    ByVal pcResponsable As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHSeleccionFases
    sSql = "UPDATE RHSeleccionFases"
    sSql = sSql & " SET cNomFase='" & pcNomFase & _
                        "',cComentarioFase='" & pcComentarioFase & _
                        "',bEjecuta=" & pbEjecuta & _
                        ",cResponsable='" & pcResponsable & _
                        "' WHERE  nCodFase=" & pnCodFase

                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHSeleccionFases:
    Err.Raise Err.Number, "ModificaRHSeleccionFases", Err.Description
End Sub

Public Function EliminaRHSeleccionFases(ByVal pnCodFase As Long)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionFases
    sSql = "DELETE FROM RHSeleccionFases" '
    sSql = sSql & " WHERE nCodFase=" & pnCodFase
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionFases:
    Err.Raise Err.Number, "EliminaRHSeleccionFases", Err.Description
End Function

Public Sub InsertaDBWebRHSeleccionProceso(ByVal pcCodProceso As String)
                    
Dim sSql As String
Dim oConecta As DConecta
Dim strDescriProceso As String
Dim intTipoProceso As Integer
Dim rs As New ADODB.Recordset
Dim strFechaIni As String
Dim strFechaFin As String
    
Dim cCodPublicacion As String
    
    On Error GoTo ErrorInsertaDBWebRHSeleccionProceso
    
    sSql = "SELECT TOP 1 TP.nConsValor, Codigo= cRHCargoDescripcion+ '-' + " & _
          " CASE WHEN Ag.cAgeCod IN('01','02','03','05','07','09') THEN 'TRUJILLO' " & _
          " WHEN Ag.cAgeCod='10' THEN 'VIRU' WHEN Ag.cAgeCod='06' THEN 'CHEPEN' " & _
          " WHEN Ag.cAgeCod='08' THEN 'HUARAZ' WHEN Ag.cAgeCod IN('12','15','16')THEN 'CHICLAYO' " & _
          " WHEN Ag.cAgeCod='13' THEN 'CAJAMARCA' WHEN Ag.cAgeCod='14' THEN 'HUARAZ' " & _
          " WHEN Ag.cAgeCod='17' THEN 'JAEN' WHEN Ag.cAgeCod='18' THEN 'LIMA' " & _
          " WHEN Ag.cAgeCod='19' THEN 'CHOCOPE' WHEN Ag.cAgeCod='20' THEN 'OTUZCO' " & _
          " WHEN Ag.cAgeCod='21' THEN 'TUMBES' END,cCodProcesoPublicacion=ISNULL(cCodProcesoPublicacion,'') FROM RHSeleccionProceso SP INNER JOIN Constante TP ON " & _
          " TP.nConsValor=SP.nTipoProceso AND TP.nConsCod=6003 INNER JOIN RHSeleccionRequerimientoProceso SRP " & _
          " ON SP.cCodProceso=SRP.cCodProceso INNER JOIN RHSeleccionRequerimiento SR ON SRP.cCodRequerimiento=SR.cCodRequerimiento " & _
          " INNER JOIN RHCargos CAR ON CAR.cRHCargoCod=SR.cCodCargoSol INNER JOIN RHCargosTabla CAT ON CAT.cRHCargoCod=CAR.cRHCargoCod " & _
          " INNER JOIN Agencias AG ON SR.cCodAgeSol=AG.cAgeCod " & _
          " WHERE SP.cCodProceso='" & pcCodProceso & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set rs = oConecta.CargaRecordSet(sSql)
        If Not rs.EOF Then
            intTipoProceso = rs("nConsValor")
            strDescriProceso = rs("Codigo")
            
            cCodPublicacion = rs("cCodProcesoPublicacion")
        End If
        
        sSql = "SELECT dFechaIni,dFechaFin FROM RHSeleccionFase WHERE nCodFase=2 AND cCodProceso='" & pcCodProceso & "'"
        Set rs = oConecta.CargaRecordSet(sSql)
        If Not rs.EOF Then
            strFechaIni = Format(rs("dFechaIni"), gsFormatoFecha)
            strFechaFin = Format(rs("dFechaFin"), gsFormatoFecha)
        End If
    End If
    
    If intTipoProceso = 0 Then
        sSql = "INSERT INTO " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProceso(cCodProceso,cDescripcion,nEstado,dFechaIni,dFechaFin,cCodPublicacion)"
        sSql = sSql & " VALUES('" & pcCodProceso & "','" & strDescriProceso & "',1,'" & strFechaIni & "','" & strFechaFin & "','" & cCodPublicacion & "')"
        oConecta.ConexionActiva.Execute sSql
    End If
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaDBWebRHSeleccionProceso:
    Err.Raise Err.Number, "InsertaDBWebRHSeleccionProceso", Err.Description
End Sub

Public Sub InsertaRHSeleccionFaseCostosDet(ByVal pnCodFase As Long, _
                                    ByVal pcCodProceso As String, _
                                    ByVal pnItem As Long, _
                                    ByVal pnMonto As Double, _
                                    ByVal pdFecha As Date, _
                                    ByVal pcGlosa As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionFaseCostosDet
    
    sSql = "INSERT INTO RHSeleccionFaseCostosDet(nCodFase, cCodProceso, nItem, nMonto,dFecha,cGlosa)"
    sSql = sSql & " VALUES(" & pnCodFase & ",'" & pcCodProceso & "'," & pnItem & "," & pnMonto & ",'" & Format(pdFecha, gsFormatoFecha) & "','" & pcGlosa & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionFaseCostosDet:
    Err.Raise Err.Number, "InsertaRHSeleccionFaseCostosDet", Err.Description
End Sub

'Public Sub ModificaRHSeleccionFaseCostosDet(ByVal pnCodFase As Long, _
'                                    ByVal pcCodProceso As String, _
'                                    ByVal pnItem As Long, _
'                                    ByVal pnMonto As Double, _
'                                    ByVal pdFecha As Date, _
'                                    ByVal pcGlosa As String)
'
'Dim sSql As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorModificaRHSeleccionFaseCostosDet
'    sSql = "UPDATE RHSeleccionFaseCostosDet"
'    sSql = sSql & " SET cNomFase='" & pcNomFase & _
'                        "',cComentarioFase='" & pcComentarioFase & _
'                        "',bEjecuta=" & pbEjecuta & _
'                        ",cResponsable='" & pcResponsable & _
'                        "' WHERE  nCodFase=" & pnCodFase
'
'
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSql
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Sub
'
'ErrorModificaRHSeleccionFaseCostosDet:
'    Err.Raise Err.Number, "ModificaRHSeleccionFaseCostosDet", Err.Description
'End Sub

Public Function EliminaRHSeleccionFaseCostosDet(ByVal pnCodFase As Long, _
                                    ByVal pcCodProceso As String, _
                                    Optional ByVal pnItem As Long = 0)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionFaseCostosDet
    sSql = "DELETE FROM RHSeleccionFaseCostosDet" '
    sSql = sSql & " WHERE nCodFase=" & pnCodFase & " AND cCodProceso='" & pcCodProceso & IIf(pnItem = 0, "'", "' AND nItem=" & pnItem)
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionFaseCostosDet:
    Err.Raise Err.Number, "EliminaRHSeleccionFaseCostosDet", Err.Description
End Function

Public Sub InsertaRHSeleccionComiteCambio(ByVal pcCodProceso As String, _
                                            ByVal pnCodFase As Long, _
                                            ByVal pcPersCodTitu As String, _
                                            ByVal pcPersCodSup As String, _
                                            ByVal pcDocumento As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionComiteCambio
    sSql = "INSERT INTO RHSeleccionComiteCambio(cCodProceso,nCodFase,cPersCodTitu,cPersCodSup,cDocumento)"
    sSql = sSql & " VALUES('" & pcCodProceso & "'," & pnCodFase & ",'" & pcPersCodTitu & "','" & pcPersCodSup & "','" & pcDocumento & "')"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    
    sSql = "UPDATE RHSeleccionProcesoComite SET cPersCodTitu='" & pcPersCodSup & "' WHERE cPersCod='" & pcPersCodTitu & "'" 'Volverlo Suplente
    oConecta.ConexionActiva.Execute sSql
    
    sSql = "UPDATE RHSeleccionProcesoComite SET cPersCodTitu= NULL WHERE cPersCod='" & pcPersCodSup & "'"   'Volverlo Titular
    oConecta.ConexionActiva.Execute sSql
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionComiteCambio:
    Err.Raise Err.Number, "InsertaRHSeleccionComiteCambio", Err.Description
End Sub

Public Sub InsertaRHSeleccionRequerimiento(ByVal pcCodRequerimiento As String, _
                                            ByVal pcCodArea As String, _
                                            ByVal pcPersCod As String, _
                                            ByVal pcCodAreaSol As String, _
                                            ByVal pcCodCargoSol As String, _
                                            ByVal pcCodAgeSol As String, _
                                            ByVal pnNumVacantes As Integer, _
                                            ByVal pdFechaIni As Date, _
                                            ByVal pdFechaFin As Date, _
                                            ByVal pnEstado As Integer, _
                                            ByVal pcUltimaActualizacion As String, _
                                            ByVal pcCargoSolNuevo As String, _
                                            ByVal pcComentario As String, _
                                            ByVal pcJustificacion As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionRequerimiento
    
    sSql = "INSERT INTO RHSeleccionRequerimiento(cCodRequerimiento, cCodArea, cPersCod, cCodAreaSol, cCodCargoSol, cCodAgeSol, nNumVacantes, dFechaIni, dFechaFin, nEstado, cUltimaActualizacion, cCargoSolNuevo, cComentario, cJustificacion)"
    sSql = sSql & " VALUES('" & pcCodRequerimiento & "','" & pcCodArea & "','" & pcPersCod & "','" & pcCodAreaSol & "','" & pcCodCargoSol & "','" & pcCodAgeSol & "'," & pnNumVacantes & ",'" & Format(pdFechaIni, gsFormatoFecha) & "','" & Format(pdFechaFin, gsFormatoFecha) & "'," & pnEstado & ",'" & pcUltimaActualizacion & "','" & pcCargoSolNuevo & "','" & pcComentario & "','" & pcJustificacion & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionRequerimiento:
    Err.Raise Err.Number, "InsertaRHSeleccionRequerimiento", Err.Description
End Sub

Public Sub ModificaRHSeleccionRequerimiento(ByVal pcCodRequerimiento As String, _
                                            Optional ByVal pcCodArea As String = "", _
                                            Optional ByVal pcPersCod As String = "", _
                                            Optional ByVal pcCodAreaSol As String = "", _
                                            Optional ByVal pcCodCargoSol As String = "", _
                                            Optional ByVal pcCodAgeSol As String = "", _
                                            Optional ByVal pnNumVacantes As Integer = -1, _
                                            Optional ByVal pdFechaIni As Date = 0, _
                                            Optional ByVal pdFechaFin As Date = 0, _
                                            Optional ByVal pnEstado As Integer = -1, _
                                            Optional ByVal pcUltimaActualizacion As String = "", _
                                            Optional ByVal pcCargoSolNuevo As String = "", _
                                            Optional ByVal pcComentario As String = "", _
                                            Optional ByVal pcJustificacion As String = "", _
                                            Optional ByVal pcMotivoRechCAP As String = "", _
                                            Optional ByVal pnTipoProceso As Integer = -1)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHSeleccionRequerimiento
    sSql = "UPDATE RHSeleccionRequerimiento SET "
    sSql = sSql & IIf(pcCodArea <> "", " cCodArea='" & pcCodArea & "',", "") & _
                IIf(pcPersCod <> "", " cPersCod='" & pcPersCod & "',", "") & _
                IIf(pcCodAreaSol <> "", " cCodAreaSol='" & pcCodAreaSol & "',", "") & _
                IIf(pcCodCargoSol <> "", " cCodCargoSol='" & pcCodCargoSol & "',", "") & _
                IIf(pcCodAgeSol <> "", " cCodAgeSol='" & pcCodAgeSol & "',", "") & _
                IIf(pnNumVacantes <> -1, " nNumVacantes=" & pnNumVacantes & ",", "") & _
                IIf(pdFechaIni <> 0, " dFechaIni='" & Format(pdFechaIni, gsFormatoFecha) & "',", "") & _
                IIf(pdFechaFin <> 0, " dFechaFin='" & Format(pdFechaFin, gsFormatoFecha) & "',", "") & _
                IIf(pnEstado <> -1, " nEstado=" & pnEstado & ",", "") & _
                IIf(pcUltimaActualizacion <> "", " cUltimaActualizacion='" & pcUltimaActualizacion & "',", "") & _
                IIf(pcCargoSolNuevo <> "", " cCargoSolNuevo='" & pcCargoSolNuevo & "',", "") & _
                IIf(pcComentario <> "", " cComentario='" & pcComentario & "',", "") & _
                IIf(pcJustificacion <> "", " cJustificacion='" & pcJustificacion & "',", "") & _
                IIf(pcMotivoRechCAP <> "", " cMotivoRechCAP='" & pcMotivoRechCAP & "',", "") & _
                IIf(pnTipoProceso <> -1, " nTipoProceso=" & pnTipoProceso & ",", "")
    
    'Eliminamos la ultima Coma(,)
    sSql = Mid(sSql, 1, Len(sSql) - 1) & " WHERE cCodRequerimiento='" & pcCodRequerimiento & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHSeleccionRequerimiento:
    Err.Raise Err.Number, "ModificaRHSeleccionRequerimiento", Err.Description
End Sub

Public Function EliminaRHSeleccionRequerimiento(ByVal pcCodRequerimiento As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionRequerimiento
    sSql = "DELETE FROM RHSeleccionRequerimiento" '
    sSql = sSql & " WHERE cCodRequerimiento='" & pcCodRequerimiento & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHSeleccionRequerimiento:
    Err.Raise Err.Number, "EliminaRHSeleccionRequerimiento", Err.Description
End Function

Public Function Obtener_Nuevo_CodigoRequerimiento() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Nuevo_CodigoRequerimiento
    
    sSql = "SELECT Codigo=ISNULL(MAX(cCodRequerimiento),0)+1 from rhseleccionrequerimiento"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Nuevo_CodigoRequerimiento = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Nuevo_CodigoRequerimiento:
    Err.Raise Err.Number, "Obtener_Nuevo_CodigoRequerimiento", Err.Description
End Function

Public Function Obtener_Codigo_PublicacionWeb(pcCodProceso As String) As String

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
Dim nNumRequerimientos As Integer
    
    On Error GoTo ErrorObtener_Codigo_PublicacionWeb
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    
        sSql = "SELECT NumReq=COUNT(cCodRequerimiento) FROM RHSeleccionRequerimientoProceso WHERE cCodProceso='" & pcCodProceso & "'"
    
        Set rs = oConecta.CargaRecordSet(sSql)
        If Not rs.EOF Then
            nNumRequerimientos = rs("NumReq")
        Else
            nNumRequerimientos = 0
        End If
        'oConecta.ConexionActiva.Execute sSql
        sSql = "SELECT TOP 1 Codigo=LEFT(cRHCargoDescripcion,3)+ REPLACE (   CASE WHEN LEFT(cRHCargoDescripcion,3)= " & _
               " SUBSTRING( REPLACE( REPLACE( REPLACE(cRHCargoDescripcion,'SUB GERENTE',' '),' DE','') ,' ','$'), " & _
               " PATINDEX('%$%',REPLACE( REPLACE(REPLACE(cRHCargoDescripcion,'SUB GERENTE',' '),' DE','') ,' ','$'))   +1 " & _
               " ,3 ) THEN '' Else SUBSTRING( REPLACE( REPLACE(REPLACE(cRHCargoDescripcion,'SUB GERENTE',' '),' DE','') ,' ','$'), " & _
               " PATINDEX('%$%', REPLACE( REPLACE(REPLACE(cRHCargoDescripcion,'SUB GERENTE',' '),' DE','') ,' ','$'))   +1 " & _
               " ,3 ) End ,'$','')+ SUBSTRING(SP.cCodProceso,5,LEN(RTRIM(SP.cCodProceso))-4) " & _
               IIf(nNumRequerimientos = 1, "+ CASE WHEN Ag.cAgeCod IN('01','02','03','05','07','09') THEN 'TRU' " & _
               " WHEN Ag.cAgeCod='10' THEN 'VIR' WHEN Ag.cAgeCod='06' THEN 'CHE'  WHEN Ag.cAgeCod='08' THEN 'HUA' WHEN Ag.cAgeCod IN('12','15','16')THEN 'CHI'" & _
               "  WHEN Ag.cAgeCod='13' THEN 'CAJ' WHEN Ag.cAgeCod='14' THEN 'HUA'  WHEN Ag.cAgeCod='17' THEN 'JAE' WHEN Ag.cAgeCod='18' THEN 'LIM' " & _
               "  WHEN Ag.cAgeCod='19' THEN 'CHO' WHEN Ag.cAgeCod='20' THEN 'OTU'  WHEN Ag.cAgeCod='21' THEN 'TUM' End + ", "+") & _
               " CONVERT(varchar(4),YEAR(SP.dFechaIni)) " & _
               " FROM RHSeleccionProceso SP INNER JOIN Constante TP ON TP.nConsValor=SP.nTipoProceso AND TP.nConsCod=6003 " & _
               " INNER JOIN RHSeleccionRequerimientoProceso SRP ON SP.cCodProceso=SRP.cCodProceso " & _
               " INNER JOIN RHSeleccionRequerimiento SR ON SRP.cCodRequerimiento=SR.cCodRequerimiento " & _
               " INNER JOIN RHCargos CAR ON CAR.cRHCargoCod=SR.cCodCargoSol INNER JOIN RHCargosTabla CAT ON CAT.cRHCargoCod=CAR.cRHCargoCod " & _
               " INNER JOIN Agencias AG ON SR.cCodAgeSol=AG.cAgeCod " & _
               " WHERE SP.cCodProceso='" & pcCodProceso & "'"
        
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_Codigo_PublicacionWeb = ""
        Else
            Obtener_Codigo_PublicacionWeb = CStr(rs("Codigo"))
        End If
    
    End If
    
    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Codigo_PublicacionWeb:
    Err.Raise Err.Number, "Obtener_Codigo_PublicacionWeb", Err.Description
End Function

Public Function Obtener_Descripcion_Proceso(pcCodProceso As String) As String

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
Dim nNumRequerimientos As Integer
    
    On Error GoTo ErrorObtener_Descripcion_Proceso
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    
        sSql = "SELECT NumReq=COUNT(cCodRequerimiento) FROM RHSeleccionRequerimientoProceso WHERE cCodProceso='" & pcCodProceso & "'"
    
        Set rs = oConecta.CargaRecordSet(sSql)
        If Not rs.EOF Then
            nNumRequerimientos = rs("NumReq")
        Else
            nNumRequerimientos = 0
        End If
        'oConecta.ConexionActiva.Execute sSql
        sSql = "SELECT TOP 1 " & _
               " Codigo= RTRIM(TP.cConsDescripcion)+' PARA '+ cRHCargoDescripcion +' ' " & _
               IIf(nNumRequerimientos = 1, "+ CASE WHEN Ag.cAgeCod IN('01','02','03','05','07','09') THEN 'TRUJILLO' " & _
               " WHEN Ag.cAgeCod='10' THEN 'VIRU' WHEN Ag.cAgeCod='06' THEN 'CHEPEN' " & _
               " WHEN Ag.cAgeCod='08' THEN 'HUAMACHUCO' WHEN Ag.cAgeCod IN('12','15','16')THEN 'CHICLAYO' " & _
               " WHEN Ag.cAgeCod='13' THEN 'CAJAMARCA' WHEN Ag.cAgeCod='14' THEN 'HUARAZ' " & _
               " WHEN Ag.cAgeCod='17' THEN 'JAEN' WHEN Ag.cAgeCod='18' THEN 'LIMA' " & _
               " WHEN Ag.cAgeCod='19' THEN 'CHOCOPE' WHEN Ag.cAgeCod='20' THEN 'OTUZCO' " & _
               " WHEN Ag.cAgeCod='21' THEN 'TUMBES' END ", "") & _
               " FROM RHSeleccionProceso SP INNER JOIN Constante TP ON TP.nConsValor=SP.nTipoProceso AND TP.nConsCod=6003 " & _
               " INNER JOIN RHSeleccionRequerimientoProceso SRP ON SP.cCodProceso=SRP.cCodProceso " & _
               " INNER JOIN RHSeleccionRequerimiento SR ON SRP.cCodRequerimiento=SR.cCodRequerimiento " & _
               " INNER JOIN RHCargos CAR ON CAR.cRHCargoCod=SR.cCodCargoSol " & _
               " INNER JOIN RHCargosTabla CAT ON CAT.cRHCargoCod=CAR.cRHCargoCod " & _
               " INNER JOIN Agencias AG ON SR.cCodAgeSol=AG.cAgeCod " & _
               " WHERE SP.cCodProceso='" & pcCodProceso & "'"
        
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_Descripcion_Proceso = ""
        Else
            Obtener_Descripcion_Proceso = CStr(rs("Codigo"))
        End If
    
    End If
    
    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Descripcion_Proceso:
    Err.Raise Err.Number, "Obtener_Descripcion_Proceso", Err.Description
End Function

Public Function Obtener_Postulantes_ExperienciaLaboral(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Postulantes_ExperienciaLaboral
    
    sSql = "select  distinct P.cPostuCod, cApellidoPat +' '+ cApellidoMat+ ' '+cNombre  NombreApellido ,p.cTelefono,cOTelefono,SE.nNota, " & _
            "Profesion=ISNULL(Pro.cConsDescripcion,''),P.cDNI,Edad=DATEDIFF(YEAR,cFechaNac,getdate()),Estado=ISNULL(c2.cConsDescripcion,''),P.cDireccion " & _
            ",Ciudad=ISNULL(CIU.cDesCiudad,'')" & _
           " FROM Postulante p inner join DbComunes..Zonas z on p.czona=z.cCodZon " & _
           " inner join Constante c1 on p.cSexo=c1.nconsvalor and c1.nconscod= '9054'" & _
           " inner join RHSeleccionPostulante rp on rp.cPostuCod=P.cPostuCod AND rp.cCodProceso='" & pcCodProceso & "'" & _
           " inner join Constante c2 on p.cEstadoCivil=c2.nconsvalor and c2.nconscod= '1020'" & _
           " inner join RHSeleccionEvaluacion SE ON SE.cPerscod=P.cPostucod AND nCodFase=6 " & _
           "    LEFT JOIN RHSeleccionCurriculumDPersonales CP ON CP.cPostuCod = P.cPostuCod " & _
           "    AND CP.nNivelAlcanzado=(SELECT MIN(nNivelAlcanzado)FROM RHSeleccionCurriculumDPersonales " & _
           "    WHERE cPostuCod=P.cPostuCod)" & _
           "    LEFT JOIN Constante PRO ON PRO.nConsValor= CP.cDescAreaEstudio AND PRO.nConsCod=9050 " & _
           "    LEFT JOIN " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProcesoCiudad PC ON PC.cCodProceso=SE.cCodProceso" & _
           "    COLLATE SQL_Latin1_General_CP437_BIN " & _
           "    LEFT JOIN " & Obtener_Numero_IP & ".DBWeb.dbo.RHCiudades CIU ON PC.cCodCiudad=CIU.cCodCiudad " & _
           " WHERE SE.cCodProceso='" & pcCodProceso & _
           "' ORDER BY nNota DESC,NombreApellido"

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Postulantes_ExperienciaLaboral = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Postulantes_ExperienciaLaboral:
    Err.Raise Err.Number, "Obtener_Postulantes_ExperienciaLaboral", Err.Description
End Function

Public Function Obtener_ExperienciaLaboral_Postulantes(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_ExperienciaLaboral_Postulantes
    sSql = "RHObtener_ExperienciaLaboral_Postulantes '" & pcCodProceso & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_ExperienciaLaboral_Postulantes = oConecta.Ejecutar(sSql)
    End If
    'oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_ExperienciaLaboral_Postulantes:
    Err.Raise Err.Number, "Obtener_ExperienciaLaboral_Postulantes", Err.Description
End Function

'Se agrego para el manejo de las ciudades
Public Function Obtener_DatosProcesoWeb_x_Estado(Optional ByVal pnEstado As Long = 500) As ADODB.Recordset
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorObtener_DatosProcesoWeb_x_Estado
    
    If pnEstado <> 500 Then
        sSql = "SELECT cCodProceso,cDescripcion,1 FROM " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProceso "
        sSql = sSql & " WHERE nEstado=" & pnEstado
    Else
        sSql = "SELECT cCodProceso,cDescripcion,1 FROM " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProceso "
    End If
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_DatosProcesoWeb_x_Estado = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_DatosProcesoWeb_x_Estado:
    Err.Raise Err.Number, "Obtener_DatosProcesoWeb_x_Estado", Err.Description
End Function

Public Function Cargar_DatosProcesoWeb(ByVal pcCodProceso As String) As ADODB.Recordset
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorCargar_DatosProcesoWeb
    
    sSql = "SELECT cCodPublicacion=ISNULL(cCodPublicacion,''),nCiudades=ISNULL(nCiudades,0) FROM " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProceso WHERE cCodProceso='" & pcCodProceso & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Cargar_DatosProcesoWeb = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorCargar_DatosProcesoWeb:
    Err.Raise Err.Number, "Cargar_DatosProcesoWeb", Err.Description
End Function

Public Function Cargar_CiudadesProcesoWeb(ByVal pcCodProceso As String) As ADODB.Recordset
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorCargar_CiudadesProcesoWeb
    
    sSql = "SELECT C.cCodCiudad,C.cDesCiudad FROM " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProceso SP INNER JOIN " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProcesoCiudad SPC" & _
           " ON SP.cCodProceso= SPC.cCodProceso INNER JOIN " & Obtener_Numero_IP & ".DBWeb.dbo.RHCiudades C " & _
           " ON SPC.cCodCiudad=C.cCodCiudad WHERE SP.cCodProceso='" & pcCodProceso & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Cargar_CiudadesProcesoWeb = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorCargar_CiudadesProcesoWeb:
    Err.Raise Err.Number, "Cargar_CiudadesProcesoWeb", Err.Description
End Function

Public Sub InsertaRHSeleccionProcesoCiudad(ByVal pcCodProceso As String, _
                        ByVal pcCodCiudad As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHSeleccionProcesoCiudad
    
    sSql = "INSERT INTO " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProcesoCiudad(cCodProceso, cCodCiudad)"
    sSql = sSql & " VALUES('" & pcCodProceso & "','" & pcCodCiudad & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHSeleccionProcesoCiudad:
    Err.Raise Err.Number, "InsertaRHSeleccionProcesoCiudad", Err.Description
End Sub

Public Sub EliminaRHSeleccionProcesoCiudad(ByVal pcCodProceso As String, _
                        ByVal pcCodCiudad As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHSeleccionProcesoCiudad
    sSql = "DELETE FROM " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProcesoCiudad"
    sSql = sSql & " WHERE cCodProceso='" & pcCodProceso & _
                        "' AND cCodCiudad='" & pcCodCiudad & "'"
 
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorEliminaRHSeleccionProcesoCiudad:
    Err.Raise Err.Number, "EliminaRHSeleccionProcesoCiudad", Err.Description
End Sub

Public Function Actualizar_DatosProcesoWeb(ByVal pcCodProceso As String, _
                                           ByVal pcCodPublicacion As String, _
                                           ByVal pnCiudades As Integer) As ADODB.Recordset
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorActualizar_DatosProcesoWeb
    
    sSql = "UPDATE " & Obtener_Numero_IP & ".DBWeb.dbo.RHSeleccionProceso " & _
           "SET cCodPublicacion='" & pcCodPublicacion & "'" & _
           ", nCiudades=" & pnCiudades & _
           " WHERE cCodProceso='" & pcCodProceso & "'"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        oConecta.ConexionActiva.Execute sSql
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorActualizar_DatosProcesoWeb:
    Err.Raise Err.Number, "Actualizar_DatosProcesoWeb", Err.Description
End Function

Public Function Cargar_CiudadesWeb() As ADODB.Recordset
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorCargar_CiudadesWeb
    
    sSql = "SELECT cCodCiudad,cDesCiudad,1 FROM " & Obtener_Numero_IP & ".DBWeb.dbo.RHCiudades"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Cargar_CiudadesWeb = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorCargar_CiudadesWeb:
    Err.Raise Err.Number, "Cargar_CiudadesWeb", Err.Description
End Function


