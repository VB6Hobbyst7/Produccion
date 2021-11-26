VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCapacitacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InsertaRHCapacitacionCursos(ByVal pnCursoCod As Long, _
                                ByVal pcCursoDescripcion As String, _
                                ByVal pcAreaCod As String, _
                                ByVal pnCursoClasificacion As Long, _
                                ByVal pnCursoTipoEvento As Long, _
                                ByVal pnCursoTipoPublico As Long, _
                                ByVal pnCursoEstado As Long)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHCapacitacionCursos
    
    sSql = "INSERT INTO RHCapacitacionCursos(cCursoDescripcion,cAreaCod,nCursoClasificacion,nCursoTipoEvento,nCursoTipoPublico,nCursoEstado) "
    sSql = sSql & " VALUES('" & pcCursoDescripcion & "','" & pcAreaCod & "'," & pnCursoClasificacion & "," & pnCursoTipoEvento & "," & pnCursoTipoPublico & "," & pnCursoEstado & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHCapacitacionCursos:
    Err.Raise Err.Number, "InsertaRHCapacitacionCursos", Err.Description
End Sub

Public Sub ModificaRHCapacitacionCursos(ByVal pnCursoCod As Long, _
                                ByVal pcCursoDescripcion As String, _
                                ByVal pcAreaCod As String, _
                                ByVal pnCursoClasificacion As Long, _
                                ByVal pnCursoTipoEvento As Long, _
                                ByVal pnCursoTipoPublico As Long, _
                                ByVal pnCursoEstado As Long)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHCapacitacionCursos
    sSql = "UPDATE RHCapacitacionCursos"
    sSql = sSql & " SET cCursoDescripcion ='" & pcCursoDescripcion & _
                        "',cAreaCod ='" & pcAreaCod & _
                        "',nCursoClasificacion=" & pnCursoClasificacion & _
                        ",nCursoTipoEvento=" & pnCursoTipoEvento & _
                        ",nCursoTipoPublico=" & pnCursoTipoPublico & _
                        " WHERE nCursoCod=" & pnCursoCod
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHCapacitacionCursos:
    Err.Raise Err.Number, "ModificaRHCapacitacionCursos", Err.Description
End Sub

Public Function EliminaRHCapacitacionCursos(ByVal pnCursoCod As Long)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHCapacitacionCursos
    sSql = "DELETE FROM RHCapacitacionCursos " '
    sSql = sSql & " WHERE  nCursoCod=" & pnCursoCod
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHCapacitacionCursos:
    Err.Raise Err.Number, "EliminaRHCapacitacionCursos", Err.Description
End Function

Public Function Obtener_Cursos_Capacitacion(Optional ByVal pbTodos As Boolean = False, _
                                            Optional ByVal pbEstructuraArbol As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Cursos_Capacitacion
    
    If pbEstructuraArbol = True Then
        sSql = "SELECT nCursoCod, Descri= cCursoDescripcion+' - '+ A.cAreaDescripcion,1" & _
           " FROM RHCapacitacionCursos CC INNER JOIN Areas A ON CC.cAreaCod=A.cAreaCod " & _
           " INNER JOIN Constante CLA ON CLA.nConsValor=nCursoClasificacion AND CLA.nConsCod=9071 " & _
           " INNER JOIN Constante TIE ON TIE.nConsValor=nCursoTipoEvento AND TIE.nConsCod=9072 " & _
           " INNER JOIN Constante TIP ON TIP.nConsValor=nCursoTipoPublico AND TIP.nConsCod=9073" & _
           " WHERE nCursoCod IN( SELECT nCursoCod FROM RHCapacitacionPrograma WHERE nProgCursoAnio=year(getdate()) )"
    Else
    
        If pbTodos = True Then
            sSql = "SELECT nCursoCod, cCursoDescripcion, A.cAreaCod,A.cAreaDescripcion,nCursoClasificacion ,Clasificacion=ISNULL(CLA.cConsDescripcion,''), " & _
                " nCursoTipoEvento,TipoEvento=ISNULL(TIE.cConsDescripcion,''), nCursoTipoPublico,TipoPublico=ISNULL(TIP.cConsDescripcion,''), nCursoEstado " & _
                " FROM RHCapacitacionCursos CC INNER JOIN Areas A ON CC.cAreaCod=A.cAreaCod " & _
                " INNER JOIN Constante CLA ON CLA.nConsValor=nCursoClasificacion AND CLA.nConsCod=9071 " & _
                " INNER JOIN Constante TIE ON TIE.nConsValor=nCursoTipoEvento AND TIE.nConsCod=9072 " & _
                " INNER JOIN Constante TIP ON TIP.nConsValor=nCursoTipoPublico AND TIP.nConsCod=9073"
        Else
            sSql = "SELECT nCursoCod, cCursoDescripcion, A.cAreaCod,A.cAreaDescripcion,nCursoClasificacion ,Clasificacion=ISNULL(CLA.cConsDescripcion,''), " & _
                " nCursoTipoEvento,TipoEvento=ISNULL(TIE.cConsDescripcion,''), nCursoTipoPublico,TipoPublico=ISNULL(TIP.cConsDescripcion,''), nCursoEstado " & _
                " FROM RHCapacitacionCursos CC INNER JOIN Areas A ON CC.cAreaCod=A.cAreaCod " & _
                " INNER JOIN Constante CLA ON CLA.nConsValor=nCursoClasificacion AND CLA.nConsCod=9071 " & _
                " INNER JOIN Constante TIE ON TIE.nConsValor=nCursoTipoEvento AND TIE.nConsCod=9072 " & _
                " INNER JOIN Constante TIP ON TIP.nConsValor=nCursoTipoPublico AND TIP.nConsCod=9073" & _
                " WHERE nCursoCod=(SELECT nCursoCod FROM RHCapacitacionPrograma WHERE nProgCursoAnio=year(getdate()) )"
        End If
    End If
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Cursos_Capacitacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Cursos_Capacitacion:
    Err.Raise Err.Number, "Obtener_Cursos_Capacitacion", Err.Description
End Function

Public Function Obtener_Cursos_asignados_Programa(pnCursoCod As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Cursos_asignados_Programa
    
    sSql = "SELECT nCursoCod, cCursoDescripcion  " & _
           " FROM RHCapacitacionPrograma WHERE nCursoCod=" & pnCursoCod
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Cursos_asignados_Programa = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Cursos_asignados_Programa:
    Err.Raise Err.Number, "Obtener_Cursos_asignados_Programa", Err.Description
End Function

Public Function Obtener_Siguiente_CodCurso() As Long

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrorObtener_Siguiente_CodCurso
    sSql = "SELECT Siguiente=ISNULL(MAX(nCursoCod),0)+1  FROM RHCapacitacionCursos "
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_Siguiente_CodCurso = 1
        Else
            Obtener_Siguiente_CodCurso = rs("Siguiente") '+ 1
        End If
    End If
    
    oConecta.CierraConexion
    
    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Siguiente_CodCurso:
    Err.Raise Err.Number, "Obtener_Siguiente_CodCurso", Err.Description
    
End Function

Public Sub InsertaRHCapacitacionPrograma(ByVal pnProgNro As Integer, _
                                ByVal pnCursoCod As Long, _
                                ByVal pnProgCursoAnio As Integer, _
                                ByVal pcProgCursoObjetivo As String, _
                                ByVal pdProgCursoFechaProbableInicio As Date, _
                                ByVal pdProgCursoFechaProbableTermino As Date, _
                                ByVal pcProgCursoPersCodADictar As String, _
                                ByVal pnProgCursoProgramado As Integer, _
                                ByVal pnProgCursoNroHoras As Integer, _
                                ByVal pnProgCursoModulo As Integer, _
                                ByVal pnProgCursoNroParticipantes As Integer, _
                                ByVal pnProgCursoTipoMoneda As Integer, _
                                ByVal pnProgCursoCostoEstimado As Double, _
                                ByVal pnProgCursoNacionalInternacional As Integer, _
                                ByVal pnProgCursoEstado As Integer)
Dim sSql As String
Dim oConecta As DConecta
Dim oCapacitacion As New NCapacitacion
    
    On Error GoTo ErrorInsertaRHCapacitacionPrograma
    Set oConecta = New DConecta
    oConecta.AbreConexion
    'oConecta.BeginTrans
    
    sSql = "INSERT INTO RHCapacitacionPrograma(nProgNro, nCursoCod, nProgCursoAnio, cProgCursoObjetivo," & _
                    "dProgCursoFechaProbableInicio, dProgCursoFechaProbableTermino,cProgCursoPersCodADictar," & _
                    "nProgCursoProgramado,nProgCursoNroHoras, nProgCursoModulo,nProgCursoNroParticipantes, " & _
                    "nProgCursoTipoMoneda,nProgCursoCostoEstimado, nProgCursoNacionalInternacional,nProgCursoEstado) " & _
            " VALUES(" & pnProgNro & "," & pnCursoCod & "," & pnProgCursoAnio & ",'" & pcProgCursoObjetivo & "','" & _
            Format(pdProgCursoFechaProbableInicio, gsFormatoFecha) & "','" & Format(pdProgCursoFechaProbableTermino, gsFormatoFecha) & "','" & pcProgCursoPersCodADictar & "'," & _
            pnProgCursoProgramado & "," & pnProgCursoNroHoras & "," & pnProgCursoModulo & "," & pnProgCursoNroParticipantes & _
            "," & pnProgCursoTipoMoneda & "," & pnProgCursoCostoEstimado & "," & pnProgCursoNacionalInternacional & "," & pnProgCursoEstado & ")"
    
    oConecta.ConexionActiva.Execute sSql
    
    Call oCapacitacion.InsertaRHCapacitacionAprobaciones(pnProgNro)
    
    'oConecta.CommitTrans
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHCapacitacionPrograma:
    Err.Raise Err.Number, "InsertaRHCapacitacionPrograma", Err.Description
    'oConecta.RollBackTrans
    Set oConecta = Nothing
End Sub

Public Sub ModificaRHCapacitacionPrograma(ByVal pnProgNro As Integer, _
                                ByVal pnCursoCod As Integer, _
                                ByVal pnProgCursoAnio As Integer, _
                                ByVal pcProgCursoObjetivo As String, _
                                ByVal pdProgCursoFechaProbableInicio As Date, _
                                ByVal pdProgCursoFechaProbableTermino As Date, _
                                ByVal pcProgCursoPersCodADictar As String, _
                                ByVal pnProgCursoProgramado As Integer, _
                                ByVal pnProgCursoNroHoras As Integer, _
                                ByVal pnProgCursoModulo As Integer, _
                                ByVal pnProgCursoNroParticipantes As Integer, _
                                ByVal pnProgCursoTipoMoneda As Integer, _
                                ByVal pnProgCursoCostoEstimado As Double, _
                                ByVal pnProgCursoNacionalInternacional As Integer)

Dim sSql As String
Dim oConecta As DConecta

On Error GoTo ErrorModificaRHCapacitacionPrograma
sSql = "UPDATE RHCapacitacionPrograma"
sSql = sSql & " SET nCursoCod =" & pnCursoCod & _
                        ",nProgCursoAnio =" & pnProgCursoAnio & _
                        ",cProgCursoObjetivo='" & pcProgCursoObjetivo & _
                        "',dProgCursoFechaProbableInicio='" & Format(pdProgCursoFechaProbableInicio, gsFormatoFecha) & _
                        "',dProgCursoFechaProbableTermino='" & Format(pdProgCursoFechaProbableTermino, gsFormatoFecha) & _
                        "',cProgCursoPersCodADictar='" & pcProgCursoPersCodADictar & _
                        "',nProgCursoProgramado=" & pnProgCursoProgramado & _
                        ",nProgCursoNroHoras=" & pnProgCursoNroHoras & _
                        ",nProgCursoModulo=" & pnProgCursoModulo & _
                        ",nProgCursoNroParticipantes=" & pnProgCursoNroParticipantes & _
                        ",nProgCursoTipoMoneda=" & pnProgCursoTipoMoneda & _
                        ",nProgCursoCostoEstimado=" & pnProgCursoCostoEstimado & _
                        ",nProgCursoNacionalInternacional=" & pnProgCursoNacionalInternacional & _
                        " WHERE nProgNro=" & pnProgNro

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorModificaRHCapacitacionPrograma:
    Err.Raise Err.Number, "ModificaRHCapacitacionPrograma", Err.Description
End Sub

Public Sub ModificaRHCapacitacionPrograma_Estado(ByVal pnProgNro As Integer, _
                                                ByVal pnProgCursoEstado As Integer)

Dim sSql As String
Dim oConecta As DConecta

On Error GoTo ErrorModificaRHCapacitacionPrograma
sSql = "UPDATE RHCapacitacionPrograma"
sSql = sSql & " SET nProgCursoEstado =" & pnProgCursoEstado & _
            " WHERE nProgNro=" & pnProgNro

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorModificaRHCapacitacionPrograma:
    Err.Raise Err.Number, "ModificaRHCapacitacionPrograma", Err.Description
End Sub


Public Function EliminaRHCapacitacionPrograma(ByVal pnProgNro As Integer)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHCapacitacionPrograma
    sSql = "DELETE FROM RHCapacitacionPrograma "
    sSql = sSql & " WHERE  nProgNro=" & pnProgNro
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHCapacitacionPrograma:
    Err.Raise Err.Number, "EliminaRHCapacitacionPrograma", Err.Description
End Function

Public Function Obtener_Siguiente_CodPrograma() As Long

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrorObtener_Siguiente_CodPrograma
    sSql = "SELECT Siguiente=COUNT(ISNULL(nProgNro,0))+1 FROM RHCapacitacionPrograma "
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_Siguiente_CodPrograma = 1
        Else
            Obtener_Siguiente_CodPrograma = rs("Siguiente") '+ 1
        End If
    End If
    
    oConecta.CierraConexion
    
    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Siguiente_CodPrograma:
    Err.Raise Err.Number, "Obtener_Siguiente_CodPrograma", Err.Description
End Function

Public Sub InsertaRHCapacitacionRequerimiento(ByVal pnReqNro As Integer, _
                                ByVal pnProgNro As Integer, _
                                ByVal pnCursoCod As Integer, _
                                ByVal pdReqFecha As Date, _
                                ByVal pcReqPersCodSolicitante As String, _
                                ByVal pnEstado As Integer)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHCapacitacionRequerimiento
    
    sSql = "INSERT INTO RHCapacitacionRequerimiento(nReqNro,nProgNro,nCursoCod,dReqFecha,cReqPersCodSolicitante,nEstado) "
    sSql = sSql & " VALUES(" & pnReqNro & "," & pnProgNro & "," & pnCursoCod & ",'" & pdReqFecha & "','" & pcReqPersCodSolicitante & "'," & pnEstado & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHCapacitacionRequerimiento:
    Err.Raise Err.Number, "InsertaRHCapacitacionRequerimiento", Err.Description
End Sub

Public Sub ModificaRHCapacitacionRequerimiento(ByVal pnReqNro As Integer, _
                                ByVal pnProgNro As Integer, _
                                ByVal pnCursoCod As Integer, _
                                ByVal pdReqFecha As Date, _
                                ByVal pcReqPersCodSolicitante As String, _
                                ByVal pnEstado As Integer)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorModificaRHCapacitacionRequerimiento
    sSql = "UPDATE RHCapacitacionRequerimiento"
    sSql = sSql & " SET nCursoCod =" & pnCursoCod & _
                        ",dReqFecha='" & pdReqFecha & _
                        "',pcReqPersCodSolicitante='" & pcReqPersCodSolicitante & _
                        "',nEstado=" & pnEstado & _
                        " WHERE nReqNro=" & pnReqNro & " AND nProgNro" & pnProgNro
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificaRHCapacitacionRequerimiento:
    Err.Raise Err.Number, "ModificaRHCapacitacionRequerimiento", Err.Description
End Sub

Public Function EliminaRHCapacitacionRequerimiento(ByVal pnReqNro As Integer, _
                                                    ByVal pnProgNro As Integer)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHCapacitacionRequerimiento
    sSql = "DELETE FROM RHCapacitacionRequerimiento " '
    sSql = sSql & " WHERE nReqNro=" & pnReqNro & " AND nProgNro" & pnProgNro
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHCapacitacionRequerimiento:
    Err.Raise Err.Number, "EliminaRHCapacitacionRequerimiento", Err.Description
End Function

'Public Sub InsertaRHCapacitacionNivelAprobacion(ByVal pcRHCargoCodReq As String, _
'                                        ByVal pcRHCargoCodAprob As String, _
'                                        ByVal pcRHCargoCodSol As String, _
'                                        ByVal pnOrden As Integer, _
'                                        ByVal pnNumAprobac As Integer)
'
'Dim sSQL As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorInsertaRHCapacitacionNivelAprobacion
'
'    sSQL = "INSERT INTO RHCapacitacionNivelAprobacion(cRHCargoCodReq,cRHCargoCodAprob,cRHCargoCodSol,nOrden,nNumAprobac)"
'    sSQL = sSQL & " VALUES('" & pcRHCargoCodReq & "','" & pcRHCargoCodAprob & "','" & pcRHCargoCodSol & "'," & pnOrden & "," & pnNumAprobac & ")"
'
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSQL
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Sub
'
'ErrorInsertaRHCapacitacionNivelAprobacion:
'    Err.Raise Err.Number, "InsertaRHCapacitacionNivelAprobacion", Err.Description
'End Sub
'
'Public Sub ModificaRHCapacitacionNivelAprobacion(ByVal pnCodNivelAprobac As Long, _
'                                            ByVal pcRHCargoCodReq As String, _
'                                            ByVal pcRHCargoCodAprob As String, _
'                                            ByVal pcRHCargoCodSol As String, _
'                                            ByVal pnOrden As Integer, _
'                                            ByVal pnNumAprobac As Integer)
'
'Dim sSQL As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorModificaRHCapacitacionNivelAprobacion
'    sSQL = "UPDATE RHCapacitacionNivelAprobacion"
'    sSQL = sSQL & " SET cRHCargoCodReq='" & pcRHCargoCodReq & _
'                        "',cRHCargoCodAprob='" & pcRHCargoCodAprob & _
'                        "',cRHCargoCodSol='" & pcRHCargoCodSol & _
'                        "',nOrden=" & pnOrden & _
'                        ",nNumAprobac=" & pnNumAprobac & _
'                        " WHERE nCodNivelAprob=" & pnCodNivelAprobac
'
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSQL
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Sub
'
'ErrorModificaRHCapacitacionNivelAprobacion:
'    Err.Raise Err.Number, "ModificaRHInduccion", Err.Description
'End Sub
'
'Public Function EliminaRHCapacitacionNivelAprobacion(ByVal pnCodNivelAprobac As Long)
'
'Dim sSQL As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorEliminaRHCapacitacionNivelAprobacion
'    sSQL = "DELETE FROM RHCapacitacionNivelAprobacion" '
'    sSQL = sSQL & " WHERE nCodNivelAprob =" & pnCodNivelAprobac
'
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSQL
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorEliminaRHCapacitacionNivelAprobacion:
'    Err.Raise Err.Number, "EliminaRHCapacitacionNivelAprobacion", Err.Description
'End Function

Public Sub InsertaRHCapacitacionNivelAprobacion(ByVal pcRHCargoCodRequiere As String, _
                                    ByVal pcRHCargoCodAprueba As String, _
                                    ByVal pnOrden As Integer, _
                                    ByVal pbTodasAgencias As Integer)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHCapacitacionNivelAprobacion
    
    sSql = "INSERT INTO RHCapacitacionNivelAprobacion(cRHCargoCodRequiere,cRHCargoCodAprueba,nOrden,bTodasAgencias)"
    sSql = sSql & " VALUES('" & pcRHCargoCodRequiere & "','" & pcRHCargoCodAprueba & "'," & pnOrden & "," & pbTodasAgencias & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHCapacitacionNivelAprobacion:
    Err.Raise Err.Number, "InsertaRHCapacitacionNivelAprobacion", Err.Description
End Sub

Public Sub ModificaRHCapacitacionNivelAprobacion(ByVal pcRHCargoCodRequiere As String, _
                                            ByVal pcRHCargoCodAprueba As String, _
                                            ByVal pnOrden As Integer, _
                                            ByVal pbTodasAgencias As Integer)

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorModificaRHCapacitacionNivelAprobacion
    sSql = "UPDATE RHCapacitacionNivelAprobacion"
    sSql = sSql & " SET nOrden=" & pnOrden & _
                    ",bTodasAgencias=" & pbTodasAgencias & _
                    " WHERE  cRHCargoCodRequiere='" & pcRHCargoCodRequiere & "' AND cRHCargoCodAprueba='" & pcRHCargoCodAprueba & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorModificaRHCapacitacionNivelAprobacion:
    Err.Raise Err.Number, "ModificaRHCapacitacionNivelAprobacion", Err.Description
End Sub

Public Function EliminaRHCapacitacionNivelAprobacion(ByVal pcRHCargoCodRequiere As String, _
                                                ByVal pcRHCargoCodAprueba As String)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHCapacitacionNivelAprobacion
    sSql = "DELETE FROM RHCapacitacionNivelAprobacion" '
    sSql = sSql & " WHERE  cRHCargoCodRequiere='" & pcRHCargoCodRequiere & "' AND cRHCargoCodAprueba='" & pcRHCargoCodAprueba & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHCapacitacionNivelAprobacion:
    Err.Raise Err.Number, "EliminaRHCapacitacionNivelAprobacion", Err.Description
End Function

Public Sub InsertaRHCapacitacionRequerimientoAprobaciones(ByVal pnReqNro As Integer, _
                                        ByVal pnProgNro As Integer, _
                                        ByVal pcPersCodSolicitante As String, _
                                        ByVal pcRHCargoCodAprob As String, _
                                        ByVal pnOrden As Long, _
                                        ByVal pcPersCodAprob As String, _
                                        ByVal pbAprobado As Integer, _
                                        ByVal pdFecha As Date, _
                                        ByVal pcComentario As String)
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHCapacitacionRequerimientoAprobaciones
    
    sSql = "INSERT INTO RHCapacitacionRequerimientoAprobaciones(nReqNro,nProgNro,cPersCodSolicitante,cRHCargoCodAprob,nOrden,cPersCodAprob,bAprobado,dFecha,cComentario)"
    sSql = sSql & " VALUES(" & pnReqNro & "," & pnProgNro & ",'" & pcPersCodSolicitante & "','" & pcRHCargoCodAprob & "'," & pnOrden & ",'" & pcPersCodAprob & "'," & pbAprobado & ",'" & pdFecha & "','" & pcComentario & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHCapacitacionRequerimientoAprobaciones:
    Err.Raise Err.Number, "InsertaRHCapacitacionRequerimientoAprobaciones", Err.Description
End Sub

Public Sub ModificaRHCapacitacionRequerimientoAprobaciones(ByVal pnReqNro As Integer, _
                                            ByVal pnProgNro As Integer, _
                                            ByVal pcPersCodSolicitante As String, _
                                            ByVal pcRHCargoCodAprob As String, _
                                            ByVal pnOrden As Long, _
                                            ByVal pcPersCodAprob As String, _
                                            ByVal pbAprobado As Integer, _
                                            ByVal pdFecha As Date, _
                                            ByVal pcComentario As String)
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorModificaRHCapacitacionRequerimientoAprobaciones
    sSql = "UPDATE RHCapacitacionRequerimientoAprobaciones"
    sSql = sSql & " SET cPersCodAprob=" & pcPersCodAprob & _
                        ",bAprobado=" & pbAprobado & _
                        ",dFecha='" & pdFecha & _
                        "',cComentario='" & pcComentario & _
                        "' WHERE nReqNro=" & pnReqNro & " AND nProgNro=" & pnProgNro & " AND cPersCodSolicitante='" & pcPersCodSolicitante & "' AND cRHCargoCodAprob='" & pcRHCargoCodAprob & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorModificaRHCapacitacionRequerimientoAprobaciones:
    Err.Raise Err.Number, "ModificaRHCapacitacionRequerimientoAprobaciones", Err.Description
End Sub

'Public Function EliminaRHCapacitacionRequerimientoAprobaciones(ByVal pcCodRequerimiento As String, _
'                                                        ByVal pcCodCargoActual As String, _
'                                                        ByVal pcPersCodReq As String)
'
'Dim sSQL As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorEliminaRHCapacitacionRequerimientoAprobaciones
'    sSQL = "DELETE FROM RHCapacitacionRequerimientoAprobaciones" '
'    sSQL = sSQL & " WHERE cCodRequerimiento='" & pcCodRequerimiento & "' AND cCodCargoActual='" & pcCodCargoActual & "' AND cPersCodReq='" & pcPersCodReq & "'"
'
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSQL
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorEliminaRHCapacitacionRequerimientoAprobaciones:
'    Err.Raise Err.Number, "EliminaRHCapacitacionRequerimientoAprobaciones", Err.Description
'End Function

Public Function Obtener_Niveles_Aprobacion() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Niveles_Aprobacion
    
    sSql = "SELECT DISTINCT cRHCargoCodReq,CargoReq=CatReq.cRHCargoDescripcion,cRHCargoCodAprob," & _
           " CargoAprob=CatAprob.cRHCargoDescripcion , nOrden FROM RHCapacitacionNivelAprobacion CNA INNER JOIN " & _
           " RHCargosTabla CatReq ON CNA.cRHCargoCodReq= CatReq.cRHCargoCod INNER JOIN RHCargosTabla CatAprob " & _
           " ON CNA.cRHCargoCodAprob= CatAprob.cRHCargoCod"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Niveles_Aprobacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Niveles_Aprobacion:
    Err.Raise Err.Number, "Obtener_Niveles_Aprobacion", Err.Description
End Function

Public Function Obtener_CargosPosibles_NivelAprobacion(ByVal pcRHCargoCodReq As String, _
                                                    ByVal pcRHCargoCodAprob As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_CargosPosibles_NivelAprobacion
    sSql = "SELECT CNA.nCodNivelAprob,CNA.cRHCargoCodSol,Cargo=Cat.cRHCargoDescripcion FROM RHCapacitacionNivelAprobacion CNA INNER JOIN RHCargosTabla Cat ON CNA.cRHCargoCodSol=Cat.cRHCargoCod " & _
            " WHERE CNA.cRHCargoCodReq='" & pcRHCargoCodReq & "' AND CNA.cRHCargoCodAprob='" & pcRHCargoCodAprob & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_CargosPosibles_NivelAprobacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_CargosPosibles_NivelAprobacion:
    Err.Raise Err.Number, "Obtener_CargosPosibles_NivelAprobacion", Err.Description
End Function

Public Function Obtener_Aprobadores_x_CargoSolicita(pcRHCargoCodRequiere As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Aprobadores_x_CargoSolicita
    
    sSql = "SELECT cRHCargoCodAprueba,cRHCargoDescripcion,nOrden,bTodasAgencias=CONVERT(int,bTodasAgencias) from " & _
           " RHCapacitacionNivelAprobacion ENA INNER JOIN RHCargosTabla Cat " & _
           " ON Cat.cRHCargoCod= ENA.cRHCargoCodAprueba " & _
           " WHERE cRHCargoCodRequiere ='" & pcRHCargoCodRequiere & "' ORDER BY nOrden"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Aprobadores_x_CargoSolicita = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtener_Aprobadores_x_CargoSolicita:
    Err.Raise Err.Number, "Obtener_Aprobadores_x_CargoSolicita", Err.Description
End Function

Public Function Obtener_Programa_Capacitacion(ByVal pnProgCursoAnio As Integer, _
                                            Optional ByVal pnProgCursoEstado As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Programa_Capacitacion
    
    sSql = "SELECT nProgNro,C.nCursoCod,C.cCursoDescripcion,FechaInicio=CONVERT(varchar(10),dProgCursoFechaProbableInicio,103),FechaTermino=CONVERT(varchar(10),dProgCursoFechaProbableTermino,103)," & _
           " Estado=CON.cConsDescripcion,nProgCursoProgramado, nProgCursoNroHoras,nProgCursoCostoEstimado,Moneda=M.cNomTab " & _
           " FROM RHCapacitacionPrograma P INNER JOIN RHCapacitacionCursos C ON P.nCursoCod=C.nCursoCod " & _
           " INNER JOIN DBComunes.dbo.TablaCod M ON M.cCodTab LIKE '03__' AND M.cValor=P.nProgCursoTipoMoneda" & _
           " INNER JOIN Constante CON ON CON.nConsValor=P.nProgCursoEstado AND CON.nConsCod=9076 " & _
           " WHERE nProgCursoAnio=" & pnProgCursoAnio & IIf(pnProgCursoEstado <> -1, " AND nProgCursoEstado=" & pnProgCursoEstado, "")
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Programa_Capacitacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Programa_Capacitacion:
    Err.Raise Err.Number, "Obtener_Programa_Capacitacion", Err.Description
End Function

Public Function Obtener_Codigo_ProgramaCapacitacion() As Integer
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorObtener_Codigo_ProgramaCapacitacion
    
    sSql = "SELECT Codigo=ISNULL(MAX(nProgNro),0)+1 FROM RHCapacitacionPrograma "
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        
        If Not rs.EOF Then
            Obtener_Codigo_ProgramaCapacitacion = rs("Codigo")
        Else
            Obtener_Codigo_ProgramaCapacitacion = 0
        End If
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Codigo_ProgramaCapacitacion:
    Err.Raise Err.Number, "Obtener_Codigo_ProgramaCapacitacion", Err.Description
End Function

Public Function Obtener_Monedas_Capacitacion() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Monedas_Capacitacion
    
    sSql = "SELECT cCodTab=RIGHT(cCodTab,1),cNomTab FROM DBComunes.dbo.TablaCod WHERE cCodTab LIKE '03__'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Monedas_Capacitacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Monedas_Capacitacion:
    Err.Raise Err.Number, "Obtener_Monedas_Capacitacion", Err.Description
End Function

Public Function Obtener_Datos_Programa_Capacitacion(ByVal pnProgNro As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_Programa_Capacitacion
    
    sSql = "SELECT nProgNro,nProgCursoAnio,C.nCursoCod,C.cCursoDescripcion,cProgCursoObjetivo,dProgCursoFechaProbableInicio,dProgCursoFechaProbableTermino," & _
           " cProgCursoPersCodADictar,PER.cPersNombre, nProgCursoModulo,nProgCursoNroParticipantes,  nProgCursoProgramado, nProgCursoNroHoras,nProgCursoCostoEstimado , nProgCursoTipoMoneda, nProgCursoNacionalInternacional, Estado = CON.cConsDescripcion " & _
           " FROM RHCapacitacionPrograma P INNER JOIN RHCapacitacionCursos C ON P.nCursoCod=C.nCursoCod INNER JOIN Persona PER ON PER.cPersCod= P.cProgCursoPersCodADictar " & _
           " INNER JOIN Constante CON ON CON.nConsValor=P.nProgCursoEstado AND CON.nConsCod=9076 " & _
           " WHERE nProgNro=" & pnProgNro

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_Programa_Capacitacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Datos_Programa_Capacitacion:
    Err.Raise Err.Number, "Obtener_Datos_Programa_Capacitacion", Err.Description
End Function


Public Sub InsertaRHCapacitacionEjecucion(ByVal pnProgNro As Integer, _
                                        ByVal pnItem As Integer, _
                                        ByVal pdEjeCursoFechaInicio As Date, _
                                        ByVal pcEjeCursoUbigeo As String, _
                                        ByVal pdEjeCursoFechaTermino As Date, _
                                        ByVal pnEjeCursoModulo As Integer, _
                                        ByVal pnEjeCursoNroHoras As Integer, _
                                        ByVal pnEjeCursoNroParticipantes As Integer, _
                                        ByVal pnEjeCursoCostoReal As Double, _
                                        ByVal pnEjeCursoTipoMoneda As Integer, _
                                        ByVal pcEjeCursoPersCodInstitucion As String, _
                                        ByVal pcEjeCursoPersCodDictar As String, _
                                        ByVal pcEjeCursoBcoCod As String, _
                                        ByVal pcEjeCursoBcoNroCuenta As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHCapacitacionEjecucion
    
    sSql = "INSERT INTO RHCapacitacionEjecucion(nProgNro, nItem, dEjeCursoFechaInicio, cEjeCursoUbigeo, dEjeCursoFechaTermino, nEjeCursoModulo, nEjeCursoNroHoras, nEjeCursoNroParticipantes, nEjeCursoCostoReal, nEjeCursoTipoMoneda, cEjeCursoPersCodInstitucion, cEjeCursoPersCodDictar, cEjeCursoBcoCod, cEjeCursoBcoNroCuenta) " & _
            " VALUES(" & pnProgNro & "," & pnItem & ",'" & Format(pdEjeCursoFechaInicio, gsFormatoFecha) & "','" & pcEjeCursoUbigeo & "','" & _
            Format(pdEjeCursoFechaTermino, gsFormatoFecha) & "'," & pnEjeCursoModulo & "," & pnEjeCursoNroHoras & "," & _
            pnEjeCursoNroParticipantes & "," & pnEjeCursoCostoReal & "," & pnEjeCursoTipoMoneda & ",'" & pcEjeCursoPersCodInstitucion & "','" & pcEjeCursoPersCodDictar & _
            "','" & pcEjeCursoBcoCod & "','" & pcEjeCursoBcoNroCuenta & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHCapacitacionEjecucion:
    Err.Raise Err.Number, "InsertaRHCapacitacionEjecucion", Err.Description
End Sub

Public Sub ModificaRHCapacitacionEjecucion(ByVal pnProgNro As Integer, _
                                        ByVal pnItem As Integer, _
                                        ByVal pdEjeCursoFechaInicio As Date, _
                                        ByVal pcEjeCursoUbigeo As String, _
                                        ByVal pdEjeCursoFechaTermino As Date, _
                                        ByVal pnEjeCursoModulo As Integer, _
                                        ByVal pnEjeCursoNroHoras As Integer, _
                                        ByVal pnEjeCursoNroParticipantes As Integer, _
                                        ByVal pnEjeCursoCostoReal As Double, _
                                        ByVal pnEjeCursoTipoMoneda As Integer, _
                                        ByVal pcEjeCursoPersCodInstitucion As String, _
                                        ByVal pcEjeCursoPersCodDictar As String, _
                                        ByVal pcEjeCursoBcoCod As String, _
                                        ByVal pcEjeCursoBcoNroCuenta As String)

Dim sSql As String
Dim oConecta As DConecta

On Error GoTo ErrorModificaRHCapacitacionEjecucion
sSql = "UPDATE RHCapacitacionEjecucion"
sSql = sSql & " SET dEjeCursoFechaInicio ='" & Format(pdEjeCursoFechaInicio, gsFormatoFecha) & _
                        "',cEjeCursoUbigeo =" & pcEjeCursoUbigeo & _
                        "',dEjeCursoFechaTermino='" & Format(pdEjeCursoFechaTermino, gsFormatoFecha) & _
                        "',nEjeCursoModulo=" & pnEjeCursoModulo & _
                        ",nEjeCursoNroHoras=" & pnEjeCursoNroHoras & _
                        ",nEjeCursoNroParticipantes=" & pnEjeCursoNroParticipantes & _
                        ",nEjeCursoCostoReal=" & pnEjeCursoCostoReal & _
                        ",nEjeCursoTipoMoneda=" & pnEjeCursoTipoMoneda & _
                        ",cEjeCursoPersCodInstitucion='" & pcEjeCursoPersCodInstitucion & _
                        "',cEjeCursoPersCodDictar='" & pcEjeCursoPersCodDictar & _
                        "',cEjeCursoBcoCod='" & pcEjeCursoBcoCod & _
                        "',cEjeCursoBcoNroCuenta='" & pcEjeCursoBcoNroCuenta & _
                        "' WHERE  nProgNro=" & pnProgNro & " AND nItem=" & pnItem

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorModificaRHCapacitacionEjecucion:
    Err.Raise Err.Number, "ModificaRHCapacitacionEjecucion", Err.Description
End Sub

Public Function EliminaRHCapacitacionEjecucion(ByVal pnProgNro As Integer, _
                                                ByVal pnItem As Integer)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHCapacitacionEjecucion
    sSql = "DELETE FROM RHCapacitacionEjecucion "
    sSql = sSql & " WHERE  nProgNro=" & pnProgNro & " AND nItem=" & pnItem
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHCapacitacionEjecucion:
    Err.Raise Err.Number, "EliminaRHCapacitacionEjecucion", Err.Description
End Function

Public Sub InsertaRHCapacitacionPlanPagos(ByVal pnProgNro As Integer, _
                                        ByVal pnItem As Integer, _
                                        ByVal pnPlanPagosNroCuota As Integer, _
                                        ByVal pdPlanPagosFechaCuota As Date, _
                                        ByVal pnPlanPagosMontoCuota As Double, _
                                        ByVal pnPlanPagosEstadoCuota As Integer, _
                                        ByVal pdPlanPagosFechaCancelacion As Date)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHCapacitacionPlanPagos
    
    If pdPlanPagosFechaCancelacion = 0 Then
        sSql = "INSERT INTO RHCapacitacionPlanPagos(nProgNro, nItem, nPlanPagosNroCuota  ,dPlanPagosFechaCuota, nPlanPagosMontoCuota, nPlanPagosEstadoCuota) " & _
            " VALUES(" & pnProgNro & "," & pnItem & "," & pnPlanPagosNroCuota & ",'" & Format(pdPlanPagosFechaCuota, gsFormatoFecha) & "'," & pnPlanPagosMontoCuota & "," & _
                pnPlanPagosEstadoCuota & ")"
    Else
        sSql = "INSERT INTO RHCapacitacionPlanPagos(nProgNro, nItem, nPlanPagosNroCuota  ,dPlanPagosFechaCuota, nPlanPagosMontoCuota, nPlanPagosEstadoCuota, dPlanPagosFechaCancelacion) " & _
            " VALUES(" & pnProgNro & "," & pnItem & "," & pnPlanPagosNroCuota & ",'" & Format(pdPlanPagosFechaCuota, gsFormatoFecha) & "'," & pnPlanPagosMontoCuota & "," & _
                pnPlanPagosEstadoCuota & ",'" & Format(pdPlanPagosFechaCancelacion, gsFormatoFecha) & "')"
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHCapacitacionPlanPagos:
    Err.Raise Err.Number, "InsertaRHCapacitacionPlanPagos", Err.Description
End Sub

Public Sub ModificaRHCapacitacionPlanPagos(ByVal pnProgNro As Integer, _
                                        ByVal pnItem As Integer, _
                                        ByVal pnPlanPagosNroCuota As Integer, _
                                        ByVal pdPlanPagosFechaCuota As Date, _
                                        ByVal pnPlanPagosMontoCuota As Double, _
                                        ByVal pdPlanPagosFechaCancelacion As Date)

Dim sSql As String
Dim oConecta As DConecta

On Error GoTo ErrorModificaRHCapacitacionPlanPagos
sSql = "UPDATE RHCapacitacionPlanPagos"
sSql = sSql & " SET dPlanPagosFechaCuota ='" & Format(pdPlanPagosFechaCuota, gsFormatoFecha) & _
                        "',nPlanPagosMontoCuota =" & pnPlanPagosMontoCuota & _
                        ",dPlanPagosFechaCancelacion='" & Format(pdPlanPagosFechaCancelacion, gsFormatoFecha) & _
                        "' WHERE  nProgNro=" & pnProgNro & " AND nItem=" & pnItem & " AND nPlanPagosNroCuota=" & pnPlanPagosNroCuota

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorModificaRHCapacitacionPlanPagos:
    Err.Raise Err.Number, "ModificaRHCapacitacionPlanPagos", Err.Description
End Sub

Public Sub ModificaRHCapacitacionPlanPagos_Estado(ByVal pnProgNro As Integer, _
                                                ByVal pnItem As Integer, _
                                                ByVal pnPlanPagosNroCuota As Integer, _
                                                ByVal pnPlanPagosEstadoCuota As Integer)

Dim sSql As String
Dim oConecta As DConecta

On Error GoTo ErrorModificaRHCapacitacionPlanPagos
sSql = "UPDATE RHCapacitacionPlanPagos"
sSql = sSql & " SET nPlanPagosEstadoCuota =" & pnPlanPagosEstadoCuota & _
            " WHERE  nProgNro=" & pnProgNro & " AND nItem=" & pnItem & " AND nPlanPagosNroCuota =" & pnPlanPagosNroCuota

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorModificaRHCapacitacionPlanPagos:
    Err.Raise Err.Number, "ModificaRHCapacitacionPlanPagos", Err.Description
End Sub

Public Function EliminaRHCapacitacionPlanPagos(ByVal pnProgNro As Integer, _
                                                ByVal pnItem As Integer, _
                                                ByVal pnPlanPagosNroCuota As Integer)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHCapacitacionPlanPagos
    sSql = "DELETE FROM RHCapacitacionPlanPagos "
    sSql = sSql & " WHERE  nProgNro=" & pnProgNro & " AND nItem=" & pnItem & " AND nPlanPagosNroCuota=" & pnPlanPagosNroCuota
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHCapacitacionPlanPagos:
    Err.Raise Err.Number, "EliminaRHCapacitacionPlanPagos", Err.Description
End Function

Public Sub InsertaRHCapacitacionParticipantes(ByVal pnProgNro As Integer, _
                                        ByVal pnItem As Integer, _
                                        ByVal pcPersCod As String, _
                                        ByVal pnGeneraMemoComunicacion As Integer, _
                                        ByVal pnConfirmaAsistencia As Integer, _
                                        ByVal pcMotivoNoAsistencia As String, _
                                        ByVal pnVerificacionAsistencia As Integer)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHCapacitacionParticipantes
    
    sSql = "INSERT INTO RHCapacitacionParticipantes(nProgNro, nItem,cPersCod ,nGeneraMemoComunicacion, nConfirmaAsistencia, cMotivoNoAsistencia, nVerificacionAsistencia) " & _
            " VALUES(" & pnProgNro & "," & pnItem & ",'" & pcPersCod & "'," & pnGeneraMemoComunicacion & "," & pnConfirmaAsistencia & ",'" & _
                pcMotivoNoAsistencia & "'," & pnVerificacionAsistencia & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHCapacitacionParticipantes:
    Err.Raise Err.Number, "InsertaRHCapacitacionParticipantes", Err.Description
End Sub

Public Sub InsertaRHCapacitacionParticipantes_Reemplazante(ByVal pnProgNro As Integer, _
                                        ByVal pnItem As Integer, _
                                        ByVal pcPersCod As String, _
                                        ByVal pcPersCodReemplaza As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHCapacitacionParticipantes_Reemplazante
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    
    sSql = "INSERT INTO RHCapacitacionParticipantes(nProgNro, nItem,cPersCod ,cPersCodReemplaza) " & _
            " VALUES(" & pnProgNro & "," & pnItem & ",'" & pcPersCod & "','" & pcPersCodReemplaza & "')"
    oConecta.ConexionActiva.Execute sSql
    
    Call EliminaRHCapacitacionParticipantes(pnProgNro, pnItem, pcPersCodReemplaza)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHCapacitacionParticipantes_Reemplazante:
    Err.Raise Err.Number, "InsertaRHCapacitacionParticipantes_Reemplazante", Err.Description
End Sub

Public Sub ModificaRHCapacitacionParticipantes(ByVal pnProgNro As Integer, _
                                                ByVal pnItem As Integer, _
                                                ByVal pcPersCod As String, _
                                                ByVal pnGeneraMemoComunicacion As Integer, _
                                                ByVal pnConfirmaAsistencia As Integer, _
                                                ByVal pcMotivoNoAsistencia As String, _
                                                ByVal pnVerificacionAsistencia As Integer)
Dim sSql As String
Dim oConecta As DConecta

On Error GoTo ErrorModificaRHCapacitacionParticipantes
sSql = "UPDATE RHCapacitacionParticipantes"
sSql = sSql & " SET nGeneraMemoComunicacion=" & pnGeneraMemoComunicacion & _
                        ",nConfirmaAsistencia =" & pnConfirmaAsistencia & _
                        ",cMotivoNoAsistencia='" & pcMotivoNoAsistencia & _
                        "',nVerificacionAsistencia=" & pnVerificacionAsistencia & _
                        " WHERE  nProgNro=" & pnProgNro & " AND nItem=" & pnItem & " AND cPersCod='" & pcPersCod & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorModificaRHCapacitacionParticipantes:
    Err.Raise Err.Number, "ModificaRHCapacitacionParticipantes", Err.Description
End Sub

Public Function EliminaRHCapacitacionParticipantes(ByVal pnProgNro As Integer, _
                                                ByVal pnItem As Integer, _
                                                ByVal pcPersCod As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHCapacitacionParticipantes
    sSql = "DELETE FROM RHCapacitacionParticipantes "
    sSql = sSql & " WHERE  nProgNro=" & pnProgNro & " AND nItem=" & pnItem & " AND cPersCod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHCapacitacionParticipantes:
    Err.Raise Err.Number, "EliminaRHCapacitacionParticipantes", Err.Description
End Function

Public Sub InsertaRHCapacitacionAprobacion(ByVal pnProgNro As Integer, _
                                            ByVal pcRHCargoCodAprob As String, _
                                            ByVal pnOrden As Integer)
Dim sSql As String
Dim oConecta As DConecta

    
    On Error GoTo ErrorInsertaRHCapacitacionAprobacion
    

    sSql = "INSERT INTO RHCapacitacionAprobacion(nProgNro, cRHCargoCodAprob, nOrden,bAprobado) " & _
            " VALUES(" & pnProgNro & ",'" & pcRHCargoCodAprob & "'," & pnOrden & ",0)"  'Por Defecto No Aprobado
    
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHCapacitacionAprobacion:
    Err.Raise Err.Number, "InsertaRHCapacitacionAprobacion", Err.Description
End Sub

Public Sub ModificaRHCapacitacionAprobacion(ByVal pnProgNro As Integer, _
                                            ByVal pcRHCargoCodAprob As String, _
                                            ByVal pcPersCodAprob As String, _
                                            ByVal pdFecha As Date, _
                                            ByVal pbAprobado As Integer)
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
Dim nEstado As Integer
Dim nNumAprocaciones As Integer

On Error GoTo ErrorModificaRHCapacitacionAprobacion

Set oConecta = New DConecta

If oConecta.AbreConexion = True Then

    sSql = "UPDATE RHCapacitacionAprobacion"
    sSql = sSql & " SET cPersCodAprob=" & pcPersCodAprob & _
                        ",dFecha ='" & Format(pdFecha, gsFormatoFecha) & _
                        "',bAprobado=" & pbAprobado & _
                        " WHERE  nProgNro=" & pnProgNro & " AND cRHCargoCodAprob='" & pcRHCargoCodAprob & "'"

    oConecta.ConexionActiva.Execute sSql

    sSql = "SELECT nProgCursoEstado FROM RHCapacitacionPrograma WHERE nProgNro=" & pnProgNro
    Set rs = oConecta.CargaRecordSet(sSql)
    nEstado = rs("nProgCursoEstado")
    Select Case nEstado
    Case 1
        sSql = "SELECT NumReg=COUNT(*) FROM RHCapacitacionAprobacion WHERE nProgNro=" & pnProgNro & " AND nOrden=1 AND bAprobado=1"
        Set rs = oConecta.CargaRecordSet(sSql)
        nNumAprocaciones = rs("NumReg")
        If nNumAprocaciones > 0 Then
            sSql = "UPDATE RHCapacitacionPrograma SET nProgCursoEstado=2 WHERE nProgNro=" & pnProgNro
            oConecta.ConexionActiva.Execute sSql
        End If
    Case 2
        sSql = "SELECT NumReg=COUNT(*) FROM RHCapacitacionAprobacion WHERE nProgNro=" & pnProgNro & " AND nOrden=2 AND bAprobado=1"
        Set rs = oConecta.CargaRecordSet(sSql)
        nNumAprocaciones = rs("NumReg")
        If nNumAprocaciones >= 3 Then   '3 son Miembros del Comite
            sSql = "UPDATE RHCapacitacionPrograma SET nProgCursoEstado=3 WHERE nProgNro=" & pnProgNro
            oConecta.ConexionActiva.Execute sSql
        End If
    End Select
    oConecta.CierraConexion
End If
Set oConecta = Nothing
Set rs = Nothing
Exit Sub

ErrorModificaRHCapacitacionAprobacion:
    Err.Raise Err.Number, "ModificaRHCapacitacionAprobacion", Err.Description
End Sub

Public Function Obtener_Programa_a_Aprobar_x_Cargo(ByVal pnProgCursoAnio As Integer, ByVal pcRHCargoCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Programa_a_Aprobar_x_Cargo
    
    sSql = " SELECT P.nProgNro,'0',C.cCursoDescripcion,FechaInicio=CONVERT(varchar(10),dProgCursoFechaProbableInicio,103),FechaTermino=CONVERT(varchar(10),dProgCursoFechaProbableTermino,103)," & _
           " Estado=CON.cConsDescripcion,CursoProgramado=CASE WHEN nProgCursoProgramado=1 THEN 'SI' ELSE 'NO' END, nProgCursoNroHoras,nProgCursoCostoEstimado,Moneda=M.cNomTab FROM RHCapacitacionPrograma P INNER JOIN RHCapacitacionCursos C " & _
           " ON P.nCursoCod=C.nCursoCod INNER JOIN DBComunes.dbo.TablaCod M ON M.cCodTab LIKE '03__' AND M.cValor=P.nProgCursoTipoMoneda INNER JOIN Constante CON ON CON.nConsValor=P.nProgCursoEstado " & _
           " AND CON.nConsCod=9076 INNER JOIN RHCapacitacionAprobacion AP ON P.nProgNro=AP.nProgNro AND AP.cRHCargoCodAprob='" & pcRHCargoCod & "' AND AP.bAprobado=0 " & _
           " AND ((P.nProgCursoEstado=1 AND AP.nOrden=1 )OR (P.nProgCursoEstado=2 AND AP.nOrden=2))" & _
           " WHERE nProgCursoAnio=" & pnProgCursoAnio

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Programa_a_Aprobar_x_Cargo = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Programa_a_Aprobar_x_Cargo:
    Err.Raise Err.Number, "Obtener_Programa_a_Aprobar_x_Cargo", Err.Description
End Function

Public Function Obtener_Ejecuciones_x_Curso(ByVal pnProgNro As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Ejecuciones_x_Curso
    
    sSql = " SELECT nProgNro, nItem,FechaInicio=CONVERT(varchar(10),dEjeCursoFechaInicio,103)," & _
           " FechaTermino=CONVERT(varchar(10),dEjeCursoFechaTermino,103),nEjeCursoNroHoras,Inst.cPersNombre, " & _
           " nEjeCursoCostoReal , Moneda = m.cNomTab " & _
           " FROM RHCapacitacionEjecucion E INNER JOIN DBComunes.dbo.TablaCod M ON M.cCodTab LIKE '03__' " & _
           " AND M.cValor=E.nEjeCursoTipoMoneda INNER JOIN Persona Inst ON Inst.cPersCod = E.cEjeCursoPersCodInstitucion" & _
           " WHERE nProgNro=" & pnProgNro
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Ejecuciones_x_Curso = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Ejecuciones_x_Curso:
    Err.Raise Err.Number, "Obtener_Ejecuciones_x_Curso", Err.Description
End Function

Public Function Obtener_Datos_Ejecucion_Programa(ByVal pnProgNro As Integer, ByVal pnItem As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_Ejecucion_Programa
    
    sSql = "SELECT nProgNro, nItem, cEjeCursoUbigeo, FechaInicio=CONVERT(varchar(10),dEjeCursoFechaInicio,103), " & _
           " FechaTermino=CONVERT(varchar(10),dEjeCursoFechaTermino,103), " & _
           " Localidad=(select rtrim(isnull((select cdeszon from dbcomunes..zonas where ccodzon like '1'+substring(z.ccodzon,2,2)+'%'),''))+ ' - ' + " & _
           " rtrim(isnull((select cdeszon from dbcomunes..zonas where ccodzon like '2'+substring(z.ccodzon,2,4)+'%'),''))+ ' - ' + " & _
           " rtrim(isnull((select cdeszon from dbcomunes..zonas where ccodzon like '3'+substring(z.ccodzon,2,6)+'%'),''))+ ' - ' + " & _
           " RTrim (cdeszon)From DBComunes..zonas Z Where ccodzon=cEjeCursoUbigeo),cEjeCursoPersCodDictar, Persona=P.cPersNombre, " & _
           " cEjeCursoPersCodInstitucion,Institucion=Inst.cPersNombre, nEjeCursoNroHoras,nEjeCursoModulo,nEjeCursoNroParticipantes, " & _
           " nEjeCursoCostoReal,nEjeCursoTipoMoneda,Moneda=M.cNomTab, cEjeCursoBcoCod , Banco = b.cPersNombre, cEjeCursoBcoNroCuenta " & _
           " FROM RHCapacitacionEjecucion E INNER JOIN DBComunes.dbo.TablaCod M ON M.cCodTab LIKE '03__' AND M.cValor=E.nEjeCursoTipoMoneda " & _
           " INNER JOIN Persona Inst ON Inst.cPersCod = E.cEjeCursoPersCodInstitucion INNER JOIN Persona P ON P.cPersCod = E.cEjeCursoPersCodDictar " & _
           " INNER JOIN Persona B ON B.cPersCod= E.cEjeCursoBcoCod" & _
           " WHERE nProgNro=" & pnProgNro & " AND nItem=" & pnItem

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_Ejecucion_Programa = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Datos_Ejecucion_Programa:
    Err.Raise Err.Number, "Obtener_Datos_Ejecucion_Programa", Err.Description
End Function

Public Function Obtener_Cuotas_x_Curso(ByVal pnProgNro As Integer, ByVal pnItem As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Cuotas_x_Curso
    
    sSql = " SELECT nPlanPagosNroCuota,dPlanPagosFechaCuota=CONVERT(varchar(10),dPlanPagosFechaCuota,103), nPlanPagosMontoCuota, ESTADO=C.cConsDescripcion,dPlanPagosFechaCancelacion=CONVERT(varchar(10),dPlanPagosFechaCancelacion,103) " & _
           " FROM RHCapacitacionPlanPagos PP INNER JOIN Constante C ON PP.nPlanPagosEstadoCuota = c.nConsValor And c.nConsCod = 9077 " & _
           " WHERE nProgNro=" & pnProgNro & " AND nItem=" & pnItem
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Cuotas_x_Curso = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Cuotas_x_Curso:
    Err.Raise Err.Number, "Obtener_Cuotas_x_Curso", Err.Description
End Function

Public Function Obtener_Item_x_Programa(pnProgNro As Integer) As Integer
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorObtener_Item_x_Programa
    
    sSql = "SELECT Codigo=ISNULL(MAX(nItem),0)+1 FROM RHCapacitacionEjecucion WHERE nProgNro= " & pnProgNro
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        
        If Not rs.EOF Then
            Obtener_Item_x_Programa = rs("Codigo")
        Else
            Obtener_Item_x_Programa = 0
        End If
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Item_x_Programa:
    Err.Raise Err.Number, "Obtener_Item_x_Programa", Err.Description
End Function

Public Sub InsertaRHCapacitacionEmpleado(ByVal pnProgNro As Integer, _
                                        ByVal pcPersCod As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHCapacitacionEmpleado
    
    sSql = "INSERT INTO RHCapacitacionEmpleado(nProgNro,cPersCod) " & _
            " VALUES(" & pnProgNro & ",'" & pcPersCod & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHCapacitacionEmpleado:
    Err.Raise Err.Number, "InsertaRHCapacitacionEmpleado", Err.Description
End Sub

Public Function EliminaRHCapacitacionEmpleado(ByVal pnProgNro As Integer, _
                                        ByVal pcPersCod As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHCapacitacionEmpleado
    sSql = "DELETE FROM RHCapacitacionEmpleado "
    sSql = sSql & " WHERE  nProgNro=" & pnProgNro & " AND cPerscod='" & pcPersCod & "'"
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHCapacitacionEmpleado:
    Err.Raise Err.Number, "EliminaRHCapacitacionEmpleado", Err.Description
End Function

Public Function Obtener_Participantes_Programa_Capacitacion(ByVal pnProgNro As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Participantes_Programa_Capacitacion
    
    sSql = "SELECT nProgNro, P.cPersCod, P.cPersNombre FROM RHCapacitacionEmpleado Pro INNER JOIN Persona P ON Pro.cPersCod=P.cPersCod" & _
           " WHERE nProgNro=" & pnProgNro

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Participantes_Programa_Capacitacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Participantes_Programa_Capacitacion:
    Err.Raise Err.Number, "Obtener_Participantes_Programa_Capacitacion", Err.Description
End Function

Public Function Obtener_Asistencia_Participantes_Programa(ByVal pnProgNro As Integer, Optional ByVal pnItem As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Asistencia_Participantes_Programa
    
    If pnItem <> -1 Then
        sSql = "SELECT P.cPersCod,P.cPersNombre,nConfirmaAsistencia=CASE WHEN nConfirmaAsistencia=1 THEN 'SI' + '" & Space(75) & "' + '1' ELSE 'NO' + '" & Space(75) & "' + '0' END ,cMotivoNoAsistencia=ISNULL(cMotivoNoAsistencia,''),nVerificacionAsistencia=CASE WHEN nVerificacionAsistencia=1 THEN 'SI' + '" & Space(75) & "' + '1' ELSE 'NO' + '" & Space(75) & "' + '0' END  " & _
           " FROM RHCapacitacionParticipantes CAP INNER JOIN Persona P ON CAP.cPersCod = P.cPersCod " & _
           " WHERE nProgNro=" & pnProgNro & " AND nItem=" & pnItem
    Else
        sSql = "SELECT P.cPersCod,P.cPersNombre,'NO' + '" & Space(75) & "' + '0','','NO' + '" & Space(75) & "' + '0' " & _
           " FROM RHCapacitacionEmpleado CEM INNER JOIN Persona P ON CEM.cPersCod = P.cPersCod " & _
           " WHERE nProgNro=" & pnProgNro
    End If
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Asistencia_Participantes_Programa = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Asistencia_Participantes_Programa:
    Err.Raise Err.Number, "Obtener_Asistencia_Participantes_Programa", Err.Description
End Function

Public Function Obtener_Documentos() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Documentos
    
    sSql = "SELECT nDocTpo,cDocDesc,1 FROM Documento"

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Documentos = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Documentos:
    Err.Raise Err.Number, "Obtener_Documentos", Err.Description
End Function

Public Sub InsertaRHCapacitacionCursos_Gastos(ByVal pnCursoCod As Integer, _
                                        ByVal pcPersCod As String, _
                                        ByVal pnItem As Integer, _
                                        ByVal pnTipoGasto As Integer, _
                                        ByVal pnDocTpo As Integer, _
                                        ByVal pcDocNro As String, _
                                        ByVal pnMonto As Double, _
                                        ByVal pcGlosa As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorInsertaRHCapacitacionCursos_Gastos
    
    sSql = "INSERT INTO RHCapacitacionCursos_Gastos(nCursoCod, cPersCod, nItem, nTipoGasto, nDocTpo,cDocNro ,nMonto, cGlosa) " & _
            " VALUES(" & pnCursoCod & ",'" & pcPersCod & "'," & pnItem & "," & pnTipoGasto & "," & pnDocTpo & ",'" & pcDocNro & "'," & pnMonto & ",'" & pcGlosa & "')"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertaRHCapacitacionCursos_Gastos:
    Err.Raise Err.Number, "InsertaRHCapacitacionCursos_Gastos", Err.Description
End Sub

Public Function EliminaRHCapacitacionCursos_Gastos(ByVal pnCursoCod As Integer, _
                                        ByVal pcPersCod As String, _
                                        ByVal pnItem As Integer)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaRHCapacitacionCursos_Gastos
    sSql = "DELETE FROM RHCapacitacionCursos_Gastos "
    sSql = sSql & " WHERE  nCursoCod=" & pnCursoCod & " AND cPerscod='" & pcPersCod & "' AND nItem=" & pnItem
                        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorEliminaRHCapacitacionCursos_Gastos:
    Err.Raise Err.Number, "EliminaRHCapacitacionCursos_Gastos", Err.Description
End Function

Public Function Obtener_CursosGastos(ByVal pnCursoCod As Integer, ByVal pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_CursosGastos
    
    sSql = "SELECT nItem,CG.nTipoGasto,C.cConsDescripcion ,CG.nDocTpo,TipoDoc=cDocDesc, cDocNro,nMonto, cGlosa" & _
           " FROM RHCapacitacionCursos_Gastos CG INNER JOIN Documento D ON CG.nDocTpo= D.nDocTpo" & _
           " INNER JOIN Constante C ON C.nConsValor=CG.nTipoGasto AND C.nConsCod=9093" & _
           " WHERE nCursoCod=" & pnCursoCod & " AND cPersCod='" & pcPersCod & "' ORDER BY nItem "

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_CursosGastos = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_CursosGastos:
    Err.Raise Err.Number, "Obtener_CursosGastos", Err.Description
End Function

Public Function Obtener_Item(ByVal pnCursoCod As Integer, ByVal pcPersCod As String) As Integer
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorObtener_Item
    
    Obtener_Item = 0
    sSql = "SELECT Item=ISNULL(MAX(nItem),0)" & _
           " FROM RHCapacitacionCursos_Gastos WHERE nCursoCod=" & pnCursoCod & " AND cPersCod='" & pcPersCod & "'"
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        Obtener_Item = CInt(rs("Item")) + 1
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Item:
    Err.Raise Err.Number, "Obtener_Item", Err.Description
End Function
