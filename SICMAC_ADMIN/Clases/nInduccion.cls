VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nInduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Obtener_Siguiente_CodInduccion() As Long

Dim sSQL As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrorObtener_Siguiente_CodInduccion
    sSQL = "SELECT Siguiente=COUNT(nCodInduccion)+1 FROM RHInduccion "
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSQL)
        If rs.EOF Then
            Obtener_Siguiente_CodInduccion = 1
        Else
            Obtener_Siguiente_CodInduccion = rs("Siguiente") '+ 1
        End If
    End If
    
    oConecta.CierraConexion
    
    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Siguiente_CodInduccion:
    Err.Raise Err.Number, "Obtener_Siguiente_CodInduccion", Err.Description

End Function

Public Function Obtener_Siguiente_CodSesion() As Long

Dim sSQL As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrorObtener_Siguiente_CodSesion
    sSQL = "SELECT Siguiente=COUNT(nCodSeccion)+1 FROM RHInduccionSesion "
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSQL)
        If rs.EOF Then
            Obtener_Siguiente_CodSesion = 1
        Else
            Obtener_Siguiente_CodSesion = rs("Siguiente") '+ 1
        End If
    End If
    
    oConecta.CierraConexion
    
    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Siguiente_CodSesion:
    Err.Raise Err.Number, "Obtener_Siguiente_CodSesion", Err.Description

End Function

Public Function Obtener_SeleccionProceso_x_Fechas(ByVal pdFecha As Date) As ADODB.Recordset
                    
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorObtener_SeleccionProceso_x_Fechas
    'Se modifico para considerar los Estados y no las Fechas
    sSQL = "SELECT cCodProceso,cDescProceso,1,nTipoContrato,dFechaIni,dFechaFin,nEstado FROM RHSeleccionProceso "
    'sSql = sSql & " WHERE '" & Format(pdFecha, gsFormatoFecha) & "' BETWEEN dFechaIni AND dFechaFin "
    sSQL = sSQL & " WHERE nEstado=3"
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_SeleccionProceso_x_Fechas = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_SeleccionProceso_x_Fechas:
    Err.Raise Err.Number, "Obtener_SeleccionProceso_x_Fechas", Err.Description
End Function

Public Function Registra_Proceso_Induccion(ByVal pnCodInduccion As Long, _
                            ByVal pcCodProceso As String, _
                            ByVal pnEstado As Long, _
                            ByVal prsInduccionFase As ADODB.Recordset, _
                            ByVal prsInduccionSesion As ADODB.Recordset)

'Dim sSql As String
'Dim oConecta As DConecta
    
    On Error GoTo ErrorRegistra_Proceso_Induccion
    
    Dim oInduccion As dInduccion
    Set oInduccion = New dInduccion
    
    Call oInduccion.InsertaRHInduccion(pnCodInduccion, _
                                        pcCodProceso, _
                                        pnEstado)
    
    While Not prsInduccionFase.EOF
        Call oInduccion.InsertaRHInduccionFase(pnCodInduccion, _
                                                prsInduccionFase("nCodTema"), _
                                                prsInduccionFase("cExpositor"), _
                                                1, prsInduccionFase("dInicio"), _
                                                prsInduccionFase("dFin"), _
                                                prsInduccionFase("dHoraInicio"), _
                                                prsInduccionFase("dHoraFin"), _
                                                prsInduccionFase("cComentario"), _
                                                prsInduccionFase("nAceptado"), _
                                                prsInduccionFase("cUltimaActualizacion"))
        While Not prsInduccionSesion.EOF
            Call oInduccion.InsertaRHInduccionSesion(pnCodInduccion, _
                                                    prsInduccionFase("nCodTema"), _
                                                    prsInduccionSesion("nCodSesion"), _
                                                      "0", _
                                                    prsInduccionSesion("dHoraInicio"), _
                                                    prsInduccionSesion("dHoraFin"))
            
            prsInduccionSesion.MoveNext
        Wend
        
        prsInduccionFase.MoveNext
    Wend
    'Set oConecta = New DConecta
    'oConecta.AbreConexion
    'oConecta.ConexionActiva.Execute sSql
    'oConecta.CierraConexion
    'Set oConecta = Nothing
    Exit Function
    
ErrorRegistra_Proceso_Induccion:
    Err.Raise Err.Number, "Registra_Proceso_Induccion", Err.Description

End Function

Public Function Llena_Sesiones(ByVal pnCodInduccion As Long, _
                                ByVal pnCodTema As Long) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorLlena_Sesiones
    sSQL = "SELECT nCodInduccion,INSES.nCodTema,nCodSeccion,dHoraInicio,dHoraFin,Tema=cDescTema FROM RHInduccionSesion INSES" & _
            " INNER JOIN RHInduccionTema INTE ON INSES.nCodTema=INTE.nCodTema "
    sSQL = sSQL & " WHERE nCodInduccion=" & pnCodInduccion & " AND INSES.nCodTema=" & pnCodTema & " ORDER BY dHoraInicio"
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Llena_Sesiones = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorLlena_Sesiones:
    Err.Raise Err.Number, "Llena_Sesiones", Err.Description
End Function

Public Function Llenar_Personas_Induccion(pcCodProceso As String)
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorLlenar_Personas_Induccion
    
    'sSql = "SELECT DISTINCT P.cPersCod, P.cPersNombre From Persona P INNER JOIN RHSeleccionFaseAprobada SFA"
    'sSql = sSql & " ON P.cPersCod = SFA.cPersCod INNER JOIN RHSeleccionProceso SP "
    'sSql = sSql & " ON SP.cCodProceso = SFA.cCodProceso  INNER JOIN RHInduccion INDUC ON SP.cCodProceso=INDUC.cCodProceso " & _
            "WHERE SFA.cCodProceso='" & pcCodProceso & "'"
    'sSql = sSql & " Order By P.cPersNombre "
    sSQL = "SELECT DISTINCT p.cPersCod,P.cPersNombre FROM Persona P INNER JOIN RHSeleccionPostulante SELP " & _
            "ON P.cPersCod=SELP.cPersCod INNER JOIN RHSeleccionProceso SP ON SELP.cCodProceso=SP.CCodProceso " & _
            "WHERE SELP.nEstado=1"  'Solo los declarados ganadores
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set Llenar_Personas_Induccion = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorLlenar_Personas_Induccion:
    Set Llenar_Personas_Induccion = Nothing
    Err.Raise Err.Number, "Llenar_Personas_Induccion", Err.Description

End Function

Public Function Obtiene_Temas() As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtiene_Temas
    sSQL = "SELECT nCodTema,cDescTema,nTipoTema,C.cConsDescripcion,cCargaHoraria,nTipoCarga,TipoCarga=C1.cConsDescripcion,nTipoInduc,TipoInduc=C2.cConsDescripcion FROM RHInduccionTema IT INNER JOIN Constante C ON IT.nTipoTema=C.nConsValor "
    sSQL = sSQL & "AND C.nConsCod='9031' INNER JOIN Constante C1 ON IT.nTipoCarga=C1.nConsValor AND C1.nConsCod='9034' INNER JOIN Constante C2 ON IT.nTipoInduc=C2.nConsValor AND C2.nConsCod='9035'" & _
            "ORDER BY nTipoInduc ASC ,nTipoTema ,nCodTema"
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtiene_Temas = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtiene_Temas:
    Err.Raise Err.Number, "Obtiene_Temas", Err.Description
End Function

Public Function Obtener_Siguiente_CodTema() As Long

Dim sSQL As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrorObtener_Siguiente_CodTema
    sSQL = "SELECT Siguiente=COUNT(nCodTema)+1 FROM RHInduccionTema "
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSQL)
        If rs.EOF Then
            Obtener_Siguiente_CodTema = 1
        Else
            Obtener_Siguiente_CodTema = rs("Siguiente") '+ 1
        End If
    End If
    
    oConecta.CierraConexion
    
    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Siguiente_CodTema:
    Err.Raise Err.Number, "Obtener_Siguiente_CodTema", Err.Description

End Function

Public Function Obtiene_Temas_x_Tipos(ByVal pnTipoTema As Long, _
                                    ByVal pnTipoInduc As Long) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtiene_Temas_x_Tipos
    sSQL = "SELECT nCodTema,cDescTema,nTipoTema,C.cConsDescripcion,cCargaHoraria,nTipoCarga,TipoCarga=C1.cConsDescripcion,nTipoInduc,TipoInduc=C2.cConsDescripcion FROM RHInduccionTema IT INNER JOIN Constante C ON IT.nTipoTema=C.nConsValor "
    sSQL = sSQL & "AND C.nConsCod='9031' INNER JOIN Constante C1 ON IT.nTipoTema=C1.nConsValor AND C1.nConsCod='9034' INNER JOIN Constante C2 ON IT.nTipoInduc=C2.nConsValor AND C2.nConsCod='9035'" & _
            "ORDER BY nTipoInduc ASC ,nTipoTema ,nCodTema" & _
            " WHERE nTipoTema=" & pnTipoTema & " AND nTipoInuc=" & pnTipoInduc
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtiene_Temas_x_Tipos = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtiene_Temas_x_Tipos:
    Err.Raise Err.Number, "Obtiene_Temas_x_Tipos", Err.Description
End Function

Public Function Llena_Fases_Induccion(ByVal pnCodInduccion As Long, Optional ByVal pnTipoTema As Integer = 0) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorLlena_Fases_Induccion
    
    Select Case pnTipoTema
        Case 0 'Induccion en Aula
            sSQL = "SELECT nCodInduccion, IFAS.nCodTema, cExpositor, dInicio, dFin, dHoraIni, dHoraFin, cComentario, nAceptado, IFAS.cUltimaActualizacion,cDescTema,Expositor=P.CPersNombre ,Aceptado=C.cConsDescripcion FROM RHInduccionFase IFAS" & _
                " INNER JOIN RHInduccionTema IT ON IFAS.nCodTema=IT.nCodTema INNER JOIN Persona P ON IFAS.cExpositor=P.cPersCod " & _
                "INNER JOIN Constante C ON C.nConsValor=IFAS.nAceptado AND nConsCod='9036'"
            sSQL = sSQL & " WHERE nCodInduccion=" & pnCodInduccion
        Case 1
            sSQL = "SELECT nCodInduccion, IFAS.nCodTema, cExpositor, dInicio, dFin, dHoraIni, dHoraFin, cComentario, nAceptado, IFAS.cUltimaActualizacion,cDescTema,Expositor=P.CPersNombre ,Aceptado=C.cConsDescripcion FROM RHInduccionFase IFAS" & _
                " INNER JOIN RHInduccionTema IT ON IFAS.nCodTema=IT.nCodTema INNER JOIN Persona P ON IFAS.cExpositor=P.cPersCod " & _
                "INNER JOIN Constante C ON C.nConsValor=IFAS.nAceptado AND nConsCod='9036'"
            sSQL = sSQL & " WHERE nCodInduccion=" & pnCodInduccion & " AND nTipoInduccion=1 ORDER BY dInicio"
        Case 2
            sSQL = "SELECT nCodInduccion, IFAS.nCodTema,dInicio, dFin, cComentario, nAceptado, IFAS.cUltimaActualizacion,cDescTema FROM RHInduccionFase IFAS" & _
                " INNER JOIN RHInduccionTema IT ON IFAS.nCodTema=IT.nCodTema "
            sSQL = sSQL & " WHERE nCodInduccion=" & pnCodInduccion & " AND nTipoInduccion=2 ORDER BY dInicio"
    End Select
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Llena_Fases_Induccion = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorLlena_Fases_Induccion:
    Err.Raise Err.Number, "Llena_Fases_Induccion", Err.Description
End Function

Public Function Buscar_Induccion_x_Proceso(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorBuscar_Induccion_x_Proceso
    sSQL = "SELECT nCodInduccion, cCodProceso,nEstado" & _
            " FROM RHInduccion WHERE cCodProceso='" & pcCodProceso & "'"
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Buscar_Induccion_x_Proceso = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorBuscar_Induccion_x_Proceso:
    Err.Raise Err.Number, "Buscar_Induccion_x_Proceso", Err.Description
End Function

Public Function Obtiene_Sesiones_x_Expositor(ByVal pcExpositor As String, ByVal pnCodInduccion As Long) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtiene_Sesiones_x_Expositor
    sSQL = "SELECT IFAS.nCodTema, IT.cDescTema,ISE.nCodSeccion, dFecha=ISE.dHoraInicio,ISE.dHoraInicio,ISE.dHoraFin" & _
            " FROM RHInduccionFase IFAS INNER JOIN RHInduccionTema IT ON IFAS.nCodTema=IT.nCodTema " & _
            " INNER JOIN RHInduccionSesion ISE ON IFAS.nCodInduccion=ISE.nCodInduccion AND IFAS.nCodTema=ISE.nCodTema" & _
            " INNER JOIN Persona P ON IFAS.cExpositor=P.cPersCod " & _
            " WHERE cExpositor='" & pcExpositor & "' AND IFAS.nCodInduccion=" & pnCodInduccion

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtiene_Sesiones_x_Expositor = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtiene_Sesiones_x_Expositor:
    Err.Raise Err.Number, "Obtiene_Sesiones_x_Expositor", Err.Description
End Function

Public Function Buscar_Temas_Repetidos(pnCodInduccion As Long, _
                                        pnCodTema As Long) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorBuscar_Temas_Repetidos
    
    sSQL = "SELECT nCodInduccion, nCodTema, cExpositor, dInicio, dFin, dHoraIni, dHoraFin, cComentario, nAceptado, cUltimaActualizacion FROM RHInduccionFase"
    sSQL = sSQL & " WHERE nCodInduccion=" & pnCodInduccion & " AND nCodTema=" & pnCodTema
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set Buscar_Temas_Repetidos = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorBuscar_Temas_Repetidos:
    Set Buscar_Temas_Repetidos = Nothing
    Err.Raise Err.Number, "Buscar_Temas_Repetidos", Err.Description
End Function

Public Function Llenar_Personas_Induccion_Asistencia(pcCodProceso As String)
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorLlenar_Personas_Induccion_Asistencia
    
    'sSql = "SELECT DISTINCT P.cPersCod, P.cPersNombre ,nTipoAsistencia='',TipoAsistencia='' From Persona P INNER JOIN RHSeleccionFaseAprobada SFA"
    'sSql = sSql & " ON P.cPersCod = SFA.cPersCod INNER JOIN RHSeleccionProceso SP "
    'sSql = sSql & " ON SP.cCodProceso = SFA.cCodProceso  INNER JOIN RHInduccion INDUC ON SP.cCodProceso=INDUC.cCodProceso " & _
    '        "WHERE SFA.cCodProceso='" & pcCodProceso & "'"
    'sSql = sSql & " Order By P.cPersNombre "
    sSQL = "SELECT DISTINCT P.cPersCod,P.cPersNombre,nTipoAsistencia='',TipoAsistencia='' FROM Persona P INNER JOIN RHSeleccionPostulante SELP ON P.cPersCod=SELP.cPersCod " & _
            "INNER JOIN RHSeleccionProceso SP ON SELP.cCodProceso=SP.CCodProceso WHERE SELP.nEstado=1 AND SP.cCodProceso='" & pcCodProceso & "'" & _
            "ORDER BY P.cPersNombre"
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set Llenar_Personas_Induccion_Asistencia = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorLlenar_Personas_Induccion_Asistencia:
    Set Llenar_Personas_Induccion_Asistencia = Nothing
    Err.Raise Err.Number, "Llenar_Personas_Induccion_Asistencia", Err.Description

End Function

Public Function Obtiene_Fases_x_Expositor(ByVal pcExpositor As String) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtiene_Fases_x_Expositor
    sSQL = "SELECT nCodInduccion, IFAS.nCodTema, cExpositor, dInicio, dFin, dHoraIni, dHoraFin, cComentario, nAceptado, IFAS.cUltimaActualizacion,cDescTema,Expositor=P.CPersNombre ,Aceptado=C.cConsDescripcion FROM RHInduccionFase IFAS" & _
            " INNER JOIN RHInduccionTema IT ON IFAS.nCodTema=IT.nCodTema INNER JOIN Persona P ON IFAS.cExpositor=P.cPersCod " & _
            "INNER JOIN Constante C ON C.nConsValor=IFAS.nAceptado AND nConsCod='9036'"
    sSQL = sSQL & " WHERE cExpositor='" & pcExpositor & "'"
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtiene_Fases_x_Expositor = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtiene_Fases_x_Expositor:
    Err.Raise Err.Number, "Obtiene_Fases_x_Expositor", Err.Description
End Function


Public Function Llenar_Personas_Induccion_Asistencia_Registrados(ByVal pnCodInduccion As Long, _
                                                    ByVal pnCodTema As Long, _
                                                    ByVal pnCodSesion As Long)
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorLlenar_Personas_Induccion_Asistencia_Registrados
    
    sSQL = "SELECT DISTINCT P.cPersCod, P.cPersNombre ,nTipoAsistencia=ISAS.nTipoAsistencia, " & _
            " TipoAsistencia = c.cConsDescripcion FROM Persona P INNER JOIN RHInduccionSesionAsistencia ISAS " & _
            " ON P.cPersCod=ISAS.cPersCod INNER JOIN RHInduccionSesion ISE ON ISAS.nCodTema=ISE.nCodTema " & _
            " AND ISAS.nCodInduccion=ISE.nCodInduccion AND ISAS.nCodSeccion=ISE.nCodSeccion INNER JOIN Constante " & _
            " C ON C.nConsValor=ISAS.nTipoAsistencia AND C.nConsCod=9033" & _
            " WHERE ISAS.nCodInduccion=" & pnCodInduccion & " AND ISAS.nCodTema=" & pnCodTema & "AND ISAS.nCodSeccion=" & pnCodSesion
    sSQL = sSQL & " Order By P.cPersNombre "
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set Llenar_Personas_Induccion_Asistencia_Registrados = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorLlenar_Personas_Induccion_Asistencia_Registrados:
    Set Llenar_Personas_Induccion_Asistencia_Registrados = Nothing
    Err.Raise Err.Number, "Llenar_Personas_Induccion_Asistencia_Registrados", Err.Description

End Function

Public Function Llenar_Personas_Induccion_Notas(ByVal pcCodProceso As String)
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorLlenar_Personas_Induccion_Notas
    
    sSQL = "SELECT DISTINCT P.cPersCod,P.cPersNombre,nNota='',bAprobado=0,cComentario='' FROM Persona P INNER JOIN RHSeleccionPostulante SELP ON P.cPersCod=SELP.cPersCod " & _
            "INNER JOIN RHSeleccionProceso SP ON SELP.cCodProceso=SP.CCodProceso WHERE SELP.nEstado=1 " & _
            "AND SP.cCodProceso='" & pcCodProceso & "'" & _
            "ORDER BY P.cPersNombre"
                
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set Llenar_Personas_Induccion_Notas = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorLlenar_Personas_Induccion_Notas:
    Set Llenar_Personas_Induccion_Notas = Nothing
    Err.Raise Err.Number, "Llenar_Personas_Induccion_Notas", Err.Description

End Function

Public Function Llenar_Personas_Induccion_Campo(ByVal pcPersCodResponsable As String, ByVal pnCodInduccion As Long)

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorLlenar_Personas_Induccion_Campo
    
'    sSQL = "SELECT DISTINCT P.cPersCod,P.cPersNombre,nNota='',bAprobado=0,cComentario='' FROM RHCargos CG INNER JOIN Persona P " & _
           " ON CG.cPersCod=P.cPersCod AND CG.cRHCargoCod = (SELECT TOP 1 cRHCargoCod FROM RHCargos WHERE cPersCod=P.cPersCod ORDER BY dRHCargoFecha DESC) " & _
           " AND P.cPersCod NOT IN(SELECT cPersCod FROM RHInduccionNota WHERE bAprobado=0 AND nTipoInduccion=1 AND nCodInduccion=" & pnCodInduccion & ")" & _
           " AND CG.cRHAgenciaCod='" & pcRHAgenciaCod & "' INNER JOIN RHSeleccionPostulante SELP ON P.cPersCod=SELP.cPersCod " & _
           " INNER JOIN RHSeleccionProceso SP ON SELP.cCodProceso=SP.CCodProceso " & _
           " WHERE SELP.nEstado=1 AND SP.cCodProceso='" & pcCodProceso & "' ORDER BY P.cPersNombre"
           
     sSQL = "SELECT DISTINCT P.cPersCod,P.cPersNombre,nNota='',bAprobado=0,cComentario='' FROM RHInduccionParticipante IP INNER JOIN Persona P " & _
           " ON P.cPersCod=IP.cPersCod AND P.cPersCod NOT IN(SELECT cPersCod FROM RHInduccionNota WHERE bAprobado=0 AND nTipoInduccion=1 AND nCodInduccion=" & pnCodInduccion & ") " & _
           " WHERE IP.nCodInduccion=" & pnCodInduccion & " AND IP.cPersCodResponsable ='" & pcPersCodResponsable & "' ORDER BY P.cPersNombre"
          
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set Llenar_Personas_Induccion_Campo = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorLlenar_Personas_Induccion_Campo:
    Set Llenar_Personas_Induccion_Campo = Nothing
    Err.Raise Err.Number, "Llenar_Personas_Induccion_Campo", Err.Description
End Function

Public Function Llenar_Personas_Induccion_Notas_Registrados(ByVal pnCodInduccion As Long, ByVal pnTipoInduccion As Integer) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorLlenar_Personas_Induccion_Notas_Registrados
    
    sSQL = "SELECT DISTINCT P.cPersCod, Participante=P.cPersNombre ,Nota=INOT.nNota," & _
           " Aprueba=CONVERT(INT,INOT.bAprobado),cComentario=ISNULL(INOT.cComentario,'')" & _
           " FROM Persona P INNER JOIN RHInduccionNota INOT ON P.cPersCod=INOT.cPersCod " & _
           " INNER JOIN RHInduccion IND ON  INOT.nCodInduccion=IND.nCodInduccion " & _
           " WHERE INOT.nCodInduccion=" & pnCodInduccion & " AND INOT.nTipoInduccion=" & pnTipoInduccion
    sSQL = sSQL & " Order By P.cPersNombre "
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set Llenar_Personas_Induccion_Notas_Registrados = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorLlenar_Personas_Induccion_Notas_Registrados:
    Set Llenar_Personas_Induccion_Notas_Registrados = Nothing
    Err.Raise Err.Number, "Llenar_Personas_Induccion_Notas_Registrados", Err.Description
End Function

'Public Function Llenar_Personas_Induccion_Notas_Registrados(pnCodInduccion As Long, _
'                                                    pnCodTema As Long, _
'                                                      pcExpositor As String)
'Dim sSql As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorLlenar_Personas_Induccion_Notas_Registrados
'
'    sSql = "SELECT DISTINCT P.cPersCod, Participante=P.cPersNombre ,Nota=INOT.nNota, " & _
'            " Aprueba=CONVERT(INT,INOT.bAprobado),cComentario=ISNULL(INOT.cComentario,'') FROM Persona P INNER JOIN RHInduccionNota INOT " & _
'            " ON P.cPersCod=INOT.cPersCod INNER JOIN RHInduccionFase IFAS ON INOT.nCodInduccion= IFAS.nCodInduccion " & _
'            " AND INOT.nCodTema=IFAS.nCodTema INNER JOIN RHInduccion IND ON  IFAS.nCodInduccion=IND.nCodInduccion" & _
'            " WHERE INOT.nCodInduccion=" & pnCodInduccion & " AND INOT.nCodTema=" & pnCodTema & _
'            "AND INOT.cExpositor='" & pcExpositor & "'"
'    sSql = sSql & " Order By P.cPersNombre "
'
'    Set oConecta = New DConecta
'    If oConecta.AbreConexion Then
'        Set Llenar_Personas_Induccion_Notas_Registrados = oConecta.CargaRecordSet(sSql)
'    End If
'    oConecta.CierraConexion
'
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorLlenar_Personas_Induccion_Notas_Registrados:
'    Set Llenar_Personas_Induccion_Notas_Registrados = Nothing
'    Err.Raise Err.Number, "Llenar_Personas_Induccion_Notas_Registrados", Err.Description
'End Function

Public Function Busca_Temas_AsignadosInduccion(ByVal pnCodTema As Long) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorBusca_Temas_AsignadosInduccion
    sSQL = "SELECT nCodInduccion, nCodTema, cExpositor, dInicio, dFin, dHoraIni, dHoraFin, cComentario, nAceptado, cUltimaActualizacion FROM RHInduccionFase " & _
            " WHERE nCodTema=" & pnCodTema
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Busca_Temas_AsignadosInduccion = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorBusca_Temas_AsignadosInduccion:
    Err.Raise Err.Number, "Busca_Temas_AsignadosInduccion", Err.Description
End Function

'Public Function Obtener_Personas_Induccion_AsignacionFuncion(pcCodProceso As String) As ADODB.Recordset
'Dim sSql As String
'Dim oConecta As DConecta
'
'    On Error GoTo ErrorObtener_Personas_Induccion_AsignacionFuncion
'
'    sSql = "SELECT DISTINCT p.cPersCod,Participante=P.cPersNombre,[Asig.Funciones]='NO',[Asig.Materiales]='NO',[Asig.Claves]='NO' FROM Persona P INNER JOIN RHSeleccionPostulante SELP " & _
'           " ON P.cPersCod=SELP.cPersCod AND P.cPersCod NOT IN( SELECT NOTA.cPersCod FROM RHInduccionNota NOTA INNER JOIN RHInduccion IND ON NOTA.nCodInduccion=IND.nCodInduccion INNER JOIN " & _
'           " RHSeleccionProceso SEP ON SEP.cCodProceso=IND.cCodProceso " & _
'           " WHERE bAprobado=0 AND SEP.cCodProceso='" & pcCodProceso & "') INNER JOIN RHSeleccionProceso SP ON SELP.cCodProceso=SP.CCodProceso " & _
'           " WHERE SELP.nEstado = 1 AND SP.cCodProceso='" & pcCodProceso & "'"
'    Set oConecta = New DConecta
'    If oConecta.AbreConexion Then
'        Set Obtener_Personas_Induccion_AsignacionFuncion = oConecta.CargaRecordSet(sSql)
'    End If
'    oConecta.CierraConexion
'
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorObtener_Personas_Induccion_AsignacionFuncion:
'    Set Obtener_Personas_Induccion_AsignacionFuncion = Nothing
'    Err.Raise Err.Number, "Obtener_Personas_Induccion_AsignacionFuncion", Err.Description
'
'End Function

Public Function Obtener_Personas_Induccion_AsignacionFuncion(ByVal pnCodInduccion As Long) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Personas_Induccion_AsignacionFuncion
    
    sSQL = "SELECT DISTINCT p.cPersCod,Participante=P.cPersNombre,[Asig.Funciones]=" & _
           " CASE WHEN INOT.bAprobado=1 THEN 'SI' ELSE 'NO' END ,'',''" & _
           " FROM Persona P INNER JOIN RHInduccionNota INOT ON P.cPersCod = INOT.cPersCod " & _
           " WHERE INOT.nTipoInduccion = 2 And INOT.nCodInduccion =" & pnCodInduccion
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set Obtener_Personas_Induccion_AsignacionFuncion = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Personas_Induccion_AsignacionFuncion:
    Set Obtener_Personas_Induccion_AsignacionFuncion = Nothing
    Err.Raise Err.Number, "Obtener_Personas_Induccion_AsignacionFuncion", Err.Description

End Function


Public Function Obtener_Personas_Induccion_AsignacionFuncion_Registrados(ByVal pnCodInduccion As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Personas_Induccion_AsignacionFuncion_Registrados
    
    sSQL = "SELECT DISTINCT p.cPersCod,Participante=P.cPersNombre, [Asig.Funciones]=CASE WHEN bFunciones=1 THEN 'SI'ELSE 'NO'END, " & _
            "[Asig.Materiales]=CASE WHEN bMateriales=1 THEN 'SI'ELSE 'NO'END, [Asig.Claves]=CASE WHEN bClaves=1 THEN 'SI'ELSE 'NO'END " & _
            " FROM RHInduccion INDUC INNER JOIN RHInduccionAsignacionFuncion IAF ON INDUC.nCodInduccion=IAF.nCodInduccion " & _
            " INNER JOIN Persona P ON IAF.cPersCod=P.CPersCod " & _
            "WHERE IAF.nCodInduccion=" & pnCodInduccion

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set Obtener_Personas_Induccion_AsignacionFuncion_Registrados = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Personas_Induccion_AsignacionFuncion_Registrados:
    Set Obtener_Personas_Induccion_AsignacionFuncion_Registrados = Nothing
    Err.Raise Err.Number, "Obtener_Personas_Induccion_AsignacionFuncion_Registrados", Err.Description
End Function

Public Function Validar_ProcInduccion_x_Estado(ByVal pnCodInduccion As Long, _
                                               ByVal pnEstado As Long) As Boolean
Dim sSQL As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorValidar_ProcInduccion_x_Estado
    sSQL = "SELECT nCodInduccion, cCodProceso,nEstado" & _
            " FROM RHInduccion WHERE nCodInduccion=" & pnCodInduccion & " AND nEstado=" & pnEstado
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSQL)
    End If
    Validar_ProcInduccion_x_Estado = Not rs.EOF
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorValidar_ProcInduccion_x_Estado:
    Err.Raise Err.Number, "Validar_ProcInduccion_x_Estado", Err.Description
End Function

Public Function Validar_FaseAprobada_x_Expositor(ByVal pnCodInduccion As Long, _
                                               ByVal pnCodTema As Long, _
                                               ByVal pcExpositor As String) As Boolean
Dim sSQL As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorValidar_FaseAprobada_x_Expositor
    sSQL = "SELECT nCodInduccion, cCodTema,cExpositor" & _
            " FROM RHInduccionFase WHERE nCodInduccion=" & pnCodInduccion & " AND nCodTema=" & pnCodTema & " AND cExpositor='" & pcExpositor & "' WHERE nAceptado=2 "   'nEstado=2==> Los Aceptados por los Expositores del Tema
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSQL)
    End If
    Validar_FaseAprobada_x_Expositor = Not rs.EOF
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorValidar_FaseAprobada_x_Expositor:
    Err.Raise Err.Number, "Validar_FaseAprobada_x_Expositor", Err.Description
End Function

Public Function Obtiene_Lista_Temas_x_Expositor(ByVal pcExpositor As String, pnCodInduccion As Long) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtiene_Lista_Temas_x_Expositor
    sSQL = "SELECT IFAS.nCodTema, IT.cDescTema, 1 FROM RHInduccionFase IFAS" & _
            " INNER JOIN RHInduccionTema IT ON IFAS.nCodTema=IT.nCodTema " & _
            " WHERE cExpositor='" & pcExpositor & "' AND IFAS.nCodInduccion=" & pnCodInduccion
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtiene_Lista_Temas_x_Expositor = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtiene_Lista_Temas_x_Expositor:
    Err.Raise Err.Number, "Obtiene_Lista_Temas_x_Expositor", Err.Description
End Function

Public Function Obtiene_CargoEvaluacion(ByVal pnCodInduccion As Long) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtiene_CargoEvaluacion
    sSQL = "SELECT TOP 1 cRHCargoCodEval,CT.cRHCargoDescripcion FROM RHInduccion IND INNER JOIN " & _
           " RHCargosTabla CT ON CT.CRHCargoCod=IND.cRHCargoCodEval " & _
           " WHERE IND.nCodInduccion=" & pnCodInduccion
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtiene_CargoEvaluacion = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtiene_CargoEvaluacion:
    Err.Raise Err.Number, "Obtiene_CargoEvaluacion", Err.Description
End Function

Public Function Validar_Cargo_x_Expositor(ByVal pcPersCod As String, ByVal pnCodInduccion As Long) As Boolean

Dim sSQL As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
Dim cRHCargoEvaluador As String

    On Error GoTo ErrorValidar_Cargo_x_Expositor
    
    Set rs = Obtiene_CargoEvaluacion(pnCodInduccion)
    
    cRHCargoEvaluador = rs("cRHCargoCodEval")
    
    sSQL = "SELECT TOP 1 cRHCargoCod FROM RHCargos WHERE cPersCod='" & pcPersCod & _
            "' ORDER BY dRHCargoFecha DESC"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set rs = oConecta.CargaRecordSet(sSQL)
    End If
    
    If cRHCargoEvaluador = rs("cRHCargoCod") Then
        Validar_Cargo_x_Expositor = True
    Else
        Validar_Cargo_x_Expositor = False
    End If
        
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorValidar_Cargo_x_Expositor:
    Err.Raise Err.Number, "Validar_Cargo_x_Expositor", Err.Description
End Function

Public Function Obtiene_Datos_x_Evaluador(ByVal pcPersCod As String) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtiene_Datos_x_Evaluador
    
    sSQL = "SELECT TOP 1 cRHCargoCod,cRHAgenciaCod FROM RHCargos WHERE cPersCod='" & pcPersCod & "'" & _
           "ORDER BY dRHCargoFecha DESC"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtiene_Datos_x_Evaluador = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtiene_Datos_x_Evaluador:
    Err.Raise Err.Number, "Obtiene_Datos_x_Evaluador", Err.Description
End Function

Public Function Obtiene_Fechas_Cronograma(ByVal pnCodInduccion As Long) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtiene_Fechas_Cronograma
    sSQL = "SELECT DISTINCT Fecha=CONVERT(varchar(10),dHoraInicio,103) FROM RHInduccionSesion " & _
           " WHERE ncodinduccion = " & pnCodInduccion
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtiene_Fechas_Cronograma = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtiene_Fechas_Cronograma:
    Err.Raise Err.Number, "Obtiene_Fechas_Cronograma", Err.Description
End Function

Public Function Obtiene_Sesiones_x_Fecha(ByVal pdFecha As Date, ByVal pnCodInduccion As Long) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtiene_Sesiones_x_Fecha
    sSQL = " SELECT IFAS.nCodTema, IT.cDescTema,ISE.nCodSeccion,ISE.dHoraInicio,ISE.dHoraFin,Expositor= P.cPersNombre +' - '+cRHCargoDescripcion " & _
           " FROM RHInduccionFase IFAS INNER JOIN RHInduccionTema IT ON IFAS.nCodTema=IT.nCodTema INNER JOIN RHInduccionSesion ISE ON " & _
           " IFAS.nCodInduccion=ISE.nCodInduccion AND IFAS.nCodTema=ISE.nCodTema INNER JOIN Persona P ON IFAS.cExpositor=P.cPersCod " & _
           " INNER JOIN RHCargos car ON Car.cPersCod=P.cPersCod AND Car.dRHCargoFecha = (SELECT MAX(dRHCargoFecha)FROM RHCargos WHERE cPersCod=P.cPersCod) " & _
           " INNER JOIN RHCargosTabla Cat ON Cat.cRHCargoCod=Car.cRHCargoCod WHERE IFAS.nCodInduccion =" & pnCodInduccion & _
           " AND CONVERT(varchar(10),ISE.dHoraInicio,103)= CONVERT(varchar(10),'" & pdFecha & "',103) ORDER BY ISE.dHoraInicio "

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtiene_Sesiones_x_Fecha = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtiene_Sesiones_x_Fecha:
    Err.Raise Err.Number, "Obtiene_Sesiones_x_Fecha", Err.Description
End Function

Public Function Obtener_ProcesosSeleccion_Activos() As ADODB.Recordset
                    
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorObtener_ProcesosSeleccion_Activos
    
    sSQL = "SELECT DISTINCT SP.cCodProceso,Descripcion= 'INDUCCION'+' '+cRHCargoDescripcion + ' ' + " & _
           " Ag.cAgeDescripcion,1 FROM RHSeleccionProceso SP INNER JOIN RHSeleccionRequerimientoProceso SRP ON " & _
           " SP.cCodProceso=SRP.cCodProceso AND SRP.cCodRequerimiento =(SELECT TOP 1 cCodRequerimiento " & _
           " FROM RHSeleccionRequerimientoProceso WHERE cCodProceso=SP.cCodProceso)" & _
           " INNER JOIN RHSeleccionRequerimiento SR ON SRP.cCodRequerimiento=SR.cCodRequerimiento " & _
           " INNER JOIN RHCargos CAR ON CAR.cRHCargoCod=SR.cCodCargoSol INNER JOIN RHCargosTabla CAT " & _
           " ON CAT.cRHCargoCod=CAR.cRHCargoCod INNER JOIN Agencias AG ON SR.cCodAgeSol=AG.cAgeCod " & _
           " WHERE SP.nEstado = 7 ORDER BY SP.cCodProceso"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_ProcesosSeleccion_Activos = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_ProcesosSeleccion_Activos:
    Err.Raise Err.Number, "Obtener_ProcesosSeleccion_Activos", Err.Description
End Function

