VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NSeleccionProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Obtener_Comite_Seleccion(ByVal pcCodProceso As String, _
                                         Optional ByVal pcPersCod As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Comite_Seleccion
    If pcPersCod = "" Then
        sSql = "SELECT P.cPersCod,Miembro=P.CPersNombre,1,SPC.nCargo,Cargo=C.cConsDescripcion + CASE WHEN ISNULL(cPersCodTitu,'')='' THEN '' ELSE ' - SUPLENTE' END," & _
           " Titular=ISNULL(T.cPersNombre,'') FROM Persona P INNER JOIN RHSeleccionProcesoComite SPC ON P.cPersCod = SPC.cPersCod INNER JOIN RHSeleccionProceso SP " & _
           " ON SPC.cCodProceso =SP.cCodProceso INNER JOIN Constante C ON C.nConsValor= SPC.nCargo AND C.nConsCod='9037' LEFT JOIN Persona T ON T.cPersCod=SPC.cPersCodTitu " & _
           " WHERE SP.cCodProceso='" & pcCodProceso & "' ORDER BY ISNULL(SPC.cPersCodTitu,'')"
    Else
        sSql = "SELECT P.cPersCod,Miembro=P.CPersNombre,1,SPC.nCargo,Cargo=C.cConsDescripcion + CASE WHEN ISNULL(cPersCodTitu,'')='' THEN '' ELSE ' - SUPLENTE' END," & _
           " Titular=ISNULL(T.cPersNombre,'') FROM Persona P INNER JOIN RHSeleccionProcesoComite SPC ON P.cPersCod = SPC.cPersCod INNER JOIN RHSeleccionProceso SP " & _
           " ON SPC.cCodProceso =SP.cCodProceso INNER JOIN Constante C ON C.nConsValor= SPC.nCargo AND C.nConsCod='9037' LEFT JOIN Persona T ON T.cPersCod=SPC.cPersCodTitu " & _
           " WHERE SP.cCodProceso='" & pcCodProceso & "' AND P.cPersCod='" & pcPersCod & "' ORDER BY ISNULL(SPC.cPersCodTitu,'')"
    End If
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Comite_Seleccion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Comite_Seleccion:
    Err.Raise Err.Number, "Obtener_Comite_Seleccion", Err.Description
End Function

'Obtiene los Miembros del Comite que no han registrado su Aprobacion de una Fase
'del Proceso de Seleccion
Public Function Obtener_ComiteSel_Aprobacion(ByVal pcCodProceso As String, _
                                            ByVal pnCodFase As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_ComiteSel_Aprobacion
    sSql = "SELECT P.cPersCod,Miembro=P.CPersNombre,1,SPC.nCargo,Cargo=C.cConsDescripcion " & _
            "FROM Persona P INNER JOIN RHSeleccionProcesoComite SPC " & _
            "ON P.cPersCod = SPC.cPersCod INNER JOIN RHSeleccionProceso SP " & _
            "ON SPC.cCodProceso =SP.cCodProceso     INNER JOIN Constante C " & _
            "ON C.nConsValor= SPC.nCargo AND C.nConsCod='9037'" & _
            "INNER JOIN RHSeleccionFaseAprobacion SPA ON SPA.cCodProceso=SP.cCodProceso" & _
            " AND SPA.cPersCod=P.cPersCod " & _
            " WHERE SP.cCodProceso='" & pcCodProceso & "'" & " AND SPA.nCodFase=" & pnCodFase
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_ComiteSel_Aprobacion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_ComiteSel_Aprobacion:
    Err.Raise Err.Number, "Obtener_ComiteSel_Aprobacion", Err.Description
End Function

Public Function Obtener_Cargos_ProcesoSelec() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Cargos_ProcesoSelec
    sSql = "SELECT cRHCargoCod,cRHCargoDescripcion FROM RHCargosTabla"
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Cargos_ProcesoSelec = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Cargos_ProcesoSelec:
    Err.Raise Err.Number, "Obtener_Cargos_ProcesoSelec", Err.Description
End Function

Public Function Obtener_Fases_Seleccion(Optional ByVal pbSoloEjecutan As Boolean = False, _
                                        Optional ByVal pnCodFase As Long = 0) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Fases_Seleccion
    'Solo las que se ejecutan..........pbSoloEjecuta=True
    If pnCodFase = 0 Then
        sSql = "SELECT nCodFase,cNomFase,1,cComentarioFase=ISNULL(cComentarioFase,'') FROM RHSeleccionFases " & IIf(pbSoloEjecutan = False, "", "WHERE bEjecuta=1")
    Else
        sSql = "SELECT nCodFase,cNomFase,cComentarioFase=ISNULL(cComentarioFase,'') FROM RHSeleccionFases WHERE nCodFase=" & pnCodFase
    End If
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Fases_Seleccion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Fases_Seleccion:
    Err.Raise Err.Number, "Obtener_Fases_Seleccion", Err.Description
End Function

Public Function Obtener_FasesSeleccion_Mantenimiento() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_FasesSeleccion_Mantenimiento
    sSql = "SELECT nCodFase,cNomFase,bEjecuta=CASE WHEN ISNULL(bEjecuta,0)=1 THEN 'SI'ELSE'NO'END,cResponsable,cComentarioFase=ISNULL(cComentarioFase,'') FROM RHSeleccionFases"
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_FasesSeleccion_Mantenimiento = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_FasesSeleccion_Mantenimiento:
    Err.Raise Err.Number, "Obtener_FasesSeleccion_Mantenimiento", Err.Description
End Function

Public Function Obtener_Datos_FaseSeleccion(ByVal pnCodFase As Long) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_FaseSeleccion
    sSql = "SELECT nCodFase,cNomFase,cComentarioFase=ISNULL(cComentarioFase,''),bEjecuta,cResponsable=ISNULL(cResponsable,'') FROM RHSeleccionFases WHERE nCodFase=" & pnCodFase
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_FaseSeleccion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Datos_FaseSeleccion:
    Err.Raise Err.Number, "Obtener_Datos_FaseSeleccion", Err.Description
End Function

Public Function Validar_Fase_Cerrada(ByVal pnCodFase As Long, _
                                    ByVal pcCodProceso As String) As Boolean

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorValidar_Fase_Cerrada
    sSql = "SELECT bCerrada=ISNULL(bCerrada,0) FROM RHSeleccionFase WHERE nCodFase=" & pnCodFase & " AND cCodProceso='" & pcCodProceso & "'"
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        Validar_Fase_Cerrada = IIf(rs("bCerrada") = 1, True, False)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorValidar_Fase_Cerrada:
    Err.Raise Err.Number, "Validar_Fase_Cerrada", Err.Description
End Function

Public Function Obtener_Fases_ProcesoSeleccion(ByVal pcCodProceso As String, _
                                               Optional ByVal pbTodos As Integer = 0, _
                                               Optional ByVal pbCerrados As Integer = 1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Fases_ProcesoSeleccion
    If pbCerrados = 1 Then
        If pbTodos = 1 Then
            sSql = "SELECT nCodFase=SFS.nCodFase,cNomFase=SFS.cNomFase,1,SF.dFechaIni,SF.dFechaFin, nNotaMin=ISNULL(convert(varchar(7),SF.nNotaMin),'')," & _
                " nNotaMax=ISNULL(convert(varchar(7),SF.nNotaMax),''), nNotaAprobatoria=ISNULL(convert(varchar(7),SF.nNotaAprobatoria),''), " & _
                " nPonderacion=ISNULL(convert(varchar(7),SF.nPonderacion),''),SF.nOrden,bEjecuta=ISNULL(SFS.bEjecuta,''),cResponsable=ISNULL(SFS.cResponsable,'') " & _
                "FROM RHSeleccionProceso SP INNER JOIN RHSeleccionFase SF ON SP.cCodProceso= SF.cCodProceso INNER JOIN RHSeleccionFases SFS ON SF.nCodFase= SFS.nCodFase " & _
                " WHERE SP.cCodProceso ='" & pcCodProceso & "' ORDER BY SF.nOrden"
        Else 'Solo los que se ejecutan
            sSql = "SELECT nCodFase=SFS.nCodFase,cNomFase=SFS.cNomFase,1,SF.dFechaIni,SF.dFechaFin, nNotaMin=ISNULL(convert(varchar(7),SF.nNotaMin),'')," & _
                " nNotaMax=ISNULL(convert(varchar(7),SF.nNotaMax),''), nNotaAprobatoria=ISNULL(convert(varchar(7),SF.nNotaAprobatoria),''), " & _
                " nPonderacion=ISNULL(convert(varchar(7),SF.nPonderacion),''),SF.nOrden,bEjecuta=ISNULL(SFS.bEjecuta,''),cResponsable=ISNULL(SFS.cResponsable,'') " & _
                "FROM RHSeleccionProceso SP INNER JOIN RHSeleccionFase SF ON SP.cCodProceso= SF.cCodProceso INNER JOIN RHSeleccionFases SFS ON SF.nCodFase= SFS.nCodFase " & _
                "WHERE SP.cCodProceso ='" & pcCodProceso & "' AND SFS.bEjecuta=1 ORDER BY SF.nOrden"
        End If
    Else    'Los Cerrados no se muestran
        sSql = "SELECT nCodFase=SFS.nCodFase,cNomFase=SFS.cNomFase,1,SF.dFechaIni,SF.dFechaFin, nNotaMin=ISNULL(convert(varchar(7),SF.nNotaMin),'')," & _
                " nNotaMax=ISNULL(convert(varchar(7),SF.nNotaMax),''), nNotaAprobatoria=ISNULL(convert(varchar(7),SF.nNotaAprobatoria),''), " & _
                " nPonderacion=ISNULL(convert(varchar(7),SF.nPonderacion),''),SF.nOrden,bEjecuta=ISNULL(SFS.bEjecuta,''),cResponsable=ISNULL(SFS.cResponsable,'') " & _
                "FROM RHSeleccionProceso SP INNER JOIN RHSeleccionFase SF ON SP.cCodProceso= SF.cCodProceso INNER JOIN RHSeleccionFases SFS ON SF.nCodFase= SFS.nCodFase " & _
                "WHERE SP.cCodProceso ='" & pcCodProceso & "' AND SFS.bEjecuta=1 AND ISNULL(SF.bCerrada,0)=0 ORDER BY SF.nOrden"
    End If
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Fases_ProcesoSeleccion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Fases_ProcesoSeleccion:
    Err.Raise Err.Number, "Obtener_Fases_ProcesoSeleccion", Err.Description
End Function

Public Function Obtener_Codigo_SeleccionProceso() As String

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrorObtener_Codigo_SeleccionProceso
    sSql = "SELECT  Codigo=ISNULL(MAX(CONVERT(int,SUBSTRING(cCodProceso,5,6))),0) FROM RHSeleccionProceso WHERE SUBSTRING(cCodProceso,1,4)=YEAR(getdate())"
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_Codigo_SeleccionProceso = "1"
        Else
            Obtener_Codigo_SeleccionProceso = CStr(rs("Codigo") + 1)
        End If
    End If
    
    oConecta.CierraConexion
    
    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Codigo_SeleccionProceso:
    Err.Raise Err.Number, "Obtener_Codigo_SeleccionProceso", Err.Description
End Function

Public Function Obtener_Datos_ProcSeleccion(ByVal pcCodProceso As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_ProcSeleccion
    sSql = "SELECT cCodProceso, cDescProceso, nTipoProceso, nTipoContrato, dFechaIni, dFechaFin, nEstado,cObservacion=ISNULL(cObservacion,''),cCodPublicacion=ISNULL(cCodProcesoPublicacion,'') FROM RHSeleccionProceso " & _
            " WHERE cCodProceso='" & pcCodProceso & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_ProcSeleccion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Datos_ProcSeleccion:
    Err.Raise Err.Number, "Obtener_Datos_ProcSeleccion", Err.Description
End Function

Public Function Obtener_Requisitos_Seleccion() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Requisitos_Seleccion
    sSql = "SELECT nCodRequisito,cDescRequisito FROM RHSeleccionRequisitos"
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Requisitos_Seleccion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Requisitos_Seleccion:
    Err.Raise Err.Number, "Obtener_Requisitos_Seleccion", Err.Description
End Function

Public Function Obtener_Siguiente_CodRequisito() As Long

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrorObtener_Siguiente_CodRequisito
    sSql = "SELECT Siguiente=COUNT(nCodRequisito)+1 FROM RHSeleccionRequisitos "
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_Siguiente_CodRequisito = 1
        Else
            Obtener_Siguiente_CodRequisito = rs("Siguiente") '+ 1
        End If
    End If
    
    oConecta.CierraConexion
    
    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Siguiente_CodRequisito:
    Err.Raise Err.Number, "Obtener_Siguiente_CodRequisito", Err.Description

End Function

Public Function Obtener_Requisitos_ProcesoSeleccion(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Requisitos_ProcesoSeleccion
    sSql = "SELECT SRS.nCodRequisito,SRS.cDescRequisito,TipoRequisito=C.cConsDescripcion FROM RHSeleccionRequisitos SRS INNER JOIN RHSeleccionRequisito SR " & _
            " ON SRS.nCodRequisito=SR.nCodRequisito INNER JOIN RHSeleccionProceso SP " & _
            " ON SR.cCodProceso=SP.cCodProceso INNER JOIN Constante C ON SR.nTipoRequisito=C.nConsValor AND C.nConsCod='9038'" & _
            " WHERE SP.cCodProceso='" & pcCodProceso & "'"
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Requisitos_ProcesoSeleccion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Requisitos_ProcesoSeleccion:
    Err.Raise Err.Number, "Obtener_Requisitos_ProcesoSeleccion", Err.Description
End Function

Public Function Buscar_Requisitos_Repetido(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorBuscar_Requisitos_Repetido
    sSql = "SELECT SRS.nCodRequisito,SRS.cDescRequisito,TipoRequisito=C.cConsDescripcion FROM RHSeleccionRequisitos SRS INNER JOIN RHSeleccionRequisito SR " & _
            " ON SRS.nCodRequisito=SR.nCodRequisito INNER JOIN RHSeleccionProceso SP " & _
            " ON SR.cCodProceso=SP.cCodProceso INNER JOIN Constante C ON SR.nTipoRequisito=C.nConsValor AND C.nConsCod='9038'" & _
            " WHERE SP.cCodProceso='" & pcCodProceso & "'"
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Buscar_Requisitos_Repetido = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorBuscar_Requisitos_Repetido:
    Err.Raise Err.Number, "Buscar_Requisitos_Repetido", Err.Description
End Function

Public Function Obtener_Personas_ProcesoSeleccion(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Personas_ProcesoSeleccion
    sSql = "SELECT DISTINCT cPersCod=P.cPostuCod,cPersNombre=RTRIM(cApellidoPat)+' '+RTRIM(cApellidoMat)+' '+ RTRIM(cNombre),Vigesimal='',nNota='',bGanador='',cComentario='' " & _
            " FROM Postulante P INNER JOIN RHSeleccionPostulante SELP ON P.cPostuCod=SELP.cPostuCod INNER JOIN RHSeleccionProceso SP ON " & _
            " SELP.CCodProceso = SP.CCodProceso " & _
            " AND SP.cCodProceso ='" & pcCodProceso & "' AND ISNULL(SELP.nEstado,0) NOT IN(4,5) AND P.cPostuCod NOT IN  (SELECT cPersCod FROM RHSeleccionEvaluacion " & _
            " WHERE bGanador=0 AND nCodFase<>12 AND cCodProceso='" & pcCodProceso & "') ORDER BY P.cPersNombre"
            'Estado 4 son los Postulantes Vetados,solo los que pasan a la otra Fase(excepto en la Entrevista)
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Personas_ProcesoSeleccion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Personas_ProcesoSeleccion:
    Err.Raise Err.Number, "Obtener_Personas_ProcesoSeleccion", Err.Description
End Function

Public Function Obtener_Personas_ProcesoSeleccion_Evaluados(ByVal pnCodFase As Long, _
                                                        ByVal pcCodProceso As String, _
                                                        Optional ByVal pcPersCodResp As String = "", _
                                                        Optional ByVal pbSoloAptos As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Personas_ProcesoSeleccion_Evaluados
    
    If pbSoloAptos = False Then
        If pcPersCodResp = "" Then
            sSql = "SELECT DISTINCT cPersCod=P.cPostuCod,cPersNombre=RTRIM(cApellidoPat)+' '+RTRIM(cApellidoMat)+' '+ RTRIM(cNombre), " & _
                   " nNota=SUM(SE.nNota),nVigesi='',bGanador=CONVERT(int,SE.bGanador),cComentario=SE.cComentario FROM RHSeleccionEvaluacion SE INNER JOIN Postulante P " & _
                " ON SE.cPersCod=P.cPostuCod " & _
                 " INNER JOIN RHSeleccionPostulante SEP ON SEP.cPostuCod = P.cPostuCod" & _
                " AND SEP.cCodProceso = SE.cCodProceso " & _
                 " WHERE SE.nCodFase=" & pnCodFase & _
                  " AND SE.cCodProceso ='" & pcCodProceso & _
                 "' AND ISNULL(SEP.nEstado,0) NOT IN (4,5) GROUP BY P.cPostuCod,cApellidoPat,cApellidoMat,cApellidoMat,cNombre , SE.bGanador, SE.cComentario " & _
                 " ORDER BY SE.nNota DESC"
            Else
                sSql = "SELECT DISTINCT cPersCod=P.cPostuCod,cPersNombre=RTRIM(cApellidoPat)+' '+RTRIM(cApellidoMat)+' '+ RTRIM(cNombre), " & _
                    " nNota=SUM(SE.nNota),nVigesi='',bGanador=CONVERT(int,SE.bGanador),cComentario=SE.cComentario FROM RHSeleccionEvaluacion SE INNER JOIN Postulante P " & _
                    " ON SE.cPersCod=P.cPostuCod " & _
                    " INNER JOIN RHSeleccionPostulante SEP ON SEP.cPostuCod = P.cPostuCod" & _
                    " AND SEP.cCodProceso = SE.cCodProceso " & _
                    " WHERE SE.nCodFase=" & pnCodFase & _
                    " AND SE.cCodProceso ='" & pcCodProceso & _
                    "' AND SE.cPersCodResp='" & pcPersCodResp & _
                    "' AND ISNULL(SEP.nEstado,0) NOT IN(4,5) GROUP BY P.cPostuCod,cApellidoPat,cApellidoMat,cApellidoMat,cNombre , SE.bGanador, SE.cComentario " & _
                    " ORDER BY SE.nNota DESC"
        End If
    Else
        If pcPersCodResp = "" Then
            sSql = "SELECT DISTINCT cPersCod=P.cPostuCod,cPersNombre=RTRIM(cApellidoPat)+' '+RTRIM(cApellidoMat)+' '+ RTRIM(cNombre), " & _
                   " nNota=SUM(SE.nNota),nVigesi='',bGanador=CONVERT(int,SE.bGanador),cComentario=SE.cComentario FROM RHSeleccionEvaluacion SE INNER JOIN Postulante P " & _
                " ON SE.cPersCod=P.cPostuCod " & _
                 " INNER JOIN RHSeleccionPostulante SEP ON SEP.cPostuCod = P.cPostuCod" & _
                " AND SEP.cCodProceso = SE.cCodProceso " & _
                 " WHERE SE.nCodFase=" & pnCodFase & _
                  " AND SE.cCodProceso ='" & pcCodProceso & _
                 "' AND ISNULL(SEP.nEstado,0) NOT IN (4,5) AND SE.bGanador=1 GROUP BY P.cPostuCod,cApellidoPat,cApellidoMat,cApellidoMat,cNombre , SE.bGanador, SE.cComentario " & _
                 " ORDER BY SE.nNota DESC"
            Else
                sSql = "SELECT DISTINCT cPersCod=P.cPostuCod,cPersNombre=RTRIM(cApellidoPat)+' '+RTRIM(cApellidoMat)+' '+ RTRIM(cNombre), " & _
                    " nNota=SUM(SE.nNota),nVigesi='',bGanador=CONVERT(int,SE.bGanador),cComentario=SE.cComentario FROM RHSeleccionEvaluacion SE INNER JOIN Postulante P " & _
                    " ON SE.cPersCod=P.cPostuCod " & _
                    " INNER JOIN RHSeleccionPostulante SEP ON SEP.cPostuCod = P.cPostuCod" & _
                    " AND SEP.cCodProceso = SE.cCodProceso " & _
                    " WHERE SE.nCodFase=" & pnCodFase & _
                    " AND SE.cCodProceso ='" & pcCodProceso & _
                    "' AND SE.cPersCodResp='" & pcPersCodResp & _
                    "' AND ISNULL(SEP.nEstado,0) NOT IN(4,5) AND SE.bGanador=1 GROUP BY P.cPostuCod,cApellidoPat,cApellidoMat,cApellidoMat,cNombre , SE.bGanador, SE.cComentario " & _
                    " ORDER BY SE.nNota DESC"
        End If
    
    End If
   
           
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Personas_ProcesoSeleccion_Evaluados = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Personas_ProcesoSeleccion_Evaluados:
    Err.Raise Err.Number, "Obtener_Personas_ProcesoSeleccion_Evaluados", Err.Description
End Function

Public Function Buscar_Fase_x_ProcesoSeleccion(ByVal pcCodProceso As String, _
                                                ByVal pnCodFase As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorBuscar_Fase_x_ProcesoSeleccion
    sSql = "SELECT SEF.nCodFase,SEF.cCodProceso FROM RHSeleccionFase SEF INNER JOIN RHSeleccionProceso SEP" & _
            " ON SEP.cCodProceso=SEF.cCodProceso " & _
            "WHERE SEF.cCodProceso ='" & pcCodProceso & "'" & " AND SEF.nCodFase=" & pnCodFase
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Buscar_Fase_x_ProcesoSeleccion = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorBuscar_Fase_x_ProcesoSeleccion:
    Err.Raise Err.Number, "Buscar_Fase_x_ProcesoSeleccion", Err.Description
End Function

Public Function Registrar_Aprobacion_Fase_ProcesoSeleccion(ByVal pnCodFase As Long, _
                                                        ByVal pcCodProceso As String, _
                                                        ByVal pcPersCod As String, _
                                                        ByVal pnAceptado As Long, _
                                                        ByVal pcComentario As String)
Dim sSql As String
Dim oConecta As DConecta
Dim oSelProceso As DSeleccionProceso
Dim rs As New ADODB.Recordset
Dim intNumAprobaciones As Integer
Dim intNumMiembrosComite As Integer
'Item a insertar en la Aprobacion de las Fases
Dim nLista As Long

    On Error GoTo ErrorRegistrar_Aprobacion_Fase_ProcesoSeleccion
        
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        'oConecta.ConexionActiva.Execute sSql
        Set oSelProceso = New DSeleccionProceso
             
        'sSQL = "SELECT Items=COUNT(*) FROM RHSeleccionFaseAprobacion " & _
                   " WHERE nCodFase=" & pnCodFase & _
                   " AND cCodProceso ='" & pcCodProceso & "'" & _
                   " AND cPersCod='" & pcPersCod & "'"
        
        'sSQL = "SELECT nLista FROM RHSeleccionEvaluacionHistorico WHERE nCodFase=" & pnCodFase & " AND cCodProceso='" & pcCodProceso & "' AND bEstado=1"
        'Set rs = oConecta.CargaRecordSet(sSQL)
        'nLista = CLng(rs("nLista"))
        nLista = 1  'La Lista va ser la unica
        
        'If rs.EOF Then  'Insertar solo en el caso de que no haya sido Aprobado/Rechazado antes por el mismo Miembro
        
        Call oSelProceso.InsertaRHSeleccionFaseAprobacion(pnCodFase, pcCodProceso, pcPersCod, nLista, pnAceptado, pcComentario)
        
        'End If
    '    If rs.EOF Then 'No existe la Aprobacion
    '        Call oSelProceso.InsertaRHSeleccionFaseAprobacion(pnCodFase, pcCodProceso, pcPersCod, nLista, pnAceptado, pcComentario)
    '    Else
    '       Call oSelProceso.ModificaRHSeleccionFaseAprobacion(pnCodFase, pcCodProceso, pcPersCod, pnAceptado, pcComentario)
    '    End If
        'If pnAceptado = 1 Then
        'Verificar si ya han aprobado la Fase del Proceso los otros Miembros del Comite
            
            sSql = "SELECT nCodFase, cCodProceso, cPersCod FROM RHSeleccionFaseAprobacion " & _
                  " WHERE nCodFase=" & pnCodFase & _
                    "AND cCodProceso ='" & pcCodProceso & "' AND nAceptado=1 AND nLista=" & nLista
        
            Set rs = oConecta.CargaRecordSet(sSql)
            intNumAprobaciones = rs.RecordCount
        
            Set rs = Obtener_Miembros_ComiteSelec(pcCodProceso, 0)
            intNumMiembrosComite = rs.RecordCount
        
            If intNumAprobaciones >= intNumMiembrosComite Then
                sSql = "UPDATE RHSeleccionFase SET bCerrada=1 WHERE cCodProceso='" & pcCodProceso & "' AND nCodFase=" & pnCodFase
            Else
                sSql = "UPDATE RHSeleccionFase SET bCerrada=0 WHERE cCodProceso='" & pcCodProceso & "' AND nCodFase=" & pnCodFase
            End If
            oConecta.ConexionActiva.Execute sSql
        'End If
    End If
    oConecta.CierraConexion

    Set oSelProceso = Nothing
    Set oConecta = Nothing
    Exit Function

ErrorRegistrar_Aprobacion_Fase_ProcesoSeleccion:
    Err.Raise Err.Number, "Registrar_Aprobacion_Fase_ProcesoSeleccion", Err.Description
End Function

Public Function Registrar_Aprobacion_Comite_Proceso(ByVal pcCodProceso As String, _
                                                        ByVal pcPersCod As String, _
                                                        ByVal pnTipoAprob As Long, _
                                                        ByVal pnAceptado As Long, _
                                                        ByVal pcComentario As String)
Dim sSql As String
Dim oConecta As DConecta
Dim oSelProceso As DSeleccionProceso
Dim rs As New ADODB.Recordset
Dim intNumAprobaciones As Integer
Dim intNumMiembrosComite As Integer

    On Error GoTo ErrorRegistrar_Aprobacion_Comite_Proceso
    
    sSql = "SELECT cCodProceso, cPersCod,nTipoAprob FROM RHSeleccionComiteAprobacion " & _
            " WHERE cCodProceso ='" & pcCodProceso & "'" & _
            " AND cPersCod='" & pcPersCod & "'" & _
            " AND nTipoAprob=" & pnTipoAprob
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        Set oSelProceso = New DSeleccionProceso
        If rs.EOF Then 'No existe la Aprobacion
            Call oSelProceso.InsertaRHSeleccionComiteAprobacion(pcCodProceso, pcPersCod, pnTipoAprob, pcComentario, pnAceptado)
        Else
            Call oSelProceso.ModificaRHSeleccionComiteAprobacion(pcCodProceso, pcPersCod, pnTipoAprob, pcComentario, pnAceptado)
        End If
        'If pnAceptado = 1 Then
        'Verificar si ya han aprobado la Fase del Proceso los otros Miembros del Comite
            sSql = "SELECT cCodProceso, cPersCod FROM RHSeleccionComiteAprobacion " & _
                  " WHERE cCodProceso ='" & pcCodProceso & _
                    "' AND nTipoAprob=" & pnTipoAprob & " AND nAceptado=1"
        
            Set rs = oConecta.CargaRecordSet(sSql)
            intNumAprobaciones = rs.RecordCount
        
            Set rs = Obtener_Miembros_ComiteSelec(pcCodProceso, 0)  'Solo Titulares
            intNumMiembrosComite = rs.RecordCount
        
            If intNumAprobaciones = intNumMiembrosComite Then
                If pnTipoAprob = 1 Then
                    sSql = "UPDATE RHSeleccionProceso SET nEstado=4 WHERE cCodProceso='" & pcCodProceso & "'"
                ElseIf pnTipoAprob = 2 Then
                    sSql = "UPDATE RHSeleccionProceso SET nEstado=7 WHERE cCodProceso='" & pcCodProceso & "'"
                Else    'Declarar Desierto
                    sSql = "UPDATE RHSeleccionProceso SET nEstado=6 WHERE cCodProceso='" & pcCodProceso & "'"
                End If
            Else
                If pnTipoAprob = 1 Then
                    sSql = "UPDATE RHSeleccionProceso SET nEstado=2 WHERE cCodProceso='" & pcCodProceso & "'"
                ElseIf pnTipoAprob = 2 Then
                    sSql = "UPDATE RHSeleccionProceso SET nEstado=3 WHERE cCodProceso='" & pcCodProceso & "'"
                Else    'En Propuesta Desierto
                    sSql = "UPDATE RHSeleccionProceso SET nEstado=5 WHERE cCodProceso='" & pcCodProceso & "'"
                End If
            End If
            oConecta.ConexionActiva.Execute sSql
        'End If
    End If
    oConecta.CierraConexion

    Set oSelProceso = Nothing
    Set oConecta = Nothing
    Exit Function

ErrorRegistrar_Aprobacion_Comite_Proceso:
    Err.Raise Err.Number, "Registrar_Aprobacion_Comite_Proceso", Err.Description
End Function

Public Function Obtener_Requerimientos_x_Estado(ByVal pnEstado As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Requerimientos_x_Estado
    
    sSql = " SELECT DISTINCT cCodCargoSol=CASE WHEN cCodCargoSol IS NULL THEN 'NUEVO' ELSE cCodCargoSol END,cCodRequerimiento," & _
               " cRHCargoDescripcion= CASE WHEN cRHCargoDescripcion IS NULL THEN rq.cCargoSolNuevo ELSE cRHCargoDescripcion END," & _
               " cCodAgeSol , cAgeDescripcion, nNumVacantes, dFechaIni, dFechaFin, rq.nEstado,C.cConsDescripcion,rq.cComentario" & _
               " from rhseleccionrequerimiento  rq" & _
               " inner join Constante C ON C.nConsValor=rq.nEstado AND C.nConsCod=9061" & _
               " inner join agencias ag on ag.cagecod=rq.ccodagesol  " & _
               " inner join Areas ar ON rq.cCodAreaSol=ar.cAreaCod left join rhcargostabla rt on rq.cCodCargoSol=rt.crhcargocod " & _
               " left join RHCargos car ON car.cRHCargoCod=rq.cCodCargoSol AND car.cRHCargoCod=rt.cRHCargoCod AND " & _
               " car.cRHAreaCod=ar.cAreaCod AND car.cRHAgenciaCod=ag.CAgeCod " & _
               " where rq.nEstado='" & pnEstado & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Requerimientos_x_Estado = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Requerimientos_x_Estado:
    Err.Raise Err.Number, "Obtener_Requerimientos_x_Estado", Err.Description
End Function

Public Function Obtener_SeleccionProceso_x_Estado(ByVal pnEstado As Long) As ADODB.Recordset
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorObtener_SeleccionProceso_x_Estado
    sSql = "SELECT cCodProceso,'Visto Bueno'='0', cDescProceso Propuesta, nTipoProceso,'Tipo Proceso'=TP.cConsDescripcion, nTipoContrato,'Tipo Contrato'=TC.cConsDescripcion, dFechaIni Inicio, dFechaFin Fin, nEstado,Estado=E.cConsDescripcion " & _
           " FROM RHSeleccionProceso SP INNER JOIN Constante TP ON SP.nTipoProceso=TP.nConsValor AND TP.nConsCod=6003" & _
           " INNER JOIN Constante TC ON SP.nTipoContrato= TC.nConsValor AND TC.nConsCod=6012 " & _
           "INNER JOIN Constante E ON SP.nEstado= E.nConsValor AND E.nConsCod=6004 "
    sSql = sSql & " WHERE nEstado=" & pnEstado
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_SeleccionProceso_x_Estado = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_SeleccionProceso_x_Estado:
    Err.Raise Err.Number, "Obtener_SeleccionProceso_x_Estado", Err.Description
End Function

Public Function Obtener_DatosProceso_x_Estado(Optional ByVal pnEstado As Long = 500) As ADODB.Recordset
                    
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorObtener_DatosProceso_x_Estado
    
    If pnEstado <> 500 Then
        sSql = "SELECT cCodProceso,cDescProceso,1,nTipoContrato,dFechaIni,dFechaFin,nEstado,cObservacion FROM RHSeleccionProceso "
        sSql = sSql & " WHERE nEstado=" & pnEstado
    Else
        sSql = "SELECT cCodProceso,cDescProceso,1,nTipoContrato,dFechaIni,dFechaFin,nEstado,cObservacion FROM RHSeleccionProceso "
    End If
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_DatosProceso_x_Estado = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_DatosProceso_x_Estado:
    Err.Raise Err.Number, "Obtener_DatosProceso_x_Estado", Err.Description
End Function


Public Function Obtener_CargosPosibles_NivelAprobacion(ByVal pcCodCargoNivel As String, _
                                                    ByVal pcCodCargoAprobac As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_CargosPosibles_NivelAprobacion
    sSql = "SELECT SNA.nCodNivelAprobac,CodCargo=SNA.cCodCargoSol,Cargo=Cat.cRHCargoDescripcion FROM RHSeleccionNivelAprobacion SNA INNER JOIN RHCargosTabla Cat ON SNA.cCodCargoSol=Cat.cRHCargoCod " & _
            " WHERE SNA.cCodCargoReq='" & pcCodCargoNivel & "' AND SNA.cCodCargoAprobac='" & pcCodCargoAprobac & "'"
            
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

Public Function Obtener_SeleccionNivelAprobacion_x_CodNivel(pnCodNivelAprobacion As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_SeleccionNivelAprobacion_x_CodNivel
    
    sSql = "SELECT nCodNivelAprobac, cCodCargoNivel,'Cargo Solicitante'=CaN.cRHCargoDescripcion , cCodCargoAprobac," & _
            " 'Cargo Aprobacion'=CaA.cRHCargoDescripcion ,Orden=nOrden, 'Mum. Aprobaciones'=nNumAprobac " & _
            " FROM RHSeleccionNivelAprobacion SNA " & _
            " INNER JOIN RHCargosTabla CaN ON SNA.cCodCargoNivel=CaN.cRHCargoCod INNER JOIN RHCargosTabla CaA " & _
            " ON SNA.cCodCargoAprobac=CaA.cRHCargoCod" & _
            " WHERE SNA.nCodNivelAprobac=" & pnCodNivelAprobacion
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_SeleccionNivelAprobacion_x_CodNivel = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_SeleccionNivelAprobacion_x_CodNivel:
    Err.Raise Err.Number, "Obtener_SeleccionNivelAprobacion_x_CodNivel", Err.Description
End Function

Public Function Obtener_SeleccionNivelAprobacion_Todos() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
'nCodNivelAprobac,
    On Error GoTo ErrorObtener_SeleccionNivelAprobacion_Todos
    sSql = "SELECT DISTINCT  cCodCargoReq,CargoSolicitante=CaN.cRHCargoDescripcion , cCodCargoAprobac," & _
            " CargoAprobacion=CaA.cRHCargoDescripcion ,Orden=nOrden, NumAprobaciones=nNumAprobac,bTodasAge=ISNULL(bTodasAge,0),bVerificaUltNivel=ISNULL(bVerificaUltNivel,0)" & _
            " FROM RHSeleccionNivelAprobacion SNA " & _
            " INNER JOIN RHCargosTabla CaN ON SNA.cCodCargoReq=CaN.cRHCargoCod INNER JOIN RHCargosTabla CaA " & _
            " ON SNA.cCodCargoAprobac=CaA.cRHCargoCod ORDER BY CaN.cRHCargoDescripcion"

    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_SeleccionNivelAprobacion_Todos = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_SeleccionNivelAprobacion_Todos:
    Err.Raise Err.Number, "Obtener_SeleccionNivelAprobacion_Todos", Err.Description
End Function

Public Function Obtener_Todos_Cargos() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Todos_Cargos
    
    sSql = "SELECT cRHCargoCod, cRHCargoDescripcion" & _
            " FROM RHCargosTabla "
            
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

Public Function Obtener_CargosSol_x_CargoReq(ByVal pcCodCargoReq As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_CargosSol_x_CargoReq
    
    sSql = "select r.cCodCargoSol,t.cRHCargoDescripcion,r.cCodCargoAprobac,bVerificaUltNivel=ISNULL(r.bVerificaUltNivel,0) from RHSeleccionNivelAprobacion r inner join RHCargosTabla t on r.cCodCargoSol = t.cRHCargoCod " & _
            " where r.cCodCargoReq='" & pcCodCargoReq & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_CargosSol_x_CargoReq = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_CargosSol_x_CargoReq:
    Err.Raise Err.Number, "Obtener_CargosSol_x_CargoReq", Err.Description
End Function

Public Function Obtener_CargosAprobac_x_CargoSol(ByVal pcCodCargoSol As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_CargosAprobac_x_CargoSol
    
    sSql = "select distinct cCodCargoAprobac,nOrden,bTodasAge From RHSeleccionNivelAprobacion " & _
            " where cCodCargosol='" & pcCodCargoSol & "'"
                
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_CargosAprobac_x_CargoSol = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_CargosAprobac_x_CargoSol:
    Err.Raise Err.Number, "Obtener_CargosAprobac_x_CargoSol", Err.Description
End Function

Public Function Obtener_NivelAprobado(ByVal pcCodCargoReq As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_NivelAprobado
    
    sSql = "select distinct cCodCargoAprobac,nOrden From RHSeleccionNivelAprobacion " & _
            " where cCodCargoReq='" & pcCodCargoReq & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_NivelAprobado = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_NivelAprobado:
    Err.Raise Err.Number, "Obtener_NivelAprobado", Err.Description
End Function

Public Function Obtener_ComiteSel_RepProceso(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_ComiteSel_RepProceso
    
    'sSql = "SELECT Miembro=P.CPersNombre, Cargo=cat.cRHCargoDescripcion + ' ('+ C.cConsDescripcion+ ')' " & _
            " FROM RHSeleccionProcesoComite SPC INNER JOIN Persona P ON SPC.cPersCod=P.cPersCod INNER JOIN Constante C " & _
            " ON C.nConsValor=SPC.nCargo AND C.nConsCod=9037 INNER JOIN rhcargos  car on  P.cperscod=car.cperscod and " & _
            " car.drhcargofecha=(select top 1 dRHCargoFecha from rhcargos where cperscod=car.cperscod order by dRHCargoFecha desc)" & _
            " INNER JOIN RHcargostabla cat on car.crhcargocod=cat.crhcargocod " & _
            "WHERE SPC.cCodProceso='" & pcCodProceso & "'"
    
    sSql = " SELECT Miembro=P.CPersNombre, Cargo=cat.cRHCargoDescripcion, TipoMiembro=C.cConsDescripcion + CASE WHEN ISNULL(cPersCodTitu,'')<>'' THEN ' - SUPLENTE' ELSE '' END  " & _
            " FROM RHSeleccionProcesoComite SPC INNER JOIN Persona P ON SPC.cPersCod=P.cPersCod INNER JOIN Constante C " & _
            " ON C.nConsValor=SPC.nCargo AND C.nConsCod=9037 INNER JOIN rhcargos  car on  P.cperscod=car.cperscod and " & _
            " car.drhcargofecha=(select top 1 dRHCargoFecha from rhcargos where cperscod=car.cperscod order by dRHCargoFecha desc)" & _
            " INNER JOIN RHcargostabla cat on car.crhcargocod=cat.crhcargocod " & _
            " WHERE SPC.cCodProceso='" & pcCodProceso & "' ORDER BY C.nConsValor"
            
            ' AND ISNULL(SPC.cPersCodTitu,'')=''"   'Solo los Titulares mostrarlos
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_ComiteSel_RepProceso = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_ComiteSel_RepProceso:
    Err.Raise Err.Number, "Obtener_ComiteSel_RepProceso", Err.Description
End Function

Public Function Obtener_FasesSel_RepProceso(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_FasesSel_RepProceso
    
    sSql = "SELECT  Etapa=SFS.cNomFase, Fechas=CONVERT(varchar(10),SF.dFechaIni,103) " & _
            " + CASE WHEN CONVERT(varchar(10),SF.dFechaIni,103)=CONVERT(varchar(10),SF.dFechaFin,103) THEN '' ELSE '-' + CONVERT(varchar(10),SF.dFechaFin,103) END," & _
            " Responsables=SFS.cResponsable FROM RHSeleccionFase SF INNER JOIN RHSeleccionFases SFS " & _
            " ON SF.nCodFase=SFS.nCodFase WHERE SF.cCodProceso='" & pcCodProceso & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_FasesSel_RepProceso = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_FasesSel_RepProceso:
    Err.Raise Err.Number, "Obtener_FasesSel_RepProceso", Err.Description
End Function

Public Function Obtener_CriteriosSel_RepProceso(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_CriteriosSel_RepProceso
    
    sSql = "SELECT  Etapa=SFS.cNomFase, PjeMAXIMO=SF.nNotaMax, Ponderacion=CONVERT(decimal(10,2), SF.nPonderacion), " & _
           " PjeMAXPonder=CONVERT(int,20*SF.nPonderacion) FROM RHSeleccionFase SF INNER JOIN RHSeleccionFases SFS " & _
           " ON SF.nCodFase=SFS.nCodFase WHERE SF.cCodProceso='" & pcCodProceso & "' AND SFS.bEjecuta=1"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_CriteriosSel_RepProceso = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_CriteriosSel_RepProceso:
    Err.Raise Err.Number, "Obtener_CriteriosSel_RepProceso", Err.Description
End Function

Public Function Obtener_numVacantes_CAP(ByVal pcRHCargoCod As String, pcAgeCod As String) As Long
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

On Error GoTo ErrorObtener_numVacantes_CAP
    
    'sSQL = "Select nvacantes=isnull(sum(nvacantes),0) From RHAsignacionPersonal  Where nvacantes<>0 and crhcargocod='" & pcRHCargoCod & "' and cAgeCod='" & pcAgeCod & "'"
    '"Select TOP 1 nvacantes=isnull(sum(nvacantes),0) From RHAsignacionPersonalVacantes Where nvacantes<>0 and crhcargocod='007006' and cRHAgenciaCod='18'GROUP BY nNroCAP ORDER BY nNroCAP DESC"
    
    sSql = "Select TOP 1 nvacantes=isnull(sum(nvacantes),0) From RHAsignacionPersonalVacantes  Where nvacantes<>0 and crhcargocod='" & pcRHCargoCod & "' and cRHAgenciaCod='" & pcAgeCod & "' GROUP BY nNroCAP ORDER BY nNroCAP DESC"
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_numVacantes_CAP = 0
        Else
            Obtener_numVacantes_CAP = rs("nvacantes")
        End If
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_numVacantes_CAP:
    Err.Raise Err.Number, "Obtener_numVacantes_CAP", Err.Description
End Function

'Public Sub Codifica_Proceso_Seleccion(ByVal pcCodProceso As String)
                    
'Dim sSql As String
'Dim oConecta As DConecta
'Dim rs As New ADODB.Recordset
    
'    On Error GoTo ErrorCodifica_Proceso_Seleccion
    
'    sSql = "INSERT INTO RHSeleccionRequisitos(nCodRequisito, cDescRequisito)"
'    sSql = sSql & " VALUES(" & pnCodRequisito & ",'" & pcDescRequisito & "')"
    
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    oConecta.ConexionActiva.Execute sSql
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Sub
    
'ErrorCodifica_Proceso_Seleccion:
'    Err.Raise Err.Number, "Codifica_Proceso_Seleccion", Err.Description
'End Sub

Public Function Obtener_NotasxPostulante_Proceso(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_NotasxPostulante_Proceso
    sSql = "sp_RHObtenerNotasxPostulante_Proceso '" & pcCodProceso & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_NotasxPostulante_Proceso = oConecta.Ejecutar(sSql)
        
    End If
    'oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_NotasxPostulante_Proceso:
    Err.Raise Err.Number, "Obtener_NotasxPostulante_Proceso", Err.Description
End Function

Public Function Obtener_AprobadosFase_SelProceso(ByVal pcCodProceso As String, _
                                                ByVal pnCodFase As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_AprobadosFase_SelProceso
    
    sSql = "SELECT cPersNombre=RTRIM(cApellidoPat)+' '+RTRIM(cApellidoMat)+' '+ RTRIM(cNombre) " & _
           "FROM RHSeleccionEvaluacion SE INNER JOIN Postulante P ON SE.cPersCod=P.cPostuCod " & _
            "WHERE SE.bGanador = 1 AND SE.nCodFase=" & pnCodFase & _
            " AND SE.cCodProceso ='" & pcCodProceso & "' ORDER BY P.cPersNombre"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_AprobadosFase_SelProceso = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_AprobadosFase_SelProceso:
    Err.Raise Err.Number, "Obtener_AprobadosFase_SelProceso", Err.Description
End Function

Public Function Obtener_DatosActa_Proceso(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_DatosActa_Proceso
    
    sSql = "SELECT TOP 1 " & _
           " Ciudad= CASE WHEN Ag.cAgeCod IN('01','02','03','05','07','09') THEN 'TRUJILLO' " & _
           " WHEN Ag.cAgeCod='10' THEN 'VIRU' " & _
           " WHEN Ag.cAgeCod='06' THEN 'CHEPEN' " & _
           " WHEN Ag.cAgeCod='08' THEN 'HUAMACHUCO'" & _
           " WHEN Ag.cAgeCod IN('12','15','16')THEN 'CHICLAYO'" & _
           " WHEN Ag.cAgeCod='13' THEN 'CAJAMARCA'  WHEN Ag.cAgeCod='14' THEN 'HUARAZ'" & _
           " WHEN Ag.cAgeCod='17' THEN 'JAEN'  WHEN Ag.cAgeCod='18' THEN 'LIMA'" & _
           " WHEN Ag.cAgeCod='19' THEN 'CHOCOPE'" & _
           " WHEN Ag.cAgeCod='20' THEN 'OTUZCO'" & _
           " WHEN Ag.cAgeCod='21' THEN 'TUMBES'" & _
           " END," & _
           " Cargo=  CAT.cRHCargoDescripcion, Fecha = CONVERT(VarChar(10), getdate(), 103)," & _
           " NroVacantes=SR.nNumVacantes, Agencia=AG.cAgeDescripcion, " & _
           " Nemonico=ISNULL(SP.cNemoProceso,''),dFechaIni=FAS.dFechaIni " & _
           " FROM RHSeleccionProceso SP INNER JOIN Constante TP ON TP.nConsValor=SP.nTipoProceso" & _
           " AND TP.nConsCod=6003" & _
           " INNER JOIN RHSeleccionRequerimientoProceso SRP ON SP.cCodProceso=SRP.cCodProceso" & _
           " INNER JOIN RHSeleccionRequerimiento SR ON SRP.cCodRequerimiento=SR.cCodRequerimiento" & _
           " INNER JOIN RHCargos CAR ON CAR.cRHCargoCod=SR.cCodCargoSol" & _
           " INNER JOIN RHCargosTabla CAT ON CAT.cRHCargoCod=CAR.cRHCargoCod" & _
           " INNER JOIN Agencias AG ON SR.cCodAgeSol=AG.cAgeCod" & _
           " INNER JOIN RHSeleccionFase FAS ON FAS.cCodProceso=SP.cCodProceso AND FAS.nCodFase=15" & _
           " WHERE SP.cCodProceso = '" & pcCodProceso & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_DatosActa_Proceso = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_DatosActa_Proceso:
    Err.Raise Err.Number, "Obtener_DatosActa_Proceso", Err.Description
End Function

Public Function Obtener_PostulantesProceso_x_Estado(ByVal pcCodProceso As String, _
                                                ByVal pnEstado As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_PostulantesProceso_x_Estado
    
    sSql = "SELECT cPersNombre=RTRIM(cApellidoPat)+' '+RTRIM(cApellidoMat)+' '+ RTRIM(cNombre)" & _
           " FROM RHSeleccionPostulante SP INNER JOIN Postulante P ON SP.cPostuCod=P.cPostuCod" & _
            " WHERE SP.cCodProceso='" & pcCodProceso & "' AND nEstado=" & pnEstado & " ORDER BY cPersNombre"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_PostulantesProceso_x_Estado = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_PostulantesProceso_x_Estado:
    Err.Raise Err.Number, "Obtener_PostulantesProceso_x_Estado", Err.Description
End Function

Public Function Obtener_CriteriosCurri_Proceso(ByVal pcCodProceso As String, _
                                               Optional ByVal pcFiltro As String = "") As ADODB.Recordset
                        
    Dim sSql As String
    Dim OCON As DConecta
    Set OCON = New DConecta
        
    If pcFiltro = "" Then
        'sSql = "SELECT cFiltro,cCondicion=cCondicon, Descripcion= ISNULL(cConsDescripcion,''), " & _
           " CE.Valor1,CE.Valor2,nPuntaje=ISNULL(nPuntaje,0) FROM RHSeleccionCriteriosEvaCurri CE LEFT JOIN Constante C ON " & _
           " CE.Valor1=C.nConsValor AND C.nConsCod='9051' AND CE.cFiltro='GRADO ACADEMICO' " & _
           " WHERE CE.cCodProceso= '" & pcCodProceso & "' ORDER BY cFiltro,nPuntaje DESC"
         sSql = "SELECT cFiltro,cCondicion=cCondicon, Descripcion= ISNULL(C1.cConsDescripcion,'')" & _
                " + CASE WHEN C2.cConsDescripcion IS NULL THEN '' ELSE '-'END + ISNULL(C2.cConsDescripcion,''), " & _
                " CE.Valor1 , CE.Valor2, nPuntaje = IsNull(nPuntaje, 0)FROM RHSeleccionCriteriosEvaCurri CE LEFT JOIN " & _
                " Constante C1 ON  CE.Valor1=C1.nConsValor AND C1.nConsCod='9051' AND CE.cFiltro='GRADO ACADEMICO' " & _
                " LEFT JOIN Constante C2 ON  CE.Valor2=C2.nConsValor AND C2.nConsCod='9052' AND CE.cFiltro='GRADO ACADEMICO' " & _
                " WHERE CE.cCodProceso= '" & pcCodProceso & "' ORDER BY cFiltro,nPuntaje DESC"
    Else
        'sSql = "SELECT cFiltro,cCondicion=cCondicon, Descripcion= ISNULL(cConsDescripcion,''), " & _
           " CE.Valor1,CE.Valor2,nPuntaje=ISNULL(nPuntaje,0) FROM RHSeleccionCriteriosEvaCurri CE LEFT JOIN Constante C ON " & _
           " CE.Valor1=C.nConsValor AND C.nConsCod='9051' AND CE.cFiltro='GRADO ACADEMICO' " & _
           " WHERE CE.cCodProceso= '" & pcCodProceso & "' AND cFiltro='" & pcFiltro & "' ORDER BY cFiltro,nPuntaje DESC"
         sSql = "SELECT cFiltro,cCondicion=cCondicon, Descripcion= ISNULL(C1.cConsDescripcion,'')" & _
                " + CASE WHEN C2.cConsDescripcion IS NULL THEN '' ELSE '-'END + ISNULL(C2.cConsDescripcion,''), " & _
                " CE.Valor1 , CE.Valor2, nPuntaje = IsNull(nPuntaje, 0)FROM RHSeleccionCriteriosEvaCurri CE LEFT JOIN " & _
                " Constante C1 ON  CE.Valor1=C1.nConsValor AND C1.nConsCod='9051' AND CE.cFiltro='GRADO ACADEMICO' " & _
                " LEFT JOIN Constante C2 ON  CE.Valor2=C2.nConsValor AND C2.nConsCod='9052' AND CE.cFiltro='GRADO ACADEMICO' " & _
                " WHERE CE.cCodProceso= '" & pcCodProceso & "' AND cFiltro='" & pcFiltro & "' ORDER BY cFiltro,nPuntaje DESC"
    End If
    If OCON.AbreConexion Then
        Set Obtener_CriteriosCurri_Proceso = OCON.CargaRecordSet(sSql)
        OCON.CierraConexion
    End If
    Set OCON = Nothing
    
End Function

Public Function Obtener_Grado_Postulante(ByVal pcCodPostulante As String, ByVal pnNivel As Long) As Long

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorObtener_Grado_Postulante
    
    sSql = "SELECT TOP 1 nNivelAlcanzado FROM RHSeleccionCurriculumDPersonales " & _
           " WHERE cPostuCod='" & pcCodPostulante & "' AND nNivelAlcanzado=" & pnNivel & " ORDER BY nNivelAlcanzado"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_Grado_Postulante = 0
        Else
            Obtener_Grado_Postulante = rs("nNivelAlcanzado")
        End If
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Grado_Postulante:
    Err.Raise Err.Number, "Obtener_Grado_Postulante", Err.Description
End Function

Public Function Obtener_Nivel_Postulante(ByVal pcCodPostulante As String, ByVal pnGrado As Long) As Long

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorObtener_Nivel_Postulante
    
    sSql = "SELECT TOP 1 nGrado=ISNULL(nGrado,0) FROM RHSeleccionCurriculumDPersonales " & _
            " WHERE cPostuCod='" & pcCodPostulante & "' AND ISNULL(nGrado,0)=" & pnGrado & " ORDER BY ISNULL(nGrado,0)"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_Nivel_Postulante = 0
        Else
            Obtener_Nivel_Postulante = rs("nGrado")
        End If
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Nivel_Postulante:
    Err.Raise Err.Number, "Obtener_Nivel_Postulante", Err.Description
End Function


Public Function Obtener_Experiencia_Postulante(ByVal pcCodPostulante As String) As Double

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
Dim dblExperiencia As Double

    On Error GoTo ErrorObtener_Experiencia_Postulante
    
    sSql = "SELECT DiasExper=SUM(DATEDIFF(DAY,dFechaIni,dFechaFin)) From RHSeleccionCurriculumDLaborales " & _
           " WHERE cPostuCod='" & pcCodPostulante & "' GROUP BY cPostuCod"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_Experiencia_Postulante = 0
        Else
            dblExperiencia = rs("DiasExper")
            Obtener_Experiencia_Postulante = dblExperiencia / 365   'Dias del Año ?
        End If
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Experiencia_Postulante:
    Err.Raise Err.Number, "Obtener_Experiencia_Postulante", Err.Description
End Function

Public Function Obtener_CodigoCompleto_Proceso(pcCodProceso As String) As String

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrorObtener_CodigoCompleto_Proceso
    sSql = "SELECT TOP 1 " & _
        " Codigo=CASE WHEN TP.nConsValor=0 THEN 'CP' WHEN TP.nConsValor=1 THEN 'CI' " & _
        " WHEN TP.nConsValor=2 THEN 'PRACT' END +'-'+ " & _
        " RIGHT(RTRIM(SP.cCodProceso),3) +'-' +  " & _
        "LEFT(cRHCargoDescripcion,3)+ REPLACE (   CASE WHEN LEFT(cRHCargoDescripcion,3)= " & _
        " SUBSTRING( REPLACE( REPLACE( REPLACE(cRHCargoDescripcion,'SUB GERENTE',' '),' DE','') ,' ','$'), " & _
        " PATINDEX('%$%',REPLACE( REPLACE(REPLACE(cRHCargoDescripcion,'SUB GERENTE',' '),' DE','') ,' ','$'))   +1 " & _
        ",3 ) THEN '' Else SUBSTRING( REPLACE( REPLACE(REPLACE(cRHCargoDescripcion,'SUB GERENTE',' '),' DE','') ,' ','$'), " & _
        " PATINDEX('%$%', REPLACE( REPLACE(REPLACE(cRHCargoDescripcion,'SUB GERENTE',' '),' DE','') ,' ','$'))   +1 " & _
        " ,3 ) End ,'$','')  +'-'+ 'RRHH'+'-'+ CASE WHEN Ag.cAgeCod IN('01','02','03','05','07','09') THEN 'TRU' " & _
        " WHEN Ag.cAgeCod='10' THEN 'VIR' WHEN Ag.cAgeCod='06' THEN 'CHE'  WHEN Ag.cAgeCod='08' THEN 'HUA' WHEN Ag.cAgeCod IN('12','15','16')THEN 'CHI' " & _
        " WHEN Ag.cAgeCod='13' THEN 'CAJ' WHEN Ag.cAgeCod='14' THEN 'HUA'  WHEN Ag.cAgeCod='17' THEN 'JAE' WHEN Ag.cAgeCod='18' THEN 'LIM' " & _
        " WHEN Ag.cAgeCod='19' THEN 'CHO' WHEN Ag.cAgeCod='20' THEN 'OTU'  WHEN Ag.cAgeCod='21' THEN 'TUM' " & _
        " END + '-'+ CONVERT(varchar(2),MONTH(SP.dFechaIni))  +'/'+CONVERT(varchar(4),YEAR(SP.dFechaIni))  FROM RHSeleccionProceso SP INNER JOIN Constante TP ON TP.nConsValor=SP.nTipoProceso " & _
        " AND TP.nConsCod=6003 " & _
        " INNER JOIN RHSeleccionRequerimientoProceso SRP ON SP.cCodProceso=SRP.cCodProceso " & _
        " INNER JOIN RHSeleccionRequerimiento SR ON SRP.cCodRequerimiento=SR.cCodRequerimiento " & _
        " INNER JOIN RHCargos CAR ON CAR.cRHCargoCod=SR.cCodCargoSol " & _
        " INNER JOIN RHCargosTabla CAT ON CAT.cRHCargoCod=CAR.cRHCargoCod " & _
        " INNER JOIN Agencias AG ON SR.cCodAgeSol=AG.cAgeCod" & _
        " WHERE SP.cCodProceso='" & pcCodProceso & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_CodigoCompleto_Proceso = ""
        Else
            Obtener_CodigoCompleto_Proceso = CStr(rs("Codigo"))
        End If
    End If
    
    sSql = " UPDATE RHSeleccionProceso SET cNemoProceso='" & Obtener_CodigoCompleto_Proceso & "' WHERE cCodProceso='" & pcCodProceso & "'"
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    
    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_CodigoCompleto_Proceso:
    Err.Raise Err.Number, "Obtener_CodigoCompleto_Proceso", Err.Description
End Function

Public Function Obtener_CuadroConsolidado(ByVal pcCodProceso As String, _
                                        ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim lr As New ADODB.Recordset
    On Error GoTo ErrorObtener_CuadroConsolidado
    sSql = "sp_RHObtener_CuadroConsolidado '" & pcCodProceso & "'," & pnTipo
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        Set Obtener_CuadroConsolidado = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtener_CuadroConsolidado:
    Err.Raise Err.Number, "Obtener_CuadroConsolidado", Err.Description
End Function

Public Function Obtener_PromedioNotas_Postulante(ByVal pcCodProceso As String, _
                                                 ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_PromedioNotas_Postulante
    sSql = "sp_RHObtener_PromedioNotas_Postulante '" & pcCodProceso & "'," & pnTipo
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_PromedioNotas_Postulante = oConecta.Ejecutar(sSql)
        
    End If
    'oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_PromedioNotas_Postulante:
    Err.Raise Err.Number, "Obtener_PromedioNotas_Postulante", Err.Description
End Function

Public Function Obtener_Siguiente_CodFase() As Long

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrorObtener_Siguiente_CodFase
    sSql = "SELECT Siguiente=MAX(nCodFase)+1 FROM RHSeleccionFases "
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_Siguiente_CodFase = 1
        Else
            Obtener_Siguiente_CodFase = rs("Siguiente") '+ 1
        End If
    End If
    
    oConecta.CierraConexion
    
    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Siguiente_CodFase:
    Err.Raise Err.Number, "Obtener_Siguiente_CodFase", Err.Description

End Function

Public Function Busca_Fases_AsignadosProceso(ByVal pnCodFase As Long) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorBusca_Fases_AsignadosProceso
    sSql = "SELECT nCodFase FROM RHSeleccionFase " & _
            " WHERE nCodFase=" & pnCodFase
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Busca_Fases_AsignadosProceso = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorBusca_Fases_AsignadosProceso:
    Err.Raise Err.Number, "Busca_Fases_AsignadosProceso", Err.Description
End Function

Public Function Registrar_Aprobacion_Gerencia_Proceso(ByVal pcCodProceso As String, _
                                                        ByVal pcPersCod As String, _
                                                        ByVal pnAceptado As Long, _
                                                        ByVal pcComentario As String)
Dim sSql As String
Dim oConecta As DConecta
Dim oSelProceso As DSeleccionProceso
Dim rs As New ADODB.Recordset
Dim intNumAprobaciones As Integer
Dim intNumMiembrosComite As Integer

    On Error GoTo ErrorRegistrar_Aprobacion_Gerencia_Proceso
    
    sSql = "SELECT cCodProceso, cPersCod FROM RHSeleccionGerenciaAprobacion " & _
            " WHERE cCodProceso ='" & pcCodProceso & "'" & _
            " AND cPersCod='" & pcPersCod & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        Set oSelProceso = New DSeleccionProceso
        If rs.EOF Then 'No existe la Aprobacion
            Call oSelProceso.InsertaRHSeleccionGerenciaAprobacion(pcCodProceso, pcPersCod, pcComentario, pnAceptado)
        Else
            Call oSelProceso.ModificaRHSeleccionGerenciaAprobacion(pcCodProceso, pcPersCod, pcComentario, pnAceptado)
        End If
        'If pnAceptado = 1 Then
        '    sSql = "UPDATE RHSeleccionProceso SET nEstado=2 WHERE cCodProceso='" & pcCodProceso & "'"   'Aprobado en Propuesta por Gerencia
        '    oConecta.ConexionActiva.Execute sSql
        'End If
        'Al Rechazar liberar el Requerimiento
        If pnAceptado = 0 Then
            'Debemos declarar el Proceso Cerrado
            sSql = "UPDATE RHSeleccionProceso SET nEstado=2 WHERE cCodProceso='" & pcCodProceso & "'"   'Rechazado en Propuesta por Gerencia
            oConecta.ConexionActiva.Execute sSql
            'Al Cerrar el Proceso debemos liberar el Requerimiento
            sSql = "UPDATE RHSeleccionRequerimiento SET nEstado=4 WHERE cCodRequerimiento IN (SELECT cCodRequerimiento FROM RHSeleccionRequerimientoProceso WHERE cCodProceso='" & pcCodProceso & "')"
            oConecta.ConexionActiva.Execute sSql
        End If
    End If
    oConecta.CierraConexion

    Set oSelProceso = Nothing
    Set oConecta = Nothing
    Exit Function

ErrorRegistrar_Aprobacion_Gerencia_Proceso:
    Err.Raise Err.Number, "Registrar_Aprobacion_Gerencia_Proceso", Err.Description
End Function

Public Function Obtener_Detalles_MovNro(ByVal pnMovNro As Long) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Detalles_MovNro
    sSql = "select  md.nMovNro,d.cDocDesc,md.cDocNro,Fecha=dbo.GetFechaMov(left(m.cMovNro,8),103), " & _
            " m.cMovDesc , c.nMovImporte from Mov m inner join MovDoc md on m.nMovNro = md.nMovNro " & _
            " inner join MovCta c on m.nMovNro = c.nMovNro inner join Documento d on d.nDocTpo = md.nDocTpo " & _
            " Where m.nMovNro =" & pnMovNro & " And c.nMovImporte > 0"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Detalles_MovNro = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Detalles_MovNro:
    Err.Raise Err.Number, "Obtener_Detalles_MovNro", Err.Description
End Function

Public Function Obtener_Movimientos_Costos(Optional ByVal pdFecha As String = "") As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Movimientos_Costos
    sSql = "select md.nMovNro,Descripcion=d.cDocDesc+' '+md.cDocNro+' '+dbo.GetFechaMov(left(m.cMovNro,8),103)+' '+ " & _
            " convert(varchar(10),c.nMovImporte) from Mov m inner join MovDoc md on m.nMovNro = md.nMovNro inner join MovCta c " & _
            " on m.nMovNro = c.nMovNro inner join Documento d on d.nDocTpo = md.nDocTpo" & _
            " Where d.nDocTpo in(1,2,3) and c.nMovImporte > 0 and c.cCtaContCod LIKE '4%' " & IIf(pdFecha = "", "", "AND dbo.GetFechaMov(left(m.cMovNro,8),103)='" & pdFecha & "'")
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Movimientos_Costos = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Movimientos_Costos:
    Err.Raise Err.Number, "Obtener_Movimientos_Costos", Err.Description
End Function

Public Function Obtener_Aprobaciones_x_Requerimiento(ByVal pcCodRequerimiento As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Aprobaciones_x_Requerimiento
    
    sSql = " Select  ccodrequerimiento,P.cPersNombre,CA.crhCargoDescripcion,nOrden,Fecha=ISNULL(CONVERT(varchar(10),rqa.dFecha,103),''),cComentario=ISNULL(cComentario,'') " & _
        " From RHSeleccionrequerimientoaprobaciones rqa inner join RhCargosTabla CA " & _
        " on CA.crhCargoCod=rqa.ccodCargoAprobac inner join Persona P on P.cPersCod= rqa.cPersCodAprobac " & _
        " where cCodRequerimiento='" & pcCodRequerimiento & "' AND bAprobado=1 ORDER BY nOrden "
                
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Aprobaciones_x_Requerimiento = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Aprobaciones_x_Requerimiento:
    Err.Raise Err.Number, "Obtener_Aprobaciones_x_Requerimiento", Err.Description
End Function

Public Function Obtener_Costos_x_FaseSelec(ByVal pnCodFase As Long, _
                                       ByVal pcCodProceso As String, _
                                       Optional ByVal pnItem As Long = 0) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Costos_x_FaseSelec
    
    If pnItem = 0 Then
        sSql = "SELECT nCodFase, cCodProceso, dFecha, cGlosa,nMonto FROM RHSeleccionFaseCostosDet" & _
            " where ncodfase=" & pnCodFase & " AND cCodProceso='" & pcCodProceso & "' ORDER by nItem "
    Else
        sSql = "SELECT nCodFase, cCodProceso, dFecha, cGlosa, nMonto FROM RHSeleccionFaseCostosDet" & _
            " where ncodfase=" & pnCodFase & " AND cCodProceso='" & pcCodProceso & "' AND nItem=" & pnItem & " ORDER by nItem "
    End If
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Costos_x_FaseSelec = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Costos_x_FaseSelec:
    Err.Raise Err.Number, "Obtener_Costos_x_FaseSelec", Err.Description
End Function

Public Function Vetar_Postulante_ProcesoSelec(ByVal pcCodProceso As String, _
                                                ByVal pcPostuCod As String, _
                                                ByVal pnCodFase As Long, _
                                                ByVal pcMotivo As String)
Dim sSql As String
Dim oConecta As DConecta
Dim oSelProceso As DSeleccionProceso
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorVetar_Postulante_ProcesoSelec
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        sSql = "INSERT INTO RHPostulanteVetado(cPostuCod,cCodProceso,nCodFase,cMotivo)" & _
               " VALUES('" & pcPostuCod & "','" & pcCodProceso & "'," & pnCodFase & ",'" & pcMotivo & "')"
        
        oConecta.ConexionActiva.Execute sSql
        
        sSql = "UPDATE RHSeleccionPostulante SET nEstado=4 WHERE cCodProceso='" & pcCodProceso & "' AND cPostuCod='" & pcPostuCod & "'"
        oConecta.ConexionActiva.Execute sSql
        
    End If
    oConecta.CierraConexion

    Set oSelProceso = Nothing
    Set oConecta = Nothing
    Exit Function

ErrorVetar_Postulante_ProcesoSelec:
    Err.Raise Err.Number, "Vetar_Postulante_ProcesoSelec", Err.Description
End Function

Public Function Obtener_NotaMax_FaseSeleccion(ByVal pnCodFase As Long, _
                                               ByVal pcCodProceso As String) As Double
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorObtener_NotaMax_FaseSeleccion
    sSql = "SELECT nNotaMax FROM RHSeleccionFase WHERE nCodFase=" & pnCodFase & " AND cCodProceso='" & pcCodProceso & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        
        If Not rs.EOF Then
            Obtener_NotaMax_FaseSeleccion = rs("nNotaMax")
        Else
            Obtener_NotaMax_FaseSeleccion = 0
        End If
    End If
    oConecta.CierraConexion

    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function

ErrorObtener_NotaMax_FaseSeleccion:
    Err.Raise Err.Number, "Obtener_NotaMax_FaseSeleccion", Err.Description
End Function

Public Function Obtener_NotaAprobatoria_FaseSeleccion(ByVal pnCodFase As Long, _
                                               ByVal pcCodProceso As String) As Double
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorObtener_NotaAprobatoria_FaseSeleccion
    sSql = "SELECT nNotaAprobatoria FROM RHSeleccionFase WHERE nCodFase=" & pnCodFase & " AND cCodProceso='" & pcCodProceso & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        
        If Not rs.EOF Then
            Obtener_NotaAprobatoria_FaseSeleccion = rs("nNotaAprobatoria")
        Else
            Obtener_NotaAprobatoria_FaseSeleccion = 0
        End If
    End If
    oConecta.CierraConexion

    Set rs = Nothing
    Set oConecta = Nothing
    Exit Function

ErrorObtener_NotaAprobatoria_FaseSeleccion:
    Err.Raise Err.Number, "Obtener_NotaAprobatoria_FaseSeleccion", Err.Description
End Function

Public Function Obtener_Aprobaciones_ComiteSelec(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Aprobaciones_ComiteSelec
    sSql = " SELECT SCA.cCodProceso,SCA.cPersCod,P.cPersNombre,SCA.nAceptado, Estado=C.cConsDescripcion,SCA.cComentario,TipoAprob = CASE WHEN SCA.nTipoAprob=1 THEN 'DE PROPUESTA' WHEN SCA.nTipoAprob=2 THEN 'DE ACTA' ELSE 'DE PROCESO DESIERTO' END " & _
           " FROM RHSeleccionComiteAprobacion SCA INNER JOIN Persona P ON SCA.cPersCod = P.cPersCod INNER JOIN Constante C " & _
           " ON SCA.nAceptado = C.nConsValor AND C.nConsCod='9069' " & _
           " WHERE SCA.cCodProceso='" & pcCodProceso & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Aprobaciones_ComiteSelec = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Aprobaciones_ComiteSelec:
    Err.Raise Err.Number, "Obtener_Aprobaciones_ComiteSelec", Err.Description
End Function

Public Function Obtener_Postulantes_ProcesoSel(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Postulantes_ProcesoSel
    sSql = "SELECT cPersCod=P.cPostuCod,cPersNombre=RTRIM(cApellidoPat)+' '+RTRIM(cApellidoMat)+' '+ RTRIM(cNombre),Estado=ISNULL(C.cConsDescripcion,'POSTULANTE') " & _
           " FROM Postulante P INNER JOIN RHSeleccionPostulante SELP ON P.cPostuCod=SELP.cPostuCod " & _
           " INNER JOIN RHSeleccionProceso SP ON SELP.cCodProceso = SP.cCodProceso " & _
           " LEFT JOIN Constante C ON C.nConsValor= SELP.nEstado AND C.nConsCod='9039' " & _
            " WHERE SP.cCodProceso ='" & pcCodProceso & "' ORDER BY P.cPersNombre"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Postulantes_ProcesoSel = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Postulantes_ProcesoSel:
    Err.Raise Err.Number, "Obtener_Postulantes_ProcesoSel", Err.Description
End Function

Public Function Cambiar_Estado_Postulante_ProcSel(ByVal pcCodProceso As String, _
                                                ByVal pcPostuCod As String, _
                                                ByVal pnEstado As Long, _
                                                Optional ByVal pcMotivo As String = "")
Dim sSql As String
Dim oConecta As DConecta
Dim oSelProceso As DSeleccionProceso
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorCambiar_Estado_Postulante_ProcSel
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        
        sSql = "UPDATE RHSeleccionPostulante SET nEstado=" & pnEstado & ",cObservacion='" & pcMotivo & "' WHERE cCodProceso='" & pcCodProceso & "' AND cPostuCod='" & pcPostuCod & "'"
        oConecta.ConexionActiva.Execute sSql
    End If
    oConecta.CierraConexion

    Set oSelProceso = Nothing
    Set oConecta = Nothing
    Exit Function

ErrorCambiar_Estado_Postulante_ProcSel:
    Err.Raise Err.Number, "Cambiar_Estado_Postulante_ProcSel", Err.Description
End Function

Public Function Obtener_Datos_ProcesoSel_para_Contrato(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_ProcesoSel_para_Contrato
    
    sSql = "SELECT TOP 1 Cargo=  CAT.cRHCargoDescripcion, CodCargo= CAT.cRHCargoCod, " & _
            " CodAgencia= AG.cAgeCod, CodArea= SR.cCodArea, TipoContrato=SP.nTipoContrato," & _
            " Fecha = CONVERT(VarChar(10), getdate(), 103), NroVacantes=SR.nNumVacantes," & _
            " Agencia = AG.cAgeDescripcion " & _
            " FROM RHSeleccionProceso SP INNER JOIN Constante TP ON TP.nConsValor=SP.nTipoProceso " & _
            " AND TP.nConsCod=6003 " & _
            " INNER JOIN RHSeleccionRequerimientoProceso SRP ON SP.cCodProceso=SRP.cCodProceso " & _
            " INNER JOIN RHSeleccionRequerimiento SR ON SRP.cCodRequerimiento=SR.cCodRequerimiento " & _
            " INNER JOIN RHCargos CAR ON CAR.cRHCargoCod=SR.cCodCargoSol " & _
            " INNER JOIN RHCargosTabla CAT ON CAT.cRHCargoCod=CAR.cRHCargoCod " & _
            " INNER JOIN Agencias AG ON SR.cCodAgeSol=AG.cAgeCod " & _
           " WHERE SP.cCodProceso = '" & pcCodProceso & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_ProcesoSel_para_Contrato = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Datos_ProcesoSel_para_Contrato:
    Err.Raise Err.Number, "Obtener_Datos_ProcesoSel_para_Contrato", Err.Description
End Function

Public Function Obtener_PromediosPostulante_Contrato(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_PromediosPostulante_Contrato
    sSql = "sp_RHObtener_PromediosPostulante_Contrato '" & pcCodProceso & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_PromediosPostulante_Contrato = oConecta.Ejecutar(sSql)
        
    End If
    'oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function
ErrorObtener_PromediosPostulante_Contrato:
    Err.Raise Err.Number, "Obtener_PromediosPostulante_Contrato", Err.Description
End Function

Public Function Obtener_Requerimientos_ProcSel(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Requerimientos_ProcSel
    sSql = " select Codigo=r.ccodrequerimiento, " & _
            "Descripcion=rt.crhcargodescripcion +' '+  careadescripcion+' '+ ag.cAgeDescripcion,1" & _
            " from rhseleccionrequerimiento r " & _
            " inner join rhseleccionrequerimientoproceso rp on r.ccodrequerimiento=rp.ccodrequerimiento" & _
            " inner join rhcargostabla rt on r.ccodcargosol=rt.crhcargocod" & _
            " inner join areas a on a.careacod=r.cCodAreaSol" & _
            " inner join agencias ag on ag.cagecod=r.ccodagesol where r.nestado=6 and ccodproceso='" & pcCodProceso & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Requerimientos_ProcSel = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Requerimientos_ProcSel:
    Err.Raise Err.Number, "Obtener_Requerimientos_ProcSel", Err.Description
End Function

Public Function Cubrir_Requerimiento_ProcSel(ByVal pcCodProceso As String, _
                                                ByVal pcPostuCod As String, _
                                                ByVal pcCodRequerimiento As String, _
                                                ByVal pcCodPersona As String, _
                                                Optional ByVal pnTipoContrato As Integer = -1)
Dim sSql As String
Dim oConecta As DConecta
Dim oSelProceso As DSeleccionProceso
Dim rs As New ADODB.Recordset

Dim cCodAgencia As String
Dim cCodCargo As String

    On Error GoTo ErrorCubrir_Requerimiento_ProcSel
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        
        sSql = "UPDATE RHSeleccionPostulante SET cCodRequerimiento='" & pcCodRequerimiento & "',cPersCod='" & pcCodPersona & "' " & _
                " WHERE cCodProceso='" & pcCodProceso & "' AND cPostuCod='" & pcPostuCod & "'"
        
        oConecta.ConexionActiva.Execute sSql
    
        'If pnTipoContrato <> -1 Then
            'sSQL = "SELECT cCodCargoSol,cCodAgeSol From RHSeleccionRequerimiento WHERE cCodRequerimiento='" & pcCodRequerimiento & "'"
        
            'Set rs = oConecta.CargaRecordSet(sSQL)
        
            'If rs.EOF Then Exit Function
        
            'cCodAgencia = rs("cCodAgeSol")
            'cCodCargo = rs("cCodCargoSol")
                    'nTipoProceso = rs("nTipoProceso")
        
            'sSQL = "UPDATE RHAsignacionPersonal SET nVacantes= nVacantes - 1 ," & _
                 IIf(pnTipoContrato = 0, "nPersPlazoIndeterminado= nPersPlazoIndeterminado + 1", "nPersPlazoFijo = nPersPlazoFijo +1") & _
                " WHERE cRHCargoCod='" & cCodCargo & "' AND cAgeCod='" & cCodAgencia & "'"
        
            'oConecta.ConexionActiva.Execute sSQL
        
        'End If
                
    End If
    
    
    oConecta.CierraConexion

    Set rs = Nothing
    Set oSelProceso = Nothing
    Set oConecta = Nothing
    Exit Function

ErrorCubrir_Requerimiento_ProcSel:
    Err.Raise Err.Number, "Cubrir_Requerimiento_ProcSel", Err.Description
End Function

Public Function Obtener_Datos_Postulante_Contrato(ByVal pcCodPostulante As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_Postulante_Contrato
    sSql = " SELECT cPostuCod, cPersNombre=RTRIM(cApellidoPat)+'/'+RTRIM(cApellidoMat)+','+ RTRIM(cNombre)," & _
           " cFechaNac , cDNI, cZona, cDireccion, cTelefono, cOTelefono=ISNULL(cOTelefono,''), cEmail, cSexo=CASE WHEN cSexo='1' THEN 'F' ELSE 'M' END, cEstadoCivil, nHijos " & _
           " From Postulante WHERE cPostuCod='" & pcCodPostulante & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_Postulante_Contrato = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function
    
ErrorObtener_Datos_Postulante_Contrato:
    Err.Raise Err.Number, "Obtener_Datos_Postulante_Contrato", Err.Description
End Function

Public Function Obtener_AprobFases_ComiteSelec(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_AprobFases_ComiteSelec
    sSql = " SELECT  SFA.cCodProceso,SFA.cPersCod,P.cPersNombre,SFA.nAceptado, Estado=C.cConsDescripcion,SFA.cComentario," & _
           " SF.cNomFase,dFecha=ISNULL(CONVERT(varchar(10),SFA.dFecha,103),''),SFA.nLista FROM RHSeleccionFaseAprobacion  SFA INNER JOIN Persona P ON SFA.cPersCod = P.cPersCod INNER JOIN Constante C " & _
           " ON SFA.nAceptado = C.nConsValor AND C.nConsCod='9069' INNER JOIN RHSeleccionFases SF ON SF.nCodFase=SFA.nCodFase " & _
           " WHERE SFA.cCodProceso='" & pcCodProceso & "' ORDER BY SFA.nCodFase,SFA.nLista,SFA.cPersCod"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_AprobFases_ComiteSelec = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_AprobFases_ComiteSelec:
    Err.Raise Err.Number, "Obtener_AprobFases_ComiteSelec", Err.Description
End Function

Public Function Cambiar_Estado_ProcesoSelec(ByVal pcCodProceso As String, _
                                            ByVal pnEstado As Long, _
                                            Optional ByVal pcObservacion As String = "")
Dim sSql As String
Dim oConecta As DConecta
Dim oSelProceso As DSeleccionProceso
Dim rs As New ADODB.Recordset

    On Error GoTo ErrorCambiar_Estado_ProcesoSelec
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
        
        sSql = "UPDATE RHSeleccionProceso SET nEstado=" & pnEstado & ",cObservacion='" & pcObservacion & "' WHERE cCodProceso='" & pcCodProceso & "'"
        oConecta.ConexionActiva.Execute sSql
    End If
    oConecta.CierraConexion

    Set oSelProceso = Nothing
    Set oConecta = Nothing
    Exit Function

ErrorCambiar_Estado_ProcesoSelec:
    Err.Raise Err.Number, "Cambiar_Estado_ProcesoSelec", Err.Description
End Function

Public Function Obtener_Fases_NoCerradas_Previas(ByVal pcCodProceso As String, ByVal pnCodFase As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Fases_NoCerradas_Previas
    sSql = "SELECT SFA1.nCodFase FROM RHSeleccionFase SFA INNER JOIN RHSeleccionFase SFA1 ON  SFA1.cCodProceso=SFA.cCodProceso " & _
           " AND SFA1.nOrden < SFA.nOrden INNER JOIN RHSeleccionFases SFS ON SFA1.nCodFase=SFS.nCodFase AND SFS.bEjecuta=1 " & _
           " AND ISNULL(SFA1.bCerrada,0)=0 " & _
           " WHERE SFA.cCodProceso='" & pcCodProceso & "' AND SFA.nCodFase=" & pnCodFase
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Fases_NoCerradas_Previas = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Fases_NoCerradas_Previas:
    Err.Raise Err.Number, "Obtener_Fases_NoCerradas_Previas", Err.Description
End Function

Public Function Obtener_Miembros_ComiteSelec(ByVal pcCodProceso As String, _
                                            Optional ByVal pnTipo As Integer = 0) As ADODB.Recordset
                                            '0: Solo Titulares, 1: Solo Suplentes y 2: Todos
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Miembros_ComiteSelec
    
    Select Case pnTipo
    
    Case 0  'Solo Titulares
            sSql = "SELECT P.cPersCod,Miembro=P.CPersNombre,1,SPC.nCargo,Cargo=C.cConsDescripcion " & _
            "FROM Persona P INNER JOIN RHSeleccionProcesoComite SPC " & _
            "ON P.cPersCod = SPC.cPersCod INNER JOIN RHSeleccionProceso SP " & _
            "ON SPC.cCodProceso =SP.cCodProceso     INNER JOIN Constante C " & _
            "ON C.nConsValor= SPC.nCargo AND C.nConsCod='9037'" & _
            " WHERE SP.cCodProceso='" & pcCodProceso & "' AND ISNULL(cPersCodTitu,'') =''"
    Case 1  'Solo Suplentes
            sSql = "SELECT P.cPersCod,Miembro=P.CPersNombre,1,SPC.nCargo,Cargo=C.cConsDescripcion " & _
            "FROM Persona P INNER JOIN RHSeleccionProcesoComite SPC " & _
            "ON P.cPersCod = SPC.cPersCod INNER JOIN RHSeleccionProceso SP " & _
            "ON SPC.cCodProceso =SP.cCodProceso     INNER JOIN Constante C " & _
            "ON C.nConsValor= SPC.nCargo AND C.nConsCod='9037'" & _
            " WHERE SP.cCodProceso='" & pcCodProceso & "' AND cPersCodTitu IS NOT NULL"
    Case 2  'Todos
            sSql = "SELECT P.cPersCod,Miembro=P.CPersNombre,1,SPC.nCargo,Cargo=C.cConsDescripcion " & _
            "FROM Persona P INNER JOIN RHSeleccionProcesoComite SPC " & _
            "ON P.cPersCod = SPC.cPersCod INNER JOIN RHSeleccionProceso SP " & _
            "ON SPC.cCodProceso =SP.cCodProceso     INNER JOIN Constante C " & _
            "ON C.nConsValor= SPC.nCargo AND C.nConsCod='9037'" & _
            " WHERE SP.cCodProceso='" & pcCodProceso & "' ORDER BY ISNULL(SPC.cPersCodTitu,'')"
    End Select
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Miembros_ComiteSelec = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Miembros_ComiteSelec:
    Err.Raise Err.Number, "Obtener_Miembros_ComiteSelec", Err.Description
End Function

Public Function Obtener_Miembro_Suplente(ByVal pcCodProceso As String, _
                                            ByVal pcPersCodTitu As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Miembro_Suplente
    sSql = "SELECT P.cPersCod,Miembro=P.CPersNombre FROM Persona P INNER JOIN RHSeleccionProcesoComite SPC " & _
           " ON P.cPersCod = SPC.cPersCod INNER JOIN RHSeleccionProceso SP ON SPC.cCodProceso =SP.cCodProceso " & _
           " WHERE cPersCodTitu='" & pcPersCodTitu & "' AND SP.cCodProceso='" & pcCodProceso & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Miembro_Suplente = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Miembro_Suplente:
    Err.Raise Err.Number, "Obtener_Miembro_Suplente", Err.Description
End Function

'Buscar si los miembros del Comite pueden revisar o no las Fases de acuerdo a su Orden  de Revision asignada
Public Function Validar_Revisiones_Previas(ByVal pcCodProceso As String, _
                                            ByVal pcPersCod As String, _
                                            ByVal pnCodFase As Long) As Boolean
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
Dim nOrdenRevision As Integer

On Error GoTo ErrorValidar_Revisiones_Previas
    
sSql = "SELECT nOrdenRevision FROM RHSeleccionProcesoComite WHERE cCodProceso='" & pcCodProceso & "' AND cPersCod='" & pcPersCod & "'"
 
Set oConecta = New DConecta
If oConecta.AbreConexion Then

    Set rs = oConecta.CargaRecordSet(sSql)
    
    If Not rs.EOF Then
        nOrdenRevision = rs("nOrdenRevision")
    Else
        nOrdenRevision = 0
    End If
    
    If nOrdenRevision = 1 Then  'El Primero en Revisar es Recursos Humanos
        Validar_Revisiones_Previas = True
    Else
    
        'Solo cuando recursos lo haya APROBADO deben poder listarlo los demas miembros del Comite
        sSql = "SELECT SFA.nCodFase,SFA.cCodProceso,SFA.cPersCod FROM RHSeleccionFaseAprobacion SFA INNER JOIN RHSeleccionProcesoComite SPC " & _
               " ON SFA.cCodProceso= SPC.cCodProceso AND SFA.cPersCod=SPC.cPersCod INNER JOIN RHSeleccionEvaluacionHistorico SEH ON SEH.nCodFase=SFA.nCodFase AND SEH.cCodProceso=SFA.cCodProceso AND SEH.nLista=SFA.nLista AND SEH.bEstado=1 " & _
               " WHERE SFA.cCodProceso='" & pcCodProceso & "' AND SFA.nCodFase=" & pnCodFase & " AND SFA.cPersCod<>'" & pcPersCod & "' AND SFA.nAceptado=1 AND nOrdenRevision <" & nOrdenRevision

        Set rs = oConecta.CargaRecordSet(sSql)
        
        If rs.EOF Then  'Deben existir las Revisiones de Orden Anterior
            Validar_Revisiones_Previas = False
        Else
            Validar_Revisiones_Previas = True
        End If
    End If
    oConecta.CierraConexion
    
End If
    Set oConecta = Nothing
    Exit Function

ErrorValidar_Revisiones_Previas:
    Err.Raise Err.Number, "Validar_Revisiones_Previas", Err.Description
    
End Function

Public Function Obtener_FasesProceso_x_CargoRequerido(ByVal pcCodCargoSol As String, _
                                                      ByVal pcCodProceso As String, _
                                                      Optional ByVal pbSoloEjecutan As Boolean = False) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
Dim rsProcesos As New ADODB.Recordset
Dim lcCodProceso As String
Dim nNumMiembros As Integer

    On Error GoTo ErrorObtener_FasesProceso_x_CargoRequerido
    
    Set oConecta = New DConecta
    
    If oConecta.AbreConexion Then
        sSql = " SELECT DISTINCT SP.cCodProceso FROM RHSeleccionProceso SP INNER JOIN " & _
               " RHSeleccionRequerimientoProceso SRP ON SRP.cCodProceso=SP.cCodProceso INNER JOIN RHSeleccionRequerimiento SR ON SRP.cCodRequerimiento= SR.cCodRequerimiento " & _
               " WHERE SR.cCodCargoSol='" & pcCodCargoSol & "' AND SP.cCodProceso<>'" & pcCodProceso & "' ORDER BY SP.cCodProceso DESC"
        
        Set rsProcesos = oConecta.CargaRecordSet(sSql)
               
        Do While Not rsProcesos.EOF
            lcCodProceso = rsProcesos("cCodProceso")
            
            sSql = "SELECT NumReg=COUNT(*) FROM RHSeleccionFase WHERE cCodProceso='" & lcCodProceso & "'"
            Set rs = oConecta.CargaRecordSet(sSql)
            If CInt(rs("NumReg")) >= 12 Then
                Exit Do
            End If
            rsProcesos.MoveNext
        Loop
        
        sSql = "SELECT SF.nCodFase,SFS.cNomFase ,SF.dFechaIni, SF.dFechaFin ,SFS.cResponsable,SF.nOrden, PjeMaxPond=(SF.nPonderacion)*20 ,SF.nNotaMin, SF.nNotaMax, SF.nNotaAprobatoria " & _
           " FROM RHSeleccionFase SF INNER JOIN RHSeleccionFases SFS ON SF.nCodFase=SFS.nCodFase AND SF.cCodProceso='" & lcCodProceso & IIf(pbSoloEjecutan = False, "' ", "' AND SFS.bEjecuta=1 ")
    
        'oConecta.ConexionActiva.Execute sSql
        Set Obtener_FasesProceso_x_CargoRequerido = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Set rs = Nothing
    Set rsProcesos = Nothing
    Exit Function

ErrorObtener_FasesProceso_x_CargoRequerido:
    Err.Raise Err.Number, "Obtener_FasesProceso_x_CargoRequerido", Err.Description
End Function

Public Function Obtener_ComiteSelec_x_CargoRequerido(ByVal pcCodCargoSol As String, _
                                                        ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
Dim rsProcesos As New ADODB.Recordset
Dim lcCodProceso As String
Dim nNumMiembros As Integer

    On Error GoTo ErrorObtener_ComiteSelec_x_CargoRequerido
    Set oConecta = New DConecta
    
    If oConecta.AbreConexion Then
        sSql = " SELECT DISTINCT SP.cCodProceso FROM RHSeleccionProceso SP INNER JOIN RHSeleccionRequerimientoProceso SRP ON SRP.cCodProceso=SP.cCodProceso " & _
              " INNER JOIN RHSeleccionRequerimiento SR ON SRP.cCodRequerimiento= SR.cCodRequerimiento WHERE SR.cCodCargoSol='" & pcCodCargoSol & "' AND SP.cCodProceso<>'" & pcCodProceso & "'" & _
              " ORDER BY SP.cCodProceso DESC"
        
        Set rsProcesos = oConecta.CargaRecordSet(sSql)
               
        Do While Not rsProcesos.EOF
            lcCodProceso = rsProcesos("cCodProceso")
            sSql = "SELECT NumReg=COUNT(*) FROM RHSeleccionProcesoComite WHERE cCodProceso='" & lcCodProceso & "'"
            Set rs = oConecta.CargaRecordSet(sSql)
            If CInt(rs("NumReg")) = 6 Then
                Exit Do
            End If
            rsProcesos.MoveNext
        Loop
        
        sSql = "SELECT P.cPersCod,Miembro=P.CPersNombre,1,SPC.nCargo,Cargo=C.cConsDescripcion + CASE WHEN cPersCodTitu IS NULL THEN '' ELSE ' - SUPLENTE' END," & _
               " cPersCodTitu=ISNULL(cPersCodTitu,''),Titular=ISNULL(T.cPersNombre,''),nOrdenRevision=ISNULL(nOrdenRevision,0) FROM Persona P INNER JOIN RHSeleccionProcesoComite SPC ON P.cPersCod = SPC.cPersCod INNER JOIN RHSeleccionProceso SP " & _
               " ON SPC.cCodProceso =SP.cCodProceso INNER JOIN Constante C ON C.nConsValor= SPC.nCargo AND C.nConsCod='9037' LEFT JOIN Persona T ON T.cPersCod=SPC.cPersCodTitu " & _
               " WHERE SP.cCodProceso=" & lcCodProceso & " ORDER BY C.nConsValor"
        
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_ComiteSelec_x_CargoRequerido = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Set rs = Nothing
    Exit Function

ErrorObtener_ComiteSelec_x_CargoRequerido:
    Err.Raise Err.Number, "Obtener_ComiteSelec_x_CargoRequerido", Err.Description
End Function

'Para cambiar el Estado del Requerimiento al ser aprobado por todos los Niveles
Public Function Validar_Aprobaciones_x_Requerimento(ByVal pcCodRequerimento As String) As Boolean

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
Dim nNumNiveles As Integer
Dim nNumAprobaciones As Integer

On Error GoTo ErrorValidar_Aprobaciones_x_Requerimento
    
Validar_Aprobaciones_x_Requerimento = False
    
'Numero de Aprobaciones que debe tener
sSql = "SELECT Niveles=COUNT(*)FROM RHSeleccionRequerimientoAprobaciones" & _
       " WHERE cCodRequerimiento='" & pcCodRequerimento & "'"
 
Set oConecta = New DConecta
If oConecta.AbreConexion Then

    Set rs = oConecta.CargaRecordSet(sSql)
    
    If Not rs.EOF Then
        nNumNiveles = rs("Niveles")
    Else
        nNumNiveles = 0
    End If
    
    sSql = "SELECT Aprobaciones=COUNT(*)FROM RHSeleccionRequerimientoAprobaciones" & _
            " WHERE cCodRequerimiento='" & pcCodRequerimento & "' AND bAprobado=1"

    Set rs = oConecta.CargaRecordSet(sSql)
        
    If rs.EOF Then
        nNumAprobaciones = 0
    Else
        nNumAprobaciones = rs("Aprobaciones")
    End If
        
    If nNumAprobaciones = nNumNiveles Then  'Aprobar Finalmente el Requerimiento
        sSql = "UPDATE RHSeleccionRequerimiento SET nEstado=4 WHERE cCodRequerimiento='" & pcCodRequerimento & "'"
        oConecta.ConexionActiva.Execute sSql
    
        'Se tiene el Requerimiento Listo para las Bases
        Validar_Aprobaciones_x_Requerimento = True
    End If
    oConecta.CierraConexion
    
End If
    Set oConecta = Nothing
    Exit Function

ErrorValidar_Aprobaciones_x_Requerimento:
    Err.Raise Err.Number, "Validar_Aprobaciones_x_Requerimento", Err.Description
End Function

Public Function Obtener_TodasFases_ProcesoSelecc(ByVal pcCodProceso As String, _
                                                Optional ByVal pbSoloEjecutan As Boolean = False) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_TodasFases_ProcesoSelecc
    
    sSql = "SELECT SF.nCodFase,SFS.cNomFase ,SF.dFechaIni, SF.dFechaFin ,SFS.cResponsable,SF.nOrden, PjeMaxPond=(SF.nPonderacion)*20 ,SF.nNotaMin, SF.nNotaMax, SF.nNotaAprobatoria " & _
           " FROM RHSeleccionFase SF INNER JOIN RHSeleccionFases SFS ON SF.nCodFase=SFS.nCodFase AND SF.cCodProceso='" & pcCodProceso & "' " & IIf(pbSoloEjecutan = False, "", "AND SFS.bEjecuta=1")
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_TodasFases_ProcesoSelecc = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_TodasFases_ProcesoSelecc:
    Err.Raise Err.Number, "Obtener_TodasFases_ProcesoSelecc", Err.Description
End Function

Public Function Obtener_PromedioNotas_PostulanteHistorico(ByVal pcCodProceso As String, _
                                                 ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_PromedioNotas_PostulanteHistorico
    sSql = "sp_RHObtener_PromedioNotas_Postulante_Historico '" & pcCodProceso & "'," & pnTipo
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_PromedioNotas_PostulanteHistorico = oConecta.Ejecutar(sSql)
        
    End If
    'oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_PromedioNotas_PostulanteHistorico:
    Err.Raise Err.Number, "Obtener_PromedioNotas_PostulanteHistorico", Err.Description
End Function

'Validar si el Miembro del Comite ya aprobo la Lista Actual
Public Function Buscar_Aprobaciones_Previas(ByVal pcCodProceso As String, _
                                            ByVal pcPersCod As String, _
                                            ByVal pnCodFase As Long) As Boolean
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

On Error GoTo ErrorBuscar_Aprobaciones_Previas
       
'No tomar en cuenta la Lista de Aprobacion
'sSQL = " SELECT SFA.cCodProceso,SFA.cPersCod,SFA.nCodFase FROM RHSeleccionFaseAprobacion SFA " & _
       " INNER JOIN RHSeleccionEvaluacionHistorico SEH ON SEH.nCodFase=SFA.nCodFase AND SEH.cCodProceso=SFA.cCodProceso AND SEH.nLista=SFA.nLista AND SEH.bEstado=1 " & _
        "WHERE SFA.cCodProceso='" & pcCodProceso & "' AND SFA.cPersCod='" & pcPersCod & "' AND SFA.nCodFase=" & pnCodFase

sSql = " SELECT SFA.cCodProceso,SFA.cPersCod,SFA.nCodFase FROM RHSeleccionFaseAprobacion SFA " & _
       " INNER JOIN RHSeleccionEvaluacion SE ON SE.nCodFase=SFA.nCodFase AND SE.cCodProceso=SFA.cCodProceso " & _
        "WHERE SFA.cCodProceso='" & pcCodProceso & "' AND SFA.cPersCod='" & pcPersCod & "' AND SFA.nCodFase=" & pnCodFase

Set oConecta = New DConecta
If oConecta.AbreConexion Then
    Set rs = oConecta.CargaRecordSet(sSql)
    Buscar_Aprobaciones_Previas = Not rs.EOF
    oConecta.CierraConexion
End If
Set oConecta = Nothing
Exit Function

ErrorBuscar_Aprobaciones_Previas:
    Err.Raise Err.Number, "Buscar_Aprobaciones_Previas", Err.Description
End Function

'Validar que se hayan ingresado Notas en la Lista Actual
Public Function Buscar_Notas_FaseAprobacion(ByVal pcCodProceso As String, _
                                            ByVal pnCodFase As Long) As Boolean
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

On Error GoTo ErrorBuscar_Notas_FaseAprobacion
             
'Por el Momento no se va manejar Lista de Aprobacion
'sSQL = " SELECT cCodProceso,nCodFase FROM RHSeleccionEvaluacionHistorico " & _
        "WHERE cCodProceso='" & pcCodProceso & "' AND nCodFase=" & pnCodFase & " AND bEstado=1"  'bEstado=1 ==>Lista Actual

sSql = " SELECT cCodProceso,nCodFase FROM RHSeleccionEvaluacion " & _
        "WHERE cCodProceso='" & pcCodProceso & "' AND nCodFase=" & pnCodFase

Set oConecta = New DConecta
If oConecta.AbreConexion Then
    Set rs = oConecta.CargaRecordSet(sSql)
    Buscar_Notas_FaseAprobacion = Not rs.EOF
    oConecta.CierraConexion
End If
Set oConecta = Nothing
Exit Function

ErrorBuscar_Notas_FaseAprobacion:
    Err.Raise Err.Number, "Buscar_Notas_FaseAprobacion", Err.Description
End Function

Public Function Obtener_numVacantes_x_RequerimientoAceptado(ByVal pcRHCargoCod As String, pcAgeCod As String) As Long
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

On Error GoTo ErrorObtener_numVacantes_x_RequerimientoAceptado
    
    sSql = "Select nvacantes=isnull(sum(nNumvacantes),0) From RHSeleccionRequerimiento  Where cCodCargoSol='" & pcRHCargoCod & "' and cCodAgeSol='" & pcAgeCod & "' AND nEstado=4 "
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_numVacantes_x_RequerimientoAceptado = 0
        Else
            Obtener_numVacantes_x_RequerimientoAceptado = rs("nvacantes")
        End If
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_numVacantes_x_RequerimientoAceptado:
    Err.Raise Err.Number, "Obtener_numVacantes_x_RequerimientoAceptado", Err.Description
End Function

Public Function Obtener_numVacantes_x_ProcesoSelecc(ByVal pcCodProceso As String) As Long
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

On Error GoTo ErrorObtener_numVacantes_x_ProcesoSelecc
    
    sSql = "SELECT Vacantes=SUM(SR.nNumVacantes) FROM RHSeleccionRequerimiento SR " & _
           " INNER JOIN RHSeleccionRequerimientoProceso SRP ON SR.cCodRequerimiento= SRP.cCodRequerimiento " & _
           " AND cCodProceso='" & pcCodProceso & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set rs = oConecta.CargaRecordSet(sSql)
        If rs.EOF Then
            Obtener_numVacantes_x_ProcesoSelecc = 0
        Else
            Obtener_numVacantes_x_ProcesoSelecc = rs("Vacantes")
        End If
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function
ErrorObtener_numVacantes_x_ProcesoSelecc:
    Err.Raise Err.Number, "Obtener_numVacantes_x_ProcesoSelecc", Err.Description
End Function

Public Function Obtener_Criterios_Admisibilidad(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Criterios_Admisibilidad
    sSql = "sp_RHObtener_Criterios_Admisibilidad '" & pcCodProceso & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Criterios_Admisibilidad = oConecta.Ejecutar(sSql)
        
    End If
    'oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Criterios_Admisibilidad:
    Err.Raise Err.Number, "Obtener_Criterios_Admisibilidad", Err.Description
End Function

Public Function Obtener_Aprobaciones_Gerencia(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Aprobaciones_Gerencia
    sSql = "SELECT P.cPersCod,cPersNombre,Estado=C.cConsDescripcion,cComentario FROM RHSeleccionGerenciaAprobacion SGA" & _
           " INNER JOIN Persona P ON SGA.cPersCod=P.cPersCod INNER JOIN Constante C  ON SGA.nAceptado = C.nConsValor AND C.nConsCod='9069'" & _
           " WHERE SGA.cCodProceso='" & pcCodProceso & "'"
            
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

Public Function Obtener_TipoProceso_x_Requerimiento(ByVal pcCodRequerimiento As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_TipoProceso_x_Requerimiento
    
    sSql = "SELECT nTipoProceso=ISNULL(nTipoProceso,0) From RHSeleccionRequerimiento WHERE cCodRequerimiento='" & pcCodRequerimiento & "'"
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_TipoProceso_x_Requerimiento = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_TipoProceso_x_Requerimiento:
    Err.Raise Err.Number, "Obtener_TipoProceso_x_Requerimiento", Err.Description
End Function

Public Function Obtener_Aprobaciones_Comite_ProcesoDesierto(ByVal pcCodProceso As String, pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Aprobaciones_Comite_ProcesoDesierto
    sSql = "SELECT SCA.cCodProceso,SCA.cPersCod FROM RHSeleccionComiteAprobacion SCA " & _
           " WHERE SCA.cCodProceso='" & pcCodProceso & "' AND SCA.cPersCod='" & pcPersCod & "' AND SCA.nTipoAprob=3"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Aprobaciones_Comite_ProcesoDesierto = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Aprobaciones_Comite_ProcesoDesierto:
    Err.Raise Err.Number, "Obtener_Aprobaciones_Comite_ProcesoDesierto", Err.Description
End Function

Public Function Verificar_Ultimo_NivelAprobacion(ByVal pcCodRequerimento As String) As Boolean

Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset
Dim nNumNiveles As Integer
Dim nNumAprobaciones As Integer

On Error GoTo ErrorVerificar_Ultimo_NivelAprobacion
    
'Numero de Aprobaciones que debe tener
sSql = "SELECT Niveles=COUNT(*)FROM RHSeleccionRequerimientoAprobaciones" & _
       " WHERE cCodRequerimiento='" & pcCodRequerimento & "'"
 
Set oConecta = New DConecta
If oConecta.AbreConexion Then

    Set rs = oConecta.CargaRecordSet(sSql)
    
    If Not rs.EOF Then
        nNumNiveles = rs("Niveles")
    Else
        nNumNiveles = 0
    End If
    
    sSql = "SELECT Aprobaciones=COUNT(*)FROM RHSeleccionRequerimientoAprobaciones" & _
            " WHERE cCodRequerimiento='" & pcCodRequerimento & "' AND bAprobado=1"

    Set rs = oConecta.CargaRecordSet(sSql)
        
    If rs.EOF Then
        nNumAprobaciones = 0
    Else
        nNumAprobaciones = rs("Aprobaciones")
    End If
        
    If nNumAprobaciones = nNumNiveles - 1 Then  'El Ultimo Nivel de Aprobacion
        Verificar_Ultimo_NivelAprobacion = True
    Else
        Verificar_Ultimo_NivelAprobacion = False
    End If
    oConecta.CierraConexion
    
End If
    Set oConecta = Nothing
    Exit Function

ErrorVerificar_Ultimo_NivelAprobacion:
    Err.Raise Err.Number, "Verificar_Ultimo_NivelAprobacion", Err.Description
End Function

Public Function Obtener_Datos_x_PersonaReq(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Datos_x_PersonaReq
    sSql = " select rh.careacod,cagenciaactual,rc.crhcargocod,crhcargodescripcion,careadescripcion,ag.cAgeDescripcion" & _
           " from rrhh rh inner join rhcargos  rc on  rh.cperscod=rc.cperscod and" & _
           " drhcargofecha=(select top 1 dRHCargoFecha from rhcargos where cperscod=rc.cperscod order by dRHCargoFecha desc)" & _
           " inner join rhcargostabla rt on rc.crhcargocod=rt.crhcargocod " & _
           " inner join areas a on a.careacod= rh.careacod" & _
           " inner join agencias ag on ag.cAgeCod=rh.cAgenciaActual " & _
           " where rh.cperscod='" & pcPersCod & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Datos_x_PersonaReq = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Datos_x_PersonaReq:
    Err.Raise Err.Number, "Obtener_Datos_x_PersonaReq", Err.Description
End Function

Public Function Obtener_Posibles_CargosSolicitados(ByVal pcAreaCod As String, _
                                                   ByVal pcCargoCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Posibles_CargosSolicitados
    sSql = "select distinct '" & pcAreaCod & "'+ r.cCodCargoSol codigo,t.cRHCargoDescripcion descripcion,1 nivel " & _
           " from RHSeleccionNivelAprobacion r inner join RHCargosTabla t on r.cCodCargoSol = t.cRHCargoCod" & _
           " where r.cCodCargoReq='" & pcCargoCod & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Posibles_CargosSolicitados = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Posibles_CargosSolicitados:
    Err.Raise Err.Number, "Obtener_Posibles_CargosSolicitados", Err.Description
End Function

Public Function Obtener_Posibles_CargosSolicitados_x_Area(ByVal pcAreaCod As String, _
                                                   ByVal pcCargoCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Posibles_CargosSolicitados_x_Area
    sSql = "select distinct '" & pcAreaCod & "'+ r.cCodCargoSol codigo,t.cRHCargoDescripcion descripcion,1 nivel " & _
           " from RHSeleccionNivelAprobacion r inner join RHCargosTabla t on r.cCodCargoSol = t.cRHCargoCod" & _
           " where r.cCodCargoReq='" & pcCargoCod & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Posibles_CargosSolicitados_x_Area = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Posibles_CargosSolicitados_x_Area:
    Err.Raise Err.Number, "Obtener_Posibles_CargosSolicitados_x_Area", Err.Description
End Function

Public Function Obtener_Historico_Requerimientos(ByVal pcAreaCod As String, _
                                                   ByVal pcAgenciaCod As String, _
                                                   ByVal pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Historico_Requerimientos
    
    sSql = "SELECT rq.cCodRequerimiento,cCodArea,rq.cPersCod,cCodAreaSol," & _
            " cCodCargoSol=CASE WHEN cCodCargoSol IS NULL THEN 'NUEVO' ELSE cCodCargoSol END,cRHCargoDescripcion= CASE WHEN cRHCargoDescripcion IS NULL THEN rq.cCargoSolNuevo ELSE cRHCargoDescripcion END," & _
            " cCodAgeSol , cAgeDescripcion, nNumVacantes, dFechaIni, dFechaFin, rq.nEstado, rq.cUltimaActualizacion,C.cConsDescripcion, " & _
            " cComentario=isnull(rqa.cComentario,'')+ CASE WHEN isnull(PersonaRec.cPersNombre,'')='' THEN '' ELSE ' -> '+ isnull(PersonaRec.cPersNombre,'')END , cJustificacion=isnull(rq.cJustificacion,''), " & _
            " Fecha = dbo.GetFechaMov(rq.cUltimaactualizacion, 103) FROM rhseleccionrequerimiento  rq inner join agencias ag on ag.cagecod=rq.ccodagesol " & _
            " inner join Constante C on rq.nEstado=C.nConsValor AND C.nConsCod='9061' left join rhcargostabla rt on rq.cCodCargoSol=rt.crhcargocod " & _
            " left join rhseleccionrequerimientoaprobaciones rqa on rqa.ccodrequerimiento=rq.ccodrequerimiento and rqa.bAprobado=0 and ISNULL(rqa.cPersCodAprobac,'')<>'' " & _
            " left join Persona PersonaRec on PersonaRec.cPersCod=rqa.cPersCodAprobac " & _
            " WHERE rq.nEstado IN(0,2) AND rq.cPersCod='" & pcPersCod & "' ORDER BY rq.cCodRequerimiento"
            
            '" WHERE ccodarea='" & pcAreaCod & "' AND rq.nEstado IN(0,2) AND cCodAgeSol='" & pcAgenciaCod & "' AND rq.cPersCod='" & pcPersCod & "' ORDER BY rq.cCodRequerimiento"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Historico_Requerimientos = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Historico_Requerimientos:
    Err.Raise Err.Number, "Obtener_Historico_Requerimientos", Err.Description
End Function

Public Function Obtener_Requerimientos_x_Aprobar(ByVal pcCargoCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Requerimientos_x_Aprobar
    
    sSql = "SELECT DISTINCT Elije='0',cCodCargoSol=CASE WHEN cCodCargoSol IS NULL THEN 'NUEVO' ELSE cCodCargoSol END,rq.cCodRequerimiento,cRHCargoDescripcion= CASE WHEN rt.cRHCargoDescripcion IS NULL THEN rq.cCargoSolNuevo " & _
               " ELSE rt.cRHCargoDescripcion END, cCodAgeSol , cAgeDescripcion,PersonaReq=P.cPersNombre,CargoReq=CatReq.cRHCargoDescripcion, nNumVacantes, rq.nEstado,c.cConsDescripcion, rq.cJustificacion, '', " & _
               " Nuevo= CASE WHEN rt.cRHCargoDescripcion IS NULL THEN 1 ELSE 0 END,PersonaAprob=ISNULL(PA.cPersNombre,''),CargoAprob=CASE WHEN PA.cPersCod IS NULL THEN '' ELSE CA.crhcargodescripcion END," & _
               " ComentarioAprob=ISNULL(SRA1.cComentario,'')," & _
               " Fecha=CONVERT(varchar(10),dFechaFin,103), Fecha= dbo.getFechaMov(rq.cUltimaactualizacion,103) ,nTipoProceso=ISNULL(CONVERT(varchar(2),nTipoProceso),''),TipoProceso=ISNULL(TIP.cConsDescripcion,'')" & _
               " FROM rhseleccionrequerimiento  rq inner join Constante C ON C.nConsValor=rq.nEstado AND C.nConsCod=9061 " & _
               " inner join agencias ag on ag.cagecod=rq.ccodagesol inner join Areas ar ON rq.cCodAreaSol=ar.cAreaCod " & _
               " inner join Persona P on rq.cPersCod=P.cPersCod inner join RHCargos CarReq ON CarReq.cPersCod=rq.cPersCod AND CarReq.drhcargofecha=(select top 1 dRHCargoFecha from rhcargos where cperscod=rq.cperscod order by dRHCargoFecha desc) " & _
               " inner join RHCargosTabla CatReq ON CarReq.cRHCargoCod=CatReq.cRHCargoCod left join rhcargostabla rt on rq.cCodCargoSol=rt.crhcargocod inner join rhseleccionRequerimientoAprobaciones SRA ON rq.ccodrequerimiento = SRA.ccodrequerimiento " & _
               " left join RHCargos car ON car.cRHCargoCod=rq.cCodCargoSol AND car.cRHCargoCod=rt.cRHCargoCod AND Car.cRHAreaCod = ar.cAreaCod And Car.cRHAgenciaCod = ag.CAgeCod inner join rhseleccionRequerimientoAprobaciones SRA1 ON " & _
               " SRA.cCodRequerimiento = SRA1.cCodRequerimiento AND ((SRA1.nOrden=SRA.nOrden-1 AND SRA1.bAprobado=1) OR (SRA1.nOrden=SRA.nOrden AND SRA1.nOrden=1 AND SRA1.bAprobado=0)) " & _
               " left join RHCargosTabla CA on CA.crhCargoCod=SRA1.cCodCargoAprobac left join Persona PA on PA.cPersCod=SRA1.cPersCodAprobac LEFT JOIN Constante TIP ON TIP.nConsValor=nTipoProceso AND TIP.nConsCod=6003" & _
               " where SRA.ccodcargoaprobac='" & pcCargoCod & "'" & _
               " AND SRA.bAprobado=0 AND rq.nEstado IN(0,1) ORDER BY rq.ccodrequerimiento "
            
            '" WHERE ccodarea='" & pcAreaCod & "' AND rq.nEstado IN(0,2) AND cCodAgeSol='" & pcAgenciaCod & "' AND rq.cPersCod='" & pcPersCod & "' ORDER BY rq.cCodRequerimiento"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Requerimientos_x_Aprobar = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Requerimientos_x_Aprobar:
    Err.Raise Err.Number, "Obtener_Requerimientos_x_Aprobar", Err.Description
End Function

Public Function Buscar_CierreProceso_Previos(ByVal pcCodProceso As String, pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorBuscar_CierreProceso_Previos
    sSql = "SELECT cCodProceso,cPersCod,nTipoAprob FROM RHSeleccionComiteAprobacion " & _
           " WHERE cCodProceso='" & pcCodProceso & "' AND cPersCod='" & pcPersCod & "'AND nTipoAprob=2"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Buscar_CierreProceso_Previos = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorBuscar_CierreProceso_Previos:
    Err.Raise Err.Number, "Buscar_CierreProceso_Previos", Err.Description
End Function

Public Function Obtener_Requerimientos_x_ReporteProceso(ByVal pcCodProceso As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Requerimientos_x_ReporteProceso
    
    sSql = " SELECT DISTINCT rq.cCodRequerimiento , rt.cRHCargoDescripcion , cCodAgeSol , cAgeDescripcion,PersonaReq=P.cPersNombre,CargoReq=CatReq.cRHCargoDescripcion, rq.nNumVacantes, rq.nEstado,c.cConsDescripcion, rq.cJustificacion, " & _
           " FechaIngreso=CONVERT(varchar(10),dFechaFin,103), Fecha= dbo.getFechaMov(rq.cUltimaactualizacion,103),TipoProceso=ISNULL(TIP.cConsDescripcion,''),Ar.cAreaDescripcion " & _
           " FROM rhseleccionrequerimiento  rq inner join RHSeleccionRequerimientoProceso SRP ON rq.cCodRequerimiento=SRP.cCodRequerimiento inner join Constante C ON C.nConsValor=rq.nEstado AND C.nConsCod=9061 " & _
           " inner join agencias ag on ag.cagecod=rq.ccodagesol inner join Areas ar ON rq.cCodAreaSol=ar.cAreaCod inner join Persona P on rq.cPersCod=P.cPersCod inner join RHCargos CarReq ON CarReq.cPersCod=rq.cPersCod AND " & _
           " CarReq.drhcargofecha=(select top 1 dRHCargoFecha from rhcargos where cperscod=rq.cperscod order by dRHCargoFecha desc) inner join RHCargosTabla CatReq ON CarReq.cRHCargoCod=CatReq.cRHCargoCod left join rhcargostabla rt on rq.cCodCargoSol=rt.crhcargocod " & _
           " LEFT JOIN RHCargos car ON car.cRHCargoCod=rq.cCodCargoSol AND car.cRHCargoCod=rt.cRHCargoCod AND Car.cRHAreaCod = ar.cAreaCod And Car.cRHAgenciaCod = ag.CAgeCod " & _
           " LEFT JOIN Constante TIP ON TIP.nConsValor=nTipoProceso AND TIP.nConsCod=6003 " & _
           " WHERE SRP.cCodProceso='" & pcCodProceso & "'"
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Requerimientos_x_ReporteProceso = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Requerimientos_x_ReporteProceso:
    Err.Raise Err.Number, "Obtener_Requerimientos_x_ReporteProceso", Err.Description
End Function

Public Function Obtener_Historico_ProcesosSeleccion(ByVal pcAnio As String, _
                                                    ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Historico_ProcesosSeleccion
    sSql = "sp_RHObtener_Historico_ProcesosSeleccion '" & pcAnio & "'," & pnTipo
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Historico_ProcesosSeleccion = oConecta.Ejecutar(sSql)
        
    End If
    'oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Historico_ProcesosSeleccion:
    Err.Raise Err.Number, "Obtener_Historico_ProcesosSeleccion", Err.Description
End Function

Public Function Obtener_Historico_Postulante(ByVal pcPostuCod As String, _
                                                    ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorObtener_Historico_Postulante
    sSql = "sp_RHObtener_Historico_Postulante '" & pcPostuCod & "'," & pnTipo
    
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Historico_Postulante = oConecta.Ejecutar(sSql)
        
    End If
    'oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Historico_Postulante:
    Err.Raise Err.Number, "Obtener_Historico_Postulante", Err.Description
End Function

Public Function Obtener_Postulantes_x_Parametros(ByVal pcParametro As String, _
                                                ByVal pnTipoParametro As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim cCadenaBusqueda As String

    On Error GoTo ErrorObtener_Postulantes_x_Parametros
    'pnTipoParametro = 0: cPostuCod, 1: cDNI, 2 : APellidos y Nombres
    
    Select Case pnTipoParametro
        Case 0
            cCadenaBusqueda = "cPostuCod='" & pcParametro & "'"
        Case 1
            cCadenaBusqueda = "cDNI='" & pcParametro & "'"
        Case 2
            cCadenaBusqueda = "RTRIM(cApellidoPat)+' '+RTRIM(cApellidoMat)+'' +RTRIM(cNombre) LIKE '" & pcParametro & "' +'%'"
    End Select
    
    sSql = " SELECT cPostuCod,Nombre= RTRIM(cApellidoPat)+' '+RTRIM(cApellidoMat)+','+RTRIM(cNombre),cDNI FROM Postulante" & _
           " WHERE " & cCadenaBusqueda
            
    Set oConecta = New DConecta
    If oConecta.AbreConexion Then
    'oConecta.ConexionActiva.Execute sSql
        Set Obtener_Postulantes_x_Parametros = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion

    Set oConecta = Nothing
    Exit Function

ErrorObtener_Postulantes_x_Parametros:
    Err.Raise Err.Number, "Obtener_Postulantes_x_Parametros", Err.Description
End Function
