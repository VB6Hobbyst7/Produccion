VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogRequeri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsCentralPer As String
Dim vsCentralCom As String
Dim vsCentralImg As String
Dim sSql As String

Public Enum TpoCargaReq
    ReqUnRegistro = 0
    ReqUnRegistroTramite = 1
    ReqTodosAreaFlex = 2
    ReqTodosFlexConsol = 3
    ReqTodosAreaTraNuevo = 4
    ReqTodosAreaTraIngreso = 5
    ReqTodosAreaTraEgreso = 6
    ReqTodosTraPrecio = 7
    ReqTodosTraCuenta = 8
    ReqTodosObten = 9
    ReqTodosFlexApro = 10
End Enum

Public Enum TpoCargaReqDet
    ReqDetUnRegistroTramite = 0
    ReqDetUnRegistroTramiteUlt = 1
    ReqDetUnRegistroPrecio = 2
    ReqDetUnRegistroConsol = 3
    ReqDetTodosFlex = 4
End Enum

Public Enum TpoCargaReqTra
    ReqTraUnRegistro = 0
    ReqTraTodosAreaMenosPrimero = 1
    ReqTraTodosArea = 2
    ReqTraTodosAreaMasDes = 3
End Enum

Public Enum TpoCargaObtDet
    ObtDetParaCtaCnt = 0
    ObtDetParaAprobar = 1
    ObtDetParaAdquisi = 2
    ObtDetParaAdquisiConsol = 3
End Enum

'''Public Enum TpoCargaAdq
'''    AdqTodosGnral = 5
'''    AdqTodosEstado = 6
'''End Enum
'''
'''Public Enum TpoCargaAdqDet
'''    AdqDetUnRegistro = 0
'''End Enum

Private Sub Class_Initialize()
    Dim oini As ClasIni
    
    Set oini = New ClasIni
        vsConexion = oini.CadenaConexion
        vsCentralPer = oini.BasePersonas
        vsCentralCom = oini.BaseComunes
        vsCentralImg = oini.BaseImagenes
    Set oini = Nothing
End Sub

'Devuelve un RecordSet con los requerimientos
Public Function CargaRequerimiento(ByVal psTpoReq As String, ByVal psTpoCarga As TpoCargaReq, ByVal psAreaCod As String, _
Optional ByVal psReqNro As String, Optional ByVal psPeriodo As String = "") As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        If psTpoCarga = ReqUnRegistro Then
            sSql = "SELECT r.cLogReqNro, r.cLogReqNecesidad, r.cLogReqRequerimiento, " & _
                "       r.cLogReqTpo, r.cLogReqPeriodo, rt.cAreaCod, a.cAreaDescripcion " & _
                " FROM  LogRequerimiento R INNER JOIN LogReqTramite RT ON r.cLogReqNro = rt.cLogReqNro " & _
                "       LEFT JOIN Areas A ON rt.cAreaCod = a.cAreaCod " & _
                " WHERE r.cLogReqNro = '" & psReqNro & "' " & _
                "       AND rt.cLogReqNro = rt.cLogReqTraNro " & _
                "       AND r.cLogReqTpo = '" & psTpoReq & "'"
        ElseIf psTpoCarga = ReqUnRegistroTramite Then
            'Un Registro(con datos del ultimo cLogReqTraNro enviado)
            sSql = "SELECT r.cLogReqNro, r.cLogReqNecesidad, r.cLogReqRequerimiento, " & _
                "       r.cLogReqTpo, r.cLogReqPeriodo, rt.cLogReqTraNro " & _
                " FROM  LogRequerimiento R INNER JOIN LogReqTramite RT ON r.cLogReqNro = rt.cLogReqNro " & _
                " WHERE r.cLogReqNro = '" & psReqNro & "' " & _
                "       AND rt.cLogReqTraNro = " & _
                "           (SELECT MAX(cLogReqTraNro) FROM LogReqTramite " & _
                "            WHERE cLogReqNro = '" & psReqNro & "' " & _
                "               AND cLogReqTraEstado <> '" & gLogReqEstadoInicio & "')" & _
                "       AND r.cLogReqTpo = '" & psTpoReq & "'"
        ElseIf psTpoCarga = ReqTodosAreaFlex Then
            'Todos los registros de un Area, para Flex (solo para edición de NUEVOS)
            sSql = "SELECT r.cLogReqNro Código, substring(r.cLogReqNecesidad,1,50) Necesidad, " & _
                "       Substring(r.cLogReqRequerimiento,1,50) Requerimiento " & _
                " FROM LogRequerimiento R JOIN LogReqTramite RT ON r.cLogReqNro = rt.cLogReqNro " & _
                "       AND r.cLogReqNro = rt.cLogReqTraNro " & _
                " WHERE rt.cAreaCod = '" & psAreaCod & "' AND rt.cAreaCodDes = ''" & _
                "       AND rt.cLogReqTraEstado = '" & gLogReqEstadoInicio & "' " & _
                "       AND r.cLogReqTpo = '" & psTpoReq & "' "
        ElseIf psTpoCarga = ReqTodosFlexConsol Then
            'Todos los registros (Todas las areas para CONSOLIDACION)
            sSql = "SELECT (SELECT cAreaCod " & _
                "           FROM LogRequerimiento M JOIN LogReqTramite N ON m.cLogReqNro = n.cLogReqNro " & _
                "                   AND n.cLogReqNro = n.cLogReqTraNro AND m.cLogReqNro = r.cLogReqNro) cAreaCod, " & _
                "          (SELECT q.cAreaDescripcion " & _
                "           FROM LogRequerimiento O JOIN LogReqTramite P ON o.cLogReqNro = p.cLogReqNro " & _
                "                   AND p.cLogReqNro = p.cLogReqTraNro AND o.cLogReqNro = r.cLogReqNro " & _
                "               LEFT JOIN Areas Q ON p.cAreaCod = q.cAreaCod) Area, " & _
                "       r.cLogReqNro Requerimiento " & _
                "       " & _
                " FROM LogRequerimiento R JOIN LogReqTramite RT ON r.cLogReqNro = rt.cLogReqNro " & _
                "       AND rt.cLogReqTraNro = (SELECT MAX(cLogReqTraNro) " & _
                "                               FROM LogReqTramite " & _
                "                               WHERE cLogReqNro = r.cLogReqNro) " & _
                " WHERE rt.cLogReqTraEstado = '" & IIf(psTpoReq = "1", gLogReqEstadoAceptado, gLogReqEstadoPrecio) & "' " & _
                "       AND r.cLogReqPeriodo = '" & psPeriodo & "' " & _
                "       AND r.cLogReqTpo = '" & psTpoReq & "'"
                
        ElseIf psTpoCarga = ReqTodosAreaTraNuevo Then
            'Todos los registros de un Area, NUEVOS
            sSql = "SELECT r.cLogReqNro Código, a.cAreaDescripcion, r.cLogReqPeriodo, " & _
                "       substring(r.cLogReqNecesidad,1,50) Necesidad, " & _
                "       substring(r.cLogReqRequerimiento,1,50) Requerimiento, " & _
                "       (SELECT c.cConsDescripcion FROM LogReqTramite RT JOIN Constante C " & _
                "        ON rt.cLogReqTraEstado = c.cConsValor AND c.cConsCod = " & gLogReqEstado & " " & _
                "        AND rt.cLogReqTraNro = r.cLogReqNro ) Estado " & _
                " FROM LogRequerimiento R JOIN LogReqTramite RT ON r.cLogReqNro = rt.cLogReqNro " & _
                "       LEFT JOIN Areas A ON rt.cAreaCod = a.cAreaCod " & _
                " WHERE rt.cAreaCod = '" & psAreaCod & "' AND rt.cAreaCodDes = '' AND " & _
                "       rt.cLogReqTraEstado = '" & gLogReqEstadoInicio & "'  AND" & _
                "       rt.cLogReqNro = (SELECT MAX(cLogReqTraNro) FROM LogReqTramite " & _
                "                       WHERE cLogReqNro = r.cLogReqNro) " & _
                "       AND r.cLogReqTpo = '" & psTpoReq & "'"
        ElseIf psTpoCarga = ReqTodosAreaTraIngreso Then
            'Todos los registros de un Area, para Flex - Tramite INGRESOS
            sSql = " SELECT  r.cLogReqNro Código, " & _
                "       (SELECT b.cAreaDescripcion FROM LogReqTramite A JOIN Areas B " & _
                "               On a.cAreaCod = b.cAreaCod WHERE a.cLogReqNro = a.cLogReqTraNro " & _
                "               And a.cLogReqNro = r.cLogReqNro ) Procedencia,  r.cLogReqPeriodo, " & _
                "       substring(r.cLogReqNecesidad, 1, 50) Necesidad, " & _
                "       substring(r.cLogReqRequerimiento, 1, 50) Requerimiento, " & _
                "       c.cConsDescripcion " & _
                " FROM LogRequerimiento R JOIN LogReqTramite RT ON r.cLogReqNro = rt.cLogReqNro " & _
                "       LEFT JOIN Constante C ON rt.cLogReqTraEstado = c.cConsValor  " & _
                "               AND c.cConsCod = " & gLogReqEstado & " AND c.cConsValor <> " & gLogReqEstado & " " & _
                " WHERE rt.cAreaCodDes = '" & psAreaCod & "' AND " & _
                "       rt.cLogReqTraNro = (SELECT MAX(cLogReqTraNro) FROM LogReqTramite " & _
                "                       WHERE cLogReqNro = r.cLogReqNro) " & _
                "       AND r.cLogReqTpo = '" & psTpoReq & "' "
        ElseIf psTpoCarga = ReqTodosAreaTraEgreso Then
            'Todos los registros de un Area, para Flex - Tramite EGRESOS
            sSql = "SELECT r.cLogReqNro Código, a.cAreaDescripcion,  r.cLogReqPeriodo," & _
                "       substring(r.cLogReqNecesidad,1,50) Necesidad," & _
                "       substring(r.cLogReqRequerimiento,1,50) Requerimiento, " & _
                "       (SELECT c.cConsDescripcion " & _
                "        FROM LogReqTramite RT JOIN Constante C " & _
                "        ON rt.cLogReqTraEstado = c.cConsValor AND c.cConsCod = " & gLogReqEstado & " " & _
                "        AND rt.cLogReqNro = r.cLogReqNro AND rt.cLogReqTraNro = " & _
                "                   (SELECT MAX(cLogReqTraNro) FROM LogReqTramite " & _
                "                   WHERE cLogReqNro = r.cLogReqNro )) Estado " & _
                " FROM LogRequerimiento R JOIN LogReqTramite RT ON r.cLogReqNro = rt.cLogReqNro " & _
                "                AND rt.cLogReqNro = rt.cLogReqTraNro  " & _
                "       LEFT JOIN Areas A ON rt.cAreaCod = a.cAreaCod " & _
                " WHERE rt.cAreaCod = '" & psAreaCod & "' AND " & _
                "       (rt.cAreaCodDes <> '' OR rt.cLogReqTraEstado = '" & gLogReqEstadoRechazado & "' " & _
                "        OR (SELECT COUNT(*) FROM LogReqTramite WHERE cLogReqNro = r.cLogReqNro) > 1 )" & _
                "       AND r.cLogReqTpo = '" & psTpoReq & "' "
                'rt.cAreaCodDes <> '' OR
        ElseIf psTpoCarga = ReqTodosTraPrecio Then
            'TODOS los registros de todas las areas para Ingresar PRECIO
            sSql = "SELECT r.cLogReqNro Código, " & _
                "       (SELECT b.cAreaDescripcion " & _
                "        FROM LogReqTramite A LEFT JOIN Areas B ON a.cAreaCod = b.cAreaCod " & _
                "        WHERE a.cLogReqNro = a.cLogReqTraNro " & _
                "        AND a.cLogReqNro = r.cLogReqNro ) Procedencia,  r.cLogReqPeriodo," & _
                "       SUBSTRING(r.cLogReqNecesidad,1,50) Necesidad, " & _
                "       SUBSTRING(r.cLogReqRequerimiento,1,50) Requerimiento, " & _
                "       c.cConsDescripcion Estado " & _
                " FROM LogRequerimiento R JOIN LogReqTramite RT ON r.cLogReqNro = rt.cLogReqNro " & _
                "       LEFT JOIN Constante C ON rt.cLogReqTraEstado = c.cConsValor AND c.cConsCod = " & gLogReqEstado & " " & _
                " WHERE rt.cLogReqTraEstado = '" & gLogReqEstadoAcepPrevio & "'" & _
                "       AND rt.cLogReqTraNro = (SELECT MAX(cLogReqTraNro) FROM LogReqTramite " & _
                "                       WHERE cLogReqNro = r.cLogReqNro) " & _
                "       AND r.cLogReqTpo = '" & psTpoReq & "'"
        ElseIf psTpoCarga = ReqTodosTraCuenta Then
            'TODOS los registros de todas las areas para Ingresar CUENTAS
            sSql = "SELECT r.cLogReqNro Código, " & _
                "       (SELECT b.cAreaDescripcion " & _
                "        FROM LogReqTramite A LEFT JOIN Areas B ON a.cAreaCod = b.cAreaCod " & _
                "        WHERE a.cLogReqNro = a.cLogReqTraNro " & _
                "        AND a.cLogReqNro = r.cLogReqNro ) Procedencia, " & _
                "       SUBSTRING(r.cLogReqNecesidad,1,50) Necesidad, " & _
                "       SUBSTRING(r.cLogReqRequerimiento,1,50) Requerimiento, " & _
                "       c.cConsDescripcion Estado " & _
                " FROM LogRequerimiento R JOIN LogReqTramite RT ON r.cLogReqNro = rt.cLogReqNro " & _
                "       LEFT JOIN Constante C ON rt.cLogReqTraEstado = c.cConsValor AND c.cConsCod = " & gLogReqEstado & " " & _
                " WHERE rt.cLogReqTraEstado = '" & gLogReqEstadoPrecio & "'" & _
                "       AND rt.cLogReqTraNro = (SELECT MAX(cLogReqTraNro) FROM LogReqTramite " & _
                "                       WHERE cLogReqNro = r.cLogReqNro) " & _
                "       AND r.cLogReqTpo = '" & psTpoReq & "'"
        ElseIf psTpoCarga = ReqTodosObten Then
            sSql = "SELECT a.cAreaDescripcion, ro.cLogReqNro " & _
                " FROM LogReqObt RO JOIN LogReqTramite RT ON ro.cLogReqNro = rt.cLogReqNro  " & _
                "       AND rt.cLogReqNro = rt.cLogReqTraNro " & _
                "       LEFT JOIN Areas A ON rt.cAreaCod = a.cAreaCod " & _
                " WHERE cLogObtNro = '" & psReqNro & "'"
        ElseIf psTpoCarga = ReqTodosFlexApro Then
            'Todos los registros (Todas las areas para Aprobación)
            sSql = "SELECT (SELECT cAreaCod " & _
                "           FROM LogRequerimiento M JOIN LogReqTramite N ON m.cLogReqNro = n.cLogReqNro " & _
                "                   AND n.cLogReqNro = n.cLogReqTraNro AND m.cLogReqNro = r.cLogReqNro) cAreaCod, " & _
                "          (SELECT q.cAreaDescripcion " & _
                "           FROM LogRequerimiento O JOIN LogReqTramite P ON o.cLogReqNro = p.cLogReqNro " & _
                "                   AND p.cLogReqNro = p.cLogReqTraNro AND o.cLogReqNro = r.cLogReqNro " & _
                "               LEFT JOIN Areas Q ON p.cAreaCod = q.cAreaCod) Area, " & _
                "       r.cLogReqNro Requerimiento,  " & _
                "       r.cLogReqNecesidad, r.cLogReqRequerimiento " & _
                " FROM LogRequerimiento R JOIN LogReqTramite RT ON r.cLogReqNro = rt.cLogReqNro " & _
                "       AND rt.cLogReqTraNro = (SELECT MAX(cLogReqTraNro) " & _
                "                               FROM LogReqTramite " & _
                "                               WHERE cLogReqNro = r.cLogReqNro) " & _
                " WHERE rt.cLogReqTraEstado = '" & gLogReqEstadoPrecio & "' " & _
                "       AND r.cLogReqPeriodo = '" & psPeriodo & "' " & _
                "       AND r.cLogReqTpo = '" & psTpoReq & "'"
        Else
            sSql = ""
        End If
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaRequerimiento = Rs
End Function

'Devuelve un RecordSet con los trámites
Public Function CargaReqTramite(ByVal psTpoCarga As TpoCargaReqTra, ByVal psReqNro As String, _
ByVal psReqTraNro As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        If psTpoCarga = ReqTraUnRegistro Then
            'Un Registro
            sSql = " SELECT rt.cLogReqTraNro, rt.cAreaCod, a.cAreaDescripcion, " & _
                "       rt.cLogReqTraComentario " & _
                " FROM LogReqTramite RT LEFT JOIN AREAS A ON rt.cAreaCod = a.cAreaCod " & _
                " WHERE  rt.cLogReqNro = '" & psReqNro & "' AND rt.cLogReqTraNro = '" & psReqTraNro & "' "
        ElseIf psTpoCarga = ReqTraTodosAreaMenosPrimero Then
            'Todos los registros de un Area menos primero sin contestar
            sSql = " SELECT rt.cLogReqTraNro, a.cAreaDescripcion, rt.cLogReqTraEstado, c.cConsDescripcion, " & _
                "       rt.cLogReqTraComentario " & _
                " FROM LogReqTramite RT LEFT JOIN AREAS A ON rt.cAreaCod = a.cAreaCod " & _
                "       LEFT JOIN Constante C ON rt.cLogReqTraEstado = c.cConsValor " & _
                "           AND c.cConsCod = '" & gLogReqEstado & "' " & _
                " WHERE  rt.cLogReqNro = '" & psReqNro & "' AND (SELECT COUNT(*) FROM LogReqTramite WHERE cLogReqNro = 'rt.cLogReqTraNro') > 1 " & _
                " ORDER BY rt.cLogReqTraNro "
        ElseIf psTpoCarga = ReqTraTodosArea Then
            'Todos los registros de un Area
            sSql = " SELECT rt.cLogReqTraNro, a.cAreaDescripcion, rt.cLogReqTraEstado, c.cConsDescripcion, " & _
                "       rt.cLogReqTraComentario " & _
                " FROM LogReqTramite RT LEFT JOIN AREAS A ON rt.cAreaCod = a.cAreaCod " & _
                "       LEFT JOIN Constante C ON rt.cLogReqTraEstado = c.cConsValor " & _
                "           AND c.cConsCod = '" & gLogReqEstado & "' " & _
                " WHERE  rt.cLogReqNro = '" & psReqNro & "' " & _
                " ORDER BY rt.cLogReqTraNro "
        ElseIf psTpoCarga = ReqTraTodosAreaMasDes Then
            sSql = " SELECT rt.cLogReqTraNro, a.cAreaDescripcion, " & gLogReqEstadoParaTramite & " cLogReqTraEstado, c.cConsDescripcion, " & _
                "       '' cLogReqTraComentario " & _
                " FROM LogReqTramite RT LEFT JOIN AREAS A ON rt.cAreaCodDes = a.cAreaCod " & _
                "       LEFT JOIN Constante C ON " & gLogReqEstadoParaTramite & " = c.cConsValor " & _
                "           AND c.cConsCod = '" & gLogReqEstado & "' " & _
                " WHERE  rt.cLogReqNro = '" & psReqNro & "' AND rt.cAreaCodDes <> '' " & _
                " AND rt.cLogReqTraNro = (SELECT MAX(cLogReqTraNro) FROM LogReqTramite " & _
                "                       WHERE cLogReqNro = rt.cLogReqNro) "
        Else
            sSql = ""
        End If
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaReqTramite = Rs
End Function


'Devuelve un RecordSet con los trámites
Public Function CuentaReqTramite(ByVal psReqNro As String) As Integer
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        sSql = "SELECT count(*) " & _
            " FROM LogReqTramite RT " & _
            " WHERE  rt.cLogReqNro = '" & psReqNro & "' "
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    CuentaReqTramite = Rs(0)
End Function

'Devuelve un RecordSet con los trámites
Public Function IsReqTramiteInicial(ByVal psReqNro As String) As Boolean
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        sSql = "SELECT rt.cLogReqNro  " & _
            " FROM LogReqTramite RT " & _
            " WHERE  rt.cLogReqNro = '" & psReqNro & "' AND rt.cLogReqNro  = rt.cLogReqTraNro " & _
            "       AND rt.cAreaCodDes = '' "
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    IsReqTramiteInicial = IIf(Rs.RecordCount > 0, True, False)
End Function

'Devuelve un RecordSet con los trámites
Public Function CuentaReqDetalle(ByVal psReqNro As String, ByVal psReqTraNro As String) As Integer
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        sSql = "SELECT count(*) " & _
            " FROM LogReqDetalle RD " & _
            " WHERE  rd.cLogReqNro = '" & psReqNro & "' AND rd.cLogReqTraNro = '" & psReqTraNro & "'"
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    CuentaReqDetalle = Rs(0)
End Function

'Devuelve un RecordSet con los trámites
Public Function CargaUltReqTraNro(ByVal psReqNro As String) As String
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        sSql = "SELECT MAX(rt.cLogReqTraNro) " & _
            " FROM LogReqTramite RT " & _
            " WHERE  rt.cLogReqNro = '" & psReqNro & "' "
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    CargaUltReqTraNro = Rs(0)
End Function

'Devuelve un RecordSet con los ReqDetalle
Public Function CargaReqDetalle(ByVal TpoCarga As TpoCargaReqDet, ByVal psReqNro As String, _
Optional ByVal psReqTraNro As String = "") As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        If TpoCarga = ReqDetUnRegistroTramiteUlt Then
            sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       ltrim(c.cConsDescripcion) + space(40) + rd.cLogReqDetRefMoneda, " & _
                "       CASE WHEN rd.nLogReqDetRefPrecio = 0 THEN Null ELSE rd.nLogReqDetRefPrecio END nLogReqDetRefPrecio " & _
                " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante C ON rd.cLogReqDetRefMoneda = c.cConsValor And c.cConsCod = " & gMoneda & " " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.cConsValor And cu.cConsCod = " & gUnidadMedida & " " & _
                " WHERE  rd.cLogReqNro = '" & psReqNro & "' AND rd.cLogReqTraNro = " & _
                "       (SELECT MAX(cLogReqTraNro) FROM LogReqDetalle WHERE cLogReqNro = '" & psReqNro & "' )"
        ElseIf TpoCarga = ReqDetUnRegistroTramite Then
            sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       ltrim(c.cConsDescripcion) + space(40) + rd.cLogReqDetRefMoneda, " & _
                "       CASE WHEN rd.nLogReqDetRefPrecio = 0 THEN Null ELSE rd.nLogReqDetRefPrecio END nLogReqDetRefPrecio " & _
                " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante C ON rd.cLogReqDetRefMoneda = c.cConsValor And cConsCod = " & gMoneda & " " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.cConsValor And cu.cConsCod = " & gUnidadMedida & " " & _
                " WHERE  rd.cLogReqNro = '" & psReqNro & "' " & _
                "       AND rd.cLogReqTraNro = '" & psReqTraNro & "' "
        ElseIf TpoCarga = ReqDetUnRegistroPrecio Then
            sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       ltrim(c.cConsDescripcion) + space(40) + rd.cLogReqDetRefMoneda, " & _
                "       CASE WHEN rd.nLogReqDetRefPrecio = 0 THEN Null ELSE rd.nLogReqDetRefPrecio END nLogReqDetRefPrecio, " & _
                "       CASE WHEN rd.nLogReqDetPrecio = 0 THEN Null ELSE rd.nLogReqDetPrecio END nLogReqDetPrecio " & _
                " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante C ON rd.cLogReqDetRefMoneda = c.cConsValor And cConsCod = " & gMoneda & " " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.cConsValor And cu.cConsCod = " & gUnidadMedida & " " & _
                " WHERE  rd.cLogReqNro = '" & psReqNro & "' AND rd.cLogReqTraNro = " & _
                "       (SELECT MAX(cLogReqTraNro) FROM LogReqDetalle WHERE cLogReqNro = '" & psReqNro & "' )"
        ElseIf TpoCarga = ReqDetUnRegistroConsol Then
            'Para Mostrar en la CONSOLIDACION - UNO
            sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       SUM(ISNULL(nLogReqDetMesCantidad,0)) nLogReqDetMesCantidad, " & _
                "       rd.nLogReqDetPrecio nLogReqDetPrecio, " & _
                "       (SUM(ISNULL(nLogReqDetMesCantidad,0)) * rd.nLogReqDetPrecio) SubTotal " & _
                " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.cConsValor And cu.cConsCod = " & gUnidadMedida & " " & _
                "       LEFT JOIN LogReqDetMes RDM ON rdm.cLogReqNro = rd.cLogReqNro " & _
                "       AND rdm.cLogReqTraNro = rd.cLogReqTraNro AND rdm.cBSCod = rd.cBSCod" & _
                " WHERE  rd.cLogReqNro ='" & psReqNro & "' " & _
                "       AND rd.cLogReqTraNro = (SELECT MAX(cLogReqTraNro) " & _
                "                               FROM LogReqDetalle " & _
                "                               WHERE cLogReqNro = rd.cLogReqNro )" & _
                " GROUP BY rd.cBSCod, bs.cBSDescripcion, nBSUnidad, cu.cConsDescripcion, rd.nLogReqDetPrecio " & _
                " ORDER BY rd.cBSCod "
        ElseIf TpoCarga = ReqDetTodosFlex Then
            'Para Mostrar en la CONSOLIDACION - TODOS
            sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       SUM(ISNULL(nLogReqDetMesCantidad,0)) nLogReqDetMesCantidad, " & _
                "       AVG(DISTINCT rd.nLogReqDetPrecio) nLogReqDetPrecio, " & _
                "       (SUM(ISNULL(nLogReqDetMesCantidad,0)) * AVG(DISTINCT rd.nLogReqDetPrecio)) SubTotal " & _
                " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.cConsValor And cu.cConsCod = " & gUnidadMedida & " " & _
                "       LEFT JOIN LogReqDetMes RDM ON rdm.cLogReqNro = rd.cLogReqNro " & _
                "       AND rdm.cLogReqTraNro = rd.cLogReqTraNro AND rdm.cBSCod = rd.cBSCod" & _
                " WHERE  rd.cLogReqNro IN ('" & psReqNro & "') " & _
                "       AND rd.cLogReqTraNro = (SELECT MAX(cLogReqTraNro) " & _
                "                               FROM LogReqDetalle " & _
                "                               WHERE cLogReqNro = rd.cLogReqNro )" & _
                " GROUP BY rd.cBSCod, bs.cBSDescripcion, nBSUnidad, cu.cConsDescripcion  " & _
                " ORDER BY rd.cBSCod "
        Else
            sSql = ""
        End If
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaReqDetalle = Rs
End Function

'Devuelve un RecordSet con los ReqDetMes
Public Function CargaReqDetMes(ByVal psReqNro As String, ByVal psReqTraNro As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        If psReqTraNro = "" Then
            sSql = "SELECT SUM(CASE WHEN c.cConsValor = " & gMesEnero & " THEN r.nLogReqDetMesCantidad ELSE Null END) Enero ," & _
                     " SUM(CASE WHEN c.cConsValor = " & gMesFebrero & " THEN r.nLogReqDetMesCantidad ELSE Null END) Febrero ," & _
                     " SUM(CASE WHEN c.cConsValor = " & gMesMarzo & " THEN r.nLogReqDetMesCantidad ELSE Null END) Marzo ," & _
                     " SUM(CASE WHEN c.cConsValor = " & gMesAbril & " THEN r.nLogReqDetMesCantidad ELSE Null END) Abril ," & _
                     " SUM(CASE WHEN c.cConsValor = " & gMesMayo & " THEN r.nLogReqDetMesCantidad ELSE Null END) Mayo ," & _
                     " SUM(CASE WHEN c.cConsValor = " & gMesJunio & " THEN r.nLogReqDetMesCantidad ELSE Null END) Junio , " & _
                     " SUM(CASE WHEN c.cConsValor = " & gMesJulio & " THEN r.nLogReqDetMesCantidad ELSE Null END) Julio , " & _
                     " SUM(CASE WHEN c.cConsValor = " & gMesAgosto & " THEN r.nLogReqDetMesCantidad ELSE Null END) Agosto , " & _
                     " SUM(CASE WHEN c.cConsValor = " & gMesSeptiembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Setiembre , " & _
                     " SUM(CASE WHEN c.cConsValor = " & gMesOctubre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Octubre , " & _
                     " SUM(CASE WHEN c.cConsValor = " & gMesNoviembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Noviembre , " & _
                     " SUM(CASE WHEN c.cConsValor = " & gMesDiciembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Diciembre " & _
                " FROM LogReqDetMes R JOIN Constante C ON r.nLogReqDetMesNro = c.cConsValor " & _
                " WHERE cLogReqNro = '" & psReqNro & "'" & _
                " AND cLogReqTraNro = (SELECT MAX(cLogReqTraNro) FROM LogReqDetalle WHERE cLogReqNro = '" & psReqNro & "' )" & _
                " AND c.cConsCod = " & gMeses & " " & _
                " GROUP BY r.cBSCod "
        Else
            sSql = "SELECT SUM(CASE WHEN c.cConsValor = " & gMesEnero & " THEN r.nLogReqDetMesCantidad ELSE Null END) Enero ," & _
                 " SUM(CASE WHEN c.cConsValor = " & gMesFebrero & " THEN r.nLogReqDetMesCantidad ELSE Null END) Febrero ," & _
                 " SUM(CASE WHEN c.cConsValor = " & gMesMarzo & " THEN r.nLogReqDetMesCantidad ELSE Null END) Marzo ," & _
                 " SUM(CASE WHEN c.cConsValor = " & gMesAbril & " THEN r.nLogReqDetMesCantidad ELSE Null END) Abril ," & _
                 " SUM(CASE WHEN c.cConsValor = " & gMesMayo & " THEN r.nLogReqDetMesCantidad ELSE Null END) Mayo ," & _
                 " SUM(CASE WHEN c.cConsValor = " & gMesJunio & " THEN r.nLogReqDetMesCantidad ELSE Null END) Junio , " & _
                 " SUM(CASE WHEN c.cConsValor = " & gMesJulio & " THEN r.nLogReqDetMesCantidad ELSE Null END) Julio , " & _
                 " SUM(CASE WHEN c.cConsValor = " & gMesAgosto & " THEN r.nLogReqDetMesCantidad ELSE Null END) Agosto , " & _
                 " SUM(CASE WHEN c.cConsValor = " & gMesSeptiembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Setiembre , " & _
                 " SUM(CASE WHEN c.cConsValor = " & gMesOctubre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Octubre , " & _
                 " SUM(CASE WHEN c.cConsValor = " & gMesNoviembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Noviembre , " & _
                 " SUM(CASE WHEN c.cConsValor = " & gMesDiciembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Diciembre " & _
            " FROM LogReqDetMes R JOIN Constante C ON r.nLogReqDetMesNro = c.cConsValor " & _
            " WHERE cLogReqNro = '" & psReqNro & "' " & _
            " AND cLogReqTraNro = '" & psReqTraNro & "' " & _
            " AND c.cConsCod = " & gMeses & " " & _
            " GROUP BY r.cBSCod "
        End If
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaReqDetMes = Rs
End Function

'Devuelve un RecordSet con Obtencion
Public Function CargaObtencion(ByVal psObtTpo As String, ByVal psPeriodo As String, _
ByVal psEstado As LogObtEstado) As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
     
    If oConec.AbreConexion(vsConexion) Then
        
        sSql = "" '"SELECT cLogReqNro " & _
            " FROM LogRequerimiento " & _
            " WHERE cLogReqTpo = '3' AND cLogReqPeriodo = '" & psPeriodo & "' " & _
            "       AND cLogReqEstado = '" & psEstado & "'"
        
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaObtencion = Rs
End Function


'Devuelve un RecordSet con ObtDetalle
Public Function CargaObtDetalle(ByVal psTpoCarga As TpoCargaObtDet, ByVal psObtNro As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        If psTpoCarga = ObtDetParaCtaCnt Then
            sSql = "SELECT od.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       od.nLogObtCantidad " & _
                " FROM LogObtDetalle OD LEFT JOIN BienesServicios BS ON od.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.cConsValor And cu.cConsCod = " & gUnidadMedida & " " & _
                " WHERE od.cLogObtNro = '" & psObtNro & "' "
        ElseIf psTpoCarga = ObtDetParaAprobar Then
            sSql = "SELECT od.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       od.nLogObtCantidad, od.nLogObtPrecio,  od.cCtaContcod " & _
                " FROM LogObtDetalle OD LEFT JOIN BienesServicios BS ON od.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.cConsValor And cu.cConsCod = " & gUnidadMedida & " " & _
                " WHERE od.cLogObtNro = '" & psObtNro & "' "
        ElseIf psTpoCarga = ObtDetParaAdquisi Then
            sSql = "SELECT od.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       od.nLogObtCantidad, od.nLogObtPrecio, " & _
                "       (od.nLogObtCantidad * od.nLogObtPrecio) SubTotal" & _
                " FROM LogObtDetalle OD LEFT JOIN BienesServicios BS ON od.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.cConsValor And cu.cConsCod = " & gUnidadMedida & " " & _
                " WHERE od.cLogObtNro = '" & psObtNro & "' "
        ElseIf psTpoCarga = ObtDetParaAdquisiConsol Then
            sSql = "SELECT od.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       SUM(IsNull(od.nLogObtCantidad,0)), " & _
                "       AVG(DISTINCT od.nLogObtPrecio) nLogObtPrecio, " & _
                "       (SUM(IsNull(od.nLogObtCantidad,0)) * AVG(DISTINCT od.nLogObtPrecio)) SubTotal" & _
                " FROM LogObtDetalle OD LEFT JOIN BienesServicios BS ON od.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.cConsValor And cu.cConsCod = " & gUnidadMedida & " " & _
                " WHERE od.cLogObtNro IN ('" & psObtNro & "') " & _
                " GROUP BY od.cBSCod, bs.cBSDescripcion, bs.nBSUnidad, cu.cConsDescripcion " & _
                " ORDER BY od.cBSCod "
        Else
            sSql = ""
        End If
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaObtDetalle = Rs
End Function

