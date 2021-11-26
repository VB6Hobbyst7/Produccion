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
    ReqUnRegistro = 1
    ReqUnRegistroTramite = 2
    ReqTodosAreaFlex = 3
    ReqTodosFlexConsol = 4
    ReqTodosAreaTraNuevo = 5
    ReqTodosAreaTraIngreso = 6
    ReqTodosAreaTraEgreso = 7
    ReqTodosAreaTraEgresoOtro = 8
    ReqTodosTraPrecio = 9
    ReqTodosTraCuenta = 10
    ReqTodosTraAproba = 11
    ReqTodosObten = 12
    ReqTodosFlexApro = 13
    ReqTodosFlexConsul = 14
    ReqTodosFlexSeleccion = 15
End Enum

Public Enum TpoCargaReqDet
    ReqDetUnRegistroPresupu = 1
    ReqDetUnRegistroTramite = 2
    ReqDetUnRegistroTramiteUlt = 3
    ReqDetUnRegistroPrecio = 4
    ReqDetUnRegistroConsol = 5
    ReqDetTodosConsol = 6
    ReqDetUnRegistroConsul = 7
    ReqDetTodosSelec = 8
End Enum

Public Enum TpoCargaReqDetMes
    ReqDetMesUnRegistro = 1
    ReqDetMesTraNro = 2
    ReqDetMesUltTraNro = 3
    ReqDetMesTodosConsol = 4
End Enum

Public Enum TpoCargaReqTra
    ReqTraUnRegistro = 0
    ReqTraTodosAreaMenosPrimero = 1
    ReqTraTodosArea = 2
    ReqTraTodosAreaMasDes = 3
End Enum

Public Enum TpoReqEstado
    ReqEstadoSolicitado = 1
    ReqEstadoaprobado = 2
    ReqEstadoObservado = 3
End Enum

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    Dim oIni As ClasIni
    
    Set oIni = New ClasIni
        vsConexion = oIni.CadenaConexion
        vsCentralPer = oIni.BasePersonas
        vsCentralCom = oIni.BaseComunes
        vsCentralImg = oIni.BaseImagenes
    Set oIni = Nothing
End Sub

'Devuelve un RecordSet con los requerimientos
Public Function CargaRequerimiento(ByVal pnTpoReq As Integer, ByVal psTpoCarga As TpoCargaReq, ByVal psAreaCod As String, _
Optional ByVal pnReqNro As Integer = 0, Optional ByVal pnPeriodo As Integer = 0) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If psTpoCarga = ReqUnRegistro Then
            sSql = "SELECT r.nLogReqNro, r.cLogReqNecesidad, r.cLogReqRequerimiento, " & _
                "       r.nLogReqTpo, (cm.cConsDescripcion + space(40) + Convert(char(1),rt.nLogReqTraMoneda)) cLogReqMoneda, " & _
                "       r.nLogReqPeriodo, rt.cAreaCod, a.cAreaDescripcion " & _
                " FROM  LogRequerimiento R INNER JOIN LogReqTramite RT ON r.nLogReqNro = rt.nLogReqNro " & _
                "       LEFT JOIN Areas A ON rt.cAreaCod = a.cAreaCod " & _
                "       LEFT JOIN Constante CM ON rt.nLogReqTraMoneda = cm.nConsValor AND cm.nConsCod = " & gMoneda & " " & _
                " WHERE r.nLogReqNro = " & pnReqNro & " " & _
                "       AND rt.nLogReqNro = rt.nLogReqTraNro " & _
                "       AND r.nLogReqTpo = " & pnTpoReq & " "
        ElseIf psTpoCarga = ReqUnRegistroTramite Then
            'Un Registro(con datos del ultimo nLogReqTraNro enviado)
            sSql = "SELECT r.nLogReqNro, r.cLogReqNecesidad, r.cLogReqRequerimiento, " & _
                "       r.nLogReqTpo, (cm.cConsDescripcion + space(40) + Convert(char(1),rt.nLogReqTraMoneda)) cLogReqMoneda, " & _
                "       r.nLogReqPeriodo, rt.nLogReqTraNro " & _
                " FROM  LogRequerimiento R INNER JOIN LogReqTramite RT ON r.nLogReqNro = rt.nLogReqNro " & _
                "       LEFT JOIN Constante CM ON rt.nLogReqTraMoneda = cm.nConsValor AND cm.nConsCod = " & gMoneda & " " & _
                " WHERE r.nLogReqNro = " & pnReqNro & " " & _
                "       AND rt.nLogReqTraNro = " & _
                "           (SELECT MAX(nLogReqTraNro) FROM LogReqTramite " & _
                "            WHERE nLogReqNro = " & pnReqNro & " " & _
                "               AND nLogReqTraEstado <> " & gLogReqEstadoInicio & ")" & _
                "       AND r.nLogReqTpo = " & pnTpoReq & " "
        ElseIf psTpoCarga = ReqTodosAreaFlex Then
            'Todos los registros de un Area, para Flex (solo para edición de NUEVOS)
            sSql = "SELECT m.cMovNro Código, substring(r.cLogReqNecesidad,1,50) Necesidad, " & _
                "       Substring(r.cLogReqRequerimiento,1,50) Requerimiento " & _
                " FROM Mov M JOIN LogRequerimiento R ON m.nMovNro = r.nLogReqNro JOIN LogReqTramite RT ON r.nLogReqNro = rt.nLogReqNro " & _
                "       AND r.nLogReqNro = rt.nLogReqTraNro " & _
                " WHERE rt.cAreaCod = '" & psAreaCod & "' AND rt.cAreaCodDes = ''" & _
                "       AND rt.nLogReqTraEstado = " & gLogReqEstadoInicio & " " & _
                "       AND rt.nLogReqNro = (SELECT MAX(nLogReqTraNro) FROM LogReqTramite " & _
                "                       WHERE nLogReqNro = r.nLogReqNro) " & _
                "       AND r.nLogReqTpo = " & pnTpoReq & " "
        ElseIf psTpoCarga = ReqTodosFlexConsol Then
            'Todos los registros (Todas las areas para CONSOLIDACION) - nLogReqNro
            sSql = "SELECT m.cMovNro , " & _
                "       (SELECT b.cAreaDescripcion FROM LogReqTramite A JOIN Areas B " & _
                "               On a.cAreaCod = b.cAreaCod WHERE a.nLogReqNro = a.nLogReqTraNro " & _
                "               And a.nLogReqNro = r.nLogReqNro ) Procedencia, " & _
                "       r.nLogReqPeriodo, (cm.cConsDescripcion + space(40) + Convert(char(1),rt.nLogReqTraMoneda)) Moneda  , c.cConsDescripcion " & _
                " FROM Mov M JOIN LogRequerimiento R ON m.nMovNro = r.nLogReqNro JOIN LogReqTramite RT ON r.nLogReqNro = rt.nLogReqNro " & _
                "       AND rt.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
                "                               FROM LogReqTramite " & _
                "                               WHERE nLogReqNro = r.nLogReqNro) " & _
                "       LEFT JOIN Constante C ON r.nLogReqTpo = c.nConsValor AND c.nConsCod = " & gLogReqTipo & " " & _
                "       LEFT JOIN Constante CM ON rt.nLogReqTraMoneda = cm.nConsValor AND cm.nConsCod = " & gMoneda & " " & _
                " WHERE  (r.nLogReqTpo = " & gLogReqTipoNormal & " AND rt.nLogReqTraEstado = " & gLogReqEstadoCuenta & ") " & _
                "       OR (r.nLogReqTpo = " & gLogReqTipoExtemporaneo & " AND rt.nLogReqTraEstado = " & gLogReqEstadoAceptado & ") " & _
                "       AND  rt.nLogReqTraEstado <> " & gLogReqEstadoRechazado & " "
                
        ElseIf psTpoCarga = ReqTodosAreaTraNuevo Then
            'Todos los registros de un Area, NUEVOS
            sSql = "SELECT m.cMovNro Código, a.cAreaDescripcion, r.nLogReqPeriodo, " & _
                "       substring(r.cLogReqNecesidad,1,50) Necesidad, " & _
                "       substring(r.cLogReqRequerimiento,1,50) Requerimiento, " & _
                "       (SELECT c.cConsDescripcion FROM LogReqTramite RT JOIN Constante C " & _
                "        ON rt.nLogReqTraEstado = c.nConsValor AND c.nConsCod = " & gLogReqEstado & " " & _
                "        AND rt.nLogReqTraNro = r.nLogReqNro ) Estado " & _
                " FROM Mov M JOIN LogRequerimiento R ON m.nMovNro = r.nLogReqNro JOIN LogReqTramite RT ON r.nLogReqNro = rt.nLogReqNro " & _
                "       LEFT JOIN Areas A ON rt.cAreaCod = a.cAreaCod " & _
                " WHERE rt.cAreaCod = '" & psAreaCod & "' AND rt.cAreaCodDes = '' AND " & _
                "       rt.nLogReqTraEstado = " & gLogReqEstadoInicio & " AND" & _
                "       rt.nLogReqNro = (SELECT MAX(nLogReqTraNro) FROM LogReqTramite " & _
                "                       WHERE nLogReqNro = r.nLogReqNro) " & _
                "       AND r.nLogReqTpo = " & pnTpoReq & " "
        ElseIf psTpoCarga = ReqTodosAreaTraIngreso Then
            'Todos los registros de un Area, para Flex - Tramite RECEPCIONADOS
            sSql = " SELECT  m.cMovNro Código, " & _
                "       (SELECT b.cAreaDescripcion FROM LogReqTramite A JOIN Areas B " & _
                "               On a.cAreaCod = b.cAreaCod WHERE a.nLogReqNro = a.nLogReqTraNro " & _
                "               And a.nLogReqNro = r.nLogReqNro ) Procedencia,  r.nLogReqPeriodo, " & _
                "       substring(r.cLogReqNecesidad, 1, 50) Necesidad, " & _
                "       substring(r.cLogReqRequerimiento, 1, 50) Requerimiento, " & _
                "       c.cConsDescripcion " & _
                " FROM Mov M JOIN LogRequerimiento R ON m.nMovNro = r.nLogReqNro JOIN LogReqTramite RT ON r.nLogReqNro = rt.nLogReqNro " & _
                "       LEFT JOIN Constante C ON rt.nLogReqTraEstado = c.nConsValor  " & _
                "               AND c.nConsCod = " & gLogReqEstado & " AND c.nConsValor <> " & gLogReqEstado & " " & _
                " WHERE rt.cAreaCodDes = '" & psAreaCod & "' AND " & _
                "       rt.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) FROM LogReqTramite " & _
                "                       WHERE nLogReqNro = r.nLogReqNro) " & _
                "       AND r.nLogReqTpo = " & pnTpoReq & " "
        ElseIf psTpoCarga = ReqTodosAreaTraEgreso Then
            'Todos los registros de un Area, para Flex - Tramite TRAMITADOS
            sSql = "SELECT m.cMovNro Código, a.cAreaDescripcion,  r.nLogReqPeriodo," & _
                "       substring(r.cLogReqNecesidad,1,50) Necesidad," & _
                "       substring(r.cLogReqRequerimiento,1,50) Requerimiento, " & _
                "       (SELECT c.cConsDescripcion " & _
                "        FROM LogReqTramite RT JOIN Constante C " & _
                "        ON rt.nLogReqTraEstado = c.nConsValor AND c.nConsCod = " & gLogReqEstado & " " & _
                "        AND rt.nLogReqNro = r.nLogReqNro AND rt.nLogReqTraNro = " & _
                "                   (SELECT MAX(nLogReqTraNro) FROM LogReqTramite " & _
                "                   WHERE nLogReqNro = r.nLogReqNro )) Estado " & _
                " FROM Mov M JOIN LogRequerimiento R ON m.nMovNro = r.nLogReqNro JOIN LogReqTramite RT ON r.nLogReqNro = rt.nLogReqNro " & _
                "                AND rt.nLogReqNro = rt.nLogReqTraNro  " & _
                "       LEFT JOIN Areas A ON rt.cAreaCod = a.cAreaCod " & _
                " WHERE rt.cAreaCod = '" & psAreaCod & "' AND " & _
                "       (rt.cAreaCodDes <> '' OR rt.nLogReqTraEstado = " & gLogReqEstadoRechazado & " " & _
                "        OR (SELECT COUNT(*) FROM LogReqTramite WHERE nLogReqNro = r.nLogReqNro) > 1 )" & _
                "       AND r.nLogReqTpo = " & pnTpoReq & " "
        ElseIf psTpoCarga = ReqTodosAreaTraEgresoOtro Then
            'Todos los registros de un Area, para Flex - Tramite TRAMITADOS OTROS
            sSql = "SELECT m.cMovNro Código," & _
                "       (SELECT b.cAreaDescripcion " & _
                "        FROM LogReqTramite A LEFT JOIN Areas B ON a.cAreaCod = b.cAreaCod " & _
                "        WHERE a.nLogReqNro = a.nLogReqTraNro " & _
                "        AND a.nLogReqNro = r.nLogReqNro ) Procedencia,  " & _
                "       r.nLogReqPeriodo, " & _
                "       substring(r.cLogReqNecesidad,1,50) Necesidad," & _
                "       substring(r.cLogReqRequerimiento,1,50) Requerimiento, " & _
                "       (SELECT c.cConsDescripcion " & _
                "        FROM LogReqTramite RT JOIN Constante C " & _
                "        ON rt.nLogReqTraEstado = c.nConsValor AND c.nConsCod = " & gLogReqEstado & " " & _
                "        AND rt.nLogReqNro = r.nLogReqNro AND rt.nLogReqTraNro = " & _
                "                   (SELECT MAX(nLogReqTraNro) FROM LogReqTramite " & _
                "                   WHERE nLogReqNro = r.nLogReqNro )) Estado " & _
                " FROM Mov M JOIN LogRequerimiento R ON m.nMovNro = r.nLogReqNro JOIN LogReqTramite RT ON r.nLogReqNro = rt.nLogReqNro " & _
                "               AND NOT rt.nLogReqTraEstado IN ('" & gLogReqEstadoCuenta & "','" & gLogReqEstadoPrecio & "','" & gLogReqEstadoConsolida & "') " & _
                "       LEFT JOIN Areas A ON rt.cAreaCod = a.cAreaCod " & _
                " WHERE rt.cAreaCod = '" & psAreaCod & "' AND " & _
                "       (LEN(rt.cAreaCodDes) > 0  " & _
                "        OR (SELECT COUNT(*) FROM LogReqTramite WHERE nLogReqNro = r.nLogReqNro) > 1 )" & _
                "       AND r.nLogReqTpo = " & pnTpoReq & " "
                
            sSql = sSql & " AND '" & psAreaCod & "' <> (SELECT a.cAreaCod " & _
                "        FROM LogReqTramite A LEFT JOIN Areas B ON a.cAreaCod = b.cAreaCod " & _
                "        WHERE a.nLogReqNro = a.nLogReqTraNro " & _
                "        AND a.nLogReqNro = r.nLogReqNro ) "
                '  AND rt.nLogReqNro = rt.nLogReqTraNro
        ElseIf psTpoCarga = ReqTodosTraPrecio Then
            'TODOS los registros de todas las areas para Ingresar PRECIO
            sSql = "SELECT m.cMovNro Código, " & _
                "       (SELECT b.cAreaDescripcion " & _
                "        FROM LogReqTramite A LEFT JOIN Areas B ON a.cAreaCod = b.cAreaCod " & _
                "        WHERE a.nLogReqNro = a.nLogReqTraNro " & _
                "        AND a.nLogReqNro = r.nLogReqNro ) Procedencia, r.nLogReqPeriodo Periodo, " & _
                "       SUBSTRING(r.cLogReqNecesidad,1,50) Necesidad, " & _
                "       SUBSTRING(r.cLogReqRequerimiento,1,50) Requerimiento, " & _
                "       c.cConsDescripcion Estado " & _
                " FROM Mov M JOIN LogRequerimiento R ON m.nMovNro = r.nLogReqNro JOIN LogReqTramite RT ON r.nLogReqNro = rt.nLogReqNro " & _
                "       LEFT JOIN Constante C ON rt.nLogReqTraEstado = c.nConsValor AND c.nConsCod = " & gLogReqEstado & " " & _
                " WHERE rt.nLogReqTraEstado = '" & gLogReqEstadoAcepPrevio & "'" & _
                "       AND rt.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) FROM LogReqTramite " & _
                "                       WHERE nLogReqNro = r.nLogReqNro) " & _
                "       AND r.nLogReqTpo = " & pnTpoReq & ""
        ElseIf psTpoCarga = ReqTodosTraCuenta Then
            'TODOS los registros de todas las areas para Ingresar CUENTAS
            sSql = "SELECT m.cMovNro Código, " & _
                "       (SELECT b.cAreaDescripcion " & _
                "        FROM LogReqTramite A LEFT JOIN Areas B ON a.cAreaCod = b.cAreaCod " & _
                "        WHERE a.nLogReqNro = a.nLogReqTraNro " & _
                "        AND a.nLogReqNro = r.nLogReqNro ) Procedencia, r.nLogReqPeriodo Periodo, " & _
                "       SUBSTRING(r.cLogReqNecesidad,1,50) Necesidad, " & _
                "       SUBSTRING(r.cLogReqRequerimiento,1,50) Requerimiento, " & _
                "       c.cConsDescripcion Estado " & _
                " FROM Mov M JOIN LogRequerimiento R ON m.nMovNro = r.nLogReqNro JOIN LogReqTramite RT ON r.nLogReqNro = rt.nLogReqNro " & _
                "       LEFT JOIN Constante C ON rt.nLogReqTraEstado = c.nConsValor AND c.nConsCod = " & gLogReqEstado & " " & _
                " WHERE rt.nLogReqTraEstado = '" & IIf(pnTpoReq = "1", gLogReqEstadoAceptado, gLogReqEstadoPrecio) & "'" & _
                "       AND rt.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) FROM LogReqTramite " & _
                "                       WHERE nLogReqNro = r.nLogReqNro) " & _
                "       AND r.nLogReqTpo = " & pnTpoReq & ""
        ElseIf psTpoCarga = ReqTodosTraAproba Then
            'TODOS los registros de todas las areas para APROBAR O RECHAZAR
            sSql = "SELECT m.cMovNro Código, " & _
                "       (SELECT b.cAreaDescripcion " & _
                "        FROM LogReqTramite A LEFT JOIN Areas B ON a.cAreaCod = b.cAreaCod " & _
                "        WHERE a.nLogReqNro = a.nLogReqTraNro " & _
                "        AND a.nLogReqNro = r.nLogReqNro ) Procedencia, r.nLogReqPeriodo Periodo, " & _
                "       SUBSTRING(r.cLogReqNecesidad,1,50) Necesidad, " & _
                "       SUBSTRING(r.cLogReqRequerimiento,1,50) Requerimiento, " & _
                "       c.cConsDescripcion Estado " & _
                " FROM Mov M JOIN LogRequerimiento R ON m.nMovNro = r.nLogReqNro JOIN LogReqTramite RT ON r.nLogReqNro = rt.nLogReqNro " & _
                "       LEFT JOIN Constante C ON rt.nLogReqTraEstado = c.nConsValor AND c.nConsCod = " & gLogReqEstado & " " & _
                " WHERE rt.nLogReqTraEstado = '" & IIf(pnTpoReq = "1", gLogReqEstadoPrecio, gLogReqEstadoCuenta) & "'" & _
                "       AND rt.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) FROM LogReqTramite " & _
                "                       WHERE nLogReqNro = r.nLogReqNro) " & _
                "       AND r.nLogReqTpo = " & pnTpoReq & ""
        ElseIf psTpoCarga = ReqTodosObten Then
            sSql = "SELECT a.cAreaDescripcion, ro.nLogReqNro " & _
                " FROM LogReqObt RO JOIN LogReqTramite RT ON ro.nLogReqNro = rt.nLogReqNro  " & _
                "       AND rt.nLogReqNro = rt.nLogReqTraNro " & _
                "       LEFT JOIN Areas A ON rt.cAreaCod = a.cAreaCod " & _
                " WHERE cLogObtNro = " & pnReqNro & ""
        ElseIf psTpoCarga = ReqTodosFlexApro Then
            'Todos los registros (Todas las areas para Aprobación)
            sSql = "SELECT (SELECT cAreaCod " & _
                "           FROM LogRequerimiento M JOIN LogReqTramite N ON m.nLogReqNro = n.nLogReqNro " & _
                "                   AND n.nLogReqNro = n.nLogReqTraNro AND m.nLogReqNro = r.nLogReqNro) cAreaCod, " & _
                "          (SELECT q.cAreaDescripcion " & _
                "           FROM LogRequerimiento O JOIN LogReqTramite P ON o.nLogReqNro = p.nLogReqNro " & _
                "                   AND p.nLogReqNro = p.nLogReqTraNro AND o.nLogReqNro = r.nLogReqNro " & _
                "               LEFT JOIN Areas Q ON p.cAreaCod = q.cAreaCod) Area, " & _
                "       r.nLogReqNro Requerimiento,  " & _
                "       r.cLogReqNecesidad, r.cLogReqRequerimiento " & _
                " FROM LogRequerimiento R JOIN LogReqTramite RT ON r.nLogReqNro = rt.nLogReqNro " & _
                "       AND rt.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
                "                               FROM LogReqTramite " & _
                "                               WHERE nLogReqNro = r.nLogReqNro) " & _
                " WHERE rt.nLogReqTraEstado = '" & gLogReqEstadoPrecio & "' " & _
                "       AND r.nLogReqPeriodo = '" & pnPeriodo & "' " & _
                "       AND r.nLogReqTpo = " & pnTpoReq & ""
        ElseIf psTpoCarga = ReqTodosFlexConsul Then
            'Todos los registros (Todas las areas para CONSULTAS)
            sSql = "SELECT m.cMovNro , " & _
                "       (SELECT b.cAreaDescripcion FROM LogReqTramite A JOIN Areas B " & _
                "               On a.cAreaCod = b.cAreaCod WHERE a.nLogReqNro = a.nLogReqTraNro " & _
                "               And a.nLogReqNro = r.nLogReqNro ) Procedencia, " & _
                "       r.nLogReqPeriodo, (cm.cConsDescripcion + space(40) + Convert(char(1),rt.nLogReqTraMoneda)) Moneda, " & _
                "       c.cConsDescripcion Estado  " & _
                " FROM Mov M JOIN LogRequerimiento R ON m.nMovNro = r.nLogReqNro JOIN LogReqTramite RT ON r.nLogReqNro = rt.nLogReqNro " & _
                "       AND rt.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
                "                               FROM LogReqTramite " & _
                "                               WHERE nLogReqNro = r.nLogReqNro) " & _
                "       LEFT JOIN Constante C ON rt.nLogReqTraEstado = c.nConsValor AND c.nConsCod = " & gLogReqEstado & " " & _
                "       LEFT JOIN Constante CM ON rt.nLogReqTraMoneda = cm.nConsValor AND cm.nConsCod = " & gMoneda & " " & _
                " WHERE  r.nLogReqTpo = " & pnTpoReq & ""
        ElseIf psTpoCarga = ReqTodosFlexSeleccion Then
            'Todos los registros para proceso de Selección
            sSql = "SELECT DISTINCT m.cMovNro ,b.cAreaDescripcion Procedencia, r.nLogReqPeriodo, " & _
                "       (cm.cConsDescripcion + space(40) + Convert(char(1),rt.nLogReqTraMoneda)) Moneda, " & _
                "       c.cConsDescripcion Estado  " & _
                " FROM Mov M JOIN LogRequerimiento R ON m.nMovNro = r.nLogReqNro " & _
                "       JOIN LogReqTramite RT ON r.nLogReqNro = rt.nLogReqNro " & _
                "           AND rt.nLogReqTraNro = r.nLogReqNro " & _
                "       JOIN LogReqDetMes RDM ON r.nLogReqNro = rdm.nLogReqNro " & _
                "           AND rdm.nLogReqTraNro = r.nLogReqNro " & _
                "       LEFT JOIN Areas B On rt.cAreaCod = b.cAreaCod " & _
                "       LEFT JOIN Constante C ON rt.nLogReqTraEstado = c.nConsValor AND c.nConsCod = " & gLogReqEstado & " " & _
                "       LEFT JOIN Constante CM ON rt.nLogReqTraMoneda = cm.nConsValor AND cm.nConsCod = " & gMoneda & " " & _
                " WHERE  r.nLogReqTpo = " & pnTpoReq & " AND rdm.nLogSelNro Is Null "
        Else
            sSql = ""
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaRequerimiento = rs
End Function

'Devuelve un RecordSet con los trámites
Public Function CargaReqTramite(ByVal psTpoCarga As TpoCargaReqTra, ByVal pnReqNro As Integer, _
ByVal pnReqTraNro As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If psTpoCarga = ReqTraUnRegistro Then
            'Un Registro
            sSql = " SELECT rt.nLogReqTraNro, rt.cAreaCod, a.cAreaDescripcion, " & _
                "       rt.cLogReqTraComentario " & _
                " FROM LogReqTramite RT LEFT JOIN AREAS A ON rt.cAreaCod = a.cAreaCod " & _
                " WHERE  rt.nLogReqNro = " & pnReqNro & " AND rt.nLogReqTraNro = " & pnReqTraNro & " "
        ElseIf psTpoCarga = ReqTraTodosAreaMenosPrimero Then
            'Todos los registros de un Area menos primero sin contestar
            sSql = " SELECT rt.nLogReqTraNro, a.cAreaDescripcion, rt.nLogReqTraEstado, c.cConsDescripcion, " & _
                "       rt.cLogReqTraComentario " & _
                " FROM LogReqTramite RT LEFT JOIN AREAS A ON rt.cAreaCod = a.cAreaCod " & _
                "       LEFT JOIN Constante C ON rt.nLogReqTraEstado = c.nConsValor " & _
                "           AND c.nConsCod = " & gLogReqEstado & " " & _
                " WHERE  rt.nLogReqNro = " & pnReqNro & " AND (SELECT COUNT(*) FROM LogReqTramite WHERE nLogReqNro = 'rt.nLogReqTraNro') > 1 " & _
                " ORDER BY rt.nLogReqTraNro "
        ElseIf psTpoCarga = ReqTraTodosArea Then
            'Todos los registros de un Area
            sSql = " SELECT rt.nLogReqTraNro, a.cAreaDescripcion, rt.nLogReqTraEstado, c.cConsDescripcion, " & _
                "       rt.cLogReqTraComentario " & _
                " FROM LogReqTramite RT LEFT JOIN AREAS A ON rt.cAreaCod = a.cAreaCod " & _
                "       LEFT JOIN Constante C ON rt.nLogReqTraEstado = c.nConsValor " & _
                "           AND c.nConsCod = " & gLogReqEstado & " " & _
                " WHERE  rt.nLogReqNro = " & pnReqNro & " " & _
                " ORDER BY rt.nLogReqTraNro "
        ElseIf psTpoCarga = ReqTraTodosAreaMasDes Then
            sSql = " SELECT rt.nLogReqTraNro, a.cAreaDescripcion, " & gLogReqEstadoParaTramite & " nLogReqTraEstado, c.cConsDescripcion, " & _
                "       '' cLogReqTraComentario " & _
                " FROM LogReqTramite RT LEFT JOIN AREAS A ON rt.cAreaCodDes = a.cAreaCod " & _
                "       LEFT JOIN Constante C ON " & gLogReqEstadoParaTramite & " = c.nConsValor " & _
                "           AND c.nConsCod = " & gLogReqEstado & " " & _
                " WHERE  rt.nLogReqNro = " & pnReqNro & " AND rt.cAreaCodDes <> '' " & _
                " AND rt.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) FROM LogReqTramite " & _
                "                       WHERE nLogReqNro = rt.nLogReqNro) "
        Else
            sSql = ""
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaReqTramite = rs
End Function


'Devuelve un RecordSet con los trámites
Public Function CuentaReqTramite(ByVal pnReqNro As Integer) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT count(*) " & _
            " FROM LogReqTramite RT " & _
            " WHERE  rt.nLogReqNro = " & pnReqNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    CuentaReqTramite = rs(0)
End Function

'Devuelve un RecordSet con los trámites
Public Function IsReqTramiteInicial(ByVal pnReqNro As Integer) As Boolean
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT rt.nLogReqNro  " & _
            " FROM LogReqTramite RT " & _
            " WHERE  rt.nLogReqNro = " & pnReqNro & " AND rt.nLogReqNro  = rt.nLogReqTraNro " & _
            "       AND rt.cAreaCodDes = '' "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    IsReqTramiteInicial = IIf(rs.RecordCount > 0, True, False)
End Function

'Devuelve un RecordSet con los trámites
Public Function CuentaReqDetalle(ByVal pnReqNro As Integer, ByVal pnReqTraNro As Integer) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT count(*) " & _
            " FROM LogReqDetalle RD " & _
            " WHERE  rd.nLogReqNro = " & pnReqNro & " AND rd.nLogReqTraNro = " & pnReqTraNro & ""
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    CuentaReqDetalle = rs(0)
End Function

'Devuelve un RecordSet con numero del ultimo tramite
Public Function CargaUltReqTraNro(ByVal pnReqNro As Integer) As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT MAX(rt.nLogReqTraNro) " & _
            " FROM LogReqTramite RT " & _
            " WHERE  rt.nLogReqNro = " & pnReqNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    CargaUltReqTraNro = rs(0)
End Function

'Devuelve un RecordSet con numero del ultimo detalle
Public Function CargaUltReqDetNro(ByVal pnReqNro As Integer) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT MAX(rd.nLogReqTraNro) " & _
            " FROM LogReqDetalle RD " & _
            " WHERE  rd.nLogReqNro = " & pnReqNro & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    CargaUltReqDetNro = rs(0)
End Function

'Devuelve un RecordSet con los ReqDetalle
Public Function CargaReqDetalle(ByVal TpoCarga As TpoCargaReqDet, ByVal pnReqNro As String, _
Optional ByVal pnReqTraNro As Integer = 0, Optional ByVal psBSCod As String = "", _
Optional ByVal pbASoles As Boolean = True, Optional ByVal pnTipCambio As Currency = 0) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If TpoCarga = ReqDetUnRegistroPresupu Then
            sSql = "SELECT IsNull(r.nLogReqTipCambio,0) nLogReqTipCambio, " & _
                "       rd.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       IsNull(rd.nLogReqDetPrecio,0) nLogReqDetPrecio, " & _
                "       IsNull(rd.nPresuCod,'') cPlaPreNro, " & _
                "       IsNull(rd.cPresuRubroCod,'') cPlaRubCod, IsNull(rd.cCtaContCod,'') cCtaContCod " & _
                " FROM LogRequerimiento R JOIN LogReqDetalle RD ON r.nLogReqNro = rd.nLogReqNro " & _
                "       LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                " WHERE  rd.nLogReqNro = " & pnReqNro & " AND rd.nLogReqTraNro = " & pnReqTraNro & "" & _
                "       AND rd.cBSCod = '" & psBSCod & "'"
        ElseIf TpoCarga = ReqDetUnRegistroTramiteUlt Then
            sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       CASE WHEN rd.nLogReqDetPrecio = 0 THEN Null ELSE rd.nLogReqDetPrecio END nLogReqDetPrecio " & _
                " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                " WHERE  rd.nLogReqNro = " & pnReqNro & " AND rd.nLogReqTraNro = " & _
                "       (SELECT MAX(nLogReqTraNro) FROM LogReqDetalle WHERE nLogReqNro = " & pnReqNro & " )"
        ElseIf TpoCarga = ReqDetUnRegistroTramite Then
            sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       CASE WHEN rd.nLogReqDetPrecio = 0 THEN Null ELSE rd.nLogReqDetPrecio END nLogReqDetPrecio " & _
                " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                " WHERE  rd.nLogReqNro = " & pnReqNro & " " & _
                "       AND rd.nLogReqTraNro = " & pnReqTraNro & " "
        ElseIf TpoCarga = ReqDetUnRegistroPrecio Then
            sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       ltrim(c.cConsDescripcion) + space(40) + rd.cLogReqDetRefMoneda, " & _
                "       CASE WHEN rd.nLogReqDetRefPrecio = 0 THEN Null ELSE rd.nLogReqDetRefPrecio END nLogReqDetRefPrecio, " & _
                "       CASE WHEN rd.nLogReqDetPrecio = 0 THEN Null ELSE rd.nLogReqDetPrecio END nLogReqDetPrecio " & _
                " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante C ON rd.cLogReqDetRefMoneda = c.nConsValor And nConsCod = " & gMoneda & " " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                " WHERE  rd.nLogReqNro = " & pnReqNro & " AND rd.nLogReqTraNro = " & _
                "       (SELECT MAX(nLogReqTraNro) FROM LogReqDetalle WHERE nLogReqNro = " & pnReqNro & " )"
        ElseIf TpoCarga = ReqDetUnRegistroConsul Then
            'Para Consultas
            'sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       CASE WHEN rd.nLogReqDetPrecio = 0 THEN Null ELSE rd.nLogReqDetPrecio END nLogReqDetPrecio " & _
                " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                " WHERE  rd.nLogReqNro = " & pnReqNro & " AND rd.nLogReqTraNro = " & _
                "       (SELECT MAX(nLogReqTraNro) FROM LogReqDetalle WHERE nLogReqNro = " & pnReqNro & " )"
        
        
            sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       SUM(ISNULL(nLogReqDetMesCantidad,0)) nLogReqDetMesCantidad, " & _
                "       ISNULL(rd.nLogReqDetPrecio,0) nLogReqDetPrecio," & _
                "       (SUM(ISNULL(nLogReqDetMesCantidad,0)) * ISNULL(rd.nLogReqDetPrecio,0)) SubTotal " & _
                " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                "       LEFT JOIN LogReqDetMes RDM ON rdm.nLogReqNro = rd.nLogReqNro " & _
                "           AND rdm.nLogReqTraNro = rd.nLogReqTraNro AND rdm.cBSCod = rd.cBSCod" & _
                " WHERE  rd.nLogReqNro ='" & pnReqNro & "' " & _
                "       AND rd.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
                "                               FROM LogReqDetalle " & _
                "                               WHERE nLogReqNro = rd.nLogReqNro )" & _
                " GROUP BY rd.cBSCod, bs.cBSDescripcion, nBSUnidad, cu.cConsDescripcion, " & _
                "       rd.nLogReqDetPrecio" & _
                " ORDER BY rd.cBSCod "
        
'                "       LEFT JOIN LogReqTramite RT ON rd.nLogReqTraNro = rt.nLogReqTraNro " & _
'                "       LEFT JOIN LogRequerimiento R ON rd.nLogReqNro = r.nLogReqNro " & _

        
        ElseIf TpoCarga = ReqDetUnRegistroConsol Then
            'Para Mostrar en la CONSOLIDACION - UNO
            If pbASoles Then
                'A SOLES
                If pnTipCambio = 0 Then
                    sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                        "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                        "       SUM(ISNULL(nLogReqDetMesCantidad,0)) nLogReqDetMesCantidad, " & _
                        "       (CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) * r.nLogReqTipCambio END) nLogReqDetPrecio," & _
                        "       (SUM(ISNULL(nLogReqDetMesCantidad,0)) * (CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) * r.nLogReqTipCambio END)) SubTotal " & _
                        " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                        "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                        "       LEFT JOIN LogReqDetMes RDM ON rdm.nLogReqNro = rd.nLogReqNro " & _
                        "       AND rdm.nLogReqTraNro = rd.nLogReqTraNro AND rdm.cBSCod = rd.cBSCod" & _
                        "       LEFT JOIN LogReqTramite RT ON rd.nLogReqTraNro = rt.nLogReqTraNro " & _
                        "       LEFT JOIN LogRequerimiento R ON rd.nLogReqNro = r.nLogReqNro " & _
                        " WHERE  rd.nLogReqNro ='" & pnReqNro & "' " & _
                        "       AND rd.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
                        "                               FROM LogReqDetalle " & _
                        "                               WHERE nLogReqNro = rd.nLogReqNro )" & _
                        " GROUP BY rd.cBSCod, bs.cBSDescripcion, nBSUnidad, cu.cConsDescripcion, " & _
                        "       rd.nLogReqDetPrecio, r.nLogReqTipCambio, rt.nLogReqTraMoneda " & _
                        " ORDER BY rd.cBSCod "
                Else
                    sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                        "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                        "       SUM(ISNULL(nLogReqDetMesCantidad,0)) nLogReqDetMesCantidad, " & _
                        "       (CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) * " & pnTipCambio & " END) nLogReqDetPrecio," & _
                        "       (SUM(ISNULL(nLogReqDetMesCantidad,0)) * (CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) * " & pnTipCambio & " END)) SubTotal " & _
                        " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                        "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                        "       LEFT JOIN LogReqDetMes RDM ON rdm.nLogReqNro = rd.nLogReqNro " & _
                        "       AND rdm.nLogReqTraNro = rd.nLogReqTraNro AND rdm.cBSCod = rd.cBSCod" & _
                        "       LEFT JOIN LogReqTramite RT ON rd.nLogReqTraNro = rt.nLogReqTraNro " & _
                        "       LEFT JOIN LogRequerimiento R ON rd.nLogReqNro = r.nLogReqNro " & _
                        " WHERE  rd.nLogReqNro ='" & pnReqNro & "' " & _
                        "       AND rd.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
                        "                               FROM LogReqDetalle " & _
                        "                               WHERE nLogReqNro = rd.nLogReqNro )" & _
                        " GROUP BY rd.cBSCod, bs.cBSDescripcion, nBSUnidad, cu.cConsDescripcion, " & _
                        "       rd.nLogReqDetPrecio, rt.nLogReqTraMoneda " & _
                        " ORDER BY rd.cBSCod "
                End If
            Else
                'A DOLARES
                sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                    "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                    "       SUM(ISNULL(nLogReqDetMesCantidad,0)) nLogReqDetMesCantidad, " & _
                    "       (CASE WHEN rt.nLogReqTraMoneda = " & gMonedaExtranjera & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) / " & pnTipCambio & " END) nLogReqDetPrecio," & _
                    "       (SUM(ISNULL(nLogReqDetMesCantidad,0)) * (CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) / " & pnTipCambio & " END)) SubTotal " & _
                    " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                    "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                    "       LEFT JOIN LogReqDetMes RDM ON rdm.nLogReqNro = rd.nLogReqNro " & _
                    "       AND rdm.nLogReqTraNro = rd.nLogReqTraNro AND rdm.cBSCod = rd.cBSCod" & _
                    "       LEFT JOIN LogReqTramite RT ON rd.nLogReqTraNro = rt.nLogReqTraNro " & _
                    "       LEFT JOIN LogRequerimiento R ON rd.nLogReqNro = r.nLogReqNro " & _
                    " WHERE  rd.nLogReqNro ='" & pnReqNro & "' " & _
                    "       AND rd.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
                    "                               FROM LogReqDetalle " & _
                    "                               WHERE nLogReqNro = rd.nLogReqNro )" & _
                    " GROUP BY rd.cBSCod, bs.cBSDescripcion, nBSUnidad, cu.cConsDescripcion, " & _
                    "       rd.nLogReqDetPrecio, rt.nLogReqTraMoneda " & _
                    " ORDER BY rd.cBSCod "
                
            End If
        ElseIf TpoCarga = ReqDetTodosConsol Then
            'Para Mostrar en la CONSOLIDACION - TODOS
            If pbASoles Then
                'A SOLES
                If pnTipCambio = 0 Then
                    sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                        "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                        "       SUM(ISNULL(nLogReqDetMesCantidad,0)) nLogReqDetMesCantidad, " & _
                        "       AVG(DISTINCT CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) * r.nLogReqTipCambio END) nLogReqDetPrecio, " & _
                        "       (SUM(ISNULL(nLogReqDetMesCantidad,0)) * AVG(DISTINCT (CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) * r.nLogReqTipCambio END))) SubTotal " & _
                        " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                        "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                        "       LEFT JOIN LogReqDetMes RDM ON rdm.nLogReqNro = rd.nLogReqNro " & _
                        "           AND rdm.nLogReqTraNro = rd.nLogReqTraNro AND rdm.cBSCod = rd.cBSCod" & _
                        "       LEFT JOIN LogReqTramite RT ON rd.nLogReqTraNro = rt.nLogReqTraNro " & _
                        "       LEFT JOIN LogRequerimiento R ON rd.nLogReqNro = r.nLogReqNro " & _
                        " WHERE  rd.nLogReqNro IN (" & pnReqNro & ") " & _
                        "       AND rd.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
                        "                               FROM LogReqDetalle " & _
                        "                               WHERE nLogReqNro = rd.nLogReqNro )" & _
                        " GROUP BY rd.cBSCod, bs.cBSDescripcion, nBSUnidad, cu.cConsDescripcion  " & _
                        " ORDER BY rd.cBSCod "
                Else
                    sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                        "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                        "       SUM(ISNULL(nLogReqDetMesCantidad,0)) nLogReqDetMesCantidad, " & _
                        "       AVG(DISTINCT CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) * " & pnTipCambio & " END) nLogReqDetPrecio, " & _
                        "       (SUM(ISNULL(nLogReqDetMesCantidad,0)) * AVG(DISTINCT (CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) * " & pnTipCambio & " END))) SubTotal " & _
                        " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                        "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                        "       LEFT JOIN LogReqDetMes RDM ON rdm.nLogReqNro = rd.nLogReqNro " & _
                        "           AND rdm.nLogReqTraNro = rd.nLogReqTraNro AND rdm.cBSCod = rd.cBSCod" & _
                        "       LEFT JOIN LogReqTramite RT ON rd.nLogReqTraNro = rt.nLogReqTraNro " & _
                        "       LEFT JOIN LogRequerimiento R ON rd.nLogReqNro = r.nLogReqNro " & _
                        " WHERE  rd.nLogReqNro IN (" & pnReqNro & ") " & _
                        "       AND rd.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
                        "                               FROM LogReqDetalle " & _
                        "                               WHERE nLogReqNro = rd.nLogReqNro )" & _
                        " GROUP BY rd.cBSCod, bs.cBSDescripcion, nBSUnidad, cu.cConsDescripcion  " & _
                        " ORDER BY rd.cBSCod "
                End If
            Else
                'A DOLARES
                sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                    "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                    "       SUM(ISNULL(nLogReqDetMesCantidad,0)) nLogReqDetMesCantidad, " & _
                    "       AVG(DISTINCT CASE WHEN rt.nLogReqTraMoneda = " & gMonedaExtranjera & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) / " & pnTipCambio & " END) nLogReqDetPrecio, " & _
                    "       (SUM(ISNULL(nLogReqDetMesCantidad,0)) * AVG(DISTINCT (CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) / " & pnTipCambio & " END))) SubTotal " & _
                    " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                    "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                    "       LEFT JOIN LogReqDetMes RDM ON rdm.nLogReqNro = rd.nLogReqNro " & _
                    "           AND rdm.nLogReqTraNro = rd.nLogReqTraNro AND rdm.cBSCod = rd.cBSCod" & _
                    "       LEFT JOIN LogReqTramite RT ON rd.nLogReqTraNro = rt.nLogReqTraNro " & _
                    "       LEFT JOIN LogRequerimiento R ON rd.nLogReqNro = r.nLogReqNro " & _
                    " WHERE  rd.nLogReqNro IN (" & pnReqNro & ") " & _
                    "       AND rd.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
                    "                               FROM LogReqDetalle " & _
                    "                               WHERE nLogReqNro = rd.nLogReqNro )" & _
                    " GROUP BY rd.cBSCod, bs.cBSDescripcion, nBSUnidad, cu.cConsDescripcion  " & _
                    " ORDER BY rd.cBSCod "
            End If
        ElseIf TpoCarga = ReqDetTodosSelec Then
            'CONSOLIDADO en proceso de SELECCION - Con restriccion en algunos MESES
            If pbASoles Then
                'A SOLES
                If pnTipCambio = 0 Then
                    sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                        "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                        "       SUM(ISNULL(nLogReqDetMesCantidad,0)) nLogReqDetMesCantidad, " & _
                        "       AVG(DISTINCT CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) * r.nLogReqTipCambio END) nLogReqDetPrecio, " & _
                        "       (SUM(ISNULL(nLogReqDetMesCantidad,0)) * AVG(DISTINCT (CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) * r.nLogReqTipCambio END))) SubTotal " & _
                        " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                        "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                        "       LEFT JOIN LogReqDetMes RDM ON rdm.nLogReqNro = rd.nLogReqNro " & _
                        "           AND rdm.nLogReqTraNro = rd.nLogReqTraNro AND rdm.cBSCod = rd.cBSCod" & _
                        "           AND rdm.nLogReqDetMesNro IN (" & psBSCod & ") " & _
                        "       LEFT JOIN LogReqTramite RT ON rd.nLogReqTraNro = rt.nLogReqTraNro And rt.nLogReqNro = rt.nLogReqTraNro " & _
                        "       LEFT JOIN LogRequerimiento R ON rd.nLogReqNro = r.nLogReqNro " & _
                        " WHERE  rd.nLogReqNro IN (" & pnReqNro & ") AND rdm.nLogSelNro Is Null " & _
                        "       AND rd.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
                        "                               FROM LogReqDetalle " & _
                        "                               WHERE nLogReqNro = rd.nLogReqNro )" & _
                        " GROUP BY rd.cBSCod, bs.cBSDescripcion, nBSUnidad, cu.cConsDescripcion  " & _
                        " ORDER BY rd.cBSCod "
                Else
                    sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                        "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                        "       SUM(ISNULL(nLogReqDetMesCantidad,0)) nLogReqDetMesCantidad, " & _
                        "       AVG(DISTINCT CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) * " & pnTipCambio & " END) nLogReqDetPrecio, " & _
                        "       (SUM(ISNULL(nLogReqDetMesCantidad,0)) * AVG(DISTINCT (CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) * " & pnTipCambio & " END))) SubTotal " & _
                        " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                        "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                        "       LEFT JOIN LogReqDetMes RDM ON rdm.nLogReqNro = rd.nLogReqNro " & _
                        "           AND rdm.nLogReqTraNro = rd.nLogReqTraNro AND rdm.cBSCod = rd.cBSCod" & _
                        "           AND rdm.nLogReqDetMesNro IN (" & psBSCod & ") " & _
                        "       LEFT JOIN LogReqTramite RT ON rd.nLogReqTraNro = rt.nLogReqTraNro And rt.nLogReqNro = rt.nLogReqTraNro  " & _
                        "       LEFT JOIN LogRequerimiento R ON rd.nLogReqNro = r.nLogReqNro " & _
                        " WHERE  rd.nLogReqNro IN (" & pnReqNro & ") AND rdm.nLogSelNro Is Null " & _
                        "       AND rd.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
                        "                               FROM LogReqDetalle " & _
                        "                               WHERE nLogReqNro = rd.nLogReqNro )" & _
                        " GROUP BY rd.cBSCod, bs.cBSDescripcion, nBSUnidad, cu.cConsDescripcion  " & _
                        " ORDER BY rd.cBSCod "
                End If
            Else
                'A DOLARES
                sSql = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
                    "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                    "       SUM(ISNULL(nLogReqDetMesCantidad,0)) nLogReqDetMesCantidad, " & _
                    "       AVG(DISTINCT CASE WHEN rt.nLogReqTraMoneda = " & gMonedaExtranjera & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) / " & pnTipCambio & " END) nLogReqDetPrecio, " & _
                    "       (SUM(ISNULL(nLogReqDetMesCantidad,0)) * AVG(DISTINCT (CASE WHEN rt.nLogReqTraMoneda = " & gMonedaNacional & " THEN ISNULL(rd.nLogReqDetPrecio,0) ELSE ISNULL(rd.nLogReqDetPrecio,0) / " & pnTipCambio & " END))) SubTotal " & _
                    " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
                    "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
                    "       LEFT JOIN LogReqDetMes RDM ON rdm.nLogReqNro = rd.nLogReqNro " & _
                    "           AND rdm.nLogReqTraNro = rd.nLogReqTraNro AND rdm.cBSCod = rd.cBSCod" & _
                    "           AND rdm.nLogReqDetMesNro IN (" & psBSCod & ") " & _
                    "       LEFT JOIN LogReqTramite RT ON rd.nLogReqTraNro = rt.nLogReqTraNro And rt.nLogReqNro = rt.nLogReqTraNro " & _
                    "       LEFT JOIN LogRequerimiento R ON rd.nLogReqNro = r.nLogReqNro " & _
                    " WHERE  rd.nLogReqNro IN (" & pnReqNro & ") AND rdm.nLogSelNro Is Null " & _
                    "       AND rd.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
                    "                               FROM LogReqDetalle " & _
                    "                               WHERE nLogReqNro = rd.nLogReqNro )" & _
                    " GROUP BY rd.cBSCod, bs.cBSDescripcion, nBSUnidad, cu.cConsDescripcion  " & _
                    " ORDER BY rd.cBSCod "
            End If
        Else
            sSql = ""
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaReqDetalle = rs
End Function


'Devuelve un RecordSet con los SelDetalle
Public Function CargaSelDetalle(ByVal pnSelNro As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT sd.cBSCod, bs.cBSDescripcion, " & _
            "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
            "       nLogSelDetCantidad, nLogSelDetPrecio, " & _
            "       (nLogSelDetCantidad * nLogSelDetPrecio) SubTotal" & _
            " FROM LogSelDetalle SD LEFT JOIN BienesServicios BS ON sd.cBSCod = bs.cBSCod " & _
            "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
            " WHERE  sd.nLogSelNro = " & pnSelNro & " "
        'Para Consultas
        'sSQL = "SELECT rd.cBSCod, bs.cBSDescripcion, " & _
            "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
            "       SUM(ISNULL(nLogReqDetMesCantidad,0)) nLogReqDetMesCantidad, " & _
            "       ISNULL(rd.nLogReqDetPrecio,0) nLogReqDetPrecio," & _
            "       (SUM(ISNULL(nLogReqDetMesCantidad,0)) * ISNULL(rd.nLogReqDetPrecio,0)) SubTotal " & _
            " FROM LogReqDetalle RD LEFT JOIN BienesServicios BS ON rd.cBSCod = bs.cBSCod " & _
            "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
            "       LEFT JOIN LogReqDetMes RDM ON rdm.nLogReqNro = rd.nLogReqNro " & _
            "           AND rdm.nLogReqTraNro = rd.nLogReqTraNro AND rdm.cBSCod = rd.cBSCod" & _
            " WHERE  rdm.nLogSelNro = " & pnReqNro & " " & _
            "       AND rd.nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
            "                               FROM LogReqDetalle " & _
            "                               WHERE nLogReqNro = rd.nLogReqNro )" & _
            " GROUP BY rd.cBSCod, bs.cBSDescripcion, nBSUnidad, cu.cConsDescripcion, " & _
            "       rd.nLogReqDetPrecio" & _
            " ORDER BY rd.cBSCod "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelDetalle = rs
End Function


'Devuelve un RecordSet con los ReqDetMes
Public Function CargaReqDetMes(ByVal TpoCarga As TpoCargaReqDetMes, ByVal pnReqNro As String, _
Optional ByVal pnReqTraNro As Integer = 0, Optional ByVal psBSCod As String = "") As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If TpoCarga = ReqDetMesUltTraNro Then
            sSql = "SELECT SUM(CASE WHEN c.nConsValor = " & gMesEnero & " THEN r.nLogReqDetMesCantidad ELSE Null END) Enero ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesFebrero & " THEN r.nLogReqDetMesCantidad ELSE Null END) Febrero ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesMarzo & " THEN r.nLogReqDetMesCantidad ELSE Null END) Marzo ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesAbril & " THEN r.nLogReqDetMesCantidad ELSE Null END) Abril ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesMayo & " THEN r.nLogReqDetMesCantidad ELSE Null END) Mayo ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesJunio & " THEN r.nLogReqDetMesCantidad ELSE Null END) Junio , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesJulio & " THEN r.nLogReqDetMesCantidad ELSE Null END) Julio , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesAgosto & " THEN r.nLogReqDetMesCantidad ELSE Null END) Agosto , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesSeptiembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Setiembre , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesOctubre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Octubre , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesNoviembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Noviembre , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesDiciembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Diciembre " & _
                " FROM LogReqDetMes R JOIN Constante C ON r.nLogReqDetMesNro = c.nConsValor " & _
                " WHERE nLogReqNro = " & pnReqNro & "" & _
                "       AND nLogReqTraNro = (SELECT MAX(nLogReqTraNro) FROM LogReqDetalle WHERE nLogReqNro = " & pnReqNro & " )" & _
                "       AND c.nConsCod = " & gMeses & " " & _
                " GROUP BY r.cBSCod "
        ElseIf TpoCarga = ReqDetMesTraNro Then
            sSql = "SELECT SUM(CASE WHEN c.nConsValor = " & gMesEnero & " THEN r.nLogReqDetMesCantidad ELSE Null END) Enero ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesFebrero & " THEN r.nLogReqDetMesCantidad ELSE Null END) Febrero ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesMarzo & " THEN r.nLogReqDetMesCantidad ELSE Null END) Marzo ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesAbril & " THEN r.nLogReqDetMesCantidad ELSE Null END) Abril ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesMayo & " THEN r.nLogReqDetMesCantidad ELSE Null END) Mayo ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesJunio & " THEN r.nLogReqDetMesCantidad ELSE Null END) Junio , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesJulio & " THEN r.nLogReqDetMesCantidad ELSE Null END) Julio , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesAgosto & " THEN r.nLogReqDetMesCantidad ELSE Null END) Agosto , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesSeptiembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Setiembre , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesOctubre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Octubre , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesNoviembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Noviembre , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesDiciembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Diciembre " & _
                " FROM LogReqDetMes R JOIN Constante C ON r.nLogReqDetMesNro = c.nConsValor " & _
                " WHERE nLogReqNro = " & pnReqNro & " " & _
                "       AND nLogReqTraNro = " & pnReqTraNro & " " & _
                "       AND c.nConsCod = " & gMeses & " " & _
                " GROUP BY r.cBSCod "
        ElseIf TpoCarga = ReqDetMesUnRegistro Then
            sSql = "SELECT c.nConsValor, c.cConsDescripcion, IsNull(nLogReqDetMesCantidad,0) nLogReqDetMesCantidad " & _
                " FROM Constante C LEFT JOIN LogReqDetMes R ON c.nConsValor = r.nLogReqDetMesNro " & _
                "       AND nLogReqNro = " & pnReqNro & " AND nLogReqTraNro = " & pnReqTraNro & " " & _
                "       AND cBSCod = '" & psBSCod & "' " & _
                " WHERE c.nConsCod = " & gMeses & "  " & _
                " ORDER BY nConsValor"
        ElseIf TpoCarga = ReqDetMesTodosConsol Then
            sSql = "SELECT SUM(CASE WHEN c.nConsValor = " & gMesEnero & " THEN r.nLogReqDetMesCantidad ELSE Null END) Enero ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesFebrero & " THEN r.nLogReqDetMesCantidad ELSE Null END) Febrero ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesMarzo & " THEN r.nLogReqDetMesCantidad ELSE Null END) Marzo ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesAbril & " THEN r.nLogReqDetMesCantidad ELSE Null END) Abril ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesMayo & " THEN r.nLogReqDetMesCantidad ELSE Null END) Mayo ," & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesJunio & " THEN r.nLogReqDetMesCantidad ELSE Null END) Junio , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesJulio & " THEN r.nLogReqDetMesCantidad ELSE Null END) Julio , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesAgosto & " THEN r.nLogReqDetMesCantidad ELSE Null END) Agosto , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesSeptiembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Setiembre , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesOctubre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Octubre , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesNoviembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Noviembre , " & _
                "       SUM(CASE WHEN c.nConsValor = " & gMesDiciembre & " THEN r.nLogReqDetMesCantidad ELSE Null END) Diciembre " & _
                " FROM LogReqDetMes R JOIN Constante C ON r.nLogReqDetMesNro = c.nConsValor " & _
                " WHERE nLogReqNro IN (" & pnReqNro & ") " & _
                "       AND nLogReqTraNro = (SELECT MAX(nLogReqTraNro) " & _
                "                           FROM LogReqDetalle " & _
                "                           WHERE nLogReqNro = r.nLogReqNro )" & _
                "       AND c.nConsCod = " & gMeses & " " & _
                " GROUP BY r.cBSCod "
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaReqDetMes = rs
End Function

Public Function CargaReqControlConsolAprobado(psPeriodo As Integer, psTipReq As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = " Select cast(nLogControlCod as varchar(3))as ConsolNro," _
             & " cast(nLogReqPeriodo as varchar(4)) +' - '+ " _
             & " CASE nLogReqTpo WHEN 1 THEN 'Regular' WHEN 2 THEN 'Extemporaneo' ELSE 'Error' END +' - '+ " _
             & " CASE nLogConsolEstado WHEN 1 THEN 'Para Aprobacion'  WHEN 2 THEN 'Eliminado' WHEN 3 THEN 'Aprobado' ELSE 'Error' END +' - '+ " _
             & " cUltimaActualizacion as DescripcionConsolidado " _
            & " from LogReqControlConsol where nLogConsolEstado in (3) and " _
            & " nLogReqPeriodo = " & psPeriodo & " And nLogReqTpo = " & psTipReq & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaReqControlConsolAprobado = rs
End Function

Public Function CargaReqControlConsolEstadopoCod(psPeriodo As Integer, psTipReq As Integer, pnCodigo As Integer) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    CargaReqControlConsolEstadopoCod = 0
    'busca en el contlo de cosolidado si existe consolidados con estado de 1 o 3 este deberia estar en el maximo
    If oConec.AbreConexion() Then
    sSql = "Select nLogConsolEstado from LogReqControlConsol " _
               & " Where nLogReqPeriodo = " & psPeriodo & " And nLogReqTpo = " & psTipReq & " " _
               & " and nLogControlCod = " & pnCodigo & ""
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    End If
    If rs.EOF Then
        CargaReqControlConsolEstadopoCod = 0
        Exit Function
    End If
    CargaReqControlConsolEstadopoCod = IIf(rs!nLogConsolEstado = Null, 0, rs!nLogConsolEstado)
End Function


Public Function CargaReqControlConsol(psPeriodo As Integer, psTipReq As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = " Select cast(nLogControlCod as varchar(3))as ConsolNro," _
             & " cast(nLogReqPeriodo as varchar(4)) +' - '+ " _
             & " CASE nLogReqTpo WHEN 1 THEN 'Regular' WHEN 2 THEN 'Extemporaneo' ELSE 'Error' END +' - '+ " _
             & " CASE nLogConsolEstado WHEN 1 THEN 'Para Aprobacion'  WHEN 2 THEN 'Eliminado' WHEN 3 THEN 'Aprobado' ELSE 'Error' END +' - '+ " _
             & " cUltimaActualizacion as DescripcionConsolidado " _
            & " from LogReqControlConsol where nLogConsolEstado in (1,3) and " _
            & " nLogReqPeriodo = " & psPeriodo & " And nLogReqTpo = " & psTipReq & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaReqControlConsol = rs
End Function

Public Function CargaDetalleGenerico(pnCodGenerico As String, pnPeriodo As Integer, pntipreq As Integer, pnNLogControlCod As Integer, pnmesIni As Integer, pnmesFin As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = " Select  c.cBSCod as [Codigo de Bien],cBSDescripcion as [Descripcion de Bien],cConsDescripcion as Unidad ,nBSUnidad ,sum(nCantidad)as Cantidad,convert(decimal(10,2),Round(sum(nSubtotal)/sum(nCantidad),2)) as [Precio Ref],sum(nSubtotal) As Monto  " _
             & " From logreqconsol_01 c , Bienesservicios bs ,Constante t " _
             & " where  bs.cBSCod = c.cBSCod and NLogControlCod = " & pnNLogControlCod & " and nLogReqDetMesNro between " & pnmesIni & " and " & pnmesFin & " " _
             & " and nLogReqTpo =" & pntipreq & " and  nLogReqPeriodo =" & pnPeriodo & " and c.cBSCod like '" & pnCodGenerico & "' +'%'  " _
             & " and bs.nBSUnidad = t.nConsValor and t.nConsCod = " & gUnidadMedida & "  " _
             & " group by c.cBSCod ,cBSDescripcion,cConsDescripcion,nBSUnidad "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaDetalleGenerico = rs
End Function

Public Function CargaReqConsolMensual(psPeriodo As Integer, psTipReq As Integer, pbtodosarea As Boolean, psAgencia As String, psArea As String, psmesIni As Integer, psmesFin As Integer, psvista As String, pnmensual As Integer, pnConsolNro As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    Dim pnTodos As Integer
    Dim pnreporte As Integer
    If oConec.AbreConexion() Then
      If pbtodosarea = True Then
              If psvista = "p" Or psvista = "q" Then
                sSql = " from LogReqConsol_01 c,bienesservicios bs,bienesservicios bd,areas a,Agencias g ,constante t where "
              Else
                sSql = " from LogReqConsol_01 c,bienesservicios bs,areas a,Agencias g ,constante t where "
              End If
              
              If psvista = "i" Or psvista = "k" Or psvista = "l" Then
                sSql = sSql + " c.cBSCod = bs.cBSCod And "
              ElseIf psvista = "m" Or psvista = "n" Or psvista = "o" Or psvista = "x" Then
              sSql = sSql + " rtrim(ltrim(bs.cBSCod)) = left(c.cBSCod,5) and "
              ElseIf psvista = "p" Or psvista = "q" Then
              sSql = sSql + " rtrim(ltrim(bs.cBSCod)) = left(c.cBSCod,5) and rtrim(ltrim(bd.cBSCod)) = c.cBSCod and"
              End If
              If psvista = "p" Or psvista = "q" Then
                sSql = sSql + "  c.cAreaCod = a.cAreaCod and " _
                & " bd.nBSUnidad = t.nConsValor and t.nConsCod =1019 and " _
                & " c.cAgenciaCod = g.cAgeCod and nLogReqTpo =" & psTipReq & " and c.NLogControlCod = " & pnConsolNro & " and " _
                & " c.nLogReqPeriodo  = " & psPeriodo & " and nLogReqDetMesNro between " & psmesIni & " and  " & psmesFin & ""
              Else
                sSql = sSql + "  c.cAreaCod = a.cAreaCod and " _
                & " bs.nBSUnidad = t.nConsValor and t.nConsCod =1019 and " _
                & " c.cAgenciaCod = g.cAgeCod and nLogReqTpo =" & psTipReq & " and c.NLogControlCod = " & pnConsolNro & " and " _
                & " c.nLogReqPeriodo  = " & psPeriodo & " and nLogReqDetMesNro between " & psmesIni & " and  " & psmesFin & ""
              End If
         Else
              If psvista = "p" Or psvista = "q" Then
                sSql = " from LogReqConsol_01 c,bienesservicios bs,bienesservicios bd,areas a,Agencias g ,constante t where "
              Else
                sSql = " from LogReqConsol_01 c,bienesservicios bs,areas a,Agencias g,constante t  where "
              End If
              If psvista = "i" Or psvista = "k" Or psvista = "l" Then
                sSql = sSql + " c.cBSCod = bs.cBSCod And "
              ElseIf psvista = "m" Or psvista = "n" Or psvista = "o" Or psvista = "x" Then
              sSql = sSql + " rtrim(ltrim(bs.cBSCod)) = left(c.cBSCod,5) and "
              ElseIf psvista = "p" Or psvista = "q" Then
              sSql = sSql + " rtrim(ltrim(bs.cBSCod)) = left(c.cBSCod,5) and rtrim(ltrim(bd.cBSCod)) = c.cBSCod and"
              End If
              sSql = sSql + " c.cAreaCod = a.cAreaCod and " _
              & " c.cAgenciaCod = g.cAgeCod  and nLogReqTpo =" & psTipReq & " and c.NLogControlCod = " & pnConsolNro & " and " _
              & " c.cAreaCod = " & psArea & " and" _
              & " bs.nBSUnidad = t.nConsValor and t.nConsCod =1019 and" _
              & " c.cAgenciaCod = " & psAgencia & "  and  " _
              & " c.nLogReqPeriodo  = " & psPeriodo & " and nLogReqDetMesNro between " & psmesIni & " and  " & psmesFin & ""
       End If
        If psvista = "i" Or psvista = "k" Or psvista = "l" Or psvista = "m" Or psvista = "n" Or psvista = "o" Or psvista = "p" Or psvista = "q" Or psvista = "x" Then
            If psvista = "i" Then
                pnreporte = 1 'Agencia,Area,Codigo Bien
            ElseIf psvista = "k" Then
                pnreporte = 2 'Area,Codigo Bien
            ElseIf psvista = "l" Then
                pnreporte = 3 'Codigo de Bien
            ElseIf psvista = "p" Then
                pnreporte = 4 'Categoria y Codigo de Bien
            ElseIf psvista = "q" Then
                pnreporte = 5 'Categoria y Codigo de Bien
            ElseIf psvista = "x" Then
                pnreporte = 6 'Categoria y Codigo de Bien
            ElseIf psvista = "m" Then
                pnreporte = 7 'Categoria y Codigo de Bien
            ElseIf psvista = "n" Then
                pnreporte = 8 'Categoria y Codigo de Bien
            ElseIf psvista = "o" Then
                pnreporte = 9 'Categoria y Codigo de Bien
            End If
            sSql = "spLogReqListaConMensual  " & pnreporte & ",'" & sSql & "' ," & pnmensual & ""
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaReqConsolMensual = rs
End Function

Public Function CargaReqListaDetalle(pntipreq As String, bareatodos As Boolean, Optional ByVal psArea As String = "", Optional ByVal psPeriodo As Integer = 2003, Optional ByVal psEstado As TpoReqEstado, Optional ByVal bestadotodos As Boolean, Optional ByVal psvista As String, Optional ByVal pnConsolNro As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = " SELECT r.nLogReqCod as  Código , g.cAgeDescripcion Agencia ,a.cAreaDescripcion as [Area ]," _
             & " r.nLogReqCod as  [Req.Nº]," _
             & " cBSDescripcion Descripcion,nLogReqDetPrecio as  Precio, " _
             & " (select  sum(nLogReqDetMesCantidad) from logreqdetmes  rdm where  rdm.nLogReqCod =rd.nLogReqCod and rdm.cBSCod = rd.cBSCod) Cantidad, " _
             & " nLogReqDetPrecio  * (select  sum(nLogReqDetMesCantidad) from logreqdetmes  rdm where  rdm.nLogReqCod =rd.nLogReqCod and rdm.cBSCod = rd.cBSCod) as Subtotal, " _
             & " Estado = CASE nLogReqEstado   WHEN 2 THEN 'Aprobado'  WHEN 1 THEN 'Solicitado'  WHEN 3 THEN 'Rechazado' ELSE 'Sin Estado' END , " _
             & " Consolidado = CASE nLogReqConsolidado   WHEN 0 THEN 'S/Consolidar'  WHEN 1 THEN 'Consolidado'  Else 'Sin Estado' END " _
             & " FROM LogRequerimiento  r,Areas A,Agencias G,logreqdetalle rd , bienesservicios bs where " _
             & " r.cAreaCod = a.cAreaCod  and r.cAgenciaCod = g.cAgeCod and " _
             & " r.nLogReqCod = rd.nLogReqCod  and " _
             & " bs.cBSCod =rd.cBSCod and "
             sSql = sSql + "  R.nLogReqTpo = '" & pntipreq & "'  and  r.nLogReqPeriodo = " & psPeriodo & "  "
            If bareatodos = False Then 'falata la agencia creo
                sSql = sSql + " and r.cAreaCod = '" & psArea & "'  "
            End If
            If bestadotodos = False Then
                sSql = sSql + " and  r.nLogReqEstado = '" & psEstado & "' "
            End If
            If psvista = "1" Then
            sSql = sSql + "and r.nLogReqCod  in(select  Distinct nLogReqCod from logreqconsol_01 where  NLogControlCod = " & pnConsolNro & " and r.nLogReqPeriodo = " & psPeriodo & " ) "
            ElseIf psvista = "2" Then
            sSql = sSql + "and r.nLogReqConsolidado = 0 "
            End If
            
            sSql = sSql + " order by g.cAgeDescripcion ,cAreaDescripcion, r.nLogReqCod asc "
            Set rs = oConec.CargaRecordSet(sSql)
            oConec.CierraConexion
    End If
    Set CargaReqListaDetalle = rs
End Function

Public Function CargaReqPrecios(psPeriodo As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "SELECT p.cBSCod,bs.cBSDescripcion, " _
             & " IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " _
             & " nLogReqCod,P.nPrecReferencial , P.cUltimaActualizacion,0 as flag " _
             & " FROM logReqBsPrecios p LEFT JOIN BienesServicios BS ON p.cBSCod = bs.cBSCod " _
             & " LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = 1019 " _
             & " where nperiodo ='" & psPeriodo & "' "
            Set rs = oConec.CargaRecordSet(sSql)
            oConec.CierraConexion
    End If
    Set CargaReqPrecios = rs
End Function

Public Function CargaReqControlConsolCodigo(psPeriodo As Integer, psTipReq As Integer) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    CargaReqControlConsolCodigo = 0
    If oConec.AbreConexion() Then
       sSql = " Select nLogControlCod from LogReqControlConsol " _
               & " Where nLogReqPeriodo = " & psPeriodo & " And nLogReqTpo = " & psTipReq & " " _
               & " and nLogControlCod =(select max(nLogControlCod)  from LogReqControlConsol  " _
               & " where  nLogReqPeriodo = " & psPeriodo & "  And nLogReqTpo = " & psTipReq & "  )"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    If rs.EOF Then
        CargaReqControlConsolCodigo = 0
        Exit Function
    End If
    CargaReqControlConsolCodigo = IIf(rs!NLogControlCod = Null, 0, rs!NLogControlCod)
End Function


Public Function CargaReqControlConsolEstado(psPeriodo As Integer, psTipReq As Integer) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    CargaReqControlConsolEstado = 0
    'busca en el contlo de cosolidado si existe consolidados con estado de 1 o 3 este deberia estar en el maximo
    If oConec.AbreConexion() Then
    sSql = "Select nLogConsolEstado from LogReqControlConsol " _
               & " Where nLogReqPeriodo = " & psPeriodo & " And nLogReqTpo = " & psTipReq & " " _
               & " and nLogControlCod =(select max(nLogControlCod)  from LogReqControlConsol " _
               & " where  nLogReqPeriodo = " & psPeriodo & "  And nLogReqTpo = " & psTipReq & "  )"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    End If
    If rs.EOF Then
        CargaReqControlConsolEstado = 0
        Exit Function
    End If
    CargaReqControlConsolEstado = IIf(rs!nLogConsolEstado = Null, 0, rs!nLogConsolEstado)
End Function

Public Function CargaReqSinConsolidar(psPeriodo As Integer, psTipReq As Integer) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    '1 Solicitados
    '2 Aprobados
    '3 Rechazados
    CargaReqSinConsolidar = 0
    If oConec.AbreConexion() Then
        sSql = "select  count(*) as cant from logrequerimiento where nLogReqEstado = 2 and  " _
               & " nLogReqConsolidado = 0 and nLogReqPeriodo = " & psPeriodo & " " _
               & " and nLogReqTpo = " & psTipReq & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    If rs.EOF Then
        CargaReqSinConsolidar = 0
        Exit Function
    End If
    CargaReqSinConsolidar = IIf(rs!Cant = Null, 0, rs!Cant)
End Function

Public Function CargaValidaReqPrecios(psPeriodo As Integer) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = " select  count(*) as cant from LogReqBsPrecios where nperiodo = " & psPeriodo & "" _
                & " and  nPrecReferencial is null "
            Set rs = oConec.CargaRecordSet(sSql)
            oConec.CierraConexion
    End If
    If rs.EOF Then
        CargaValidaReqPrecios = 0
        Exit Function
    End If
    CargaValidaReqPrecios = IIf(rs!Cant = Null, 0, rs!Cant)
End Function

Public Function CargaListaReqPrecios(pnPeriodo As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
            sSql = "select p.cBSCod,cBSDescripcion,cConsDescripcion,nLogReqCod ,nPrecReferencial from " _
                   & " LogReqBsPrecios p , bienesservicios bs ,constante c " _
                   & " where  nperiodo = " & pnPeriodo & " and  p.cBSCod  = bs.cBSCod and " _
                   & " nConsCod = 1019 And c.nConsValor = bs.nBSUnidad "
            Set rs = oConec.CargaRecordSet(sSql)
            oConec.CierraConexion
    End If
    Set CargaListaReqPrecios = rs
End Function



Public Function CargaPrecioReq(psBS As String, psReqNro As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "select  nLogReqDetPrecio from logreqdetalle " _
               & " where nLogReqCod = '" & psReqNro & "' and cBSCod = '" & psBS & "' "
            Set rs = oConec.CargaRecordSet(sSql)
            oConec.CierraConexion
    End If
    Set CargaPrecioReq = rs
End Function


Public Function CargaReqPreciosProm(psPeriodo As Integer, psBS As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "select  d.nLogReqCod ,d.nLogReqDetPrecio  ,d.cBSCod   from logreqdetalle d, logrequerimiento r " _
               & " where  d.nLogReqCod = r.nLogReqCod and r.nLogReqPeriodo = " & psPeriodo & " and  d.cBSCod = '" & psBS & "' "
            Set rs = oConec.CargaRecordSet(sSql)
            oConec.CierraConexion
    End If
    Set CargaReqPreciosProm = rs
End Function

Public Function CargaReqConsol(psPeriodo As Integer, psTipReq As Integer, pbtodosarea As Boolean, psAgencia As String, psArea As String, psmesIni As Integer, psmesFin As Integer, psvista As String, pnConsolNro As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    Dim pnTodos As Integer
    Dim pncategoria As Integer
    If oConec.AbreConexion() Then
       If psvista = "d" Then
        sSql = " select cAgeDescripcion  as Agencia , cAreaDescripcion as Area ,cConsDescripcion as Mes,cBSDescripcion as [Descripcion Bien] ," _
               & " nPrecReferencial as [Precio Referencial],nCantidad as Cantidad, nSubTotal as SubTotal " _
               & " from LogReqConsol_01 c,areas a,agencias g ,bienesservicios bs ,constante n" _
               & " where c.cBSCod = bs.cBSCod and " _
               & " c.cAreaCod =   a.cAreaCod and " _
               & " c.cAgenciaCod  = g.cAgeCod and n.nConsValor = c.nLogReqDetMesNro and n.nConsCod = 1010 and " _
               & " nLogReqPeriodo = " & psPeriodo & " and nLogReqTpo = " & psTipReq & " and  c.NLogControlCod = " & pnConsolNro & " and  " _
               & " nLogReqDetMesNro between " & psmesIni & "  and  " & psmesFin & ""
               If pbtodosarea = False Then
               sSql = sSql + " and  c.cAgenciaCod  = '" & psAgencia & "' and c.cAreaCod  ='" & psArea & "' "
               End If
               sSql = sSql + " order by cAgeDescripcion, cAreaDescripcion,nLogReqDetMesNro desc "
        ElseIf psvista = "r" Or psvista = "f" Or psvista = "g" Or psvista = "h" Then
            If pbtodosarea = True Then
                pnTodos = 2
            Else
                pnTodos = 1
            End If
            If psvista = "r" Then
                pncategoria = 1 'Codigo Bien
            ElseIf psvista = "f" Then
                pncategoria = 2 'Categoria de Bien
            ElseIf psvista = "g" Then
                pncategoria = 3 'Categoria de Bien
             ElseIf psvista = "h" Then
                pncategoria = 4 'Categoria de Bien
            End If
            sSql = "spLogReqListaConsolidado  " & psmesIni & "," & psmesFin & "," & psPeriodo & " ," & psTipReq & " ," & pnTodos & ",'" & psArea & "','" & psAgencia & "'," & pncategoria & "," & pnConsolNro & ""
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaReqConsol = rs
End Function

