VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DPigImpresiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim sql As String
Dim oConn As DConecta
Private Sub Class_Initialize()
Dim oIni As ClasIni
Dim sConn As String

    Set oIni = New ClasIni
    sConn = oIni.CadenaConexion
    Set oIni = Nothing
    
Set oConn = New DConecta
If oConn.AbreConexion(sConn) = False Then
    Call RaiseError(MyUnhandledError, "DPigActualizaBD:Initialize. Error en Conexion a Base de datos")
End If

End Sub

Private Sub Class_Terminate()
    oConn.CierraConexion
    Set oConn = Nothing
End Sub

'******************************************************************************************
' CAFF -  10/10/2002 - Devuelve datos para impresion de las guias de la Remesa - Cabecera
'******************************************************************************************
Public Function GetDatosColocPigGuia(ByVal psNumGuia As String, ByVal pnTipoGuia As Integer) As Recordset

On Error GoTo ErrorGetDatos
    sql = "SELECT G.cNumDoc, nTotItem, SUM(nTotPiezasCont) TotalPiezas, nPesoBruto, nPesoNeto, nTotTasacion, " _
          & "C1.cConsDescripcion Origen, C2.cConsDescripcion Destino, C3.cConsDescripcion Motivo " _
          & "FROM COLOCPIGGUIA G INNER JOIN CONSTANTE C1 ON G.nOrigen = C1.nConsValor " _
          & "INNER JOIN CONSTANTE C2 ON G.nDestino = C2.nConsValor " _
          & "INNER JOIN CONSTANTE C3 ON G.nMotivo = C3.nConsValor " _
          & "INNER JOIN COLOCPIGGUIADET GD ON GD.cNumDoc = G.cNumDoc " _
          & "WHERE C1.nConsCod = 8004 AND C2.nConsCod = 8004 AND C3.nConsCod = 8015 AND G.cNumDoc = '" & psNumGuia _
          & "' AND nTipoGuia = " & pnTipoGuia _
          & " GROUP BY G.cNumDoc, nTotItem, nPesoBruto, nPesoNeto, nTotTasacion, C1.cConsDescripcion, C2.cConsDescripcion, " _
          & "C3.cConsDescripcion"
          
    Set GetDatosColocPigGuia = oConn.CargaRecordSet(sql)
      
    Exit Function

ErrorGetDatos:
    Err.Raise Err.Number, "Error En Proceso <<GetDatosColocPigGuia>>", Err.Description

End Function

'**************************************************************************************************
' CAFF -  10/10/2002 - Devuelve datos para impresion de las guias de la Remesa - Detalle por Lotes
'**************************************************************************************************
Public Function GetDatosColocPigGuiaDetLote(ByVal psNumDoc As String) As Recordset

On Error GoTo ErrorGetDatos
    sql = "SELECT GD.cCtaCod, nTotPiezasCont, SUM(J.nPesoBruto) nPBruto, SUM(J.nPesoNeto) nPNeto, " _
        & "SUM(nTasacion + nTasacionAdicional) Tasacion, cUser " _
        & "FROM COLOCPIGGUIADET GD INNER JOIN COLOCPIGJOYATASACION J ON GD.cCtaCod = J.cCtaCod " _
        & "INNER JOIN COLOCPIGGUIA G ON G.cNumDoc = GD.cNumDoc " _
        & "INNER JOIN COLOCPIGNO C ON C.cCtaCod = J.cCtaCod AND C.nTipoTasacion = J.nTipoTasacion " _
        & "INNER JOIN RRHH RH ON J.cPersCod = RH.cPersCod " _
        & "WHERE G.cNumDoc = '" & psNumDoc & "'  AND nTipoGuia = 1 " _
        & "GROUP BY GD.cCtaCod, nTotPiezasCont, cUser"
    
    Set GetDatosColocPigGuiaDetLote = oConn.CargaRecordSet(sql)
      
    Exit Function

ErrorGetDatos:
    Err.Raise Err.Number, "Error En Proceso <<GetDatosColocPigGuiaDetLote>>", Err.Description

End Function

'***************************************************************************************************
' CAFF -  10/10/2002 - Devuelve datos para impresion de las guias de la Remesa - Detalle por Piezas
'***************************************************************************************************
Public Function GetDatosColocPigGuiaDetPieza(ByVal psNumDoc As String) As Recordset

On Error GoTo ErrorGetDatos
    sql = "SELECT GD.cCtaCod, GD.nItemPieza, cDescripcion, J.nPesoBruto, J.nPesoNeto, nTasacion + nTasacionAdicional Tasacion, " _
        & "cConsDescripcion, cUser " _
        & "FROM COLOCPIGGUIADET GD INNER JOIN COLOCPIGJOYATASACION J ON GD.cCtaCod = J.cCtaCod " _
        & " AND GD.nItemPieza = J.nItemPieza " _
        & "INNER JOIN COLOCPIGGUIA G ON G.cNumDoc = GD.cNumDoc  " _
        & "INNER JOIN COLOCPIGNO C ON C.cCtaCod = J.cCtaCod AND C.nTipoTasacion = J.nTipoTasacion " _
        & "INNER JOIN CONSTANTE C1 ON J.nMaterial = C1.nConsValor INNER JOIN RRHH RH ON RH.cPersCod = J.cPersCod " _
        & "WHERE nConsCod = 8003 AND nTipoGuia = 2 AND G.cNumDoc = '" & psNumDoc & "'"
    
    Set GetDatosColocPigGuiaDetPieza = oConn.CargaRecordSet(sql)
      
    Exit Function

ErrorGetDatos:
    Err.Raise Err.Number, "Error En Proceso <<GetDatosColocPigGuiaDetPieza>>", Err.Description

End Function

'***************************************************************************************************
' CAFF -  10/10/2002 - Devuelve datos para impresion de la Guia de Remision
'***************************************************************************************************
Public Function GetDatosColocPigGuiaRemCab(ByVal psNumDoc As String) As Recordset

On Error GoTo ErrorGetDatos

    sql = "SELECT A1.cAgeDescripcion Origen, A1.cAgeDireccion DirOrigen, A2.cAgeDescripcion Destino, A2.cAgeDireccion DirDestino, G.cPersCod, " _
        & "cPersNombre, cPersDireccDomicilio " _
        & "FROM ColocPigGuia G LEFT JOIN AGENCIAS A1 ON A1.cAgeCod = nOrigen " _
        & "LEFT JOIN AGENCIAS A2 ON A2.cAgeCod = Case nDestino WHEN 99 Then '01' ELSE nDestino END " _
        & "INNER JOIN PERSONA P ON  P.cPersCod =  G.cPersCod " _
        & "WHERE cNumDoc = '" & psNumDoc & "'"
    
    Set GetDatosColocPigGuiaRemCab = oConn.CargaRecordSet(sql)
    
    Exit Function

ErrorGetDatos:
    Err.Raise Err.Number, "Error En Proceso <<GetDatosColocPigGuiaRemCab>>", Err.Description

End Function

'***************************************************************************************************
' CAFF -  10/10/2002 - Devuelve datos para impresion de la Guia de Remision - Detalle de la Guia
'***************************************************************************************************
Public Function GetDatosColocPigGuiaRemDet(ByVal psNumDoc As String, Optional ByVal psTipoGuia) As Recordset

On Error GoTo ErrorGetDatos

    sql = "SELECT cConsDescripcion, COUNT(*) Piezas,  SUM(nPesoNeto) Peso " _
        & "FROM COLOCPIGGUIADET GD INNER JOIN COLOCPIGJOYATASACION J ON GD.cCtaCod = J.cCtaCod "
        If psTipoGuia = 2 Then
            sql = sql & " AND GD.nItemPieza = J.nItemPieza "
        End If
    sql = sql & "INNER JOIN COLOCPIGNO C ON C.cCtaCod = J.cCtaCod AND C.nTipoTasacion = J.nTipoTasacion " _
        & "INNER JOIN CONSTANTE C1 ON C1.nConsValor = nMaterial " _
        & "WHERE nConsCod = 8003 AND cNumDoc = '" & psNumDoc & "'" _
        & "GROUP BY cConsDescripcion"
    
    Set GetDatosColocPigGuiaRemDet = oConn.CargaRecordSet(sql)
    
    Exit Function

ErrorGetDatos:
    Err.Raise Err.Number, "Error En Proceso <<GetDatosColocPigGuiaRemDet>>", Err.Description

End Function

Public Function GetDatosCabMartillero(ByVal pnNumRemate As Integer, ByVal pnTipoProceso As Integer) As Recordset

On Error GoTo ErrorGetDatos

    sql = "SELECT CT.cCtaCod, SUM(nPesoBruto) PesoNeto, SUM(nTasacion + nTasacionAdicional) Tasacion, " _
        & "SUM(nValorDeuda) ValorDeuda, SUM(nValorProceso) ValorBase " _
        & "FROM COLOCPIGJOYATASACION CT INNER JOIN COLOCPIGNO CP ON CT.cCtaCod = CP.cCtaCod AND CT.nTipoTasacion = CP.nTipoTasacion " _
        & "INNER JOIN COLOCPIGPROCESO CR ON CT.CCTACOD = CR.CCTACOD AND CT.nItemPieza= CR.nItemPieza " _
        & "WHERE CT.nTipoTasacion = 3 And nRemate = " & pnNumRemate & " AND nTipoProceso = " & pnTipoProceso _
        & " AND nSituacionPieza = " & gPigSituacionDisponible _
        & " GROUP BY CT.cCtaCod"

    Set GetDatosCabMartillero = oConn.CargaRecordSet(sql)

    Exit Function
    
ErrorGetDatos:
    Err.Raise Err.Number, "Error En Proceso <<GetDatosCabMartillero>>", Err.Description
    
End Function

Public Function GetDatosDetMartillero(ByVal pnNumRemate As Integer, ByVal pnTipoProceso As Integer, ByVal psCtaCod As String) As Recordset

On Error GoTo ErrorGetDatos

    sql = "SELECT CT.nItemPieza, C.cConsDescripcion TipoJoya, nMaterial, " _
        & "SubTipoJoya = (SELECT cConsDescripcion FROM CONSTANTE WHERE nConsValor = CT.nSubTipoJoya AND nConsCod = 8002 ), " _
        & "Material = (SELECT cConsDescripcion FROM CONSTANTE WHERE nConsValor = CT.nMaterial AND nConsCod = 8003), " _
        & "nPesoBruto, cDescripcion, cDescripcionAdic, nPesoNeto, nValorProceso, " _
        & "Conservacion = (SELECT cConsDescripcion FROM CONSTANTE WHERE nConsValor = CT.nConservacion AND nConsCod = 8006), " _
        & "ValorGramo = (SELECT nValor FROM COLOCPIGPRECIOMATERIAL WHERE nMaterial = CT.nMaterial AND nTipoValor = 3) " _
        & "FROM COLOCPIGJOYATASACION CT INNER JOIN COLOCPIGNO CP ON CT.cCtaCod = CP.cCtaCod AND CT.nTipoTasacion = CP.nTipoTasacion " _
        & "INNER JOIN COLOCPIGPROCESO CR ON CT.cCtaCod = CR.cCtaCod AND CT.nItemPieza= CR.nItemPieza " _
        & "INNER JOIN CONSTANTE C ON CT.nTipoJoya = C.nConsValor " _
        & "WHERE CT.nTipoTasacion = 3 And C.nConsCod = 8001 AND nRemate = " & pnNumRemate & " AND nTipoProceso = " & pnTipoProceso _
        & " AND CT.cCtaCod = '" & psCtaCod & "' AND nSituacionPieza = " & gPigSituacionDisponible _
        & "ORDER BY CT.nItemPieza"

    Set GetDatosDetMartillero = oConn.CargaRecordSet(sql)

    Exit Function
    
ErrorGetDatos:
    Err.Raise Err.Number, "Error En Proceso <<GetDatosDetMartillero>>", Err.Description

End Function

Public Function dObtieneGuiaRecep(ByVal psNumGuia As String) As Recordset
Dim Rs As Recordset

On Error GoTo dError
sql = "SELECT cNumDoc, Motivo = (SELECT cConsDescripcion FROM CONSTANTE WHERE nConsCod = 8015 AND nConsValor = G.nMotivo), " _
        & "Origen = (SELECT cConsDescripcion FROM CONSTANTE WHERE nConsCod = 8004 AND nConsValor = G.nOrigen), " _
        & "Destino = (SELECT cConsDescripcion FROM CONSTANTE WHERE nConsCod = 8004 AND nConsValor = G.nDestino), " _
        & "nTotItem, TipoGuia = (SELECT cConsDescripcion FROM CONSTANTE WHERE nConsCod = 8010 AND nConsValor = G.nTipoGuia), " _
        & "nMotivo, nPesoNeto, nPesoBruto, nTotTasacion, cSerDocSunat, cNumDocSunat " _
        & "FROM COLOCPIGGUIA G WHERE cNumDoc = '" & psNumGuia & "'"
        
Set Rs = oConn.CargaRecordSet(sql)
Set dObtieneGuiaRecep = Rs

Set Rs = Nothing

Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Guia Recepcion <<dObtieneGuiaRecep>>", Err.Description

End Function

