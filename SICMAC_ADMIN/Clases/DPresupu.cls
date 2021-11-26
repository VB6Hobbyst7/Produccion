VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DPresupu"
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

Public Enum TpoCargaPlaPri
    PlaPriCbo = 1
    PlaPriFlex = 2
End Enum

Public Enum TpoCargaPlaPre
    PlaPreCbo = 1
    PlaPreFlex = 2
End Enum

Public Enum TpoCargaPlaRub
    PlaRubCbo = 1
    PlaRubFlex = 2
End Enum

Public Enum TpoCargaPlaRubMes
    PlaRubMesPreReq = 1
    PlaRubMesFlex = 2
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

'Devuelve un RecordSet con PresuClase
Public Function CargaPresuClase(ByVal psTpoCarga As TpoCargaPlaPri) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If psTpoCarga = PlaPriCbo Then
            sSql = "SELECT cPresuDescripcion + space(40) + Convert(char(1),nPresuCod) " & _
                " FROM PresuClase "
        ElseIf psTpoCarga = PlaPriFlex Then
            sSql = "SELECT nPresuCod, cPresuDescripcion " & _
                " FROM PresuClase "
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaPresuClase = rs
End Function

'Devuelve un RecordSet con PlaPresupuesto
Public Function CargaPlaPresupu(ByVal psTpoCarga As TpoCargaPlaPre, _
Optional ByVal psPeriodo As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If psTpoCarga = PlaPreCbo Then
            sSql = "SELECT distinct  pri.cPresuDescripcion + space(40) + LTrim(Str(Pri.nPresuCod)) " & _
                " FROM PresuClase Pri JOIN PresuRubroAnio Pre ON pri.nPresuCod = pre.nPresuCod "
            
            sSql = sSql + IIf(Len(Trim(psPeriodo)) = 0, " AND pre.cPeriodo = '" & psPeriodo & "'", "")
        ElseIf psTpoCarga = PlaPreFlex Then
            sSql = "SELECT pre.nPlaPreNro, pri.cPresuDescripcion " & _
                " FROM PresuClase Pri JOIN PlaPresupuesto Pre ON pri.nPresuCod = pre.nPresuCod "
            
            sSql = sSql + IIf(Len(Trim(psPeriodo)) = 0, " AND pre.cPeriodo = '" & psPeriodo & "'", "")
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaPlaPresupu = rs
End Function

'Devuelve un RecordSet con PlaRubro
Public Function CargaPlaRubro(ByVal psTpoCarga As TpoCargaPlaRub, _
ByVal psPlaPreNro As String, ByVal psPeriodo As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If psTpoCarga = PlaRubCbo Then
            sSql = " SELECT LTrim(r.cPresuRubCod) + Space(4) + r.cPresuRubDescripcion " & _
                " FROM PresuRubroAnio R " & _
                " WHERE r.nPresuCod = '" & psPlaPreNro & "' And nPresuAnio = " & psPeriodo & "" & _
                " ORDER BY r.cPresuRubCod "
        ElseIf psTpoCarga = PlaRubFlex Then
            sSql = "SELECT cPlaRubCod, cPlaRubDescripcion, nPlaRubMonto " & _
                " FROM PlaRubro " & _
                " WHERE cPlaPreNro = '" & psPlaPreNro & "' "
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaPlaRubro = rs
End Function

'Devuelve un RecordSet con PlaRubCta
Public Function CargaPlaRubCta(ByVal psPlaPreNro As String, _
ByVal psPlaRubCod As String, ByVal psPeriodo As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT cCtaContCod " & _
            " FROM PresuRubroCta " & _
            " WHERE nPresuCod = '" & psPlaPreNro & "' " & _
            "       AND cPresuRubCod = '" & psPlaRubCod & "' And nPresuAnio = '" & psPeriodo & "' "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaPlaRubCta = rs
End Function

'Devuelve un RecordSet con PlaRubMes
Public Function CargaPlaRubMes(ByVal psTpoCarga As TpoCargaPlaRubMes, _
ByVal psPlaPreNro As String, ByVal psPlaRubCod As String, ByVal psPeriodo As String, _
Optional ByVal psReqNro As String, Optional ByVal psReqTraNro As String, _
Optional ByVal psBSCod As String, Optional ByVal pnTipCambio As Currency) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If psTpoCarga = PlaRubMesFlex Then
            sSql = "SELECT rm.nPresuMes, cm.cConsDescripcion, rm.nPresuRubMesMonIni, rm.nPresuRubMesMonEje, " & _
                "       rm.nPresuRubMesMonRes " & _
                " FROM PresuRubroMes RM JOIN Constante CM ON rm.nPresuMes = cm.nConsValor AND cm.nConsCod = " & gMeses & " " & _
                " WHERE rm.nPresuCod = '" & psPlaPreNro & "' AND rm.cPresuRubCod = '" & psPlaRubCod & "'"
        ElseIf psTpoCarga = PlaRubMesPreReq Then
            sSql = "SELECT rm.nPresuMes, substring(cm.cConsDescripcion,1,10), rm.nPresuRubMesMonIni, rm.nPresuRubMesMonEje,  " & _
                "       rm.nPresuRubMesMonRes, Round(IsNull(rdm.nLogReqDetMesCantidad,0) * (rd.nLogReqDetPrecio * " & pnTipCambio & " ),2) nLogReqDetMesCantidad" & _
                " FROM PresuRubroMes RM JOIN Constante CM ON rm.nPresuMes = cm.nConsValor AND cm.nConsCod = " & gMeses & " " & _
                "           AND rm.nPresuCod = '" & psPlaPreNro & "' AND rm.cPresuRubCod = '" & psPlaRubCod & "' And rm.nPresuAnio = '" & psPeriodo & "' " & _
                "       LEFT JOIN LogReqDetMes RDM ON cm.nConsValor = rdm.nLogReqDetMesNro " & _
                "           AND rdm.nLogReqNro = '" & psReqNro & "' AND rdm.nLogReqTraNro = " & psReqTraNro & " " & _
                "           AND rdm.cBSCod = '" & psBSCod & "' " & _
                "       LEFT JOIN LogReqDetalle RD ON rdm.nLogReqNro = rd.nLogReqNro And rdm.nLogReqTraNro = rd.nLogReqTraNro " & _
                "           AND rdm.cBSCod = rd.cBSCod "
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaPlaRubMes = rs
End Function

