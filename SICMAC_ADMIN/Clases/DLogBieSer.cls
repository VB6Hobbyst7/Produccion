VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogBieSer"
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

Public Enum TpoCargaBS
    BsUnRegistro = 0
    BsTodosArbol = 1
    BsTodosFlex = 2
    BsSuperiorFlex = 3
    BsTodos = 4
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

'Devuelva un RecordSet con los Bienes/Servicios
Public Function CargaBS(ByVal pnTpoCarga As TpoCargaBS, Optional ByVal psCodBS As String = "") As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If pnTpoCarga = BsUnRegistro Then
            'Un registro
            sSql = "SELECT bs.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(c.cConsDescripcion) + space(40) + Convert(char(3),bs.nBSUnidad)),'') cConsUnidad  " & _
                "     , bs.bSerie Serie, bs.bVerifica Verificable, bs.bContiene Contenido, bs.bVigente Vigente, BS.bCorrela Correlativo, BS.nPorDeprecia Per_Depre, BS.nBSTipDep Tipo,isnull(BS.nStockMinimo,0) StockMin " & _
                " FROM BienesServicios BS LEFT JOIN Constante C On bs.nBSUnidad = c.nConsValor " & _
                "           AND nConsCod = " & gUnidadMedida & " " & _
                " WHERE bs.cBSCod = '" & psCodBS & "' "
        ElseIf pnTpoCarga = BsTodosArbol Then
            'Todos los registros para el Arbol
            sSql = "SELECT RTrim(bs.cBSCod) A, bs.cBSDescripcion Descripción, len(bs.cBSCod) " & _
                " FROM BienesServicios BS " & _
                " ORDER BY bs.cBSCod "
        ElseIf pnTpoCarga = BsTodosFlex Then
            'Todos los registros para flex
            sSql = " SELECT BS.cBSCod Código, BS.cBSCod Código, BS.cBSDescripcion Descripción, BS.bSerie Serie, BS.bVerifica Verificable, BS.bContiene Contenido, BS.bCorrela Correlativo, BS.nPorDeprecia Per_Depre, nBSTipDep Tipo,isnull(BS.nStockMinimo,0) StockMin " & _
                 "   FROM BienesServicios BS Where bs.cBSCod LIKE '" & psCodBS & "%' " & _
                 "   ORDER BY bs.cBSCod "
        ElseIf pnTpoCarga = BsSuperiorFlex Then
            'Registros superiores para flex
            sSql = "SELECT bs.cBSCod Código, bs.cBSCod Código, bs.cBSDescripcion Descripción , bs.bSerie Serie, bs.bVerifica Verificable, bs.bContiene Contenido, BS.bCorrela Correlativo, BS.nPorDeprecia Per_Depre, nBSTipDep Tipo,isnull(BS.nStockMinimo,0) StockMin  " & _
                " FROM BienesServicios BS WHERE bs.cBSCod LIKE '" & psCodBS & "%' AND " & _
                "       LEN(bs.cBSCod) = (SELECT MIN(LEN(cBSCod)) FROM BienesServicios WHERE cBSCod LIKE '" & psCodBS & "%') " & _
                " ORDER BY bs.cBSCod "
        ElseIf pnTpoCarga = BsTodos Then
            'Todos los registros para flex
            sSql = "SELECT bs.cBSCod , bs.cBSDescripcion, c.cConsDescripcion" & _
                " FROM BienesServicios BS LEFT JOIN Constante C On bs.nBSUnidad = c.nConsValor " & _
                "       AND nConsCod = " & gUnidadMedida & " " & _
                " ORDER BY bs.cBSCod "
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaBS = rs
    Set rs = Nothing
End Function

'Agrega un Bien/Servicio
Public Function GrabaBS(ByVal psBSCod As String, ByVal psDescripcion As String, ByVal pnUnidad As Integer, ByVal pbSerie As Boolean, ByVal pbVerifica As Boolean, ByVal pbContenido As Boolean, ByVal pbVigente As Boolean, ByVal pbCorrela As Boolean, ByVal psPorDeprecia As String, ByVal psTipoDeprecia As String, ByVal psActualizacion As String, Optional ByVal psStockMin As Currency) As Integer
    Dim sActualiza As String
    Dim oConec As DConecta
    Dim rs As Recordset
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        sSql = "SELECT cBSCod FROM BienesServicios " & _
            " WHERE cBSCod = '" & psBSCod & "' "
        Set rs = oConec.CargaRecordSet(sSql)
        
        If psPorDeprecia = "" Then psPorDeprecia = "0"
        If psTipoDeprecia = "" Then psTipoDeprecia = "0"
        
        If rs.RecordCount = 0 Then
            sActualiza = Left(psActualizacion, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
            
            sSql = "INSERT INTO BienesServicios(cBSCod, cBSDescripcion, " & _
                "       nBSUnidad, bSerie, bVerifica, bContiene, bVigente, bCorrela, cUltimaActualizacion, nPorDeprecia, nBSTipDep,nStockMinimo) " & _
                " VALUES ('" & psBSCod & "','" & psDescripcion & "'," & _
                        pnUnidad & "," & IIf(pbSerie, 1, 0) & ", " & IIf(pbVerifica, 1, 0) & ", " & IIf(pbContenido, 1, 0) & "," & IIf(pbVigente, 1, 0) & "," & IIf(pbCorrela, 1, 0) & ",'" & sActualiza & "'," & psPorDeprecia & "," & psTipoDeprecia & "," & psStockMin & ")"
            oConec.Ejecutar sSql
            
            GrabaBS = 0
        Else
            GrabaBS = 2
        End If
        oConec.CierraConexion
    Else
        GrabaBS = 1
    End If
End Function

'Modifica un Bien/Servicio
Public Function ModificaBS(ByVal psBSCod As String, ByVal psDescripcion As String, _
ByVal pnUnidad As Integer, ByVal pbSerie As Boolean, ByVal pbVerifica As Boolean, ByVal pbContenido As Boolean, ByVal pbVigente As Boolean, ByVal pbCorrela As Boolean, ByVal psPorDeprecia As String, ByVal psTipoDeprecia As String, ByVal psActualizacion As String, ByVal psStockMin As Currency) As Integer
    Dim sActualiza As String
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        sSql = "SELECT cBSCod FROM BienesServicios " & _
            " WHERE cBSCod = '" & psBSCod & "' "
        Set rs = oConec.CargaRecordSet(sSql)
        
        If psPorDeprecia = "" Then psPorDeprecia = "0"
        If psTipoDeprecia = "" Then psTipoDeprecia = "0"
        
        If rs.RecordCount = 1 Then
            sActualiza = Left(psActualizacion, 8) & Format(oConec.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
            
            sSql = " UPDATE BienesServicios SET cBSDescripcion = '" & psDescripcion & "', " & _
                   "       nBSUnidad = " & pnUnidad & ", bSerie = " & IIf(pbSerie, 1, 0) & ", bVerifica = " & IIf(pbVerifica, 1, 0) & ", bContiene  = " & IIf(pbContenido, 1, 0) & ", bVigente = " & IIf(pbVigente, 1, 0) & ", bCorrela = " & IIf(pbCorrela, 1, 0) & ", " & _
                   "       cUltimaActualizacion = '" & sActualiza & "', nPorDeprecia = " & psPorDeprecia & ", nBSTipDep = " & psTipoDeprecia & ",nStockMinimo = " & psStockMin & "" & _
                   " WHERE cBSCod = '" & psBSCod & "'"
            oConec.Ejecutar sSql
            
                        
            sSql = " UPDATE BienesServicios SET " & _
                   "       cUltimaActualizacion = '" & sActualiza & "', nPorDeprecia = " & psPorDeprecia & ", nBSTipDep = " & psTipoDeprecia & ", bSerie = " & IIf(pbSerie, 1, 0) & ", bVerifica = " & IIf(pbVerifica, 1, 0) & ", bContiene  = " & IIf(pbContenido, 1, 0) & ", bVigente = " & IIf(pbVigente, 1, 0) & ", bCorrela = " & IIf(pbCorrela, 1, 0) & "" & _
                   " WHERE cBSCod Like '" & psBSCod & "%'"
            oConec.Ejecutar sSql
            
'            If Mid(psBSCod, 2, 1) = "2" Then
'                sSQL = "Update BSActivosFijos Set nPeriodoDep = " & psPeriodo & " WHERE cBSCod = '" & psBSCod & "'"
'                oConec.Ejecutar sSQL
'            End If
            
            ModificaBS = 0
        Else
            ModificaBS = 2
        End If
        oConec.CierraConexion
    Else
        ModificaBS = 1
    End If
End Function

'Elimina un Bien/Servicio y su relacion con ProveedorBS
Public Function EliminaBS(ByVal psBSCod As String) As Integer
    Dim oConec As DConecta
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        sSql = "DELETE ProveedorBS WHERE cBSCod = '" & psBSCod & "'"
        oConec.Ejecutar sSql
        
        sSql = "DELETE BienesServicios WHERE cBSCod = '" & psBSCod & "'"
        oConec.Ejecutar sSql
        
        oConec.CierraConexion
        EliminaBS = 0
    Else
        EliminaBS = 1
    End If
End Function

'Devuelve un nuevo código para los Bienes/Servicios
Public Function GeneraBSCodNue(ByVal psCodBSPadre As String) As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    Dim sValor As String
        
    If oConec.AbreConexion() Then
        sSql = "SELECT isnull(MAX(bs.cBSCod),0) Campo" & _
            " FROM BienesServicios BS WHERE bs.cBSCod LIKE '" & psCodBSPadre & "_%' AND " & _
            "   LEN(bs.cBSCod) = (SELECT MIN(LEN(cBSCod)) FROM BienesServicios WHERE cBSCod LIKE '" & psCodBSPadre & "_%')"
        
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        sValor = rs!Campo
        If Val(sValor) > 0 Then
            GeneraBSCodNue = Val(sValor) + 1
        Else
            If Len(psCodBSPadre) = 5 Then
                GeneraBSCodNue = psCodBSPadre & "001"
            Else
                GeneraBSCodNue = psCodBSPadre & "001"
            End If
        End If
    Else
        GeneraBSCodNue = ""
    End If
End Function

Public Function EsUltimoNivel(psBS As String) As Boolean
    Dim sqlBS As String
    Dim rsBS As ADODB.Recordset
    Set rsBS = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlBS = " Select cBSCod from BienesServicios Where cBSCod Like '" & psBS & "_%' And cBSCod <> '" & psBS & "'"
    
    If oCon.AbreConexion Then
        Set rsBS = oCon.CargaRecordSet(sqlBS)
        If rsBS.EOF And rsBS.EOF Then
            EsUltimoNivel = True
        Else
            EsUltimoNivel = False
        End If
    End If
    
    Set oCon = Nothing
    rsBS.Close
    Set rsBS = Nothing
End Function
'*** BRGO 20110801 ******************************************
Public Function ObtenerDatosBienAdjudicado(ByVal NumAdj As Integer) As ADODB.Recordset
    Dim oconect As New DConecta
    Dim sSql As String
    On Error GoTo RegBienAdjudErr
    oconect.AbreConexion
    
    sSql = "stp_sel_ObtenerBienAdjudicado " & NumAdj
    Set ObtenerDatosBienAdjudicado = oconect.CargaRecordSet(sSql)
    oconect.CierraConexion: Set oconect = Nothing
    Exit Function
RegBienAdjudErr:
   'Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Function
Public Function ObtenerListaBienesAdjudicado() As ADODB.Recordset

    Dim oconect As New DConecta
    Dim sSql As String
    On Error GoTo RegBienAdjudErr
    oconect.AbreConexion
    
    sSql = "stp_sel_ObtenerListaBienesAdjudicados '','',1"
    Set ObtenerListaBienesAdjudicado = oconect.CargaRecordSet(sSql)
    oconect.CierraConexion: Set oconect = Nothing
    Exit Function
RegBienAdjudErr:
   'Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Function

Public Function RegistrarTasaBienAdjudicado(ByVal BienAdjudCod As Integer, ByVal FecTasacion As Date, ByVal ValTasacion As Double, ByVal sMovNro As String)

    Dim oconect As New DConecta
    Dim sSql As String
    On Error GoTo RegBienAdjudErr
    oconect.AbreConexion
    
    sSql = "stp_ins_BienAdjudicadoTasacion " & BienAdjudCod & ",'" & Format(FecTasacion, "MM/dd/yyyy") & "'," & ValTasacion & ",'" & sMovNro & "'"
    oconect.CargaRecordSet (sSql)
    oconect.CierraConexion: Set oconect = Nothing
    Exit Function
RegBienAdjudErr:
   'Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Function
Public Function ListarTasasBienAdjudicado(ByVal BienAdjudCod As Integer) As ADODB.Recordset

    Dim oconect As New DConecta
    Dim sSql As String
    On Error GoTo RegBienAdjudErr
    oconect.AbreConexion
    
    sSql = "stp_sel_BienAdjudicadoTasacion " & BienAdjudCod
    Set ListarTasasBienAdjudicado = oconect.CargaRecordSet(sSql)
    oconect.CierraConexion: Set oconect = Nothing
    Exit Function
RegBienAdjudErr:
   'Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Function
'** END BRGO ***************************************************
'JUCS TI-ERS002-2017
Public Function RegistrarVentaBienAdjudicado(ByVal BienAdjudCod As Integer, ByVal FecVenta As Date, ByVal ValVenta As Double, ByVal sMovNro As String)

    Dim oconect As New DConecta
    Dim sSql As String
    On Error GoTo RegBienAdjudErr
    oconect.AbreConexion
    
    sSql = "stp_ins_ColocBienesAdjudVenta " & BienAdjudCod & ",'" & Format(FecVenta, "MM/dd/yyyy") & "'," & ValVenta & ",'" & sMovNro & "'"
    oconect.CargaRecordSet (sSql)
    oconect.CierraConexion: Set oconect = Nothing
    Exit Function
RegBienAdjudErr:
   'Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Function

Public Function ListarMontoVentaBienAdjudicado(ByVal BienAdjudCod As Integer) As ADODB.Recordset

    Dim oconect As New DConecta
    Dim sSql As String
    On Error GoTo RegBienAdjudErr
    oconect.AbreConexion
    
    sSql = "stp_sel_BienAdjudicadoHistVenta " & BienAdjudCod
    Set ListarMontoVentaBienAdjudicado = oconect.CargaRecordSet(sSql)
    oconect.CierraConexion: Set oconect = Nothing
    Exit Function
RegBienAdjudErr:
   'Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones: GetLibroDiario Method")
End Function
'END JUCS TI-ERS002-2017
