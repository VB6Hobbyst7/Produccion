VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCatalogoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'NAGL ERS 042 - CATALOGO DE PRODUCTOS
Public Function CargarTipoProd(ByVal psTipo As String, Optional psRelac As String = "", Optional psVig As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_Catalogo_CargarProdCred '" & psTipo & "', '" & psRelac & "','" & psVig & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set CargarTipoProd = rs
    Set oConec = Nothing
End Function

Public Function ObtieneRelacProdEquiv(ByVal psTpoCredRelac As String, Optional psTpo As String = "") As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_Catalogo_ObtieneRelacCatalogoEquiv '" & psTpoCredRelac & "', '" & psTpo & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set ObtieneRelacProdEquiv = rs
    Set oConec = Nothing
End Function

Public Function ObtieneCorrelativoNewProd(ByVal psRelac As String) As Integer
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    Dim rs As New ADODB.Recordset
    Dim nCorr As Integer
    sSql = "Exec stp_sel_Catalogo_CorrelativoProdNew '" & psRelac & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    If Not rs.BOF And Not rs.EOF Then
        nCorr = rs!nCorrelativo
    Else
        nCorr = 0
    End If
    oConec.CierraConexion
    ObtieneCorrelativoNewProd = nCorr
    Set oConec = Nothing
End Function

Public Function CargaTiposModulos() As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_Catalogo_CargarTipoModulo"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set CargaTiposModulos = rs
    Set oConec = Nothing
End Function

Public Sub EliminaNewProducto(ByVal psTpoProd As String)
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "Exec stp_del_Catalogo_NewProducto '" & psTpoProd & "'"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Sub
End Sub

Public Function HabilitarProducto(ByVal psTpoProd As String) As String
   Dim sSql As String
   Dim cValHab As String
   Dim oConec As New COMConecta.DCOMConecta
   Dim rs As New ADODB.Recordset
   sSql = "Exec stp_upd_Catalogo_HabilitarProducto '" & psTpoProd & "'"
   oConec.AbreConexion
   Set rs = oConec.CargaRecordSet(sSql)
   If Not rs.BOF And Not rs.EOF Then
        cValHab = rs!psHabilitar
   Else
        cValHab = "No"
   End If
   oConec.CierraConexion
   HabilitarProducto = cValHab
   Set oConec = Nothing
End Function

Public Function DeshabilitarMigracionRel(ByVal psTpoProdAnt As String, ByVal psTpoProdNew As String) As String
   Dim sSql As String
   Dim cValDeshabRel As String
   Dim oConec As New COMConecta.DCOMConecta
   Dim rs As New ADODB.Recordset
   sSql = "Exec stp_upd_DeshabMigracion '" & psTpoProdAnt & "','" & psTpoProdNew & "'"
   oConec.AbreConexion
   Set rs = oConec.CargaRecordSet(sSql)
   If Not rs.BOF And Not rs.EOF Then
        cValDeshabRel = rs!psDeshabMigracion
   Else
        cValDeshabRel = "No"
   End If
   oConec.CierraConexion
   DeshabilitarMigracionRel = cValDeshabRel
   Set oConec = Nothing
End Function

Public Function ObtieneValRegNewProd(ByVal psConsValor As String, Optional psTipoConsul As String = "") As String
   Dim sSql As String
   Dim cValReg As String
   Dim oConec As New COMConecta.DCOMConecta
   Dim rs As New ADODB.Recordset
   sSql = "Exec stp_sel_Catalogo_ObtieneValConstProd '" & psConsValor & "','" & psTipoConsul & "'"
   oConec.AbreConexion
   Set rs = oConec.CargaRecordSet(sSql)
   If Not rs.BOF And Not rs.EOF Then
        cValReg = rs!psRegistro
   Else
        cValReg = "No"
   End If
   oConec.CierraConexion
   ObtieneValRegNewProd = cValReg
   Set oConec = Nothing
End Function 'Para observar si se puede proceder con la Eliminación/Edición

Public Sub GrabarNewProducto(ByVal pMatrizListRelProd As Variant, ByVal pMatrizNewProd As Variant, ByVal psMovNro As String) '(ByVal psTpoProdAnt As String, ByVal psTpoProdNew As String, ByVal psDescripProdNew As String, ByVal psMovNro As String)
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    Dim i As Integer
    Dim psTpoProdAnt As String, psTpoProdNew As String
    Dim psDescripProdNew As String, psTpoIns As String, psMovNrOp As String
    oConec.AbreConexion
    
    For i = 1 To UBound(pMatrizNewProd)
        psTpoProdAnt = ""
        psTpoProdNew = pMatrizNewProd(i, 0)
        psDescripProdNew = pMatrizNewProd(i, 1)
        psMovNrOp = ""
        psTpoIns = "NewProd"
        sSql = "Exec stp_ins_Catalogo_NewProductoEquiv '" & psTpoProdAnt & "','" & psTpoProdNew & "','" & psDescripProdNew & "','" & psMovNrOp & "', '" & psTpoIns & "'"
        oConec.Ejecutar sSql
    Next i

    For i = 0 To UBound(pMatrizListRelProd) - 1
        psTpoProdAnt = pMatrizListRelProd(i, 0)
        psTpoProdNew = pMatrizListRelProd(i, 1)
        psDescripProdNew = pMatrizListRelProd(i, 2)
        psTpoIns = ""
        sSql = "Exec stp_ins_Catalogo_NewProductoEquiv '" & psTpoProdAnt & "','" & psTpoProdNew & "','" & psDescripProdNew & "','" & psMovNro & "','" & psTpoIns & "'"
        oConec.Ejecutar sSql
    Next i
    
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Sub
End Sub

'Condiciones
Public Function CargarCboCondiciones(ByVal pnParCod As Long, Optional psDescrip As String = "") As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_Catalogo_CargaCondiciones " & pnParCod & ",'" & psDescrip & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set CargarCboCondiciones = rs
    Set oConec = Nothing
End Function

Public Function ObtieneParametroPrincipal(ByVal pnParValor As Long, ByVal psParDescrip As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_Catalogo_ObtieneParametroPrincipal " & pnParValor & ",'" & psParDescrip & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set ObtieneParametroPrincipal = rs
    Set oConec = Nothing
End Function

Public Function ObtieneListaParametro(Optional psParam As String, Optional pnTpoParam As Long, Optional psEvalRang As String) As ADODB.Recordset  'psHabRang
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_Catalogo_ListarParametros '" & psParam & "'," & pnTpoParam & ",'" & psEvalRang & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set ObtieneListaParametro = rs
    Set oConec = Nothing
End Function

'************GRABAR DATOS DE LAS CONDICIONES Y REQUISITOS*************
Public Function ObtieneCodIdCondRequi(ByVal psTipo As String, ByVal psParProd As String, ByVal pnParCod As Long, ByVal pnValor As Long, ByVal psOperInicio As String, ByVal psRangoInicio As String, ByVal psOperFin As String, ByVal psRangoFin As String, ByVal pnUndRango As Long, ByVal pnIdModulo As Long, psMovNro As String) As Long
   Dim sSql As String
   Dim pnCodIdDepen As Long
   Dim oConec As New COMConecta.DCOMConecta
   Dim rs As New ADODB.Recordset
   sSql = "Exec stp_sel_Catalogo_ObtieneCodIdCondRequ '" & psTipo & "','" & psParProd & "'," & pnParCod & "," & pnValor & ",'" & psOperInicio & "','" & psRangoInicio & "','" & psOperFin & "','" & psRangoFin & "'," & pnUndRango & ", " & pnIdModulo & ",'" & psMovNro & "'"
   oConec.AbreConexion
   Set rs = oConec.CargaRecordSet(sSql)
   If Not rs.BOF And Not rs.EOF Then
        pnCodIdDepen = rs!nCodIdDepen
   Else
        pnCodIdDepen = 0
   End If
   oConec.CierraConexion
   ObtieneCodIdCondRequi = pnCodIdDepen
   Set oConec = Nothing
End Function
Public Sub GrabarDatosCondicionRequisitos(ByVal optCondRequ As String, ByVal pnCodIdDepen As Long, ByVal psTpoProd As String, ByVal pnParCod As Long, ByVal pnParValor As Long, ByVal psRang1 As String, ByVal psRang2 As String, ByVal psRang3 As String, ByVal psRang4 As String, ByVal pnUnidRango As Long, ByVal pnIdModulo As Long, ByVal psMovNro As String)
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "Exec stp_ins_Catalogo_CondicionesRequisitos '" & optCondRequ & "', " & pnCodIdDepen & ",'" & psTpoProd & "'," & pnParCod & "," & pnParValor & ",'" & psRang1 & "','" & psRang2 & "','" & psRang3 & "','" & psRang4 & "'," & pnUnidRango & "," & pnIdModulo & ",'" & psMovNro & "'"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Sub
End Sub

Public Sub LimpiarDataProducto(ByVal optCondRequ As String, ByVal psTpoProd As String, ByVal psMovNroHist As String) 'NAGL 20190923 Agregó psMovNroHist
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    oConec.AbreConexion
     sSql = "Exec stp_del_Catalogo_CondicionesRequisitos '" & optCondRequ & "', '" & psTpoProd & "', '" & psMovNroHist & "'" 'NAGL 20190923 Agregó psMovNroHist
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Sub
End Sub

Public Sub RegresaDatosAnteriorProducto(ByVal optCondRequ As String, ByVal psTpoProd As String)
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "Exec stp_ins_Catalogo_RegresaDatosAnteriorProducto '" & optCondRequ & "', '" & psTpoProd & "'"
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Sub
End Sub '***NAGL 20190923
'*******************************************************************

Public Function GetListadoCondicionesRequisitos(ByVal psTipoCarac As String, ByVal psParProd As String, Optional psResultCols As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    If psTipoCarac = "Cond" Then
        sSql = "Exec stp_sel_Catalogo_ObtieneListaCondicionesxProducto '" & psParProd & "','" & psResultCols & "'"
    Else
        sSql = "Exec stp_sel_Catalogo_ObtieneListaRequisitosxProducto '" & psParProd & "','" & psResultCols & "'"
    End If
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set GetListadoCondicionesRequisitos = rs
    Set oConec = Nothing
End Function

Public Function ObtieneParamDescripInt(ByVal pnTipoParam As Long, ByVal pnTipoParamDet As Long) As String
   Dim sSql As String
   Dim cValDet As String
   Dim oConec As New COMConecta.DCOMConecta
   Dim rs As New ADODB.Recordset
   sSql = "Exec stp_sel_Catalogo_ObtieneParametroDescripDet " & pnTipoParam & "," & pnTipoParamDet & ""
   oConec.AbreConexion
   Set rs = oConec.CargaRecordSet(sSql)
   If Not rs.BOF And Not rs.EOF Then
        cValDet = rs!DescParam
   Else
        cValDet = ""
   End If
   oConec.CierraConexion
   ObtieneParamDescripInt = cValDet
   Set oConec = Nothing
End Function

Public Function CargaMensajesCondiciones(ByVal psTpoCod As String) As String
   Dim sSql As String
   Dim psReg As String
   Dim cMsgCond As String
   Dim oConec As New COMConecta.DCOMConecta
   Dim rs As New ADODB.Recordset
   sSql = "Exec stp_sel_CatProd_CargaMensajesCondiciones '" & psTpoCod & "'"
   oConec.AbreConexion
   Set rs = oConec.CargaRecordSet(sSql)
   If Not rs.BOF And Not rs.EOF Then
        cMsgCond = rs!Mensaje
   Else
        cMsgCond = ""
   End If
   oConec.CierraConexion
   CargaMensajesCondiciones = cMsgCond
   Set oConec = Nothing
End Function

'****PARA LA CONFIG.TASAS
Public Function ObtieneCaractCondRequixProducto(ByVal psTpoProd As String, ByVal pdFecha As Date)
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_Catalogo_GetCondReqGeneralxProducto '" & psTpoProd & "','" & Format(pdFecha, "yyyymmdd") & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set ObtieneCaractCondRequixProducto = rs
    Set oConec = Nothing
End Function


'JOEP
Public Sub GrabarDatosCheckList(ByVal pnTipo As Integer, ByVal pnCodCatg As Integer, ByVal pnCodProd As Integer, ByVal pcItem As String, Optional ByVal pcDescrp As String = "-1", Optional ByVal pnCantConf As Integer = -1, Optional ByVal pnNivel As Integer = -1, Optional ByVal pnTpDoc As Integer = -1, Optional ByVal pnEstCheck As Integer = -1, Optional ByVal pnParCod As Long = -1, Optional ByVal pnParValor As Long = -1, Optional ByVal pcOpeIni As String = "-1", Optional ByVal pnMontoIni As Currency = -1, Optional ByVal pcOpeFin As String = "-1", Optional ByVal pnMontoFin As Currency = -1, Optional ByVal pnUnidMed As Integer = -1, Optional ByVal pcNroMov As String, Optional ByVal pnEstado As Integer)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim oBase As COMDCredito.DCOMCredActBD
Set oBase = New COMDCredito.DCOMCredActBD
       sSql = "exec stp_ins_CatalogoCheckListGuardaDocDetCond " & pnTipo & "," & pnCodCatg & "," & pnCodProd & ",'" & pcItem & "','" & pcDescrp & "'," & pnCantConf & "," & pnNivel & "," & pnTpDoc & "," & pnEstCheck & "," & pnParCod & "," & pnParValor & ",'" & pcOpeIni & "'," & pnMontoIni & ",'" & pcOpeFin & "'," & pnMontoFin & "," & pnUnidMed & ",'" & pcNroMov & "'," & pnEstado & ""
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
Set oConecta = Nothing
End Sub
Public Sub EliminaDatosCheckList(ByVal pnTipo As Integer, ByVal pnCodCatg As Integer, ByVal pnCodProd As Integer, Optional ByVal pcItem As String, Optional ByVal pnCantConf As Integer)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim oBase As COMDCredito.DCOMCredActBD
Set oBase = New COMDCredito.DCOMCredActBD
       sSql = "exec stp_del_CatalogoCheckListEliminaDocDetCond " & pnTipo & "," & pnCodCatg & "," & pnCodProd & ",'" & pcItem & "'," & pnCantConf & ""
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
Set oConecta = Nothing
End Sub
Public Function ObtenerItemCheckList(ByVal pnItem As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim oBase As COMDCredito.DCOMCredActBD
Set oBase = New COMDCredito.DCOMCredActBD
       sSql = "exec stp_sel_CatalogoCheckListObtenerItem " & pnItem & ""
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
   Set ObtenerItemCheckList = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Function CargaFlexCheckList(ByVal pnTipo As Integer, ByVal pnCodCatg As Integer, ByVal pnCodProd As Integer, Optional ByVal pnCantConfi As Integer = 0, Optional ByVal pnItem As String = "") As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_CatalogoCheckListObtieneDocDetCod " & pnTipo & "," & pnCodCatg & "," & pnCodProd & "," & pnCantConfi & ",'" & pnItem & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set CargaFlexCheckList = rs
    Set oConec = Nothing
End Function

Public Function ObtieneListaParametroCheckList(ByVal pnOpcion As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_CatalogoCheckListListaParametro " & pnOpcion & ""
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set ObtieneListaParametroCheckList = rs
    Set oConec = Nothing
End Function
'JOEP
