VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
'** Nombre : DBien
'** Descripción : Para manejo de Bienes Logistica creado segun ERS059-2013
'** Creación : EJVG, 20130520 17:00:00 PM
'********************************************************************
Option Explicit
Dim dbConec As DConecta
Dim lsSQL As String

Private Sub Class_Initialize()
    Set dbConec = New DConecta
    If Not dbConec.AbreConexion Then
        Call RaiseError(MyUnhandledError, "DBien:Initialize Method. Error en Conexion a Base de datos")
    End If
End Sub
Private Sub Class_Terminate()
    If Not dbConec Is Nothing Then
        dbConec.CierraConexion
        Set dbConec = Nothing
    End If
End Sub
Public Sub dBeginTrans()
    dbConec.BeginTrans
End Sub
Public Sub dRollbackTrans()
    dbConec.RollBackTrans
    dbConec.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub
Public Sub dCommitTrans()
    dbConec.CommitTrans
End Sub
Public Function ListarOrdenCompra(ByVal pnMoneda As Moneda) As ADODB.Recordset
    Set dbConec = New DConecta
    lsSQL = "Exec stp_sel_ERS0592013_ListarOrdenCompra '" & pnMoneda & "'"
    dbConec.AbreConexion
    Set ListarOrdenCompra = dbConec.CargaRecordSet(lsSQL)
    dbConec.CierraConexion
End Function
Public Function ListarOrdenCompraDet(ByVal pnMovNro As Long) As ADODB.Recordset
    Set dbConec = New DConecta
    lsSQL = "Exec stp_sel_ERS0592013_ListarOrdenCompraDet " & pnMovNro
    dbConec.AbreConexion
    Set ListarOrdenCompraDet = dbConec.CargaRecordSet(lsSQL)
    dbConec.CierraConexion
End Function
Public Sub ActivarBien(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, ByVal pnTpoBien As Integer)
    lsSQL = "Exec stp_upd_ERS0592013_ActivarBien " & pnMovNro & "," & pnMovItem & "," & pnTpoBien
    dbConec.Ejecutar (lsSQL)
End Sub
Public Function RecuperaCorrelativoAF(ByVal psBSCod As String, ByVal psAgeCod As String) As String
    Set dbConec = New DConecta
    Dim rs As New ADODB.Recordset
    lsSQL = "Exec stp_sel_ERS0592013_CorrelativoAF '" & psBSCod & "','" & psAgeCod & "'"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        RecuperaCorrelativoAF = rs!cCorrelativo
    End If
    Set rs = Nothing
    dbConec.CierraConexion
End Function
Public Function RecuperaCorrelativoBND(ByVal psAgeCod As String) As String
    Set dbConec = New DConecta
    Dim rs As New ADODB.Recordset
    lsSQL = "Exec stp_sel_ERS0592013_CorrelativoBND '" & psAgeCod & "'"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        RecuperaCorrelativoBND = rs!cCorrelativo
    End If
    Set rs = Nothing
    dbConec.CierraConexion
End Function
Public Function RecuperaCodigoBAN(ByVal psBSCod As String) As Integer
    Set dbConec = New DConecta
    Dim rs As New ADODB.Recordset
    lsSQL = "Select dbo.fnc_getBANCod('" & psBSCod & "') nCodBAN"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(lsSQL)
    RecuperaCodigoBAN = 0
    If Not rs.EOF Then
        RecuperaCodigoBAN = rs!nCodBAN
    End If
    Set rs = Nothing
    dbConec.CierraConexion
End Function
Public Function RecuperaSubCtaxBANCod(ByVal pnBANCod As Integer) As String
    Dim oConec As New DConecta
    Dim rs As New ADODB.Recordset
    Dim sSQL As String
    sSQL = "Select dbo.fnc_getSubCtaxBANCod(" & pnBANCod & ") cSubCtaCod"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    If Not rs.EOF Then
        RecuperaSubCtaxBANCod = rs!cSubCtaCod
    End If
    oConec.CierraConexion
    Set rs = Nothing
    Set oConec = Nothing
End Function
Public Function RecuperaMejoraComponentePaArbol() As ADODB.Recordset
    Set dbConec = New DConecta
    lsSQL = "Exec stp_sel_ERS0592013_ActivoCompuestoPaArbol"
    dbConec.AbreConexion
    Set RecuperaMejoraComponentePaArbol = dbConec.CargaRecordSet(lsSQL)
    dbConec.CierraConexion
End Function
Public Sub InsertaBienNoDepreciable(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnDocOrigen As Integer, ByVal psInventarioCod As String, ByVal psObjetoCod As String, ByVal psNombre As String, ByVal pnBSValor As Currency, ByVal pdActivacion As Date, ByVal psAreCod As String, ByVal psAgeCod As String, ByVal psMarca As String, ByVal psSerie As String, ByVal psModelo As String, ByVal pnMovNroAF As Long)
    lsSQL = "Exec stp_ins_ERS0592013_BienNoDepreciable " & pnMovNro & "," & pnMovItem & "," & pnDocOrigen & ",'" & psInventarioCod & "','" & psObjetoCod & "','" & psNombre & "'," & pnBSValor & ",'" & Format(pdActivacion, "yyyymmdd") & "','" & psAreCod & "','" & psAgeCod & "','" & psMarca & "','" & psSerie & "','" & psModelo & "'," & pnMovNroAF
    dbConec.Ejecutar (lsSQL)
End Sub
Public Function InsertaBienActivoCompuesto(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnDocOrigen As Integer, ByVal psInventarioCod As String, ByVal psObjetoCod As String, ByVal psNombre As String, ByVal pnBSValor As Currency, ByVal pdActivacion As Date, ByVal psAreCod As String, ByVal psAgeCod As String, ByVal psPersCod As String, ByVal psMarca As String, ByVal psSerie As String, ByVal psModelo As String) As Long
    Dim rs As New ADODB.Recordset
    lsSQL = "Exec stp_ins_ERS0592013_ActivoCompuesto " & pnMovNro & "," & pnMovItem & "," & pnDocOrigen & ",'" & psInventarioCod & "','" & psObjetoCod & "','" & psNombre & "'," & pnBSValor & ",'" & Format(pdActivacion, "yyyymmdd") & "','" & psAreCod & "','" & psAgeCod & "','" & psPersCod & "','" & psMarca & "','" & psSerie & "','" & psModelo & "'"
    Set rs = dbConec.CargaRecordSet(lsSQL)
    InsertaBienActivoCompuesto = rs!nId
    Set rs = Nothing
End Function
Public Sub InsertaBienActivoCompuestoDet(ByVal pnId As Long, ByRef pnItem As Integer, ByVal pnDocOrigen As Integer, ByVal nTpoComponente As Integer, ByVal psInventarioCod As String, ByVal psObjetoCod As String, ByVal psNombre As String, ByVal pnBSValor As Currency, ByVal pdActivacion As Date, ByVal pnDepreciaTributMes As Integer, ByVal pnDepreciaContabMes As Integer, ByVal psAreCod As String, ByVal psAgeCod As String, ByVal psPersCod As String, ByVal psMarca As String, ByVal psSerie As String, ByVal psModelo As String, ByVal pnMovNroAF As Long)
    Dim rs As New Recordset
    lsSQL = "Exec stp_ins_ERS0592013_ActivoCompuestoDet " & pnId & "," & pnItem & "," & pnDocOrigen & "," & nTpoComponente & ",'" & psInventarioCod & "','" & psObjetoCod & "','" & psNombre & "'," & pnBSValor & ",'" & Format(pdActivacion, "yyyymmdd") & "'," & pnDepreciaTributMes & "," & pnDepreciaContabMes & ",'" & psAreCod & "','" & psAgeCod & "','" & psPersCod & "','" & psMarca & "','" & psSerie & "','" & psModelo & "'," & pnMovNroAF
    Set rs = dbConec.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        pnItem = rs!nItem
    End If
    Set rs = Nothing
End Sub
Public Sub InsertaBienMejoraComponente(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, ByVal pnACId As Integer, ByVal pnACItem As Integer, ByVal pdMejora As Date)
    lsSQL = "Exec stp_ins_ERS0592013_MejoraComponente " & pnMovNro & "," & pnMovItem & "," & pnACId & "," & pnACItem & ",'" & Format(pdMejora, "yyyymmdd hh:mm:ss") & "'"
    dbConec.Ejecutar (lsSQL)
End Sub
Public Sub InsertaBienNoActivable(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnDocOrigen As Integer, ByVal psObjetoCod As String)
    lsSQL = "Exec stp_ins_ERS0592013_BienNoActivable " & pnMovNro & "," & pnMovItem & "," & pnDocOrigen & ",'" & psObjetoCod & "'"
    dbConec.Ejecutar (lsSQL)
End Sub
Public Sub InsertaAF(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovNroAF As Long)
    lsSQL = "Exec stp_ins_ERS0592013_AF " & pnMovNro & "," & pnMovItem & "," & pnMovNroAF
    dbConec.Ejecutar (lsSQL)
End Sub
Public Function RecuperaBienActivoCompuesto(ByVal pnId As Long) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_ActivoCompuesto " & pnId
    oConect.AbreConexion
    Set RecuperaBienActivoCompuesto = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function RecuperaCategoriasBienPaObjeto(ByVal pbAreaAgeTodos As Boolean, ByVal psAreaAgeCod As String) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_CategoriasBienMntxAreaAgePaObjeto " & IIf(pbAreaAgeTodos, 1, 0) & ",'" & psAreaAgeCod & "'"
    oConect.AbreConexion
    Set RecuperaCategoriasBienPaObjeto = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function RecuperaSeriesPaObjeto(ByVal psAreAgeCod As String, ByVal psBSCodCate As String, Optional ByVal pbBuscaxDeterioro As Boolean = False, Optional ByVal psAge As String = "") As ADODB.Recordset
    'NAGL 20191222 Agregó Optional ByVal psAge As String = ""
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_SeriesPaObjeto '" & psAreAgeCod & "','" & psBSCodCate & "'," & IIf(pbBuscaxDeterioro, 1, 0) & ",'" & psAge & "'"
    oConect.AbreConexion
    Set RecuperaSeriesPaObjeto = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function RecuperaSeriesBajadosPaObjeto(Optional ByVal psAreAgeCod As String = "", Optional ByVal psBSCodCate As String = "", Optional ByVal pbBuscaxDestino As Boolean = False) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_SeriesBajadosPaObjeto '" & psAreAgeCod & "','" & psBSCodCate & "'," & IIf(pbBuscaxDestino, 1, 0)
    oConect.AbreConexion
    Set RecuperaSeriesBajadosPaObjeto = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function RecuperaSeriesConActivaciones(ByVal psAreAgeCod As String, ByVal psBSCodCate As String) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_SeriesConActivaciones '" & psAreAgeCod & "','" & psBSCodCate & "'"
    oConect.AbreConexion
    Set RecuperaSeriesConActivaciones = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function RecuperaAFxMov(ByVal pnMovNro As Long) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_SeriesxMov " & pnMovNro
    oConect.AbreConexion
    Set RecuperaAFxMov = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function RecuperaBienxMantenimiento(ByVal pbAreaAgeTodos As Boolean, ByVal psAreaAgeCod As String, ByVal pbCategoriaTodos As Boolean, ByVal psCategoriaCod As String) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_BienxMantenimiento " & IIf(pbAreaAgeTodos, 1, 0) & ",'" & psAreaAgeCod & "'," & IIf(pbCategoriaTodos, 1, 0) & ",'" & psCategoriaCod & "'"
    oConect.AbreConexion
    Set RecuperaBienxMantenimiento = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function RecuperaBienxMantenimiento2(ByVal psAreaAgeCod As String, ByVal psCategoriaCod As String, ByVal psSerieCod As String) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_BienxMantenimiento2 '" & psAreaAgeCod & "','" & psCategoriaCod & "','" & psSerieCod & "'"
    oConect.AbreConexion
    Set RecuperaBienxMantenimiento2 = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function RecuperaBienxTransferencia(ByVal pbUnico As Boolean, ByVal pnMovNro As Long, ByVal pnId As Long) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_BienxTransferencia " & IIf(pbUnico, 1, 0) & "," & pnMovNro & "," & pnId
    oConect.AbreConexion
    Set RecuperaBienxTransferencia = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function RecuperaBienxAjusteVidaUtil(ByVal psBSCodCate As String, ByVal pnTpoBusqueda As Integer, ByVal psAreaAgeCod As String, Optional pdFecha As Date = "01/01/1900", Optional psCodSerie As String = "") As ADODB.Recordset
    'NAGL 20191023 Agregó pdFecha, psCodSerie
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_BienxAjusteVidaUtil '" & psBSCodCate & "'," & pnTpoBusqueda & ",'" & psAreaAgeCod & "','" & Format(pdFecha, "yyyymmdd") & "','" & psCodSerie & "'" 'NAGL 20191023 Agregó pdFecha, psCodSerie Según RFC1910190001
    oConect.AbreConexion
    Set RecuperaBienxAjusteVidaUtil = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function RecuperaHistorialVidaUtil(ByVal pnMovNro As Long) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_HistorialVidaUtil " & pnMovNro
    oConect.AbreConexion
    Set RecuperaHistorialVidaUtil = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function RealizaronDepreciacion(ByVal pnAnio As Integer, ByVal pnMes As Integer) As Boolean
    Dim oConect As New DConecta
    Dim rs As New ADODB.Recordset
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_RealizaronDepreciacion " & pnAnio & "," & pnMes
    RealizaronDepreciacion = False
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(sSQL)
    If Not rs.EOF Then
        RealizaronDepreciacion = rs!bExiste
    End If
    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
End Function
Public Function TieneDepreciacion(ByVal psSerie As String) As Boolean
    Dim oConect As New DConecta
    Dim rs As New ADODB.Recordset
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_TieneDepreciacion '" & psSerie & "'"
    TieneDepreciacion = False
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(sSQL)
    If Not rs.EOF Then
        TieneDepreciacion = rs!bExiste
    End If
    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
End Function
Public Function GetAFBienes(Optional ByVal psAreaAgeCod As String = "") As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_GetAFBSSerie '" & psAreaAgeCod & "'"
    oConect.AbreConexion
    Set GetAFBienes = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function GetAFBienesBaja(Optional ByVal psAreaAgeCod As String = "") As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_GetAFBSSerieBajas '" & psAreaAgeCod & "'"
    oConect.AbreConexion
    Set GetAFBienesBaja = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function GetAFBienesFull(ByVal psAreaAgeCod As String) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    'sSQL = "Exec stp_sel_ERS0592013_GetAFBSFull '" & psCodBien & "','" & psAreaAgeCod & "'"
    sSQL = "Exec stp_sel_ERS0592013_GetAFBSSerieFull '" & psAreaAgeCod & "'"
    oConect.AbreConexion
    Set GetAFBienesFull = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function GetAFBienesPaBaja(ByVal psAreaAgeCod As String, ByVal psCodBien As String, ByVal psSerie As String) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_GetAFBSSerieVigentes '" & psAreaAgeCod & "','" & psCodBien & "','" & psSerie & "'"
    oConect.AbreConexion
    Set GetAFBienesPaBaja = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function GetAFBienesPaDeterioro(ByVal psAreaAgeCod As String, ByVal psCodBien As String, ByVal psSerie As String) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_GetAFBSSeriePaDeterioro '" & psAreaAgeCod & "','" & psCodBien & "','" & psSerie & "'"
    oConect.AbreConexion
    Set GetAFBienesPaDeterioro = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function GetAFBienesBajadosPaDestino(ByVal psAreaAgeCod As String, ByVal psCodBien As String, ByVal psSerie As String) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_GetAFBSSerieBajadosPaDestino '" & psAreaAgeCod & "','" & psCodBien & "','" & psSerie & "'"
    oConect.AbreConexion
    Set GetAFBienesBajadosPaDestino = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function GetAFBienesPaKardex(ByVal psBSCod As String, ByVal pdFechaIni As Date, ByVal pdFechaFin As Date) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_GetAFBSSeriexKardex '" & psBSCod & "','" & Format(pdFechaIni, "yyyymmdd") & "','" & Format(pdFechaFin, "yyyymmdd") & "'"
    oConect.AbreConexion
    Set GetAFBienesPaKardex = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function GetAFDepreciacionxKardex(ByVal psSerie As String) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_DepreciacionxKardex '" & psSerie & "'"
    oConect.AbreConexion
    Set GetAFDepreciacionxKardex = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function GetAFDepreciacionxKardexConsol(ByVal psFechaIni As String, ByVal psFechaFin As String) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_DepreciacionxKardexConsol '" & psFechaIni & "','" & psFechaFin & "'"
    oConect.AbreConexion
    Set GetAFDepreciacionxKardexConsol = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Sub ActualizarAF(ByVal pnMovNro As Long, _
                            Optional ByVal psInventarioCod As String = "", Optional ByVal pdFechaIng As Date = "1900-01-01", _
                            Optional ByVal psNombre As String = "", Optional ByVal pnDepreTributMes As Integer = -1, Optional ByVal pnDepreContabMes As Integer = -1, _
                            Optional ByVal psAreaCod As String = "", Optional ByVal psAgeCod As String, Optional ByVal psPersCod As String = "", _
                            Optional ByVal psMarca As String = "", Optional ByVal psSerie As String = "", Optional psModelo As String = "")
    'Optional ByVal pbActivaAF As Boolean = False, Optional ByVal pnActivaAF_Id As Long = 0, _
                            Optional ByVal pbActivaAF_AC As Boolean = False, Optional ByVal pnActivaAF_AC_Id As Long = 0, Optional ByVal pnActivaAF_AC_Item As Long = 0, _
                            Optional ByVal pbActivaBND As Boolean = False, Optional ByVal pnActivaBND_Id As Long = 0, _
                            Optional ByVal pbActivaBND_AC As Boolean = False, Optional ByVal pnActivaBND_AC_Id As Long = 0, Optional ByVal pnActivaBND_AC_Item As Long = 0,
    Dim lsSQL As String
    If psInventarioCod <> "" Then
        lsSQL = lsSQL & " cSerie = '" & psInventarioCod & "',"
    End If
    If DateDiff("D", pdFechaIng, CDate("1900-01-01")) <> 0 Then
        lsSQL = lsSQL & " dActivacion = '" & Format(pdFechaIng, "yyyymmdd hh:mm:ss") & "',"
        '******Agregado by NAGL 20191222 Según RFC1910190001**********
        lsSQL = lsSQL & " dCompra = '" & Format(pdFechaIng, "yyyymmdd hh:mm:ss") & "',"
        lsSQL = lsSQL & " dBase = '" & Format(pdFechaIng, "yyyymmdd hh:mm:ss") & "',"
        '******************END NAGL***********************************
    End If
    If psNombre <> "" Then
        lsSQL = lsSQL & " cDescripcion = '" & psNombre & "',"
    End If
    If pnDepreTributMes <> -1 Then
        lsSQL = lsSQL & " nBSPerDepreciaTrib = " & pnDepreTributMes & ","
    End If
    If pnDepreContabMes <> -1 Then
        lsSQL = lsSQL & " nBSPerDeprecia = " & pnDepreContabMes & ","
    End If
    If psAreaCod <> "" Then
        lsSQL = lsSQL & " cAreCod = '" & psAreaCod & "',"
    End If
    If psAgeCod <> "" Then
        lsSQL = lsSQL & " cAgeCod = '" & psAgeCod & "',"
    End If
    If psPersCod <> "" Then
        lsSQL = lsSQL & " cPersCod = '" & psPersCod & "',"
    End If
    If psMarca <> "" Then
        lsSQL = lsSQL & " vMarca = '" & psMarca & "',"
    End If
    If psSerie <> "" Then
        lsSQL = lsSQL & " vSerie = '" & psSerie & "',"
    End If
    If psModelo <> "" Then
        lsSQL = lsSQL & " vModelo = '" & psModelo & "',"
    End If
    If Len(lsSQL) > 0 Then
        lsSQL = "UPDATE BSActivoFijo SET " & Left(lsSQL, Len(lsSQL) - 1) & " WHERE nMovNro = " & pnMovNro
        dbConec.Ejecutar (lsSQL)
    End If
End Sub
Public Sub ActualizarVidaUtilAF(ByVal pnMovNroAF As Long, ByVal pnMovNro As Long, ByVal pnDepreContabMes As Integer)
    lsSQL = "UPDATE BSActivoFijoVidaUtil SET nBSPerDeprecia = " & pnDepreContabMes & " WHERE nMovNroAF = " & pnMovNroAF & " AND nMovNro = " & pnMovNro
    dbConec.Ejecutar (lsSQL)
End Sub
Public Sub InsertaVidaUtilAF(ByVal pnMovNroAF As Long, ByVal pnMovNro As Long, ByVal pnDepreContabMes As Integer, Optional psMotivo As String = "")
    'NAGL Agregó 20191223
    'lsSQL = "INSERT INTO BSActivoFijoVidaUtil (nMovNroAF,nMovNro,nBSPerDeprecia,psMotivo) " & _
            '"VALUES (" & pnMovNroAF & "," & pnMovNro & "," & pnDepreContabMes & ")" 'Comentado by NAGL 20191223
    lsSQL = "Exec stp_ins_ActivoFijoVidaUtil " & pnMovNroAF & "," & pnMovNro & "," & pnDepreContabMes & ",'" & psMotivo & "'"
    'NAGL 20191223 Según RFC1910190001
    dbConec.Ejecutar (lsSQL)
End Sub
Public Sub ActualizarActivoCompuesto(ByVal pnId As Long, Optional ByVal psAreaCod As String = "", _
                                        Optional ByVal psAgeCod As String = "", Optional ByVal psPersCod As String = "")
    Dim lsSQL As String
    If psAreaCod <> "" Then
        lsSQL = lsSQL & " cAreaCod = '" & psAreaCod & "',"
    End If
    If psAgeCod <> "" Then
        lsSQL = lsSQL & " cAgeCod = '" & psAgeCod & "',"
    End If
    If psPersCod <> "" Then
        lsSQL = lsSQL & " cPersCod = '" & psPersCod & "',"
    End If
    If Len(lsSQL) > 0 Then
        lsSQL = "UPDATE BSActivoCompuesto SET " & Left(lsSQL, Len(lsSQL) - 1) & " WHERE nId = " & pnId
        dbConec.Ejecutar (lsSQL)
    End If
End Sub
Public Sub InsertaDestinoBajaAF(ByVal pnMovNro As Long, ByVal pnDestino As Integer, ByVal pdFecha As Date, ByVal psDescripcion As String)
    lsSQL = "stp_ins_ERS0592013_BajaDestino " & pnMovNro & "," & pnDestino & ",'" & Format(pdFecha, "yyyymmdd") & "','" & psDescripcion & "'"
    dbConec.Ejecutar (lsSQL)
End Sub
Public Sub InsertaDeterioroBien(ByVal pnMovNro As Long, ByVal pdFecha As Date, ByVal psGlosa As String)
    lsSQL = "stp_ins_ERS0592013_BajaDestino " & pnMovNro & ",'" & Format(pdFecha, "yyyymmdd") & "','" & psGlosa & "'"
    dbConec.Ejecutar (lsSQL)
End Sub
Public Sub ActualizaDestinoBajaAF(ByVal pnMovNro As Long, ByVal pnDestino As Integer, ByVal pdFecha As Date, ByVal psDescripcion As String)
    lsSQL = "stp_upd_ERS0592013_BajaDestino " & pnMovNro & "," & pnDestino & ",'" & Format(pdFecha, "yyyymmdd") & "','" & psDescripcion & "'"
    dbConec.Ejecutar (lsSQL)
End Sub

'*** PEAC 20140530
Public Function RecuperaAreaBien(ByVal pbAreaAgeTodos As Boolean, ByVal psAreaAgeCod As String) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ERS0592013_CategoriasBienMntxAreaAgePaObjeto " & IIf(pbAreaAgeTodos, 1, 0) & ",'" & psAreaAgeCod & "'"
    oConect.AbreConexion
    Set RecuperaAreaBien = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function

'*** PEAC 20140602
Public Function RecuperaAreaAgenciaBienParaTransf(ByVal pcSerieBien As String) As ADODB.Recordset
    Dim oConect As New DConecta
    Dim sSQL As String
    sSQL = "Exec stp_sel_ObtieneSeriesConActivaciones '" & pcSerieBien & "' "
    oConect.AbreConexion
    Set RecuperaAreaAgenciaBienParaTransf = oConect.CargaRecordSet(sSQL)
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function ValidaDepreciacionDia(ByVal pnMovNro As Long, pdFecha As Date) As Boolean
    Dim oConect As New DConecta
    Dim rs As New ADODB.Recordset
    Dim sSQL As String
    sSQL = "Exec stp_sel_ValidaDepreciacionDia '" & pnMovNro & "','" & Format(pdFecha, "yyyymmdd") & "'"
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(sSQL)
    ValidaDepreciacionDia = rs!bExiste
    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
End Function 'NAGL 20191222 Según RFC1910190001

Public Function HabilitaAjusteVidaUtil() As Boolean
    Dim oConect As New DConecta
    Dim rs As New ADODB.Recordset
    Dim sSQL As String
    sSQL = "Exec stp_sel_HabilitaModAjusteVidaUtil"
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(sSQL)
    HabilitaAjusteVidaUtil = rs!bHabilita
    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
End Function 'NAGL 20191222 Según RFC1910190001

Public Sub ActualizaVidaUtilAF_NewGlosa(ByVal pnMovNroAF As Long, ByVal pdFecha As Date, ByVal pnDepreContabMes As Integer, Optional psMotivo As String = "")
    lsSQL = "Exec stp_upd_ActivoFijoVidaUtil " & pnMovNroAF & ",'" & Format(pdFecha, "yyyymmdd") & "'," & pnDepreContabMes & ",'" & psMotivo & "'"
    dbConec.Ejecutar (lsSQL)
End Sub 'NAGL 20191223 Según RFC1910190001
