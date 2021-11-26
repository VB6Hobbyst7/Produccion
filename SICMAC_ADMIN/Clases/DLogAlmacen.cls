VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogAlmacen"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsCentralPer As String
Dim vsCentralCom As String
Dim vsCentralImg As String
Dim sSQL As String

Public Enum TpoCargaC
    CUnRegistro = 1
    CTodosEstado = 5
End Enum

Public Enum TpoCargaA
    AUnRegistro = 1
    ATodos = 5
    ATodosMenosUno = 6
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

'Devuelva un RecordSet con las Contrataciones
Public Function CargaContratacion(ByVal pnTpoCarga As TpoCargaC, _
Optional ByVal pnConNro As Long, Optional ByVal pnEstado As LogConEstado) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        If pnTpoCarga = CUnRegistro Then
            sSQL = "SELECT p.cPersNombre, m.cConsDescripcion, s.nLogSelTipCambio, " & _
                "       s.nLogSelNro, s.nLogSelCotNro " & _
                " FROM LogContratacion C " & _
                "       JOIN LogSeleccion S ON c.nLogSelNro = s.nLogSelNro " & _
                "       LEFT JOIN Constante M ON s.nLogSelMoneda = m.nConsValor " & _
                "           AND m.nConsCod = " & gMoneda & " " & _
                "       JOIN LogSelCotiza SC ON s.nLogSelCotNro = sc.nLogSelCotNro " & _
                "           AND s.nLogSelNro = sc.nLogSelNro " & _
                "       LEFT JOIN Persona P ON sc.cPersCod = p.cPersCod " & _
                " WHERE nLogConNro = " & pnConNro & " "
        ElseIf pnTpoCarga = CTodosEstado Then
            sSQL = "SELECT m.cMovNro Código, c.dLogConReg Orden, p.cPersNombre Proveedor" & _
                " FROM Mov M JOIN LogContratacion C On m.nMovNro = c.nLogConNro " & _
                "       JOIN LogSeleccion S ON c.nLogSelNro = s.nLogSelNro " & _
                "       JOIN LogSelCotiza SC ON s.nLogSelCotNro = sc.nLogSelCotNro " & _
                "           AND s.nLogSelNro = sc.nLogSelNro " & _
                "       LEFT JOIN Persona P ON sc.cPersCod = p.cPersCod " & _
                " WHERE c.nLogConEstado IN (" & pnEstado & ")"
        End If
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If

    Set CargaContratacion = rs
    Set rs = Nothing
End Function

'Devuelve un RecordSet con LogConDetalle (LogSeleccion - LogSelCotiza - LogSelCotDetalle)
Public Function CargaConDetalle(ByVal pnConNro As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSQL = "SELECT cd.cBSCod, bs.cBSDescripcion, " & _
            "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
            "       cd.nLogConDetPrecio, cd.nLogConDetCantidad " & _
            "       " & _
            " FROM LogConDetalle CD LEFT JOIN BienesServicios BS ON cd.cBSCod = bs.cBSCod " & _
            "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = " & gUnidadMedida & " " & _
            " WHERE cd.nLogConNro = " & pnConNro & " " & _
            " ORDER BY cd.cBSCod "
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    Set CargaConDetalle = rs
End Function

'Operación - Documento
Public Function CargaOpeDoc(ByVal pnOpeCod As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSQL = " SELECT a.nDocTpo, d.cDocDesc " & _
            " FROM " & vsCentralCom & "OpeDoc a JOIN " & vsCentralCom & "Documento D ON d.nDocTpo = a.nDocTpo " & _
            " WHERE cOpeCod = " & pnOpeCod & " "
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    Set CargaOpeDoc = rs
End Function

'Documento - Impuesto
Public Function CargaDocImp(ByVal pnTpoDoc As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSQL = " SELECT di.cCtaContCod, '',i.cImpAbrev, i.nImpTasa " & _
            " FROM " & vsCentralCom & "DocImpuesto DI JOIN " & vsCentralCom & "Impuesto I ON di.cCtaContCod = i.cCtaContCod " & _
            " WHERE di.nDocTpo = " & pnTpoDoc & " "
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    Set CargaDocImp = rs
End Function

Public Function CargaSubasta(Optional ByVal pnEstadoDoc As Integer = 0) As ADODB.Recordset
    Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sql As String
    
    sql = " Select MD.cDocNro, PE.cPersNombre + '[' + Str(MD.nMovNro) + ']', 2 Nivel from BSSubasta BSS" _
         & " Inner Join MovDoc MD On BSS.nMovNro = MD.nMovNro" _
         & " Inner Join BienesServicios BS On  BS.cBSCod = BSS.cBSCod" _
         & " Inner Join MovGasto MG On BSS.nMovNro = MG.nMovNro" _
         & " Inner Join Persona  PE On PE.cPersCod = MG.cPersCod" _
         & " Left  Join BSSubastaDet BSSD On BSS.nMovNro = BSSD.nMovNro And BSS.cBSCod = BSSD.cBSCod" _
         & " Where nSubEstado = " & pnEstadoDoc & ""
    oConec.AbreConexion
    
    Set CargaSubasta = oConec.CargaRecordSet(sql)
    Set oConec = Nothing
End Function

Public Function CargaBSSubasta() As ADODB.Recordset
    Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sql As String
    
    sql = " Select BSS.cBSCod, Rtrim(BS.cBSDescripcion)+ ' {LOTE NRO = ' + Rtrim(LTrim(BSS.nMovNro)) + '}-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])' Descripcion , 1" _
        & " From BSSubasta BSS Inner Join BienesServicios BS On BS.cBSCod = BSS.cBSCod" _
        & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCOd = '1019'"
    oConec.AbreConexion
    
    Set CargaBSSubasta = oConec.CargaRecordSet(sql)
    Set oConec = Nothing
End Function

Public Function CargaBSSubastaStock(psBSCod As String, pnMovNro As Long) As ADODB.Recordset
    Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sql As String
    
    sql = " Select  nVentaValor / nVentaCant Precio, nVentaCant Cantidad From BSSubasta " _
        & " Where nMovNro =  " & pnMovNro & " And cBSCod = '" & psBSCod & "'"
    oConec.AbreConexion
    
    Set CargaBSSubastaStock = oConec.CargaRecordSet(sql)
    Set oConec = Nothing
End Function


Public Function CargaBSSubastaSerie() As ADODB.Recordset
    Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sql As String
    
    sql = " Select BSS.cBSCod, Rtrim(BS.cBSDescripcion)+ ' {LOTE NRO = ' + Rtrim(LTrim(BSS.nMovNro)) + '}-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])' Descripcion , 1" _
        & " From BSSubasta BSS Inner Join BienesServicios BS On BS.cBSCod = BSS.cBSCod" _
        & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCOd = '1019'"
    oConec.AbreConexion
    
    Set CargaBSSubastaSerie = oConec.CargaRecordSet(sql)
    Set oConec = Nothing
End Function


Public Function CargaSubastaDetalle(Optional ByVal pnMovNro As Integer = 0) As ADODB.Recordset
    Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sql As String
    
    sql = " Select BSS.cBSCod, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])' Descripcion,nAdjudicaValor,nAdjudicaCant from BSSubasta BSS" _
        & " Inner Join MovDoc MD On BSS.nMovNro = MD.nMovNro" _
        & " Inner Join BienesServicios BS On  BS.cBSCod = BSS.cBSCod" _
        & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCOd = '1019'" _
        & " Where BSS.nMovNro = " & pnMovNro
    oConec.AbreConexion
    
    Set CargaSubastaDetalle = oConec.CargaRecordSet(sql)
    Set oConec = Nothing
End Function

Public Function CargaSubastaDetalleSerie(Optional ByVal pnMovNro As Integer = 0) As ADODB.Recordset
    Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sql As String
    
    sql = " Select BS.cBSCod, BSSD.cSerie from BSSubasta BSS" _
        & " Inner Join MovDoc MD On BSS.nMovNro = MD.nMovNro" _
        & " Inner Join BienesServicios BS On  BS.cBSCod = BSS.cBSCod" _
        & " Left  Join BSSubastaDet BSSD On BSS.nMovNro = BSSD.nMovNro And BSS.cBSCod = BSSD.cBSCod" _
        & " Where BSS.nMovNro = " & pnMovNro
    oConec.AbreConexion
    
    Set CargaSubastaDetalleSerie = oConec.CargaRecordSet(sql)
    Set oConec = Nothing
End Function

Public Function CargaStock(ByVal pdFecha As Date, pnProdTpo As Integer, ByVal pnAlmacenCod As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sql As String
    
    sql = " Select cBSCod, cBSDescripcion, nMonto, nStock, Cta, cCtaContDesc From (" _
        & " Select BS.cBSCod, BBSS.cBSDescripcion, nMonto, nStock, (Select Top 1 replace(cCtaContCod,'AG','" & Format(pnAlmacenCod, "00") & "') From CtaBS Where BS.cBSCod Like cObjetoCod + '%' And cOpeCod In (Select cOpeCod From LogOpeProd Where nProdTpo = " & pnProdTpo & " ))  Cta From BSSaldos BS" _
        & " Inner Join BienesServicios BBSS On BS.cBsCod = BBSS.cBSCod" _
        & " Where dSaldo = '" & Format(pdFecha, gsFormatoFecha) & "' and nAlmCod = " & pnAlmacenCod & ") As AA" _
        & " Inner Join CtaCont Cta On AA.Cta = Cta.cCtaContCod order by cta"
    oConec.AbreConexion
    
    Set CargaStock = oConec.CargaRecordSet(sql)
    Set oConec = Nothing
End Function

'Carga Almacenes
Public Function CargaAlmacen(ByVal pnTpoCarga As TpoCargaA, _
Optional ByVal pnAlmacen As Integer = 0) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If pnTpoCarga = ATodos Then
            'Todos almacenes
            sSQL = " SELECT  cAlmDescripcion , nAlmCod " & _
                " FROM Almacen "
        ElseIf pnTpoCarga = ATodosMenosUno Then
            'Todos menos uno (para tranferencia)
            sSQL = " SELECT  cAlmDescripcion , nAlmCod " & _
                " FROM Almacen " & _
                " WHERE nAlmCod <> " & pnAlmacen & ""
        ElseIf pnTpoCarga = AUnRegistro Then
            'Lista detalle de almacen
            sSQL = " SELECT nAlmCod, cAlmDescripcion, cAlmDireccion " & _
                " FROM Almacen " & _
                " WHERE nAlmCod = " & pnAlmacen & ""
        Else
            sSQL = " "
        End If
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    Set CargaAlmacen = rs
End Function

'Almacen - BS+

Public Function CargaAlmacenBS(ByVal pnAlmacen As Integer, ByVal pnAlmacenTpo As Integer, _
    ByVal pdFecha As Date, Optional ByVal pbStockCero As Boolean = True, Optional pbValidaInconsistencias As Boolean = False, Optional pnValor As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim tmpSql  As String
    Dim TempQ As String

    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    'CUANDO SE PIDE CON STOCK CERO
    If pbValidaInconsistencias Then
        tmpSql = " Where  ( (IsNull(Ini.nStock, 0) + IsNull(ING.Cant, 0) - IsNull(SAL.Cant, 0) <> 0 And IsNull(INI.nMonto,0) + IsNull(ING.Valor,0) + IsNull(SAL.Valor,0) <= 0) Or (IsNull(Ini.nStock, 0) + IsNull(ING.Cant, 0) - IsNull(SAL.Cant, 0) = 0 And IsNull(INI.nMonto,0) + IsNull(ING.Valor,0) + IsNull(SAL.Valor,0) <> 0) )  And len(bs.cbscod) = (Select Max(len(cbscod)) From  BienesServicios)   "
    Else
        tmpSql = IIf(pbStockCero = False, " Where IsNull(Ini.nStock, 0) + IsNull(ING.Cant, 0) - IsNull(SAL.Cant, 0) <> 0 And len(bs.cbscod) = (Select Max(len(cbscod)) From  BienesServicios)   ", " Where len(bs.cbscod) = (Select Max(len(cbscod)) From  BienesServicios) ")
    End If

    Dim ldFecha As Date
    Dim lsFecha As String

    If oConec.AbreConexion() Then
        sSQL = " Select * From (Select BS.cBSCod, BS.cBSDescripcion, IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, dbo.LogGetStock(" & pnAlmacen & ",cBsCod," & pnAlmacenTpo & ") nAlmBSStock, dbo.LogGetPrecioPromedio(1,cBsCod," & pnAlmacenTpo & ") nAlmBSPrePromedio From BienesServicios BS" _
             & " LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = 1019" _
             & " where Len(BS.cBSCod)  = (Select Max(len(cbscod)) From  BienesServicios) ) As AAA " & tmpSql & " Order by cBSCod"

        sSQL = " Select max(dsaldo) Fecha,  convert(varchar(8),dateadd(day,1,  max(dsaldo)),112) lsFecha From BSSaldos  WITH (NOLOCK) Where dSaldo <= '" & Format(pdFecha, gsFormatoFecha) & "' And nAlmCod = " & pnAlmacen & ""
        Set rs = oConec.CargaRecordSet(sSQL)

        If Not rs.EOF And Not rs.BOF Then

            If IsNull(rs!fecha) Then
                ldFecha = DateAdd("m", -1, (pdFecha + 1)) - 1
                lsFecha = Format(DateAdd("m", -1, (pdFecha + 1)), gsFormatoMovFecha)
            Else
                ldFecha = rs!fecha
                lsFecha = rs!lsFecha
            End If
        End If
        rs.Close

        If pnValor = 1 Then
           TempQ = "AND (IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0)) <= BS.nStockMinimo And bVigente = 1 "
        End If
        If pnValor = 2 Then
           TempQ = "AND (IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0)) >= 0 And bVigente = 1"
        Else
           TempQ = " AND (IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0)) <= BS.nStockMinimo " & _
                   " AND (IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0)) >= 0 And bVigente = 1 "
        End If

        If pbValidaInconsistencias Then
            TempQ = ""
        End If

        If pnAlmacenTpo <> 1 Then
           sSQL = " Select BS.cBSCod, BS.cBSDescripcion, CO.cConsDescripcion cConsUnidad, IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0) nAlmBSStock, (IsNull(INI.nMonto,0) + IsNull(ING.Valor,0) + IsNull(SAL.Valor,0)) TotalSaldo, (IsNull(INI.nMonto,0) + IsNull(ING.Valor,0) + IsNull(SAL.Valor,0)) / (case when IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0) = 0 then 1 else IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0) end) nAlmBSPrePromedio, (Select Top 1 Replace(cCtaContCod,'AG','" & Format(pnAlmacen, "00") & "') cCtaContCod from CtaBS where  cOpeCod = '591201' and BS.cBSCod Like cObjetoCod + '%' Order By Len(cObjetoCod) Desc) cCtaContCod ,isnull(BS.nStockMinimo,'')nStockMinimo  From bienesservicios BS WITH (NOLOCK)" _
             & " Left Join Constante CO WITH (NOLOCK) ON CO.nConsValor = BS.nBSUnidad And nConsCod = 1019" _
             & " Left Join (select cBSCod, nStock, nMonto from bssaldos WITH (NOLOCK) where dsaldo = '" & Format(ldFecha, gsFormatoFecha) & "' And nAlmCod = " & pnAlmacen & " And nAlmTpo = " & pnAlmacenTpo & ") As INI" _
             & "            On BS.cBSCod = INI.cBSCod" _
             & " Left Join (Select MB.cBSCod, Isnull(Sum(nMovCant) ,0) Cant, Isnull(Sum(nMovImporte) ,0) Valor From Mov M  WITH (NOLOCK) " _
             & "            Inner Join MovCant MC  WITH (NOLOCK) On M.nMovNro = MC.nMovNro" _
             & "            Inner Join MovBS MB  WITH (NOLOCK) On M.nMovNro = MB.nMovNro And MB.nMovItem = MC.nMovItem And nMovBSOrden = " & pnAlmacen & "" _
             & "            Inner Join MovCta MCTA  WITH (NOLOCK) On MB.nMovNro = MCTA.nMovNro And MB.nMovItem = MCTA.nMovItem" _
             & "            Where M.nMovEstado = '20' And M.nMovFlag = 0 And M.cOpeCod In (Select cOpeCod from LogOpeProd Where nOpeTpo = 1 And nProdTpo = " & pnAlmacenTpo & ")" _
             & "                  And M.cMovNro > '" & lsFecha & "' And M.cMovNro < '" & Format(DateAdd("d", 1, pdFecha), gsFormatoMovFecha) & "' Group by MB.cBSCod) As ING On BS.cBSCod = ING.cBSCod" _
             & " Left Join (Select MB.cBSCod , Isnull(Sum(nMovCant) ,0) Cant, Isnull(Sum(nMovImporte) ,0) Valor From Mov M WITH (NOLOCK) " _
             & "            Inner Join MovCant MC  WITH (NOLOCK) On M.nMovNro = MC.nMovNro" _
             & "            Inner Join MovBS MB  WITH (NOLOCK) On M.nMovNro = MB.nMovNro And MB.nMovItem = MC.nMovItem And nMovBSOrden = " & pnAlmacen & " " _
             & "            Inner Join MovCta MCTA  WITH (NOLOCK) On MB.nMovNro = MCTA.nMovNro And MB.nMovItem = MCTA.nMovItem"
             'EJVG20120403 Para que figuren las transferencias entre agencias
             '& "            Where M.nMovEstado in ('10','22') And M.nMovFlag = 0 And M.cOpeCod In (Select cOpeCod from LogOpeProd Where nOpeTpo = 2 And nProdTpo = " & pnAlmacenTpo & ")"
             sSQL = sSQL & " Where M.nMovEstado in ('10','22') And (M.nMovFlag = 0 Or (M.nMovFlag = 4 And (Select count(*) from Mov M1 Inner Join MovRef MR1 on M1.nMovNro = MR1.nMovNro Where MR1.nMovNroRef = M.nMovNro And M1.nMovEstado = 20 And M1.nMovFlag = 0) = 1)) And M.cOpeCod In (Select cOpeCod from LogOpeProd Where nOpeTpo = 2 And nProdTpo = " & pnAlmacenTpo & ")" _
             & "                And M.cMovNro > '" & lsFecha & "' And M.cMovNro < '" & Format(DateAdd("d", 1, pdFecha), gsFormatoMovFecha) & "' " _
             & "            Group By MB.cBSCod ) As SAL ON BS.cBSCod = SAL.cBSCod " _
             & " " & tmpSql & TempQ _
             & " Order by cCtaContCod "
         Else
             sSQL = " Select BS.cBSCod, BS.cBSDescripcion, CO.cConsDescripcion cConsUnidad, " _
             & " IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0) nAlmBSStock, (IsNull(INI.nMonto,0) + IsNull(ING.Valor,0) + IsNull(SAL.Valor,0)) TotalSaldo, " _
             & " (IsNull(INI.nMonto,0) + IsNull(ING.Valor,0) + IsNull(SAL.Valor,0)) / (case when IsNull(INI.nStock,0) + IsNull(ING.Cant,0) + IsNull(SAL.Cant,0) = 0 then 1 else IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0) end) nAlmBSPrePromedio, " _
             & " ING.cDocNro,ING.nDocTpo, (Select Top 1 Replace(cCtaContCod,'AG','" & Format(pnAlmacen, "00") & "') cCtaContCod from CtaBS where  cOpeCod = '591201' and BS.cBSCod Like cObjetoCod + '%' Order By Len(cObjetoCod) Desc) cCtaContCod ,isnull(BS.nStockMinimo,'')nStockMinimo " _
             & " From BienesServicios BS WITH (NOLOCK)" _
             & " Left Join Constante CO WITH (NOLOCK) ON CO.nConsValor = BS.nBSUnidad And nConsCod = 1019" _
             & " Left Join (select cBSCod, nStock, nMonto from bssaldos WITH (NOLOCK) where dsaldo = '" & Format(ldFecha, gsFormatoFecha) & "' And nAlmCod = " & pnAlmacen & " And nAlmTpo = " & pnAlmacenTpo & ") As INI" _
             & "            On BS.cBSCod = INI.cBSCod " _
             & " inner Join (Select MB.cBSCod, Isnull(Sum(nMovCant) ,0) Cant, Isnull(Sum(nMovImporte) ,0) Valor, " _
             & "            MVD.cDocNro, MVD.nDocTpo      From Mov M  WITH (NOLOCK) " _
             & "            Inner join MovDoc MVD on M.nMovNro = MVD.nMovNro " _
             & "            Inner Join MovCant MC  WITH (NOLOCK) On M.nMovNro = MC.nMovNro" _
             & "            Inner Join MovBS MB  WITH (NOLOCK) On M.nMovNro = MB.nMovNro And MB.nMovItem = MC.nMovItem And nMovBSOrden = " & pnAlmacen & "" _
             & "            Inner Join MovCta MCTA  WITH (NOLOCK) On MB.nMovNro = MCTA.nMovNro And MB.nMovItem = MCTA.nMovItem" _
             & "            Where M.nMovEstado = '20' And M.nMovFlag = 0 And M.cOpeCod In (Select cOpeCod from LogOpeProd Where nOpeTpo = 1 And nProdTpo = " & pnAlmacenTpo & ")" _
             & "                  And M.cMovNro > '" & lsFecha & "' And M.cMovNro < '" & Format(DateAdd("d", 1, pdFecha), gsFormatoMovFecha) & "' and MVD.nDocTpo = 42 " _
             & "            Group by MB.cBSCod, MVD.cDocNro, MVD.nDocTpo ) As ING On BS.cBSCod = ING.cBSCod" _
             & " Left Join (Select MB.cBSCod , Isnull(Sum(nMovCant) ,0) Cant, Isnull(Sum(nMovImporte) ,0) Valor From Mov M WITH (NOLOCK) " _
             & "            Inner Join MovCant MC  WITH (NOLOCK) On M.nMovNro = MC.nMovNro" _
             & "            Inner Join MovBS MB  WITH (NOLOCK) On M.nMovNro = MB.nMovNro And MB.nMovItem = MC.nMovItem And nMovBSOrden = " & pnAlmacen & " " _
             & "            Inner Join MovCta MCTA  WITH (NOLOCK) On MB.nMovNro = MCTA.nMovNro And MB.nMovItem = MCTA.nMovItem" _
             & "            Where M.nMovEstado in ('10','22') And M.nMovFlag = 0 And M.cOpeCod In (Select cOpeCod from LogOpeProd Where nOpeTpo = 2 And nProdTpo = " & pnAlmacenTpo & ")" _
             & "                And M.cMovNro > '" & lsFecha & "' And M.cMovNro < '" & Format(DateAdd("d", 1, pdFecha), gsFormatoMovFecha) & "' " _
             & "            Group By MB.cBSCod ) As SAL ON BS.cBSCod = SAL.cBSCod " _
             & " ORDER BY ING.cDocNro " & TempQ
          End If

'            sSQL = " Select * From (Select BS.cBSCod, BS.cBSDescripcion, IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, dbo.LogGetStockSubasta(" & pnAlmacen & ",cBsCod) nAlmBSStock, dbo.LogGetPrecioPromedioSubasta(1,cBsCod) nAlmBSPrePromedio From BienesServicios BS" _
                 & " LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = 1019" _
                 & " where Len(BS.cBSCod)  = (Select Max(len(cbscod)) From  BienesServicios) And BS.cBSCod In (Select distinct cBSCod from MovBs)) As AAA " & tmpSql & " Order by cBSCod"
 '           sSQL = " Select * From (Select BS.cBSCod, BS.cBSDescripcion, IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, dbo.LogGetStockEmbargo(" & pnAlmacen & ",cBsCod) nAlmBSStock, dbo.LogGetPrecioPromedioEmbargo(1,cBsCod) nAlmBSPrePromedio From BienesServicios BS" _
                 & " LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = 1019" _
                 & " where Len(BS.cBSCod)  = (Select Max(len(cbscod)) From  BienesServicios) And BS.cBSCod In (Select distinct cBSCod from MovBs)) As AAA " & tmpSql & " Order by cBSCod"
                 'And BS.cBSCod In (Select distinct cBSCod from MovBs)
        
        '[TORE ADD: Se paso la logica a procedimiento almacenado - RFC1902200002]
        sSQL = "exec  stp_sel_ObtenerInventarioAlmacen " & pnAlmacen & ", " & pnAlmacenTpo & ", '" & Format(pdFecha, "yyyyMMdd") & "', " & IIf(pbStockCero = False, 0, 1) & ", " & IIf(pbValidaInconsistencias = False, 0, 1) & ", " & pnValor
        
        
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    Set CargaAlmacenBS = rs
End Function

'Public Function CargaAlmacenBS(ByVal pnAlmacen As Integer, ByVal pnAlmacenTpo As Integer, _
'    ByVal pdFecha As Date, Optional ByVal pbStockCero As Boolean = True, Optional pbValidaInconsistencias As Boolean = False, Optional pnValor As Integer) As ADODB.Recordset
'    Dim oConec As DConecta
'    Dim rs As ADODB.Recordset
'    Dim tmpSql  As String
'    Dim TempQ As String
'
'    Set oConec = New DConecta
'    Set rs = New ADODB.Recordset
'
'    'CUANDO SE PIDE CON STOCK CERO
'    If pbValidaInconsistencias Then
'        tmpSql = " Where  ( (IsNull(Ini.nStock, 0) + IsNull(ING.Cant, 0) - IsNull(SAL.Cant, 0) <> 0 And IsNull(INI.nMonto,0) + IsNull(ING.Valor,0) + IsNull(SAL.Valor,0) = 0) Or (IsNull(Ini.nStock, 0) + IsNull(ING.Cant, 0) - IsNull(SAL.Cant, 0) = 0 And IsNull(INI.nMonto,0) + IsNull(ING.Valor,0) + IsNull(SAL.Valor,0) <> 0) )  And len(bs.cbscod) = (Select Max(len(cbscod)) From  BienesServicios)   "
'        Else
'        tmpSql = IIf(pbStockCero = False, " Where IsNull(Ini.nStock, 0) + IsNull(ING.Cant, 0) - IsNull(SAL.Cant, 0) <> 0 And len(bs.cbscod) = (Select Max(len(cbscod)) From  BienesServicios)   ", " Where len(bs.cbscod) = (Select Max(len(cbscod)) From  BienesServicios) ")
'    End If
'
'    Dim ldFecha As Date
'    Dim lsFecha As String
'
'    If oConec.AbreConexion(vsConexion) Then
'        sSql = " Select * From (Select BS.cBSCod, BS.cBSDescripcion, IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, dbo.LogGetStock(" & pnAlmacen & ",cBsCod," & pnAlmacenTpo & ") nAlmBSStock, dbo.LogGetPrecioPromedio(1,cBsCod," & pnAlmacenTpo & ") nAlmBSPrePromedio From BienesServicios BS" _
'             & " LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = 1019" _
'             & " where Len(BS.cBSCod)  = (Select Max(len(cbscod)) From  BienesServicios) ) As AAA " & tmpSql & " Order by cBSCod"
'
'        sSql = " Select max(dsaldo) Fecha,  convert(varchar(8),dateadd(day,1,  max(dsaldo)),112) lsFecha From BSSaldos  WITH (NOLOCK) Where dSaldo <= '" & Format(pdFecha, gsFormatoFecha) & "'"
'        Set rs = oConec.CargaRecordSet(sSql)
'
'
'
'        ldFecha = rs!Fecha
'        lsFecha = rs!lsFecha
'        rs.Close
'
'        If pnValor = 1 Then
'           TempQ = "AND (IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0)) < BS.nStockMinimo "
'        End If
'        If pnValor = 2 Then
'           TempQ = "AND (IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0))>0"
'        Else
'           TempQ = " AND (IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0)) < BS.nStockMinimo " & _
'                   " AND (IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0))>0 "
'        End If
'
'        If pnAlmacenTpo <> 1 Then
'           sSql = " Select BS.cBSCod, BS.cBSDescripcion, CO.cConsDescripcion cConsUnidad, IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0) nAlmBSStock, (IsNull(INI.nMonto,0) + IsNull(ING.Valor,0) + IsNull(SAL.Valor,0)) TotalSaldo, (IsNull(INI.nMonto,0) + IsNull(ING.Valor,0) + IsNull(SAL.Valor,0)) / (case when IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0) = 0 then 1 else IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0) end) nAlmBSPrePromedio, (Select Top 1 cCtaContCod from CtaBS where  cOpeCod = '591201' and BS.cBSCod Like cObjetoCod + '%' Order By Len(cObjetoCod) Desc) cCtaContCod ,BS.nStockMinimo  From bienesservicios BS WITH (NOLOCK)" _
'             & " Left Join Constante CO WITH (NOLOCK) ON CO.nConsValor = BS.nBSUnidad And nConsCod = 1019" _
'             & " Left Join (select cBSCod, nStock, nMonto from bssaldos WITH (NOLOCK) where dsaldo = '" & Format(ldFecha, gsFormatoFecha) & "' And nAlmCod = " & pnAlmacen & " And nAlmTpo = " & pnAlmacenTpo & ") As INI" _
'             & "            On BS.cBSCod = INI.cBSCod" _
'             & " Left Join (Select MB.cBSCod, Isnull(Sum(nMovCant) ,0) Cant, Isnull(Sum(nMovImporte) ,0) Valor From Mov M  WITH (NOLOCK) " _
'             & "            Inner Join MovCant MC  WITH (NOLOCK) On M.nMovNro = MC.nMovNro" _
'             & "            Inner Join MovBS MB  WITH (NOLOCK) On M.nMovNro = MB.nMovNro And MB.nMovItem = MC.nMovItem And nMovBSOrden = " & pnAlmacen & "" _
'             & "            Inner Join MovCta MCTA  WITH (NOLOCK) On MB.nMovNro = MCTA.nMovNro And MB.nMovItem = MCTA.nMovItem" _
'             & "            Where M.nMovEstado = '20' And M.nMovFlag = 0 And M.cOpeCod In (Select cOpeCod from LogOpeProd Where nOpeTpo = 1 And nProdTpo = " & pnAlmacenTpo & ")" _
'             & "                  And M.cMovNro > '" & lsFecha & "' And M.cMovNro < '" & Format(DateAdd("d", 1, pdFecha), gsFormatoMovFecha) & "' Group by MB.cBSCod) As ING On BS.cBSCod = ING.cBSCod" _
'             & " Left Join (Select MB.cBSCod , Isnull(Sum(nMovCant) ,0) Cant, Isnull(Sum(nMovImporte) ,0) Valor From Mov M WITH (NOLOCK) " _
'             & "            Inner Join MovCant MC  WITH (NOLOCK) On M.nMovNro = MC.nMovNro" _
'             & "            Inner Join MovBS MB  WITH (NOLOCK) On M.nMovNro = MB.nMovNro And MB.nMovItem = MC.nMovItem And nMovBSOrden = " & pnAlmacen & " " _
'             & "            Inner Join MovCta MCTA  WITH (NOLOCK) On MB.nMovNro = MCTA.nMovNro And MB.nMovItem = MCTA.nMovItem" _
'             & "            Where M.nMovEstado in ('10','22') And M.nMovFlag = 0 And M.cOpeCod In (Select cOpeCod from LogOpeProd Where nOpeTpo = 2 And nProdTpo = " & pnAlmacenTpo & ")" _
'             & "                And M.cMovNro > '" & lsFecha & "' And M.cMovNro < '" & Format(DateAdd("d", 1, pdFecha), gsFormatoMovFecha) & "' " _
'             & "            Group By MB.cBSCod ) As SAL ON BS.cBSCod = SAL.cBSCod " _
'             & " " & tmpSql & TempQ
'         Else
'             sSql = " Select BS.cBSCod, BS.cBSDescripcion, CO.cConsDescripcion cConsUnidad, " _
'             & " IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0) nAlmBSStock, (IsNull(INI.nMonto,0) + IsNull(ING.Valor,0) + IsNull(SAL.Valor,0)) TotalSaldo, " _
'             & " (IsNull(INI.nMonto,0) + IsNull(ING.Valor,0) + IsNull(SAL.Valor,0)) / (case when IsNull(INI.nStock,0) + IsNull(ING.Cant,0) + IsNull(SAL.Cant,0) = 0 then 1 else IsNull(INI.nStock,0) + IsNull(ING.Cant,0) - IsNull(SAL.Cant,0) end) nAlmBSPrePromedio, " _
'             & " ING.cDocNro,ING.nDocTpo, (Select Top 1 cCtaContCod from CtaBS where  cOpeCod = '591201' and BS.cBSCod Like cObjetoCod + '%' Order By Len(cObjetoCod) Desc) cCtaContCod ,BS.nStockMinimo " _
'             & " From BienesServicios BS WITH (NOLOCK)" _
'             & " Left Join Constante CO WITH (NOLOCK) ON CO.nConsValor = BS.nBSUnidad And nConsCod = 1019" _
'             & " Left Join (select cBSCod, nStock, nMonto from bssaldos WITH (NOLOCK) where dsaldo = '" & Format(ldFecha, gsFormatoFecha) & "' And nAlmCod = " & pnAlmacen & " And nAlmTpo = " & pnAlmacenTpo & ") As INI" _
'             & "            On BS.cBSCod = INI.cBSCod " _
'             & " inner Join (Select MB.cBSCod, Isnull(Sum(nMovCant) ,0) Cant, Isnull(Sum(nMovImporte) ,0) Valor, " _
'             & "            MVD.cDocNro, MVD.nDocTpo      From Mov M  WITH (NOLOCK) " _
'             & "            Inner join MovDoc MVD on M.nMovNro = MVD.nMovNro " _
'             & "            Inner Join MovCant MC  WITH (NOLOCK) On M.nMovNro = MC.nMovNro" _
'             & "            Inner Join MovBS MB  WITH (NOLOCK) On M.nMovNro = MB.nMovNro And MB.nMovItem = MC.nMovItem And nMovBSOrden = " & pnAlmacen & "" _
'             & "            Inner Join MovCta MCTA  WITH (NOLOCK) On MB.nMovNro = MCTA.nMovNro And MB.nMovItem = MCTA.nMovItem" _
'             & "            Where M.nMovEstado = '20' And M.nMovFlag = 0 And M.cOpeCod In (Select cOpeCod from LogOpeProd Where nOpeTpo = 1 And nProdTpo = " & pnAlmacenTpo & ")" _
'             & "                  And M.cMovNro > '" & lsFecha & "' And M.cMovNro < '" & Format(DateAdd("d", 1, pdFecha), gsFormatoMovFecha) & "' and MVD.nDocTpo = 42 " _
'             & "            Group by MB.cBSCod, MVD.cDocNro, MVD.nDocTpo ) As ING On BS.cBSCod = ING.cBSCod" _
'             & " Left Join (Select MB.cBSCod , Isnull(Sum(nMovCant) ,0) Cant, Isnull(Sum(nMovImporte) ,0) Valor From Mov M WITH (NOLOCK) " _
'             & "            Inner Join MovCant MC  WITH (NOLOCK) On M.nMovNro = MC.nMovNro" _
'             & "            Inner Join MovBS MB  WITH (NOLOCK) On M.nMovNro = MB.nMovNro And MB.nMovItem = MC.nMovItem And nMovBSOrden = " & pnAlmacen & " " _
'             & "            Inner Join MovCta MCTA  WITH (NOLOCK) On MB.nMovNro = MCTA.nMovNro And MB.nMovItem = MCTA.nMovItem" _
'             & "            Where M.nMovEstado in ('10','22') And M.nMovFlag = 0 And M.cOpeCod In (Select cOpeCod from LogOpeProd Where nOpeTpo = 2 And nProdTpo = " & pnAlmacenTpo & ")" _
'             & "                And M.cMovNro > '" & lsFecha & "' And M.cMovNro < '" & Format(DateAdd("d", 1, pdFecha), gsFormatoMovFecha) & "' " _
'             & "            Group By MB.cBSCod ) As SAL ON BS.cBSCod = SAL.cBSCod " _
'             & " ORDER BY ING.cDocNro " & TempQ
'          End If
'
''            sSQL = " Select * From (Select BS.cBSCod, BS.cBSDescripcion, IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, dbo.LogGetStockSubasta(" & pnAlmacen & ",cBsCod) nAlmBSStock, dbo.LogGetPrecioPromedioSubasta(1,cBsCod) nAlmBSPrePromedio From BienesServicios BS" _
'                 & " LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = 1019" _
'                 & " where Len(BS.cBSCod)  = (Select Max(len(cbscod)) From  BienesServicios) And BS.cBSCod In (Select distinct cBSCod from MovBs)) As AAA " & tmpSql & " Order by cBSCod"
' '           sSQL = " Select * From (Select BS.cBSCod, BS.cBSDescripcion, IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, dbo.LogGetStockEmbargo(" & pnAlmacen & ",cBsCod) nAlmBSStock, dbo.LogGetPrecioPromedioEmbargo(1,cBsCod) nAlmBSPrePromedio From BienesServicios BS" _
'                 & " LEFT JOIN Constante CU ON bs.nBSUnidad = cu.nConsValor And cu.nConsCod = 1019" _
'                 & " where Len(BS.cBSCod)  = (Select Max(len(cbscod)) From  BienesServicios) And BS.cBSCod In (Select distinct cBSCod from MovBs)) As AAA " & tmpSql & " Order by cBSCod"
'                 'And BS.cBSCod In (Select distinct cBSCod from MovBs)
'        Set rs = oConec.CargaRecordSet(sSql)
'        oConec.CierraConexion
'    End If
'    Set CargaAlmacenBS = rs
'End Function

'Verifica existencia de B/S en Almacen
Public Function ExisAlmacenBS(ByVal pnAlmacen As Integer, _
ByVal psBSCod As String) As Boolean
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSQL = " SELECT cBSCod " & _
            " FROM AlmacenBS " & _
            " WHERE nAlmCod = " & pnAlmacen & " AND cBSCod = '" & psBSCod & "'"
        Set rs = oConec.CargaRecordSet(sSQL)
        If rs.RecordCount > 0 Then
            ExisAlmacenBS = True
        Else
            ExisAlmacenBS = False
        End If
        oConec.CierraConexion
    End If
End Function

'Reporte existencia de B/S en Almacen
'Public Function ReporteAlmacenBS(ByVal lsNomAge As String, ByVal lsEmpresa As String, ByVal ldFecSis As Date, ByVal pnAlmacen As Integer, pnProdTpo As Long, pdFechaProceso As Date, Optional ByVal pbStockCero As Boolean = True, Optional ByVal pbValidaInconsistencias As Boolean = False) As String
'    Dim rs As ADODB.Recordset
'    Dim lsCadena As String
'    Dim lnItem As Long
'    Dim lnContador As Long
'    Dim lnPagina As Long
'    Dim lsCorr As String * 5
'    Dim lsBSCod As String * 15
'    Dim lsBSDescripcion As String * 40
'    Dim lsConsUnidad As String * 18
'    Dim lsAlmBSStock As String * 10
'    Dim lsAlmPrecio As String * 11
'    Dim lsAlmValor As String * 11
'
'    Dim lnAlmBSStock As Currency
'    Dim lnAlmPrecio As Currency
'    Dim lnAlmValor As Currency
'    Dim xProdTpo As String
'    Dim xLote As String
'
'    lnAlmBSStock = 0
'    lnAlmPrecio = 0
'    lnAlmValor = 0
'
'    Dim lsItem As String * 5
'
'    If pnProdTpo = 0 Then
'       xProdTpo = "BIENES"
'    ElseIf pnProdTpo = 1 Then
'       xProdTpo = "BIENES ADJUDICADOS"
'    ElseIf pnProdTpo = 2 Then
'       xProdTpo = "BIENES EMBARGADOS"
'    End If
'
'    Set rs = New ADODB.Recordset
'    Set rs = CargaAlmacenBS(pnAlmacen, pnProdTpo, pdFechaProceso, pbStockCero, pbValidaInconsistencias)
'
'    If Not (rs.EOF And rs.BOF) Then
'        lnContador = 1
'        lsCadena = ""
'        lsCadena = lsCadena & CabeceraPagina("Inventario " & IIf(pbValidaInconsistencias, "INCONSISTENCIAS ", "") & "de " & xProdTpo & " del Almacen (" & Format(pdFechaProceso, gsFormatoFechaView) & ") : ", lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, "")
'        lsCadena = lsCadena & Encabezado("Item;5;Codigo;8; ;7;Descripcion;20; ;20;Unidad;12; ;6;Stock;10; ;8;P.Prom.;8; ;2;P.Valor;8; ;4;Cta Cont;9;", lnItem)
'        xLote = ""
'        While Not rs.EOF
'            lsCorr = Format(lnContador, " 0000")
'            lsBSCod = rs!cBsCod
'            lsBSDescripcion = rs!cBSDescripcion
'            lsConsUnidad = rs!cConsUnidad & ""
'
'            RSet lsAlmBSStock = Format(rs!nAlmBSStock, "#,##0.00")
'            RSet lsAlmPrecio = Format(rs!nAlmBSPrePromedio, "#,##0.00")
'            RSet lsAlmValor = Format(rs!TotalSaldo, "#,##0.00")
'
'            lnAlmBSStock = lnAlmBSStock + rs!nAlmBSStock
'            lnAlmPrecio = lnAlmPrecio + rs!nAlmBSPrePromedio
'            lnAlmValor = lnAlmValor + (rs!TotalSaldo)
'
'            If pnProdTpo = 1 Then
'               If xLote <> rs!cDocNro Then
'                  If lnContador > 1 Then
'                     lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'                  End If
'                  lsCadena = lsCadena & " LOTE: " & rs!cDocNro & oImpresora.gPrnSaltoLinea
'                  lsCadena = lsCadena & " ----- " & String(Len(rs!cDocNro), "-") & oImpresora.gPrnSaltoLinea
'                  xLote = rs!cDocNro
'               End If
'            End If
'            lsCadena = lsCadena & lsCorr & "  " & lsBSCod & "  " & lsBSDescripcion & "  " & lsConsUnidad & lsAlmBSStock & lsAlmPrecio & " " & lsAlmValor & "  " & rs!cCtaContCod & oImpresora.gPrnSaltoLinea
'
'
'            lnContador = lnContador + 1
'            lnItem = lnItem + 1
'
'            If lnItem = 54 Then
'                lnItem = 0
'                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
'                lsCadena = lsCadena & CabeceraPagina("Inventario " & IIf(pbValidaInconsistencias, "INCONSISTENCIAS ", "") & "de Bienes del Almacen (" & Format(pdFechaProceso, gsFormatoFechaView) & ") : ", lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, "")
'                lsCadena = lsCadena & Encabezado("Item;5;Codigo;8; ;7;Descripcion;20; ;20;Unidad;12; ;6;Stock;10; ;8;P.Prom.;8; ;2;P.Valor;8; ;4;Cta Cont;9;", lnItem)
'            End If
'
'            rs.MoveNext
'        Wend
'    End If
'
'    lsCorr = ""
'    lsBSCod = ""
'    lsBSDescripcion = ""
'    lsConsUnidad = ""
'    RSet lsAlmBSStock = Format(lnAlmBSStock, "#,##0.00")
'    RSet lsAlmPrecio = Format(lnAlmPrecio, "#,##0.00")
'    RSet lsAlmValor = Format(lnAlmValor, "#,##0.00")
'
'    lsCadena = lsCadena & String(128, "=") & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsCorr & "  " & lsBSCod & "  " & lsBSDescripcion & "  " & lsConsUnidad & lsAlmBSStock & lsAlmPrecio & " " & lsAlmValor & oImpresora.gPrnSaltoLinea
'
'    ReporteAlmacenBS = lsCadena
'End Function

'********Nuevo Reporte AlmacenBS ******
'Reporte existencia de B/S en Almacen
Public Function ReporteAlmacenBS(ByVal lsNomAge As String, ByVal lsEmpresa As String, ByVal ldFecSis As Date, ByVal pnAlmacen As Integer, pnProdTpo As Long, pdFechaProceso As Date, Optional ByVal pbStockCero As Boolean = True, Optional ByVal pbValidaInconsistencias As Boolean = False, Optional ByVal pnValor As Integer) As String
    Dim rs As ADODB.Recordset
    Dim lsCadena As String
    Dim lnItem As Long
    Dim lnContador As Long
    Dim lnPagina As Long
    Dim lsCorr As String * 5
    Dim lsBSCod As String * 15
    Dim lsBSDescripcion As String * 40
    Dim lsConsUnidad As String * 18
    Dim lsAlmBSStock As String * 10
    Dim lsAlmPrecio As String * 11
    Dim lsAlmValor As String * 11
    
    Dim lnAlmBSStock As Currency
    Dim lnAlmPrecio As Currency
    Dim lnAlmValor As Currency
    Dim xProdTpo As String
    Dim xLote As String
    
    lnAlmBSStock = 0
    lnAlmPrecio = 0
    lnAlmValor = 0

    Dim lsItem As String * 5
    
    If pnProdTpo = 0 Then
       xProdTpo = "BIENES"
    ElseIf pnProdTpo = 1 Then
       xProdTpo = "BIENES ADJUDICADOS"
    ElseIf pnProdTpo = 2 Then
       xProdTpo = "BIENES EMBARGADOS"
    End If
    
    Set rs = New ADODB.Recordset
    Set rs = CargaAlmacenBS(pnAlmacen, pnProdTpo, pdFechaProceso, pbStockCero, pbValidaInconsistencias, pnValor)
    
    If Not (rs.EOF And rs.BOF) Then
        lnContador = 1
        lsCadena = ""
        lsCadena = lsCadena & oImpresora.gPrnCondensadaON
        lsCadena = lsCadena & CabeceraPagina("Inventario " & IIf(pbValidaInconsistencias, "INCONSISTENCIAS ", "") & "de " & xProdTpo & " del Almacen (" & Format(pdFechaProceso, gsFormatoFechaView) & ") : ", lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, "")
        lsCadena = lsCadena & Encabezado("Item;5;Codigo;8; ;7;Descripcion;20; ;20;Unidad;12; ;6;Stock;10; ;8;P.Prom.;8; ;2;P.Valor;8; ;4;Cta Cont;9;", lnItem)
        xLote = ""
        While Not rs.EOF
            lsCorr = Format(lnContador, " 0000")
            lsBSCod = rs!cBSCod
            lsBSDescripcion = rs!cBSDescripcion
            lsConsUnidad = rs!cConsUnidad & ""
            
            RSet lsAlmBSStock = Format(rs!nAlmBSStock, "#,##0.00")
            RSet lsAlmPrecio = Format(rs!nAlmBSPrePromedio, "#,##0.00")
            RSet lsAlmValor = Format(rs!TotalSaldo, "#,##0.00")
            
            lnAlmBSStock = lnAlmBSStock + rs!nAlmBSStock
            lnAlmPrecio = lnAlmPrecio + rs!nAlmBSPrePromedio
            lnAlmValor = lnAlmValor + (rs!TotalSaldo)
            
            If pnProdTpo = 1 Then
               If xLote <> rs!cDocNro Then
                  If lnContador > 1 Then
                     lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
                  End If
                  lsCadena = lsCadena & " LOTE: " & rs!cDocNro & oImpresora.gPrnSaltoLinea
                  lsCadena = lsCadena & " ----- " & String(Len(rs!cDocNro), "-") & oImpresora.gPrnSaltoLinea
                  xLote = rs!cDocNro
               End If
            End If
            lsCadena = lsCadena & lsCorr & "  " & lsBSCod & "  " & lsBSDescripcion & "  " & lsConsUnidad & lsAlmBSStock & lsAlmPrecio & " " & lsAlmValor & "  " & rs!cCtaContCod & oImpresora.gPrnSaltoLinea
            
            
            lnContador = lnContador + 1
            lnItem = lnItem + 1
            
            If lnItem = 54 Then
                lnItem = 0
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Inventario " & IIf(pbValidaInconsistencias, "INCONSISTENCIAS ", "") & "de Bienes del Almacen (" & Format(pdFechaProceso, gsFormatoFechaView) & ") : ", lnPagina, lnItem, lsNomAge, lsEmpresa, ldFecSis, "")
                lsCadena = lsCadena & Encabezado("Item;5;Codigo;8; ;7;Descripcion;20; ;20;Unidad;12; ;6;Stock;10; ;8;P.Prom.;8; ;2;P.Valor;8; ;4;Cta Cont;9;", lnItem)
            End If
            
            rs.MoveNext
        Wend
    End If
    
    lsCorr = ""
    lsBSCod = ""
    lsBSDescripcion = ""
    lsConsUnidad = ""
    RSet lsAlmBSStock = Format(lnAlmBSStock, "#,##0.00")
    RSet lsAlmPrecio = Format(lnAlmPrecio, "#,##0.00")
    RSet lsAlmValor = Format(lnAlmValor, "#,##0.00")
    
    lsCadena = lsCadena & String(128, "=") & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsCorr & "  " & lsBSCod & "  " & lsBSDescripcion & "  " & lsConsUnidad & lsAlmBSStock & lsAlmPrecio & " " & lsAlmValor & oImpresora.gPrnSaltoLinea
    
    ReporteAlmacenBS = lsCadena
End Function



Public Function GetBienesProv(Optional psPersProveedor As String = "") As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    If psPersProveedor <> "" Then
        sqlA = " Select RTrim(BS.cBSCod) Codigo, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])', len(BS.cBSCod) Nivel " _
             & " From BienesServicios BS" _
             & " Inner Join ProveedorBS PBS On BS.cBSCod like PBS.cBSCod + '%'" _
             & " Inner Join Constante Co On CO.nConsValor = BS.nBSUnidad  And CO.nConsCOd = '1019'" _
             & " Where cPersCod = '" & psPersProveedor & "'"
    Else
        sqlA = " Select RTrim(BS.cBSCod) Codigo, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])', len(BS.cBSCod) Nivel " _
             & " From BienesServicios BS" _
             & " Inner Join Constante Co On CO.nConsValor = BS.nBSUnidad  And CO.nConsCOd = '1019'"
    End If
    Set GetBienesProv = oCon.CargaRecordSet(sqlA)
    
    Set oCon = Nothing
End Function

Public Function VerificarCierrecBSCod(pdFecha As Date, pnAlmacen As Integer, pnAlmTpo As Integer) As ADODB.Recordset
    'ANPS
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
   
    sqlA = "Exec sp_VerificarCierrecBSCod '" & Format(pdFecha, "yyyymmdd") & "', " & CInt(pnAlmacen) & ", " & CInt(pnAlmTpo)

    Set VerificarCierrecBSCod = oCon.CargaRecordSet(sqlA)
    Set oCon = Nothing
        If VerificarCierrecBSCod!cBSCod = 0 Then
                       MsgBox ("Aún no se ha realizado el cierre de almacén del mes anterior" & vbNewLine & "No se podrá procesar ")
            Exit Function
        End If
   
    'FIN ANPS

End Function


Public Function GetLogAlmacenStock(psAlmacen As String, pdFecha As Date, pnProdTpo As Integer, Optional lsCad As String, Optional Cod As Integer) As ADODB.Recordset
'COMENTADO ANPS
'    Dim sqlA  As String
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'
'    oCon.AbreConexion
'    'anpscBSCod
''    sqlA = " Select lischeck=cBSCod,cBSCod, cBSDescripcion + ' ('  + cConsDescripcion + ')' Descripcion, 0, 0 From BienesServicios BS " _
''         & "    Left Join Constante CO On CO.nConsCod = 1019 And CO.nConsValor = BS.nBSUnidad" _
''         & " Where Len(cBSCod) >= 10 And cBSCod Not In (Select cBSCod From BSSaldos BSS Where  dSaldo between '" & Format(pdFecha, "yyyymmdd") & " 00:00:00' And '" & Format(pdFecha, "yyyymmdd") & " 23:59:59' And nAlmCod = " & psAlmacen & " And nAlmTpo = " & pnProdTpo & ")  and BS.bVigente=1" _
''         & "    Union" _
''         & " Select lischeck='0',BS.cBSCod, cBSDescripcion + ' ('  + cConsDescripcion + ')' Descripcion, nStock, nMonto from BSSaldos BSS" _
''         & "    Inner Join BienesServicios BS On BSS.cBSCod = BS.cBSCod" _
''         & "    Left Join Constante CO On CO.nConsCod = 1019 And CO.nConsValor = BS.nBSUnidad" _
''         & " Where  BS.bVigente=1 and dSaldo between '" & Format(pdFecha, "yyyymmdd") & " 00:00:00' And '" & Format(pdFecha, "yyyymmdd") & " 23:59:59' And nAlmCod = " & psAlmacen & "  And nAlmTpo = " & pnProdTpo
'
'    Set GetLogAlmacenStock = oCon.CargaRecordSet(sqlA)
'
'    Set oCon = Nothing
    
    'ANPS
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
   
    sqlA = "Exec stp_sel_GetLogAlmacenStock " & CInt(psAlmacen) & ",'" & Format(pdFecha, "yyyymmdd") & "', " & CInt(pnProdTpo) & " ," & " '" & lsCad & "' ," & Cod & " "

    Set GetLogAlmacenStock = oCon.CargaRecordSet(sqlA)
    Set oCon = Nothing
 
        If GetLogAlmacenStock!lischeck <> lsCad And Cod = 1 Then
            MsgBox "No se ha encontrado dato !", vbCritical, "Aviso"
            Exit Function
        End If
   
    'FIN ANPS

End Function

Public Function GetBienes() As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    sqlA = " Select cBSCod, cBSDescripcion + ' ('  + cConsDescripcion + ')' Descripcion From BienesServicios BS " _
         & "    Left Join Constante CO On CO.nConsCod = 1019 And CO.nConsValor = BS.nBSUnidad" _
         & " Where Len(cBSCod) >= 10 "
    Set GetBienes = oCon.CargaRecordSet(sqlA)
    
    Set oCon = Nothing
End Function

'Public Function GetIngresosAlmacen(pdIni As Date, pdFin As Date, Optional pnAlmcen As Integer = 1, Optional pbOrdenCtaCont As Boolean = False) As ADODB.Recordset
'    Dim sqlA  As String
'    Dim oCon As DConecta
'    Dim sqlOrder As String
'
'    Set oCon = New DConecta
'
'
'    If Not pbOrdenCtaCont Then
'        sqlOrder = " Order by MD.cDocNro"
'    Else
'        sqlOrder = " Order by MCTA.cCtaContCod, MD.cDocNro"
'    End If
'
'    oCon.AbreConexion
'
'    sqlA = " Select M.cMovNro, M.nMovNro,M.cMovDesc,MD.cDocNro,MD.dDocFecha,MBS.cSerie nMovCant,BS.cBSCod,BS.cBSDescripcion,CO.cConsDescripcion,PE.cPersNombre, Abs(MBS.nValor) nMovImporte, BS.cBSCod cCtaContCod ," _
'         & " IsNull((Select Top 1 MFAC.cDocNro From MovDoc MFAC Where MFAC.nMovNro = M.nMOvNro And nDocTpo = 1),'') FACTURA," _
'         & " IsNull((Select Top 1 MGUI.cDocNro From MovDoc MGUI Where MGUI.nMovNro = M.nMOvNro And nDocTpo = 20),'') GUIA" _
'         & " From Mov M" _
'         & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
'         & " Inner Join MovCant MC On MC.nMovNro = M.nMovNro" _
'         & " Inner Join MovBSSerie MBS On MC.nMovNro = MBS.nMovNro And MC.nMovItem = MBS.nMovItem  " _
'         & " Inner Join BienesServicios BS On MBS.cBSCod = BS.cBSCod" _
'         & " Left  Join MovGasto MG On MG.nMovNro = M.nMovNro" _
'         & " Left  Join Persona PE On MG.cPersCod = PE.cPersCod" _
'         & " Left  Join Movcta MCTA On MCTA.nMovNro = MC.nMovNro And MC.nMovItem = MCTA.nMovItem  And nMovImporte >= 0" _
'         & " Left  Join Constante CO On CO.nConsCod = '1019' And CO.nConsValor = BS.nBSUnidad" _
'         & " where nMovEstado = '20' and nDocTpo = '42' And nMovFlag = '" & MovFlag.gMovFlagVigente & "' And " _
'         & " MD.dDocFecha Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & " 23:59:59'  And MBS.cBSCod Like '112%'" _
'         & sqlOrder
'
'
'    sqlA = " Select M.cMovNro, M.nMovNro,M.cMovDesc,MD.cDocNro,MD.dDocFecha,MC.nMovCant,BS.cBSCod,BS.cBSDescripcion,CO.cConsDescripcion,PE.cPersNombre, Abs(MCTA.nMovImporte) nMovImporte, MCTA.cCtaContCod ," _
'         & "  IsNull((Select Top 1 MFAC.cDocNro From MovDoc MFAC Where MFAC.nMovNro = M.nMOvNro And nDocTpo = 1),'')  FACTURA," _
'         & "  IsNull((Select Top 1 MGUI.cDocNro From MovDoc MGUI Where MGUI.nMovNro = M.nMOvNro And nDocTpo = 20),'') GUIA" _
'         & " From Mov M" _
'         & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
'         & " Inner Join MovCant MC On MC.nMovNro = M.nMovNro And MC.nMovCant > 0 " _
'         & " Inner Join MovBS MBS On MC.nMovNro = MBS.nMovNro And MC.nMovItem = MBS.nMovItem And nMovBSOrden = " & pnAlmcen & "   " _
'         & " Inner Join BienesServicios BS On MBS.cBSCod = BS.cBSCod" _
'         & " Left  Join MovGasto MG On MG.nMovNro = M.nMovNro" _
'         & " Left  Join Persona PE On MG.cPersCod = PE.cPersCod" _
'         & " Left  Join Movcta MCTA On MCTA.nMovNro = MC.nMovNro And MC.nMovItem = MCTA.nMovItem  And nMovImporte >  0" _
'         & " Left  Join Constante CO On CO.nConsCod = '1019' And CO.nConsValor = BS.nBSUnidad" _
'         & " where nMovEstado = '20' and nDocTpo = '42' And nMovFlag = '" & MovFlag.gMovFlagVigente & "' And " _
'         & " MD.dDocFecha Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & " 23:59:59'  " _
'         & sqlOrder
'
'    Set GetIngresosAlmacen = oCon.CargaRecordSet(sqlA)
'
'    Set oCon = Nothing
'End Function

Public Function GetIngresosAlmacen(pdIni As Date, pdFin As Date, Optional psPersCod As String, Optional pnAlmcen As Integer = 1, Optional pbOrdenCtaCont As Boolean = False) As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Dim sqlOrder As String
    Dim pnOrden As Integer 'NAGL 20180423
    Set oCon = New DConecta
    
    
    If Not pbOrdenCtaCont Then
        sqlOrder = " Order by MD.cDocNro"
        pnOrden = 1 'NAGL 20180423
    Else
        sqlOrder = " Order by MCTA.cCtaContCod, MD.cDocNro"
        pnOrden = 2 'NAGL 20180423
    End If
    
    oCon.AbreConexion
    'sqlA = " Select M.cMovNro, M.nMovNro,M.cMovDesc,MD.cDocNro,MD.dDocFecha,MC.nMovCant,BS.cBSCod,BS.cBSDescripcion,CO.cConsDescripcion,PE.cPersNombre, Abs(MCTA.nMovImporte) nMovImporte, MCTA.cCtaContCod ," _
         '& "  IsNull((Select Top 1 MFAC.cDocNro From MovDoc MFAC Where MFAC.nMovNro = M.nMOvNro And nDocTpo = 1),'')  FACTURA," _
         '& "  IsNull((Select Top 1 MGUI.cDocNro From MovDoc MGUI Where MGUI.nMovNro = M.nMOvNro And nDocTpo = 20),'') GUIA" _
         '& " From Mov M" _
         '& " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
         '& " Inner Join MovCant MC On MC.nMovNro = M.nMovNro" _
         '& " Inner Join MovBS MBS On MC.nMovNro = MBS.nMovNro And MC.nMovItem = MBS.nMovItem And nMovBSOrden = " & pnAlmcen & "   " _
         '& " Inner Join BienesServicios BS On MBS.cBSCod = BS.cBSCod" _
         '& " Left  Join MovGasto MG On MG.nMovNro = M.nMovNro" _
         '& " Left  Join Persona PE On MG.cPersCod = PE.cPersCod" _
         '& " Left  Join Movcta MCTA On MCTA.nMovNro = MC.nMovNro And MC.nMovItem = MCTA.nMovItem  And nMovImporte >= 0" _
         '& " Left  Join Constante CO On CO.nConsCod = '1019' And CO.nConsValor = BS.nBSUnidad" _
         '& " where nMovEstado = '20' and nDocTpo = '42' And nMovFlag = '" & MovFlag.gMovFlagVigente & "' And " _
         '& " MD.dDocFecha Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & " 23:59:59'" _
         '& " " & IIf(psPersCod = "", "", " and PE.cPersCod='" & psPersCod & "'") _
         '& sqlOrder
         'Comentado by NAGL 20180423
    sqlA = "Exec stp_sel_GetIngresosAlmacen '" & Format(pdIni, "yyyymmdd") & "','" & Format(pdFin, "yyyymmdd") & "', '" & psPersCod & "' ," & pnAlmcen & ", " & pnOrden & " "
    'Procedimiento Agregado by NAGL 20180423
    Set GetIngresosAlmacen = oCon.CargaRecordSet(sqlA)
    Set oCon = Nothing
End Function

Public Function GetIngresosAlmacenAF(pdIni As Date, pdFin As Date, Optional pnAlmcen As Integer = 1) As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    sqlA = " Select M.cMovNro, MCTA.nMovItem, PE.cPersCod, BS.bSerie, M.nMovNro,M.cMovDesc,MD.cDocNro,MD.dDocFecha,MC.nMovCant,BS.cBSCod,BS.cBSDescripcion,CO.cConsDescripcion,PE.cPersNombre, Abs(MCTA.nMovImporte) nMovImporte, MCTA.cCtaContCod ," _
         & "  IsNull((Select Top 1 MFAC.cDocNro From MovDoc MFAC Where MFAC.nMovNro = M.nMOvNro And nDocTpo = 1),'')  FACTURA," _
         & "  IsNull((Select Top 1 MGUI.cDocNro From MovDoc MGUI Where MGUI.nMovNro = M.nMOvNro And nDocTpo = 20),'') GUIA" _
         & " From Mov M" _
         & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
         & " Inner Join MovCant MC On MC.nMovNro = M.nMovNro" _
         & " Inner Join MovBS MBS On MC.nMovNro = MBS.nMovNro And MC.nMovItem = MBS.nMovItem And nMovBSOrden = " & pnAlmcen & "   " _
         & " Inner Join BienesServicios BS On MBS.cBSCod = BS.cBSCod" _
         & " Left  Join MovGasto MG On MG.nMovNro = M.nMovNro" _
         & " Left  Join Persona PE On MG.cPersCod = PE.cPersCod" _
         & " Left  Join Movcta MCTA On MCTA.nMovNro = MC.nMovNro And MC.nMovItem = MCTA.nMovItem  And nMovImporte >= 0" _
         & " Left  Join Constante CO On CO.nConsCod = '1019' And CO.nConsValor = BS.nBSUnidad" _
         & " where nMovEstado = '20' and nDocTpo = '42' And nMovFlag = '0' And " _
         & " MD.dDocFecha Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & " 23:59:59' And Left(BS.cBsCod,3) In ('112','113') " _
         & " Order by MD.cDocNro"
    Set GetIngresosAlmacenAF = oCon.CargaRecordSet(sqlA)
    
    Set oCon = Nothing
End Function

Public Function GetSalidasAlmacen(pdIni As Date, pdFin As Date, Optional pnAlmcen As Integer = 1, Optional pbOrdenCtaCont As Boolean = False) As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlOrden As String
    Dim lnOrden As Integer '*** PEAC 20100903
    
    If Not pbOrdenCtaCont Then
        sqlOrden = " Order by MD.cDocNro"
        lnOrden = 1 '*** PEAC 20100903
    Else
        sqlOrden = " Order by MCTA.cCtaContCod, MD.cDocNro"
        lnOrden = 2 '*** PEAC 20100903
    End If

    oCon.AbreConexion
    
    '*** PEAC 20100903
    'sqlA = " Select Distinct M.nMovNro, M.cMovNro, M.cMovDesc, MD.cDocNro, MD.dDocFecha,MC.nMovCant,BS.cBSCod,BS.cBSDescripcion,CO.cConsDescripcion,PE.cPersNombre, ISNULL(MCTA.cCtaContCod,'')cCtaContCod , ISNULL(MCTA.nMovImporte,0) * -1 nMovImporte" _
         & " , IsNull((Select Top 1 MDFAC.cDocNro From MovDoc MDFAC Where MDFAC.nMovNro = M.nMovNro And MDFAC.nDocTpo = 1),'') FACTURA " _
         & " , IsNull((Select Top 1 MDGUIA.cDocNro From MovDoc MDGUIA Where MDGUIA.nMovNro = M.nMovNro And MDGUIA.nDocTpo = 20),'') GUIA" _
         & " ,isnull(ag.cAgeDescripcion,'')cAgeDescripcion,isnull(a.cAreaDescripcion,'')cAreaDescripcion ," _
         & " isnull(a1.cSubctaCod,'')cSubctaCod " _
         & " From Mov M " _
         & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
         & " Inner Join MovCant MC On MC.nMovNro = M.nMovNro" _
         & " Inner Join MovBS MBS On MC.nMovNro = MBS.nMovNro And MC.nMovItem = MBS.nMovItem And nMovBSOrden = " & pnAlmcen & "   " _
         & " Inner Join BienesServicios BS On MBS.cBSCod = BS.cBSCod" _
         & " Left  Join MovGasto MG On MG.nMovNro = M.nMovNro" _
         & " Left  Join Persona PE On MG.cPersCod = PE.cPersCod" _
         & " left JOIN MovObjAreaAgencia moa on moa.nMovNro =m.nMovNro" _
         & " left JOIN Agencias ag on ag.cAgecod=moa.cAgeCod " _
         & " left JOIN Areas a on a.cAreaCod=moa.cAreaCod " _
         & " left JOIN AreaAgencia a1 ON a1.cAgecod=moa.cAgecod and a1.cAreaCod=moa.cAreaCod " _
         & " Left  Join Movcta MCTA On MCTA.nMovNro = MC.nMovNro And MC.nMovItem = MCTA.nMovItem  And nMovImporte <= 0" _
         & " Left  Join Constante CO On CO.nConsCod = '1019' And CO.nConsValor = BS.nBSUnidad" _
         & " where nMovEstado in ('10','22') and nDocTpo = '71' And nMovFlag = '0' And " _
         & " MD.dDocFecha Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & " 23:59:59'" _
         & sqlOrden
         
      sqlA = " exec stp_sel_GetSalidasAlmacen '" & Format(pdIni, "yyyymmdd") & "','" & Format(pdFin, "yyyymmdd") & "'," & pnAlmcen & "," & lnOrden
      
    '*** FIN PEAC
    
    Set GetSalidasAlmacen = oCon.CargaRecordSet(sqlA)
    
    Set oCon = Nothing
End Function

Public Function GetSalidasAlmacenAF(pdIni As Date, pdFin As Date, Optional pnAlmcen As Integer = 1) As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta

    oCon.AbreConexion
    sqlA = " Select M.nMovNro,MC.nMovItem, M.cMovNro, M.cMovDesc, MD.cDocNro, MD.dDocFecha,MC.nMovCant,BS.cBSCod,BS.cBSDescripcion,CO.cConsDescripcion,PE.cPersNombre, MCTA.cCtaContCod , MCTA.nMovImporte * -1 nMovImporte, MOAA.cAreaCod, MOAA.cAgeCod, BS.bSerie, BS.nPorDeprecia  " _
         & " , IsNull((Select Top 1 MDFAC.cDocNro From MovDoc MDFAC Where MDFAC.nMovNro = M.nMovNro And MDFAC.nDocTpo = 1),'') FACTURA " _
         & " , IsNull((Select Top 1 MDGUIA.cDocNro From MovDoc MDGUIA Where MDGUIA.nMovNro = M.nMovNro And MDGUIA.nDocTpo = 20),'') GUIA" _
         & " From Mov M" _
         & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
         & " Inner Join MovCant MC On MC.nMovNro = M.nMovNro" _
         & " Inner Join MovBS MBS On MC.nMovNro = MBS.nMovNro And MC.nMovItem = MBS.nMovItem And nMovBSOrden = " & pnAlmcen & "   " _
         & " Inner Join MovObjAreaAgencia MOAA On MOAA.nMovNro = MBS.nMovNro And MOAA.nMovItem = MBS.nMovItem " _
         & " Inner Join BienesServicios BS On MBS.cBSCod = BS.cBSCod" _
         & " Left  Join MovGasto MG On MG.nMovNro = M.nMovNro" _
         & " Left  Join Persona PE On MG.cPersCod = PE.cPersCod" _
         & " Left  Join Movcta MCTA On MCTA.nMovNro = MC.nMovNro And MC.nMovItem = MCTA.nMovItem  And nMovImporte <= 0" _
         & " Left  Join Constante CO On CO.nConsCod = '1019' And CO.nConsValor = BS.nBSUnidad" _
         & " where nMovEstado in ('10','22') and nDocTpo = '71' And nMovFlag = '0' And " _
         & " MD.dDocFecha Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & " 23:59:59'" _
         & " And Left(BS.cBSCod,3) in ('112','113') " _
         & " Order by BS.cBsCod"
    Set GetSalidasAlmacenAF = oCon.CargaRecordSet(sqlA)
    
    Set oCon = Nothing
End Function

Public Function GetProveedorSerie(psSerie As String, psBSCod As String) As String
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    oCon.AbreConexion
    
    sqlA = " Select ISNull(dbo.LogObtenerProveedorBSerie('" & psSerie & "','" & psBSCod & "'),'')"
    Set rs = oCon.CargaRecordSet(sqlA)
    
    If rs.EOF And rs.BOF Then
        GetProveedorSerie = ""
    Else
        GetProveedorSerie = rs.Fields(0)
    End If
    
    Set oCon = Nothing
End Function

Public Function GetProveedorFactura(psSerie As String, psBSCod As String) As String
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    oCon.AbreConexion
    
    sqlA = " Select ISNull(dbo.LogObtenerFacturaBSerie('" & psSerie & "','" & psBSCod & "'),'')"
    Set rs = oCon.CargaRecordSet(sqlA)
    
    If rs.EOF And rs.BOF Then
        GetProveedorFactura = ""
    Else
        GetProveedorFactura = rs.Fields(0)
    End If
    
    Set oCon = Nothing
End Function

Public Function GetSerieMov(psBSCod As String, psDoc As String, pnTipoDoc As TpoDoc, pnMovItem As Integer) As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    oCon.AbreConexion
    
    sqlA = " Select Distinct cSerie, IsNull(nIGV,0) nIGV, IsNull(nValor,0) nValor From MovBSSerie a" _
         & " Inner join MovDoc b on a.nMovNro = b.nMovNro" _
         & " Inner join Mov    c on a.nMovNro = c.nMovNro" _
         & " where cBSCod = '" & psBSCod & "' and cdocnro = '" & psDoc & "' And nDocTpo = " & pnTipoDoc & " And a.nMovItem = " & pnMovItem & " And c.nMovFlag = " & MovFlag.gMovFlagVigente
    Set GetSerieMov = oCon.CargaRecordSet(sqlA)
    
    Set oCon = Nothing
End Function

Public Function GetPeriodoDeprecia(psBSCod As String) As Long
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    
    sqlA = "Select nPorDeprecia From BienesServicios Where cBSCod = '" & psBSCod & "'"
    GetPeriodoDeprecia = oCon.CargaRecordSet(sqlA).Fields(0)
    
    Set oCon = Nothing
End Function

Public Function GetCorrelaSerie(psBSCod As String) As String
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    oCon.AbreConexion
    
    sqlA = " Select dbo.LogObtenerCorrelaBSerie('" & psBSCod & "')"
    Set rs = oCon.CargaRecordSet(sqlA)
    
    If IsNull(rs.Fields(0)) Then
        GetCorrelaSerie = "00000000"
    Else
        GetCorrelaSerie = Mid(rs.Fields(0), 6)
    End If
    
    Set oCon = Nothing
End Function

Public Function GetCorrelaSerieSal(psBSCod As String) As String
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    oCon.AbreConexion
    
    sqlA = " Select dbo.LogObtenerCorrelaBSerieSal('" & psBSCod & "')"
    Set rs = oCon.CargaRecordSet(sqlA)
    
    If IsNull(rs.Fields(0)) Then
        GetCorrelaSerieSal = "0"
    Else
        GetCorrelaSerieSal = Val(Mid(rs.Fields(0), 6))
    End If
    
    Set oCon = Nothing
End Function


Public Function InsertMovBsControl(pnAnio As Integer, pnMovNro As Long, pcBSCod As String, pcSerie As String, pnBSValor As Currency, pnBSIGV As Currency, pnBSDeprecia As Currency, pbDepreciado As Integer, pdActivacion As String, pcAreCod As String, pcAgeCod As String, pnBSPerDeprecia As Currency, pnBSPerDepAct As Currency, pcDescripcion As String, pbDeprecia As Integer, pbIntangible As Integer, pbForsado As Integer, pcNroID As String, pcNroPatrimonio As String, pcProveedor As String, pcCtaCont As String, pcFactura As String, pbIngresos As Boolean) As Boolean
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
    Dim sql As String
    
    If pbIngresos Then
        sql = " BSControl "
    Else
        sql = " BSControlSal "
    End If
    
    sqlA = " Insert " & sql & "" _
         & " (nAnio,nMovNro,cBSCod,cSerie," _
         & " nBSValor,nIGV,nBSDeprecia,bDepreciado,dActivacion," _
         & " cAreCod,cAgeCod,nBSPerDeprecia,nBSPerDepAct," _
         & " cDescripcion,bDeprecia,bIntangible,bForsado,cNroID,cNroPatrimonio,cProveedor,cCtaCont,cFactura)" _
         & " Values(" & pnAnio & "," & pnMovNro & ",'" & pcBSCod & "','" & pcSerie & "'," _
         & " " & pnBSValor & "," & pnBSIGV & "," & pnBSDeprecia & "," & pbDepreciado & ",'" & pdActivacion & "','" _
         & "" & pcAreCod & "','" & pcAgeCod & "'," & pnBSPerDeprecia & "," & pnBSPerDepAct & ",'" _
         & "" & pcDescripcion & "'," & pbDeprecia & "," & pbIntangible & "," & pbForsado & ",'" & pcNroID & "','" & pcNroPatrimonio & "','" & pcProveedor & "','" & pcCtaCont & "','" & pcFactura & "')"
         
    oCon.Ejecutar sqlA
End Function

Public Function ValidaMovBsControl(pnAnio As Integer, pnMovNro As Long, pcBSCod As String, pcSerie As String, pbIngresos As Boolean) As Boolean
    Dim sqlA  As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    Dim sql As String
    
    If pbIngresos Then
        sql = " BSControl "
    Else
        sql = " BSControlSal "
    End If
    
    sqlA = " Select nMovNro From " & sql & "" _
         & " Where nAnio = " & pnAnio & " And nMovNro = " & pnMovNro & " And cBSCod = '" & pcBSCod & "' And cSerie = '" & pcSerie & "' "
    Set rs = oCon.CargaRecordSet(sqlA)

    If rs.EOF And rs.BOF Then
        ValidaMovBsControl = False
    Else
        ValidaMovBsControl = True
    End If
    
    Set rs = Nothing
    Set oCon = Nothing
End Function


Public Function EliminaMovBsControl(pdIni As Date, pdFin As Date, Optional pbIngresos As Boolean = True) As Boolean
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
    Dim sql As String
    
    If pbIngresos Then
        sql = " BSControl "
    Else
        sql = " BSControlSal "
    End If
    
    sqlA = " Delete " & sql & "" _
         & " Where dActivacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin + 1, gsFormatoFecha) & "'"
    oCon.Ejecutar sqlA
End Function

Public Function GetRepLogMensual(psNumIngresos As String, psNumSalidas As String, pdFechaAnt As Date, pdFechaFinal As Date, psTitulo As String, pgsNomAge As String, pgsNomEmpresa As String, pgsfecsis As Date, pnTpoProd As Integer, Optional pnAlmcen As Integer = 1) As String
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As New ADODB.Recordset
    Dim rt As New ADODB.Recordset
    Dim cAnioMesAnt As String
    
    Dim lsCadena As String
    
    Dim lnItem As Long
    Dim lnPagina As Long
    
    Dim lnSAnt As Currency
    Dim lnSIng As Currency
    Dim lnSSal As Currency
    Dim lnSAjus As Currency
    Dim lnSNue As Currency
    
    Dim lsCta As String * 12
    Dim lsCtaDesc As String * 23
    Dim lsSAnt As String * 17
    Dim lsSIng As String * 17
    Dim lsSSal As String * 17
    Dim lsSAjus As String * 17
    Dim lsSNue As String * 17
    
    Dim lnAjusteAnt As Currency
    Dim lnMontoCalc As Currency
    
    oCon.AbreConexion
    
    cAnioMesAnt = Format(pdFechaAnt, "YYYYMM" + "01")
    'Adicional - Ñique
    sqlA = "Select distinct c.cCtaContCod,c.nMovImporte From MovCta c Inner Join Mov m On c.nMovNro = m.nMovNro " & _
           " inner join (Select Distinct cObjetoCod as cBSCod, Replace(cCtaContCod,'AG','" & Format(pnAlmcen, "00") & "') cCtaContCod From CtaBS) b on b.cCtaContCod = c.cCtaContCod " & _
           " Where m.cMovNro Like '" & cAnioMesAnt & "%' And m.cOpeCod = '701123'  And m.nMovFlag = 0 " & _
           " order by c.cCtaContCod "

    Set rt = oCon.CargaRecordSet(sqlA)
    
    sqlA = " Select CBS.cCtaContCod, (SELECT cCtaContDesc FROM ctacont ww where ww.cctacontcod = CBS.cCtaContCod) cCtaContDesc, IsNull(CTA.Monto,0) MontoIni, IsNull(CING.Monto,0) MontoIng, IsNull(CSAL.Monto,0) MontoSal, IsNull(CTA1.Monto,0) MontoFinal," _
         & " IsNull((Select Sum(nMovImporte) nMovImporte From MovCta MCA" _
         & "         Inner Join Mov M On MCA.nMovNro = M.nMovNro" _
         & "  Where cMovNro Like '" & Format(DateAdd("d", 1, pdFechaAnt), gsFormatoMovFecha) & "%' And cOpeCod = '701123' And MCA.cCtaContCod = CBS.cCtaContCod" _
         & "        And MCA.cCtaContCod In (Select Distinct Replace(cCtaContCod,'AG','" & Format(pnAlmcen, "00") & "') cCtaContCod From CtaBS) And M.nMovFlag = " & MovFlag.gMovFlagVigente & "),0) As CtaAjuste " _
         & " From (Select distinct Replace(cCtaContCod,'AG','" & Format(pnAlmcen, "00") & "') cCtaContCod From CtaBS) As CBS" _
         & " Left Join ( Select Sum(nMOnto) MOnto, cta from " _
         & "                (Select *, (Select Top 1 Replace(cCtaContCod,'AG','" & Format(pnAlmcen, "00") & "') cCtaContCod from ctabs where BS.cBSCod Like cObjetoCod + '%' And cOpeCod In (Select cOpeCod from LogOpeProd Where nProdTpo = " & pnTpoProd & ") order by len(cObjetoCod) desc ) Cta" _
         & "             From BSSaldos BS Where nAlmCod = " & pnAlmcen & " And Datediff(day,dSaldo,'" & Format(pdFechaAnt, gsFormatoFecha) & "') = 0 ) As T" _
         & "             Group by cta) As CTA On CTA.Cta = CBS.cCtaContCod" _
         & " Left Join ( Select Sum(nMovImporte) Monto, cCtaContCod From MovCta" _
         & "             Where nMovNro in ('" & psNumIngresos & "') And nMovImporte > 0" _
         & "             Group By cCtaContCod ) As CING On CBS.cCtaContCod = CING.cCtaContCod" _
         & " Left Join ( Select Sum(nMovImporte) Monto, cCtaContCod From MovCta" _
         & " Where nMovNro in ('" & psNumSalidas & "') And nMovImporte < 0 Group By cCtaContCod" _
         & " ) As CSAL On CBS.cCtaContCod = CSAL.cCtaContCod" _
         & " Left Join ( Select Sum(nMOnto) MOnto, cta from" _
         & "                (Select *, (Select Top 1 Replace(cCtaContCod,'AG','" & Format(pnAlmcen, "00") & "') cCtaContCod from ctabs where BS.cBSCod Like cObjetoCod + '%' And cOpeCod In (Select cOpeCod from LogOpeProd Where nProdTpo = " & pnTpoProd & ") order by len(cObjetoCod) desc) Cta" _
         & "            From BSSaldos BS Where nAlmCod = " & pnAlmcen & " And nAlmTpo = " & pnTpoProd & " And Datediff(day,dSaldo,'" & Format(pdFechaFinal, gsFormatoFecha) & "') = 0 ) As T" _
         & " Group by cta) As CTA1 On CTA1.Cta = CBS.cCtaContCod" _
         & " Order by CBS.cCtaContCod"
    Set rs = oCon.CargaRecordSet(sqlA)
        
    lsCadena = ""
    lnSAnt = 0
    lnSIng = 0
    lnSSal = 0
    lnSAjus = 0
    lnSNue = 0

    lsCadena = lsCadena & CabeceraPagina(psTitulo, lnPagina, lnItem, pgsNomAge, pgsNomEmpresa, pgsfecsis)
    lsCadena = lsCadena & Encabezado("Cta.Cont;10; ;5;Descripcion;20; ;5;Saldo.Ant;15; ;6;Saldo.Ing;15; ;6;Saldo.Sal;15; ;6;Saldo.Ajus;15; ;6;Saldo.Fin;15;", lnItem)
    
    While Not rs.EOF
        RSet lsCta = rs!cCtaContCod
        
        'Ñique - Inicio Modificacion ------------------------------------------------------
        'Verifica si ha tenido ajuste el mes anterior
        
        lnAjusteAnt = 0
        If Not (rt.EOF And rt.BOF) Then
           rt.MoveFirst
           rt.Find "cCtaContCod = '" & Trim(rs!cCtaContCod) & "'", , adSearchForward, 1
           If Not rt.EOF Then
              Do While rt!cCtaContCod = Trim(rs!cCtaContCod)
                 lnAjusteAnt = lnAjusteAnt + rt!NMOVIMPORTE
                 rt.MoveNext
                 If rt.EOF Then
                    Exit Do
                 End If
              Loop
           End If
        End If
        
        'Ñique - Fin Modificacion ---------------------------------------------------------
        
        RSet lsSAnt = Format(rs!MontoIni + lnAjusteAnt, "#,##0.00")
        RSet lsSIng = Format(rs!MontoIng, "#,##0.00")
        RSet lsSSal = Format(rs!MontoSal, "#,##0.00")
        lsCtaDesc = rs!cCtaContDesc & ""
        
        If (rs!MontoIni <> 0 Or rs!MontoIng Or rs!MontoSal <> 0) And Month(pdFechaFinal) = 1 Then
            RSet lsSAjus = Format(rs!CtaAjuste, "#,##0.00")
        Else
            RSet lsSAjus = Format(0, "#,##0.00")
        End If
        If rs!cCtaContCod = "19111801" Then
            RSet lsSAjus = Format(0, "#,##0.00")
        End If
        
        RSet lsSNue = Format(rs!MontoFinal, "#,##0.00")
        
        lnSAnt = lnSAnt + rs!MontoIni
        lnSIng = lnSIng + rs!MontoIng
        lnSSal = lnSSal + rs!MontoSal
        If (rs!MontoIni <> 0 Or rs!MontoIng Or rs!MontoSal <> 0) And Month(pdFechaFinal) = 1 Then
            If rs!cCtaContCod <> "19111801" Then
                lnSAjus = lnSAjus + rs!CtaAjuste
                RSet lsSNue = Format(rs!MontoFinal + rs!CtaAjuste, "#,##0.00")
                lnSNue = lnSNue + rs!CtaAjuste
            End If
        End If
        lnSNue = lnSNue + rs!MontoFinal
    
        'Ñique - Verificación adicional de cálculos ---------------------------------
        
        lnMontoCalc = rs!MontoIni + lnAjusteAnt + rs!MontoIng + rs!MontoSal
        If lnMontoCalc <> Format(lsSNue, "########0.00") Then
           MsgBox "Diferencia entre lo calculado " + Format(lnMontoCalc - lsSNue, "########0.00") + " en la cuenta " + rs!cCtaContCod + "..." + Space(10), vbInformation
        End If
        'lsCadena = lsCadena & lsCta & "    " & lsSAnt & "    " & lsSIng & " " & lsSSal & " " & lsSAjus & " " & lsSNue & " " & lnMontoCalc & oImpresora.gPrnSaltoLinea
        'Fin verificacion -----------------------------------------------------------
        lsCadena = lsCadena & lsCta & "  " & lsCtaDesc & "  " & lsSAnt & "    " & lsSIng & "    " & lsSSal & "    " & lsSAjus & "    " & lsSNue & oImpresora.gPrnSaltoLinea
        
        lnItem = lnItem + 1
        
        If lnItem > 54 Then
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
            lsCadena = lsCadena & CabeceraPagina(psTitulo, lnPagina, lnItem, pgsNomAge, pgsNomEmpresa, pgsfecsis)
            lsCadena = lsCadena & Encabezado("Cta.Cont;15; ;5;Saldo.Ant;15; ;6;Saldo.Ing;15; ;6;Saldo.Sal;15; ;6;Saldo.Ajus;15; ;6;Saldo.Fin;15;", lnItem)
        End If
        
        rs.MoveNext
    Wend
    
    lsCta = ""
    RSet lsSAnt = Format(lnSAnt, "#,##0.00")
    RSet lsSIng = Format(lnSIng, "#,##0.00")
    RSet lsSSal = Format(lnSSal, "#,##0.00")
    RSet lsSAjus = Format(lnSAjus, "#,##0.00")
    RSet lsSNue = Format(lnSNue, "#,##0.00")
    
    lsCadena = lsCadena & Encabezado(" ;35; ;6;" & Trim(lsSAnt) & ";15; ;6;" & Trim(lsSIng) & ";15; ;6;" & Trim(lsSSal) & ";15; ;6;" & Trim(lsSAjus) & ";15; ;6;" & Trim(lsSNue) & ";15;", lnItem)
    
    'RSet lsSAnt = Format(lnSAnt, "#,##0.00")
    'RSet lsSIng = Format(lnSIng, "#,##0.00")
    'RSet lsSSal = Format(lnSSal, "#,##0.00")
    'RSet lsSAjus = Format(lnSAjus, "#,##0.00")
    'RSet lsSNue = Format(lnSNue, "#,##0.00")
    
    'lsCadena = lsCadena & lsCta & "    " & lsSAnt & "    " & lsSIng & "    " & lsSSal & "    " & lsSAjus & "    " & lsSNue & oImpresora.gPrnSaltoLinea
    
    Set oCon = Nothing
    Set rs = Nothing
    
    GetRepLogMensual = lsCadena
End Function

Public Function GetSerieBien(psBienCod As String, psAlmacen As String, psIncluidos As String) As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    sqlA = " Select BSS.cSerie, BS.cBSDescripcion From Mov M" _
         & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro" _
         & " Inner Join MovBSSerie MBSS On MBSS.nMovNro = MBS.nMovNro And MBSS.nMovItem = MBS.nMovItem " _
         & " Inner Join BienesServicios BS On MBSS.cBSCod = BS.cBSCod " _
         & " Inner Join BSSerie BSS On BSS.cBSCod = MBSS.cBSCod And BSS.cSerie = MBSS.cSerie" _
         & " where BSS.cBSCod = '" & psBienCod & "' And nBienEstado in (0,2) And M.nMovEstado = '20' And nMovBSOrden  = '" & psAlmacen & "' And BSS.cSerie Not In (" & psIncluidos & ")" _
         & " Order By BSS.cSerie"
         
         Set GetSerieBien = oCon.CargaRecordSet(sqlA)
    
    Set oCon = Nothing
End Function

Public Function VerfBienSerie(psBSCod As String) As Boolean
'    Dim sqlA  As String
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'
'    oCon.AbreConexion
'    sqlA = " Select RTrim(BS.cBSCod) Codigo, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])', len(BS.cBSCod) Nivel " _
'         & " From BienesServicios BS" _
'         & " Inner Join ProveedorBS PBS On BS.cBSCod like PBS.cBSCod + '%'" _
'         & " Inner Join Constante Co On CO.nConsValor = BS.nBSUnidad  And CO.nConsCOd = '1019'" _
'         & " Where cPersCod = '" & psPersCod & "'"
'    Set GetBienesProv = oCon.CargaRecordSet(sqlA)
'
'    Set oCon = Nothing
End Function

Public Function GetStock1(ByVal psAlmacen As String, psBSCod As String) As Double
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
  
   If oCon.AbreConexion() Then
        sql = " select dbo.LogGetStock(" & psAlmacen & ",'" & psBSCod & "')"
        GetStock1 = oCon.CargaRecordSet(sql).Fields(0)
   End If
   
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function SetEstadoSubasta(ByVal pbRechazo As Boolean, pnMovNro As Long) As Integer
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
  
   If oCon.AbreConexion() Then
        sql = " Update BSSubasta " _
            & " Set nSubEstado = " & IIf(pbRechazo, 1, 2) & "" _
            & " Where nMovNro = " & pnMovNro
        oCon.Ejecutar sql
   End If
   
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetPrePromedio(ByVal psAlmacen As String, psBSCod As String, pnProdTpo As Integer) As Double
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
  
   If oCon.AbreConexion() Then
        sql = " select dbo.LogGetPrecioPromedio(" & psAlmacen & ",'" & psBSCod & "'," & pnProdTpo & ")"
        Set rs = oCon.CargaRecordSet(sql)
        
        If rs.EOF And rs.BOF Then
            GetPrePromedio = 0
        Else
            GetPrePromedio = rs.Fields(0)
        End If
   End If
   
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetLogKardex(ByVal psAlmacen As String, psBSCod As String, pdFecIni As Date, pdFecFin As Date, pnProdTpo As Integer, Optional psDocTpoIng As String = "42", Optional pbMuestraExtornados As Boolean = False) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim pnOrden As Integer 'NAGL TIC1801160002
   Dim sql As String
   
   '***NAGL TIC1801160002****'
   If Not pbMuestraExtornados Then
        pnOrden = 1
   Else
        pnOrden = 2
   End If
   '***NAGL 20180110***'
   If oCon.AbreConexion() Then
        'sql = " Select M.nMovNro, MB.nMovItem, M.nMovEstado, convert(datetime,left(cMovNro,8)) Fecha, MD.cDocNro, MCANT.nMovCant , PE.cPersNombre, MC.nMovImporte/ Case MCANT.nMovCant when 0 then 1 else MCANT.nMovCant end PrePromedio, MC.nMovImporte, MOAA.cAreaCod, MOAA.cAgeCod, Area.cAreaDescripcion , AGE.cAgeDescripcion, BS.cBSDescripcion From Mov M" _
            & " Inner Join MovBS MB On M.nMovNro = MB.nMovNro And nMovBSOrden = " & psAlmacen & "" _
            & " Left Join MovCTA MC On M.nMovNro = MC.nMovNro And MB.nMovItem = MC.nMovItem" _
            & " Inner Join MovCant MCANT On MB.nMovNro = MCANT.nMovNro And MB.nMovItem = MCANT.nMovItem And MCANT.nMovCant > 0 " _
            & " Inner Join BienesServicios BS On MB.cBSCod = BS.cBSCod" _
            & " Left Join MovDoc MD On M.nMovNro = MD.nMovNro And MD.nDocTpo = " & psDocTpoIng & "" _
            & " Left Join MovGasto MG On M.nMovNro = MG.nMovNro" _
            & " Left Join Persona PE On MG.cPersCod = PE.cPersCod" _
            & " Left Join MovObjAreaAgencia MOAA On MB.nMovNro = MOAA.nMovNro And MB.nMovItem = MOAA.nMovItem" _
            & " Left Join Areas AREA On AREA.cAreaCod = MOAA.cAreaCod" _
            & " Left Join Agencias AGE On AGE.cAgeCod = MOAA.cAgeCod" _
            & " Where M.nMovEstado in ('20') And MB.cBSCod = '" & psBSCod & "' And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')  And convert(datetime,left(cMovNro,8)) Between '" & Format(pdFecIni, gsFormatoFecha) & "' And '" & Format(pdFecFin, gsFormatoFecha) & "' And M.cOpeCod In (Select cOpeCod from LogOpeProd Where nOpeTpo = 1 And nProdTpo = " & pnProdTpo & ")" _
            & " Union" _
            & " Select M.nMovNro, MB.nMovItem, M.nMovEstado, convert(datetime,left(cMovNro,8)) Fecha, MD.cDocNro, MCANT.nMovCant , PE.cPersNombre, MC.nMovImporte/ Case MCANT.nMovCant when 0 then 1 else MCANT.nMovCant end PrePromedio, MC.nMovImporte, MOAA.cAreaCod, MOAA.cAgeCod, Area.cAreaDescripcion , AGE.cAgeDescripcion, BS.cBSDescripcion From Mov M" _
            & " Inner Join MovBS MB On M.nMovNro = MB.nMovNro And nMovBSOrden = " & psAlmacen & "" _
            & " Left  Join MovCTA MC On M.nMovNro = MC.nMovNro And MB.nMovItem = MC.nMovItem" _
            & " Inner Join MovCant MCANT On MB.nMovNro = MCANT.nMovNro And MB.nMovItem = MCANT.nMovItem" _
            & " Inner Join BienesServicios BS On MB.cBSCod = BS.cBSCod" _
            & " Left Join MovDoc MD On M.nMovNro = MD.nMovNro " _
            & " Left Join MovGasto MG On M.nMovNro = MG.nMovNro" _
            & " Left Join Persona PE On MG.cPersCod = PE.cPersCod" _
            & " Left Join MovObjAreaAgencia MOAA On MB.nMovNro = MOAA.nMovNro And MB.nMovItem = MOAA.nMovItem" _
            & " Left Join Areas AREA On AREA.cAreaCod = MOAA.cAreaCod" _
            & " Left Join Agencias AGE On AGE.cAgeCod = MOAA.cAgeCod" _
            & " Where M.nMovEstado in ('10','22') And MB.cBSCod = '" & psBSCod & "' And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')   And convert(datetime,left(cMovNro,8)) Between '" & Format(pdFecIni, gsFormatoFecha) & "' And '" & Format(pdFecFin, gsFormatoFecha) & "' And M.cOpeCod In (Select cOpeCod from LogOpeProd Where nOpeTpo = 2 And nProdTpo = " & pnProdTpo & ") Order By convert(datetime,left(cMovNro,8)),  M.nMovNro"
            
       '-------------------------------------------------------------------
       'MODIFICACION: realizada por error en la conversión a tipo Fecha para
       '              el Bien 1110502702
       '       FECHA: 26-Mayo-2005
       '-------------------------------------------------------------------
       '***********************Comentado by NAGL 20170110*************************
        'sql = " Select left(cMovNro,8),M.nMovNro, MB.nMovItem, M.nMovEstado, convert(datetime,left(cMovNro,8)) Fecha, MD.cDocNro, MCANT.nMovCant , PE.cPersNombre, MC.nMovImporte/ Case MCANT.nMovCant when 0 then 1 else MCANT.nMovCant end PrePromedio, MC.nMovImporte, MOAA.cAreaCod, MOAA.cAgeCod, Area.cAreaDescripcion , AGE.cAgeDescripcion, BS.cBSDescripcion, M.nMovFlag From Mov M" _
            '& " Inner Join MovBS MB On M.nMovNro = MB.nMovNro And nMovBSOrden = " & psAlmacen & "" _
            '& " Left Join MovCTA MC On M.nMovNro = MC.nMovNro And MB.nMovItem = MC.nMovItem" _
            '& " Inner Join MovCant MCANT On MB.nMovNro = MCANT.nMovNro And MB.nMovItem = MCANT.nMovItem And MCANT.nMovCant > 0 " _
            '& " Inner Join BienesServicios BS On MB.cBSCod = BS.cBSCod" _
            '& " Left Join MovDoc MD On M.nMovNro = MD.nMovNro And MD.nDocTpo = " & psDocTpoIng & "" _
            '& " Left Join MovGasto MG On M.nMovNro = MG.nMovNro" _
            '& " Left Join Persona PE On MG.cPersCod = PE.cPersCod" _
            '& " Left Join MovObjAreaAgencia MOAA On MB.nMovNro = MOAA.nMovNro And MB.nMovItem = MOAA.nMovItem" _
            '& " Left Join Areas AREA On AREA.cAreaCod = MOAA.cAreaCod" _
            '& " Left Join Agencias AGE On AGE.cAgeCod = MOAA.cAgeCod" _
            '& " Where MB.cBSCod = '" & psBSCod & "' " & IIf(pbMuestraExtornados, " And cOpeCod <> '591101'", " And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') And  M.nMovEstado in ('20')  ") & " And (left(cMovNro,8) Between '" & Format(pdFecIni, "YYYYMMDD") & "' And '" & Format(pdFecFin, "YYYYMMDD") & "') And M.cOpeCod In (Select cOpeCod from LogOpeProd Where nOpeTpo = 1 And nProdTpo = " & pnProdTpo & ")" _
            '& " Union" _
            '& " Select left(cMovNro,8),M.nMovNro, MB.nMovItem, M.nMovEstado, convert(datetime,left(cMovNro,8)) Fecha, MD.cDocNro, MCANT.nMovCant , PE.cPersNombre, MC.nMovImporte/ Case MCANT.nMovCant when 0 then 1 else MCANT.nMovCant end PrePromedio, MC.nMovImporte, MOAA.cAreaCod, MOAA.cAgeCod, Area.cAreaDescripcion , AGE.cAgeDescripcion, BS.cBSDescripcion, M.nMovFlag  From Mov M" _
            '& " Inner Join MovBS MB On M.nMovNro = MB.nMovNro And nMovBSOrden = " & psAlmacen & "" _
            '& " Left  Join MovCTA MC On M.nMovNro = MC.nMovNro And MB.nMovItem = MC.nMovItem" _
            '& " Inner Join MovCant MCANT On MB.nMovNro = MCANT.nMovNro And MB.nMovItem = MCANT.nMovItem" _
            '& " Inner Join BienesServicios BS On MB.cBSCod = BS.cBSCod" _
            '& " Left Join MovDoc MD On M.nMovNro = MD.nMovNro " _
            '& " Left Join MovGasto MG On M.nMovNro = MG.nMovNro" _
            '& " Left Join Persona PE On MG.cPersCod = PE.cPersCod" _
            '& " Left Join MovObjAreaAgencia MOAA On MB.nMovNro = MOAA.nMovNro And MB.nMovItem = MOAA.nMovItem" _
            '& " Left Join Areas AREA On AREA.cAreaCod = MOAA.cAreaCod" _
            '& " Left Join Agencias AGE On AGE.cAgeCod = MOAA.cAgeCod" _
            '& " Where M.nMovEstado in ('10','22') And MB.cBSCod = '" & psBSCod & "' " & IIf(pbMuestraExtornados, "", " And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') ") & " And (left(cMovNro,8) Between '" & Format(pdFecIni, "YYYYMMDD") & "' And '" & Format(pdFecFin, "YYYYMMDD") & "') And M.cOpeCod In (Select cOpeCod from LogOpeProd Where nOpeTpo = 2 And nProdTpo = " & pnProdTpo & ") Order By left(cMovNro,8),  M.nMovNro"
          '************************END NAGL***************************************
      sql = "Exec stp_sel_ObtieneLogKardex '" & psAlmacen & "', '" & psBSCod & "', '" & Format(pdFecIni, "yyyymmdd") & "', '" & Format(pdFecFin, "yyyymmdd") & "', " & pnProdTpo & ", '" & psDocTpoIng & "', " & pnOrden & ""
      '**********NAGL TIC1801160002*******************
      Set GetLogKardex = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function InsertaBSSaldos(ByVal psAlmacen As String, psBSCod As String, pdFecha As Date, pnStock As Long, pnValor As Currency, pnProdTpo As Integer) As Boolean
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Delete BSSaldos" _
          & " Where nAlmCod = " & psAlmacen & " And cBSCod = '" & psBSCod & "' And dSaldo Between '" & Format(pdFecha, gsFormatoFecha) & "' And '" & Format(pdFecha & " 23:59:59", gsFormatoFecha) & "' And nAlmTpo = " & pnProdTpo
      oCon.Ejecutar sql
      
      sql = " Insert BSSaldos(nAlmCod, cBSCod, dSaldo, nStock, nMonto, nAlmTpo)" _
          & " Values(" & psAlmacen & ",'" & psBSCod & "','" & Format(pdFecha, gsFormatoFecha) & "'," & pnStock & "," & pnValor & "," & pnProdTpo & ")"
      oCon.Ejecutar sql
      
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function InsertaBSSaldosDia(ByVal psAlmacen As String, prs As ADODB.Recordset, pdFecha As Date, pnProdTpo As Integer) As Boolean
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
  'ANPS*******
   Dim pssAlmCod As Integer
   Dim psAlmtpo As Integer
   Dim psBSCod As String
   Dim psStock As Currency
   Dim psMonto As Currency

     If oCon.AbreConexion() Then
      prs.MoveFirst
      While Not prs.EOF
     
        If prs.Fields(0) = "1" Then
            pssAlmCod = psAlmacen
            psAlmtpo = pnProdTpo
            psBSCod = prs.Fields(1)
            psStock = prs.Fields(3)
            psMonto = prs.Fields(4)
      sql = " exec stp_int_BSSaldos '" & Format(pdFecha, "yyyymmdd") & "'," & pssAlmCod & "," & psAlmtpo & ",'" & psBSCod & "'," & psStock & "," & psMonto
      
       oCon.CargaRecordSet (sql)
      
      End If
        prs.MoveNext
      Wend
      oCon.CierraConexion
   End If
   
   
'Comentado anps
'   If oCon.AbreConexion() Then
'      prs.MoveFirst
'      'COMENTADO ANPS
''      sql = " Delete BSSaldos " _
''          & " Where nAlmCod = " & psAlmacen & " And dSaldo Between '" & Format(pdFecha, gsFormatoFecha) & "' And '" & Format(pdFecha & " 23:59:59", gsFormatoFecha) & "' And nAlmTpo = " & pnProdTpo
''      oCon.Ejecutar sql
'
'      While Not prs.EOF
'      'ANPS*******
'        If prs.Fields(0) = "1" Then
''            sql = " Insert BSSaldos(nAlmCod, cBSCod, dSaldo, nStock, nMonto, nAlmTpo)" _
''                & " Values(" & psAlmacen & ",'" & prs.Fields(1) & "','" & Format(pdFecha, gsFormatoFecha) & "'," & prs.Fields(3) & "," & prs.Fields(4) & "," & pnProdTpo & ")"
'
'           sql = " declare @count int " _
'            & " set @count=(select count(cBSCod) from BSSaldos where  nAlmCod =" & psAlmacen & "  And dSaldo ='" & Format(pdFecha, gsFormatoFecha) & "' And nAlmTpo = " & pnProdTpo & " and cBSCod='" & prs.Fields(1) & "')" _
'            & " if @count = 1 " _
'               & " begin " _
'                    & " update bs set bs.nStock=" & prs.Fields(3) & " , bs.nMonto=" & prs.Fields(4) & " from BSSaldos bs where  bs.nAlmCod =" & psAlmacen & "   And bs.dSaldo ='" & Format(pdFecha, gsFormatoFecha) & "' And bs.nAlmTpo = " & pnProdTpo & " and bs.cBSCod='" & prs.Fields(1) & "'" _
'                & " End " _
'            & " Else " _
'                & " begin " _
'                 & "   Insert BSSaldos(nAlmCod, cBSCod, dSaldo, nStock, nMonto, nAlmTpo) " _
'                 & "   Values(" & psAlmacen & ",'" & prs.Fields(1) & "','" & Format(pdFecha, gsFormatoFecha) & "'," & prs.Fields(3) & "," & prs.Fields(4) & "," & pnProdTpo & ")" _
'                & " End " _
'        'FIN ANPS*********
'            oCon.Ejecutar sql
'        End If
'        prs.MoveNext
'      Wend
'      oCon.CierraConexion
'   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function
Public Function GetLogMovimientoDia(ByVal psAlmacen As String, psBSCod As String, pdFecha As Date) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
    
   If oCon.AbreConexion() Then
      sql = " Select nSaldo, nMonto From dbo.LogGetMovimientoDia(" & psAlmacen & ",'" & psBSCod & "','" & Format(pdFecha, gsFormatoFecha) & "')"
      Set GetLogMovimientoDia = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetLogMovimientoDiaRango(ByVal psAlmacen As String, psBSCod As String, pnProdTpo As Integer, pdFecha As Date, pdFechaFin As Date) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
    
   If oCon.AbreConexion() Then
      sql = " Select nSaldo, nMonto From dbo.LogGetMovimientoRangoDia(" & psAlmacen & ",'" & psBSCod & "','" & Format(pdFecha, gsFormatoFecha) & "','" & Format(pdFechaFin, gsFormatoFecha) & "'," & pnProdTpo & ")"
      Set GetLogMovimientoDiaRango = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetStock(ByVal psAlmacen As String, psBSCod As String, pnProdTpo As Integer, Optional pdFecha As Date = "01/01/1899", Optional pnMontoIni As Currency = 0) As Double
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim sql As String
  
   If oCon.AbreConexion() Then
        If pdFecha = CDate("01/01/1899") Then
            sql = " select dbo.LogGetStock(" & psAlmacen & ",'" & psBSCod & "'," & pnProdTpo & ")"
            Set rs = oCon.CargaRecordSet(sql)
            GetStock = rs.Fields(0)
        Else
            sql = " Select Top 1 nStock, Convert(Varchar(10),dSaldo,103) dSaldo , nMonto, dSaldo Fecha From BSSaldos Where nAlmTpo = " & pnProdTpo & " And cBSCod = '" & psBSCod & "' And nAlmCod = " & psAlmacen & " And dSaldo < '" & Format(pdFecha, gsFormatoFecha) & "' order by Fecha Desc"
            Set rs = oCon.CargaRecordSet(sql)
            
            If rs.EOF And rs.BOF Then
                GetStock = 0
                pnMontoIni = 0
                pdFecha = DateAdd("d", -1, DateAdd("m", -1, DateAdd("d", 1, pdFecha)))
            Else
                pdFecha = Format(rs.Fields(1), gsFormatoFechaView)
                pnMontoIni = rs.Fields(2)
                GetStock = rs.Fields(0)
            End If
        End If
   End If
   
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function VerfBSCorrela(psBSCod As String) As Boolean
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
    
   If oCon.AbreConexion() Then
      sql = " Select bCorrela From BienesServicios Where cBSCod = '" & psBSCod & "'"
      Set rs = oCon.CargaRecordSet(sql)
      
      If rs!bCorrela Then
        VerfBSCorrela = True
      Else
        VerfBSCorrela = False
      End If
      
      oCon.CierraConexion
   End If
   
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetBSCorrelaIni(psBSCod As String) As Long
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim lnNum As Long
   Dim lnNumAF As Long
       
   lnNumAF = CLng(GetCorrelaSerie(psBSCod)) + 1
   
   If oCon.AbreConexion() Then
      sql = " Select  Max(SubString(cSerie,6,8)) cSerie From BSserie Where cBSCOd = '" & psBSCod & "' And IsNumeric(SubString(cSerie,6,8)) = 1"
      Set rs = oCon.CargaRecordSet(sql)
      
      If rs.EOF And rs.BOF Then
            lnNum = 1
      Else
            If IsNull(rs!cSerie) Then
                lnNum = 1
                GetBSCorrelaIni = 1
            Else
                lnNum = Val(rs!cSerie) + 1
                GetBSCorrelaIni = Val(rs!cSerie) + 1
            End If
            
      End If
      
      If lnNumAF > lnNum Then
            GetBSCorrelaIni = lnNumAF
      Else
            GetBSCorrelaIni = lnNum
      End If
      
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetBSCorrelaIniMov(psBSCod As String, pnMovNro As Long) As Long
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim lnNum As Long
   Dim lnNumAF As Long
       
   lnNumAF = CInt(GetCorrelaSerie(psBSCod)) + 1
   
   If oCon.AbreConexion() Then
      sql = " Select Top 1 SubString(cSerie,6,8) cSerie  from MovBSSerie Where cBSCOd = '" & psBSCod & "' And IsNumeric(SubString(cSerie,6,8)) = 1 And nMOvNro = " & pnMovNro & " Order by Convert(int,SubString(cSerie,6,8)) Asc"
      Set rs = oCon.CargaRecordSet(sql)
      
      If rs.EOF And rs.BOF Then
            lnNum = 1
      Else
            GetBSCorrelaIniMov = Val(rs!cSerie)
      End If
        
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

'Public Function GetStockEmbargo(ByVal psAlmacen As String, psBSCod As String, Optional pdFecha As Date = "01/01/1899", Optional pnMontoIni As Currency = 0) As Double
'   On Error GoTo CargaOpeGruErr
'   Dim oCon As DConecta
'   Set oCon = New DConecta
'   Dim rs As ADODB.Recordset
'   Set rs = New ADODB.Recordset
'   Dim sql As String
'
'   If oCon.AbreConexion() Then
'        If pdFecha = CDate("01/01/1899") Then
'            sql = " select dbo.LogGetStockEmbargo(" & psAlmacen & ",'" & psBSCod & "')"
'            Set rs = oCon.CargaRecordSet(sql)
'            GetStockEmbargo = rs.Fields(0)
'        Else
'            sql = " Select Top 1 nStock, Convert(Varchar(10),dSaldo,103) dSaldo , nMonto From BSSaldos Where cBSCod = '" & psBSCod & "' And nAlmCod = " & psAlmacen & " And dSaldo < '" & Format(pdFecha, gsFormatoFecha) & "' order by dSaldo Desc"
'            Set rs = oCon.CargaRecordSet(sql)
'
'            If rs.EOF And rs.BOF Then
'                GetStockEmbargo = 0
'                pnMontoIni = 0
'            Else
'                pdFecha = Format(rs.Fields(1), gsFormatoFechaView)
'                pnMontoIni = rs.Fields(2)
'                GetStockEmbargo = rs.Fields(0)
'            End If
'        End If
'   End If
'
'   Set oCon = Nothing
'   Exit Function
'CargaOpeGruErr:
'   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
'End Function

'Public Function GetAFBSSerie(psBSCod As String, pnAnio As Integer, Optional psTipoRepo As String = "", Optional pnTodos As Integer = 0) As ADODB.Recordset
Public Function GetAFBSSerie(psBSCod As String, pnAnio As Integer, Optional psTipoRepo As String = "", Optional pnTodos As Integer = 0, Optional ByVal psAreaAgeCod As String = "") As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
    
   If oCon.AbreConexion() Then
      
'      sql = " Select cSerie, convert(varchar(10),dCompra,103), nBSValor V_His, cDescripcion   From BSActivoFijo BSAF" _
          & " Inner Join BienesServicios BS On BSAF.cBSCod = BS.cBSCod Where BS.cBSCod = '" & psBSCod & "' And nAnio = " & pnAnio & " And bBaja = 0 Order by dCompra, cSerie"
    '*** PEAC 20101129
    'sql = " exec stp_sel_GetAFBSSerie '" & psBSCod & "'," & pnAnio & ",'" & psTipoRepo & "'," & pnTodos
    sql = " exec stp_sel_GetAFBSSerie '" & psBSCod & "'," & pnAnio & ",'" & psTipoRepo & "'," & pnTodos & ",'" & psAreaAgeCod & "'" 'EJVG20130626
          
      Set GetAFBSSerie = oCon.CargaRecordSet(sql)
      
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetBNDBSSerie(psBSCod As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
    
   If oCon.AbreConexion() Then
      sql = " Select cSerie, convert(varchar(10),dCompra,103), nBSValor V_His, BS.cBSDescripcion From BSNoDepreciables BSAF" _
          & " Inner Join BienesServicios BS On BSAF.cBSCod = BS.cBSCod Where BS.cBSCod = '" & psBSCod & "' And bBaja = 0 Order by dCompra, cSerie"
      Set GetBNDBSSerie = oCon.CargaRecordSet(sql)
      
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetAFBSDetalle(psBSCod As String, psSerie As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
    
   If oCon.AbreConexion() Then
      
      'sql = "Select Top 1 cAreCod, cAgeCod, nMovNro, nAnio From BSActivoFijo BSAF Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "' Order By nAnio Desc"
      
      '*** PEAC 20101129
      sql = " exec stp_sel_GetAFBSDetalle '" & psBSCod & "','" & psSerie & "'"
      
      Set GetAFBSDetalle = oCon.CargaRecordSet(sql)
      
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetBNDBSDetalle(psBSCod As String, psSerie As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
    
   If oCon.AbreConexion() Then
      sql = "Select Top 1 cAreCod, cAgeCod From BSNoDepreciables BSAF Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
      Set GetBNDBSDetalle = oCon.CargaRecordSet(sql)
      
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetAFBSDetallePersona(psBSCod As String, psSerie As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
    
   If oCon.AbreConexion() Then
      sql = "select Top 1 cPersCod, IsNull((Select cPersNombre From Persona WHere cPersCod = BSAF.cPersCod),'') NomPers, nMovNro, nAnio From BSActivoFijo BSAF Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "' Order By nAnio Desc"
      Set GetAFBSDetallePersona = oCon.CargaRecordSet(sql)
      
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetBNDBSDetallePersona(psBSCod As String, psSerie As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
    
   If oCon.AbreConexion() Then
      sql = "Select Top 1 cPersCod, IsNull((Select cPersNombre From Persona WHere cPersCod = BSAF.cPersCod),'') NomPers From BSNoDepreciables BSAF Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "' "
      Set GetBNDBSDetallePersona = oCon.CargaRecordSet(sql)
      
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function


Public Function GetAFBienes() As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta

    oCon.AbreConexion

    sqlA = " Select Distinct RTrim(BS.cBSCod) Codigo, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])', len(BS.cBSCod) Nivel " _
         & " From BienesServicios BS" _
         & " Inner Join Constante Co On CO.nConsValor = BS.nBSUnidad  And CO.nConsCOd = '1019'" _
         & " Inner Join BSActivoFijo BSAF On BSAF.cBSCod = BS.cBSCod  Where BSAF.bBaja = 0 and len(BS.cBSCod)>5 " _
         & " Union " _
         & " Select BS.cBsCod, cBSDescripcion, 2 Nivel  from bienesservicios BS Where cBSCod Like '112__'" _
         & " UNION" _
         & " Select C.cCtaContCod Codigo, C.cCtaContDesc Descripcion, 3 Nivel From CtaCont C WITH (NOLOCK) " _
         & " Inner Join CtasIntangibles CI WITH (NOLOCK) On C.cCtaContCod = CI.cCtaContCod" _
         & " ORDER BY RTRIM(BS.cBSCod) "

    Set GetAFBienes = oCon.CargaRecordSet(sqlA)

    Set oCon = Nothing
End Function
                

Public Function GetBienesAlmacen(Optional psAlmacen As String = "", Optional psSubProducto As String = "", Optional pnAddCtasIntangible As Boolean = False, Optional pbFilfroPedidoAgencia As Boolean = False) As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
     
    oCon.AbreConexion
     
    If psAlmacen <> "" Then
        If psSubProducto = "" Then
            sqlA = " Select RTrim(BS.cBSCod) Codigo, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])', len(BS.cBSCod) Nivel " _
                 & " From BienesServicios BS  WITH (NOLOCK)  " _
                 & " Inner Join AlmacenBS PBS WITH (NOLOCK) On BS.cBSCod like PBS.cBSCod + '%'" _
                 & " Inner Join Constante CO  WITH (NOLOCK) On CO.nConsValor = BS.nBSUnidad  And CO.nConsCOd = '1019'" _
                 & " Where nAlmCod = " & psAlmacen & " ORDER BY RTRIM(BS.cBSCod) "
        Else
            sqlA = " Select RTrim(BS.cBSCod) Codigo, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])', len(BS.cBSCod) Nivel " _
                 & " From BienesServicios BS  WITH (NOLOCK) " _
                 & " Inner Join AlmacenBS PBS WITH (NOLOCK) On BS.cBSCod like PBS.cBSCod + '%'" _
                 & " Inner Join Constante CO  WITH (NOLOCK) On CO.nConsValor = BS.nBSUnidad  And CO.nConsCOd = '1019'" _
                 & " Where nAlmCod = " & psAlmacen & " And Left(BS.cBSCod,2) In ('" & psSubProducto & "')" _
                 & " ORDER BY RTRIM(BS.cBSCod)"
        End If
    Else
        If psSubProducto = "" Then
            sqlA = " Select RTrim(BS.cBSCod) Codigo, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])', len(BS.cBSCod) Nivel " _
                 & " From BienesServicios BS WITH (NOLOCK) " _
                 & " Inner Join Constante CO WITH (NOLOCK) On CO.nConsValor = BS.nBSUnidad  And CO.nConsCOd = '1019' " _
                 & " ORDER BY RTRIM(BS.cBSCod)"
        Else
            If Not pnAddCtasIntangible Then
                If pbFilfroPedidoAgencia Then
                    sqlA = " Select RTrim(BS.cBSCod) Codigo, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])', len(BS.cBSCod) Nivel " _
                         & " From BienesServicios BS WITH (NOLOCK) " _
                         & " Inner Join Constante CO WITH (NOLOCK) On CO.nConsValor = BS.nBSUnidad  And CO.nConsCOd = '1019'" _
                         & " Inner Join LogPoderesObjeto LPO WITH (NOLOCK) On BS.cBSCod Like LPO.cObjetoCod + '%' And LPO.nTipo = 1" _
                         & " Where bVigente = 1" _
                         & " ORDER BY RTRIM(BS.cBSCod)"
                Else
                    sqlA = " Select RTrim(BS.cBSCod) Codigo, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])', len(BS.cBSCod) Nivel " _
                         & " From BienesServicios BS WITH (NOLOCK) " _
                         & " Inner Join Constante CO WITH (NOLOCK) On CO.nConsValor = BS.nBSUnidad  And CO.nConsCOd = '1019'" _
                         & " Where Substring(BS.cBSCod,2,2) In ('" & psSubProducto & "') And bVigente = 1" _
                         & " ORDER BY RTRIM(BS.cBSCod)"
                End If
            Else
                sqlA = " Select RTrim(BS.cBSCod) Codigo, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])' Descripcion, len(BS.cBSCod) Nivel " _
                     & " From BienesServicios BS WITH (NOLOCK) " _
                     & " Inner Join Constante CO WITH (NOLOCK) On CO.nConsValor = BS.nBSUnidad  And CO.nConsCOd = '1019'" _
                     & " Where Substring(BS.cBSCod,2,2) In ('" & psSubProducto & "') And bVigente = 1" _
                     & " UNION" _
                     & " Select C.cCtaContCod Codigo, C.cCtaContDesc Descripcion, 3 Nivel From CtaCont C WITH (NOLOCK) " _
                     & " Inner Join CtasIntangibles CI WITH (NOLOCK) On C.cCtaContCod = CI.cCtaContCod" _
                     & " ORDER BY Codigo"
            End If
        End If
    End If
    Set GetBienesAlmacen = oCon.CargaRecordSet(sqlA)
    
    Set oCon = Nothing
End Function

Public Function GetBNDBienes() As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    sqlA = " Select Distinct RTrim(BS.cBSCod) Codigo, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])', len(BS.cBSCod) Nivel " _
         & " From BienesServicios BS" _
         & " Inner Join Constante Co On CO.nConsValor = BS.nBSUnidad  And CO.nConsCOd = '1019'" _
         & " Inner Join BSNoDepreciables BSAF On BSAF.cBSCod = BS.cBSCod Where BSAF.bBaja = 0" _
         & " ORDER BY RTRIM(BS.cBSCod) "

    Set GetBNDBienes = oCon.CargaRecordSet(sqlA)
    
    Set oCon = Nothing
End Function



Public Function AFActualizaAreaAge(psArea As String, psAge As String, pnAnio As Long, psBSCod As String, psSerie As String) As Boolean
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    'sqlA = " Update BSActivoFijo " _
         & " Set cAreCod = '" & psArea & "', cAgeCod = '" & psAge & "'" _
         & " Where nAnio = " & pnAnio & " And cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
         
    '*** PEAC 20101129
    sqlA = " exec stp_upd_AFActualizaAreaAge '" & psArea & "','" & psAge & "'," & pnAnio & ",'" & psBSCod & "','" & psSerie & "'"
         
    oCon.Ejecutar sqlA
    AFActualizaAreaAge = 1
    
    Set oCon = Nothing
End Function

Public Function AFActualizaBaja(pnAnio As Long, psBSCod As String, psSerie As String, pdBaja As Date) As Boolean
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    sqlA = " Update BSActivoFijo " _
         & " Set bBaja = 1, dBaja = '" & Format(pdBaja, gsFormatoFecha) & "'" _
         & " Where nAnio = " & pnAnio & " And cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
    oCon.Ejecutar sqlA
    AFActualizaBaja = 1
    
    Set oCon = Nothing
End Function

Public Function BNDActualizaBaja(psBSCod As String, psSerie As String, pdBaja As Date) As Boolean
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    sqlA = " Update BSNoDepreciables " _
         & " Set bBaja = 1, dBaja = '" & Format(pdBaja, gsFormatoFecha) & "'" _
         & " Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
    oCon.Ejecutar sqlA
    BNDActualizaBaja = 1
    
    Set oCon = Nothing
End Function

Public Function GetAlmacenDetAF(pdIni As Date, pdFin As Date, Optional pbIngresos As Boolean = True) As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Dim sqlOrder As String
    
    Set oCon = New DConecta
    
    
    If pbIngresos Then
        sqlOrder = " BSControl"
    Else
        sqlOrder = " BSControlSal"
    End If
    
    oCon.AbreConexion
    sqlA = " Select cBSCod, cSerie, cDescripcion, Isnull(AREA.cAreaCod,'') cAreaCod,  IsNull(AREA.cAreadescripcion,'') cAreadescripcion,  IsNull(BSC.cAgeCod,'') cAgeCod ," _
         & " (Select cAgeDescripcion from Agencias AGE Where AGE.cAgeCod = Isnull(BSC.cAgeCod,(Select top 1 cUbicaCod from areaagencia where cAreaCod = BSC.cAreCod)))  Agencia, nBSValor, dActivacion," _
         & " (Select cPersNombre From Persona where cPersCod = BSC.cProveedor) Proveedor ," _
         & " cCtaCont , cFactura" _
         & " from " & sqlOrder & " BSC" _
         & " Left Join Areas AREA On BSC.cAreCod = AREA.cAreaCod" _
         & " Where dActivacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & " 23:59:59'" _
         & " Order by Left(cBSCod,3), cdescripcion, cSerie"
    Set GetAlmacenDetAF = oCon.CargaRecordSet(sqlA)
    
    Set oCon = Nothing
End Function

Public Function CierreMesLogistica(pdFecha As Date, Optional pnAlmacen As Long) As Boolean
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    
    sql = "Select Count(distinct dSaldo) From BSSaldos Where nAlmCod = " & pnAlmacen & " and dSaldo >= '" & Format(pdFecha, gsFormatoFecha) & "' "
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.Fields(0) = 0 Then
        CierreMesLogistica = False
    Else
        CierreMesLogistica = True
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    Set rs = Nothing
End Function

Public Function SetSerieMov(psBSCod As String, psDoc As String, pnTipoDoc As Integer, pnCantidad As Long, pnMovNro As Long, pnMovItem As Long, pgdFecSis As Date) As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lnI As Integer
    Dim lnContador As Long
    Dim lsSerie
    oCon.AbreConexion
    
    sqlA = "Select IsNull(Max(right(cSerie,8)),0) From BSSerie Where cBSCod = '" & psBSCod & "' And IsNumeric(Right(cSerie,8)) = 1"
    Set rs = oCon.CargaRecordSet(sqlA)
    
    lnContador = rs.Fields(0) + 1
    
    For lnI = 1 To pnCantidad
        lsSerie = Trim(Str(Year(pgdFecSis))) & "-" & Format(lnContador, "00000000")
        lnContador = lnContador + 1
        sqlA = "Insert BSSerie (cBSCod,cSerie,nBienEstado) values('" & psBSCod & "','" & lsSerie & "',1)"
        oCon.Ejecutar sqlA
        sqlA = "Insert MovBSSerie (nMovNro,nMovItem,cBSCod,cSerie) values(" & pnMovNro & "," & pnMovItem & ",'" & psBSCod & "','" & lsSerie & "')"
        oCon.Ejecutar sqlA
    Next lnI
    
    Set oCon = Nothing
End Function

Public Function AFActualizaPersona(psPersCod As String, pnAnio As Long, psBSCod As String, psSerie As String) As Boolean
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    sqlA = " Update BSActivoFijo " _
         & " Set cPersCod = '" & psPersCod & "'" _
         & " Where nAnio = " & pnAnio & " And cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
    oCon.Ejecutar sqlA
    AFActualizaPersona = 1
    
    Set oCon = Nothing
End Function

Public Function BNDActualizaPersona(psPersCod As String, psBSCod As String, psSerie As String) As Boolean
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    sqlA = " Update BSNoDepreciables " _
         & " Set cPersCod = '" & psPersCod & "'" _
         & " Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
    oCon.Ejecutar sqlA
    BNDActualizaPersona = 1
    
    Set oCon = Nothing
End Function


Public Sub GetPrecioActivo(psBSCod As String, psSerie As String, pnIGV As Currency, pnValor As Currency)
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    sqlA = " Select nIGV , nValor From BSSerie " _
          & " Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
    
    Set rs = oCon.CargaRecordSet(sqlA)
    
    If Not (rs.EOF And rs.BOF) Then
        pnIGV = IIf(IsNull(rs!nIGV), 0, rs!nIGV)
        pnValor = IIf(IsNull(rs!nValor), 0, rs!nValor)
    Else
        pnIGV = 0
        pnValor = 0
    End If
    Set oCon = Nothing
End Sub

Public Function TransferenciaActivoFijo(pAnio As Long, pdIni As Date, pdFin As Date) As Boolean
    Dim sqlA As String
    Dim sqlD As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    
    oCon.AbreConexion
    
    sqlD = " Delete BSActivoFijo" _
         & " Where dActivacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & "'" _
         & "       And nAnio = " & pAnio & " --And ltrim(rtrim(str(nMovNro)))+cBSCod+cSerie In" _
         & "           --(Select Ltrim(Rtrim(Str(nMovNro)))+cBSCod+cSerie" _
         & "           --         From BSControlSal BSCS" _
         & "           --         Where BSCS.cBsCod Like '112%' And dActivacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & "')"
        
    If IIf(oConst.LeeConstSistema(gConstSistBitIGVxActivoCredFiscal) = "1", True, False) Then
        sqlA = " Insert BSActivoFijo (nAnio, nMovNro, cBsCod, cSerie, nBSValor,nIGV, nBSDeprecia, bDepreciado, dActivacion, cAreCod, cAgeCod, nBSPerDeprecia, nBSPerDepAct, cDescripcion, bDeprecia, bIntangible, bForsado, cNroID, cNroPatrimonio, dCompra, dBase, Ban, bBaja, cPersCod)" _
             & " Select " & pAnio & " nAnio,nMovNro,cBSCod,cSerie,nBSValor,nIGV,nBSDeprecia,bDepreciado,dActivacion,cAreCod, isnull(cAgeCod,''),(Select nPorDeprecia from BienesServicios BS Where BSCS.cBSCod = BS.cBSCod) PerDep ,nBSPerDepAct,cDescripcion,bDeprecia,bIntangible,bForsado,cNroID,cNroPatrimonio,dActivacion,dActivacion,(Select nBSTipDep from BienesServicios BS Where BSCS.cBSCod = BS.cBSCod) TipoDep,0,Null From BSControlSal BSCS" _
             & " Where BSCS.cBsCod Like '112%' And dActivacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & "'"
    Else
        sqlA = " Insert BSActivoFijo (nAnio, nMovNro, cBsCod, cSerie, nBSValor, nIGV, nBSDeprecia, bDepreciado, dActivacion, cAreCod, cAgeCod, nBSPerDeprecia, nBSPerDepAct, cDescripcion, bDeprecia, bIntangible, bForsado, cNroID, cNroPatrimonio, dCompra, dBase, Ban, bBaja, cPersCod)" _
             & " Select " & pAnio & " nAnio,nMovNro,cBSCod,cSerie,nBSValor, 0, nBSDeprecia,bDepreciado,dActivacion,cAreCod, isnull(cAgeCod,''),(Select nPorDeprecia from BienesServicios BS Where BSCS.cBSCod = BS.cBSCod) PerDep ,nBSPerDepAct,cDescripcion,bDeprecia,bIntangible,bForsado,cNroID,cNroPatrimonio,dActivacion,dActivacion,(Select nBSTipDep from BienesServicios BS Where BSCS.cBSCod = BS.cBSCod) TipoDep,0,Null From BSControlSal BSCS" _
             & " Where BSCS.cBsCod Like '112%' And dActivacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & "'"
    End If
    
    oCon.BeginTrans
        oCon.Ejecutar sqlD
        oCon.Ejecutar sqlA
    oCon.CommitTrans
    
    Set oCon = Nothing
End Function

Public Function DatosAdjudicados(pnMovNro As Long, pnMovItem As Integer) As ADODB.Recordset
Dim sSQL As String
Dim oConn As DConecta

Set oConn = New DConecta

sSQL = "Select * from BSAdjudicados Where nMovNro = " & pnMovNro & " and nMovItem = " & pnMovItem & " "
If oConn.AbreConexion Then
   Set DatosAdjudicados = oConn.CargaRecordSet(sSQL)
Else
   Set DatosAdjudicados = Nothing
End If
End Function

Public Function AFActualizaBajaLote(pnAnio As Long, psBSCod As String, psSerie As String, pdBaja As Date, oConexion As ADODB.Connection) As Boolean
    Dim sqlA  As String
    
    sqlA = " Update BSActivoFijo " _
         & " Set bBaja = 1, dBaja = '" & Format(pdBaja, gsFormatoFecha) & "'" _
         & " Where nAnio = " & pnAnio & " And cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
    oConexion.Execute sqlA
    
    AFActualizaBajaLote = 1
End Function

Public Function GetBSObjSalida(psBSCod As String, psOpeCod As String) As ADODB.Recordset
    Dim sSQL As String
    Dim oConn As DConecta
    
    Set oConn = New DConecta
    
    sSQL = " Select Distinct cObjetoCod cBSCod, cObjetoDesc + ' [' + cCtaContCod  + ']' cBSDescripcion , nObjetoNiv from BSObj BSO" _
         & " inner Join Objeto O On BSO.cObj = O.cObjetoCod" _
         & " Where BSO.cOpeCod = '" & psOpeCod & "' And '" & psBSCod & "' Like BSO.cBSCod + '%'" _
         & " And Len(BSO.cBSCod) = ( Select Max(nBSCod) from BSObj BSO" _
         & "                         inner Join Objeto O On BSO.cObj = O.cObjetoCod" _
         & "                         Where BSO.cOpeCod = '" & psOpeCod & "' And '" & psBSCod & "' Like BSO.cBSCod + '%')"
    
    sSQL = sSQL & " Union Select Left('" & psBSCod & "',1) cBSCod, 'OBJETO DESTINO DE BIENES' cBSDescripcion , 0 nObjetoNiv Order By nObjetoNiv, cBSCod"
    If oConn.AbreConexion Then
       Set GetBSObjSalida = oConn.CargaRecordSet(sSQL)
    Else
       Set GetBSObjSalida = Nothing
    End If
End Function


Public Function GetKardexActivo(ByVal psBSCod As String, ByVal psSerie As String) As Recordset ' ByVal pdDeprecia As Date
Dim sSQL As String
Dim oCon As DConecta
Set oCon = New DConecta

oCon.AbreConexion

On Error GoTo GetKardexActivoError
sSQL = "SELECT M.cMovNro, M.cOpeCod, cOpeDesc ,nMovImporte, nBSValor ,MC.nMovImporte Depreciacion" & _
       " FROM BSActivoFijo BSAF " & _
       " Inner Join MovBSAF On BSAF.nAnio = MovBSAF.nAnio And BSAF.nMovNro = MovBSAF.nMovNro And BSAF.cBSCod = MovBSAF.cBSCod  And BSAF.cSerie = MovBSAF.cSerie " & _
       " Inner Join Mov M On M.nMovNro = MovBSAF.nMovNroRef " & _
       " Inner Join MovCta MC On M.nMovNro = MC.nMovNro And MC.nMovItem = MovBSAF.nMovItem " & _
       " Inner Join OpeTpo OP On OP.cOpeCod = M.cOpeCod " & _
       " where BSAF.cBSCod = '" & psBSCod & "' And BSAF.cSerie = '" & psSerie & "' " & _
       " And M.cOpeCod <> '581202'  And M.nMovFlag = 0 And M.nMovEstado = 25 " & _
       " Order By cMovNro"
      '***Estos campos del query van arriba falta definir la Fecha
   ' " Convert(Decimal(19,2),Round(((Round(dbo.GetFactorIPM (dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "') * nBSValor,10)) * (Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase,  Format(pdDeprecia, gsFormatoFecha)) Then BSAF.nBSPerDepAct + Datediff(Month,dBase,  Format(pdDeprecia, gsFormatoFecha)) Else BSAF.nBSPerDeprecia End)) / BSAF.nBSPerDeprecia,2) - dbo.GetLogDepAcum(BSAF.nAnio,BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581202',  '" & Format(pdDeprecia, gsFormatoFecha) & "')) DepAjusteMes,
       ' " Round(dbo.GetFactorIPM (dActivacion,  '" & Format(pdDeprecia, gsFormatoFecha) & "'),10) FactorAjuste,dactivacion
       
   Set GetKardexActivo = oCon.CargaRecordSet(sSQL)
   
    Exit Function
GetKardexActivoError:
    Call RaiseError(MyUnhandledError, "DLogAlmacen:GetKardexActivo Function")

End Function

Public Function GetLogMovBienes(ByVal pdFechaI As String, ByVal pdFechaF As String, ByVal pParametro As Integer, Optional ByVal pnArea As String = "", Optional ByVal pnAgencia As String = "", Optional ByVal psBSCod As String = "", Optional ByVal psSerie As String = "", Optional ByVal psOpeTpo As String = "", Optional ByVal pnTipo As Long = 0) As Recordset
'*** PEAC 20120426 psOpeTpo
Dim sSQL As String
Dim oCon As DConecta
Set oCon = New DConecta

oCon.AbreConexion

On Error GoTo GetLogMovBienesError

'Select Case pParametro
'Case 0
'sSQL = " Select substring(m.cMovNro,1,8) Fecha,f.nAnio,f.cBSCod,f.cSerie, f.cDescripcion Descripcion, " & _
'       " (Select cAgeDescripcion From Agencias " & _
'       " where cAgeCod = (Select cAgeCod FROM MovObjAreaAgencia " & _
'       " where nMovNro = m.nMovNro and nMovItem =1 )) AgenciaOrigen, " & _
'       " (Select cAreaDescripcion From Areas " & _
'       " where cAreaCod = (Select cAreaCod FROM MovObjAreaAgencia " & _
'       " where nMovNro = m.nMovNro and nMovItem =1 )) AreaOrigen, " & _
'       " (Select cAgeDescripcion From Agencias " & _
'       " where cAgeCod = (Select cAgeCod FROM MovObjAreaAgencia " & _
'       " where nMovNro = m.nMovNro and nMovItem =2 )) AgenciaDestino," & _
'       " (Select cAreaDescripcion From Areas " & _
'       " where cAreaCod = (Select cAreaCod FROM MovObjAreaAgencia " & _
'       " where nMovNro = m.nMovNro and nMovItem =2 )) Areadestino " & _
'       " From MOV m JOIN MOVBSAF mb ON  mb.nMovNroRef = m.nMovNro " & _
'       " JOIN BSActivoFijo f ON f.nAnio =mb.nAnio and f.cBSCod =mb.cBSCod and f.cSerie =mb.cSerie " & _
'       " Where mb.nMovItem = 1 And m.cOpecod = 581204 And f.bBaja = 0 " & _
'       " and f.cBSCod = '" & psBSCod & "' and f.cSerie ='" & psSerie & "' " & _
'       " and substring(m.cMovNro,1,8) >= '" & pdFechaI & "' and substring(m.cMovNro,1,8) < = '" & pdFechaF & "' "
'
'Case 1
'    sSQL = " Select substring(m.cMovNro,1,8) Fecha,f.nAnio,f.cBSCod,f.cSerie, f.cDescripcion Descripcion, " & _
'       " (Select cAgeDescripcion From Agencias " & _
'       " where cAgeCod = (Select cAgeCod FROM MovObjAreaAgencia " & _
'       " where nMovNro = m.nMovNro and nMovItem =1 )) AgenciaOrigen, " & _
'       " (Select cAreaDescripcion From Areas " & _
'       " where cAreaCod = (Select cAreaCod FROM MovObjAreaAgencia " & _
'       " where nMovNro = m.nMovNro and nMovItem =1 )) AreaOrigen, " & _
'       " (Select cAgeDescripcion From Agencias " & _
'       " where cAgeCod = (Select cAgeCod FROM MovObjAreaAgencia " & _
'       " where nMovNro = m.nMovNro and nMovItem =2 )) AgenciaDestino," & _
'       " (Select cAreaDescripcion From Areas " & _
'       " where cAreaCod = (Select cAreaCod FROM MovObjAreaAgencia " & _
'       " where nMovNro = m.nMovNro and nMovItem =2 )) Areadestino " & _
'       " From MOV m JOIN MOVBSAF mb ON  mb.nMovNroRef = m.nMovNro " & _
'       " JOIN BSActivoFijo f ON f.nAnio =mb.nAnio and f.cBSCod =mb.cBSCod and f.cSerie =mb.cSerie " & _
'       " JOIN MovObjAreaAgencia moa ON moa.nMovNro = m.nMovNro and moa.nMovItem =1 " & _
'       " Where mb.nMovItem = 1 And m.cOpecod = 581204 And f.bBaja = 0 " & _
'       " and substring(m.cMovNro,1,8) >= '" & pdFechaI & "' and substring(m.cMovNro,1,8) < = '" & pdFechaF & "' " & _
'       " and moa.cAreaCod= '" & pnArea & "' and moa.cAgeCod= '" & pnAgencia & "'"
'  End Select

sSQL = "exec stp_sel_GetLogMovBienes '" & pdFechaI & "','" & pdFechaF & "','" & psBSCod & "','" & psSerie & "','" & psOpeTpo & "'," & pnTipo

   Set GetLogMovBienes = oCon.CargaRecordSet(sSQL)

    Exit Function
GetLogMovBienesError:
    Call RaiseError(MyUnhandledError, "DLogAlmacen:GetLogMovBienes Function")

End Function

Public Function GetLogAlmacenSaldo(pdFecha As Date) As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
        sqlA = "spLogSaldosFinalCta '" & Format(pdFecha, gsFormatoFecha) & "'"
    'sqlA = " Select cBSCod, cBSDescripcion + ' ('  + cConsDescripcion + ')' Descripcion, 0, 0 From BienesServicios BS " _
    '     & "    Left Join Constante CO On CO.nConsCod = 1019 And CO.nConsValor = BS.nBSUnidad" _
    '     & " Where Len(cBSCod) >= 10 And cBSCod Not In (Select cBSCod From BSSaldos BSS Where dSaldo between '" & Format(pdFecha, gsFormatoFecha) & " 00:00:00' And '" & Format(pdFecha, gsFormatoFecha) & " 23:59:59' And nAlmCod = " & psAlmacen & " And nAlmTpo = " & pnProdTpo & ")" _
    '     & "    Union" _
    '     & " Select BS.cBSCod, cBSDescripcion + ' ('  + cConsDescripcion + ')' Descripcion, nStock, nMonto from BSSaldos BSS" _
    '     & "    Inner Join BienesServicios BS On BSS.cBSCod = BS.cBSCod" _
    '     & "    Left Join Constante CO On CO.nConsCod = 1019 And CO.nConsValor = BS.nBSUnidad" _
    '     & " Where dSaldo between '" & Format(pdFecha, gsFormatoFecha) & " 00:00:00' And '" & Format(pdFecha, gsFormatoFecha) & " 23:59:59' And nAlmCod = " & psAlmacen & "  And nAlmTpo = " & pnProdTpo

    Set GetLogAlmacenSaldo = oCon.CargaRecordSet(sqlA)
    
    Set oCon = Nothing
End Function
Public Function GetKardexActivoConsol(psFecha As String) As Recordset ' ByVal pdDeprecia As Date
Dim sSQL As String
Dim oCon As DConecta
Set oCon = New DConecta

oCon.AbreConexion

On Error GoTo GetKardexActivoConsolError
'ALPA 20081222************************************************************************
'sSql = "SELECT M.cMovNro, BSAF.cSerie, BSAF.cDescripcion, M.cOpeCod, cOpeDesc ,nMovImporte, nBSValor ,MC.nMovImporte Depreciacion," & _
'       "       BSAF.cAgeCod, A.cAgeDescripcion, BSAF.cAreCod, AR.cAreaDescripcion, BSAF.dActivacion, " & _
'       "       BSAF.nBSPerDeprecia " & _
'       " FROM BSActivoFijo BSAF " & _
'       " Inner Join MovBSAF On BSAF.nAnio = MovBSAF.nAnio And BSAF.nMovNro = MovBSAF.nMovNro And BSAF.cBSCod = MovBSAF.cBSCod  And BSAF.cSerie = MovBSAF.cSerie " & _
'       " Inner Join Mov M On M.nMovNro = MovBSAF.nMovNroRef " & _
'       " Inner Join MovCta MC On M.nMovNro = MC.nMovNro And MC.nMovItem = MovBSAF.nMovItem " & _
'       " Inner Join OpeTpo OP On OP.cOpeCod = M.cOpeCod " & _
'       " Inner Join Agencias A On BSAF.cAgeCod = A.cAgecod" & _
'       " Inner Join Areas AR On BSAF.cAreCod = AR.cAreaCod" & _
'       " where M.cOpeCod <> '581202'  And M.nMovFlag = 0 And M.nMovEstado = 25 " & _
'       "       And Substring(M.cMovNro,8) <= '" & psFecha & "'" & _
'       " Order By BSAF.cSerie, cMovNro"
'ALPA 20090504*******************************************************************************
sSQL = "SELECT DISTINCT SUBSTRING(M.cMovNro,1,8) cMovNro, BSAF.cSerie, BSAF.cDescripcion, M.cOpeCod, cOpeDesc ,nMovImporte, nBSValor ,MC.nMovImporte Depreciacion,"
'sSql = "SELECT M.cMovNro, BSAF.cSerie, BSAF.cDescripcion, M.cOpeCod, cOpeDesc ,nMovImporte, nBSValor ,MC.nMovImporte Depreciacion,"
'********************************************************************************************
sSQL = sSQL & "       BSAF.cAgeCod, A.cAgeDescripcion, BSAF.cAreCod, AR.cAreaDescripcion, BSAF.dActivacion, "
sSQL = sSQL & " BSAF.nBSPerDeprecia,BSAF.dCompra "
sSQL = sSQL & " FROM BSActivoFijo BSAF "
sSQL = sSQL & " Inner Join MovBSAF On BSAF.nMovNro = MovBSAF.nMovNro And BSAF.cBSCod = MovBSAF.cBSCod  And BSAF.cSerie = MovBSAF.cSerie "
'sSql = sSql & " Inner Join MovBSAF On BSAF.nAnio = MovBSAF.nAnio And BSAF.nMovNro = MovBSAF.nMovNro And BSAF.cBSCod = MovBSAF.cBSCod  And BSAF.cSerie = MovBSAF.cSerie "
sSQL = sSQL & " Inner Join Mov M On M.nMovNro = MovBSAF.nMovNroRef "
sSQL = sSQL & " Inner Join MovCta MC On M.nMovNro = MC.nMovNro And MC.nMovItem = MovBSAF.nMovItem "
sSQL = sSQL & " Inner Join OpeTpo OP On OP.cOpeCod = M.cOpeCod "
sSQL = sSQL & " Inner Join Agencias A On BSAF.cAgeCod = A.cAgecod"
sSQL = sSQL & " Inner Join Areas AR On BSAF.cAreCod = AR.cAreaCod"
sSQL = sSQL & " where M.cOpeCod <> '581202'  And M.nMovFlag = 0 And M.nMovEstado = 25 "
sSQL = sSQL & "    And Substring(M.cMovNro,1,8) <= '" & psFecha & "'"

sSQL = sSQL & "    Order By BSAF.cSerie, cMovNro"

'**********************************************************************************
   Set GetKardexActivoConsol = oCon.CargaRecordSet(sSQL)
   
    Exit Function
GetKardexActivoConsolError:
    Call RaiseError(MyUnhandledError, "DLogAlmacen:GetKardexActivo Function")

End Function

Public Function GetBienesActivosFijos() As ADODB.Recordset
    Dim sqlA  As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    'On Error GoTo GetBienesActivosFijosError
    sqlA = "Select Rtrim(BA.cSerie) Codigo, BA.cBSCod + ' '+ BS.cBSDescripcion + ' ' + BA.cDescripcion +'['+ BA.cSerie+ ']' as Descripcion, Len(BA.cSerie) As Nivel " _
         & " from BSActivoFijo BA " _
         & " Inner Join BienesServicios BS On BA.cBSCod = BS.cBSCod" _
         & " Where bBaja = 0 Order By BA.cBSCod"
    
    Set GetBienesActivosFijos = oCon.CargaRecordSet(sqlA)
  Set oCon = Nothing
    'Exit Function
'GetBienesActivosFijosError:
 '   Call RaiseError(MyUnhandledError, "DlogAlmacen:GetBienesActivosFijos Function")
    
End Function
Public Function GetAFBSSerieMod(psBSCod As String, pnAnio As Integer) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
    
   If oCon.AbreConexion() Then
      
      'sql = " Select cSerie, cDescripcion, convert(varchar(10),dCompra,103), nBSValor V_His  From BSActivoFijo BSAF" _
          & " Inner Join BienesServicios BS On BSAF.cBSCod = BS.cBSCod Where BS.cBSCod = '" & psBSCod & "' And nAnio = " & pnAnio & " And bBaja = 0 Order by dCompra, cSerie"
          
       '*** PEAC 20101129
        sql = " exec stp_sel_GetAFBSSerieMod '" & psBSCod & "'," & pnAnio
          
      Set GetAFBSSerieMod = oCon.CargaRecordSet(sql)
      
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Sub LimpiaSaldoBSSaldoCierreMes(ByVal psBSCod As String, ByVal psAlmacen As String, ByVal pnProdTpo As Integer, ByVal pdFecha As Date)
    On Error GoTo LimpiaSaldoBSSaldosCierreMesErr
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sql As String
    
    If oCon.AbreConexion() Then
        sql = "Exec stp_sel_LimpiaSaldoCierreAlmacen '" & psBSCod & "'," & psAlmacen & "," & pnProdTpo & ",'" & Format(pdFecha, gsFormatoFecha) & "'"
        oCon.Ejecutar (sql)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Sub
LimpiaSaldoBSSaldosCierreMesErr:
    Call RaiseError(MyUnhandledError, "DLogAlmacen:LimpiaSaldoBSSaldosCierreMes Method")
End Sub

'*** PEAC 20120327
Public Function BuscaNuevaSerie(ByVal psSerie As String, ByRef psDescrip As String) As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oConect As New DConecta
    On Error GoTo BuscarMovErr
    BuscaNuevaSerie = False
    
    If oConect.AbreConexion = False Then
        Exit Function
    End If
    sql = "exec stp_sel_BuscaNuevaSerie '" & psSerie & "'"
    Set rs = oConect.CargaRecordSet(sql)
    BuscaNuevaSerie = Not rs.EOF
    psDescrip = IIf(Not rs.EOF, rs!cDescripcion, "")
    rs.Close
    Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
BuscarMovErr:
    Call RaiseError(MyUnhandledError, "DLogAlmacen:BuscarNuevaserie ethod")

End Function


'*** PEAC 20120327
Public Function BuscaSiSerieFueDepre(ByVal psSerie As String) As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oConect As New DConecta
    On Error GoTo BuscarMovErr
    BuscaSiSerieFueDepre = False
    
    If oConect.AbreConexion = False Then
        Exit Function
    End If
    sql = "exec stp_sel_BuscaSiSerieFueDepre '" & psSerie & "'"
    Set rs = oConect.CargaRecordSet(sql)
    BuscaSiSerieFueDepre = Not rs.EOF
    rs.Close
    Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
BuscarMovErr:
    Call RaiseError(MyUnhandledError, "DLogAlmacen:BuscaSiSerieFueDepre Method")

End Function


