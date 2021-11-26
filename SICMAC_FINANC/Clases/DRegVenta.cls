VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRegVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9E031F038A"
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset
Public Function CargaRegistro(pnDocTpo As Long, psDocNro As String, dFecha1 As Date, dFecha2 As Date, Optional psAreaAgencia As String = "") As Recordset
Dim sCond As String
   On Error GoTo CargaRegistroErr
   Set dbConec = New DConecta
'   If pnDocTpo > 0 Then
'      sCond = " and rv.nDocTpo = " & pnDocTpo
'   End If
'   If psDocNro <> "" Then
'      'sCond = sCond & " and rv.cDocNro = '" & psDocNro & "'" Comentado by NAGL ERS012-2017
'       sCond = sCond & " and rv.cDocNro = " & " IIF('" & psDocNro & "' IN (select cDocNro from RegVenta where cDocNro = '" & psDocNro & "'),'" & psDocNro & "',Substring('" & psDocNro & "',2,4) + Substring('" & psDocNro & "',5,Len('" & psDocNro & "'))) " & ""
'       '******NAGL ERS012-2017 20170710 ********'
'   End If
   If dbConec.AbreConexion() Then
'      If dFecha1 = dFecha2 Then
''ALPA 20090925*************************************************
''         psSql = "SELECT rv.*, dbo.Pstanombre(p.cPersNombre,1) cPersNombre, d.cDocAbrev FROM RegVenta rv JOIN Documento d ON d.nDocTpo = rv.nDocTpo LEFT JOIN Persona p ON p.cPersCod = rv.cPersCod WHERE rv.dDocFecha = '" & Format(dFecha1, "mm/dd/yyyy hh:mm:ss") & "' " & sCond & " ORDER BY  convert(varchar(8),dDocFecha,112), rv.nDocTpo, cDocNro "
'          psSql = "         SELECT cDocNroNew = (case when Substring(rv.cDocNro,1,4) = (SELECT Right(cComprobanteNro,4) FROM ColocPigNroComprobanteAdj where Right(cComprobanteNro,4) =  Substring(rv.cDocNro,1,4) and cAgeCod = Right(rv.cAreaAgeCod,2)) Then rv.cDocNro  Else  '0' +  Substring(rv.cDocNro,1,3) + Substring(rv.cDocNro,5,Len(rv.cDocNro)) End)," 'NAGL ERS 012-2017
'          psSql = psSql & " rv.*, dbo.Pstanombre(p.cPersNombre,1) cPersNombre,d.cDocAbrev "
'          psSql = psSql & " FROM RegVenta rv JOIN Documento d ON d.nDocTpo = rv.nDocTpo "
'          psSql = psSql & "     LEFT JOIN Persona p ON p.cPersCod = rv.cPersCod "
'          psSql = psSql & " WHERE rv.dDocFecha = '" & Format(dFecha1, "mm/dd/yyyy hh:mm:ss") & "' " & sCond & " ORDER BY  convert(varchar(8),dDocFecha,112), rv.nDocTpo, cDocNro "
''**************************************************************
'      Else
''ALPA 20090925*************************************************
''      psSql = "SELECT rv.*, dbo.Pstanombre(p.cPersNombre,1) cPersNombre , "
''      psSql = psSql & " IsNull((Select cPersIDnro From PersID PID Where cPersIDTpo = 2 And PID.cPersCod = p.cPersCod),isnull((Select cPersIDnro From PersID PID Where cPersIDTpo = 1 And PID.cPersCod = p.cPersCod),'')) cRuc,"
''      psSql = psSql & " d.cDocAbrev FROM RegVenta rv JOIN Documento d ON d.nDocTpo = rv.nDocTpo LEFT JOIN Persona p ON p.cPersCod = rv.cPersCod "
''      psSql = psSql & " WHERE convert(datetime,convert(varchar(10),rv.dDocFecha,101)) BETWEEN '" & Format(dFecha1, gsFormatoFecha) & "' and '" & Format(dFecha2, gsFormatoFecha) & "' " & sCond & " ORDER BY  convert(varchar(8),dDocFecha,112), rv.nDocTpo, cDocNro "
'      psSql = "         SELECT cDocNroNew = (case when Substring(rv.cDocNro,1,4) = (SELECT Right(cComprobanteNro,4) FROM ColocPigNroComprobanteAdj where Right(cComprobanteNro,4) =  Substring(rv.cDocNro,1,4) and cAgeCod = Right(rv.cAreaAgeCod,2)) Then rv.cDocNro  Else  '0' +  Substring(rv.cDocNro,1,3) + Substring(rv.cDocNro,5,Len(rv.cDocNro)) End)," 'NAGL ERS 012-2017
'      psSql = psSql & " rv.*, dbo.Pstanombre(p.cPersNombre,1) cPersNombre,d.cDocAbrev, "
'      'psSql = psSql & " isnull((Select top 1 cPersIDnro From PersID PID Where PID.cPersCod = p.cPersCod order by cPersIDTpo DESC),'') cRuc," /**Comments PASI20170421**/
'      psSql = psSql & " isnull((Select top 1 cPersIDnro = case when cPersIDTpo = '99' then '00' else cPersIDnro end From PersID PID Where PID.cPersCod = p.cPersCod order by cPersIDTpo DESC),'') cRuc," 'PASI20170421
'      'psSql = psSql & " isnull((Select top 1 cPersIDTpo From PersID PID Where PID.cPersCod = p.cPersCod  order by cPersIDTpo DESC),'') cTipoDoc," /**Comments PASI20170421**/
'      psSql = psSql & " isnull((Select top 1 cPersIDTpo = case when cPersIDTpo = '99' then '00' else cPersIDTpo end From PersID PID Where PID.cPersCod = p.cPersCod  order by cPersIDTpo DESC),'') cTipoDoc," 'PASI20170421
'      psSql = psSql & " d.cDocAbrev FROM RegVenta rv JOIN Documento d ON d.nDocTpo = rv.nDocTpo LEFT JOIN Persona p ON p.cPersCod = rv.cPersCod "
'    'Modificado YIHU20150220
'      psSql = psSql & " WHERE convert(datetime,convert(varchar(10),rv.dDocFecha,101)) BETWEEN '" & Format(dFecha1, gsFormatoFecha) & "' and '" & Format(dFecha2, gsFormatoFecha) & "' " & sCond & " ORDER BY  convert(varchar(8),dDocFecha,112), rv.nDocTpo, cDocNro "
'      'psSql = psSql & " WHERE convert(datetime,convert(varchar(10),rv.dDocFecha,101)) BETWEEN '" & Format(dFecha1, "mm/dd/yyyy hh:mm:ss") & "' and '" & Format(dFecha2, "mm/dd/yyyy hh:mm:ss") & "' " & sCond & " ORDER BY  convert(varchar(8),dDocFecha,112), rv.nDocTpo, cDocNro "
'      'END YIHU
'**************************************************************
'       End If

      psSql = "Exec stp_sel_CargaRegistroRegVenta " & pnDocTpo & ", '" & psDocNro & "', '" & Format(dFecha1, "mm/dd/yyyy hh:mm:ss") & "','" & Format(dFecha2, "mm/dd/yyyy hh:mm:ss") & "','" & psAreaAgencia & "' " '********NAGL  ERS 012-2017
      Set CargaRegistro = dbConec.CargaRecordSet(psSql, adLockReadOnly)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaRegistroErr:
   Call RaiseError(MyUnhandledError, "DRegVenta:CargaRegistro Method")
End Function

'ALPA 20090925****************************************************
'Public Sub InsertaVenta(psOpeTpo As String, pnDocTpo As Long, psDocNro As String, pdDocFecha As Date, psPersCod As String, psCtaCod As String, psDescrip As String, pnVVenta As Currency, pnIGV As Currency, pnPVentas As Currency)
'*****************************************************************
Public Sub InsertaVenta(psOpeTpo As String, pnDocTpo As Long, psDocNro As String, pdDocFecha As Date, psPersCod As String, psCtaCod As String, psDescrip As String, pnVVenta As Currency, pnIGV As Currency, pnPVentas As Currency, psDocNroRefe As String, pdDocFecRef As Date, ByVal pnDocTpoRefe As Integer, _
                        ByVal pnMovNro As Long, pnMoneda As Integer, Optional pnTipoCambio As Currency = 0, Optional psAreaAgecod As String = "")
'YIHU20150218-ERS-181-2014, pnTipoCambio, pnMoneda
'PASI20161007 ERS0532016 psAreaAgeCod
   On Error GoTo InsertaVentaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then

'
''ALPA 20090925**************************************************
''      psSql = "INSERT RegVenta (cOpeTpo,nDocTpo,cDocNro, dDocFecha, cPersCod, cCtaCod, cDescrip, nVVenta, nIGV, nPVenta) " _
''            & "VALUES   ('" & psOpeTpo & "','" & pnDocTpo & "','" & psDocNro & "','" & Format(pdDocFecha, "mm/dd/yyyy hh:mm:ss") & "', '" _
''            & psPersCod & "','" & psCtaCod & "','" & psDescrip & "'," & pnVVenta & "," & pnIGV & "," & pnPVentas & ")"
'    psSql = "INSERT RegVenta (cOpeTpo,nDocTpo,cDocNro, dDocFecha, cPersCod, cCtaCod, cDescrip, nVVenta, nIGV, nPVenta,cDocNroRefe, dDocRefeFec,nDocTpoRefe,nMovNro, nTipoCambio, nMoneda,cAreaAgeCod) "
'    psSql = psSql & "VALUES   ('" & psOpeTpo & "','" & pnDocTpo & "','" & psDocNro & "','" & Format(pdDocFecha, "mm/dd/yyyy hh:mm:ss") & "', '"
'    psSql = psSql & psPersCod & "','" & psCtaCod & "','" & psDescrip & "'," & pnVVenta & "," & pnIGV & "," & pnPVentas & ",'" & psDocNroRefe & "', '" & Format(pdDocFecRef, "mm/dd/yyyy hh:mm:ss") & "'," & pnDocTpoRefe & "," & pnMovNro & "," & pnTipoCambio & "," & pnMoneda & ",'" & psAreaAgecod & "')"
''*****************************************************************
'YIHU20150218-ERS-181-2014, pnTipoCambio, pnMoneda
      psSql = "Exec stp_ins_RegistroVenta '" & psOpeTpo & "','" & pnDocTpo & "','" & psDocNro & "','" & Format(pdDocFecha, "mm/dd/yyyy hh:mm:ss") & "','" & psPersCod & "','" & psCtaCod & "','" & psDescrip & "'," & pnVVenta & "," & pnIGV & "," & pnPVentas & ",'" & psDocNroRefe & "', '" & Format(pdDocFecRef, "mm/dd/yyyy hh:mm:ss") & "'," & pnDocTpoRefe & "," & pnMovNro & "," & pnTipoCambio & "," & pnMoneda & ",'" & psAreaAgecod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaVentaErr:
   Call RaiseError(MyUnhandledError, "DRegVenta:InsertaVenta Method")
End Sub

'Public Sub ActualizaVenta(psOpeTpo As String, pnDocTpo As Long, psDocNro As String, pdDocFecha As Date, psPersCod As String, psCtaCod As String, psDescrip As String, pnVVenta As Currency, pnIGV As Currency, pnPVentas As Currency, pnDocTpoAnt As Long, psDocNroAnt As String, psDocFechaAnt As String)
Public Sub ActualizaVenta(psOpeTpo As String, pnDocTpo As Long, psDocNro As String, pdDocFecha As Date, psPersCod As String, psCtaCod As String, psDescrip As String, pnVVenta As Currency, pnIGV As Currency, pnPVentas As Currency, pnDocTpoAnt As Long, psDocNroAnt As String, psDocFechaAnt As String, psDocNroRefe As String, _
                        pdDocFecRef As Date, pnDocTpoRefe As Integer, Optional pnTipoCambio As Currency = 0)

'YIHU20150218-ERS181-2014, pnTipoCambio
On Error GoTo ActualizaVentaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
    'GIPO 02/12/2016
      psSql = "Exec sp_upd_ActualizarRegVenta '" & psOpeTpo & "', "
      psSql = psSql & pnDocTpo & ","
      psSql = psSql & "'" & psDocNro & "',"
      psSql = psSql & "'" & Format(pdDocFecha, "mm/dd/yyyy hh:mm:ss") & "', "
      psSql = psSql & "'" & psPersCod & "', "
      psSql = psSql & "'" & psCtaCod & "', "
      psSql = psSql & "'" & psDescrip & "', "
      psSql = psSql & pnVVenta & "," & pnIGV & "," & pnPVentas & ","
      psSql = psSql & "'" & psDocNroRefe & "', "
      psSql = psSql & "'" & Format(pdDocFecRef, "mm/dd/yyyy hh:mm:ss") & "',"
      psSql = psSql & pnDocTpoRefe & ", "
      psSql = psSql & pnTipoCambio & ", "
      psSql = psSql & pnDocTpoAnt & ", "
      psSql = psSql & "'" & psDocNroAnt & "', "
      psSql = psSql & "'" & Format(psDocFechaAnt, "mm/dd/yyyy hh:mm:ss") & "'  "
      'psSql = psSql & " WHERE nDocTpo = '" & pnDocTpoAnt & "' and cDocNro = '" & psDocNroAnt & "' and dDocFecha = '" & psDocFechaAnt & "'"
'*******************************************************************
'YIHU20150218 ERS 181, pnTipoCambio. Se agregó el campo TipoCambio
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizaVentaErr:
   Call RaiseError(MyUnhandledError, "DRegVenta:ActualizaVenta Method")
End Sub

Public Function CargaSerieValida(psDocNroSerie As String, psAgencia As String, psAreaAgen As String, psTipoAccion As String) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    On Error GoTo ErroCargaSerieValida
    
    psSql = "Exec sp_sel_CargaSerieValida '" & psDocNroSerie & "','" & psAgencia & "', '" & psAreaAgen & "','" & psTipoAccion & "'"
    oConec.AbreConexion
    Set CargaSerieValida = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    
    Exit Function
ErroCargaSerieValida:
    Call RaiseError(MyUnhandledError, "Carga Serie Valida")
End Function 'NAGL 20170801

Public Sub EliminaVenta(pnDocTpo As Long, psDocNro As String, pdDocFecha As Date, pOpeTpo As String, Optional ByVal psMovNro As String = "")
   On Error GoTo EliminaVentaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      
      '*** PEAC 20110425
      'psSql = "DELETE RegVenta WHERE nDocTpo = '" & pnDocTpo & "' and cDocNro = '" & psDocNro & "' and dDocFecha = '" & Format(pdDocFecha, "mm/dd/yyyy hh:mm:ss") & "'"
      
      psSql = "exec stp_sel_EliminaVenta '" & pnDocTpo & "','" & psDocNro & "','" & Format(pdDocFecha, "yyyymmdd") & "','" & pOpeTpo & "','" & psMovNro & "'"
      
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaVentaErr:
   Call RaiseError(MyUnhandledError, "DRegVenta:EliminaVenta Method")
End Sub

Public Function CargaRegOperacion(Optional psOpeTpo As String = "") As Recordset
   On Error GoTo CargaRegistroErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "SELECT cOpeTpo, cOpeDesc FROM RegVentaOpe " & IIf(psOpeTpo <> "", " WHERE cOpeTpo = '" & psOpeTpo & "'", "")
      Set CargaRegOperacion = dbConec.CargaRecordSet(psSql, adLockReadOnly)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaRegistroErr:
   Call RaiseError(MyUnhandledError, "DRegVenta:CargaRegistro Method")
End Function

Public Sub MigrarIngresosCustodia()
   On Error GoTo MigrarIngresosCustodiaErr

    'your code goes here...

   Exit Sub
MigrarIngresosCustodiaErr:
   Call RaiseError(MyUnhandledError, "DRegVenta:MigrarIngresosCustodia Method")
End Sub

Public Sub MigrarRemateJoyas()
   On Error GoTo MigrarRemateJoyasErr

    'your code goes here...

   Exit Sub
MigrarRemateJoyasErr:
   Call RaiseError(MyUnhandledError, "DRegVenta:MigrarRemateJoyas Method")
End Sub

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

End Sub



Public Function CargaCuarta(psAnio As String, psMes As String) As Recordset
Dim sql As String
   On Error GoTo CargaCuartaErr
   Set dbConec = New DConecta
  
  
   If dbConec.AbreConexion() Then
    '**Modificado por ALPA 20080214 ***********************
      'sql = "SP_FinRentaCuarta '" & psAnio & "', '" & psMes & "'  "
      sql = "stp_sel_FinRentaCuarta '" & psAnio & "', '" & psMes & "'  "
      '****************************************************
      
      
      Set CargaCuarta = dbConec.CargaRecordSet(sql)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaCuartaErr:
   Call RaiseError(MyUnhandledError, "DRegVenta:CargaCuarta Method")
End Function

'*** PEAC 20091116
Public Function CargaEntidades() As Recordset
   On Error GoTo ErrorCargaEntidades
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      
      psSql = " exec stp_sel_ObtieneEntidadesFinancieras "
      
      Set CargaEntidades = dbConec.CargaRecordSet(psSql, adLockReadOnly)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
ErrorCargaEntidades:
   Call RaiseError(MyUnhandledError, "DRegVenta : CargaEntidades")
End Function

'*** PEAC 20110425
Public Function VerificaVentaExistente(pnDocTpo As Long, psDocNro As String, pdDocFecha As Date, pOpeTpo As String) As Recordset
   On Error GoTo VerificaVentaExistenteErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then

    psSql = "exec stp_sel_VerificaVentaExistente '" & pnDocTpo & "','" & psDocNro & "','" & Format(pdDocFecha, "yyyymmdd") & "','" & pOpeTpo & "'"
      Set VerificaVentaExistente = dbConec.CargaRecordSet(psSql, adLockReadOnly)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
VerificaVentaExistenteErr:
   Call RaiseError(MyUnhandledError, "DRegVenta:VerificaVentaExistente Method")
End Function
'PASI20161006 ERS0532016
Public Function DocAutDevuelveNroComprobante(ByVal psAgeCod As String) As ADODB.Recordset
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        Set DocAutDevuelveNroComprobante = dbConec.CargaRecordSet("exec stp_sel_DocAutDevuelveNroComprobante '" & psAgeCod & "'", adLockReadOnly)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
End Function
Public Function ListaDocAutorizado(ByVal pdFechaIni As Date, ByVal pdFechaDFin As Date, ByVal psAreaAgecod As String) As ADODB.Recordset
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        Set ListaDocAutorizado = dbConec.CargaRecordSet("exec stp_sel_ListaDocAutorizado '" & Format(pdFechaIni, "YYYYMMdd") & "','" & Format(pdFechaDFin, "YYYYMMdd") & "','" & psAreaAgecod & "'", adLockReadOnly)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
End Function
Public Function ObtieneDocAutorizadoxNroDoc(ByVal psDocNro As String) As ADODB.Recordset
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        Set ObtieneDocAutorizadoxNroDoc = dbConec.CargaRecordSet("exec stp_sel_ObtieneDocAutorizadoxNroDoc '" & psDocNro & "'", adLockReadOnly)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
End Function
Public Function ObtieneDatosAgencia(ByVal psAgeCod As String) As ADODB.Recordset
Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        Set ObtieneDatosAgencia = dbConec.CargaRecordSet("stp_sel_ObtieneDatosAgencia '" & psAgeCod & "'", adLockReadOnly)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
End Function
Public Sub ActualizaDocAutorizado(psOpeTpo As String, pnDocTpo As Long, psDocNro As String, pdDocFecha As Date, psPersCod As String, psCtaCod As String, psDescrip As String, pnVVenta As Currency, pnIGV As Currency, pnPVentas As Currency, pnDocTpoAnt As Long, psDocNroAnt As String, psDocFechaAnt As String, psDocNroRefe As String, _
                        pdDocFecRef As Date, pnDocTpoRefe As Integer, pnMoneda As Integer, pnTipoCambio As Currency, psAreaAgecod As String)
On Error GoTo ActualizaDocErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      '*****************ERS 012-2017 20170710**********************
      psSql = " Exec stp_upd_ActualizaDocAutorizado '" & psOpeTpo & "', '" & pnDocTpo & "', '" & psDocNro & "', '" & Format(pdDocFecha, "mm/dd/yyyy hh:mm:ss") & "', '" & psPersCod & "', '" & psCtaCod & "', '" & psDescrip & "', " & pnVVenta & ", " & pnIGV & ", " & pnPVentas & ", "
      psSql = psSql & " '" & psDocNroRefe & "', '" & Format(pdDocFecRef, "mm/dd/yyyy hh:mm:ss") & "'," & pnDocTpoRefe & ", " & pnTipoCambio & ", '" & psAreaAgecod & "', " & pnMoneda & ", "
      psSql = psSql & " '" & pnDocTpoAnt & "', '" & psDocNroAnt & "', '" & psDocFechaAnt & "'"
      
      'psSql = "UPDATE RegVenta SET cOpeTpo = '" & psOpeTpo & "', "
      'psSql = psSql & "nDocTpo = '" & pnDocTpo & "',"
      'psSql = psSql & "cDocNro = '" & psDocNro & "',"
      'psSql = psSql & "dDocFecha = '" & Format(pdDocFecha, "mm/dd/yyyy hh:mm:ss") & "', "
      'psSql = psSql & "cPersCod = '" & psPersCod & "', "
      'psSql = psSql & "cCtaCod = '" & psCtaCod & "', "
      'psSql = psSql & "cDescrip = '" & psDescrip & "', "
      'psSql = psSql & "nVVenta = " & pnVVenta & ", nIGV = " & pnIGV & ", nPVenta = " & pnPVentas & ","
      'psSql = psSql & "cDocNroRefe = '" & psDocNroRefe & "', "
      'psSql = psSql & "dDocRefeFec = '" & Format(pdDocFecRef, "mm/dd/yyyy hh:mm:ss") & "', "
      'psSql = psSql & "nDocTpoRefe = " & pnDocTpoRefe & ", "
      'psSql = psSql & "nTipoCambio = " & pnTipoCambio & ", "
      'psSql = psSql & "cAreaAgeCod = '" & psAreaAgecod & "',"
      'psSql = psSql & "nMoneda = " & pnMoneda
      'psSql = psSql & " WHERE nDocTpo = '" & pnDocTpoAnt & "' and cDocNro = '" & psDocNroAnt & "' and dDocFecha = '" & psDocFechaAnt & "'"
      'COMENTADO BY NAGL ERS 012-2017
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizaDocErr:
   Call RaiseError(MyUnhandledError, "DRegVenta:ActualizaDocAutorizado Method")
End Sub
Public Function EsDocAutorizado(ByVal psDocNro As String) As Boolean
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        Set prs = dbConec.CargaRecordSet("stp_sel_EsDocAutorizado '" & psDocNro & "'", adLockReadOnly)
        If Not (prs.EOF And prs.BOF) Then EsDocAutorizado = IIf(Not (prs.EOF And prs.BOF), IIf(Not prs!nValor = 0, True, False), False)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
End Function
'END PASI


