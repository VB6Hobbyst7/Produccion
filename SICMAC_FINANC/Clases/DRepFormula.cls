VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRepFormula"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A7EDA0000BB"
Option Base 0
Option Explicit
Dim sCentralCom As String
Dim dbConec   As DConecta

Dim psSql As String

'EJVG20130416 ***
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
    dbConec.RollbackTrans
    dbConec.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub
Public Sub dCommitTrans()
    dbConec.CommitTrans
End Sub
'END EJVG *******
Public Function CargaRepFormula(Optional psCodigo As String = "", Optional psOpeCod As String = "", Optional plLike As Boolean = False, Optional LockType As LockTypeEnum = adLockReadOnly, Optional plOpeDesc As Boolean = False, Optional lsFormaOrdenRep As String = "INT") As Recordset
Dim sCond As String
   On Error GoTo CargaRepFormulaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sCond = ""
      sCond = IIf(psCodigo = "", "", " WHERE rc.cCodigo " & IIf(plLike, " LIKE '", "= '") & psCodigo & IIf(plLike, "%'", "'"))
      sCond = sCond & IIf(psOpeCod = "", "", IIf(sCond = "", " WHERE", " AND") & " rc.cOpeCod " & IIf(plLike, " LIKE '", "= '") & psOpeCod & IIf(plLike, "%'", "'"))
      If plOpeDesc Then
         psSql = "SELECT rc.cCodigo, rc.cOpeCod, Isnull(rc.cDescrip,'')cDescrip, ISNULL(rc.cFormula,'')cFormula , o.cOpeDesc, cObserva, bIngresoManual FROM RepBaseFormula rc LEFT JOIN " & sCentralCom & "OpeTpo o ON o.cOpeCod = rc.cOpeCod " & sCond & " ORDER BY rc.cOpeCod, " & IIf(lsFormaOrdenRep = "INT", "right(replicate('0',20)+RTRIM(rc.cCodigo),20) ", "rc.cCodigo")
      Else
         psSql = "SELECT cCodigo, cOpeCod, Isnull(rc.cDescrip,'')cDescrip, ISNULL(rc.cFormula,'')cFormula, cObserva, bIngresoManual  FROM RepBaseFormula rc " & sCond & " ORDER BY cOpeCod, " & IIf(lsFormaOrdenRep = "INT", "right(replicate('0',20)+RTRIM(rc.cCodigo),20) ", "rc.cCodigo")
      End If
      Set CargaRepFormula = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaRepFormulaErr:
   Call RaiseError(MyUnhandledError, "DMntRepFormula:CargaRepFormula Method")
End Function

Public Sub ReorganizaOrdenRep(psOpeCod As String, psCodigo As String)
   On Error GoTo ReorganizaOrdenRepErr
    
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "UPDATE RepBaseFormula Set cCodigo=CONVERT(VARCHAR(20),CONVERT(INT,cCodigo)) + 1 Where CONVERT(INT,cCodigo) >= " & Val(psCodigo) & " and cOpeCod = '" & psOpeCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing

   Exit Sub
ReorganizaOrdenRepErr:
   Call RaiseError(MyUnhandledError, "DMntRepFormula:ReorganizaOrdenRep Method")
End Sub

Public Sub InsertaRepFormula(psCodigo As String, psOpeCod As String, psDescrip As String, psFormula As String, psObserva As String, psMovNro As String, pnValor As Integer)
On Error GoTo InsertaRepFormulaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT INTO RepBaseFormula(cCodigo,cOpeCod, cDescrip, cFormula, cObserva, cUltimaActualizacion, bIngresoManual) " _
            & " VALUES('" & psCodigo & "','" & psOpeCod & "','" & psDescrip & "','" & psFormula & "','" & psObserva & "', '" & psMovNro & "', " & pnValor & ")"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaRepFormulaErr:
   Call RaiseError(MyUnhandledError, "DMntRepFormula:InsertaRepFormula Method")
End Sub

Public Sub ActualizaRepFormula(psCodigo As String, psOpeCod As String, psDescrip As String, psFormula As String, psObserva As String, psMovNro As String, pnValor As Integer)
   On Error GoTo ActualizaRepFormulaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "UPDATE RepBaseFormula Set cDescrip = '" & psDescrip & "', cFormula = '" & psFormula & "', cObserva='" & psObserva & "', cUltimaActualizacion = '" & psMovNro & "', bIngresoManual = " & pnValor & "  Where cCodigo = '" & psCodigo & "' and cOpeCod = '" & psOpeCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizaRepFormulaErr:
   Call RaiseError(MyUnhandledError, "DMntRepFormula:ActualizaRepFormula Method")
End Sub

Public Sub EliminaRepFormula(psCodigo As String, psOpeCod As String)
   On Error GoTo EliminaRepFormulaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE RepBaseFormula Where cCodigo = '" & psCodigo & "' and cOpeCod = '" & psOpeCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaRepFormulaErr:
   Call RaiseError(MyUnhandledError, "DMntRepFormula:EliminaRepFormula Method")
End Sub


Public Function ExisteCodigo(psOpeCod As String, psCodigo As String) As Boolean
Dim reg As ADODB.Recordset
Dim ban As Boolean

   On Error GoTo GeneraCodigoFilaErr
 
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "SELECT cCodigo, cOpeCod From dbo.RepBaseFormula WHERE (cCodigo='" & Trim(psCodigo) & "') AND (cOpeCod='" & psOpeCod & "')"
      Set reg = dbConec.CargaRecordSet(psSql)
      If reg.BOF Then
        ExisteCodigo = False
      Else
        ExisteCodigo = True
      End If
      reg.Close
      Set reg = Nothing
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
 
   
   Exit Function
GeneraCodigoFilaErr:
   Call RaiseError(MyUnhandledError, "DMntRepFormula:GeneraCodigoFila Method")
End Function
 


Public Function GeneraCodigoFila(psOpeCod As String) As String
   On Error GoTo GeneraCodigoFilaErr


   Exit Function
GeneraCodigoFilaErr:
   Call RaiseError(MyUnhandledError, "DMntRepFormula:GeneraCodigoFila Method")
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
   Dim oIni As New ClasIni
   sCentralCom = oIni.BaseComunes
    'EJVG20130416 ***
    Set dbConec = New DConecta
    If Not dbConec.AbreConexion Then
        Call RaiseError(MyUnhandledError, "DRepFormula:Initialize Method. Error en Conexion a Base de datos")
    End If
    'END EJVG ******
   Set oIni = Nothing
End Sub
'EJVG20130413 ***
Public Function RecuperaConfigRepNotasEstado(ByVal psOpeCod As String) As ADODB.Recordset
    On Error GoTo ErrRecuperaConfigRepNotaEstado
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "Exec NIIF_stp_sel_NotasEstadoConfig '" & psOpeCod & "'"
        Set RecuperaConfigRepNotasEstado = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrRecuperaConfigRepNotaEstado:
   Call RaiseError(MyUnhandledError, "DRepFormula:RecuperaConfigRepNotaEstado Method")
End Function
Public Function RecuperaConfigRepNotasEstadoDetalle(ByVal psOpeCod As String, ByVal pnId As Long, ByVal pnNotaEstado As Long) As ADODB.Recordset
    On Error GoTo ErrRecuperaConfigRepNotaEstadoDetalle
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "Exec NIIF_stp_sel_NotasEstadoConfigDet " & pnId & ",'" & psOpeCod & "'," & pnNotaEstado
        Set RecuperaConfigRepNotasEstadoDetalle = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrRecuperaConfigRepNotaEstadoDetalle:
   Call RaiseError(MyUnhandledError, "DRepFormula:RecuperaConfigRepNotasEstadoDetalle Method")
End Function
Public Function InsertarNotasEstadoConfig(ByVal psOpeCod As String, ByVal psMovNro As String) As Long
    On Error GoTo ErrInsertarNotasEstadoConfig
    Dim rs As New ADODB.Recordset
    
    InsertarNotasEstadoConfig = 0
    psSql = "Exec NIIF_stp_ins_NotasEstadoConfig '" & psOpeCod & "','" & psMovNro & "'"
    Set rs = dbConec.CargaRecordSet(psSql)
    If Not RSVacio(rs) Then
        InsertarNotasEstadoConfig = rs!nId
    End If
    Set rs = Nothing
    Exit Function
ErrInsertarNotasEstadoConfig:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertarNotasEstadoConfig Method")
End Function
Public Sub InsertarNotasEstadoConfigDet(ByVal pnId As Long, ByVal psOpeCod As String, ByVal pnNotaEstado As Long, ByVal psDescripcion As String, ByVal psFormula As String, ByVal pbPeriodo As Boolean, ByVal pbComentario As Boolean)
    On Error GoTo ErrInsertarNotasEstadoConfigDet
    psSql = "Exec NIIF_stp_ins_NotasEstadoConfigDet " & pnId & ",'" & psOpeCod & "'," & pnNotaEstado & ",'" & psDescripcion & "','" & psFormula & "'," & IIf(pbPeriodo, 1, 0) & "," & IIf(pbComentario, 1, 0)
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrInsertarNotasEstadoConfigDet:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertarNotasEstadoConfigDet Method")
End Sub
'Public Sub InsertarNotasEstadoConfigDetDet(ByVal pnID As Integer, ByVal psOpeCod As String, ByVal pnNotaEstado As Integer, ByVal pnOrden As Integer, ByVal pnTpoDetalle As Integer, ByVal psDescripcion As String, ByVal pbNegrita As Boolean, ByVal pnNivel As Integer, ByVal psFormula1 As String, ByVal psFormula2 As String, ByVal psFormula3 As String, ByVal psFormula4 As String, ByVal psFormula5 As String)
Public Sub InsertarNotasEstadoConfigDetDet(ByVal pnId As Integer, ByVal psOpeCod As String, ByVal pnNotaEstado As Integer, ByVal pnOrden As Integer, ByVal pnTpoDetalle As Integer, ByVal psDescripcion As String, ByVal pbNegrita As Boolean, ByVal pnNivel As Integer, ByVal psFormula1 As String, ByVal psFormula2 As String, ByVal psFormula3 As String, ByVal psFormula4 As String, ByVal psFormula5 As String, ByVal psFormula1_2012 As String, ByVal psFormula2_2012 As String, ByVal psFormula3_2012 As String, ByVal psFormula4_2012 As String, ByVal psFormula5_2012 As String)
    On Error GoTo ErrInsertarNotasEstadoConfigDet
    'psSql = "Exec NIIF_stp_ins_NotasEstadoConfigDetDet " & pnID & ",'" & psOpeCod & "'," & pnNotaEstado & "," & pnOrden & "," & pnTpoDetalle & ",'" & psDescripcion & "'," & IIf(pbNegrita, 1, 0) & "," & pnNivel & ",'" & psFormula1 & "','" & psFormula2 & "','" & psFormula3 & "','" & psFormula4 & "','" & psFormula5 & "'"
    psSql = "Exec NIIF_stp_ins_NotasEstadoConfigDetDet " & pnId & ",'" & psOpeCod & "'," & pnNotaEstado & "," & pnOrden & "," & pnTpoDetalle & ",'" & psDescripcion & "'," & IIf(pbNegrita, 1, 0) & "," & pnNivel & ",'" & psFormula1 & "','" & psFormula2 & "','" & psFormula3 & "','" & psFormula4 & "','" & psFormula5 & "','" & psFormula1_2012 & "','" & psFormula2_2012 & "','" & psFormula3_2012 & "','" & psFormula4_2012 & "','" & psFormula5_2012 & "'"
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrInsertarNotasEstadoConfigDet:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertarNotasEstadoConfigDet Method")
End Sub
Public Function ObtenerUltimoNivelConfig(ByVal psOpeCod As String) As Integer
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrObtenerUltimoNivelConfig
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "Exec NIIF_stp_sel_ObtieneUltimoNivelConfig '" & psOpeCod & "'"
        Set rs = dbConec.CargaRecordSet(psSql)
        If Not rs.EOF Then
            ObtenerUltimoNivelConfig = rs!nNivelMax
        End If
        dbConec.CierraConexion
    End If
    Set rs = Nothing
    Set dbConec = Nothing
    Exit Function
ErrObtenerUltimoNivelConfig:
   Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerUltimoNivelConfig Method")
End Function
'END EJVG *******
'ALPA 20130518****************************************************************************
Public Sub InsertarEEFF(ByVal psOpeCod As String, ByVal pnCorrelativo As Integer, ByVal pnNivel As Integer, ByVal psDescripcion As String, ByVal psFormulaCons As String, ByVal psFormulaAgen As String, ByVal pnTipo As Integer, ByVal psCodUser As String)
    On Error GoTo ErrInsertarEEFF
    Dim dbConecta As DConecta
    Set dbConecta = New DConecta
    dbConecta.AbreConexion
        psSql = "Exec NIIF_stp_ins_NIIFRepFor '" & psOpeCod & "','" & pnCorrelativo & "'," & pnNivel & ",'" & psDescripcion & "','" & psFormulaCons & "','" & psFormulaAgen & "'," & pnTipo & ",'" & psCodUser & "'"
        dbConecta.Ejecutar (psSql)
    dbConecta.CierraConexion
    Exit Sub
ErrInsertarEEFF:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertarNotasEstadoConfigDet Method")
End Sub

Public Sub EliminarEEFF(ByVal psOpeCod As String)
    On Error GoTo ErrEliminarEEFF
    Dim dbConecta As DConecta
    Set dbConecta = New DConecta
    dbConecta.AbreConexion
        psSql = "Exec NIIF_stp_del_NIIFRepFor '" & psOpeCod & "'"
        dbConecta.Ejecutar (psSql)
    dbConecta.CierraConexion
    Exit Sub
ErrEliminarEEFF:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertarNotasEstadoConfigDet Method")
End Sub
Public Function ObtenerEEFF(ByVal psOpeCod As String, Optional ByVal pnAnio As Integer = 2014, Optional ByVal pnMes As Integer = 12) As ADODB.Recordset
    On Error GoTo ErrEliminarEEFF
    Dim dbConecta As DConecta
    Set dbConecta = New DConecta
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    dbConecta.AbreConexion
        'psSql = "Exec NIIF_stp_sel_NIIFRepFor '" & psOpeCod & "'"
        psSql = "Exec NIIF_stp_sel_NIIFRepFor '" & psOpeCod & "'," & pnAnio & "," & pnMes 'EJVG20140912
        Set oRs = dbConecta.CargaRecordSet(psSql)
    dbConecta.CierraConexion
    Set ObtenerEEFF = oRs
    Set oRs = Nothing
    Exit Function
ErrEliminarEEFF:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerEEFF Method")
End Function
Public Function ObtenerEEFFxId(ByVal psOpeCod As String, ByVal pnId As Integer, ByVal psAge As String) As String
    On Error GoTo ErrEliminarEEFF
    Dim dbConecta As DConecta
    Set dbConecta = New DConecta
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    dbConecta.AbreConexion
    psSql = "Exec stp_sel_ObtenerNIIFRepForxnCorreint '" & psOpeCod & "'," & pnId & ""
    Set oRs = dbConecta.CargaRecordSet(psSql)
    ObtenerEEFFxId = ""
    If Not (oRs.BOF Or oRs.EOF) Then
        Do While Not oRs.EOF
            If psAge = "" Then
                ObtenerEEFFxId = oRs!cFormulaCons
            Else
                ObtenerEEFFxId = oRs!cFormulaAgen
            End If
            oRs.MoveNext
        Loop
    End If
    dbConecta.CierraConexion
    Set oRs = Nothing
    Exit Function
ErrEliminarEEFF:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerEEFF Method")
End Function
Public Function ObtenerDistCreditosDirectos(ByVal pdFecha As Date, ByVal psMoneda As String) As Currency
    On Error GoTo ErrObtenerDistCreditosDirectos
    Dim dbConecta As DConecta
    Set dbConecta = New DConecta
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Dim lnSaldo As Currency
    lnSaldo = 0
    dbConecta.AbreConexion
        psSql = "Exec NIIF_stp_sel_DistCreditosDirectos '" & Format(pdFecha, "YYYY/MM/DD") & "','" & psMoneda & "'"
    Set oRs = dbConecta.CargaRecordSet(psSql)
    If Not (oRs.BOF Or oRs.EOF) Then
        lnSaldo = IIf(IsNull(oRs!nSaldoFinal), 0, oRs!nSaldoFinal)
    End If
    Set oRs = Nothing
    dbConecta.CierraConexion
    ObtenerDistCreditosDirectos = lnSaldo
    Exit Function
ErrObtenerDistCreditosDirectos:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerEEFF Method")
End Function
'*****************************************************************************************
'COMENTADO PASI 20140315 *****************************************************************
'*** PEAC 20131125 -
'Public Function CargarRegimenesRepoFormulaBCR() As Recordset
'Dim sCond As String
'   On Error GoTo CargarRegimenesRepoFormulaBCRErr
'   Set dbConec = New DConecta
'   If dbConec.AbreConexion() Then
'        psSql = "exec stp_sel_CargarRegimenesRepoFormulaBCR"
'        Set CargarRegimenesRepoFormulaBCR = dbConec.CargaRecordSet(psSql)
'        dbConec.CierraConexion
'   End If
'   Set dbConec = Nothing
'   Exit Function
'CargarRegimenesRepoFormulaBCRErr:
'   Call RaiseError(MyUnhandledError, "DMntRepFormula:CargarRegimenesRepoFormulaBCR Method")
'End Function

'*** PEAC 20131126
'Public Sub InsertaRep1FormulaBCR(ByVal psNumCol As String, ByVal psDescripcion As String, ByVal cCodigo As String, ByVal pnCodRegimen As Integer, ByVal psValor As String, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal psMovNro As String, ByVal pnValor As Integer, ByVal pnNumColRef As Integer)
'On Error GoTo InsertaRep1FormulaBCRErr
'   Set dbConec = New DConecta
'   If dbConec.AbreConexion() Then
'
'      psSql = "exec stp_ins_InsertaRep1FormulaBCR '" & psNumCol & "','" & psDescripcion & "','" & cCodigo & "'," & pnCodRegimen & ",'" & psValor & "'," & pnItem & "," & pnMoneda & ",'" & psMovNro & "'," & pnValor & "," & pnNumColRef
'
'      dbConec.Ejecutar psSql
'      dbConec.CierraConexion
'   End If
'   Set dbConec = Nothing
'   Exit Sub
'InsertaRep1FormulaBCRErr:
'   Call RaiseError(MyUnhandledError, "DMntRepFormula:InsertaRep1FormulaBCR Method")
'End Sub

'*** PEAC 20131127
'Public Sub EliminaRep1FormulaBCR(psNumCol As String, psMoneda As String)
'   On Error GoTo EliminaRep1FormulaBCRErr
'   Set dbConec = New DConecta
'   If dbConec.AbreConexion() Then
'      psSql = "EXEC stp_del_EliminaRep1FormulaBCR '" & psNumCol & "','" & psMoneda & "'"
'      dbConec.Ejecutar psSql
'      dbConec.CierraConexion
'   End If
'   Set dbConec = Nothing
'   Exit Sub
'EliminaRep1FormulaBCRErr:
'   Call RaiseError(MyUnhandledError, "DMntRepFormula:EliminaRep1FormulaBCR Method")
'End Sub

'*** PEAC 20131127
'Public Function CargaRep1FormulaBCR(psMoneda As String) As Recordset
'Dim sCond As String
'   On Error GoTo CargaRep1FormulaBCRErr
'   Set dbConec = New DConecta
'   If dbConec.AbreConexion() Then
'      psSql = "exec stp_sel_CargaRep1FormulaBCR '" & psMoneda & "'"
'      Set CargaRep1FormulaBCR = dbConec.CargaRecordSet(psSql)
'      dbConec.CierraConexion
'   End If
'   Set dbConec = Nothing
'   Exit Function
'CargaRep1FormulaBCRErr:
'   Call RaiseError(MyUnhandledError, "DMntRepFormula:CargaRep1FormulaBCR Method")
'End Function

'*** PEAC 20131210
'Public Sub ModificaRep1FormulaBCR(ByVal psNumCol As String, ByVal psDescripcion As String, ByVal cCodigo As String, ByVal pnCodRegimen As Integer, ByVal psValor As String, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal psMovNro As String, ByVal pnValor As Integer, ByVal pnNumColRef As Integer)
'On Error GoTo InsertaRep1FormulaBCRErr
'   Set dbConec = New DConecta
'   If dbConec.AbreConexion() Then
'
'      psSql = "exec stp_upd_Rep1FormulaBCR '" & psNumCol & "','" & psDescripcion & "','" & cCodigo & "'," & pnCodRegimen & ",'" & psValor & "'," & pnItem & "," & pnMoneda & ",'" & psMovNro & "'," & pnValor & "," & pnNumColRef
'
'      dbConec.Ejecutar psSql
'      dbConec.CierraConexion
'   End If
'   Set dbConec = Nothing
'   Exit Sub
'InsertaRep1FormulaBCRErr:
'   Call RaiseError(MyUnhandledError, "DMntRepFormula:InsertaRep1FormulaBCR Method")
'End Sub

'*** PEAC 20131217
'Public Function CargarColumnasRepoFormulaBCR(ByVal nRep As Integer, ByVal nMoneda As Integer) As Recordset
'Dim sCond As String
'   On Error GoTo CargarColumnasRepoFormulaBCRErr
'   Set dbConec = New DConecta
'   If dbConec.AbreConexion() Then
'        psSql = "exec stp_sel_CargarColumnasRepoFormulaBCR " & nRep & "," & nMoneda
'        Set CargarColumnasRepoFormulaBCR = dbConec.CargaRecordSet(psSql)
'        dbConec.CierraConexion
'   End If
'   Set dbConec = Nothing
'   Exit Function
'CargarColumnasRepoFormulaBCRErr:
'   Call RaiseError(MyUnhandledError, "DMntRepFormula:CargarColumnasRepoFormulaBCR Method")
'End Function

'*** PEAC 20131210
'Public Sub ModificaRep1ItemOrdenFormulaBCR(ByVal psNumCol As String, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal psMovNro As String)
'On Error GoTo ModificaRep1ItemOrdenFormulaBCRErr
'   Set dbConec = New DConecta
'   If dbConec.AbreConexion() Then
'
'      psSql = "exec stp_upd_Rep1ItemOrdenFormulaBCR '" & psNumCol & "'," & pnItem & "," & pnMoneda & ",'" & psMovNro & "'"
'
'      dbConec.Ejecutar psSql
'      dbConec.CierraConexion
'   End If
'   Set dbConec = Nothing
'   Exit Sub
'ModificaRep1ItemOrdenFormulaBCRErr:
'   Call RaiseError(MyUnhandledError, "DMntRepFormula:ModificaRep1ItemOrdenFormulaBCR Method")
'End Sub
'COMENTADO PASI FIN*******************************************************************************************
'***PASI 20140207 TI-ERS102-2013
Public Function CargarRegimenesRepoFormulaBCR() As Recordset
Dim sCond As String
   On Error GoTo CargarRegimenesRepoFormulaBCRErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_CargarRegimenesRep1FormulaBCR"
        Set CargarRegimenesRepoFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargarRegimenesRepoFormulaBCRErr:
   Call RaiseError(MyUnhandledError, "DMntRepFormula:CargarRegimenesRepoFormulaBCR Method")
End Function

'***PASI 20140207 TI-ERS102-2013
Public Sub InsertaRep1FormulaBCR(ByVal psNumCol As String, ByVal psDescripcion As String, ByVal cCodigo As String, ByVal pnCodRegimen As Integer, ByVal psValor As String, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal psMovNro As String, ByVal pnValor As Integer)
On Error GoTo InsertaRep1FormulaBCRErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
   
      psSql = "exec stp_ins_InsertaRep1FormulaBCR '" & psNumCol & "','" & psDescripcion & "','" & cCodigo & "'," & pnCodRegimen & ",'" & psValor & "'," & pnItem & "," & pnMoneda & ",'" & psMovNro & "'," & pnValor
      
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaRep1FormulaBCRErr:
   Call RaiseError(MyUnhandledError, "DMntRepFormula:InsertaRep1FormulaBCR Method")
End Sub

'***PASI 20140207 TI-ERS102-2013
Public Sub EliminaRep1FormulaBCR(psNumCol As String, pnMoneda As Integer)
   On Error GoTo EliminaRep1FormulaBCRErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "exec stp_del_EliminaRep1FormulaBCR '" & psNumCol & "','" & pnMoneda & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaRep1FormulaBCRErr:
   Call RaiseError(MyUnhandledError, "DMntRepFormula:EliminaRep1FormulaBCR Method")
End Sub

'***PASI 20140207 TI-ERS102-2013
Public Function CargaRep1FormulaBCR(ByVal pnMoneda As Integer) As Recordset
   On Error GoTo CargaRep1FormulaBCRErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "exec stp_sel_CargaRep1FormulaBCR '" & pnMoneda & "'"
      Set CargaRep1FormulaBCR = dbConec.CargaRecordSet(psSql)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaRep1FormulaBCRErr:
   Call RaiseError(MyUnhandledError, "DMntRepFormula:CargaRep1FormulaBCR Method")
End Function

'***PASI 20140207 TI-ERS102-2013
Public Sub ModificaRep1FormulaBCR(ByVal psNumCol As String, ByVal psDescripcion As String, ByVal cCodigo As String, ByVal pnCodRegimen As Integer, ByVal psValor As String, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal psMovNro As String, ByVal pnValor As Integer)
On Error GoTo ModificaRep1FormulaBCRBCRErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
   
      psSql = "exec stp_upd_ActualizaRep1FormulaBCR '" & psNumCol & "','" & psDescripcion & "','" & cCodigo & "'," & pnCodRegimen & ",'" & psValor & "'," & pnItem & "," & pnMoneda & ",'" & psMovNro & "'," & pnValor
      
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ModificaRep1FormulaBCRBCRErr:
   Call RaiseError(MyUnhandledError, "DMntRepFormula:ModificaRep1FormulaBCR Method")
End Sub

'***PASI 20140207 TI-ERS102-2013
Public Function CargarColumnasRepoFormulaBCR(ByVal nRep As Integer, ByVal nMoneda As Integer) As Recordset
Dim sCond As String
   On Error GoTo CargarColumnasRepoFormulaBCRErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_CargarColumnasRepoFormulaBCR " & nRep & "," & nMoneda
        Set CargarColumnasRepoFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargarColumnasRepoFormulaBCRErr:
   Call RaiseError(MyUnhandledError, "DMntRepFormula:CargarColumnasRepoFormulaBCR Method")
End Function

'***PASI 20140207 TI-ERS102-2013
Public Sub ModificaRep1ItemOrdenFormulaBCR(ByVal psNumCol As String, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal psMovNro As String)
On Error GoTo ModificaRep1ItemOrdenFormulaBCRErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
   
      psSql = "exec stp_upd_Rep1ItemOrdenFormulaBCR '" & psNumCol & "'," & pnItem & "," & pnMoneda & ",'" & psMovNro & "'"
      
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ModificaRep1ItemOrdenFormulaBCRErr:
   Call RaiseError(MyUnhandledError, "DMntRepFormula:ModificaRep1ItemOrdenFormulaBCR Method")
End Sub

'***PASI 20140207 TI-ERS102-2013
Public Function CargaRep2ColFormulaBCR(pnMoneda As Integer) As ADODB.Recordset
    On Error GoTo CargaRep2ColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_CargaRep2ColFormulaBCR " & pnMoneda
        Set CargaRep2ColFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
CargaRep2ColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:CargaRep2ColFormulaBCR Method")
End Function

'***PASI 20140207 TI-ERS102-2013
Public Sub InsertaRep2ColFormulaBCR(ByVal psCodigo As String, ByVal pnMoneda As Integer, ByVal psDescripcion As String, ByVal pnItem As Integer, ByVal psMovNro As String)
On Error GoTo InsertaRep2ColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_ins_InsertaRep2ColFormulaBCR '" & psCodigo & "'," & pnMoneda & ",'" & psDescripcion & "'," & pnItem & ",'" & psMovNro & "'"
        
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
InsertaRep2ColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertaRep2ColFormulaBCR Method")
End Sub
 '***PASI 20140207 TI-ERS102-2013
Public Sub ModificaRep2ColFormulaBCR(ByVal pcCodigo As String, ByVal pnMoneda As Integer, ByVal psDescripcion As String, ByVal pnItem As Integer, ByVal psMovNro As String)
On Error GoTo ModificaRep2ColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_upd_ActualizaRep2ColFormulaBCR '" & pcCodigo & "'," & pnMoneda & ",'" & psDescripcion & "'," & pnItem & ",'" & psMovNro & "'"
        
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
ModificaRep2ColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ModificaRep2ColFormulaBCR Method")
 End Sub
 
 '***PASI 20140210 TI-ERS102-2013
 Public Sub InsertaRep2SubColFormulaBCR(ByVal psCodSwif As String, ByVal psCodigo As String, ByVal psDescripcion As String, ByVal psValor As String, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal pnValor As Integer, ByVal psMovNro As String)
 On Error GoTo InsertaRep2SubColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_ins_InsertaRep2SubColFormulaBCR '" & psCodSwif & "','" & psCodigo & "','" & psDescripcion & "','" & psValor & "'," & pnItem & "," & pnMoneda & "," & pnValor & ",'" & psMovNro & "'"
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
InsertaRep2SubColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertaRep2SubColFormulaBCR Method")
 End Sub

'***PASI 20140210 TI-ERS102-2013
Public Function CargarComboRep2ColFormulaBCR(pnMoneda As Integer) As Recordset
    On Error GoTo CargarComboRep2ColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_CargarComboRep2ColFormulaBCR " & pnMoneda
        Set CargarComboRep2ColFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
CargarComboRep2ColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:CargarComboRep2ColFormulaBCR Method")
End Function

'***PASI 20140210 TI-ERS102-2013
Public Function CargarRep2SubColFormulaBCR(pnMoneda As Integer) As Recordset
    On Error GoTo CargarRep2SubColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_CargarRep2SubColFormulaBCR " & pnMoneda
        Set CargarRep2SubColFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
CargarRep2SubColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:CargarRep2SubColFormulaBCR Method")
End Function

'***PASI 20140212 TI-ERS102-2013
Public Sub ModificaRep2SubColFormulaBCR(ByVal psCodSwif As String, ByVal psCodigo As String, ByVal psDescripcion As String, ByVal psValor As String, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal pnValor As Integer, ByVal psMovNro As String)
    On Error GoTo ModificaRep2SubColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion Then
        psSql = "exec stp_upd_ActualizaRep2SubColFormulaBCR '" & psCodSwif & "','" & psCodigo & "','" & psDescripcion & "','" & psValor & "'," & pnItem & "," & pnMoneda & "," & pnValor & ",'" & psMovNro & "'"
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
ModificaRep2SubColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ModificaRep2SubColFormulaBCR Method")
End Sub

'*** PASI 20140212 TI-ERS102-2013
Public Sub EliminaRep2SubColFormulaBCR(ByVal pnItem As Integer, ByVal pnMoneda As Integer)
    On Error GoTo EliminaRep2SubColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion Then
        psSql = "exec stp_del_EliminaRep2SubColFormulaBCR " & pnItem & "," & pnMoneda
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
EliminaRep2SubColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:EliminaRep2SubColFormulaBCR Method")
End Sub

'*** PASI 20140213 TI-ERS102-2013
Public Sub ModificaRep2SubColItemOrdenFormulaBCR(ByVal pnValorRef As Integer, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal psMovNro As String)
    On Error GoTo ModificaRep2SubColItemOrdenFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_upd_Rep2SubColItemOrdenFormulaBCR " & pnValorRef & "," & pnItem & "," & pnMoneda & ",'" & psMovNro & "'"
        
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
ModificaRep2SubColItemOrdenFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ModificaRep2SubColItemOrdenFormulaBCR Method")
End Sub

'*** PASI 20140214 TI-ERS102-2013
Public Function CargarTipoObligacionesRep3FormulaBCR() As Recordset
    On Error GoTo CargarTipoObligacionesRep3FormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_CargarTipoObligacionesRep3FormulaBCR"
        Set CargarTipoObligacionesRep3FormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
CargarTipoObligacionesRep3FormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:CargarTipoObligacionesRep3FormulaBCR Method")
End Function

'*** PASI 20140214 TI-ERS102-2013
Public Function CargaRep3ColFormulaBCR(pnMoneda As Integer) As Recordset
    On Error GoTo CargaRep3ColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_CargaRep3ColFormulaBCR " & pnMoneda
        Set CargaRep3ColFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
CargaRep3ColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:CargaRep3ColFormulaBCR Method")
End Function

'*** PASI 20140214 TI-ERS102-2013
Public Sub InsertaRep3ColFormulaBCR(ByVal psCodigo As String, ByVal pnMoneda As Integer, ByVal psDescripcion As String, ByVal pnCodObl As Integer, ByVal pnItem As Integer, ByVal psMovNro As String)
On Error GoTo InsertaRep3ColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_ins_InsertaRep3ColFormulaBCR '" & psCodigo & "'," & pnMoneda & ",'" & psDescripcion & "'," & pnCodObl & "," & pnItem & ",'" & psMovNro & "'"
        
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
InsertaRep3ColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertaRep3ColFormulaBCR Method")
End Sub
'*** PASI 20140214 TI-ERS102-2013
Public Function CargarDestinoObligacionRep3FormulaBCR() As Recordset
    On Error GoTo CargarDestinoObligacionRep3FormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_CargarDestinoObligacionRep3FormulaBCR"
        Set CargarDestinoObligacionRep3FormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
CargarDestinoObligacionRep3FormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:CargarDestinoObligacionRep3FormulaBCR Method")
End Function

'*** PASI 20140214 TI-ERS102-2013
Public Function CargarComboRep3ColFormulaBCR(ByVal pnMoneda As Integer, ByVal pnCodObl As Integer) As Recordset
    On Error GoTo CargarComboRep3ColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_CargarComboRep3ColFormulaBCR " & pnMoneda & "," & pnCodObl
        Set CargarComboRep3ColFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
CargarComboRep3ColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:CargarComboRep3ColFormulaBCR Method")
End Function

'*** PASI 20140215 TI-ERS102-2013
Public Function CargarRep3SubColFormulaBCR(ByVal pnMoneda As Integer) As Recordset
    On Error GoTo CargarRep3SubColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_CargarRep3SubColFormulaBCR " & pnMoneda
        Set CargarRep3SubColFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
CargarRep3SubColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:CargarRep3SubColFormulaBCR Method")
End Function

'*** PASI 20140215 TI-ERS102-2013
Public Sub InsertaRep3SubColformulaBCR(ByVal psInstitucion As String, ByVal psCodSwif As String, ByVal pnDest As Integer, ByVal pdPeriodoDes As Date, ByVal pdPeriodoHas As Date, ByVal pnPlazoProm As Double, ByVal psCodigo As String, ByVal psValor As String, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal psMovNro As String, ByVal pnValor As Integer, ByVal pnCodOblRef As Integer)
    On Error GoTo InsertaRep3SubColformulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_ins_InsertaRep3SubColformulaBCR '" & psInstitucion & "','" & psCodSwif & "'," & pnDest & ",'" & Format(pdPeriodoDes, "yyyyMMdd") & "','" & Format(pdPeriodoHas, "yyyyMMdd") & "'," & pnPlazoProm & ",'" & psCodigo & "','" & psValor & "'," & pnItem & "," & pnMoneda & ",'" & psMovNro & "'," & pnValor & "," & pnCodOblRef
         dbConec.Ejecutar psSql
         dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
InsertaRep3SubColformulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertaRep3SubColformulaBCR Method")
End Sub

'*** PASI 20140217 TI-ERS102-2013
Public Sub ModificarRep3SubColFormulaBCR(ByVal psInstitucion As String, ByVal psCodSwif As String, ByVal pnDest As Integer, ByVal pdPeriodoDes As Date, ByVal pdPeriodoHas As Date, ByVal pnPlazoProm As Double, ByVal psCodigo As String, ByVal psValor As String, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal psMovNro As String, ByVal pnValor As Integer, ByVal pnCodOblRef As Integer)
    On Error GoTo ModificarRep3SubColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion Then
        psSql = "exec stp_upd_ActualizaRep3SubColFormulaBCR '" & psInstitucion & "','" & psCodSwif & "'," & pnDest & ",'" & Format(pdPeriodoDes, "yyyyMMdd") & "','" & Format(pdPeriodoHas, "yyyyMMdd") & "'," & pnPlazoProm & ",'" & psCodigo & "','" & psValor & "'," & pnItem & "," & pnMoneda & ",'" & psMovNro & "'," & pnValor & "," & pnCodOblRef
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
ModificarRep3SubColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ModificarRep3SubColFormulaBCR Method")
End Sub
 '*** PASI 20140220 TI-ERS102-2013
 Public Function ObtenerCabeceraRep1BaseFormulaBCR(pnMoneda As Integer) As Recordset
    On Error GoTo ObtenerCabeceraRep1BaseFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_ObtenerCabeceraRep1BaseFormulaBCR " & pnMoneda
        Set ObtenerCabeceraRep1BaseFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ObtenerCabeceraRep1BaseFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ModificarRep3SubColFormulaBCR Method")
 End Function

'*** PASI 20140222 TI-ERS102-2013
Public Function ObtenerRep1ColumnaCtaSaldo(ByVal psCtamas As String, psCtamenos As String, ByVal pdFecha As Date) As Recordset
    On Error GoTo ObtenerRep1ColumnaCtaSaldoErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_ObtenerRep1ColumnaCtaSaldo '" & psCtamas & "','" & psCtamenos & "','" & Format(pdFecha, "yyyyMMdd") & "'"
        Set ObtenerRep1ColumnaCtaSaldo = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ObtenerRep1ColumnaCtaSaldoErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerRep1ColumnaCtaSaldo Method")
End Function
 '*** PASI 20140224 TI-ERS102-2013
 Public Function ObtenerRep1ValorColumna(ByVal pnCol As Integer, ByVal pnMoneda As Integer) As Recordset
     On Error GoTo ObtenerRep1ValorColumnaErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_ObtenerRep1ValorColumna " & pnCol & "," & pnMoneda
        Set ObtenerRep1ValorColumna = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ObtenerRep1ValorColumnaErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerRep1ValorColumna Method")
 End Function

'*** PASI 20140225 TI-ERS102-2013
Public Function ObtenerRep2SubColdeColFormulaBCR(ByVal psCodigo As String, ByVal pnMoneda As Integer) As Recordset
    On Error GoTo ObtenerRep2SubColdeColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_ObtenerRep2SubColdeColFormulaBCR '" & psCodigo & "'," & pnMoneda
        Set ObtenerRep2SubColdeColFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ObtenerRep2SubColdeColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerRep2SubColdeColFormulaBCR Method")
End Function
 '*** PASI 20140225 TI-ERS102-2013
 Public Sub EliminaRep2ColFormulaBCR(ByVal psCodigo As String, ByVal pnMoneda As Integer)
    On Error GoTo EliminaRep2ColFormulaBCRERR
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_del_EliminaRep2ColFormulaBCR '" & psCodigo & "'," & pnMoneda
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
EliminaRep2ColFormulaBCRERR:
    Call RaiseError(MyUnhandledError, "DRepFormula:EliminaRep2ColFormulaBCR Method")
 End Sub

'*** PASI 20140226 TI-ERS102-2013
Public Function ObtenerColRep1FormulaBCR(ByVal pnMoneda As Integer) As Recordset
    On Error GoTo ObtenerColRep1FormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "stp_sel_ObtenerColRep1FormulaBCR " & pnMoneda
        Set ObtenerColRep1FormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ObtenerColRep1FormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerColRep1FormulaBCR Method")
End Function

'*** PASI 20140225 TI-ERS102-2013
Public Function ObtenerCabeceraRep2FormulaBCR(ByVal pnMoneda As Integer) As Recordset
    On Error GoTo ObtenerCabeceraRep2FormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_ObtenerCabeceraRep2FormulaBCR " & pnMoneda
        Set ObtenerCabeceraRep2FormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ObtenerCabeceraRep2FormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerCabeceraRep2FormulaBCR Method")
End Function

'*** PASI 20140225 TI-ERS102-2013
Public Function ObtenerSubColRep2FormulaBCR(ByVal pnMoneda As Integer) As Recordset
    On Error GoTo ObtenerSubColRep2FormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "stp_sel_ObtenerSubColRep2FormulaBCR " & pnMoneda
        Set ObtenerSubColRep2FormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ObtenerSubColRep2FormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerSubColRep2FormulaBCR Method")
End Function

'*** PASI 20140225 TI-ERS102-2013
Public Function ObtenerRep2SubColCtaSaldo(ByVal pdFecha As Date, ByVal psCta As String, ByVal pnMoneda As Integer) As Recordset
    On Error GoTo ObtenerRep2SubColCtaSaldoErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "stp_sel_ObtenerRep2SubColCtaSaldo '" & Format(pdFecha, "yyyyMMdd") & "','" & psCta & "'," & pnMoneda
        Set ObtenerRep2SubColCtaSaldo = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ObtenerRep2SubColCtaSaldoErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerRep2SubColCtaSaldo Method")
End Function

'*** PASI 20140228 TI-ERS102-2013
Public Sub ActualizarRep3ColFormulaBCR(ByVal psCodigo As String, ByVal pnMoneda As Integer, ByVal psDescripcion As String, ByVal nCodObl As Integer, ByVal pnItem As Integer, ByVal psMovNro As String)
    On Error GoTo ActualizarRep2ColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion Then
        psSql = "stp_upd_ActualizarRep3ColFormulaBCR '" & psCodigo & "'," & pnMoneda & ",'" & psDescripcion & "'," & nCodObl & "," & pnItem & ",'" & psMovNro & "'"
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
ActualizarRep2ColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ActualizarRep2ColFormulaBCR Method")
End Sub

'*** PASI 20140303 TI-ERS102-2013
 Public Sub EliminaRep3ColFormulaBCR(ByVal psCodigo As String, ByVal pnMoneda As Integer)
    On Error GoTo EliminaRep3ColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_del_EliminaRep3ColFormulaBCR '" & psCodigo & "'," & pnMoneda
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
EliminaRep3ColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:EliminaRep3ColFormulaBCRErr Method")
 End Sub
 '*** PASI 20140303 TI-ERS102-2013
 Public Function ObtenerRep3SubColdeColFormulaBCR(ByVal psCodigo As String, ByVal pnMoneda As Integer) As Recordset
    On Error GoTo ObtenerRep3SubColdeColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_ObtenerRep3SubColdeColFormulaBCR '" & psCodigo & "'," & pnMoneda
        Set ObtenerRep3SubColdeColFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ObtenerRep3SubColdeColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerRep3SubColdeColFormulaBCR Method")
End Function

'*** PASI 20140303 TI-ERS102-2013
Public Sub ModificaRep3SubColItemOrdenFormulaBCR(ByVal pnValorRef As Integer, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal psMovNro As String)
    On Error GoTo ModificaRep3SubColItemOrdenFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_upd_Rep3SubColItemOrdenFormulaBCR " & pnValorRef & "," & pnItem & "," & pnMoneda & ",'" & psMovNro & "'"
        
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
ModificaRep3SubColItemOrdenFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ModificaRep3SubColItemOrdenFormulaBCR Method")
End Sub

'*** PASI 20140303 TI-ERS102-2013
Public Sub EliminaRep3SubColFormulaBCR(ByVal pnItem As Integer, ByVal pnMoneda As Integer)
    On Error GoTo EliminaRep3SubColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion Then
        psSql = "exec stp_del_EliminaRep3SubColFormulaBCR " & pnItem & "," & pnMoneda
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
EliminaRep3SubColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:EliminaRep3SubColFormulaBCR Method")
End Sub
 
'*** PASI 20140303 TI-ERS102-2013
Public Sub InsertaRep4ColFormulaBCR(ByVal psCodigo As String, ByVal pnMoneda As Integer, ByVal psDescripcion As String, ByVal pnItem As Integer, ByVal psMovNro As String)
On Error GoTo InsertaRep4ColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_ins_InsertaRep4ColFormulaBCR '" & psCodigo & "'," & pnMoneda & ",'" & psDescripcion & "'," & pnItem & ",'" & psMovNro & "'"
        
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
InsertaRep4ColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertaRep4ColFormulaBCR Method")
End Sub

'*** PASI 20140303 TI-ERS102-2013
Public Function CargaRep4ColFormulaBCR(pnMoneda As Integer) As ADODB.Recordset
    On Error GoTo CargaRep4ColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_CargaRep4ColFormulaBCR " & pnMoneda
        Set CargaRep4ColFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
CargaRep4ColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:CargaRep4ColFormulaBCR Method")
End Function

'*** PASI 20140303 TI-ERS102-2013
Public Function CargarComboRep4ColFormulaBCR(pnMoneda As Integer) As Recordset
    On Error GoTo CargarComboRep4ColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_CargarComboRep4ColFormulaBCR " & pnMoneda
        Set CargarComboRep4ColFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
CargarComboRep4ColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:CargarComboRep4ColFormulaBCR Method")
End Function

'*** PASI 20140303 TI-ERS102-2013
Public Function CargarRep4SubColFormulaBCR(ByVal pnMoneda As Integer) As Recordset
    On Error GoTo CargarRep4SubColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_CargarRep4SubColFormulaBCR " & pnMoneda
        Set CargarRep4SubColFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
CargarRep4SubColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:CargarRep4SubColFormulaBCR Method")
End Function

'*** PASI 20140304 TI-ERS102-2013
Public Sub InsertaRep4SubColFormulaBCR(ByVal psDescripcion As String, ByVal psCodSwif As String, ByVal psCodigo As String, ByVal psValor As String, ByVal pnValor As Integer, ByVal pbAplicaPer As Boolean, ByVal pdFechaDes As Date, ByVal pdFechaHas As Date, ByVal pnPlazoProm As Double, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal psMovNro As String)
    On Error GoTo InsertaRep4SubColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_ins_InsertaRep4SubColFormulaBCR '" & psDescripcion & "','" & psCodSwif & "','" & psCodigo & "','" & psValor & "'," & pnValor & "," & IIf(pbAplicaPer, "true", "false") & ",'" & Format(pdFechaDes, "yyyyMMdd") & "','" & Format(pdFechaHas, "yyyyMMdd") & "'," & pnPlazoProm & "," & pnItem & "," & pnMoneda & ",'" & psMovNro & "'"
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
InsertaRep4SubColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertaRep4SubColFormulaBCR Method")
End Sub

'*** PASI 20140304 TI-ERS102-2013
Public Sub ModificaRep4ColFormulaBCR(ByVal pcCodigo As String, ByVal pnMoneda As Integer, ByVal psDescripcion As String, ByVal pnItem As Integer, ByVal psMovNro As String)
On Error GoTo ModificaRep4ColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_upd_ActualizaRep4ColFormulaBCR '" & pcCodigo & "'," & pnMoneda & ",'" & psDescripcion & "'," & pnItem & ",'" & psMovNro & "'"
        
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
ModificaRep4ColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ModificaRep4ColFormulaBCR Method")
 End Sub
 
'*** PASI 20140304 TI-ERS102-2013
Public Function ObtenerRep4SubColdeColFormulaBCR(ByVal psCodigo As String, ByVal pnMoneda As Integer) As Recordset
    On Error GoTo ObtenerRep4SubColdeColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_ObtenerRep4SubColdeColFormulaBCR '" & psCodigo & "'," & pnMoneda
        Set ObtenerRep4SubColdeColFormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ObtenerRep4SubColdeColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerRep4SubColdeColFormulaBCR Method")
End Function

'*** PASI 20140304 TI-ERS102-2013
 Public Sub EliminaRep4ColFormulaBCR(ByVal psCodigo As String, ByVal pnMoneda As Integer)
    On Error GoTo EliminaRep4ColFormulaBCRERR
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_del_EliminaRep4ColFormulaBCR '" & psCodigo & "'," & pnMoneda
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
EliminaRep4ColFormulaBCRERR:
    Call RaiseError(MyUnhandledError, "DRepFormula:EliminaRep4ColFormulaBCR Method")
 End Sub

'*** PASI 20140304 TI-ERS102-2013
Public Sub ModificaRep4SubColFormulaBCR(ByVal psDescripcion As String, ByVal psCodSwif As String, ByVal psCodigo As String, ByVal psValor As String, ByVal pnValor As Integer, ByVal pbAplicaPer As Boolean, ByVal pdFechaDes As Date, ByVal pdFechaHas As Date, ByVal pnPlazoProm As Double, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal psMovNro As String)
    On Error GoTo ModificaRep4SubColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_upd_ActualizaRep4SubColFormulaBCR '" & psDescripcion & "','" & psCodSwif & "','" & psCodigo & "','" & psValor & "'," & pnValor & "," & IIf(pbAplicaPer, "true", "false") & ",'" & Format(pdFechaDes, "yyyyMMdd") & "','" & Format(pdFechaHas, "yyyyMMdd") & "'," & pnPlazoProm & "," & pnItem & "," & pnMoneda & ",'" & psMovNro & "'"
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
ModificaRep4SubColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ModificaRep4SubColFormulaBCR Method")
End Sub
 
'*** PASI 20140304 TI-ERS102-2013
Public Sub ModificaRep4SubColItemOrdenFormulaBCR(ByVal pnValorRef As Integer, ByVal pnItem As Integer, ByVal pnMoneda As Integer, ByVal psMovNro As String)
    On Error GoTo ModificaRep4SubColItemOrdenFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_upd_Rep4SubColItemOrdenFormulaBCR " & pnValorRef & "," & pnItem & "," & pnMoneda & ",'" & psMovNro & "'"
        
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
ModificaRep4SubColItemOrdenFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ModificaRep4SubColItemOrdenFormulaBCR Method")
End Sub
'*** PASI 20140303 TI-ERS102-2013
Public Sub EliminaRep4SubColFormulaBCR(ByVal pnItem As Integer, ByVal pnMoneda As Integer)
    On Error GoTo EliminaRep4SubColFormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion Then
        psSql = "exec stp_del_EliminaRep4SubColFormulaBCR " & pnItem & "," & pnMoneda
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
EliminaRep4SubColFormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:EliminaRep4SubColFormulaBCR Method")
End Sub

'*** PASI 20140304 TI-ERS102-2013
Public Function ObtenerCabeceraRep3FormulaBCR(ByVal pnCodObl As Integer, ByVal pnMoneda As Integer) As Recordset
    On Error GoTo ObtenerCabeceraRep4FormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_ObtenerCabeceraRep3FormulaBCR " & pnCodObl & "," & pnMoneda
        Set ObtenerCabeceraRep3FormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ObtenerCabeceraRep4FormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerCabeceraRep4FormulaBCR Method")
End Function

'*** PASI 20140305 TI-ERS102-2013
Public Function ObtenerSubColRep3FormulaBCR(ByVal pnCodObl As Integer, ByVal pnMoneda As Integer) As Recordset
    On Error GoTo ObtenerSubColRep3FormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_ObtenerSubColRep3FormulaBCR " & pnCodObl & "," & pnMoneda
        Set ObtenerSubColRep3FormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ObtenerSubColRep3FormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerSubColRep3FormulaBCR Method")
End Function

'*** PASI 20140305 TI-ERS102-2013
Public Function ObtenerCabeceraRep4FormulaBCR(ByVal pnMoneda As Integer) As Recordset
    On Error GoTo ObtenerCabeceraRep4FormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "exec stp_sel_ObtenerCabeceraRep4FormulaBCR " & pnMoneda
        Set ObtenerCabeceraRep4FormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ObtenerCabeceraRep4FormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerCabeceraRep4FormulaBCR Method")
End Function

'*** PASI 20140305 TI-ERS102-2013
Public Function ObtenerSubColRep4FormulaBCR(ByVal pnMoneda As Integer) As Recordset
    On Error GoTo ObtenerSubColRep4FormulaBCRErr
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "stp_sel_ObtenerSubColRep4FormulaBCR " & pnMoneda
        Set ObtenerSubColRep4FormulaBCR = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ObtenerSubColRep4FormulaBCRErr:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerSubColRep4FormulaBCR Method")
End Function
'FRHU 20131223 RQ13657
Public Function RecuperaConfigFlujoEfectivo() As ADODB.Recordset
    On Error GoTo ErrRecuperaConfigFlujoEfectivo
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "Exec stp_sel_EstadoFlujoEfectivo"
        Set RecuperaConfigFlujoEfectivo = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrRecuperaConfigFlujoEfectivo:
   Call RaiseError(MyUnhandledError, "DRepFormula:RecuperaConfigFlujoEfectivo Method")
End Function
Public Sub RegistrarConfigFlujoEfectivo(ByVal pnId As Integer, ByVal pcDescripcion As String, ByVal pnNivel As Integer, ByVal pnTipo As Integer, ByVal pcValor As String, ByVal pbPeriodo As Boolean, ByVal pbEstado As Boolean, ByVal pnOrden As Integer)
    On Error GoTo ErrRegistrarConfigFlujoEfectivo
    psSql = "Exec stp_ins_EstadoFlujoEfectivoConfig " & pnId & ",'" & pcDescripcion & "'," & pnNivel & "," & pnTipo & ",'" & pcValor & "'," & IIf(pbPeriodo, 1, 0) & "," & IIf(pbEstado, 1, 0) & "," & pnOrden
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrRegistrarConfigFlujoEfectivo:
    Call RaiseError(MyUnhandledError, "DRepFormula:RegistrarConfigFlujoEfectivo Method")
End Sub
Public Sub InsertarHojasDeTrabajoFE(ByVal pnId As Integer, ByVal pnTipo As Integer, ByVal pnNivel As Integer, ByVal psDescripcion As String, _
                                    ByVal psForMen As String, ByVal psForMay As String, ByVal psAjusDebe As String, ByVal psAjusHaber As String, ByVal psOperDebe As String, _
                                    ByVal psOperHaber As String, ByVal psInveDebe As String, ByVal psInveHaber As String, ByVal psFinaDebe As String, ByVal psFinaHaber As String, _
                                    ByVal pbEstado As Boolean, ByVal pnOrden As Integer, ByVal pbMov As Boolean)
    On Error GoTo ErrInsertarHojasDeTrabajoFE
    psSql = "Exec stp_ins_upd_FEHojaDeTrabajo " & pnId & "," & pnTipo & "," & pnNivel & ",'" & psDescripcion & "','" & psForMen & "','" & psForMay & "','" & psAjusDebe & "','" & psAjusHaber & "','" & psOperDebe & "','" & psOperHaber & "','" & psInveDebe & "','" & psInveHaber & "','" & psFinaDebe & "','" & psFinaHaber & "'," & IIf(pbEstado, 1, 0) & "," & pnOrden & "," & IIf(pbMov, 1, 0)
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrInsertarHojasDeTrabajoFE:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertarHojasDeTrabajoFE Method")
End Sub
Public Function RecuperaHojaTrabajoFE(ByVal pnTipo As Integer) As ADODB.Recordset
    On Error GoTo ErrRecuperaHojaTrabajoFE
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "Exec stp_sel_FEHojaDeTrabajo " & pnTipo
        Set RecuperaHojaTrabajoFE = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrRecuperaHojaTrabajoFE:
   Call RaiseError(MyUnhandledError, "DRepFormula:RecuperaHojaTrabajoFE Method")
End Function
Public Sub ModificarEstadoHojaTrabajo(ByVal pnId As Integer)
    On Error GoTo ErrModificarEstadoHojaTrabajo
    psSql = "Exec stp_upd_EstadoFEHojaDeTrabajo " & pnId
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrModificarEstadoHojaTrabajo:
    Call RaiseError(MyUnhandledError, "DRepFormula:ModificarEstadoHojaTrabajo Method")
End Sub
'FIN FRHU 20131223
'FRHU 20140104 RQ 13658
Public Function ObtenerEstadoFlujoEfectivo() As ADODB.Recordset
    On Error GoTo ErrObtenerEstadoFlujoEfectivo
    Dim dbConecta As DConecta
    Set dbConecta = New DConecta
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    dbConecta.AbreConexion
        psSql = "Exec stp_sel_EstadoFlujoEfectivoConfig"
        Set oRs = dbConecta.CargaRecordSet(psSql)
    dbConecta.CierraConexion
    Set ObtenerEstadoFlujoEfectivo = oRs
    Set oRs = Nothing
    Exit Function
ErrObtenerEstadoFlujoEfectivo:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtenerEstadoFlujoEfectivo Method")
End Function
'FIN FRHU 20140104
'EJVG20140102 ***
Public Sub ActualizaNotaEstadoxNotaComplementaria(ByVal pnId As Long, ByVal psOpeCod As String, ByVal pnNotaEstadoNew As Long, ByVal pnNotaEstadoOld As Long)
    On Error GoTo ErrActualizaNotaEstadoxNotaComplementaria
    psSql = "Exec NIIF_stp_upd_MigrarNotaEstadoxNotaComplementaria " & pnId & ",'" & psOpeCod & "'," & pnNotaEstadoNew & "," & pnNotaEstadoOld
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrActualizaNotaEstadoxNotaComplementaria:
    Call RaiseError(MyUnhandledError, "DRepFormula:ActualizaNotaEstadoxNotaComplementaria Method")
End Sub
Public Sub LimpiaNotaEstadosxNuevasNotasComplementaria(ByVal psOpeCod As String)
    On Error GoTo ErrLimpiaNotaEstadosxNuevasNotasComplementaria
    psSql = "Exec NIIF_stp_upd_LimpiaNotaEstadosxNuevasNotasComplementaria '" & psOpeCod & "'"
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrLimpiaNotaEstadosxNuevasNotasComplementaria:
    Call RaiseError(MyUnhandledError, "DRepFormula:LimpiaNotaEstadosxNuevasNotasComplementaria Method")
End Sub
Public Sub ActualizaNotaComplementaria(ByVal psOpeCod As String, ByVal pnNotaEstado As Integer, ByVal pnSeccion As Integer, ByVal pnColumna As Integer, ByVal psFormulaNotaComplementaria As String, ByVal psFormulaNotaComplementaria_2012 As String)
    On Error GoTo ErrActualizaNotaComplementaria
    psSql = "Exec NIIF_stp_upd_ActualizaNotaComplementaria '" & psOpeCod & "'," & pnNotaEstado & "," & pnSeccion & "," & pnColumna & "," & IIf(Len(psFormulaNotaComplementaria) = 0, "NULL", "'" & psFormulaNotaComplementaria & "'") & "," & IIf(Len(psFormulaNotaComplementaria_2012) = 0, "NULL", "'" & psFormulaNotaComplementaria_2012 & "'")
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrActualizaNotaComplementaria:
    Call RaiseError(MyUnhandledError, "DRepFormula:ActualizaNotaComplementaria Method")
End Sub
Public Function ListaNotaComplementariasAll(ByVal psOpeCod As String) As ADODB.Recordset
    On Error GoTo ErrListaNotaComplementarias
    psSql = "Exec NIIF_stp_sel_ListaNotaComplementariasAll '" & psOpeCod & "'"
    Set ListaNotaComplementariasAll = dbConec.CargaRecordSet(psSql)
    Exit Function
ErrListaNotaComplementarias:
    Call RaiseError(MyUnhandledError, "DRepFormula:ListaNotaComplementariasAll Method")
End Function
Public Function ListaCalidadCrediticiaActivosFinancieros(ByVal pdFecha As Date, ByVal pbResumen As Boolean) As ADODB.Recordset
    On Error GoTo ErrListaCalidadCrediticiaActivosFinancieros
    psSql = "Exec NIIF_stp_sel_CalidadCrediticiaActivosFinancieros '" & Format(pdFecha, "yyyymmdd") & "'," & IIf(pbResumen, 1, 0)
    Set ListaCalidadCrediticiaActivosFinancieros = dbConec.CargaRecordSet(psSql)
    Exit Function
ErrListaCalidadCrediticiaActivosFinancieros:
    Call RaiseError(MyUnhandledError, "DRepFormula:ListaCalidadCrediticiaActivosFinancieros Method")
End Function
Public Function ListaCreditosVencidosxDiasAtraso(ByVal pdFecha As Date) As ADODB.Recordset
    On Error GoTo ErrListaCreditosVencidosxDiasAtraso
    psSql = "Exec NIIF_stp_sel_CreditosVencidosxDiasAtraso '" & Format(pdFecha, "yyyymmdd") & "'"
    Set ListaCreditosVencidosxDiasAtraso = dbConec.CargaRecordSet(psSql)
    Exit Function
ErrListaCreditosVencidosxDiasAtraso:
    Call RaiseError(MyUnhandledError, "DRepFormula:ListaCreditosVencidosxDiasAtraso Method")
End Function
Public Function ListaConcentracionRiesgoxSector(ByVal pdFecha As Date) As ADODB.Recordset
    On Error GoTo ErrListaConcentracionRiesgoxSector
    psSql = "Exec NIIF_stp_sel_ConcentracionRiesgoxSector '" & Format(pdFecha, "yyyymmdd") & "'"
    Set ListaConcentracionRiesgoxSector = dbConec.CargaRecordSet(psSql)
    Exit Function
ErrListaConcentracionRiesgoxSector:
    Call RaiseError(MyUnhandledError, "DRepFormula:ListaConcentracionRiesgoxSector Method")
End Function
'END EJVG *******
'EJVG20140909 ***
Public Function ListaProyeccionxOperacion(ByVal psOpeCod As String, ByVal pnAnio As Integer) As ADODB.Recordset
    On Error GoTo ErrListaProyeccionxOperacion
    psSql = "Exec stp_sel_ERS0572014_NIIFProyeccionxOperacion '" & psOpeCod & "'," & pnAnio
    Set ListaProyeccionxOperacion = dbConec.CargaRecordSet(psSql)
    Exit Function
ErrListaProyeccionxOperacion:
    Call RaiseError(MyUnhandledError, "DRepFormula:ListaProyeccionxOperacion Method")
End Function
Public Sub InsertarProyeccionxOperacion(ByVal psOpeCod As String, ByVal pnAnio As Integer, ByVal pnNota As Long, ByVal pnMes As Integer, ByVal pnMonto As Currency)
    On Error GoTo ErrInsertarProyeccionxOperacion
    psSql = "Exec stp_ins_ERS0572014_NIIFProyeccion '" & psOpeCod & "'," & pnAnio & "," & pnNota & "," & pnMes & "," & pnMonto
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrInsertarProyeccionxOperacion:
    Call RaiseError(MyUnhandledError, "DRepFormula:InsertarProyeccionxOperacion Method")
End Sub
Public Sub LimpiarProyeccionxOperacion(ByVal psOpeCod As String, ByVal pnAnio As Integer)
    On Error GoTo ErrLimpiarProyeccionxOperacion
    psSql = "Exec stp_del_ERS0572014_NIIFProyeccionxOperacion '" & psOpeCod & "'," & pnAnio
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrLimpiarProyeccionxOperacion:
    Call RaiseError(MyUnhandledError, "DRepFormula:LimpiarProyeccionxOperacion Method")
End Sub
Public Sub GuardarProyeccionTemporal(ByVal psOpeCod As String)
    On Error GoTo ErrGuardarProyeccionTemporal
    psSql = "Exec stp_ins_ERS0572014_GuardaTemporalProyeccion '" & psOpeCod & "'"
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrGuardarProyeccionTemporal:
    Call RaiseError(MyUnhandledError, "DRepFormula:GuardarProyeccionTemporal Method")
End Sub
Public Sub MigrarProyeccion(ByVal psOpeCod As String, ByVal pnNotaAnt As Long, ByVal pnNotaAct As Long)
    On Error GoTo ErrMigrarProyeccion
    psSql = "Exec stp_ins_ERS0572014_MigrarProyeccion '" & psOpeCod & "'," & pnNotaAnt & "," & pnNotaAct
    dbConec.Ejecutar (psSql)
    Exit Sub
ErrMigrarProyeccion:
    Call RaiseError(MyUnhandledError, "DRepFormula:MigrarProyeccion Method")
End Sub
'END EJVG *******
'PASI20150310
Public Function ObtieneRep1NumColCAJAME() As Integer
On Error GoTo ErrObtieneRep1NumColCAJAME
    Dim rs As ADODB.Recordset
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "stp_sel_ObtieneRep1NumColCAJAME"
        Set rs = dbConec.CargaRecordSet(psSql)
        If Not rs.EOF And Not rs.BOF Then
             ObtieneRep1NumColCAJAME = rs!NumCol
        Else
             ObtieneRep1NumColCAJAME = 0
        End If
    End If
    Set dbConec = Nothing
    Exit Function
ErrObtieneRep1NumColCAJAME:
    Call RaiseError(MyUnhandledError, "DRepFormula:ObtieneRep1NumColCAJAME Method")
End Function
'END PASI
