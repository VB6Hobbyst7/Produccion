VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim dbConec As COMConecta.DCOMConecta
Dim psSql As String
Dim prs   As ADODB.Recordset
Dim oError As COMConecta.COMErrorHandling

Dim sCentralCom As String

Public Function CargaOpeGru(Optional psOpeGruCod As String, Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaOpeGruErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      psSql = "SELECT cOpeGruCod, cOpeGruDesc, 1 as Nivel FROM " & sCentralCom & "OpeGrupo " & IIf(psOpeGruCod = "", "", "WHERE cOpeGruCod = '" & psOpeGruCod & "'")
      Set CargaOpeGru = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Sub InsertaOpeGru(psOpeGruCod As String, psOpeGruDesc As String, psMovNro As String)
   On Error GoTo InsertaOpeErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeGrupo (cOpeGruCod, cOpeGruDesc, cUltimaActualizacion) VALUES ('" & psOpeGruCod & "', '" & psOpeGruDesc & "', '" & psMovNro & "') "
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaOpeErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:InsertaOpeGru Method")
End Sub

Public Sub ActualizaOpeGru(psOpeGruCod As String, psOpeGruDesc As String, psMovNro As String)
   On Error GoTo InsertaOpeErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      psSql = "UPDATE " & sCentralCom & "OpeGrupo SET cOpeGruDesc = '" & psOpeGruDesc & "', cUltimaActualizacion = '" & psMovNro & "' WHERE cOpeGruCod = '" & psOpeGruCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaOpeErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:InsertaOpeGru Method")
End Sub

Public Sub EliminaOpeGru(psOpeGruCod As String)
   On Error GoTo InsertaOpeErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeGrupo WHERE cOpeGruCod = '" & psOpeGruCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaOpeErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:InsertaOpeGru Method")
End Sub

Public Function CargaOpeTpo(psOpeCod As String, Optional plLike As Boolean = False, Optional LockType As LockTypeEnum = adLockReadOnly, Optional pnVisible As Integer = -1, Optional pnNivel As Integer = 0) As Recordset
Dim lsFiltro As String
Dim oVar As New COMFunciones.FCOMVarPublicas

   On Error GoTo CargaOpeTpoErr
   lsFiltro = ""
   If pnVisible <> -1 Then
        lsFiltro = " and cOpeVisible = '" & pnVisible & "'"
   End If
   If pnNivel > 0 Then
        lsFiltro = lsFiltro & " and nOpeNiv = " & pnNivel & " "
   End If
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      psSql = "SELECT cOpeCod,cOpeDesc,nOpeNiv, CASE WHEN cOpeVisible='" & gOpeTpoVisibleSi & "' THEN '" & oVar.gsSI & "' ELSE '" & oVar.gsNO & "' END cOpeVisible, ISNULL(cOpeGruCod,'') cOpeGruCod FROM " & sCentralCom & "OpeTpo" & IIf(psOpeCod = "", "", " WHERE cOpeCod " & IIf(plLike, " LIKE '", "= '") & psOpeCod & IIf(plLike, "%'", "'")) & lsFiltro
      Set CargaOpeTpo = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeTpoErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:CargaOpeTpo Method")
End Function

Public Sub InsertaOpeTpo(psOpeCod As String, psOpeDesc As String, psOpeVisible As String, pnOpeNiv As Integer, psOpeTpo As String, psMovNro As String)
   On Error GoTo InsertaOpeTpoErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeTpo (cOpeCod, cOpeDesc, cOpeVisible, nOpeNiv, cOpeGruCod, cUltimaActualizacion) VALUES ('" & psOpeCod & "', '" & Replace(psOpeDesc, "'", "''") & "', '" & psOpeVisible & "', " & pnOpeNiv & ", '" & psOpeTpo & "','" & psMovNro & "') "
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaOpeTpoErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:InsertaOpeTpo Method")
End Sub

Public Sub ActualizaOpeTpo(psOpeCod As String, psOpeDesc As String, psOpeVisible As String, pnOpeNiv As Integer, psOpeTpo As String, psMovNro As String)
   On Error GoTo ActualziaOpeTpoErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      psSql = "UPDATE " & sCentralCom & "OpeTpo SET cOpeDesc = '" & psOpeDesc & "', cOpeVisible = '" & psOpeVisible & "', nOpeNiv = " & pnOpeNiv & ", cOpeGruCod = '" & psOpeTpo & "', cUltimaActualizacion = '" & psMovNro & "' " _
            & "WHERE cOpeCod = '" & psOpeCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualziaOpeTpoErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:ActualziaOpeTpo Method")
End Sub

Public Sub EliminaOpeTpo(psOpeCod As String)
Dim lsError As String
Dim lsTrans As Boolean
   On Error GoTo EliminaOpeTpoErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      lsTrans = True
      dbConec.BeginTrans
      psSql = "DELETE " & sCentralCom & "OpeDoc WHERE cOpeCod = '" & psOpeCod & "' "
      dbConec.Ejecutar psSql
      psSql = "DELETE " & sCentralCom & "OpeObj WHERE cOpeCod = '" & psOpeCod & "' "
      dbConec.Ejecutar psSql
      psSql = "DELETE " & sCentralCom & "OpeCta WHERE cOpeCod = '" & psOpeCod & "' "
      dbConec.Ejecutar psSql
      psSql = "DELETE " & sCentralCom & "OpeTpo WHERE cOpeCod = '" & psOpeCod & "' "
      dbConec.Ejecutar psSql
      dbConec.CommitTrans
      lsTrans = False
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaOpeTpoErr:
   lsError = Err.Description
   If lsTrans Then
      dbConec.RollbackTrans
   End If
   Set dbConec = Nothing
   Err.Raise Err.Number + oError.MyUnhandledError, "DOperación: EliminaOpeTpo", lsError
End Sub
Public Function CargaOpeCta(psOpeCod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "", Optional LockType As LockTypeEnum = adLockReadOnly, Optional pbGetClase As Boolean = False, Optional pnNivSuperior As Integer = 0) As Recordset
   On Error GoTo CargaOpeCtaErr
   Set dbConec = New COMConecta.DCOMConecta
   Dim oVar As New COMFunciones.FCOMVarPublicas
   Dim lsFiltroDH As String
   Dim lsFiltroOrden As String
    If psDH <> "" Then
        lsFiltroDH = " AND a.cOpeCtaDH ='" & psDH & "' "
    End If
    If psOpeCtaOrden <> "" Then
        lsFiltroOrden = " And a.cOpeCtaOrden='" & psOpeCtaOrden & "' "
    End If
   If dbConec.AbreConexion() Then
      psSql = "select " _
         & "   a.cCtaContCod, dbo.GetCtaContDesc(a.cCtaContCod," & pnNivSuperior & ",0) cCtaContDesc, 1 as Nivel ,  " _
         & "  CASE WHEN a.cOpeCtaDH = '" & oVar.gsContHaber & "' THEN '" & oVar.gsContHaberDesc & "' " _
         & "       WHEN a.cOpeCtaDH = '" & oVar.gsContDebe & "' THEN '" & oVar.gsContDebeDesc & "' END  as Clase, " _
         & "  CASE WHEN a.cOpeCtaOpc = '" & COMDConstantes.gOpeCtaCaracObligatorio & "' then '" & oVar.gsOpeCtaCaracterObligaDesc & "' WHEN a.cOpeCtaOpc='" & COMDConstantes.gOpeCtaCaracObligatorio & "' then '" & oVar.gsOpeCtaCaracterOpcionDesc & "' END as Caracter, " _
         & "  a.cOpeCod, a.cOpeCtaDH, a.cOpeCtaOpc, a.cCtaContN, a.cOpeCtaOrden " & IIf(pbGetClase, ", cls.cCtaCaracter ", "") _
         & "FROM " & sCentralCom & "OpeCta as a " & IIf(pbGetClase, " JOIN " & sCentralCom & "CtaContClase cls ON a.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' ", "") _
         & "WHERE a.cOpeCod LIKE '" & psOpeCod & "' " & lsFiltroDH & lsFiltroOrden _
         & "order by a.cOpeCtaDH, a.cOpeCtaOrden, a.cCtaContCod "
      Set CargaOpeCta = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeCtaErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:CargaOpeCta Method")
End Function

Public Function CargaOpeCtaArbol(psOpeCod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "") As Recordset
   On Error GoTo CargaOpeCtaErr
   Set dbConec = New COMConecta.DCOMConecta
   Dim lsFiltroDH As String
   Dim lsFiltroOrden As String
    If psDH <> "" Then
        lsFiltroDH = " AND a.cOpeCtaDH ='" & psDH & "' "
    End If
    If psOpeCtaOrden <> "" Then
        lsFiltroOrden = " And a.cOpeCtaOrden='" & psOpeCtaOrden & "' "
    End If
   If dbConec.AbreConexion() Then
      psSql = "SELECT DISTINCT a.cCtaContCod, b.cCtaContDesc, 1 as nNivel " _
         & "FROM " & sCentralCom & "OpeCta as a JOIN " & sCentralCom & "CtaCont as b ON b.cCtaContCod = a.cCtaContCod " _
         & "WHERE a.cOpeCod LIKE '" & psOpeCod & "%' " & lsFiltroDH & lsFiltroOrden _
         & "ORDER BY a.cCtaContCod "
      Set CargaOpeCtaArbol = dbConec.CargaRecordSet(psSql, adLockReadOnly)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeCtaErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:CargaOpeCta Method")
End Function

Public Sub InsertaOpeCta(psOpeCod As String, psOpeCtaOrden As String, psCtaCod As String, psOpeCtaDH As String, psOpeCtaOpc As String, psCtaCodN As String, psMovNro As String)
   On Error GoTo InsertaOpeCtaErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeCta (cOpeCod, cOpeCtaOrden , cCtaContCod, cOpeCtaDH, cOpeCtaOpc, cCtaContN, cUltimaActualizacion) VALUES ('" & psOpeCod & "', '" & psOpeCtaOrden & "','" & psCtaCod & "','" & psOpeCtaDH & "','" & psOpeCtaOpc & "', '" & psCtaCodN & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaOpeCtaErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:InsertaOpeCta Method")
End Sub

Public Sub EliminaOpeCta(psOpeCod As String, psOpeCtaOrden As String, psCtaCod As String, psOpeCtaDH As String)
   On Error GoTo EliminaOpeCtaErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeCta WHERE cOpeCod = '" & psOpeCod & "' and cOpeCtaOrden = '" & psOpeCtaOrden & "' and cCtaContCod = '" & psCtaCod & "' and cOpeCtaDH = '" & psOpeCtaDH & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaOpeCtaErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:EliminaOpeCta Method")
End Sub

Public Function CargaOpeObj(psOpeCod As String, Optional LockType As LockTypeEnum = adLockReadOnly, Optional psOpeObjOrden As String = "") As Recordset
   On Error GoTo CargaOpeObjErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      psSql = "SELECT a.cOpeObjOrden ,a.cObjetoCod, b.cObjetoDesc, a.cOpeObjValor, " & _
         "       a.nOpeObjNiv, a.cOpeObjFiltro, a.cOpeCod" & _
         "  FROM  " & sCentralCom & "OpeObj as a JOIN " & sCentralCom & "Objeto as b ON b.cObjetoCod = a.cObjetoCod " & _
         " WHERE a.cOpeCod = '" & psOpeCod & "' " & IIf(psOpeObjOrden = "", "", " and cOpeObjOrden = '" & psOpeObjOrden & "' ")
      Set CargaOpeObj = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeObjErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:CargaOpeObj Method")
End Function

Public Sub InsertaOpeObj(psOpeCod As String, psOpeObjOrden As String, psObjetoCod As String, psOpeObjValor As String, pnOpeObjNiv As Integer, psOpeObjFiltro As String, psMovNro As String)
   On Error GoTo InsertaOpeObjErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeObj (cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjValor, nOpeObjNiv, cOpeObjFiltro, cUltimaActualizacion) " _
            & "VALUES ('" & psOpeCod & "','" & psOpeObjOrden & "','" & psObjetoCod & "','" & psOpeObjValor & "'," & pnOpeObjNiv & ",'" & psOpeObjFiltro & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaOpeObjErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:InsertaOpeObj Method")
End Sub

Public Sub EliminaOpeObj(psOpeCod As String, psOpeObjOrden As String, psObjetoCod As String)
   On Error GoTo EliminaOpeObjErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeObj WHERE cOpeCod = '" & psOpeCod & "' and cOpeObjOrden = '" & psOpeObjOrden & "' and cObjetoCod = '" & psObjetoCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaOpeObjErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:EliminaOpeObj Method")
End Sub

Public Function CargaOpeDoc(psOpeCod As String, Optional pnDocTpo As TpoDoc = -1, Optional psMetodo As OpeDocMetodo = -1, Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaOpeDocErr
   Set dbConec = New COMConecta.DCOMConecta
   Dim lsFiltro As String
   lsFiltro = ""
   If Not pnDocTpo = -1 Then
      lsFiltro = " AND a.nDocTpo = " & pnDocTpo & " "
   End If
   If Not psMetodo = -1 Then
        lsFiltro = lsFiltro & " AND a.cOpeDocMetodo ='" & psMetodo & "'"
   End If
   If dbConec.AbreConexion() Then
      psSql = "SELECT a.cOpeCod, a.nDocTpo, d.cDocDesc, " & _
       "CASE WHEN a.cOpeDocEstado = '11'  THEN 'Obligatorio - Debe Existir' " & _
       "     WHEN a.cOpeDocEstado = '12'  THEN 'Obligatorio - No debe Existir' " & _
       "     WHEN a.cOpeDocEstado = '21'  THEN 'Opcional - Debe Existir' " & _
       "     WHEN a.cOpeDocEstado = '22'  THEN 'Opcional - No debe Existir' END as Estado, " & _
       "CASE WHEN a.cOpeDocMetodo = '1' THEN 'Nro. de Movimiento' " & _
       "     WHEN a.cOpeDocMetodo = '2' THEN 'Autogenerado' " & _
       "     WHEN a.cOpeDocMetodo = '3' THEN 'Digitado' END as Metodo , D.cDocAbrev " & _
       "FROM " & sCentralCom & "OpeDoc a JOIN " & sCentralCom & "Documento d ON d.nDocTpo = a.nDocTpo " & _
       "WHERE cOpeCod = '" & psOpeCod & "' " & lsFiltro & " ORDER BY a.nDocTpo "
       
      Set CargaOpeDoc = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeDocErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:CargaOpeDoc Method")
End Function

Public Sub InsertaOpeDoc(psOpeCod As String, pnDocTpo As Long, psOpeDocEstado As String, psOpeDocMetodo As String, psMovNro As String)
   On Error GoTo InsertaOpeDocErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeDoc (cOpeCod, nDocTpo, cOpeDocEstado, cOpeDocMetodo, cUltimaActualizacion) VALUES ('" & psOpeCod & "', '" & pnDocTpo & "', '" & psOpeDocEstado & "','" & psOpeDocMetodo & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaOpeDocErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:InsertaOpeDoc Method")
End Sub

Public Sub EliminaOpeDoc(psOpeCod As String, pnDocTpo As Long)
   On Error GoTo EliminaOpeDocErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeDoc WHERE cOpeCod = '" & psOpeCod & "' and nDocTpo = " & pnDocTpo & ""
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaOpeDocErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:EliminaOpeDoc Method")
End Sub

Private Sub Class_Initialize()
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Set oImp = New COMDConstSistema.DCOMImpresoras
    Set oError = New COMConecta.COMErrorHandling
    
    Dim oImpresora As New COMFunciones.FCOMVarImpresion
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As New COMConecta.DCOMClasIni
sCentralCom = oIni.BaseComunes
Set oIni = Nothing
End Sub

' Parches
Public Function EmiteDocOpe(ByVal psOpeCod As String, ByVal psOpeDocEstado As COMDConstantes.OpeDocEstado, ByVal psOpeDocMetodo As COMDConstantes.OpeDocMetodo) As String
On Error GoTo ErrorEmiteDocOpe
Dim rs As ADODB.Recordset
Dim sSql As String
Dim lsFiltroOpeDocEstado As String
Dim lsFiltroOpeDocMetodo As String
Set rs = New ADODB.Recordset


lsFiltroOpeDocEstado = "  AND cOpeDocEstado ='" & psOpeDocEstado & "'  "
lsFiltroOpeDocMetodo = "  AND cOpeDocMetodo ='" & psOpeDocMetodo & "'  "
Dim oConect As COMConecta.DCOMConecta
Set oConect = New COMConecta.DCOMConecta
EmiteDocOpe = ""
If oConect.AbreConexion = False Then Exit Function
sSql = "Select nDocTpo from " & sCentralCom & "OpeDoc where cOpeCod = '" & psOpeCod & "' " & lsFiltroOpeDocEstado & lsFiltroOpeDocMetodo
Set rs = oConect.CargaRecordSet(sSql)
If Not rs.EOF And Not rs.BOF Then
    EmiteDocOpe = rs!nDocTpo
End If
rs.Close: Set rs = Nothing
oConect.CierraConexion
Set oConect = Nothing
Exit Function
ErrorEmiteDocOpe:
    Err.Raise vbObjectError + 100, "ErrorEmiteDocOpe", "Error emitir documento en Operacion"
End Function

'FUNCION QUE DEVUELVE EL CODIGO DE OPERACION QUE FUE DEFINIDO CON UN DOCUMENTO ESPECIFICO
'Para Arendir Cuenta : -1 Efectivo; 48 Orden de Pago, XX Cheque, 46 Carta
Public Function EmiteOpeDoc(ByVal psOpeCod As String, ByVal psDocTpo As COMDConstantes.TpoDoc, Optional ByVal psDH As String = "") As String
On Error GoTo ErrorEmiteOpeDoc
Dim rs As ADODB.Recordset
Dim sSql As String
Dim lsFiltroDH As String
Set rs = New ADODB.Recordset

Dim oConect As COMConecta.DCOMConecta
Set oConect = New COMConecta.DCOMConecta
EmiteOpeDoc = ""
If oConect.AbreConexion = False Then Exit Function
        
If psDH <> "" Then
    lsFiltroDH = " AND OC.cOpeCtaDH ='" & psDH & "'"
End If

sSql = "Select OC.COPECOD, OD.nDocTpo  " _
    & " From   " & sCentralCom & "OpeCta OC Left Join " & sCentralCom & "OpeDoc OD on OD.cOpeCod = OC.cOpeCod  " _
    & " Where  Substring(OC.cOpeCod,1,LEN('" & psOpeCod & "')) ='" & Mid(psOpeCod, 1, Len(psOpeCod)) & "' and OD.nDocTpo" & IIf(psDocTpo = -1, " IS NULL ", "= " & psDocTpo & " ") _
    & lsFiltroDH _
    & " GROUP BY OC.COPECOD, OD.nDocTpo  "
    
Set rs = oConect.CargaRecordSet(sSql)
If Not rs.EOF And Not rs.BOF Then
    EmiteOpeDoc = rs!cOpecod
End If
rs.Close: Set rs = Nothing
oConect.CierraConexion
Set oConect = Nothing
Exit Function
ErrorEmiteOpeDoc:
    Err.Raise vbObjectError + 100, "ErrorEmiteDocOpe", "Error emitir documento en Operacion"
End Function

Public Function EmiteSubCtaOpeObj(ByVal psOpeCod As String, ByVal psCtaContCod As String) As String
Dim SQL As String
Dim rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
Set oConect = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
Dim lsSubCta As String
lsSubCta = ""
If oConect.AbreConexion = False Then Exit Function

SQL = "Select  cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjFiltro, " _
    & " CASE " _
    & "         WHEN OO.COBJETOCOD ='" & Format(ObjEntidadesFinancieras, "00") & "' THEN " _
    & "                 ISNULL((SELECT cCtaIFSubCta FROM CTAIFFILTRO WHERE (cPersCod = OO.cOpeObjFiltro OR cCtaIfCod = OO.cOpeObjFiltro ) AND cCtaContCod in ('" & psCtaContCod & "')),'') " _
    & "         WHEN OO.COBJETOCOD ='" & ObjCMACAgenciaArea & "' THEN " _
    & "                 ISNULL((SELECT cSubCtaCod FROM CtaAreaAgeFiltro WHERE cAreaCod + cAgeCod LIKE OO.cOpeObjFiltro AND cCtaContCod in ('" & psCtaContCod & "')) , '') " _
    & "         WHEN OO.COBJETOCOD ='" & ObjCMACAgencias & "' THEN " _
    & "                 ISNULL((Select cSubCtaCod from AGENCIAS WHERE cAgeCod LIKE OO.cOpeObjFiltro ),'') " _
    & "         WHEN OO.COBJETOCOD ='" & ObjCMACArea & "' THEN " _
    & "                 ISNULL((SELECT cSubCtaCod FROM AREAS WHERE cAreaCod LIKE OO.cOpeObjFiltro ),'') " _
    & "         Else    ISNULL((SELECT cCtaObjSubCta FROM CTAOBJFILTRO WHERE COBJETOCOD LIKE OO.cOpeObjFiltro AND cCtaContCod in ('" & psCtaContCod & "')),'')  END as SUBCUENTA " _
    & " From    OpeObj OO " _
    & " Where   OO.cOpecod Like '" & psOpeCod & "' " _
    & " ORDER BY OO.COPEOBJORDEN "

Set rs = oConect.CargaRecordSet(SQL)
If Not rs.EOF And Not rs.BOF Then
    Do While Not rs.EOF
        lsSubCta = lsSubCta + IIf(IsNull(rs!SUBCUENTA), "", rs!SUBCUENTA)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing
EmiteSubCtaOpeObj = lsSubCta
End Function

Public Function GetOpeCtaDebeHaber(ByVal psOpeCod As String, ByVal psCtaContCod As String, Optional ByVal psOpeCtaOrden As String = "0") As String
Dim SQL As String
Dim rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

SQL = " SELECT   cOpeCtaDH " _
    & " FROM    " & sCentralCom & "OPECTANEG OC " _
    & " WHERE    COPECOD = '" & psOpeCod & "' and cCtaContcod ='" & psCtaContCod & "'"
GetOpeCtaDebeHaber = ""
Set rs = oConect.CargaRecordSet(SQL)
If Not rs.EOF And Not rs.BOF Then
    GetOpeCtaDebeHaber = rs!cOpeCtaDH
End If
rs.Close
Set rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function EmiteOpeCtasNivel(ByVal psOpeCod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "0") As ADODB.Recordset
Dim SQL As String
Dim rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
Dim lsDHFiltro As String
Dim lsOrdenFiltro As String
If psDH <> "" Then
    lsDHFiltro = " AND cOpeCtaDH ='" & psDH & "'"
End If
If psOpeCtaOrden <> "" Then
   lsOrdenFiltro = " and OC.cOpeCtaOrden='" & psOpeCtaOrden & "'"
End If

SQL = "SELECT   OC.cCtaContCod , ISNULL(dbo.GetCtaContDesc(OC.cCtaContCod,1,1),'') cCtaContDesc, 1 as Nivel " _
    & "FROM     " & sCentralCom & "OPECTANEG OC " _
    & "         " _
    & " WHERE   COPECOD = '" & psOpeCod & "' " & lsDHFiltro

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

Set rs = oConect.CargaRecordSet(SQL)
Set EmiteOpeCtasNivel = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

'Public Function GetOpeObj(ByVal psOpeCod As String, ByVal psOpeObjOrden As String, _
'                        Optional psCtaContCod As String = "", _
'                        Optional ByVal psFiltro As String, Optional ByVal psFiltroAdd As String = "") As ADODB.Recordset
'Dim sql As String
'Dim rs As ADODB.Recordset
'Dim rs1 As ADODB.Recordset
'Dim oConect As COMConecta.DCOMConecta
'Dim oRHAreas As DActualizaDatosArea
'Dim oContFunct As NContFunciones
'Dim oCtaIf As NCajaCtaIF
'Dim oEfect As Defectivo
'Dim lsFiltro As String
'Dim lsFiltroIf As String
'Dim lsFiltroCta As String
'
'Set oCtaIf = New NCajaCtaIF
'Set oContFunct = New NContFunciones
'Set oConect = New COMConecta.DCOMConecta
'Set oRHAreas = New DActualizaDatosArea
'Set oEfect = New Defectivo
'
'Set rs = New ADODB.Recordset
'Set rs1 = New ADODB.Recordset
'If oConect.AbreConexion = False Then
'    Exit Function
'End If
'sql = " SELECT cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjFiltro, nOpeObjNiv " _
'    & " FROM " & sCentralCom & "OPEOBJ " _
'    & " WHERE COPECOD ='" & psOpeCod & "' and cOpeObjOrden='" & psOpeObjOrden & "'"
'
'Set rs = oConect.CargaRecordSet(sql)
'If Not rs.EOF And Not rs.BOF Then
'    If psFiltro = "" Then
'        lsFiltro = Trim(rs!cOpeObjFiltro) & psFiltroAdd
'    Else
'        lsFiltro = psFiltro & psFiltroAdd
'    End If
'    Select Case Val(rs!cObjetoCod)
'            Case ObjCMACAgencias
'                Set rs1 = oRHAreas.GetAgencias(lsFiltro)
'            Case ObjCMACAgenciaArea
'                Set rs1 = oRHAreas.GetAgenciasAreas(lsFiltro)
'            Case ObjCMACArea
'                Set rs1 = oRHAreas.GetAreas(lsFiltro)
'            Case ObjEntidadesFinancieras
'                lsFiltroIf = lsFiltro
'                lsFiltroCta = lsFiltro
'                If rs!nOpeObjNiv = 1 Then
'                    Set rs1 = oCtaIf.GetInstFinancieras(lsFiltro)
'                Else
'                    Set rs1 = oCtaIf.CargaCtasIF(Mid(psOpeCod, 3, 1), lsFiltro, rs!nOpeObjNiv)
'                End If
'            Case ObjDescomEfectivo
'                Set rs1 = oEfect.GetBilletajes(lsFiltro)
'            Case ObjPersona
'                Set rs1 = Nothing
'            Case Else
'                'Set rs1 = oContFunct.GetObjetos(Val(IIf(lsFiltro = "", rs!cObjetoCod, rs!cOpeObjFiltro)), IIf(rs!cOpeObjFiltro = "", true, false))
'                Set rs1 = oContFunct.GetObjetos(Val(IIf(lsFiltro = "", rs!cObjetoCod, rs!cOpeObjFiltro)), False)
'        End Select
'End If
'rs.Close
'Set rs = Nothing
'oConect.CierraConexion: Set oConect = Nothing
'
'Set oCtaIf = Nothing
'Set oContFunct = Nothing
'Set oConect = Nothing
'Set oRHAreas = Nothing
'Set oEfect = Nothing
'
'Set GetOpeObj = rs1
'
'End Function

'Public Function GetRsOpeObj(ByVal psOpeCod As String, ByVal psOpeObjOrden As String, Optional ByVal psFiltroObj As String, Optional ByVal psFiltroAdd As String = "") As ADODB.Recordset
'Dim sql As String
'Dim rs As ADODB.Recordset
'Dim rs1 As ADODB.Recordset
'Dim oConect As COMConecta.DCOMConecta
'Dim oRHAreas As DActualizaDatosArea
'Dim oContFunct As NContFunciones
'Dim oCtaIf As NCajaCtaIF
'Dim oEfect As Defectivo
'Dim lsFiltro As String
'Dim lsFiltroIf As String
'Dim lsFiltroCta As String
'
'Set oCtaIf = New NCajaCtaIF
'Set oContFunct = New NContFunciones
'Set oConect = New COMConecta.DCOMConecta
'Set oRHAreas = New DActualizaDatosArea
'Set oEfect = New Defectivo
'
'Set rs = New ADODB.Recordset
'Set rs1 = New ADODB.Recordset
'If oConect.AbreConexion = False Then
'    Exit Function
'End If
'sql = " SELECT cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjFiltro, nOpeObjNiv " _
'    & " FROM " & sCentralCom & "OPEOBJ " _
'    & " WHERE COPECOD ='" & psOpeCod & "' and cOpeObjOrden='" & psOpeObjOrden & "'"
'
'Set rs = oConect.CargaRecordSet(sql)
'If Not rs.EOF And Not rs.BOF Then
'    If psFiltroObj = "" Then
'        lsFiltro = Trim(rs!cOpeObjFiltro) & psFiltroAdd
'    Else
'        lsFiltro = psFiltroObj & psFiltroAdd
'    End If
'    Select Case Val(rs!cObjetoCod)
'            Case ObjCMACAgencias
'                Set rs1 = oRHAreas.GetAgencias(lsFiltro)
'            Case ObjCMACAgenciaArea
'                Set rs1 = oRHAreas.GetAgenciasAreas(lsFiltro)
'            Case ObjCMACArea
'                Set rs1 = oRHAreas.GetAreas(lsFiltro)
'            Case ObjEntidadesFinancieras
'                Set rs1 = oCtaIf.CargaCtasIF(Mid(psOpeCod, 3, 1), lsFiltro, rs!nOpeObjNiv)
'            Case ObjDescomEfectivo
'                Set rs1 = oEfect.GetBilletajes(lsFiltro)
'            Case ObjPersona
'                Set rs1 = Nothing
'            Case Else
'                Set rs1 = oContFunct.GetObjetos(Val(IIf(lsFiltro = "", rs!cObjetoCod, rs!cOpeObjFiltro)), IIf(rs!cOpeObjFiltro = "", True, False))
'        End Select
'End If
'rs.Close
'Set rs = Nothing
'oConect.CierraConexion: Set oConect = Nothing
'
'Set oCtaIf = Nothing
'Set oContFunct = Nothing
'Set oConect = Nothing
'Set oRHAreas = Nothing
'Set oEfect = Nothing
'
'Set GetRsOpeObj = rs1
'
'End Function

Public Function CargaListaCuentasOperacion(psOpeCod As String, psCtaCod As String) As String
On Error GoTo CargaListaCuentasOperacionErr
Dim lsCtas   As String
   Set dbConec = New COMConecta.DCOMConecta
   dbConec.AbreConexion
   psSql = "SELECT cCtaContCod FROM " & sCentralCom & "CtaCont WHERE cCtaContCod LIKE '" & psCtaCod & "_%' and cCtaContCod like '__[12]%'"
   Set prs = dbConec.CargaRecordSet(psSql)
   If Not prs.EOF Then
      Do While Not prs.EOF
         lsCtas = lsCtas & ",'" & prs!cCtaContCod & "'"
         prs.MoveNext
      Loop
   Else
      lsCtas = lsCtas & ",'" & psCtaCod & "'"
   End If
   CargaListaCuentasOperacion = Mid(lsCtas, 2, Len(lsCtas))
   dbConec.CierraConexion
   Set dbConec = Nothing
   Exit Function
CargaListaCuentasOperacionErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:CargaListaCuentasOperacion Method")
End Function

Public Function CargaListaDocsOperacion(psOpeCod As String) As String
On Error GoTo CargaListaDocsOperacionErr
Dim lsDocs   As String
   Set prs = CargaOpeDoc(psOpeCod)
   If prs.EOF Then
      Err.Raise 50001, "DOperacion", "No se asignaron Documentos a Operación"
      prs.Close: Set prs = Nothing
      Exit Function
   End If
   lsDocs = "'" & prs!nDocTpo & "'"
   prs.MoveNext
   Do While Not prs.EOF
      lsDocs = lsDocs & ",'" & prs!nDocTpo & "'"
      prs.MoveNext
   Loop
   CargaListaDocsOperacion = lsDocs
   Exit Function
CargaListaDocsOperacionErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:CargaListaDocsOperacion Method")
End Function

Public Function CargaOpeCtaUltimoNivel(psOpeCod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "") As Recordset
   On Error GoTo CargaOpeCtaUltimoNivelErr
   Set dbConec = New COMConecta.DCOMConecta
   Dim lsFiltroDH As String
   Dim lsFiltroOrden As String
    If psDH <> "" Then
        lsFiltroDH = " AND a.cOpeCtaDH ='" & psDH & "' "
    End If
    If psOpeCtaOrden <> "" Then
        lsFiltroOrden = " And a.cOpeCtaOrden='" & psOpeCtaOrden & "' "
    End If
   If dbConec.AbreConexion() Then
      psSql = "SELECT b.cCtaContCod, b.cCtaContDesc " _
         & "FROM " & sCentralCom & "OpeCta as a JOIN " & sCentralCom & "CtaCont as b ON b.cCtaContCod LIKE a.cCtaContCod+'%' " _
         & "WHERE a.cOpeCod LIKE '" & psOpeCod & "%' " & lsFiltroDH & lsFiltroOrden _
         & "ORDER BY b.cCtaContCod "
      Set CargaOpeCtaUltimoNivel = dbConec.CargaRecordSet(psSql, adLockReadOnly)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeCtaUltimoNivelErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DOperacion:CargaOpeCtaUltimoNivel Method")
End Function


Public Function GetFiltroObjetos(ByVal pnTipoObj As COMDConstantes.TpoObjetos, ByVal psCtaContCod As String, ByVal psObjetoCod As String, Optional lbMuestraCta As Boolean = True, Optional lbExisteFiltro As Boolean = False) As String
    On Error GoTo GetCtaObjFiltroErr
    Dim SQL As String
    Dim rs   As New ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    Dim lsPersCtaIF As String
    Dim lsCtaIf As String
    Dim lsTpoIf As String
    
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    GetFiltroObjetos = ""
    Select Case pnTipoObj
        Case ObjCMACAgencias
            SQL = "SELECT cAgeCod,  cSubCtaCod as SubCta  FROM Agencias where cAgeCod='" & psObjetoCod & "'"
        Case ObjCMACArea
            SQL = "SELECT cAreaCod, cSubCtaCod  as SubCta  FROM AREAS where cAreaCod='" & psObjetoCod & "'"
        Case ObjCMACAgenciaArea
           SQL = "     SELECT   AA.cAreaCod, AA.cAgeCod, " _
                & "             CASE " _
                & "                  WHEN CF.cSubctaCod IS NULL THEN AA.cSubCtaCod " _
                & "                  Else CF.cSubCtaCod END As SubCta " _
                & "    FROM      AREAAGENCIA AA " _
                & "             LEFT JOIN ( Select cAreaCod , cAgeCod , cSubCtaCod " _
                & "                          From CtaAreaAgeFiltro " _
                & "                          WHERE  cCtaContcod in ('" & psCtaContCod & "')) AS CF " _
                & "             ON CF.cAreacod = AA.cAreaCod and CF.cAgeCod = AA.cAgeCod " _
                & "    WHERE    AA.cAgeCod='" & Mid(psObjetoCod, 4, 2) & "' AND  AA.cAreaCod='" & Mid(psObjetoCod, 1, 3) & "'"
                
        Case ObjEntidadesFinancieras
            'If Len(psObjetoCod) > 15 Then
                lsTpoIf = Mid(psObjetoCod, 1, 2)
                lsPersCtaIF = Mid(psObjetoCod, 4, 13)
                lsCtaIf = Mid(psObjetoCod, 18, 10)
            'Else
            '    If psPersCodIf <> "" Then Exit Function
            '    lsPersCtaIF = psPersCodIf
            '    lsCtaIf = psObjetoCod
            'End If
            If Len(psObjetoCod) = 13 Then
                SQL = "select cPersCod, cSubCtaContcod SubCta FROM InstitucionFinanc where cPersCod = '" & psObjetoCod & "' "
            Else
                SQL = "Select cPersCod, cCtaIfCod, cCtaContCod, cCtaIFSubCta as SubCta From CtaIFFiltro WHERE  cPersCod = '" & lsPersCtaIF & "' and cIFTpo='" & lsTpoIf & "' AND cCtaIfCod='" & lsCtaIf & "' AND cCtaContCod in ('" & psCtaContCod & "')"
            End If
        Case Else
            SQL = "SELECT cCtaObjSubCta as SubCta FROM CtaObjFiltro WHERE cCtaContCod in ('" & psCtaContCod & "') and cObjetoCod = '" & psObjetoCod & "'"
    End Select
    Set rs = oConect.CargaRecordSet(SQL)
    If Not rs.EOF Then
        lbExisteFiltro = True
        GetFiltroObjetos = IIf(lbMuestraCta, psCtaContCod, "") & rs!SubCta
    End If
    If GetFiltroObjetos = "" Then
        GetFiltroObjetos = IIf(lbMuestraCta, psCtaContCod, "")
    End If
    Exit Function
GetCtaObjFiltroErr:
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GetCtaObjFiltro Method")
End Function

'DE ACA EN ADELANTE ES PARA EL MANTENIMIENTO DEL GRUPO DE OPERACIONES

Public Function InsertaGruposOpe(ByVal psOpeCod As String, ByVal psCodGrupoOpe As String) As Integer
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta

On Error GoTo InsertaGruposOpeErr

Set oCon = New COMConecta.DCOMConecta
InsertaGruposOpe = 1
SQL = "INSERT INTO GruposOpe(cGrupoCod, cOpeCod) Values('" & psCodGrupoOpe & "','" & psOpeCod & "')"

oCon.AbreConexion
oCon.Ejecutar SQL
oCon.CierraConexion

InsertaGruposOpe = 0
Set oCon = Nothing

Exit Function
InsertaGruposOpeErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function DeleteGruposOpe(ByVal psOpeCod As String, ByVal psCodGrupoOpe As String) As Integer
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta

On Error GoTo DeleteGruposOpeErr

Set oCon = New COMConecta.DCOMConecta
DeleteGruposOpe = 1
SQL = "DELETE GruposOpe WHERE cGrupoCod='" & psCodGrupoOpe & "' AND cOpeCod='" & psOpeCod & "'"

oCon.AbreConexion
oCon.Ejecutar SQL
oCon.CierraConexion

DeleteGruposOpe = 0
Set oCon = Nothing

Exit Function
DeleteGruposOpeErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function InsertaOpeGrupo(ByVal psGrupoCod As String, ByVal psGrupoNombre As String, ByVal pnOrden As Integer, _
                                ByVal psIngEgr As String, ByVal pnEfectivo As Integer, ByVal pnTipoGrupo As Integer, ByVal psOPSuma As String) As Integer

Dim SQL As String
Dim oCon As COMConecta.DCOMConecta

On Error GoTo InsertaOpeGrupoErr

Set oCon = New COMConecta.DCOMConecta
InsertaOpeGrupo = 1
SQL = " INSERT INTO GrupoOpe(cGrupoCod, cGrupoNombre, nOrden, cIngEgr, nEfectivo, nTipoGrupo, cOPSuma) " _
    & " VALUES('" & psGrupoCod & "','" & Replace(psGrupoNombre, "'", "''") & "'," & pnOrden & ",'" & psIngEgr & "'," _
    & pnEfectivo & "," & pnTipoGrupo & ",'" & psOPSuma & "')"

oCon.AbreConexion
oCon.Ejecutar SQL
oCon.CierraConexion

InsertaOpeGrupo = 0
Set oCon = Nothing

Exit Function
InsertaOpeGrupoErr:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
Public Function ActOpeGrupo(ByVal psGrupoCod As String, ByVal psGrupoNombre As String, ByVal pnOrden As Integer, _
                                ByVal psIngEgr As String, ByVal pnEfectivo As Integer, ByVal pnTipoGrupo As Integer, ByVal psOPSuma As String) As Integer

Dim SQL As String
Dim oCon As COMConecta.DCOMConecta

On Error GoTo ActOpeGrupoErr

Set oCon = New COMConecta.DCOMConecta
ActOpeGrupo = 1

SQL = " UPDATE  GrupoOpe " _
    & " SET     cGrupoNombre='" & Replace(psGrupoNombre, "'", "''") & "', " _
    & "         nOrden=" & pnOrden & "," _
    & "         cIngEgr='" & psIngEgr & "'," _
    & "         nEfectivo=" & pnEfectivo & "," _
    & "         nTipoGrupo=" & pnTipoGrupo & "," _
    & "         cOPSuma= '" & psOPSuma & "' " _
    & " WHERE  cGrupoCod='" & psGrupoCod & "'"

oCon.AbreConexion
oCon.Ejecutar SQL
oCon.CierraConexion

ActOpeGrupo = 0
Set oCon = Nothing

Exit Function
ActOpeGrupoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function DeleteOpeGrupo(ByVal psGrupoCod As String) As Integer
Dim SQL As String
Dim oCon As COMConecta.DCOMConecta

On Error GoTo DeleteOpeGrupoErr
Set oCon = New COMConecta.DCOMConecta

DeleteOpeGrupo = 1
SQL = " Delete GrupoOpe  WHERE  cGrupoCod='" & psGrupoCod & "'"

oCon.AbreConexion
oCon.Ejecutar SQL
oCon.CierraConexion

DeleteOpeGrupo = 0

Set oCon = Nothing
Exit Function
DeleteOpeGrupoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


'***********************************************
'DE ACA EN ADELANTE TODO LO QUE ERA EL nGrupoOpe

Public Function GetGrupoOpe() As ADODB.Recordset
Dim SQL As String
Dim rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta


If oConect.AbreConexion = False Then Exit Function

SQL = " Select  cGrupoCod, cGrupoNombre, nOrden, cIngEgr,nEfectivo,cOPSuma,cConsDescripcion, nTipoGrupo " _
    & " From    GrupoOpe G " _
    & "         JOIN CONSTANTE ON G.nTipoGrupo = nConsValor AND  nConsCod=1031 ORDER BY cGrupoCod  "

Set rs = oConect.CargaRecordSet(SQL)
Set GetGrupoOpe = rs
oConect.CierraConexion
Set oConect = Nothing

End Function

Public Function GetOpeGrupo(Optional ByVal psGrupo As String = "") As ADODB.Recordset
Dim SQL As String
Dim rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta

If oConect.AbreConexion = False Then Exit Function
SQL = " Select  O.cOpeCod, O.cOpeDesc " _
    & " From    Opetpo O JOIN GruposOpe GO ON GO.cOpeCod=O.cOpeCod " _
    & " Where   GO.cGrupoCod ='" & psGrupo & "'"

Set rs = oConect.CargaRecordSet(SQL)
Set GetOpeGrupo = rs
oConect.CierraConexion
Set oConect = Nothing

End Function

Public Function GetOperaciones() As ADODB.Recordset
Dim SQL As String
Dim rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta

If oConect.AbreConexion = False Then Exit Function
SQL = " Select  O.cOpeCod, O.cOpeDesc " _
    & " From    Opetpo o "
    

Set rs = oConect.CargaRecordSet(SQL)
Set GetOperaciones = rs
oConect.CierraConexion
Set oConect = Nothing

End Function

'Public Function GrabaOPeGrupo(ByVal psOpeCod As String, ByVal psCodGrupoOpe As String) As Integer
'Dim oOpeGrup As DGrupoOpe
'Set oOpeGrup = New DGrupoOpe
'oOpeGrup.InsertaGruposOpe psOpeCod, psCodGrupoOpe
'Set oOpeGrup = Nothing
'End Function

'Public Function EliminaGruposOpe(ByVal psOpeCod As String, ByVal psCodGrupoOpe As String) As Integer
'Dim oOpeGrup As DGrupoOpe
'Set oOpeGrup = New DGrupoOpe
'oOpeGrup.DeleteGruposOpe psOpeCod, psCodGrupoOpe
'Set oOpeGrup = Nothing
'End Function

'Public Function InsertaOpeGrupo(ByVal psGrupoCod As String, ByVal psGrupoNombre As String, ByVal pnOrden As Integer, _
'                                ByVal psIngEgr As String, ByVal pnEfectivo As Integer, ByVal pnTipoGrupo As Integer, ByVal psOPSuma As String) As Integer
'Dim oOpeGrup As DGrupoOpe
'Set oOpeGrup = New DGrupoOpe
'InsertaOpeGrupo = oOpeGrup.InsertaOpeGrupo(psGrupoCod, psGrupoNombre, pnOrden, psIngEgr, pnEfectivo, pnTipoGrupo, psOPSuma)
'Set oOpeGrup = Nothing
'End Function

'Public Function ActualizaOpeGrupo(ByVal psGrupoCod As String, ByVal psGrupoNombre As String, ByVal pnOrden As Integer, _
                                    ByVal psIngEgr As String, ByVal pnEfectivo As Integer, ByVal pnTipoGrupo As Integer, ByVal psOPSuma As String) As Integer
'Dim oOpeGrup As DGrupoOpe
'Set oOpeGrup = New DGrupoOpe
'ActualizaOpeGrupo = oOpeGrup.ActOpeGrupo(psGrupoCod, psGrupoNombre, pnOrden, psIngEgr, pnEfectivo, pnTipoGrupo, psOPSuma)
'Set oOpeGrup = Nothing
'End Function

'Public Function DeleteOpeGrupo(ByVal psGrupoCod As String) As Integer
'Dim oOpeGrup As DGrupoOpe
'Set oOpeGrup = New DGrupoOpe
'DeleteOpeGrupo = oOpeGrup.DeleteOpeGrupo(psGrupoCod)
'Set oOpeGrup = Nothing
'End Function

Public Function VerificaGruposOpe(ByVal psGrupoCod As String, ByVal psOpeCod As String) As Boolean
Dim SQL As String
Dim rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
VerificaGruposOpe = False
Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta

If oConect.AbreConexion = False Then Exit Function
SQL = " Select * From GruposOpe Where cGrupoCod='" & psGrupoCod & "' and cOpeCod='" & psOpeCod & "'"
Set rs = oConect.CargaRecordSet(SQL)
If Not rs.EOF And Not rs.BOF Then
    VerificaGruposOpe = True
End If
oConect.CierraConexion
Set oConect = Nothing

End Function

Public Function ObtenerValoresPonderados() As ADODB.Recordset
    Dim SQL As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As New ADODB.Recordset
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    SQL = "Select RTRIM(cDescripcion) + ' [' + Rtrim(Ltrim(Str(nValor))) + ']' cNomTab, nOpeGrupoCod cValor, nValor nRanIniTab From OpeTpoValoresPonderados"
    Set rs = oCon.CargaRecordSet(SQL)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerValoresPonderados = rs
    Else
        Set ObtenerValoresPonderados = Nothing
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function


Public Function ObtenerValoresPonderadosRelRangoFec(ByVal pdFechaI As Date, ByVal pdFechaF As Date, ByVal psParametro As String, ByVal psUsuario As String, Optional ByVal psPeso As String) As ADODB.Recordset
    Dim SQL As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As New ADODB.Recordset
    Dim oFunV As New COMFunciones.FCOMVarPublicas
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    If Trim(psPeso) = "" Then
        SQL = " Select Right(M.cMovNro,4) cCodUsu,  Count(Distinct Convert(varchar(25),M.cMovNro,113)) Num"
    Else
        SQL = " Select Right(M.cMovNro,4) cCodUsu,  Count(Distinct Convert(varchar(25),M.cMovNro,113)) * " & psPeso & " Num"
    End If
    SQL = SQL & " From Mov M" _
            & " Inner Join OpeTpoValoresPonderadosRel OP On M.cOpeCod = OP.cOpeCod" _
            & " Where   Left(M.cMovNro,8) Between '" & Format(CDate(pdFechaI), oFunV.gsFormatoMovFecha) & "' And '" & Format(DateAdd("d", 1, CDate(pdFechaF)), oFunV.gsFormatoMovFecha) & "' And (M.nMovFlag = " & MovFlag.gMovFlagVigente & ")" _
            & "          " & IIf(Right(psParametro, 4) = "9999", "  And OP.nOpeGrupoCod <> 0", "  And OP.nOpeGrupoCod = " & Trim(Right(psParametro, 4))) & " And Right(M.cMovNro,4) In (" & psUsuario & ")" _
            & " Group By Right(M.cMovNro,4) Order by cCodUsu"

    Set rs = oCon.CargaRecordSet(SQL)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerValoresPonderadosRelRangoFec = rs
    Else
        Set ObtenerValoresPonderadosRelRangoFec = Nothing
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ObtenerValoresPonderadosRelFec(ByVal pdFechaI As Date, ByVal lnMes As String, ByVal psParametro As String, ByVal psUsuario As String, Optional ByVal psPeso As String) As ADODB.Recordset
    Dim SQL As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As New ADODB.Recordset
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    If Trim(psPeso) = "" Then
        SQL = " Select Right(M.cMovNro,4) cCodUsu,  Count(Distinct Convert(varchar(25),M.cMovNro,113)) Num"
    Else
        SQL = " Select Right(M.cMovNro,4) cCodUsu,  Count(Distinct Convert(varchar(25),M.cMovNro,113)) * " & psPeso & " Num"
    End If
    SQL = SQL & " From Mov M" _
              & " Inner Join OpeTpoValoresPonderadosRel OP On M.cOpeCod = OP.cOpeCod" _
              & " Where   Left(M.cMovNro,6) = '" & Format(DateAdd("m", lnMes, CDate(pdFechaI)), "yyyymm") & "' And (M.nMovFlag = " & MovFlag.gMovFlagVigente & ")" _
              & "          " & IIf(Right(psParametro, 4) = "9999", "  And OP.nOpeGrupoCod <> 0", "  And OP.nOpeGrupoCod = " & Trim(Right(psParametro, 4))) & " And Right(M.cMovNro,4) In (" & psUsuario & ")" _
              & " Group By Right(M.cMovNro,4) Order by cCodUsu"

    Set rs = oCon.CargaRecordSet(SQL)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerValoresPonderadosRelFec = rs
    Else
        Set ObtenerValoresPonderadosRelFec = Nothing
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function



