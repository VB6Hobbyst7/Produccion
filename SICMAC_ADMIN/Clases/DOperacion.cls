VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOperacion"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sConexion As String
Dim sCentralCom As String

Public Function CargaOpeGru(Optional psOpeGruCod As String, Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaOpeGruErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "SELECT cOpeGruCod, cOpeGruDesc,1 as Nivel FROM " & sCentralCom & "OpeGrupo " & IIf(psOpeGruCod = "", "", "WHERE cOpeGruCod = '" & psOpeGruCod & "'")
      Set CargaOpeGru = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Sub InsertaOpeGru(psOpeGruCod As String, psOpeGruDesc As String, psMovNro As String)
   On Error GoTo InsertaOpeErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeGrupo (cOpeGruCod, cOpeGruDesc, cUltimaActualizacion) VALUES ('" & psOpeGruCod & "', '" & psOpeGruDesc & "', '" & psMovNro & "') "
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaOpeErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeGru Method")
End Sub

Public Sub ActualizaOpeGru(psOpeGruCod As String, psOpeGruDesc As String, psMovNro As String)
   On Error GoTo InsertaOpeErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "UPDATE " & sCentralCom & "OpeGrupo SET cOpeGruDesc = '" & psOpeGruDesc & "', cUltimaActualizacion = '" & psMovNro & "' WHERE cOpeGruCod = '" & psOpeGruCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaOpeErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeGru Method")
End Sub

Public Sub EliminaOpeGru(psOpeGruCod As String)
   On Error GoTo InsertaOpeErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeGrupo WHERE cOpeGruCod = '" & psOpeGruCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaOpeErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeGru Method")
End Sub

Public Function CargaOpeTpo(psOpeCod As String, Optional plLike As Boolean = False, Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaOpeTpoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "SELECT cOpeCod,cOpeDesc, CASE WHEN cOpeVisible='" & gOpeTpoVisibleSi & "' THEN '" & gsSI & "' ELSE '" & gsNO & "' END cOpeVisible,nOpeNiv, cOpeGruCod FROM " & sCentralCom & "OpeTpo" & IIf(psOpeCod = "", "", " WHERE cOpeCod " & IIf(plLike, " LIKE '", "= '") & psOpeCod & IIf(plLike, "%'", "'"))
      Set CargaOpeTpo = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeTpoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeTpo Method")
End Function

Public Sub InsertaOpeTpo(psOpeCod As String, psOpeDesc As String, psOpeVisible As String, pnOpeNiv As Integer, psOpeTpo As String, psMovNro As String)
   On Error GoTo InsertaOpeTpoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeTpo (cOpeCod, cOpeDesc, cOpeVisible, nOpeNiv, cOpeGruCod, cUltimaActualizacion) VALUES ('" & psOpeCod & "', '" & Replace(psOpeDesc, "'", "''") & "', '" & psOpeVisible & "', " & pnOpeNiv & ", '" & psOpeTpo & "','" & psMovNro & "') "
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaOpeTpoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeTpo Method")
End Sub

Public Sub ActualizaOpeTpo(psOpeCod As String, psOpeDesc As String, psOpeVisible As String, pnOpeNiv As Integer, psOpeTpo As String, psMovNro As String)
   On Error GoTo ActualziaOpeTpoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "UPDATE " & sCentralCom & "OpeTpo SET cOpeDesc = '" & psOpeDesc & "', cOpeVisible = '" & psOpeVisible & "', nOpeNiv = " & pnOpeNiv & ", cOpeGruCod = '" & psOpeTpo & "', cUltimaActualizacion = '" & psMovNro & "' " _
            & "WHERE cOpeCod = '" & psOpeCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualziaOpeTpoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:ActualziaOpeTpo Method")
End Sub

Public Sub EliminaOpeTpo(psOpeCod As String)
   On Error GoTo EliminaOpeTpoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeDoc WHERE cOpeCod = '" & psOpeCod & "' "
      dbConec.AdicionaCmdBatch psSql
      psSql = "DELETE " & sCentralCom & "OpeObj WHERE cOpeCod = '" & psOpeCod & "' "
      dbConec.AdicionaCmdBatch psSql
      psSql = "DELETE " & sCentralCom & "OpeCta WHERE cOpeCod = '" & psOpeCod & "' "
      dbConec.AdicionaCmdBatch psSql
      psSql = "DELETE " & sCentralCom & "OpeTpo WHERE cOpeCod = '" & psOpeCod & "' "
      dbConec.AdicionaCmdBatch psSql
      dbConec.EjecutarBatch
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaOpeTpoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:EliminaOpeTpo Method")
End Sub

Public Function CargaOpeCta(psOpeCod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "", Optional LockType As LockTypeEnum = adLockReadOnly, Optional pbGetClase As Boolean = False) As Recordset
   On Error GoTo CargaOpeCtaErr
   Set dbConec = New DConecta
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
         & "  a.cCtaContCod, b.cCtaContDesc,   1 as Nivel ,  " _
         & "  CASE WHEN a.cOpeCtaDH = '" & gsContHaber & "' THEN '" & gsContHaberDesc & "' " _
         & "       WHEN a.cOpeCtaDH = '" & gsContDebe & "' THEN '" & gsContDebeDesc & "' END  as Clase, " _
         & "  CASE WHEN a.cOpeCtaOpc = '" & gOpeCtaCaracObligatorio & "' then '" & gsOpeCtaCaracterObligaDesc & "' WHEN a.cOpeCtaOpc='" & gOpeCtaCaracObligatorio & "' then '" & gsOpeCtaCaracterOpcionDesc & "' END as Caracter, " _
         & "  a.cOpeCod, a.cOpeCtaDH, a.cOpeCtaOpc, a.cCtaContN, a.cOpeCtaOrden " & IIf(pbGetClase, ", cls.cCtaCaracter ", "") _
         & "  FROM OpeCta as a JOIN CtaCont as b ON b.cCtaContCod = a.cCtaContCod " _
         & IIf(pbGetClase, " JOIN CtaContClase cls ON a.cCtaContCod LIKE cls.cCtaContCod + '%' ", "") _
         & " where a.cOpeCod = '" & psOpeCod & "' " & lsFiltroDH & lsFiltroOrden _
         & " order by a.cOpeCtaDH"
      Set CargaOpeCta = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeCta Method")
End Function

Public Sub InsertaOpeCta(psOpeCod As String, psOpeCtaOrden As String, psCtaCod As String, psOpeCtaDH As String, psOpeCtaOpc As String, psCtaCodN As String, psMovNro As String)
   On Error GoTo InsertaOpeCtaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeCta (cOpeCod, cOpeCtaOrden , cCtaContCod, cOpeCtaDH, cOpeCtaOpc, cCtaContN, cUltimaActualizacion) VALUES ('" & psOpeCod & "', '" & psOpeCtaOrden & "','" & psCtaCod & "','" & psOpeCtaDH & "','" & psOpeCtaOpc & "', '" & psCtaCodN & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeCta Method")
End Sub

Public Sub EliminaOpeCta(psOpeCod As String, psOpeCtaOrden As String, psCtaCod As String, psOpeCtaDH As String)
   On Error GoTo EliminaOpeCtaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeCta WHERE cOpeCod = '" & psOpeCod & "' and cOpeCtaOrden = '" & psOpeCtaOrden & "' and cCtaContCod = '" & psCtaCod & "' and cOpeCtaDH = '" & psOpeCtaDH & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:EliminaOpeCta Method")
End Sub

Public Function CargaOpeObj(psOpeCod As String, Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaOpeObjErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "SELECT a.cOpeObjOrden ,a.cObjetoCod, b.cObjetoDesc, a.cOpeObjValor, " & _
         "       a.nOpeObjNiv, a.cOpeObjFiltro, a.cOpeCod" & _
         "  FROM  " & sCentralCom & "OpeObj as a JOIN " & sCentralCom & "Objeto as b ON b.cObjetoCod = a.cObjetoCod " & _
         " WHERE a.cOpeCod = '" & psOpeCod & "'"
      Set CargaOpeObj = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeObjErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeObj Method")
End Function

Public Sub InsertaOpeObj(psOpeCod As String, psOpeObjOrden As String, psObjetoCod As String, psOpeObjValor As String, pnOpeObjNiv As Integer, psOpeObjFiltro As String, psMovNro As String)
   On Error GoTo InsertaOpeObjErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeObj (cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjValor, nOpeObjNiv, cOpeObjFiltro, cUltimaActualizacion) " _
            & "VALUES ('" & psOpeCod & "','" & psOpeObjOrden & "','" & psObjetoCod & "','" & psOpeObjValor & "'," & pnOpeObjNiv & ",'" & psOpeObjFiltro & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaOpeObjErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeObj Method")
End Sub

Public Sub EliminaOpeObj(psOpeCod As String, psOpeObjOrden As String, psObjetoCod As String)
   On Error GoTo EliminaOpeObjErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeObj WHERE cOpeCod = '" & psOpeCod & "' and cOpeObjOrden = '" & psOpeObjOrden & "' and cObjetoCod = '" & psObjetoCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaOpeObjErr:
   Call RaiseError(MyUnhandledError, "DOperacion:EliminaOpeObj Method")
End Sub

Public Function CargaOpeDoc(psOpeCod As String, Optional psMetodo As String, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
   On Error GoTo CargaOpeDocErr
   Set dbConec = New DConecta
   Dim lsFiltroMet As String
   If psMetodo <> "" Then
        lsFiltroMet = " AND a.cOpeDocMetodo ='" & psMetodo & "'"
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
       "WHERE cOpeCod = '" & psOpeCod & "' " & lsFiltroMet
       
      Set CargaOpeDoc = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeDocErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeDoc Method")
End Function

Public Function CargaOpeDocEstado(psOpeCod As String, Optional psMetodo As String, Optional psOpeEstado As String, Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaOpeDocErr
   Set dbConec = New DConecta
   Dim lsFiltroMet As String
   Dim lsFiltroEstado As String
   If psMetodo <> "" Then
        lsFiltroMet = " AND a.cOpeDocMetodo ='" & psMetodo & "'"
   End If
   If psOpeEstado <> "" Then
        lsFiltroEstado = " AND a.cOpeDocEstado like '" & psOpeEstado & "'"
   End If
   
   If dbConec.AbreConexion() Then
      psSql = " SELECT a.cOpeCod, a.nDocTpo, d.cDocDesc, " & _
       " CASE WHEN a.cOpeDocEstado = '11'  THEN 'Obligatorio - Debe Existir' " & _
       "      WHEN a.cOpeDocEstado = '12'  THEN 'Obligatorio - No debe Existir' " & _
       "      WHEN a.cOpeDocEstado = '21'  THEN 'Opcional - Debe Existir' " & _
       "      WHEN a.cOpeDocEstado = '22'  THEN 'Opcional - No debe Existir' END as Estado, " & _
       " CASE WHEN a.cOpeDocMetodo = '1' THEN 'Nro. de Movimiento' " & _
       "      WHEN a.cOpeDocMetodo = '2' THEN 'Autogenerado' " & _
       "      WHEN a.cOpeDocMetodo = '3' THEN 'Digitado' END as Metodo , D.cDocAbrev " & _
       " FROM " & sCentralCom & "OpeDoc a JOIN " & sCentralCom & "Documento d ON d.nDocTpo = a.nDocTpo " & _
       " WHERE cOpeCod = '" & psOpeCod & "' " & lsFiltroMet & lsFiltroEstado
       
      Set CargaOpeDocEstado = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeDocErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeDoc Method")
End Function

Public Sub InsertaOpeDoc(psOpeCod As String, psDocTpo As String, psOpeDocEstado As String, psOpeDocMetodo As String, psMovNro As String)
   On Error GoTo InsertaOpeDocErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeDoc (cOpeCod, nDocTpo, cOpeDocEstado, cOpeDocMetodo, cUltimaActualizacion) VALUES ('" & psOpeCod & "', '" & psDocTpo & "', '" & psOpeDocEstado & "','" & psOpeDocMetodo & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaOpeDocErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeDoc Method")
End Sub

Public Sub EliminaOpeDoc(psOpeCod As String, psDocTpo As String)
   On Error GoTo EliminaOpeDocErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeDoc WHERE cOpeCod = '" & psOpeCod & "' and nDocTpo = '" & psDocTpo & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaOpeDocErr:
   Call RaiseError(MyUnhandledError, "DOperacion:EliminaOpeDoc Method")
End Sub

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

Dim oIni As New ClsIni.ClasIni
sConexion = oIni.CadenaConexion
sCentralCom = oIni.BaseComunes
Set oIni = Nothing
End Sub
' Parches
Public Function EmiteDocOpe(ByVal psOpeCod As String, ByVal psOpeDocEstado As OpeDocEstado, ByVal psOpeDocMetodo As OpeDocMetodo) As String
On Error GoTo ErrorEmiteDocOpe
Dim rs As ADODB.Recordset
Dim sSql As String
Dim lsFiltroOpeDocEstado As String
Dim lsFiltroOpeDocMetodo As String
Set rs = New ADODB.Recordset


lsFiltroOpeDocEstado = "  AND cOpeDocEstado ='" & psOpeDocEstado & "'  "
lsFiltroOpeDocMetodo = "  AND cOpeDocMetodo ='" & psOpeDocMetodo & "'  "
Dim oconect As DConecta
Set oconect = New DConecta
EmiteDocOpe = ""
If oconect.AbreConexion = False Then Exit Function
sSql = "Select nDocTpo from " & sCentralCom & "OpeDoc where cOpeCod = '" & psOpeCod & "' " & lsFiltroOpeDocEstado & lsFiltroOpeDocMetodo
Set rs = oconect.CargaRecordSet(sSql)
If Not rs.EOF And Not rs.BOF Then
    EmiteDocOpe = rs!nDocTpo
End If
rs.Close: Set rs = Nothing
oconect.CierraConexion
Set oconect = Nothing
Exit Function
ErrorEmiteDocOpe:
    Err.Raise vbObjectError + 100, "ErrorEmiteDocOpe", "Error emitir documento en Operacion"
End Function
Public Function EmiteOpeDoc(ByVal psOpeCod As String, ByVal psDocTpo As TpoDoc, Optional ByVal psDH As String = "") As String
On Error GoTo ErrorEmiteOpeDoc
Dim rs As ADODB.Recordset
Dim sSql As String
Dim lsFiltroDH As String
Set rs = New ADODB.Recordset

Dim oconect As DConecta
Set oconect = New DConecta
EmiteOpeDoc = ""
If oconect.AbreConexion = False Then Exit Function
        
If psDH <> "" Then
    lsFiltroDH = " AND OC.cOpeCtaDH ='" & psDH & "'"
End If

sSql = "Select OC.COPECOD, OD.nDocTpo  " _
    & " From   " & sCentralCom & "OpeCta OC Left Join " & sCentralCom & "OpeDoc OD on OD.cOpeCod = OC.cOpeCod  " _
    & " Where  Substring(OC.cOpeCod,1,LEN('" & psOpeCod & "')) ='" & Mid(psOpeCod, 1, Len(psOpeCod)) & "' and OD.nDocTpo" & IIf(psDocTpo = -1, " IS NULL ", "='" & psDocTpo & "' ") _
    & lsFiltroDH _
    & " GROUP BY OC.COPECOD, OD.nDocTpo  "
    
Set rs = oconect.CargaRecordSet(sSql)
If Not rs.EOF And Not rs.BOF Then
    EmiteOpeDoc = rs!COPECOD
End If
rs.Close: Set rs = Nothing
oconect.CierraConexion
Set oconect = Nothing
Exit Function
ErrorEmiteOpeDoc:
    Err.Raise vbObjectError + 100, "ErrorEmiteDocOpe", "Error emitir documento en Operacion"
End Function
Public Function EmiteOpeCta(ByVal psOpeCod As String, ByVal psDH As String, Optional psOpeCtaOrden As String = "0", Optional ByVal psObjetoCod As String = "", Optional pnFiltro As TpoObjetos, Optional lbMuestraSubCta As Boolean = True) As String
    On Error GoTo EmiteOpeCodErr
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oconect As DConecta
    Dim lsListaCtasCont As String
    Dim lsSubCta As String
    Dim oCtaIf As NCajaCtaIF
    Dim oCtaCont As DCtaCont
    Dim oContFunc As NContFunciones
    
    Set oCtaCont = New DCtaCont
    Set oCtaIf = New NCajaCtaIF
    Set oContFunc = New NContFunciones
    
    Set oconect = New DConecta
    Set rs = New ADODB.Recordset
    Dim lnNumCtas As Integer
    Dim lsNuevaCta As String
    Dim lsCuenta As String
    Dim lsAux As String
    If oconect.AbreConexion = False Then Exit Function
    
    sSql = " SELECT  OC.cCtaContCod , OC.cOpeCtaOrden, OC.cOpeCtaDH, CC.cCtaContDesc " _
        & " FROM     " & sCentralCom & "OPECTA OC  JOIN CTACONT CC ON CC.CCTACONTCOD = OC.CCTACONTCOD   " _
        & " WHERE    OC.cOpeCod ='" & psOpeCod & "'  AND OC.cOpeCtaDH = '" & psDH & "' AND OC.cOpeCtaOrden = '" & psOpeCtaOrden & "'"
    lnNumCtas = 0
    Set rs = oconect.CargaRecordSet(sSql)
    lnNumCtas = rs.RecordCount
    Do While Not rs.EOF
        lsListaCtasCont = rs!cCtaContCod
        If lsListaCtasCont <> "" Then
            'lsListaCtasCont = Mid(lsListaCtasCont, 1, Len(lsListaCtasCont) - 3)
            lsNuevaCta = lsListaCtasCont
            lsSubCta = EmiteSubCtaOpeObj(psOpeCod, lsListaCtasCont)
            lsAux = ""
            If lsSubCta = "" Then
                If psObjetoCod <> "" Or lsSubCta <> "" Then
                    lsSubCta = ""
                    lsAux = oContFunc.GetFiltroObjetos(pnFiltro, lsListaCtasCont, psObjetoCod, lbMuestraSubCta)
                Else
                    lsAux = IIf(lbMuestraSubCta, lsListaCtasCont, "")
                End If
            Else
                lsAux = IIf(lbMuestraSubCta, lsListaCtasCont, "") & lsSubCta
            End If
            If lsNuevaCta = lsAux And lnNumCtas > 1 And lsSubCta = "" Then
                lsAux = ""
            Else
                lsCuenta = lsAux
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    oconect.CierraConexion
    Set oconect = Nothing
    EmiteOpeCta = lsCuenta
'    If lsListaCtasCont <> "" Then
'        lsListaCtasCont = Mid(lsListaCtasCont, 1, Len(lsListaCtasCont) - 3)
'        lsNuevaCta = lsListaCtasCont
'        rs.Close: Set rs = Nothing
'        oConect.CierraConexion
'        Set oConect = Nothing
'        lsSubCta = EmiteSubCtaOpeObj(psOpeCod, lsListaCtasCont)
'        If lsSubCta = "" Then
'            If psObjetoCod <> "" Or lsSubCta <> "" Then
'                lsSubCta = ""
'                Select Case pnFiltro
'                    Case CtaOBjFiltroIF
'                            EmiteOpeCta = oCtaIf.GetCtaIfFiltro(lsListaCtasCont, Mid(psObjetoCod, 15, 10), Mid(psObjetoCod, 1, 13), lbMuestraSubCta)
'                    Case CtaOBjFiltroObj
'                            EmiteOpeCta = oCtaCont.GetCtaObjFiltro(lsListaCtasCont, psObjetoCod, lbMuestraSubCta)
'                    Case Else
'                            EmiteOpeCta = oContFunc.GetFiltroObjetos(pnFiltro, lsListaCtasCont, psObjetoCod, , lbMuestraSubCta)
'                        'Exit Function
'                End Select
'            Else
'                EmiteOpeCta = IIf(lbMuestraSubCta, lsListaCtasCont, "")
'            End If
'        Else
'            EmiteOpeCta = IIf(lbMuestraSubCta, lsListaCtasCont, "") & lsSubCta
'        End If
'        If lsNuevaCta = EmiteOpeCta And lnNumCtas > 1 Then
'            EmiteOpeCta = ""
'        End If
'    End If
    Set oCtaIf = Nothing
    Set oCtaCont = Nothing
    
    Exit Function
EmiteOpeCodErr:
    Call RaiseError(MyUnhandledError, "DMov:EmiteOpeCod Method")
End Function
Public Function EmiteSubCtaOpeObj(ByVal psOpeCod As String, ByVal psCtaContCod As String) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta
Set oconect = New DConecta
Set rs = New ADODB.Recordset
Dim lsSubCta As String
lsSubCta = ""
If oconect.AbreConexion = False Then Exit Function
'SQL = "Select  cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjFiltro, " _
    & " CASE " _
    & "         WHEN OO.COBJETOCOD ='" & ObjEntidadesFinancieras & "' THEN " _
    & "                 ISNULL((SELECT cCtaIFSubCta FROM CTAIFFILTRO WHERE (cPersCod = OO.cOpeObjFiltro OR cCtaIfCod = OO.cOpeObjFiltro ) AND cCtaContCod in ('" & psCtaContCod & "')),'') " _
    & "         WHEN OO.COBJETOCOD ='" & ObjCMACAgenciaArea & "' THEN " _
    & "                 ISNULL((SELECT cSubCtaCod FROM CtaAreaAgeFiltro WHERE (cAreaCod LIKE OO.cOpeObjFiltro OR cAgeCod LIKE OO.cOpeObjFiltro) AND cCtaContCod in ('" & psCtaContCod & "')) , " _
    & "                 (SELECT cSubCtaCod FROM AREAAGENCIA WHERE (cAreaCod LIKE OO.cOpeObjFiltro OR cAgeCod LIKE OO.cOpeObjFiltro) )) " _
    & "         WHEN OO.COBJETOCOD ='" & ObjCMACAgencias & "' THEN " _
    & "                 ISNULL((Select cSubCtaCod from AGENCIAS WHERE cAgeCod LIKE OO.cOpeObjFiltro ),'') " _
    & "         WHEN OO.COBJETOCOD ='" & ObjCMACArea & "' THEN " _
    & "                 ISNULL((SELECT cSubCtaCod FROM AREAS WHERE cAreaCod LIKE OO.cOpeObjFiltro ),'') " _
    & "         Else " _
    & "                 (SELECT cCtaObjSubCta FROM CTAOBJFILTRO WHERE COBJETOCOD LIKE OO.cOpeObjFiltro AND cCtaContCod in ('" & psCtaContCod & "'))  END as SUBCUENTA " _
    & " From    OpeObj OO " _
    & " Where   OO.cOpecod Like '" & psOpeCod & "' " _
    & " ORDER BY OO.COPEOBJORDEN "

sql = "Select  cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjFiltro, " _
    & " CASE " _
    & "         WHEN OO.COBJETOCOD ='" & ObjEntidadesFinancieras & "' THEN " _
    & "                 ISNULL((SELECT cCtaIFSubCta FROM CTAIFFILTRO WHERE (cPersCod = OO.cOpeObjFiltro OR cCtaIfCod = OO.cOpeObjFiltro ) AND cCtaContCod in ('" & psCtaContCod & "')),'') " _
    & "         WHEN OO.COBJETOCOD ='" & ObjCMACAgenciaArea & "' THEN " _
    & "                 ISNULL((SELECT cSubCtaCod FROM CtaAreaAgeFiltro WHERE cAreaCod + cAgeCod LIKE OO.cOpeObjFiltro AND cCtaContCod in ('" & psCtaContCod & "')) , '') " _
    & "         Else " _
    & "                 (SELECT cCtaObjSubCta FROM CTAOBJFILTRO WHERE COBJETOCOD LIKE OO.cOpeObjFiltro AND cCtaContCod in ('" & psCtaContCod & "'))  END as SUBCUENTA " _
    & " From    OpeObj OO " _
    & " Where   OO.cOpecod Like '" & psOpeCod & "' " _
    & " ORDER BY OO.COPEOBJORDEN "

Set rs = oconect.CargaRecordSet(sql)
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
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta

Set rs = New ADODB.Recordset
Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function

sql = " SELECT   cOpeCtaDH " _
    & " FROM    " & sCentralCom & "OPECTA OC " _
    & " WHERE    COPECOD = '" & psOpeCod & "' and cCtaContcod ='" & psCtaContCod & "' and   OC.cOpeCtaOrden='" & psOpeCtaOrden & "' "
GetOpeCtaDebeHaber = ""
Set rs = oconect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetOpeCtaDebeHaber = rs!cOpeCtaDH
End If
rs.Close
Set rs = Nothing
oconect.CierraConexion: Set oconect = Nothing
End Function
Public Function EmiteOpeCtasNivel(ByVal psOpeCod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "0") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta
Dim lsDHFiltro As String
Dim lsOrdenFiltro As String
If psDH <> "" Then
    lsDHFiltro = " AND cOpeCtaDH ='" & psDH & "'"
End If
If psOpeCtaOrden <> "" Then
   lsOrdenFiltro = " and OC.cOpeCtaOrden='" & psOpeCtaOrden & "'"
End If

sql = "SELECT   C.cCtaContCod , C.cCtaContDesc, 1 as Nivel " _
    & "FROM     " & sCentralCom & "OPECTA OC " _
    & "         JOIN     " & sCentralCom & "CTACONT C ON C.cCtaContCod = OC.cCtaContCod " _
    & " WHERE   COPECOD = '" & psOpeCod & "' " & lsDHFiltro & lsOrdenFiltro

Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function

Set rs = oconect.CargaRecordSet(sql)
Set EmiteOpeCtasNivel = rs
oconect.CierraConexion: Set oconect = Nothing
End Function

Public Function GetOpeObj(ByVal psOpeCod As String, ByVal psOpeObjOrden As String, _
                        Optional psCtaContCod As String = "", _
                        Optional ByVal psFiltro As String, Optional ByVal psFiltroAdd As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim oconect As DConecta
Dim oRHAreas As DActualizaDatosArea
Dim oContFunct As NContFunciones
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo
Dim lsFiltro As String
Dim lsFiltroIf As String
Dim lsFiltroCta As String

Set oCtaIf = New NCajaCtaIF
Set oContFunct = New NContFunciones
Set oconect = New DConecta
Set oRHAreas = New DActualizaDatosArea
Set oEfect = New Defectivo

Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
If oconect.AbreConexion = False Then
    Exit Function
End If
sql = " SELECT cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjFiltro, nOpeObjNiv " _
    & " FROM " & sCentralCom & "OPEOBJ " _
    & " WHERE COPECOD ='" & psOpeCod & "' and cOpeObjOrden='" & psOpeObjOrden & "'"
    
Set rs = oconect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    If psFiltro = "" Then
        lsFiltro = Trim(rs!cOpeObjFiltro) & psFiltroAdd
    Else
        lsFiltro = psFiltro & psFiltroAdd
    End If
    Select Case Val(rs!cObjetoCod)
            Case ObjCMACAgencias
                Set rs1 = oRHAreas.GetAgencias(lsFiltro)
            Case ObjCMACAgenciaArea
                Set rs1 = oRHAreas.GetAgenciasAreas(lsFiltro)
            Case ObjCMACArea
                Set rs1 = oRHAreas.GetAreas(lsFiltro)
            Case ObjEntidadesFinancieras
                lsFiltroIf = lsFiltro
                lsFiltroCta = lsFiltro
                'Set rs1 = oCtaIf.GetCtasInstFinancieras(lsFiltro, psCtaContCod)
                Set rs1 = oCtaIf.CargaCtasIF(Mid(psOpeCod, 3, 1), lsFiltro, rs!nOpeObjNiv)
            Case ObjDescomEfectivo
                Set rs1 = oEfect.GetBilletajes(lsFiltro)
            Case ObjPersona
                Set rs1 = Nothing
            Case Else
                'Set rs1 = oContFunct.GetObjetos(Val(IIf(lsFiltro = "", rs!cObjetoCod, rs!cOpeObjFiltro)), IIf(rs!cOpeObjFiltro = "", true, false))
                Set rs1 = oContFunct.GetObjetos(Val(IIf(lsFiltro = "", rs!cObjetoCod, rs!cOpeObjFiltro)), False)
        End Select
End If
rs.Close
Set rs = Nothing
oconect.CierraConexion: Set oconect = Nothing

Set oCtaIf = Nothing
Set oContFunct = Nothing
Set oconect = Nothing
Set oRHAreas = Nothing
Set oEfect = Nothing

Set GetOpeObj = rs1

End Function
Public Function GetRsOpeObj(ByVal psOpeCod As String, ByVal psOpeObjOrden As String, Optional ByVal psFiltroObj As String, Optional ByVal psFiltroAdd As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim oconect As DConecta
Dim oRHAreas As DActualizaDatosArea
Dim oContFunct As NContFunciones
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo
Dim lsFiltro As String
Dim lsFiltroIf As String
Dim lsFiltroCta As String

Set oCtaIf = New NCajaCtaIF
Set oContFunct = New NContFunciones
Set oconect = New DConecta
Set oRHAreas = New DActualizaDatosArea
Set oEfect = New Defectivo

Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
If oconect.AbreConexion = False Then
    Exit Function
End If
sql = " SELECT cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjFiltro, nOpeObjNiv " _
    & " FROM " & sCentralCom & "OPEOBJ " _
    & " WHERE COPECOD ='" & psOpeCod & "' and cOpeObjOrden='" & psOpeObjOrden & "'"
    
Set rs = oconect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    If psFiltroObj = "" Then
        lsFiltro = Trim(rs!cOpeObjFiltro) & psFiltroAdd
    Else
        lsFiltro = psFiltroObj & psFiltroAdd
    End If
    Select Case Val(rs!cObjetoCod)
            Case ObjCMACAgencias
                Set rs1 = oRHAreas.GetAgencias(lsFiltro)
            Case ObjCMACAgenciaArea
                Set rs1 = oRHAreas.GetAgenciasAreas(lsFiltro)
            Case ObjCMACArea
                Set rs1 = oRHAreas.GetAreas(lsFiltro)
            Case ObjEntidadesFinancieras
                Set rs1 = oCtaIf.CargaCtasIF(Mid(psOpeCod, 3, 1), lsFiltro, rs!nOpeObjNiv)
            Case ObjDescomEfectivo
                Set rs1 = oEfect.GetBilletajes(lsFiltro)
            Case ObjPersona
                Set rs1 = Nothing
            Case Else
                Set rs1 = oContFunct.GetObjetos(Val(IIf(lsFiltro = "", rs!cObjetoCod, rs!cOpeObjFiltro)), IIf(rs!cOpeObjFiltro = "", True, False))
        End Select
End If
rs.Close
Set rs = Nothing
oconect.CierraConexion: Set oconect = Nothing

Set oCtaIf = Nothing
Set oContFunct = Nothing
Set oconect = Nothing
Set oRHAreas = Nothing
Set oEfect = Nothing

Set GetRsOpeObj = rs1

End Function


