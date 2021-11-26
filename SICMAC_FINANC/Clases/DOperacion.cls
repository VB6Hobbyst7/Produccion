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
Dim dbconec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sCentralCom As String

Public Function CargaOpeGru(Optional psOpeGruCod As String, Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaOpeGruErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "SELECT cOpeGruCod, cOpeGruDesc, 1 as Nivel FROM " & sCentralCom & "OpeGrupo " & IIf(psOpeGruCod = "", "", "WHERE cOpeGruCod = '" & psOpeGruCod & "'")
      Set CargaOpeGru = dbconec.CargaRecordSet(psSql, LockType)
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Sub InsertaOpeGru(psOpeGruCod As String, psOpeGruDesc As String, psMovNro As String)
   On Error GoTo InsertaOpeErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeGrupo (cOpeGruCod, cOpeGruDesc, cUltimaActualizacion) VALUES ('" & psOpeGruCod & "', '" & psOpeGruDesc & "', '" & psMovNro & "') "
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
InsertaOpeErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeGru Method")
End Sub

Public Sub ActualizaOpeGru(psOpeGruCod As String, psOpeGruDesc As String, psMovNro As String)
   On Error GoTo InsertaOpeErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "UPDATE " & sCentralCom & "OpeGrupo SET cOpeGruDesc = '" & psOpeGruDesc & "', cUltimaActualizacion = '" & psMovNro & "' WHERE cOpeGruCod = '" & psOpeGruCod & "'"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
InsertaOpeErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeGru Method")
End Sub

Public Sub EliminaOpeGru(psOpeGruCod As String)
   On Error GoTo InsertaOpeErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeGrupo WHERE cOpeGruCod = '" & psOpeGruCod & "'"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
InsertaOpeErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeGru Method")
End Sub

Public Function CargaOpeTpo(psOpecod As String, Optional plLike As Boolean = False, Optional LockType As LockTypeEnum = adLockReadOnly, Optional pnVisible As Integer = -1, Optional pnNivel As Integer = 0, Optional pssqlAddOrOPTPO As String = "") As Recordset
Dim lsFiltro As String
   On Error GoTo CargaOpeTpoErr
   lsFiltro = ""
   If pnVisible <> -1 Then
        lsFiltro = " and cOpeVisible = '" & pnVisible & "'"
   End If
   If pnNivel > 0 Then
        lsFiltro = lsFiltro & " and nOpeNiv = " & pnNivel & " "
   End If
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "SELECT cOpeCod,Isnull(cOpeDesc,'')cOpeDesc,Isnull(nOpeNiv,'')nOpeNiv, CASE WHEN cOpeVisible='" & gOpeTpoVisibleSi & "' THEN '" & gsSI & "' ELSE '" & gsNO & "' END cOpeVisible, ISNULL(cOpeGruCod,'') cOpeGruCod FROM " & sCentralCom & "OpeTpo" & IIf(psOpecod = "", "", " WHERE (cOpeCod " & IIf(plLike, " LIKE '", "= '") & psOpecod & IIf(plLike, "%')", "')") & "  " & pssqlAddOrOPTPO) & lsFiltro
      Set CargaOpeTpo = dbconec.CargaRecordSet(psSql, LockType)
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Function
CargaOpeTpoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeTpo Method")
End Function

Public Function CargaOpeTpo1(psOpecod As String, Optional plLike As Boolean = False, Optional LockType As LockTypeEnum = adLockReadOnly, Optional pnVisible As Integer = -1, Optional pnNivel As Integer = 0, Optional pssqlAddOrOPTPO As String = "") As Recordset
Dim lsFiltro As String
   On Error GoTo CargaOpeTpoErr
   lsFiltro = ""
   If pnVisible <> -1 Then
        lsFiltro = " and cOpeVisible = '" & pnVisible & "'"
   End If
   If pnNivel > 0 Then
        lsFiltro = lsFiltro & " and nOpeNiv = " & pnNivel & " "
   End If
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "SELECT cOpeCod,Isnull(cOpeDesc,'')cOpeDesc,Isnull(nOpeNiv,'')nOpeNiv, CASE WHEN cOpeVisible='" & gOpeTpoVisibleSi & "' THEN '" & gsSI & "' ELSE '" & gsNO & "' END cOpeVisible, ISNULL(cOpeGruCod,'') cOpeGruCod FROM " & sCentralCom & "OpeTpo" & IIf(psOpecod = "", "", " WHERE cOpeVisible = 1 and (cOpeCod " & IIf(plLike, " LIKE '", "= '") & psOpecod & IIf(plLike, "%')", "')") & "  " & pssqlAddOrOPTPO) & lsFiltro
      Set CargaOpeTpo1 = dbconec.CargaRecordSet(psSql, LockType)
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Function
CargaOpeTpoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeTpo Method")
End Function

Public Sub InsertaOpeTpo(psOpecod As String, psOpeDesc As String, psOpeVisible As String, pnOpeNiv As Integer, psOpeTpo As String, psMovNro As String)
   On Error GoTo InsertaOpeTpoErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeTpo (cOpeCod, cOpeDesc, cOpeVisible, nOpeNiv, cOpeGruCod, cUltimaActualizacion) VALUES ('" & psOpecod & "', '" & Replace(psOpeDesc, "'", "''") & "', '" & psOpeVisible & "', " & pnOpeNiv & ", '" & psOpeTpo & "','" & psMovNro & "') "
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
InsertaOpeTpoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeTpo Method")
End Sub

Public Sub ActualizaOpeTpo(psOpecod As String, psOpeDesc As String, psOpeVisible As String, pnOpeNiv As Integer, psOpeTpo As String, psMovNro As String)
   On Error GoTo ActualziaOpeTpoErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "UPDATE " & sCentralCom & "OpeTpo SET cOpeDesc = '" & psOpeDesc & "', cOpeVisible = '" & psOpeVisible & "', nOpeNiv = " & pnOpeNiv & ", cOpeGruCod = '" & psOpeTpo & "', cUltimaActualizacion = '" & psMovNro & "' " _
            & "WHERE cOpeCod = '" & psOpecod & "'"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
ActualziaOpeTpoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:ActualziaOpeTpo Method")
End Sub

Public Sub EliminaOpeTpo(psOpecod As String)
Dim lsError As String
Dim lsTrans As Boolean
   On Error GoTo EliminaOpeTpoErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      lsTrans = True
      dbconec.BeginTrans
      psSql = "DELETE " & sCentralCom & "OpeDoc WHERE cOpeCod = '" & psOpecod & "' "
      dbconec.Ejecutar psSql
      psSql = "DELETE " & sCentralCom & "OpeObj WHERE cOpeCod = '" & psOpecod & "' "
      dbconec.Ejecutar psSql
      psSql = "DELETE " & sCentralCom & "OpeCta WHERE cOpeCod = '" & psOpecod & "' "
      dbconec.Ejecutar psSql
      psSql = "DELETE " & sCentralCom & "OpeTpo WHERE cOpeCod = '" & psOpecod & "' "
      dbconec.Ejecutar psSql
      dbconec.CommitTrans
      lsTrans = False
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
EliminaOpeTpoErr:
   lsError = Err.Description
   If lsTrans Then
      dbconec.RollbackTrans
   End If
   Set dbconec = Nothing
   Err.Raise Err.Number + MyUnhandledError, "DOperación: EliminaOpeTpo", lsError
End Sub
Public Function CargaOpeCta(psOpecod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "", Optional LockType As LockTypeEnum = adLockReadOnly, Optional pbGetClase As Boolean = False, Optional pnNivSuperior As Integer = 0, Optional pnOpeCtaOpc As Integer = -1) As Recordset
   On Error GoTo CargaOpeCtaErr
   Set dbconec = New DConecta
   Dim lsFiltroDH As String
   Dim lsFiltroOrden As String
    If psDH <> "" Then
        lsFiltroDH = " AND a.cOpeCtaDH ='" & psDH & "' "
    End If
    If psOpeCtaOrden <> "" Then
        lsFiltroOrden = " And a.cOpeCtaOrden='" & psOpeCtaOrden & "' "
    End If
    If pnOpeCtaOpc > -1 Then
      lsFiltroOrden = " AND a.cOpeCtaOpc = '" & pnOpeCtaOpc & "' "
    End If
   If dbconec.AbreConexion() Then
      psSql = "select " _
         & "   a.cCtaContCod, ISNULL(dbo.GetCtaContDesc(a.cCtaContCod," & pnNivSuperior & ",0),'') cCtaContDesc, 1 as Nivel ,  " _
         & "  CASE WHEN a.cOpeCtaDH = '" & gsContHaber & "' THEN '" & gsContHaberDesc & "' " _
         & "       WHEN a.cOpeCtaDH = '" & gsContDebe & "' THEN '" & gsContDebeDesc & "' END  as Clase, " _
         & "  CASE WHEN a.cOpeCtaOpc = '" & gOpeCtaCaracObligatorio & "' then '" & gsOpeCtaCaracterObligaDesc & "' WHEN a.cOpeCtaOpc='" & gOpeCtaCaracOpcional & "' then '" & gsOpeCtaCaracterOpcionDesc & "' END as Caracter, " _
         & "  a.cOpeCod, a.cOpeCtaDH, a.cOpeCtaOpc, a.cCtaContN, a.cOpeCtaOrden " & IIf(pbGetClase, ", cls.cCtaCaracter ", "") _
         & "FROM " & sCentralCom & "OpeCta as a " & IIf(pbGetClase, " JOIN " & sCentralCom & "CtaContClase cls ON a.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' ", "") _
         & "WHERE a.cOpeCod LIKE '" & psOpecod & "' " & lsFiltroDH & lsFiltroOrden _
         & "order by a.cOpeCtaOrden, a.cOpeCtaDH, a.cCtaContCod "
      Set CargaOpeCta = dbconec.CargaRecordSet(psSql, LockType)
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Function
CargaOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeCta Method")
End Function

Public Function CargaOpeCtaIF(psOpecod As String, psIFCod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "") As Recordset
   On Error GoTo CargaOpeCtaErr
   Set dbconec = New DConecta
   Dim lsFiltroDH As String
   Dim lsFiltroOrden As String
    If psDH <> "" Then
        lsFiltroDH = " AND a.cOpeCtaDH ='" & psDH & "' "
    End If
    If psOpeCtaOrden <> "" Then
        lsFiltroOrden = " And a.cOpeCtaOrden='" & psOpeCtaOrden & "' "
    End If
   If dbconec.AbreConexion() Then
      psSql = "SELECT Distinct a.cCtaContCod, ISNULL(dbo.GetCtaContDesc(a.cCtaContCod,1,0),'') cCtaContDesc, " _
         & "   a.cOpeCod, a.cOpeCtaDH, a.cOpeCtaOpc, a.cOpeCtaOrden " _
         & "FROM " & sCentralCom & "OpeCta as a JOIN CtaIFFiltro cif ON cif.cCtaContCod = a.cCtaContCod and cif.cIFTpo = '" & Left(psIFCod, 2) & "' and cif.cPersCod = '" & Mid(psIFCod, 4, 13) & "' " _
         & "WHERE a.cOpeCod LIKE '" & psOpecod & "' " & lsFiltroDH & lsFiltroOrden _
         & "order by a.cOpeCtaDH, a.cOpeCtaOrden, a.cCtaContCod "
      Set CargaOpeCtaIF = dbconec.CargaRecordSet(psSql, adLockReadOnly)
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Function
CargaOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeCtaIF Method")
End Function


Public Function CargaOpeCtaArbol(psOpecod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "") As Recordset
   On Error GoTo CargaOpeCtaErr
   Set dbconec = New DConecta
   Dim lsFiltroDH As String
   Dim lsFiltroOrden As String
    If psDH <> "" Then
        lsFiltroDH = " AND a.cOpeCtaDH ='" & psDH & "' "
    End If
    If psOpeCtaOrden <> "" Then
        lsFiltroOrden = " And a.cOpeCtaOrden='" & psOpeCtaOrden & "' "
    End If
   If dbconec.AbreConexion() Then
      psSql = "SELECT DISTINCT a.cCtaContCod, b.cCtaContDesc, 1 as nNivel " _
         & "FROM " & sCentralCom & "OpeCta as a JOIN " & sCentralCom & "CtaCont as b ON b.cCtaContCod = a.cCtaContCod " _
         & "WHERE a.cOpeCod LIKE '" & psOpecod & "%' " & lsFiltroDH & lsFiltroOrden _
         & "ORDER BY a.cCtaContCod "
      Set CargaOpeCtaArbol = dbconec.CargaRecordSet(psSql, adLockReadOnly)
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Function
CargaOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeCta Method")
End Function

Public Sub InsertaOpeCta(psOpecod As String, psOpeCtaOrden As String, psCtaCod As String, psOpeCtaDH As String, psOpeCtaOpc As String, psCtaCodN As String, psMovNro As String)
   On Error GoTo InsertaOpeCtaErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeCta (cOpeCod, cOpeCtaOrden , cCtaContCod, cOpeCtaDH, cOpeCtaOpc, cCtaContN, cUltimaActualizacion) VALUES ('" & psOpecod & "', '" & psOpeCtaOrden & "','" & psCtaCod & "','" & psOpeCtaDH & "','" & psOpeCtaOpc & "', '" & psCtaCodN & "','" & psMovNro & "')"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
InsertaOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeCta Method")
End Sub

Public Sub EliminaOpeCta(psOpecod As String, psOpeCtaOrden As String, psCtaCod As String, psOpeCtaDH As String)
   On Error GoTo EliminaOpeCtaErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeCta WHERE cOpeCod = '" & psOpecod & "' and cOpeCtaOrden = '" & psOpeCtaOrden & "' and cCtaContCod = '" & psCtaCod & "' and cOpeCtaDH = '" & psOpeCtaDH & "'"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
EliminaOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:EliminaOpeCta Method")
End Sub

Public Function CargaOpeObj(psOpecod As String, Optional LockType As LockTypeEnum = adLockReadOnly, Optional psOpeObjOrden As String = "") As Recordset
   On Error GoTo CargaOpeObjErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "SELECT a.cOpeObjOrden ,a.cObjetoCod, b.cObjetoDesc, a.cOpeObjValor, " & _
         "       a.nOpeObjNiv, a.cOpeObjFiltro, a.cOpeCod" & _
         "  FROM  " & sCentralCom & "OpeObj as a JOIN " & sCentralCom & "Objeto as b ON b.cObjetoCod = a.cObjetoCod " & _
         " WHERE a.cOpeCod = '" & psOpecod & "' " & IIf(psOpeObjOrden = "", "", " and cOpeObjOrden = '" & psOpeObjOrden & "' ")
      Set CargaOpeObj = dbconec.CargaRecordSet(psSql, LockType)
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Function
CargaOpeObjErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeObj Method")
End Function

Public Sub InsertaOpeObj(psOpecod As String, psOpeObjOrden As String, psObjetoCod As String, psOpeObjValor As String, pnOpeObjNiv As Integer, psOpeObjFiltro As String, psMovNro As String)
   On Error GoTo InsertaOpeObjErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeObj (cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjValor, nOpeObjNiv, cOpeObjFiltro, cUltimaActualizacion) " _
            & "VALUES ('" & psOpecod & "','" & psOpeObjOrden & "','" & psObjetoCod & "','" & psOpeObjValor & "'," & pnOpeObjNiv & ",'" & psOpeObjFiltro & "','" & psMovNro & "')"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
InsertaOpeObjErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeObj Method")
End Sub

Public Sub EliminaOpeObj(psOpecod As String, psOpeObjOrden As String, psObjetoCod As String)
   On Error GoTo EliminaOpeObjErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeObj WHERE cOpeCod = '" & psOpecod & "' and cOpeObjOrden = '" & psOpeObjOrden & "' and cObjetoCod = '" & psObjetoCod & "'"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
EliminaOpeObjErr:
   Call RaiseError(MyUnhandledError, "DOperacion:EliminaOpeObj Method")
End Sub

Public Function CargaOpeDoc(psOpecod As String, Optional pnDocTpo As TpoDoc = -1, Optional psMetodo As OpeDocMetodo = -1, Optional LockType As LockTypeEnum = adLockReadOnly, Optional nTipoDoc As Integer = 0) As Recordset
   On Error GoTo CargaOpeDocErr
   Set dbconec = New DConecta
   Set dbconec = New DConecta '1
   Dim lsFiltro As String
   lsFiltro = ""
   If Not pnDocTpo = -1 Then
      lsFiltro = " AND a.nDocTpo = " & pnDocTpo & " "
   End If
   If Not psMetodo = -1 Then
        lsFiltro = lsFiltro & " AND a.cOpeDocMetodo ='" & psMetodo & "'"
   End If
   If dbconec.AbreConexion() Then
   '*******Comentado por alpa
   '***28/02/2008
   If psMetodo = -1 Then
   psSql = "stp_sel_DocumentosOperaciones '" & psOpecod & "', '" & nTipoDoc & "'"
   Else
   psSql = "SELECT a.cOpeCod, a.nDocTpo, d.cDocDesc, " & _
       "CASE WHEN a.cOpeDocEstado = '11'  THEN 'Obligatorio - Debe Existir' " & _
       "     WHEN a.cOpeDocEstado = '12'  THEN 'Obligatorio - No debe Existir' " & _
       "     WHEN a.cOpeDocEstado = '21'  THEN 'Opcional - Debe Existir' " & _
       "     WHEN a.cOpeDocEstado = '22'  THEN 'Opcional - No debe Existir' END as Estado, " & _
       "CASE WHEN a.cOpeDocMetodo = '1' THEN 'Nro. de Movimiento' " & _
       "     WHEN a.cOpeDocMetodo = '2' THEN 'Autogenerado' " & _
       "     WHEN a.cOpeDocMetodo = '3' THEN 'Digitado' END as Metodo , D.cDocAbrev " & _
       "FROM " & sCentralCom & "OpeDoc a JOIN " & sCentralCom & "Documento d ON d.nDocTpo = a.nDocTpo " & _
       "WHERE cOpeCod = '" & psOpecod & "' " & lsFiltro & " ORDER BY a.nDocTpo "
    End If
      Set CargaOpeDoc = dbconec.CargaRecordSet(psSql, LockType)
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Function
CargaOpeDocErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeDoc Method")
End Function

Public Sub InsertaOpeDoc(psOpecod As String, pnDocTpo As Long, psOpeDocEstado As String, psOpeDocMetodo As String, psMovNro As String)
   On Error GoTo InsertaOpeDocErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeDoc (cOpeCod, nDocTpo, cOpeDocEstado, cOpeDocMetodo, cUltimaActualizacion) VALUES ('" & psOpecod & "', '" & pnDocTpo & "', '" & psOpeDocEstado & "','" & psOpeDocMetodo & "','" & psMovNro & "')"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
InsertaOpeDocErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeDoc Method")
End Sub

Public Sub EliminaOpeDoc(psOpecod As String, pnDocTpo As Long)
   On Error GoTo EliminaOpeDocErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeDoc WHERE cOpeCod = '" & psOpecod & "' and nDocTpo = " & pnDocTpo & ""
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
EliminaOpeDocErr:
   Call RaiseError(MyUnhandledError, "DOperacion:EliminaOpeDoc Method")
End Sub

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As New DClassIni
sCentralCom = oIni.BaseComunes
Set oIni = Nothing
End Sub
' Parches
Public Function EmiteDocOpe(ByVal psOpecod As String, ByVal psOpeDocEstado As OpeDocEstado, ByVal psOpeDocMetodo As OpeDocMetodo) As String
On Error GoTo ErrorEmiteDocOpe
Dim rs As ADODB.Recordset
Dim sSql As String
Dim lsFiltroOpeDocEstado As String
Dim lsFiltroOpeDocMetodo As String
Set rs = New ADODB.Recordset


lsFiltroOpeDocEstado = "  AND cOpeDocEstado ='" & psOpeDocEstado & "'  "
lsFiltroOpeDocMetodo = "  AND cOpeDocMetodo ='" & psOpeDocMetodo & "'  "
Dim oConect As DConecta
Set oConect = New DConecta
EmiteDocOpe = ""
If oConect.AbreConexion = False Then Exit Function
sSql = "Select nDocTpo from " & sCentralCom & "OpeDoc where cOpeCod = '" & psOpecod & "' " & lsFiltroOpeDocEstado & lsFiltroOpeDocMetodo
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
Public Function EmiteOpeDoc(ByVal psOpecod As String, ByVal psDocTpo As TpoDoc, Optional ByVal psDH As String = "", Optional ByVal psMetodo As String = "") As String
On Error GoTo ErrorEmiteOpeDoc
Dim rs As ADODB.Recordset
Dim sSql As String
Dim lsFiltroDH As String
Dim lsFiltroM As String
Set rs = New ADODB.Recordset

Dim oConect As DConecta
Set oConect = New DConecta
EmiteOpeDoc = ""
If oConect.AbreConexion = False Then Exit Function
        
If psDH <> "" Then
    lsFiltroDH = " AND OC.cOpeCtaDH ='" & psDH & "'"
End If
If psMetodo <> "" Then
   lsFiltroM = "  and cOpeDocEstado = " & psMetodo & " "
End If
sSql = "Select OC.COPECOD, OD.nDocTpo  " _
    & " From   " & sCentralCom & "OpeCta OC Left Join " & sCentralCom & "OpeDoc OD on OD.cOpeCod = OC.cOpeCod " & lsFiltroM _
    & " Where  Substring(OC.cOpeCod,1,LEN('" & psOpecod & "')) ='" & Mid(psOpecod, 1, Len(psOpecod)) & "' and OD.nDocTpo" & IIf(psDocTpo = -1, " IS NULL ", "= " & psDocTpo & " ") _
    & lsFiltroDH _
    & " GROUP BY OC.COPECOD, OD.nDocTpo  "
    
Set rs = oConect.CargaRecordSet(sSql)
If Not rs.EOF And Not rs.BOF Then
    EmiteOpeDoc = rs!cOpeCod
End If
rs.Close: Set rs = Nothing
oConect.CierraConexion
Set oConect = Nothing
Exit Function
ErrorEmiteOpeDoc:
    Err.Raise vbObjectError + 100, "ErrorEmiteDocOpe", "Error emitir documento en Operacion"
End Function
Public Function EmiteOpeCta(ByVal psOpecod As String, ByVal psDH As String, Optional psOpeCtaOrden As String = "0", Optional ByVal psObjetoCod As String = "", Optional pnFiltro As TpoObjetos, Optional lbMuestraSubCta As Boolean = True, Optional pbBuscaenOpeObj As Boolean = True) As String
    On Error GoTo EmiteOpeCodErr
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConect As DConecta
    Dim lsListaCtasCont As String
    Dim lsSubCta As String
    Dim oCtaIf As NCajaCtaIF
    Dim oCtaCont As DCtaCont
    Dim oContFunc As NContFunciones
    
    Set oCtaCont = New DCtaCont
    Set oCtaIf = New NCajaCtaIF
    Set oContFunc = New NContFunciones
    
    Set oConect = New DConecta
    Set rs = New ADODB.Recordset
    Dim lnNumCtas As Integer
    Dim lsNuevaCta As String
    Dim lsCuenta As String
    Dim lsAux As String
    Dim lbExisteFiltro As Boolean
    Dim lbExisteCuenta As Boolean
    If oConect.AbreConexion = False Then Exit Function
    'Probado por Pepe Para Trujillo y Lima
    sSql = " SELECT  OC.cCtaContCod , OC.cOpeCtaOrden, OC.cOpeCtaDH, CC.cCtaContDesc " _
        & " FROM     " & sCentralCom & "OPECTA OC  JOIN CTACONT CC ON CC.CCTACONTCOD = OC.CCTACONTCOD   " _
        & " WHERE    OC.cOpeCod ='" & psOpecod & "'  AND OC.cOpeCtaDH = '" & psDH & "' AND OC.cOpeCtaOrden = '" & psOpeCtaOrden & "'"
    lnNumCtas = 0
    Set rs = oConect.CargaRecordSet(sSql)
    lnNumCtas = rs.RecordCount
    lbExisteCuenta = False
    Do While Not rs.EOF
        lsListaCtasCont = rs!cCtaContCod
        If lsListaCtasCont <> "" Then
            'lsListaCtasCont = Mid(lsListaCtasCont, 1, Len(lsListaCtasCont) - 3)
            lsNuevaCta = lsListaCtasCont
            If pbBuscaenOpeObj Then
              
                lsSubCta = EmiteSubCtaOpeObj(psOpecod, lsListaCtasCont)
                
            End If
             lsAux = ""
            If lsSubCta = "" Then
                If psObjetoCod <> "" Or lsSubCta <> "" Then
                    lsSubCta = ""
                    lsAux = oContFunc.GetFiltroObjetos(pnFiltro, lsListaCtasCont, psObjetoCod, lbMuestraSubCta, lbExisteFiltro)
                    If lbExisteFiltro Then
                        lbExisteCuenta = True
                    End If
                Else
                    lsAux = IIf(lbMuestraSubCta, lsListaCtasCont, "")
                End If
            Else
                lsAux = IIf(lbMuestraSubCta, lsListaCtasCont, "") & lsSubCta
            End If
            If lsNuevaCta = lsAux And Not lbExisteFiltro And lnNumCtas > 1 And lsSubCta = "" Then
                lsAux = ""
            Else
                lsCuenta = lsAux
            End If
        End If
        lbExisteFiltro = False
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    EmiteOpeCta = lsCuenta
    Set oCtaIf = Nothing
    Set oCtaCont = Nothing
    Exit Function
EmiteOpeCodErr:
    Call RaiseError(MyUnhandledError, "DMov:EmiteOpeCod Method")
End Function
Public Function EmiteSubCtaOpeObj(ByVal psOpecod As String, ByVal psCtaContCod As String, Optional ByVal psOpeObjOrden = "") As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Set oConect = New DConecta
Set rs = New ADODB.Recordset
Dim lsSubCta As String
lsSubCta = ""
If oConect.AbreConexion = False Then Exit Function
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

'Probado por Pepe Para Lima y Trujillo
sql = "Select  cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjFiltro, " _
    & " CASE " _
    & "         WHEN OO.COBJETOCOD ='" & Format(ObjEntidadesFinancieras, "00") & "' THEN " _
    & "                 ISNULL((SELECT Distinct cCtaIFSubCta FROM CTAIFFILTRO WHERE (cPersCod = OO.cOpeObjFiltro OR cCtaIfCod = OO.cOpeObjFiltro ) AND cCtaContCod in ('" & psCtaContCod & "')),'') " _
    & "         WHEN OO.COBJETOCOD ='" & ObjCMACAgenciaArea & "' THEN " _
    & "                 ISNULL((SELECT Distinct cSubCtaCod FROM CtaAreaAgeFiltro WHERE cAreaCod + cAgeCod LIKE OO.cOpeObjFiltro AND cCtaContCod in ('" & psCtaContCod & "')) , '') " _
    & "         WHEN OO.COBJETOCOD ='" & ObjCMACAgencias & "' THEN " _
    & "                 ISNULL((Select Distinct cSubCtaCod from AGENCIAS WHERE cAgeCod LIKE OO.cOpeObjFiltro ),'') " _
    & "         WHEN OO.COBJETOCOD ='" & ObjCMACArea & "' THEN " _
    & "                 ISNULL((SELECT Distinct cSubCtaCod FROM AREAS WHERE cAreaCod LIKE OO.cOpeObjFiltro ),'') " _
    & "         Else    ISNULL((SELECT cCtaObjSubCta FROM CTAOBJFILTRO WHERE COBJETOCOD LIKE OO.cOpeObjFiltro AND cCtaContCod in ('" & psCtaContCod & "')),'')  END as SUBCUENTA " _
    & " From    OpeObj OO " _
    & " Where   OO.cOpecod Like '" & psOpecod & "' " & IIf(psOpeObjOrden = "", "", " and cOpeObjOrden = '" & psOpeObjOrden & "' ") _
    & " ORDER BY OO.COPEOBJORDEN "

Set rs = oConect.CargaRecordSet(sql)
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

Public Function GetOperacionDesc(ByVal psOpecod As String) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sql = " SELECT   cOpeDesc " _
    & " FROM    " & sCentralCom & "OPETpo OC " _
    & " WHERE    COPECOD = '" & psOpecod & "' "
GetOperacionDesc = ""
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetOperacionDesc = psOpecod & " - " & rs!cOpeDesc
End If
RSClose rs
oConect.CierraConexion: Set oConect = Nothing

End Function


Public Function GetOpeCtaDebeHaber(ByVal psOpecod As String, ByVal psCtaContCod As String, Optional ByVal psOpeCtaOrden As String = "0") As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sql = " SELECT   cOpeCtaDH " _
    & " FROM    " & sCentralCom & "OPECTA OC " _
    & " WHERE    COPECOD = '" & psOpecod & "' and cCtaContcod ='" & psCtaContCod & "' and   OC.cOpeCtaOrden='" & psOpeCtaOrden & "' "
GetOpeCtaDebeHaber = ""
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetOpeCtaDebeHaber = rs!cOpeCtaDH
End If
rs.Close
Set rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function EmiteOpeCtasNivel(ByVal psOpecod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "0") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsDHFiltro As String
Dim lsOrdenFiltro As String
Dim lsOrden As String

If psDH <> "" Then
    lsDHFiltro = " AND cOpeCtaDH ='" & psDH & "'"
End If
If psOpeCtaOrden <> "" Then
   lsOrdenFiltro = " and OC.cOpeCtaOrden='" & psOpeCtaOrden & "'"
End If
lsOrden = " ORDER BY OC.cCtaContCod "

sql = "SELECT   OC.cCtaContCod , ISNULL(dbo.GetCtaContDesc(OC.cCtaContCod,1,1),'') cCtaContDesc, 1 as Nivel " _
    & "FROM     " & sCentralCom & "OPECTA OC " _
    & "         " _
    & " WHERE   COPECOD = '" & psOpecod & "' " & lsDHFiltro & lsOrdenFiltro & lsOrden
    

Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

Set rs = oConect.CargaRecordSet(sql)
Set EmiteOpeCtasNivel = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetOpeObj(ByVal psOpecod As String, ByVal psOpeObjOrden As String, _
                        Optional psCtaContCod As String = "", _
                        Optional ByVal psFiltro As String, Optional ByVal psFiltroAdd As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim oConect As DConecta
Dim oRHAreas As DActualizaDatosArea
Dim oContFunct As NContFunciones
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo
Dim lsFiltro As String
Dim lsFiltroIf As String
Dim lsFiltroCta As String

Set oCtaIf = New NCajaCtaIF
Set oContFunct = New NContFunciones
Set oConect = New DConecta
Set oRHAreas = New DActualizaDatosArea
Set oEfect = New Defectivo

Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
If oConect.AbreConexion = False Then
    Exit Function
End If
sql = " SELECT cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjFiltro, nOpeObjNiv " _
    & " FROM " & sCentralCom & "OPEOBJ " _
    & " WHERE COPECOD ='" & psOpecod & "' and cOpeObjOrden='" & psOpeObjOrden & "'"
    
Set rs = oConect.CargaRecordSet(sql)
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
                If rs!nOpeObjNiv = 1 Then
                    Set rs1 = oCtaIf.GetInstFinancieras(lsFiltro)
                Else
                    Set rs1 = oCtaIf.CargaCtasIF(Mid(psOpecod, 3, 1), lsFiltro, rs!nOpeObjNiv)
                End If
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
oConect.CierraConexion: Set oConect = Nothing

Set oCtaIf = Nothing
Set oContFunct = Nothing
Set oConect = Nothing
Set oRHAreas = Nothing
Set oEfect = Nothing

Set GetOpeObj = rs1

End Function
'ALPA 20110703
'EJVG20131118 Se agregó psListaPersCod
Public Function GetRsOpeObj(ByVal psOpecod As String, ByVal psOpeObjOrden As String, Optional ByVal psFiltroObj As String, Optional ByVal psFiltroAdd As String = "", Optional ByVal bMostrarSoloVigentes As Boolean = False, Optional ByVal psListaPersCod As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim oConect As DConecta
Dim oRHAreas As DActualizaDatosArea
Dim oContFunct As NContFunciones
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo
Dim lsFiltro As String
Dim lsFiltroIf As String
Dim lsFiltroCta As String

Set oCtaIf = New NCajaCtaIF
Set oContFunct = New NContFunciones
Set oConect = New DConecta
Set oRHAreas = New DActualizaDatosArea
Set oEfect = New Defectivo

Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
If oConect.AbreConexion = False Then
    Exit Function
End If
sql = " SELECT cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjFiltro, nOpeObjNiv " _
    & " FROM " & sCentralCom & "OPEOBJ " _
    & " WHERE COPECOD ='" & psOpecod & "' and cOpeObjOrden='" & psOpeObjOrden & "'"
    
Set rs = oConect.CargaRecordSet(sql)
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
                If rs!nOpeObjNiv = 1 Then
                    Set rs1 = oCtaIf.GetInstFinancieras(lsFiltro)
                Else
                    Set rs1 = oCtaIf.CargaCtasIF(Mid(psOpecod, 3, 1), lsFiltro, rs!nOpeObjNiv, , False, bMostrarSoloVigentes, psListaPersCod)
                End If
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
oConect.CierraConexion: Set oConect = Nothing

Set oCtaIf = Nothing
Set oContFunct = Nothing
Set oConect = Nothing
Set oRHAreas = Nothing
Set oEfect = Nothing

Set GetRsOpeObj = rs1

End Function

Public Function CargaListaCuentasOperacion(psOpecod As String, psCtaCod As String) As String
On Error GoTo CargaListaCuentasOperacionErr
Dim lsCtas   As String
   Set dbconec = New DConecta
   dbconec.AbreConexion
   psSql = "SELECT cCtaContCod FROM " & sCentralCom & "CtaCont WHERE cCtaContCod LIKE '" & psCtaCod & "_%' and cCtaContCod like '__[12]%'"
   Set prs = dbconec.CargaRecordSet(psSql)
   If Not prs.EOF Then
      Do While Not prs.EOF
         lsCtas = lsCtas & ",'" & prs!cCtaContCod & "'"
         prs.MoveNext
      Loop
   Else
      lsCtas = lsCtas & ",'" & psCtaCod & "'"
   End If
   CargaListaCuentasOperacion = Mid(lsCtas, 2, Len(lsCtas))
   dbconec.CierraConexion
   Set dbconec = Nothing
   Exit Function
CargaListaCuentasOperacionErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaListaCuentasOperacion Method")
End Function

Public Function CargaListaDocsOperacion(psOpecod As String, Optional nTipoDoc As Integer = 0) As String
On Error GoTo CargaListaDocsOperacionErr
Dim lsDocs   As String
   Set prs = CargaOpeDoc(psOpecod, , , , nTipoDoc)
   'Optional pnDocTpo As TpoDoc = -1, Optional psMetodo As OpeDocMetodo = -1, Optional LockType As LockTypeEnum = adLockReadOnly, Optional nTipoDoc As Integer = 0
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
   Call RaiseError(MyUnhandledError, "DOperacion:CargaListaDocsOperacion Method")
End Function

Public Function CargaOpeCtaUltimoNivel(psOpecod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "") As Recordset
   On Error GoTo CargaOpeCtaUltimoNivelErr
   
   'Probado por Pepe Para Lima y Trujillo
   Set dbconec = New DConecta
   Dim lsFiltroDH As String
   Dim lsFiltroOrden As String
   
    If psDH <> "" Then
        lsFiltroDH = " AND a.cOpeCtaDH ='" & psDH & "' "
    End If
    If psOpeCtaOrden <> "" Then
        lsFiltroOrden = " And a.cOpeCtaOrden='" & psOpeCtaOrden & "' "
    End If
    
   If dbconec.AbreConexion() Then
      'psSql = "SELECT b.cCtaContCod, b.cCtaContDesc " _
         '& "FROM " & sCentralCom & "OpeCta as a JOIN " & sCentralCom & "CtaCont as b ON b.cCtaContCod LIKE a.cCtaContCod+'%' " _
         '& "WHERE a.cOpeCod LIKE '" & psOpecod & "%' " & lsFiltroDH & lsFiltroOrden _
         '& "ORDER BY b.cCtaContCod "
      psSql = "Exec stp_sel_ObtieneCargaOpeCtaUltimoNivel '" & psOpecod & "','" & psDH & "','" & psOpeCtaOrden & "'" 'Agregado by NAGL 20190612 Según INC1903010011
      Set CargaOpeCtaUltimoNivel = dbconec.CargaRecordSet(psSql, adLockReadOnly)
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Function
CargaOpeCtaUltimoNivelErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeCtaUltimoNivel Method")
End Function

Public Sub EliminaCtaBS(psOpecod As String, psCtaCod As String, psBSCod As String)
   On Error GoTo EliminaOpeCtaErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "CtaBS WHERE cOpeCod = '" & psOpecod & "' and cCtaContCod = '" & psCtaCod & "' and cObjetoCod = '" & psBSCod & "'"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
EliminaOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:EliminaOpeCta Method")
End Sub

Public Sub InsertaCtaBS(psOpecod As String, psCtaCod As String, psBSCod As String)
   On Error GoTo InsertaCtaBSErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "INSERT INTO " & sCentralCom & "CtaBS (cCtaContCod, nCtaObjOrden, cObjetoCod, cOpeCod, nCtaObjNiv, cCtaObjFiltro, cCtaObjImpre, cCtaObjOpc) VALUES ('" & psCtaCod & "',1,'" & psBSCod & "', '" & psOpecod & "', 2, '', '', 0 )"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
InsertaCtaBSErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaCtaBS Method")
End Sub

Public Function GetOperacionRSRefer(psOpecod As String, Optional pnTipoRef = 9, Optional pbBitCentral As Boolean = False) As ADODB.Recordset
On Error GoTo ErrGetOperacionReferencia
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
        If pbBitCentral Then
            psSql = "SELECT ore.cOpeCodRef, o.cOpeDesc " _
                  & "FROM OpeRef ore JOIN OpeTpo o ON o.cOpeCod = ore.cOpeCodRef " _
                  & "WHERE ore.cOpeCod = '" & psOpecod & "' and nTpoRef = " & pnTipoRef
        Else
            psSql = "SELECT ore.cOpeCodRef, o.cNomOpe cOpeDesc " _
                  & "FROM OpeRef ore JOIN DBComunes..Operacion o ON o.cCodOpe = ore.cOpeCodRef " _
                  & "WHERE ore.cOpeCod = '" & psOpecod & "' and nTpoRef = " & pnTipoRef
        End If
        Set GetOperacionRSRefer = dbconec.CargaRecordSet(psSql)
        dbconec.CierraConexion
   End If
   Set dbconec = Nothing
Exit Function
ErrGetOperacionReferencia:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetOperacionRefencia(psOpecod As String, Optional pnTipoRef = 9) As String
Dim rs As ADODB.Recordset
Dim lsOpes As String
On Error GoTo ErrGetOperacionReferencia
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
        psSql = "SELECT cOpeCodRef FROM OpeRef WHERE cOpeCod = '" & psOpecod & "' and nTpoRef = " & pnTipoRef
        Set rs = dbconec.CargaRecordSet(psSql)
        If Not rs.EOF Then
            If rs.RecordCount = 1 Then
               GetOperacionRefencia = rs!cOpeCodRef
            Else
               lsOpes = RSMuestraLista(rs)
               GetOperacionRefencia = Mid(lsOpes, 2, Len(lsOpes) - 2)
            End If
        End If
        dbconec.CierraConexion
   End If
   Set dbconec = Nothing
Exit Function
ErrGetOperacionReferencia:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ValidaCtaCont(psCtaCont As String) As Boolean
Dim rs As ADODB.Recordset
Dim sql As String
On Error GoTo ErrGetOperacionReferencia
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
        sql = "Select cCtaContCod From CtaCont where cCtaContCod Like '" & psCtaCont & "%' And nCtaEstado = 1"
        Set rs = dbconec.CargaRecordSet(sql)
        If rs.RecordCount = 1 Then
            ValidaCtaCont = True
        Else
            ValidaCtaCont = False
        End If
        dbconec.CierraConexion
   End If
   Set dbconec = Nothing
Exit Function
ErrGetOperacionReferencia:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function CargaOpeCtaPlantillaNeg(psCtaCont As String, psOpecod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "", Optional LockType As LockTypeEnum = adLockReadOnly, Optional pbGetClase As Boolean = False, Optional pnNivSuperior As Integer = 0, Optional pnOpeCtaOpc As Integer = -1) As Recordset
    Dim dbconec As DConecta
   On Error GoTo CargaOpeCtaErr
   Set dbconec = New DConecta
   Dim lsFiltroDH As String
   Dim lsFiltroOrden As String
    If psDH <> "" Then
        lsFiltroDH = " AND a.cOpeCtaDH ='" & psDH & "' "
    End If
    If psOpeCtaOrden <> "" Then
        lsFiltroOrden = " And a.cOpeCtaOrden='" & psOpeCtaOrden & "' "
    End If
    If pnOpeCtaOpc > -1 Then
      lsFiltroOrden = " AND a.cOpeCtaOpc = '" & pnOpeCtaOpc & "' "
    End If
   If dbconec.AbreConexion() Then
      psSql = "select " _
         & "   a.cCtaContCod, ISNULL(dbo.GetCtaContDesc(a.cCtaContCod," & pnNivSuperior & ",0),'') cCtaContDesc, 1 as Nivel ,  " _
         & "  CASE WHEN a.cOpeCtaDH = '" & gsContHaber & "' THEN '" & gsContHaberDesc & "' " _
         & "       WHEN a.cOpeCtaDH = '" & gsContDebe & "' THEN '" & gsContDebeDesc & "' END  as Clase, " _
         & "  CASE WHEN a.cOpeCtaOpc = '" & gOpeCtaCaracObligatorio & "' then '" & gsOpeCtaCaracterObligaDesc & "' WHEN a.cOpeCtaOpc='" & gOpeCtaCaracOpcional & "' then '" & gsOpeCtaCaracterOpcionDesc & "' END as Caracter, " _
         & "  a.cOpeCod, a.cOpeCtaDH, a.cOpeCtaOpc, a.cCtaContN, a.cOpeCtaOrden " & IIf(pbGetClase, ", cls.cCtaCaracter ", "") _
         & "FROM " & sCentralCom & "OpeCta as a " & IIf(pbGetClase, " JOIN " & sCentralCom & "CtaContClase cls ON a.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' ", "") _
         & "WHERE a.cCtaContCod like '" & psCtaCont & "%'  and a.cOpeCod LIKE '" & psOpecod & "' " & lsFiltroDH & lsFiltroOrden _
         & "order by a.cOpeCtaOrden, a.cOpeCtaDH, a.cCtaContCod "
      Set CargaOpeCtaPlantillaNeg = dbconec.CargaRecordSet(psSql, LockType)
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Function
CargaOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeCta Method")
End Function

'*** PEAC 20100708
Public Function CargaAreaSoli() As Recordset
   Dim dbconec As DConecta
   On Error GoTo CargaOpeDocErr
   Set dbconec = New DConecta
   
   If dbconec.AbreConexion() Then
        psSql = " exec stp_sel_ObtieneAreasSoporte "
        Set CargaAreaSoli = dbconec.CargaRecordSet(psSql)
        dbconec.CierraConexion
   End If
   
   Set dbconec = Nothing
   Exit Function
CargaOpeDocErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaAreaSoli")
End Function

'*** PEAC 20100708
Public Function CargaPoliSeguPatri(ByVal nTipoBus As Integer, Optional ByVal psPersCod As String = "@", Optional ByVal psRegistroDel As String = "", Optional ByVal psRegistroAl As String = "") As Recordset
   Dim dbconec As DConecta
   On Error GoTo CargaPoliSeguPatriErr
   Set dbconec = New DConecta
   
   If dbconec.AbreConexion() Then
        psSql = " EXEC stp_sel_CargaPolizaSeguPatri " & nTipoBus & ",'" & psPersCod & "','" & psRegistroDel & "','" & psRegistroAl & "'"
        Set CargaPoliSeguPatri = dbconec.CargaRecordSet(psSql)
        dbconec.CierraConexion
   End If
   
   Set dbconec = Nothing
   Exit Function
CargaPoliSeguPatriErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaPoliSeguPatri")
End Function

'*** PEAC 20200922
Public Sub InsertaPoliSeguPatri(ByVal psCodPoli As String, ByVal pnTipoSegu As Integer, ByVal psPersCodAse As String, ByVal pnMontoPrima As Double, ByVal psVigeDel As String, ByVal psVigeAl As String, ByVal pnMoneda As Integer, ByVal psDescrip As String, ByVal psMatrizGastos As Variant, ByVal pnCantAgeSel As Integer, ByVal psMovNro As String)
   Dim dbconec As DConecta
   
   Dim lcCodAge As String, lnPorcen As Double
   Dim I As Integer, j As Integer
   
   On Error GoTo InsertaPoliSeguPatriErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then

    dbconec.BeginTrans
    psSql = " exec stp_ins_PolizaSeguPatri '" & psCodPoli & "'," & pnTipoSegu & ",'" & psPersCodAse & "'," & pnMontoPrima & ",'" & psVigeDel & "','" & psVigeAl & "'," & pnMoneda & ",'" & psDescrip & "','" & psMovNro & "'"
    dbconec.Ejecutar psSql

    For I = 1 To pnCantAgeSel
        For j = 1 To 3
            If j = 1 Then
                lcCodAge = psMatrizGastos(j, I)
            ElseIf j = 3 Then
                lnPorcen = psMatrizGastos(j, I)
            End If
        Next j
        psSql = " exec stp_ins_PolizaSeguPatriPorenDet '" & psCodPoli & "','" & lcCodAge & "'," & pnTipoSegu & "," & lnPorcen
        dbconec.Ejecutar psSql
    Next I
    dbconec.CommitTrans
    
    dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
InsertaPoliSeguPatriErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaPoliSeguPatri Method")
End Sub

'*** PEAC 20100922
Public Function CargaDatosPoliSeguPatri(ByVal psNumPoliza As String, ByVal pnTipoPoliza As Integer, ByVal psPersCodAse As String) As Recordset
    Dim dbconec As DConecta
   On Error GoTo CargaDatosPoliSeguPatriErr
   Set dbconec = New DConecta
   
   If dbconec.AbreConexion() Then
        psSql = " exec stp_sel_ObtieneDatosPoliSeguPatri '" & psNumPoliza & "'," & pnTipoPoliza & ",'" & psPersCodAse & "'"
        Set CargaDatosPoliSeguPatri = dbconec.CargaRecordSet(psSql)
        dbconec.CierraConexion
   End If

   Set dbconec = Nothing
   Exit Function
CargaDatosPoliSeguPatriErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaDatosPoliSeguPatri")
End Function

'*** PEAC 20200922
Public Sub ActualizaPoliSeguPatri(ByVal psCodPoli As String, ByVal pnTipoSegu As Integer, ByVal psPersCodAse As String, ByVal pnMontoPrima As Double, ByVal psVigeDel As String, ByVal psVigeAl As String, ByVal pnMoneda As Integer, ByVal psDescrip As String, ByVal psMatrizGastos As Variant, ByVal pnCantAgeSel As Integer, ByVal psMovNro As String)
   Dim dbconec As DConecta
   
   Dim I As Integer, j As Integer
   Dim lcCodAge As String, lnPorcen As Double
   
   On Error GoTo ActualizaPoliSeguPatriErr
   Set dbconec = New DConecta
   
   If dbconec.AbreConexion() Then
      
        dbconec.BeginTrans
        
        psSql = " exec stp_upd_PolizaSeguPatri '" & psCodPoli & "'," & pnTipoSegu & ",'" & psPersCodAse & "'," & pnMontoPrima & ",'" & psVigeDel & "','" & psVigeAl & "'," & pnMoneda & ",'" & psDescrip & "','" & psMovNro & "'"
        dbconec.Ejecutar psSql
          
        psSql = " exec stp_del_PolizaSeguPatriPorenDet '" & psCodPoli & "'"
        dbconec.Ejecutar psSql
        
        For I = 1 To pnCantAgeSel
            For j = 1 To 3
                If j = 1 Then
                    lcCodAge = psMatrizGastos(j, I)
                ElseIf j = 3 Then
                    lnPorcen = psMatrizGastos(j, I)
                End If
            Next j
            psSql = " exec stp_ins_PolizaSeguPatriPorenDet '" & psCodPoli & "','" & lcCodAge & "'," & pnTipoSegu & "," & lnPorcen
            dbconec.Ejecutar psSql
        Next I
          
        dbconec.CommitTrans
        dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
ActualizaPoliSeguPatriErr:
   Call RaiseError(MyUnhandledError, "DOperacion:ActualizaPoliSeguPatri Method")
End Sub

'*** PEAC 20100922
Public Sub EliminaPolizaSeguPatri(ByVal psCodPoli As String, ByVal pnTipoSegu As Integer, ByVal psPersCodAse As String, ByVal psMovNro As String)
    Dim dbconec As DConecta
   On Error GoTo EliminaPolizaSeguPatriErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = " exec stp_upd_DesactivaPolizaSeguPatri '" & psCodPoli & "'," & pnTipoSegu & ",'" & psPersCodAse & "','" & psMovNro & "'"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
EliminaPolizaSeguPatriErr:
   Call RaiseError(MyUnhandledError, "DOperacion:EliminaPolizaSeguPatri Method")
End Sub

'*** PEAC 20100927
Public Function VerificaAsiento(ByVal psOpecod As String, ByVal psFecha As String, ByVal psTipo As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    
    Dim sSql As String
    
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
                
        sSql = " exec stp_sel_VerificaAsiento '" & psOpecod & "','" & psTipo & "','" & psFecha & "'"
        
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If

    Set VerificaAsiento = rs
    Set rs = Nothing
End Function

'*** PEAC 20100928
Public Function ObtieneCtasPoliSeguPatri(ByVal psOpecod As String, ByVal psDH As String, ByVal pnConcepto As Integer) As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    
    If oConec.AbreConexion() Then
        sql = " exec stp_sel_ObtieneCtasPoliSeguPatri '" & psOpecod & "','" & psDH & "'," & pnConcepto
        Set rs = oConec.CargaRecordSet(sql)
        oConec.CierraConexion
    End If
    If rs.EOF And rs.BOF Then
        ObtieneCtasPoliSeguPatri = ""
    Else
        ObtieneCtasPoliSeguPatri = rs!cCtaContCod   '*** rs.Fields(0)
    End If
End Function

'*** PEAC 20100929
Public Function ObtieneCtasIngresosGastosPoliSeguPatri() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = " exec stp_sel_ObtieneCtasIngresoGastosPoliSeguPatri "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set ObtieneCtasIngresosGastosPoliSeguPatri = rs
    Set rs = Nothing
End Function

'*** PEAC 20100929
Public Sub EliminaAsientoCont(ByVal pnMovNro As Long)
    Dim dbconec As DConecta
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = " exec stp_upd_ExtornaAsientoCont " & pnMovNro
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
End Sub

'*** PEAC 20200922
Public Sub InsertaFondoSeguDepo(ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal pnTrim As Integer, ByVal pnTasa As Double, ByVal pnTC As Double, ByVal pnTotDepMN As Double, ByVal pnTotDepME As Double, ByVal psOpecod As String, ByVal psFecha As String, ByVal pdFecsis As Date, ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnTCVenta As Double, ByVal pnTCFijo As Double, ByVal psMovNro As String)
    Dim dbconec As DConecta
    Dim oVeri As DOperacion
    Dim rs As ADODB.Recordset, rs1 As ADODB.Recordset
    Dim oCont As NContFunciones
    Set oCont = New NContFunciones
    Dim ldFechaRegistro As Date
    Dim lsSQL As String
    Dim lnMovNro As Long
    Dim nConta As Integer
    Dim I As Integer
    Dim lsMovNro As String, lsCtaCont As String

    Dim lnTotDebeME As Double

    Dim oDep As DOperacion
    Set oDep = New DOperacion

    Dim lnTotHaber As Double, lnHaber As Double, lnDebe As Double
    Dim lnDebeME As Double
    Dim lnTotDebe As Double
    Dim lcCtaDif As String
    Dim lnTotDepMEDebeMN As Double, lnTotDepMEHaberMN As Double
    Dim lnPrimaSol As Double, lnPrimaDol As Double
    
   On Error GoTo InsertaFondoSeguDepoErr
   Set dbconec = New DConecta
    
   If dbconec.AbreConexion() Then
        lsSQL = " exec stp_ins_FondoSeguroDeposito " & pnAnio & "," & pnMes & "," & pnTrim & "," & pnTasa & "," & pnTC & "," & pnTotDepMN & "," & pnTotDepME & ",'" & psMovNro & "'"
        dbconec.Ejecutar lsSQL
        dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   
   Exit Sub
InsertaFondoSeguDepoErr:
     dbconec.RollbackTrans
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaFondoSeguDepo Method")
End Sub

'*** PEAC 20200922
Public Sub ActualizaFondoSeguDepo(ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal pnTrim As Integer, ByVal pnTasa As Double, ByVal pnTC As Double, ByVal pnTotDepMN As Double, ByVal pnTotDepME As Double, ByVal psMovNro As String)
   Dim dbconec As DConecta
     
   On Error GoTo ActualizaFondoSeguDepoErr
   Set dbconec = New DConecta
   
   If dbconec.AbreConexion() Then
             
        psSql = " exec stp_upd_FondoSeguroDeposito " & pnAnio & "," & pnMes & "," & pnTrim & "," & pnTasa & "," & pnTC & "," & pnTotDepMN & "," & pnTotDepME & ",'" & psMovNro & "'"
        dbconec.Ejecutar psSql
                    
        dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
ActualizaFondoSeguDepoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:ActualizaFondoSeguDepo Method")
End Sub


'*** PEAC 20100708
Public Function CargaFondoSeguDepo(ByVal nTipoBus As Integer, Optional ByVal psRegistroDel As String = "", Optional ByVal psRegistroAl As String = "") As Recordset
   Dim dbconec As DConecta
   On Error GoTo CargaFondoSeguDepoErr
   Set dbconec = New DConecta
   
   If dbconec.AbreConexion() Then
        psSql = " EXEC stp_sel_CargaFondoSeguDepo " & nTipoBus & ",'" & psRegistroDel & "','" & psRegistroAl & "'"
        Set CargaFondoSeguDepo = dbconec.CargaRecordSet(psSql)
        dbconec.CierraConexion
   End If
   
   Set dbconec = Nothing
   Exit Function
CargaFondoSeguDepoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaFondoSeguDepo")
End Function

'*** PEAC 20100922
Public Function CargaDatosFondoSeguDepo(ByVal pnAnio As Integer, ByVal pnMes As Integer) As Recordset
    Dim dbconec As DConecta
   On Error GoTo CargaDatosFondoSeguDepoErr
   Set dbconec = New DConecta
   
   If dbconec.AbreConexion() Then
        psSql = " exec stp_sel_ObtieneDatosFondoSeguDepo " & pnAnio & "," & pnMes
        Set CargaDatosFondoSeguDepo = dbconec.CargaRecordSet(psSql)
        dbconec.CierraConexion
   End If

   Set dbconec = Nothing
   Exit Function
CargaDatosFondoSeguDepoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaDatosFondoSeguDepo")
End Function

'*** PEAC 20100922
Public Sub EliminaFondoSeguDepo(ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal psMovNro As String)
    Dim dbconec As DConecta
   On Error GoTo EliminaFondoSeguDepoErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      
      dbconec.BeginTrans
      
      psSql = " exec stp_upd_DesactivaFondoSeguDepo " & pnAnio & "," & pnMes & ",'" & psMovNro & "'"
      dbconec.Ejecutar psSql
      
'      psSql = " exec stp_upd_DesactivaAsientoFondoSeguDepo '" & psCodPoli & "'," & pnTipoSegu & ",'" & psPersCodAse & "','" & psMovNro & "'"
'      dbConec.Ejecutar psSql
      
      dbconec.CommitTrans
      
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
EliminaFondoSeguDepoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:EliminaFondoSeguDepo Method")
End Sub

'*** PEAC 20100929
Public Function VerificaAsientoCont(ByVal psOpecod As String, ByVal psFecha As String, Optional ByVal psAges As String = "%%") As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = " exec stp_sel_VerificaAsientoCont '" & psOpecod & "','" & psFecha & "','" & psAges & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set VerificaAsientoCont = rs
    Set rs = Nothing
End Function

'*** PEAC 20100922
Public Function CargaTipoCambioDelDia(ByVal psFecha As String, ByVal nTipoTC As Integer) As Double
    Dim dbconec As DConecta
    Dim rs As ADODB.Recordset
    
    On Error GoTo CargaTipoCambioDelDiaErr
    Set dbconec = New DConecta

    If dbconec.AbreConexion() Then
        psSql = " exec stp_sel_ObtieneTipoCambioDelDia '" & psFecha & "'"
        Set rs = dbconec.CargaRecordSet(psSql)
        
        If Not (rs.EOF And rs.BOF) Then
            If nTipoTC = 1 Then
                CargaTipoCambioDelDia = rs!PonderadoVenta
            Else
                CargaTipoCambioDelDia = rs!Fijo
            End If
        Else
            CargaTipoCambioDelDia = 0
        End If
        
        'Set CargaTipoCambioDelDia = dbConec.CargaRecordSet(psSql)
        
        dbconec.CierraConexion
    End If

    Set dbconec = Nothing
    Exit Function
CargaTipoCambioDelDiaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaTipoCambioDelDia")
End Function


'*** PEAC 20100922
Public Function CargaPrima(ByVal pnAnio As Integer, ByVal pnTrim As Integer, ByVal pnMes As Integer) As Double
    Dim dbconec As DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    On Error GoTo CargaPrimaErr
    Set dbconec = New DConecta

    If dbconec.AbreConexion() Then
        psSql = " exec stp_sel_ObtienePrimaFSD " & pnAnio & "," & pnTrim & "," & pnMes
        Set rs = dbconec.CargaRecordSet(psSql)
        
'        If Not (rs.EOF And rs.BOF) Then
'            If pnMoneda = 1 Then
'                CargaPrima = rs!nPrimaS
'            Else
'                CargaPrima = rs!nPrimaD
'            End If
'        Else
'            CargaPrima = 0
'        End If
        
        dbconec.CierraConexion
    End If

    Set dbconec = Nothing
    Exit Function
CargaPrimaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaPrima")
End Function

'*** PEAC 20200922
Public Sub GeneraAsientoContableFSD(ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal pnTrim As Integer, ByVal pnTasa As Double, ByVal pnTC As Double, ByVal pnTotDepMN As Double, ByVal pnTotDepME As Double, ByVal psOpecod As String, ByVal psFecha As String, ByVal pdFecsis As Date, ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnTCVenta As Double, ByVal pnTCFijo As Double, ByVal psMovNro As String)

    Dim dbconec As DConecta
    Dim oVeri As DOperacion
    Dim rs As ADODB.Recordset, rs1 As ADODB.Recordset
    Dim oCont As NContFunciones
    Set oCont = New NContFunciones
    Dim ldFechaRegistro As Date
    Dim lsSQL As String
    Dim lnMovNro As Long
    Dim nConta As Integer
    Dim I As Integer
    Dim lsMovNro As String, lsCtaCont As String

    Dim lnTotDebeME As Double

    Dim oDep As DOperacion
    Set oDep = New DOperacion

    Dim oPrevio As clsPrevioFinan
    Set oPrevio = New clsPrevioFinan
    
    Dim oAsiento As NContImprimir
    Set oAsiento = New NContImprimir
    
    Dim lsMovNroS As String, lsMovNroD As String
    'Dim nConta As Integer,
    Dim lnTotHaber As Double, lnHaber As Double, lnDebe As Double
    Dim lnDebeME As Double
    Dim lnTotDebe As Double
    Dim lcCtaDif As String
    Dim lnTotDepMEDebeMN As Double, lnTotDepMEHaberMN As Double
    Dim lnPrimaSol As Double, lnPrimaDol As Double
    Dim lnMontoPriS As Double, lnMontoPriD As Double
    Dim lnTotDebeNew As Double, lnTotDebeMENew As Double
    
   On Error GoTo GeneraAsientoContableFSDErr
   Set dbconec = New DConecta
       
    Set rs = New ADODB.Recordset
   
    Set oVeri = New DOperacion
    Set rs = oVeri.VerificaAsientoCont(psOpecod, psFecha)
    Set rs1 = oVeri.ObtieneCtasIngresosGastosPoliSeguPatri
    Set oVeri = Nothing

    Call oDep.CargaPrima(pnAnio, pnTrim, pnMes)

    lnMontoPriS = oDep.ObtieneMontosPrima(pnMes, 1)
    lnMontoPriD = oDep.ObtieneMontosPrima(pnMes, 2)

    If rs.EOF And rs.BOF Then

        Dim oMov As DMov
        Set oMov = New DMov
        
        ldFechaRegistro = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & pnMes & "/" & pnAnio)))

        oMov.BeginTrans
            
            For I = 1 To 2 '' para la moneda 1 soles 2 dolar
                Set rs = Nothing
                
                lsMovNro = oMov.GeneraMovNro(pdFecsis, psCodAge, psCodUser)
                oMov.InsertaMov lsMovNro, psOpecod, "Reg. Provisión Prima FSD - " & CStr(ldFechaRegistro) & " - " & Trim(CStr(pnTrim)) & " Trim. " & IIf(I = 1, "MN", "ME")
                lnMovNro = oMov.GetnMovNro(lsMovNro)

                If I = 1 Then
                    lsMovNroS = lsMovNro
                Else
                    lsMovNroD = lsMovNro
                End If

                lsSQL = " exec stp_sel_AgenciaPorcentajeGastos "
                Set rs = oMov.CargaRecordSet(lsSQL)
            
                nConta = 0: lnTotHaber = 0: lnHaber = 0: lnDebe = 0
                lnTotDebe = 0: lnTotDepMEDebeMN = 0: lnTotDepMEHaberMN = 0
                lnTotDebeMENew = 0: lnTotDebeNew = 0
            
                Do While Not rs.EOF
                    nConta = nConta + 1
                    lsCtaCont = oDep.ObtieneCtasPoliSeguPatri(psOpecod, "D", 0)
                    lsCtaCont = Replace(lsCtaCont, "AG", rs!cAgeCod)
                    lsCtaCont = Replace(lsCtaCont, "M", Trim(CStr(I)))
                    
                    If I = 1 Then
                        lnDebe = Round((lnMontoPriS * rs!nPorcenCarteraAho) / 100, 2)
                        oMov.InsertaMovCta lnMovNro, nConta, lsCtaCont, lnDebe
                        
                        lnTotDebeNew = lnTotDebeNew + lnDebe
                        
                    Else
                    
                        lnTotDepMEDebeMN = Round(lnMontoPriD * pnTCVenta, 2)
                    
                        lnDebe = Round((lnTotDepMEDebeMN * rs!nPorcenCarteraAho) / 100, 2)
                        lnDebeME = Round((lnMontoPriD * rs!nPorcenCarteraAho) / 100, 2)
                        oMov.InsertaMovCta lnMovNro, nConta, lsCtaCont, lnDebe
                        oMov.InsertaMovMe lnMovNro, nConta, lnDebeME
                                                
                        lnTotDebeMENew = lnTotDebeMENew + lnDebeME
                        
                    End If
                rs.MoveNext
                Loop
                lsCtaCont = oDep.ObtieneCtasPoliSeguPatri(psOpecod, "H", 0)
                lsCtaCont = Replace(lsCtaCont, "M", Trim(CStr(I)))
                nConta = nConta + 1
                If I = 1 Then
                    
                    If Round(lnTotDebeNew, 2) > Round(lnMontoPriS, 2) Then
                        'oMov.act lnMovNro, nConta - 1, , lnDebe - (lnTotDebeNew - lnMontoPriS)
                    ElseIf Round(lnMontoPriS, 2) > Round(lnTotDebeNew, 2) Then
                        oMov.ActualizaMovCta lnMovNro, nConta - 1, , lnDebe + (lnMontoPriS - lnTotDebeNew)
                    End If
                
                    oMov.InsertaMovCta lnMovNro, nConta, lsCtaCont, lnMontoPriS * -1
                Else
                    
                    If Round(lnTotDebeMENew, 2) > Round(lnMontoPriD, 2) Then
                        oMov.ActualizaMovMe lnMovNro, nConta - 1, Round(lnDebeME - (lnTotDebeMENew - lnMontoPriD), 2)
                    ElseIf Round(lnMontoPriD, 2) > Round(lnTotDebeMENew, 2) Then
                        oMov.ActualizaMovMe lnMovNro, nConta - 1, Round(lnDebeME + (lnMontoPriD - lnTotDebeMENew), 2)
                    End If
                    
                    lnTotDepMEHaberMN = Round(lnMontoPriD * pnTCFijo, 2)
                    
                    oMov.InsertaMovCta lnMovNro, nConta, lsCtaCont, lnTotDepMEHaberMN * -1
                    oMov.InsertaMovMe lnMovNro, nConta, lnMontoPriD * -1
                
                    nConta = nConta + 1
                    
                    If lnTotDepMEHaberMN > lnTotDepMEDebeMN Then
                        lcCtaDif = rs1!cCtaGasto
                        oMov.InsertaMovCta lnMovNro, nConta, lcCtaDif, (lnTotDepMEHaberMN - lnTotDepMEDebeMN)
                    ElseIf lnTotDepMEHaberMN < lnTotDepMEDebeMN Then
                        lcCtaDif = rs1!cCtaIngreso
                        oMov.InsertaMovCta lnMovNro, nConta, lcCtaDif, (lnTotDepMEDebeMN - lnTotDepMEHaberMN) * -1
                    End If
                    
                End If
            Next I

        oMov.CommitTrans
    End If
    
   Set oMov = Nothing
   
   For I = 1 To 2
    oPrevio.Show oAsiento.ImprimeAsientoContable(IIf(I = 1, lsMovNroS, lsMovNroD), 60, 80, "A S I E N T O   C O N T A B L E"), "ASIENTO CONTABLE", True
   Next I
   
   
   Exit Sub
GeneraAsientoContableFSDErr:
    oMov.RollbackTrans
   Call RaiseError(MyUnhandledError, "DOperacion:GeneraAsientoContableFSD Method")
End Sub

'*** PEAC 20100922
Public Function ObtieneMontosPrima(ByVal pnMes As Integer, ByVal pnMoneda As Integer) As Double
    Dim dbconec As DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    On Error GoTo ObtieneMontosPrimaErr
    Set dbconec = New DConecta

    If dbconec.AbreConexion() Then
        psSql = " exec stp_sel_ObtieneMontosPrimaFSD " & pnMes
        Set rs = dbconec.CargaRecordSet(psSql)
        
        If Not (rs.EOF And rs.BOF) Then
            If pnMoneda = 1 Then
                ObtieneMontosPrima = rs!nPrimaS
            Else
                ObtieneMontosPrima = rs!nPrimaD
            End If
        Else
            ObtieneMontosPrima = 0
        End If
        
        dbconec.CierraConexion
    End If

    Set dbconec = Nothing
    Exit Function
ObtieneMontosPrimaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:ObtieneMontosPrima")
End Function

'*** PEAC 20101012
Public Sub AnulaAsientoContFSD(ByVal psFecha As String)
   On Error GoTo AnulaAsientoContFSDErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
   
      psSql = " exec stp_upd_AnulaAsientoContFSD '" & psFecha & "'"
      
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
AnulaAsientoContFSDErr:
   Call RaiseError(MyUnhandledError, "DOperacion:AnulaAsientoContFSD Method")
End Sub


'*** PEAC 20101012
Public Sub AnulaRegistroFSD(ByVal pnAnio As Integer, ByVal pnMes As Integer)
   On Error GoTo AnulaRegistroFSDErr
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
   
      psSql = " exec stp_upd_AnulaRegistroFSD " & pnAnio & "," & pnMes
      
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
AnulaRegistroFSDErr:
   Call RaiseError(MyUnhandledError, "DOperacion:AnulaRegistroFSD Method")
End Sub

'*** MADM 20110505
Public Function DevuelveGastosCreditos() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = " exec stp_sel_RecuperaGastosRecupFinan "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set DevuelveGastosCreditos = rs
    Set rs = Nothing
End Function

Public Function dObtieneDatosAsignaGastoCredRecup(ByVal psCtaCod As String) As ADODB.Recordset
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Set dbconec = New DConecta
Set lrs = New ADODB.Recordset
On Error GoTo dError
If dbconec.AbreConexion() Then
lsSQL = " stp_sel_RecuperaGastoCreditoFinan '" & psCtaCod & "'"

Set lrs = dbconec.CargaRecordSet(lsSQL)
   End If
Set dObtieneDatosAsignaGastoCredRecup = lrs
Set lrs = Nothing
Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosContrato>>", Err.Description
End Function

Public Sub nGastoRecupAsignaNuevo(ByVal psCtaCod As String, ByVal psFechaHora As String, _
        ByVal psMovNro As String, ByVal pnNroGastoCta As Integer, ByVal psFecAsigna As String, ByVal pnColocConceptoCod As Integer, _
        ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, ByVal pnColocRecGastoEstado As ColocRecGastoEstado, _
        ByVal psMotivoGasto As String, ByVal pnNewSaldoGasto As Currency, Optional pbEjecBatch As Boolean = False, Optional ByVal pnMovNro As Long = 0)
 
    Call dUpdateColocRecup(psCtaCod, , , , , pnNewSaldoGasto + pnMonto, , , , , , False)
    Call dInsertColocRecupGastos(psCtaCod, pnNroGastoCta, pnColocConceptoCod, psFecAsigna, pnMonto, pnMontoPagado, _
            pnColocRecGastoEstado, psMotivoGasto, False, pnMovNro)
    Call dInsertRecupGastoMovFinan(pnMovNro, psCtaCod, pnMonto)
End Sub

Public Sub dInsertRecupGastoMovFinan(ByVal pnMovNro As Long, ByVal psCtaCod As String, ByVal pnSaldoGastoAnt As Currency)
Dim lsSQL As String
Set dbconec = New DConecta
    If dbconec.AbreConexion() Then
      lsSQL = "stp_ins_RecupGastoMovFinan " & pnMovNro & ",'" & psCtaCod & "', " & pnSaldoGastoAnt & " "
      dbconec.Ejecutar lsSQL
    End If
    Set dbconec = Nothing
   Exit Sub
End Sub

Public Sub dUpdateColocRecup(ByVal psCtaCod As String, Optional ByVal psFecIngRecup As String = "@", _
        Optional ByVal pnComisionCod As Integer = "-1", Optional ByVal pnSaldoIntComp As Currency = -1, _
        Optional ByVal pnSaldoIntMor As Currency = -1, Optional ByVal pnSaldoGasto As Currency = -1, _
        Optional ByVal pnIntCompGen As Currency = -1, Optional ByVal psMetLiquid As String = "@", _
        Optional ByVal pnTipCj As Integer = -1, _
        Optional ByVal pnDemanda As Integer = -1, Optional ByVal pnNroCalend As Integer = -1, _
        Optional pbEjecBatch As Boolean = False, Optional ByVal psComentario As String = "", _
        Optional ByVal pnVisMetodo As Integer = -1, Optional ByVal pnVisNegociacion As Integer = -1, _
        Optional ByVal pnVisCancelacion As Integer = -1, _
        Optional ByVal pnIntMoraGen As Currency = -1) 'DAOR 20070809, pnIntMoraGen

Dim lsSQL As String

Dim lsSQL1 As String
Dim nVistoMetodo As Integer
Dim nVisNegociacion As Integer
Dim nVisCancelacion As Integer
Set dbconec = New DConecta

    If dbconec.AbreConexion() Then
            lsSQL = "UPDATE ColocRecup SET "
        
            If psFecIngRecup <> "@" Then
                lsSQL = lsSQL & " dIngRecup = '" & psFecIngRecup & "',"
            End If
            If pnComisionCod <> -1 Then
                lsSQL = lsSQL & " nComisionCod = " & pnComisionCod & ","
            End If
            
            If pnSaldoIntComp = -4 Then ' Si se le envia (-4) disminuye la Cantidad
                lsSQL = lsSQL & " nSaldoIntComp  = nSaldoIntComp - " & pnSaldoIntComp & ","
            ElseIf pnSaldoIntComp = -3 Then ' Si se le envia (-3) aumenta la Cantidad
                lsSQL = lsSQL & " nSaldoIntComp  = nSaldoIntComp + " & pnSaldoIntComp & ","
            ElseIf pnSaldoIntComp <> -1 Then
                lsSQL = lsSQL & " nSaldoIntComp = " & pnSaldoIntComp & ","
            End If
            
            If pnSaldoIntMor = -4 Then ' Si se le envia (-4) disminuye la Cantidad
                lsSQL = lsSQL & " nSaldoIntMor = nSaldoIntMor - " & pnSaldoIntMor & ","
            ElseIf pnSaldoIntMor = -3 Then ' Si se le envia (-3) aumenta la Cantidad
                lsSQL = lsSQL & " nSaldoIntMor = nSaldoIntMor + " & pnSaldoIntMor & ","
            ElseIf pnSaldoIntMor <> -1 Then
                lsSQL = lsSQL & " nSaldoIntMor = " & pnSaldoIntMor & ","
            End If
            
            If pnSaldoGasto = -4 Then ' Si se le envia (-4) disminuye la cantidad
                lsSQL = lsSQL & " nSaldoGasto = nSaldoGasto - " & pnSaldoGasto & ","
            ElseIf pnSaldoGasto = -3 Then ' Si se le envia (-3) aumenta la cantidad
                lsSQL = lsSQL & " nSaldoGasto = nSaldoGasto + " & pnSaldoGasto & ","
            ElseIf pnSaldoGasto <> -1 Then
                lsSQL = lsSQL & " nSaldoGasto = " & pnSaldoGasto & ","
            End If
            
            If pnIntCompGen = -3 Then 'Si se le envia (-3) aumenta la cantidad
                lsSQL = lsSQL & " nIntCompGen = nIntCompGen + " & pnIntCompGen & ","
            ElseIf pnIntCompGen <> -1 Then
                lsSQL = lsSQL & " nIntCompGen  = " & pnIntCompGen & ","
            End If
            
            '**DAOR 20070809 ****************
            If pnIntMoraGen <> -1 Then
                lsSQL = lsSQL & " nIntMoraGen  = " & pnIntMoraGen & ","
            End If
            '********************************
            
            If psMetLiquid <> "@" Then
                lsSQL = lsSQL & " cMetLiquid = '" & psMetLiquid & "',"
            End If
            If pnTipCj <> -1 Then
                lsSQL = lsSQL & " nTipCJ = " & pnTipCj & ","
            End If
            If pnDemanda <> -1 Then
                lsSQL = lsSQL & " nDemanda = " & pnDemanda & ","
            End If
            If pnNroCalend <> -1 Then
                lsSQL = lsSQL & " nNroCalen = " & pnNroCalend & ","
            End If
              
            If pnVisMetodo <> -1 Then
                lsSQL = lsSQL & " nVisMetodo = " & pnVisMetodo & ","
            End If
            If pnVisNegociacion <> -1 Then
                lsSQL = lsSQL & " nVisNegociacion = " & pnVisNegociacion & ","
            End If
            If pnVisCancelacion <> -1 Then
                lsSQL = lsSQL & " nVisCancelacion = " & pnVisCancelacion & ","
            End If
            '---------------------------------------------------------
              
            lsSQL = Left(lsSQL, Len(lsSQL) - 1)
            
            lsSQL = lsSQL & " , cComentario = '" & psComentario & "'"
            lsSQL = lsSQL & " WHERE cCtaCod ='" & psCtaCod & "'"
            
            If pbEjecBatch Then
                dbconec.AdicionaCmdBatch lsSQL
            Else
                dbconec.Ejecutar lsSQL
            End If
       Set dbconec = Nothing
   Exit Sub
End If
End Sub

Public Sub dInsertColocRecupGastos(ByVal psCtaCod As String, ByVal pnNroGastoCta As Integer, _
        ByVal pnPrdConceptoCod As Integer, _
        ByVal psFecAsigna As String, ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, _
        ByVal pnColocRecGastoEstado As ColocRecGastoEstado, ByVal psMotivoGasto As String, _
        Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pnMovNro As Long = 0)

Dim lsSQL As String
Set dbconec = New DConecta
    If dbconec.AbreConexion() Then
            
            lsSQL = "INSERT ColocRecupGastos (cCtaCod,nNroGastoCta, nPrdConceptoCod, dAsigna,nMonto, " _
                & " nMontoPagado, nColocRecGastoEstado,cMotivoGasto,nMovNro ) " _
                & " VALUES ('" & psCtaCod & "'," & pnNroGastoCta & "," & pnPrdConceptoCod & ",'" & psFecAsigna & "'," & pnMonto & "," _
                & pnMontoPagado & "," & pnColocRecGastoEstado & ",'" & psMotivoGasto & "' ," & pnMovNro & " )"
        
            If pbEjecBatch Then
                dbconec.AdicionaCmdBatch lsSQL
            Else
                dbconec.Ejecutar lsSQL
            End If
    End If
    Set dbconec = Nothing
   Exit Sub
End Sub
'END MADM

'*** PEAC 20110530
Public Function ObtieneAgencias() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = " exec stp_sel_Agencias "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set ObtieneAgencias = rs
    Set rs = Nothing
End Function

'*** PEAC 20110530
Public Function ObtieneCtasResumenCredPigno() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = " exec stp_sel_ObtieneCtasResumenCredPigno "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set ObtieneCtasResumenCredPigno = rs
    Set rs = Nothing
End Function

'*** PEAC 20110706
Public Function ObtieneProductos() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = " exec stp_sel_RS_ListaProductos "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set ObtieneProductos = rs
    Set rs = Nothing
End Function

Public Function CargaCtaCont() As Recordset
   On Error GoTo CargaCtaContErr
   Set dbconec = New DConecta
  dbconec.AbreConexion
      psSql = "exec stp_sel_CtaContxOperacion "
  Set CargaCtaCont = dbconec.CargaRecordSet(psSql)
  dbconec.CierraConexion
   
   Set dbconec = Nothing
   Exit Function
CargaCtaContErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeCta Method")
End Function

Public Function CargaOperacionxCtaContOperacion() As Recordset
   On Error GoTo CargaOperacionxCtaContOperacionErr
   Set dbconec = New DConecta
  dbconec.AbreConexion
      psSql = "exec stp_sel_CtaContxOperacion_Operacion "
  Set CargaOperacionxCtaContOperacion = dbconec.CargaRecordSet(psSql)
  dbconec.CierraConexion
   
   Set dbconec = Nothing
   Exit Function
CargaOperacionxCtaContOperacionErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeCta Method")
End Function


Public Function CargaListaOperacionxCtaContOperacion(ByVal psPersCod As String, ByVal psCtaContCod As String, ByVal psCtaIFCod As String) As Recordset
   On Error GoTo CargaListaOperacionxCtaContOperacionErr
   Set dbconec = New DConecta
      dbconec.AbreConexion
      psSql = "exec stp_sel_CtaIFFiltro_Operaciones '" & psPersCod & "','" & psCtaContCod & "','" & psCtaIFCod & "'"
      Set CargaListaOperacionxCtaContOperacion = dbconec.CargaRecordSet(psSql)
      dbconec.CierraConexion
   
   Set dbconec = Nothing
   Exit Function
CargaListaOperacionxCtaContOperacionErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeCta Method")
End Function

Public Function CargaListaOperacionxOpeCta(ByVal psOpecod As String, ByVal psTipoCta As String, ByVal psOrden As String) As Recordset
   On Error GoTo CargaListaOperacionxOpeCtaErr
   Set dbconec = New DConecta
      dbconec.AbreConexion
      psSql = "exec stp_sel_OpeCta_Operaciones '" & psOpecod & "', '" & psTipoCta & "','" & psOrden & "'"
      Set CargaListaOperacionxOpeCta = dbconec.CargaRecordSet(psSql)
      dbconec.CierraConexion
   
   Set dbconec = Nothing
   Exit Function
CargaListaOperacionxOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeCta Method")
End Function


Public Sub RegistraOpeCta(ByVal psOpecod As String, ByVal psOpeCtaOrden As String, ByVal psCtaContCod As String, ByVal psOpeCtaDH As String, ByVal psOpeCtaOpc As String, ByVal psCtaContN As String, ByVal psMovNro As String)
   Dim dbconec As DConecta
     
   On Error GoTo RegistraOpeCtaErr
   Set dbconec = New DConecta
   
        dbconec.AbreConexion
             
        psSql = " exec stp_sel_InsertarOpeCta '" & psOpecod & "', '" & psOpeCtaOrden & "', '" & psCtaContCod & "', '" & psOpeCtaDH & "', '" & psOpeCtaOpc & "', '" & psCtaContN & "','" & psMovNro & "'"
        dbconec.Ejecutar psSql
                    
        dbconec.CierraConexion

   Set dbconec = Nothing
   Exit Sub
RegistraOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:ActualizaFondoSeguDepo Method")
End Sub

Public Sub RegistraCtaFiltro(ByVal psCtaContCod As String, ByVal psPersCod As String, ByVal psCtaIfSubCta As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String)
   Dim dbconec As DConecta
     
   On Error GoTo RegistraCtaFiltroErr
   Set dbconec = New DConecta
   
        dbconec.AbreConexion
             
        psSql = " exec stp_sel_InsertarCtaIFFiltro '" & psCtaContCod & "', '" & psPersCod & "', '" & psCtaIfSubCta & "', '" & psIFTpo & "', '" & psCtaIFCod & "'"
        dbconec.Ejecutar psSql
                    
        dbconec.CierraConexion

   Set dbconec = Nothing
   Exit Sub
RegistraCtaFiltroErr:
   Call RaiseError(MyUnhandledError, "DOperacion:ActualizaFondoSeguDepo Method")
End Sub
Public Sub EliminaCtaIFFiltro(ByVal psCtaContCod As String, ByVal psPersCod As String, ByVal psCtaIfSubCta As String, sCtaIFCod As String)
   Dim dbconec As DConecta
     
   On Error GoTo EliminaCtaIFFiltroErr
   Set dbconec = New DConecta
   
        dbconec.AbreConexion
             
        psSql = " exec stp_del_CtaIFFiltro '" & psCtaContCod & "', '" & psPersCod & "', '" & psCtaIfSubCta & "', '" & sCtaIFCod & "'"
        dbconec.Ejecutar psSql
        dbconec.CierraConexion

   Set dbconec = Nothing
   Exit Sub
EliminaCtaIFFiltroErr:
   Call RaiseError(MyUnhandledError, "DOperacion:ActualizaFondoSeguDepo Method")
End Sub

Public Sub EliminaOpeCtaAdeu(psOpecod As String, psOpeCtaOrden As String, psCtaContCod As String, psDH As String)
   Dim dbconec As DConecta
     
   On Error GoTo EliminaOpeCtaAdeuErr
   Set dbconec = New DConecta
   
        dbconec.AbreConexion
             
        psSql = " exec stp_del_OpeCta '" & psOpecod & "', '" & psOpeCtaOrden & "', '" & psCtaContCod & "','" & psDH & "'"
        dbconec.Ejecutar psSql
                    
        dbconec.CierraConexion

   Set dbconec = Nothing
   Exit Sub
EliminaOpeCtaAdeuErr:
   Call RaiseError(MyUnhandledError, "DOperacion:ActualizaFondoSeguDepo Method")
End Sub

Public Function CargaCtaContSubCta(ByVal psCtaContSC As String) As Recordset
   On Error GoTo CargaCtaContSubCtaErr
   Set dbconec = New DConecta
  dbconec.AbreConexion
      psSql = "exec stp_sel_CtaContxOperacionSubCta '" & psCtaContSC & "'"
  Set CargaCtaContSubCta = dbconec.CargaRecordSet(psSql)
  dbconec.CierraConexion
   
   Set dbconec = Nothing
   Exit Function
CargaCtaContSubCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeCta Method")
End Function

'***Agregado por ELRO el 20111011, según Acta 278-2011/TI-D
Public Function listarCuentasOperacion(ByVal psCtaCod As String) As ADODB.Recordset
    On Error GoTo listarCuentasOperacionErr
    Set dbconec = New DConecta
    Set prs = New ADODB.Recordset
    
    dbconec.AbreConexion
   
    psSql = "SELECT cCtaContCod, cCtaContDesc FROM CtaCont WHERE cCtaContCod LIKE '" & psCtaCod & "_%' and cCtaContCod like '__[12]%'"
    
    Set prs = dbconec.CargaRecordSet(psSql)
    
    If Not prs.EOF Then
        Set listarCuentasOperacion = prs
    End If
    
    dbconec.CierraConexion
    
    Set dbconec = Nothing
    Set prs = Nothing
    
    Exit Function
    
listarCuentasOperacionErr:
   Call RaiseError(MyUnhandledError, "Metodo:listarCuentasOperacion")
End Function

'***Fin Agregado por ELRO**********************************
'ALPA 20120503*********************************************
Public Function EmiteOpeCtaLeasing(ByVal psOpecod As String, ByVal pnOpeOrden As Integer) As String
    On Error GoTo EmiteOpeCtaLeasingErr
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConect As DConecta
    Dim lsListaCtasCont As String
    Dim lsSubCta As String
    Dim oCtaIf As NCajaCtaIF
    Dim oCtaCont As DCtaCont
    Dim oContFunc As NContFunciones
    
    Set oCtaCont = New DCtaCont
    Set oCtaIf = New NCajaCtaIF
    Set oContFunc = New NContFunciones
    
    Set oConect = New DConecta
    Set rs = New ADODB.Recordset
    Dim lnNumCtas As Integer
    Dim lsNuevaCta As String
    Dim lsCuenta As String
    Dim lsAux As String

    If oConect.AbreConexion = False Then Exit Function
    'Probado por ALPA Para Leasing
    sSql = " SELECT cCtaContCod "
    sSql = sSql & " FROM CtaContArrendamiento "
    sSql = sSql & " WHERE  nOrden =" & pnOpeOrden & "  AND cOpeCod = '" & psOpecod & "' "
    lnNumCtas = 0
    Set rs = oConect.CargaRecordSet(sSql)
    
    If Not (rs.BOF Or rs.EOF) Then
        lsListaCtasCont = rs!cCtaContCod
    End If
    rs.Close: Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    EmiteOpeCtaLeasing = lsListaCtasCont
    Set oCtaIf = Nothing
    Set oCtaCont = Nothing
    Exit Function
EmiteOpeCtaLeasingErr:
    Call RaiseError(MyUnhandledError, "DMov:EmiteOpeCod Method")
End Function
'**********************************************************
'*** Agregado por PASI20140327 TI-ERS001-2014
Public Function ObtenerCtasAmortIntangible(ByVal psCodOpe As String) As ADODB.Recordset
    On Error GoTo ErrorObtenerCtaAmortIntangible
    Set dbconec = New DConecta
    Set prs = New ADODB.Recordset
    dbconec.AbreConexion
    psSql = "stp_sel_ERS0012014_ObtenerCtasAmortIntangible " & psCodOpe
    Set prs = dbconec.CargaRecordSet(psSql)
    
    If Not prs.EOF Then
        Set ObtenerCtasAmortIntangible = prs
    End If
    dbconec.CierraConexion
    
    Set dbconec = Nothing
    Set prs = Nothing
    
    Exit Function
ErrorObtenerCtaAmortIntangible:
   Call RaiseError(MyUnhandledError, "Metodo:ObtenerCtaDAmortIntangible")
End Function
'*** End PASI
'PASI20151202*****
Public Function ObtieneCtaHaberxProvInversion(ByVal pnTpoInv As Integer, Optional ByVal nTpoInvSub As Integer = 0) As String
On Error GoTo ERRORObtieneCtaHaberxProxInversion
    Set dbconec = New DConecta
    Set prs = New ADODB.Recordset
    dbconec.AbreConexion
    psSql = "stp_sel_ObtieneCtaHaberxProvInversion " & pnTpoInv & "," & nTpoInvSub
    Set prs = dbconec.CargaRecordSet(psSql)
    If Not prs.EOF Then
        ObtieneCtaHaberxProvInversion = prs!CtaCont
    Else
        ObtieneCtaHaberxProvInversion = ""
    End If
    Set dbconec = Nothing
    Set prs = Nothing
Exit Function
ERRORObtieneCtaHaberxProxInversion:
   Call RaiseError(MyUnhandledError, "Metodo:ObtieneCtaHaberxProvInversion")
End Function
Public Function ObtieneCtaDebexProvInversion(ByVal pnTpoInv As Integer, Optional ByVal nTpoInvSub As Integer = 0) As String
On Error GoTo ERRORObtieneCtaDebexProvInversion
    Set dbconec = New DConecta
    Set prs = New ADODB.Recordset
    dbconec.AbreConexion
    psSql = "stp_sel_ObtieneCtaDebexProvInversion " & pnTpoInv & "," & nTpoInvSub
    Set prs = dbconec.CargaRecordSet(psSql)
    If Not prs.EOF Then
        ObtieneCtaDebexProvInversion = prs!CtaCont
    Else
        ObtieneCtaDebexProvInversion = ""
    End If
    Set dbconec = Nothing
    Set prs = Nothing
Exit Function
ERRORObtieneCtaDebexProvInversion:
   Call RaiseError(MyUnhandledError, "Metodo:ObtieneCtaDebexProvInversion")
End Function
Public Function ObtieneCtaProvxCancelInversion(ByVal pnTpoInv As Integer) As String
On Error GoTo ERRORObtieneCtaProvxCancelInversion
    Set dbconec = New DConecta
    Set prs = New ADODB.Recordset
    dbconec.AbreConexion
    psSql = "stp_sel_ObtieneCtaProvxCancelInversion " & pnTpoInv
    Set prs = dbconec.CargaRecordSet(psSql)
    If Not prs.EOF Then
        ObtieneCtaProvxCancelInversion = prs!CtaCont
    Else
        ObtieneCtaProvxCancelInversion = ""
    End If
    Set dbconec = Nothing
    Set prs = Nothing
Exit Function
ERRORObtieneCtaProvxCancelInversion:
   Call RaiseError(MyUnhandledError, "Metodo:ObtieneCtaProvxCancelInversion")
End Function
Public Function ObtieneCtaInteresesxCancelInversion(ByVal pnTpoInv As Integer) As String
On Error GoTo ERRORObtieneCtaInteresesxCancelInversion
    Set dbconec = New DConecta
    Set prs = New ADODB.Recordset
    dbconec.AbreConexion
    psSql = "stp_sel_ObtieneCtaInteresesxCancelInversion " & pnTpoInv
    Set prs = dbconec.CargaRecordSet(psSql)
    If Not prs.EOF Then
        ObtieneCtaInteresesxCancelInversion = prs!CtaCont
    Else
        ObtieneCtaInteresesxCancelInversion = ""
    End If
    Set dbconec = Nothing
    Set prs = Nothing
Exit Function
ERRORObtieneCtaInteresesxCancelInversion:
   Call RaiseError(MyUnhandledError, "Metodo:ObtieneCtaInteresesxCancelInversion")
End Function
Public Function ObtieneCtaHaberCapxCancelInversion(ByVal pnTpoInv As Integer) As String
On Error GoTo ERRORObtieneCtaHaberCapxCancelInversion
    Set dbconec = New DConecta
    Set prs = New ADODB.Recordset
    dbconec.AbreConexion
    psSql = "stp_sel_ObtieneCtaHaberCapxCancelInversion " & pnTpoInv
    Set prs = dbconec.CargaRecordSet(psSql)
    If Not prs.EOF Then
        ObtieneCtaHaberCapxCancelInversion = prs!CtaCont
    Else
        ObtieneCtaHaberCapxCancelInversion = ""
    End If
    Set dbconec = Nothing
    Set prs = Nothing
Exit Function
ERRORObtieneCtaHaberCapxCancelInversion:
   Call RaiseError(MyUnhandledError, "Metodo:ObtieneCtaHaberCapxCancelInversion")
End Function
'end PASI***
