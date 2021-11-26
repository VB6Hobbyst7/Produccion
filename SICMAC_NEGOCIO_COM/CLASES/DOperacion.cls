VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim DBConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sCentralCom As String

Public Function CargaOpeGru(Optional psOpeGruCod As String, Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaOpeGruErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      psSql = "SELECT cOpeGruCod, cOpeGruDesc, 1 as Nivel FROM " & sCentralCom & "OpeGrupo " & IIf(psOpeGruCod = "", "", "WHERE cOpeGruCod = '" & psOpeGruCod & "'")
      Set CargaOpeGru = DBConec.CargaRecordSet(psSql, LockType)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Sub InsertaOpeGru(psOpeGruCod As String, psOpeGruDesc As String, psMovNro As String)
   On Error GoTo InsertaOpeErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeGrupo (cOpeGruCod, cOpeGruDesc, cUltimaActualizacion) VALUES ('" & psOpeGruCod & "', '" & psOpeGruDesc & "', '" & psMovNro & "') "
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaOpeErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeGru Method")
End Sub

Public Sub ActualizaOpeGru(psOpeGruCod As String, psOpeGruDesc As String, psMovNro As String)
   On Error GoTo InsertaOpeErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      psSql = "UPDATE " & sCentralCom & "OpeGrupo SET cOpeGruDesc = '" & psOpeGruDesc & "', cUltimaActualizacion = '" & psMovNro & "' WHERE cOpeGruCod = '" & psOpeGruCod & "'"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaOpeErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeGru Method")
End Sub

Public Sub EliminaOpeGru(psOpeGruCod As String)
   On Error GoTo InsertaOpeErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeGrupo WHERE cOpeGruCod = '" & psOpeGruCod & "'"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaOpeErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeGru Method")
End Sub

Public Function CargaOpeTpo(psOpeCod As String, Optional plLike As Boolean = False, Optional LockType As LockTypeEnum = adLockReadOnly, Optional pnVisible As Integer = -1, Optional pnNivel As Integer = 0) As Recordset
Dim lsFiltro As String
   On Error GoTo CargaOpeTpoErr
   lsFiltro = ""
   If pnVisible <> -1 Then
        lsFiltro = " and cOpeVisible = '" & pnVisible & "'"
   End If
   If pnNivel > 0 Then
        lsFiltro = lsFiltro & " and nOpeNiv = " & pnNivel & " "
   End If
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      psSql = "SELECT cOpeCod,cOpeDesc,nOpeNiv, CASE WHEN cOpeVisible='" & gOpeTpoVisibleSi & "' THEN '" & gsSI & "' ELSE '" & gsNO & "' END cOpeVisible, ISNULL(cOpeGruCod,'') cOpeGruCod FROM " & sCentralCom & "OpeTpo" & IIf(psOpeCod = "", "", " WHERE cOpeCod " & IIf(plLike, " LIKE '", "= '") & psOpeCod & IIf(plLike, "%'", "'")) & lsFiltro
      Set CargaOpeTpo = DBConec.CargaRecordSet(psSql, LockType)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaOpeTpoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeTpo Method")
End Function

Public Sub InsertaOpeTpo(psOpeCod As String, psOpeDesc As String, psOpeVisible As String, pnOpeNiv As Integer, psOpeTpo As String, psMovNro As String)
   On Error GoTo InsertaOpeTpoErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeTpo (cOpeCod, cOpeDesc, cOpeVisible, nOpeNiv, cOpeGruCod, cUltimaActualizacion) VALUES ('" & psOpeCod & "', '" & Replace(psOpeDesc, "'", "''") & "', '" & psOpeVisible & "', " & pnOpeNiv & ", '" & psOpeTpo & "','" & psMovNro & "') "
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaOpeTpoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeTpo Method")
End Sub

Public Sub ActualizaOpeTpo(psOpeCod As String, psOpeDesc As String, psOpeVisible As String, pnOpeNiv As Integer, psOpeTpo As String, psMovNro As String)
   On Error GoTo ActualziaOpeTpoErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      psSql = "UPDATE " & sCentralCom & "OpeTpo SET cOpeDesc = '" & psOpeDesc & "', cOpeVisible = '" & psOpeVisible & "', nOpeNiv = " & pnOpeNiv & ", cOpeGruCod = '" & psOpeTpo & "', cUltimaActualizacion = '" & psMovNro & "' " _
            & "WHERE cOpeCod = '" & psOpeCod & "'"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
ActualziaOpeTpoErr:
   Call RaiseError(MyUnhandledError, "DOperacion:ActualziaOpeTpo Method")
End Sub

Public Sub EliminaOpeTpo(psOpeCod As String)
Dim lsError As String
Dim lsTrans As Boolean
   On Error GoTo EliminaOpeTpoErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      lsTrans = True
      DBConec.BeginTrans
      psSql = "DELETE " & sCentralCom & "OpeDoc WHERE cOpeCod = '" & psOpeCod & "' "
      DBConec.Ejecutar psSql
      psSql = "DELETE " & sCentralCom & "OpeObj WHERE cOpeCod = '" & psOpeCod & "' "
      DBConec.Ejecutar psSql
      psSql = "DELETE " & sCentralCom & "OpeCta WHERE cOpeCod = '" & psOpeCod & "' "
      DBConec.Ejecutar psSql
      psSql = "DELETE " & sCentralCom & "OpeTpo WHERE cOpeCod = '" & psOpeCod & "' "
      DBConec.Ejecutar psSql
      DBConec.CommitTrans
      lsTrans = False
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
EliminaOpeTpoErr:
   lsError = Err.Description
   If lsTrans Then
      DBConec.RollbackTrans
   End If
   Set DBConec = Nothing
   Err.Raise Err.Number + MyUnhandledError, "DOperación: EliminaOpeTpo", lsError
End Sub
Public Function CargaOpeCta(psOpeCod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "", Optional LockType As LockTypeEnum = adLockReadOnly, Optional pbGetClase As Boolean = False, Optional pnNivSuperior As Integer = 0) As Recordset
   On Error GoTo CargaOpeCtaErr
   Set DBConec = New DConecta
   Dim lsFiltroDH As String
   Dim lsFiltroOrden As String
    If psDH <> "" Then
        lsFiltroDH = " AND a.cOpeCtaDH ='" & psDH & "' "
    End If
    If psOpeCtaOrden <> "" Then
        lsFiltroOrden = " And a.cOpeCtaOrden='" & psOpeCtaOrden & "' "
    End If
   If DBConec.AbreConexion() Then
      psSql = "select " _
         & "   a.cCtaContCod, dbo.GetCtaContDesc(a.cCtaContCod," & pnNivSuperior & ",0) cCtaContDesc, 1 as Nivel ,  " _
         & "  CASE WHEN a.cOpeCtaDH = '" & gsContHaber & "' THEN '" & gsContHaberDesc & "' " _
         & "       WHEN a.cOpeCtaDH = '" & gsContDebe & "' THEN '" & gsContDebeDesc & "' END  as Clase, " _
         & "  CASE WHEN a.cOpeCtaOpc = '" & gOpeCtaCaracObligatorio & "' then '" & gsOpeCtaCaracterObligaDesc & "' WHEN a.cOpeCtaOpc='" & gOpeCtaCaracObligatorio & "' then '" & gsOpeCtaCaracterOpcionDesc & "' END as Caracter, " _
         & "  a.cOpeCod, a.cOpeCtaDH, a.cOpeCtaOpc, a.cCtaContN, a.cOpeCtaOrden " & IIf(pbGetClase, ", cls.cCtaCaracter ", "") _
         & "FROM " & sCentralCom & "OpeCta as a " & IIf(pbGetClase, " JOIN " & sCentralCom & "CtaContClase cls ON a.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' ", "") _
         & "WHERE a.cOpeCod LIKE '" & psOpeCod & "' " & lsFiltroDH & lsFiltroOrden _
         & "order by a.cOpeCtaDH, a.cOpeCtaOrden, a.cCtaContCod "
      Set CargaOpeCta = DBConec.CargaRecordSet(psSql, LockType)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeCta Method")
End Function

Public Function CargaOpeCtaArbol(psOpeCod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "") As Recordset
   On Error GoTo CargaOpeCtaErr
   Set DBConec = New DConecta
   Dim lsFiltroDH As String
   Dim lsFiltroOrden As String
    If psDH <> "" Then
        lsFiltroDH = " AND a.cOpeCtaDH ='" & psDH & "' "
    End If
    If psOpeCtaOrden <> "" Then
        lsFiltroOrden = " And a.cOpeCtaOrden='" & psOpeCtaOrden & "' "
    End If
   If DBConec.AbreConexion() Then
      psSql = "SELECT DISTINCT a.cCtaContCod, b.cCtaContDesc, 1 as nNivel " _
         & "FROM " & sCentralCom & "OpeCta as a JOIN " & sCentralCom & "CtaCont as b ON b.cCtaContCod = a.cCtaContCod " _
         & "WHERE a.cOpeCod LIKE '" & psOpeCod & "%' " & lsFiltroDH & lsFiltroOrden _
         & "ORDER BY a.cCtaContCod "
      Set CargaOpeCtaArbol = DBConec.CargaRecordSet(psSql, adLockReadOnly)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeCta Method")
End Function

Public Sub InsertaOpeCta(psOpeCod As String, psOpeCtaOrden As String, psctacod As String, psOpeCtaDH As String, psOpeCtaOpc As String, psCtaCodN As String, psMovNro As String)
   On Error GoTo InsertaOpeCtaErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeCta (cOpeCod, cOpeCtaOrden , cCtaContCod, cOpeCtaDH, cOpeCtaOpc, cCtaContN, cUltimaActualizacion) VALUES ('" & psOpeCod & "', '" & psOpeCtaOrden & "','" & psctacod & "','" & psOpeCtaDH & "','" & psOpeCtaOpc & "', '" & psCtaCodN & "','" & psMovNro & "')"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeCta Method")
End Sub

Public Sub EliminaOpeCta(psOpeCod As String, psOpeCtaOrden As String, psctacod As String, psOpeCtaDH As String)
   On Error GoTo EliminaOpeCtaErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeCta WHERE cOpeCod = '" & psOpeCod & "' and cOpeCtaOrden = '" & psOpeCtaOrden & "' and cCtaContCod = '" & psctacod & "' and cOpeCtaDH = '" & psOpeCtaDH & "'"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
EliminaOpeCtaErr:
   Call RaiseError(MyUnhandledError, "DOperacion:EliminaOpeCta Method")
End Sub

Public Function CargaOpeObj(psOpeCod As String, Optional LockType As LockTypeEnum = adLockReadOnly, Optional psOpeObjOrden As String = "") As Recordset
   On Error GoTo CargaOpeObjErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      psSql = "SELECT a.cOpeObjOrden ,a.cObjetoCod, b.cObjetoDesc, a.cOpeObjValor, " & _
         "       a.nOpeObjNiv, a.cOpeObjFiltro, a.cOpeCod" & _
         "  FROM  " & sCentralCom & "OpeObj as a JOIN " & sCentralCom & "Objeto as b ON b.cObjetoCod = a.cObjetoCod " & _
         " WHERE a.cOpeCod = '" & psOpeCod & "' " & IIf(psOpeObjOrden = "", "", " and cOpeObjOrden = '" & psOpeObjOrden & "' ")
      Set CargaOpeObj = DBConec.CargaRecordSet(psSql, LockType)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaOpeObjErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeObj Method")
End Function

Public Sub InsertaOpeObj(psOpeCod As String, psOpeObjOrden As String, psObjetoCod As String, psOpeObjValor As String, pnOpeObjNiv As Integer, psOpeObjFiltro As String, psMovNro As String)
   On Error GoTo InsertaOpeObjErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeObj (cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjValor, nOpeObjNiv, cOpeObjFiltro, cUltimaActualizacion) " _
            & "VALUES ('" & psOpeCod & "','" & psOpeObjOrden & "','" & psObjetoCod & "','" & psOpeObjValor & "'," & pnOpeObjNiv & ",'" & psOpeObjFiltro & "','" & psMovNro & "')"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaOpeObjErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeObj Method")
End Sub

Public Sub EliminaOpeObj(psOpeCod As String, psOpeObjOrden As String, psObjetoCod As String)
   On Error GoTo EliminaOpeObjErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeObj WHERE cOpeCod = '" & psOpeCod & "' and cOpeObjOrden = '" & psOpeObjOrden & "' and cObjetoCod = '" & psObjetoCod & "'"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
EliminaOpeObjErr:
   Call RaiseError(MyUnhandledError, "DOperacion:EliminaOpeObj Method")
End Sub

Public Function CargaOpeDoc(psOpeCod As String, Optional pnDocTpo As TpoDoc = -1, Optional psMetodo As OpeDocMetodo = -1, Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaOpeDocErr
   Set DBConec = New DConecta
   Dim lsFiltro As String
   lsFiltro = ""
   If Not pnDocTpo = -1 Then
      lsFiltro = " AND a.nDocTpo = " & pnDocTpo & " "
   End If
   If Not psMetodo = -1 Then
        lsFiltro = lsFiltro & " AND a.cOpeDocMetodo ='" & psMetodo & "'"
   End If
   If DBConec.AbreConexion() Then
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
       
      Set CargaOpeDoc = DBConec.CargaRecordSet(psSql, LockType)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaOpeDocErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeDoc Method")
End Function

Public Sub InsertaOpeDoc(psOpeCod As String, pnDocTpo As Long, psOpeDocEstado As String, psOpeDocMetodo As String, psMovNro As String)
   On Error GoTo InsertaOpeDocErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "OpeDoc (cOpeCod, nDocTpo, cOpeDocEstado, cOpeDocMetodo, cUltimaActualizacion) VALUES ('" & psOpeCod & "', '" & pnDocTpo & "', '" & psOpeDocEstado & "','" & psOpeDocMetodo & "','" & psMovNro & "')"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaOpeDocErr:
   Call RaiseError(MyUnhandledError, "DOperacion:InsertaOpeDoc Method")
End Sub

Public Sub EliminaOpeDoc(psOpeCod As String, pnDocTpo As Long)
   On Error GoTo EliminaOpeDocErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "OpeDoc WHERE cOpeCod = '" & psOpeCod & "' and nDocTpo = " & pnDocTpo & ""
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
EliminaOpeDocErr:
   Call RaiseError(MyUnhandledError, "DOperacion:EliminaOpeDoc Method")
End Sub

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As New COMConecta.DCOMClasIni
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
Dim oConect As DConecta
Set oConect = New DConecta
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
Public Function EmiteOpeDoc(ByVal psOpeCod As String, ByVal psDocTpo As TpoDoc, Optional ByVal psDH As String = "") As String
On Error GoTo ErrorEmiteOpeDoc
Dim rs As ADODB.Recordset
Dim sSql As String
Dim lsFiltroDH As String
Set rs = New ADODB.Recordset

Dim oConect As DConecta
Set oConect = New DConecta
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
    EmiteOpeDoc = rs!copecod
End If
rs.Close: Set rs = Nothing
oConect.CierraConexion
Set oConect = Nothing
Exit Function
ErrorEmiteOpeDoc:
    Err.Raise vbObjectError + 100, "ErrorEmiteDocOpe", "Error emitir documento en Operacion"
End Function
Public Function EmiteOpeCta(ByVal psOpeCod As String, ByVal psDH As String, Optional psOpeCtaOrden As String = "0", Optional ByVal psObjetoCod As String = "", Optional pnFiltro As TpoObjetos, Optional lbMuestraSubCta As Boolean = True) As String
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
    
    sSql = " SELECT  OC.cCtaContCod , OC.cOpeCtaOrden, OC.cOpeCtaDH, CC.cCtaContDesc " _
        & " FROM     " & sCentralCom & "OPECTA OC  JOIN CTACONT CC ON CC.CCTACONTCOD = OC.CCTACONTCOD   " _
        & " WHERE    OC.cOpeCod ='" & psOpeCod & "'  AND OC.cOpeCtaDH = '" & psDH & "' AND OC.cOpeCtaOrden = '" & psOpeCtaOrden & "'"
    lnNumCtas = 0
    Set rs = oConect.CargaRecordSet(sSql)
    lnNumCtas = rs.RecordCount
    lbExisteCuenta = False
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
Public Function EmiteSubCtaOpeObj(ByVal psOpeCod As String, ByVal psCtaContCod As String) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Set oConect = New DConecta
Set rs = New ADODB.Recordset
Dim lsSubCta As String
lsSubCta = ""
If oConect.AbreConexion = False Then Exit Function

sql = "Select  cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjFiltro, " _
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
Public Function GetOpeCtaDebeHaber(ByVal psOpeCod As String, ByVal psCtaContCod As String, Optional ByVal psOpeCtaOrden As String = "0") As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

sql = " SELECT   cOpeCtaDH " _
    & " FROM    " & sCentralCom & "OPECTANEG OC " _
    & " WHERE    COPECOD = '" & psOpeCod & "' and cCtaContcod ='" & psCtaContCod & "'"
GetOpeCtaDebeHaber = ""
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetOpeCtaDebeHaber = rs!cOpeCtaDH
End If
rs.Close
Set rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function EmiteOpeCtasNivel(ByVal psOpeCod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "0") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsDHFiltro As String
Dim lsOrdenFiltro As String
If psDH <> "" Then
    lsDHFiltro = " AND cOpeCtaDH ='" & psDH & "'"
End If
If psOpeCtaOrden <> "" Then
   lsOrdenFiltro = " and OC.cOpeCtaOrden='" & psOpeCtaOrden & "'"
End If

sql = "SELECT   OC.cCtaContCod , ISNULL(dbo.GetCtaContDesc(OC.cCtaContCod,1,1),'') cCtaContDesc, 1 as Nivel " _
    & "FROM     " & sCentralCom & "OPECTANEG OC " _
    & "         " _
    & " WHERE   COPECOD = '" & psOpeCod & "' " & lsDHFiltro

Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function

Set rs = oConect.CargaRecordSet(sql)
Set EmiteOpeCtasNivel = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetOpeObj(ByVal psOpeCod As String, ByVal psOpeObjOrden As String, _
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
    & " WHERE COPECOD ='" & psOpeCod & "' and cOpeObjOrden='" & psOpeObjOrden & "'"
    
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
                    Set rs1 = oCtaIf.CargaCtasIF(Mid(psOpeCod, 3, 1), lsFiltro, rs!nOpeObjNiv)
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
Public Function GetRsOpeObj(ByVal psOpeCod As String, ByVal psOpeObjOrden As String, Optional ByVal psFiltroObj As String, Optional ByVal psFiltroAdd As String = "") As ADODB.Recordset
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
    & " WHERE COPECOD ='" & psOpeCod & "' and cOpeObjOrden='" & psOpeObjOrden & "'"
    
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
oConect.CierraConexion: Set oConect = Nothing

Set oCtaIf = Nothing
Set oContFunct = Nothing
Set oConect = Nothing
Set oRHAreas = Nothing
Set oEfect = Nothing

Set GetRsOpeObj = rs1

End Function

Public Function CargaListaCuentasOperacion(psOpeCod As String, psctacod As String) As String
On Error GoTo CargaListaCuentasOperacionErr
Dim lsCtas   As String
   Set DBConec = New DConecta
   DBConec.AbreConexion
   psSql = "SELECT cCtaContCod FROM " & sCentralCom & "CtaCont WHERE cCtaContCod LIKE '" & psctacod & "_%' and cCtaContCod like '__[12]%'"
   Set prs = DBConec.CargaRecordSet(psSql)
   If Not prs.EOF Then
      Do While Not prs.EOF
         lsCtas = lsCtas & ",'" & prs!cCtaContCod & "'"
         prs.MoveNext
      Loop
   Else
      lsCtas = lsCtas & ",'" & psctacod & "'"
   End If
   CargaListaCuentasOperacion = Mid(lsCtas, 2, Len(lsCtas))
   DBConec.CierraConexion
   Set DBConec = Nothing
   Exit Function
CargaListaCuentasOperacionErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaListaCuentasOperacion Method")
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
   Call RaiseError(MyUnhandledError, "DOperacion:CargaListaDocsOperacion Method")
End Function

Public Function CargaOpeCtaUltimoNivel(psOpeCod As String, Optional ByVal psDH As String = "", Optional ByVal psOpeCtaOrden As String = "") As Recordset
   On Error GoTo CargaOpeCtaUltimoNivelErr
   Set DBConec = New DConecta
   Dim lsFiltroDH As String
   Dim lsFiltroOrden As String
    If psDH <> "" Then
        lsFiltroDH = " AND a.cOpeCtaDH ='" & psDH & "' "
    End If
    If psOpeCtaOrden <> "" Then
        lsFiltroOrden = " And a.cOpeCtaOrden='" & psOpeCtaOrden & "' "
    End If
   If DBConec.AbreConexion() Then
      psSql = "SELECT b.cCtaContCod, b.cCtaContDesc " _
         & "FROM " & sCentralCom & "OpeCta as a JOIN " & sCentralCom & "CtaCont as b ON b.cCtaContCod LIKE a.cCtaContCod+'%' " _
         & "WHERE a.cOpeCod LIKE '" & psOpeCod & "%' " & lsFiltroDH & lsFiltroOrden _
         & "ORDER BY b.cCtaContCod "
      Set CargaOpeCtaUltimoNivel = DBConec.CargaRecordSet(psSql, adLockReadOnly)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaOpeCtaUltimoNivelErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeCtaUltimoNivel Method")
End Function


Public Function GetFiltroObjetos(ByVal pnTipoObj As TpoObjetos, ByVal psCtaContCod As String, ByVal psObjetoCod As String, Optional lbMuestraCta As Boolean = True, Optional lbExisteFiltro As Boolean = False) As String
    On Error GoTo GetCtaObjFiltroErr
    Dim sql As String
    Dim rs   As New ADODB.Recordset
    Dim oConect As DConecta
    Dim lsPersCtaIF As String
    Dim lsCtaIf As String
    Dim lsTpoIf As String
    
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
    GetFiltroObjetos = ""
    Select Case pnTipoObj
        Case ObjCMACAgencias
            sql = "SELECT cAgeCod,  cSubCtaCod as SubCta  FROM Agencias where cAgeCod='" & psObjetoCod & "'"
        Case ObjCMACArea
            sql = "SELECT cAreaCod, cSubCtaCod  as SubCta  FROM AREAS where cAreaCod='" & psObjetoCod & "'"
        Case ObjCMACAgenciaArea
           sql = "     SELECT   AA.cAreaCod, AA.cAgeCod, " _
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
                sql = "select cPersCod, cSubCtaContcod SubCta FROM InstitucionFinanc where cPersCod = '" & psObjetoCod & "' "
            Else
                sql = "Select cPersCod, cCtaIfCod, cCtaContCod, cCtaIFSubCta as SubCta From CtaIFFiltro WHERE  cPersCod = '" & lsPersCtaIF & "' and cIFTpo='" & lsTpoIf & "' AND cCtaIfCod='" & lsCtaIf & "' AND cCtaContCod in ('" & psCtaContCod & "')"
            End If
        Case Else
            sql = "SELECT cCtaObjSubCta as SubCta FROM CtaObjFiltro WHERE cCtaContCod in ('" & psCtaContCod & "') and cObjetoCod = '" & psObjetoCod & "'"
    End Select
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF Then
        lbExisteFiltro = True
        GetFiltroObjetos = IIf(lbMuestraCta, psCtaContCod, "") & rs!SubCta
    End If
    If GetFiltroObjetos = "" Then
        GetFiltroObjetos = IIf(lbMuestraCta, psCtaContCod, "")
    End If
    Exit Function
GetCtaObjFiltroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GetCtaObjFiltro Method")
End Function

'***Agregado por ELRO 20111021, según Acta 277-2011/TI-D
Public Function listarOperacionCtaCta(ByVal psOpeCtaDH As String, _
                            ByVal psMoneda As String, _
                            Optional ByVal psOpeCod As String = "") As ADODB.Recordset
    On Error GoTo listarOperacionCtaCtaErr
    
    Dim sql As String
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
        
    If oConect.AbreConexion = False Then
        Exit Function
    End If
    
    sql = " exec stp_sel_ListarOperacionCuentaContable '" & psOpeCtaDH & "', '" & psMoneda & "', '" & psOpeCod & "' "
    Set listarOperacionCtaCta = oConect.CargaRecordSet(sql)
    
    oConect.CierraConexion: Set oConect = Nothing
    Exit Function
    
listarOperacionCtaCtaErr:
    Call RaiseError(MyUnhandledError, "DOperacion:listarOperacionCtaCta Method")

End Function

Public Function listarCuentasEntidadesFinacieras(ByVal psComparar As String, _
                                                 ByVal psMoneda As String) As ADODB.Recordset
    On Error GoTo listarCuentasEntidadesFinacierasErr
    Dim sql As String
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
    
    If oConect.AbreConexion = False Then
        Exit Function
    End If
    
    sql = "exec stp_sel_ListarEntidadFinanciera '" & psComparar & "', '" & psMoneda & "' "
    Set listarCuentasEntidadesFinacieras = oConect.CargaRecordSet(sql)
    
    oConect.CierraConexion: Set oConect = Nothing
    Exit Function
    
listarCuentasEntidadesFinacierasErr:
    Call RaiseError(MyUnhandledError, "DOperacion:listarCuentasEntidadesFinacieras Method")

End Function

Public Function recuperaTipoCuentaEntidadFinaciera(ByVal psCodCtaIf As String) As String

    On Error GoTo recuperaTipoCuentaEntidadFinacieraErr
    
    Dim sql As String
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
    Dim rsTipoCuenta As ADODB.Recordset
    Set rsTipoCuenta = New ADODB.Recordset
    
    If oConect.AbreConexion = False Then
        Exit Function
    End If
    
    sql = "exec stp_sel_RecuperarTipoCuentaEntidadFinanciera '" & psCodCtaIf & "' "
    Set rsTipoCuenta = oConect.CargaRecordSet(sql)
    
    If Not rsTipoCuenta.EOF Then
         recuperaTipoCuentaEntidadFinaciera = rsTipoCuenta.Fields(1)
    End If
    
    oConect.CierraConexion: Set oConect = Nothing
    Set rsTipoCuenta = Nothing
    Exit Function

recuperaTipoCuentaEntidadFinacieraErr:
    Call RaiseError(MyUnhandledError, "DOperacion:recuperaTipoCuentaEntidadFinaciera Method")

End Function

'***Fin Agregado por ELRO*******************************

