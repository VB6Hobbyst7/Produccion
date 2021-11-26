VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCtaCont"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A7B5FCF038A"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim dbConec As DConecta
Dim oConec  As DClassIni
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sCentralCom As String

Public Function CargaCtaCont(Optional psFiltroCta As String = "", Optional psTabla As String = "CtaCont", Optional LockType As LockTypeEnum = adLockReadOnly, Optional plAlias As Boolean = False) As Recordset
On Error GoTo CargaCtaContErr
   If psTabla = "CtaCont" Then
      psTabla = sCentralCom & psTabla
   End If
      If Not plAlias Then
         psSql = "SELECT cCtaContCod, cCtaContDesc FROM " & psTabla & IIf(psFiltroCta = "", "", " WHERE " & psFiltroCta)
      Else
         psSql = "SELECT cCtaContCod Código, cCtaContDesc Descripción FROM " & psTabla & IIf(psFiltroCta = "", "", " WHERE " & psFiltroCta)
      End If
      Set CargaCtaCont = dbConec.CargaRecordSet(psSql, LockType)
   Exit Function
CargaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaCont Method")
End Function

Public Function GetCtaContDesc(psCta As String) As String
On Error GoTo CargaCtaContErr
    psSql = "SELECT dbo.GetCtaContDesc(cCtaContCod,2,1) cCtaContDesc " _
          & "FROM CtaCont WHERE cCtaContCod = '" & psCta & "' "
    Set prs = dbConec.CargaRecordSet(psSql)
    If Not prs.EOF Then
        GetCtaContDesc = prs!cCtaContDesc
    End If
   Exit Function
CargaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaCont Method")
End Function

Public Function GetAreaAgencia(ByVal psCta As String) As ADODB.Recordset
On Error GoTo CargaCtaContErr
    psSql = "SELECT isnull(nCtaEstado,0)nCtaEstado,isnull(bAgencia,0)bAgencia " _
          & "FROM CtaCont WHERE cCtaContCod = '" & psCta & "' "
    Set prs = dbConec.CargaRecordSet(psSql)
    If Not prs.EOF Then
        Set GetAreaAgencia = dbConec.CargaRecordSet(psSql)
    End If
   Exit Function
CargaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaCont Method")

End Function

Public Sub InsertaCtaCont(psCtaCod As String, psCtaDesc As String, psMovNro As String, Optional psTabla As String = "CtaCont", Optional psMoneda As String = "")
   On Error GoTo InsertaCtaContErr
      If psMoneda = "" Or psTabla = "CtaContBase" Then
         psSql = "INSERT " & psTabla & " (cCtaContCod, cCtaContDesc, cUltimaActualizacion) " _
               & "VALUES ('" & psCtaCod & "', '" & psCtaDesc & "','" & psMovNro & "')"
         dbConec.Ejecutar psSql
      Else
         If psTabla = sCentralCom & "CtaCont" Then
            psSql = sCentralCom & "sp_InsertaCtaCont '" & psCtaCod & "','" & psCtaDesc & "', '" & psMovNro & "', '" & psMoneda & "'"
            dbConec.Ejecutar psSql
         End If
      End If
   Exit Sub
InsertaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaCtaCont Method")
End Sub

Public Sub InsertaCtaContHC(psCtaCod As String, psCtaDesc As String, psMovNro As String, Optional psTabla As String = "CtaCont", Optional psMoneda As String = "", Optional pbEstado As Integer, Optional pbAgencia As Integer, Optional pEstaGen As Integer, Optional pbHistorico As Integer = 0, Optional psCtaContCodHistorico As String = "")
   On Error GoTo InsertaCtaContErr
      If psMoneda = "" Or psTabla = "CtaContBase" Then
         psSql = "INSERT " & psTabla & " (cCtaContCod, cCtaContDesc, cUltimaActualizacion,nCtaEstado,bAgencia) " _
               & "VALUES ('" & psCtaCod & "', '" & psCtaDesc & "','" & psMovNro & "','" & pbEstado & "','" & pbAgencia & "')"
         dbConec.Ejecutar psSql
      Else
         If psTabla = sCentralCom & "CtaCont" Then
           psSql = sCentralCom & "sp_InsertaCtaContHC '" & psCtaCod & "','" & psCtaDesc & "', '" & psMovNro & "', '" & psMoneda & "','" & pbEstado & "','" & pbAgencia & "','" & pbHistorico & "','" & psCtaContCodHistorico & "'"
            dbConec.Ejecutar psSql
            If pEstaGen = 1 Then
               psSql = "UPDATE  CtaCont  SET nCtaEstado = '" & pbEstado & "' WHERE cCtaContCod LIKE '" & psCtaCod & "%" & "'"
               dbConec.Ejecutar psSql
            End If
         End If
      End If
   Exit Sub
InsertaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaCtaCont Method")
End Sub



Public Sub ActualizaCtaCont(psCtaCod As String, psCtaDesc As String, psMovNro As String, Optional psTabla As String = "CtaCont")
   On Error GoTo ActualizaCtaContErr
         psSql = "UPDATE " & psTabla & " SET cCtaContDesc = '" & psCtaDesc & "', cUltimaActualizacion = '" & psMovNro & "' WHERE cCtaContCod LIKE '" & psCtaCod & "'"
         dbConec.Ejecutar psSql
   Exit Sub
ActualizaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:ActualizaCtaCont Method")
End Sub

Public Sub EliminaCtaCont(psCtaCod As String, Optional psTabla As String = "CtaCont")
   On Error GoTo EliminaCtaContErr
      If psTabla = "CtaCont" Then
         psTabla = sCentralCom & psTabla
      End If
         psSql = "DELETE " & psTabla & " WHERE cCtaContCod LIKE '" & psCtaCod & "'"
         dbConec.Ejecutar psSql
   Exit Sub
EliminaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:EliminaCtaCont Method")
End Sub
Public Function CargaCtaObj(Optional psCtaCod As String = "", Optional psFiltro As String = "", Optional lObjDesc As Boolean = False) As ADODB.Recordset
   On Error GoTo CargaCtaObjErr
      If lObjDesc Then
         psSql = "SELECT co.cCtaContCod, co.nCtaObjOrden, co.cObjetoCod, o.cObjetoDesc, co.nCtaObjNiv, co.cCtaObjFiltro, co.cCtaObjImpre, o.nObjetoNiv FROM " & sCentralCom & "CtaObj co LEFT JOIN " & sCentralCom & "Objeto o ON o.cObjetoCod =  co.cObjetoCod " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "'", IIf(psFiltro <> "", " WHERE " & psFiltro, "")) & " order by co.nctaobjorden  "
      Else
         psSql = "SELECT cCtaContCod, nCtaObjOrden, cObjetoCod, nCtaObjNiv, cCtaObjFiltro, cCtaObjImpre FROM " & sCentralCom & "CtaObj " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "'", IIf(psFiltro <> "", " WHERE " & psFiltro, "")) & "order by nctaobjorden "
      End If
      Set CargaCtaObj = dbConec.CargaRecordSet(psSql)
   Exit Function
CargaCtaObjErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaObj Method")
End Function

Private Function ImpreFiltro(ObjCod As String, ObjFiltro As String) As String
Dim N As Integer
ImpreFiltro = ""
For N = 1 To Len(ObjFiltro)
    If Mid(ObjFiltro, N, 1) = "X" Then
       If Len(ImpreFiltro) = 0 Then
          ImpreFiltro = ""
       End If
       ImpreFiltro = ImpreFiltro & Mid(ObjCod, N, 1)
    End If
Next
End Function

Public Function CargaCtaObjFiltro(Optional psCtaCod As String = "", Optional psObjetoCod As String = "", Optional psFiltro As String = "", Optional lObjDesc As Boolean = False) As ADODB.Recordset
   On Error GoTo CargaCtaObjFiltroErr
      If lObjDesc Then
         psSql = "SELECT co.cCtaContCod, co.nCtaObjOrden, co.cObjetoCod, o.cObjetoDesc, co.cCtaObjSubCta cSubCtaCod FROM " & sCentralCom & "CtaObjFiltro co LEFT JOIN " & sCentralCom & "Objeto o ON o.cObjetoCod =  co.cObjetoCod " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "' ", "") & IIf(psFiltro <> "", IIf(psCtaCod = "", " WHERE ", " AND ") & psFiltro, "") & IIf(psObjetoCod = "", "", " and co.cObjetoCod LIKE '" & psObjetoCod & "%'")
      Else
         psSql = "SELECT cCtaContCod, nCtaObjOrden, cObjetoCod, cCtaObjSubCta cSubCtaCod FROM " & sCentralCom & "CtaObjFiltro co " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "' ", "") & IIf(psFiltro <> "", IIf(psCtaCod = "", " WHERE ", " AND ") & psFiltro, "") & IIf(psObjetoCod = "", "", " and co.cObjetoCod LIKE '" & psObjetoCod & "%'")
      End If
      Set CargaCtaObjFiltro = dbConec.CargaRecordSet(psSql)
   Exit Function
CargaCtaObjFiltroErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaObjFiltro Method")
End Function

Public Sub InsertaCtaObjFiltro(psCtaCod As String, psCtaObjOrden As String, psObjetoCod As String, psCtaObjSubCta As String, psMovNro As String)
   On Error GoTo InsertaCtaObjFiltroErr
      psSql = "INSERT " & sCentralCom & "CtaObjFiltro (cCtaContCod, nCtaObjOrden, cObjetoCod, cCtaObjSubCta, cUltimaActualizacion) " _
            & "VALUES ('" & psCtaCod & "', '" & psCtaObjOrden & "','" & psObjetoCod & "','" & psCtaObjSubCta & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
   Exit Sub
InsertaCtaObjFiltroErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaCtaObjFiltro Method")
End Sub

Public Sub EliminaCtaObjFiltro(psCtaCod As String, Optional psCtaObjOrden As String = "", Optional psObjCod As String = "")
   On Error GoTo EliminaCtaObjFiltroErr
      psSql = "DELETE " & sCentralCom & "CtaObjFiltro WHERE cCtaContCod = '" & psCtaCod & "' " & IIf(psCtaObjOrden = "", "", " and nCtaObjOrden = '" & psCtaObjOrden & "' ") & IIf(psObjCod = "", "", " and cObjetoCod = '" & psObjCod & "'")
      dbConec.Ejecutar psSql
   Exit Sub
EliminaCtaObjFiltroErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:EliminaCtaObjFiltro Method")
End Sub

Public Sub CargaConvertidor()
   On Error GoTo CargaConvertidorErr
   Exit Sub
CargaConvertidorErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaConvertidor Method")
End Sub

Public Sub InsertaConvertidor()
   On Error GoTo InsertaConvertidorErr
   Exit Sub
InsertaConvertidorErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaConvertidor Method")
End Sub

Public Sub ActualizaConvertidor()
   On Error GoTo ActualizaConvertidorErr
   Exit Sub
ActualizaConvertidorErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:ActualizaConvertidor Method")
End Sub

Public Sub EliminaConvertidor()
   On Error GoTo EliminaConvertidorErr
   Exit Sub
EliminaConvertidorErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:EliminaConvertidor Method")
End Sub

Public Function ExisteCuenta(psCtaContCod As String, Optional plEsInstancia As Boolean = False) As Boolean
   On Error GoTo ExisteCuentaErr
   ExisteCuenta = True
      psSql = "SELECT cCtaContCod FROM " & sCentralCom & "CtaCont WHERE nCtaEstado =1 and cCtaContCod LIKE '" & psCtaContCod & "%'"
      Set prs = dbConec.CargaRecordSet(psSql)
      If prs.EOF Then
         ExisteCuenta = False
         Err.Raise 50001, "DCtaCont:ExisteCuenta", "Cuenta Contable no existe...!"
      Else
         If prs!cCtaContCod <> psCtaContCod Then
             ExisteCuenta = False
             Err.Raise 50001, "DCtaCont:ExisteCuenta", "Cuenta Contable no existe...!"
         End If
      End If
      If plEsInstancia Then
         If prs.RecordCount > 1 Then
             ExisteCuenta = False
             Err.Raise 50001, "DCtaCont:ExisteCuenta", "Cuenta Contable no es ultima Instancia"
         End If
      End If
   'ExisteCuenta = True
   Exit Function
ExisteCuentaErr:
    RSClose prs
    
End Function

Public Function ExisteCuentaCad(psCtaContCod As String, Optional plEsInstancia As Boolean = False) As String
   On Error GoTo ExisteCuentaErr
   ExisteCuentaCad = ""
      psSql = "SELECT cCtaContCod FROM " & sCentralCom & "CtaCont WHERE cCtaContCod LIKE '" & psCtaContCod & "%'"
      Set prs = dbConec.CargaRecordSet(psSql)
      If prs.EOF Then
         ExisteCuentaCad = "Cuenta Contable no existe...!"
      Else
         If prs!cCtaContCod <> psCtaContCod Then
             ExisteCuentaCad = "Cuenta Contable no existe...!"
         End If
      End If
      If plEsInstancia Then
         If prs.RecordCount > 1 Then
             ExisteCuentaCad = "Cuenta Contable no es ultima Instancia"
         End If
      End If
   Exit Function
ExisteCuentaErr:
   RSClose prs
   'Call RaiseError(MyUnhandledError, "DCtaCont:ExisteCuentaCad Method")
End Function

Public Function CtaInstancia(psCtaCod As String, Optional psTabla As String = "CtaCont") As Boolean
   On Error GoTo CtaInstancia
   CtaInstancia = True
   If psTabla = "CtaCont" Then
      psTabla = sCentralCom & psTabla
   End If
   psSql = "SELECT cCtaContCod FROM " & psTabla & " WHERE cCtaContCod LIKE '" & psCtaCod & "%'"
      Set prs = dbConec.CargaRecordSet(psSql)
      If prs.RecordCount > 1 Then
         CtaInstancia = False
      End If
   Exit Function
CtaInstancia:
   Call RaiseError(MyUnhandledError, "DCtaCont:CuentaInstancia Method")
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
   Set oConec = New DClassIni
   sCentralCom = oConec.BaseComunes
   Set oConec = Nothing
   Set dbConec = New DConecta
   dbConec.AbreConexion
End Sub

Public Function GetCtaObjFiltro(ByVal psCtaContCod As String, psObjetoCod As String, Optional lbMuestraCta As Boolean = True) As String
    On Error GoTo GetCtaObjFiltroErr
    Dim sql As String
    Dim rs   As New ADODB.Recordset
    Dim oConect As DConecta
    
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
    GetCtaObjFiltro = ""
    sql = "SELECT cCtaObjSubCta FROM " & sCentralCom & "CtaObjFiltro WHERE cCtaContCod = '" & psCtaContCod & "' and cObjetoCod = '" & psObjetoCod & "'"
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF Then
        GetCtaObjFiltro = IIf(lbMuestraCta, psCtaContCod, "") & rs!cCtaObjSubCta
    End If
    If GetCtaObjFiltro = "" Then
        GetCtaObjFiltro = IIf(lbMuestraCta, psCtaContCod, "")
    End If
    Exit Function
GetCtaObjFiltroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GetCtaObjFiltro Method")
End Function
Public Function GetCtaObjetos(ByVal psCtaContCod As String, Optional psFiltro As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim oRHAreas As DActualizaDatosArea
Dim oContFunct As NContFunciones
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo


Set oCtaIf = New NCajaCtaIF
Set oEfect = New Defectivo
'Set oCtaCont = New DCtaCont
Set oContFunct = New NContFunciones
Set oRHAreas = New DActualizaDatosArea

Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set rs = CargaCtaObj(psCtaContCod, psFiltro, True)
If Not rs.EOF And Not rs.BOF Then
    Select Case Val(rs!cObjetoCod)
            Case ObjCMACAgencias
                Set rs1 = oRHAreas.GetAgencias(rs!cCtaObjFiltro)
            Case ObjCMACAgenciaArea
                Set rs1 = oRHAreas.GetAgenciasAreas(rs!cCtaObjFiltro)
            Case ObjCMACArea
                Set rs1 = oRHAreas.GetAreas(rs!cCtaObjFiltro)
            Case ObjEntidadesFinancieras
                Set rs1 = oCtaIf.GetCtasInstFinancieras(rs!cCtaObjFiltro, psCtaContCod)
            Case ObjDescomEfectivo
                Set rs1 = oEfect.GetBilletajes(rs!cCtaObjFiltro)
            Case ObjPersona
                Set rs1 = Nothing
            Case Else
                Set rs1 = oContFunct.GetObjetos(Val(IIf(rs!cCtaObjFiltro = "", rs!cObjetoCod, rs!cCtaObjFiltro)))
        End Select
End If
rs.Close
Set rs = Nothing
Set oContFunct = Nothing
Set oRHAreas = Nothing
Set oCtaIf = Nothing
Set oEfect = Nothing


Set GetCtaObjetos = rs1
End Function
Public Function GetCtaInstancia(psCtaCod As String, Optional psTabla As String = "CtaCont") As ADODB.Recordset
   On Error GoTo GetCtaInstancia
   If psTabla = "CtaCont" Then
      psTabla = sCentralCom & psTabla
   End If
   psSql = "SELECT c.cCtaContCod, c.cCtaContDesc FROM " & psTabla & " c, " & psTabla & " c1 " _
         & "WHERE c.cctacontcod LIKE '" & psCtaCod & "%' and c1.cctacontcod like c.cctacontcod + '%' " _
         & "GROUP BY c.cctacontcod , c.cctacontdesc " _
         & "HAVING count(*) = 1 "
      Set GetCtaInstancia = dbConec.CargaRecordSet(psSql)
   Exit Function
GetCtaInstancia:
   Call RaiseError(MyUnhandledError, "DCtaCont:GetCuentaInstancia Method")
End Function

Public Function CargaCtaContClase(Optional psCtaCod As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
On Error GoTo CargaCtaContClaseErr
      psSql = "SELECT cCtaContCod, cCtaCaracter, nCtaMoneda FROM " & sCentralCom & "CtaContClase " & IIf(psCtaCod = "", "", " WHERE '" & psCtaCod & "' LIKE cCtaContCod + '%' ")
      Set CargaCtaContClase = dbConec.CargaRecordSet(psSql, LockType)
   Exit Function
CargaCtaContClaseErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaContClase Method")
End Function


'**********************
'*** NUEVAS FUNCIONES
'**********************

Public Sub InsertaCtaObj(psCtaCod As String, psCtaObjOrden As String, psObjetoCod As String, pnCtaObjNiv As Integer, psCtaObjFiltro As String, psCtaObjImpre As String, psMovNro As String, Optional pnObjetoNiv As Integer = 0)
Dim lsError As String
Dim lTrans As Boolean
   On Error GoTo InsertaCtaObjErr
      dbConec.BeginTrans
      lTrans = True
      psSql = "INSERT " & sCentralCom & "CtaObj (cCtaContCod, nCtaObjOrden, cObjetoCod, nCtaObjNiv, cCtaObjFiltro, cCtaObjImpre, cUltimaActualizacion) " _
            & "VALUES ('" & psCtaCod & "', '" & psCtaObjOrden & "','" & psObjetoCod & "'," & pnCtaObjNiv & ",'" & psCtaObjFiltro & "','" & psCtaObjImpre & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      If psCtaObjImpre <> "" Then
         Set prs = dbConec.CargaRecordSet(sCentralCom & "spGetTreeObj '" & psObjetoCod & "'," & pnObjetoNiv + pnCtaObjNiv & ",'" & psCtaObjFiltro & "'", adLockReadOnly)
         Do While Not prs.EOF
            If prs!nObjetoNiv = pnCtaObjNiv + pnObjetoNiv Then
               InsertaCtaObjFiltro psCtaCod, psCtaObjOrden, prs!cObjetoCod, ImpreFiltro(prs!cObjetoCod, psCtaObjImpre), psMovNro
               'dbConec.Ejecutar psSql
            End If
            prs.MoveNext
         Loop
         If prs.State = adStateOpen Then prs.Close: Set prs = Nothing
      End If
      dbConec.CommitTrans
      lTrans = False
   Exit Sub
InsertaCtaObjErr:
   lsError = Err.Description
   If lTrans Then
      dbConec.RollbackTrans
   End If
   dbConec.CierraConexion
   Err.Raise Err.Number + MyUnhandledError, "DCtaCont:InsertaCtaObj Method", lsError
End Sub

Public Sub EliminaCtaObj(psCtaCod As String, psObjCod As String)
Dim lsError As String
Dim lsTrans As Boolean
   On Error GoTo EliminaCtaObjErr
      lsTrans = True
      dbConec.BeginTrans
      psSql = "DELETE " & sCentralCom & "CtaObjFiltro WHERE cCtaContCod = '" & psCtaCod & "' and cObjetoCod LIKE '" & psObjCod & "%'"
      dbConec.Ejecutar psSql
      psSql = "DELETE " & sCentralCom & "CtaObj WHERE cCtaContCod = '" & psCtaCod & "' and cObjetoCod = '" & psObjCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CommitTrans
      lsTrans = False
   Exit Sub
EliminaCtaObjErr:
   lsError = Err.Description
   If lsTrans Then
      dbConec.RollbackTrans
   End If
   dbConec.CierraConexion
   Err.Raise Err.Number + MyUnhandledError, "DCtaCont:EliminaCtaObj Method", lsError
End Sub

Public Function CargaCtaAreaAgeFiltro(Optional psCtaCod As String = "", Optional psObjetoCod As String = "", Optional psFiltro As String = "", Optional lObjDesc As Boolean = False) As ADODB.Recordset
   On Error GoTo CargaCtaAreaAgeFiltroErr
      If lObjDesc Then
         psSql = "SELECT co.cCtaContCod, co.cCtaObjOrden, co.cAreaCod+co.cAgeCod cObjetoCod, " _
               & "       ISNULL(ag.cAgeDescripcion,a.cAreaDescripcion) cObjetoDesc, " _
               & "       co.cSubCtaCod " _
               & "FROM CTAAREAAGEFILTRO co LEFT JOIN Areas a ON a.cAreaCod = co.cAreaCod " _
               & "      LEFT JOIN Agencias ag ON ag.cAgeCod = co.cAgeCod " _
               & "WHERE co.cCtaContCod = '" & psCtaCod & "' " & IIf(psObjetoCod = "", "", " and co.cObjetoCod = '" & psObjetoCod & "' ")
      Else
         psSql = "SELECT co.cCtaContCod, co.cCtaObjOrden, co.cAreaCod+co.cAgeCod cObjetoCod " _
               & "       co.cSubCtaCod " _
               & "FROM CTAAREAAGEFILTRO co  " _
               & "WHERE co.cCtaContCod = '" & psCtaCod & "' " & IIf(psObjetoCod = "", "", " and co.cObjetoCod = '" & psObjetoCod & "' ")
      End If
      Set CargaCtaAreaAgeFiltro = dbConec.CargaRecordSet(psSql)
   Exit Function
CargaCtaAreaAgeFiltroErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaAreaAgeFiltro Method")
End Function

Public Sub InsertaCtaAreaAgeObjFiltro(psCtaCod As String, psCtaObjOrden As String, psObjetoCod As String, psAreaCod As String, psAgeCod As String, psSubCtaCod As String, psMovNro As String)
   On Error GoTo InsertaCtaObjFiltroErr
      psSql = "INSERT CtaAreaAgeFiltro (cCtaContCod, cCtaObjOrden, cObjetoCod, cAreaCod, cAgeCod, cSubCtaCod) " _
            & "VALUES ('" & psCtaCod & "', '" & psCtaObjOrden & "','" & psObjetoCod & "','" & psAreaCod & "', '" & psAgeCod & "','" & psSubCtaCod & "')"
      dbConec.Ejecutar psSql
   Exit Sub
InsertaCtaObjFiltroErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaCtaObjFiltro Method")
End Sub

Public Sub EliminaCtaAreaAgeFiltro(psCtaCod As String, Optional psCtaObjOrden As String = "", Optional psObjCod As String = "", Optional psAreaCod As String = "", Optional psAgeCod As String = "")
   On Error GoTo EliminaCtaFiltroErr
      psSql = "DELETE CtaAreaAgeFiltro WHERE cCtaContCod = '" & psCtaCod & "' " & IIf(psCtaObjOrden = "", "", " and cCtaObjOrden = '" & psCtaObjOrden & "' ") & IIf(psObjCod = "", "", " and cObjetoCod = '" & psObjCod & "'") & IIf(psAreaCod = "", "", " and cAreaCod = '" & psAreaCod & "' ") & IIf(psAgeCod = "", "", " and cAgeCod = '" & psAgeCod & "' ")
      dbConec.Ejecutar psSql
   Exit Sub
EliminaCtaFiltroErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:EliminaCtaAreaAgeFiltro Method")
End Sub

Private Sub Class_Terminate()
If Not dbConec Is Nothing Then
    dbConec.CierraConexion
End If
Set dbConec = Nothing
End Sub

Public Function VerificaExisteCuenta(psCtaContCod As String, Optional plEsInstancia As Boolean = False) As String
On Error GoTo ExisteCuentaErr
   VerificaExisteCuenta = ""
      psSql = "SELECT cCtaContCod FROM " & sCentralCom & "CtaCont WHERE cCtaContCod LIKE '" & psCtaContCod & "'" 'NAGL Cambió de "%'" a "'" , para tomar la Subcuenta correspondiente según correo 20190201
      Set prs = dbConec.CargaRecordSet(psSql)
      If prs.EOF Then
         VerificaExisteCuenta = "Cuenta Contable " & psCtaContCod & " no existe...!" & vbCrLf & "Favor Consultar con Contabilidad o comunicarse con Sistemas"
         Exit Function
      Else
         If prs!cCtaContCod <> psCtaContCod Then
             VerificaExisteCuenta = "Cuenta Contable " & psCtaContCod & " no existe...!" & vbCrLf & "Favor Consultar con Contabilidad o de comunicarse con Sistemas"
             Exit Function
         End If
      End If
      If plEsInstancia Then
         If prs.RecordCount > 1 Then
             VerificaExisteCuenta = " Cuenta Contable " & psCtaContCod & " no es esta Configurada en el Ultimo Nivel " & vbCrLf & "Favor Consultar con Contabilidad o  de comunicarse con Sistemas"
             Exit Function
         End If
      End If
   VerificaExisteCuenta = ""
   Exit Function
ExisteCuentaErr:
    RSClose prs
   Call RaiseError(MyUnhandledError, "DCtaCont:ExisteCuenta Method")
End Function

Public Function GetEsUltimoNivel(psCuenta As String, Optional psSubCuenta As String = "") As Boolean
Dim rs As New ADODB.Recordset
Dim sCuenta As String
Dim sSql As String
Dim oCon As New DConecta

On Error GoTo GetEsUltimoNivelErr

sCuenta = psCuenta & psSubCuenta

sSql = "Select cCtaContCod "
sSql = sSql & " From CtaCont "
sSql = sSql & " Where cCtaContCod Like '" & Trim(sCuenta) & "%' And cCtaContCod <> '" & Trim(sCuenta) & "' And nCtaEstado = 1"
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
    If rs.EOF And rs.BOF Then
        RSClose rs
        GetEsUltimoNivel = True
    Else
        RSClose rs
        GetEsUltimoNivel = False
    End If
    
Exit Function
GetEsUltimoNivelErr:
    Call RaiseError(MyUnhandledError, "DACGAdeudados:GetEsUltimoNivel Method")

End Function
Public Sub InsertaCtaContHisto(psCtaCod As String, psCtaDesc As String, psMovNro As String, Optional psTabla As String = "CtaCont", Optional psMoneda As String = "")
    Dim rs As ADODB.Recordset
    Dim nValorVerifica As Integer
    Set rs = New ADODB.Recordset
   On Error GoTo InsertaCtaContHistoErr
       nValorVerifica = 0
       psSql = " select * from CtaContHisto where cCtaContCod='" & psCtaCod & "' and cUltimaActualizacion='" & psMovNro & "'"
       Set rs = dbConec.CargaRecordSet(psSql)
       If Not (rs.BOF Or rs.EOF) Then
        nValorVerifica = 1
       End If
       If nValorVerifica = 0 Then
       psSql = "INSERT CtaContHisto (cCtaContCod, cCtaContDesc, cUltimaActualizacion) " _
               & "VALUES ('" & psCtaCod & "', '" & psCtaDesc & "','" & psMovNro & "')"
       dbConec.Ejecutar psSql
       End If
   Exit Sub
InsertaCtaContHistoErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaCtaContHisto Method")
End Sub

'MAVM 20090810 **************
Public Function DarCuotasPendientes(ByVal sPersCod As String, ByVal sCtaIFCod As String) As ADODB.Recordset
On Error GoTo DarCuotasPendientes
        Dim sSql As String
        sSql = "exec stp_sel_DarCuotasPendientes '" & sPersCod & "', '" & sCtaIFCod & "'"
        Set DarCuotasPendientes = dbConec.CargaRecordSet(sSql)
        Exit Function
DarCuotasPendientes:
        Set DarCuotasPendientes = Nothing
        Err.Raise Err.Number, "Cuotas Pendientes", Err.Description
End Function

Public Function DarPagarexCodPersona(ByVal sPersCod As String) As ADODB.Recordset
On Error GoTo DarPagarexCodPersona
        Dim sSql As String
        sSql = "exec stp_sel_DarPagarexCodPersona '" & sPersCod & "'"
        Set DarPagarexCodPersona = dbConec.CargaRecordSet(sSql)
        Exit Function
DarPagarexCodPersona:
        Set DarPagarexCodPersona = Nothing
        Err.Raise Err.Number, "Pagare Cod Pers", Err.Description
End Function

Public Function DarTotales_ReporteSBS(ByVal sPersCod As String, ByVal sFecha As Date) As ADODB.Recordset
On Error GoTo DarTotales_ReporteSBS
        Dim sSql As String
        sSql = "exec stp_sel_DarTotales_ReporteSBS '" & sPersCod & "', '" & Format(sFecha, "yyyymmdd") & "'"
        Set DarTotales_ReporteSBS = dbConec.CargaRecordSet(sSql)
        Exit Function
DarTotales_ReporteSBS:
        Set DarTotales_ReporteSBS = Nothing
        Err.Raise Err.Number, "Cod Pers Adeudados", Err.Description
End Function

Public Function DarCodPersona_ReporteSBS(Optional sCtaContCod As String, Optional sNombre As String, Optional sAccion As String) As ADODB.Recordset
On Error GoTo DarCodPersona_ReporteSBS
        Dim sSql As String
        sSql = "exec stp_sel_DarCodPersona_ReporteSBS '" & sCtaContCod & "', '" & sNombre & "', '" & sAccion & "'"
        Set DarCodPersona_ReporteSBS = dbConec.CargaRecordSet(sSql)
        Exit Function
DarCodPersona_ReporteSBS:
        Set DarCodPersona_ReporteSBS = Nothing
        Err.Raise Err.Number, "Cod Pers Adeudados", Err.Description
End Function

Public Function DarEntidadesFinac() As ADODB.Recordset
On Error GoTo DarEntidadesFinac
        Dim sSql As String
        sSql = "exec stp_sel_EntidadesFinac_ReporteSBS"
        Set DarEntidadesFinac = dbConec.CargaRecordSet(sSql)
        Exit Function
DarEntidadesFinac:
        Set DarEntidadesFinac = Nothing
        Err.Raise Err.Number, "Cta Cont Adeudados", Err.Description
End Function

Public Function DarCtaContAdeudados() As ADODB.Recordset
On Error GoTo DarCtaContAdeudados
        Dim sSql As String
        sSql = "exec stp_sel_CtaContAdeudados"
        Set DarCtaContAdeudados = dbConec.CargaRecordSet(sSql)
        Exit Function
DarCtaContAdeudados:
        Set DarCtaContAdeudados = Nothing
        Err.Raise Err.Number, "Cta Cont Adeudados", Err.Description
End Function

Public Function DarReporteSBS12_FinanciamientoRecibido(ByVal dsFecha As Date, ByVal sCtaCont As String) As ADODB.Recordset
On Error GoTo DarReporteSBS12_FinanciamientoRecibido
        Dim sSql As String
        sSql = "exec stp_sel_ReporteSBS12_FinanciamientoRecibido '" & Mid(Format(dsFecha, "yyyymmdd"), 5, 2) & "', '" & Mid(Format(dsFecha, "yyyymmdd"), 1, 4) & "', '" & sCtaCont & "'"
        Set DarReporteSBS12_FinanciamientoRecibido = dbConec.CargaRecordSet(sSql)
        Exit Function
DarReporteSBS12_FinanciamientoRecibido:
        Set DarReporteSBS12_FinanciamientoRecibido = Nothing
        Err.Raise Err.Number, "Financiamiento Recibido", Err.Description
End Function
'****************************
'ALPA 20090930*******************************
Public Function FactorAjusteRiesgoOperac(ByVal sBalAnio As String, ByVal sBalMes As String, ByVal nTipo As Integer) As ADODB.Recordset
On Error GoTo FactorAjusteRiesgoOperac
        Dim sSql As String
        sSql = "exec stp_sel_FactorAjusteRiesgoOperac '" & sBalAnio & "', '" & sBalMes & "', '" & nTipo & "'"
        Set FactorAjusteRiesgoOperac = dbConec.CargaRecordSet(sSql)
        Exit Function
FactorAjusteRiesgoOperac:
        Set FactorAjusteRiesgoOperac = Nothing
        Err.Raise Err.Number, "Financiamiento Recibido", Err.Description
End Function

Public Function Reporte3PatrimonioEfec(ByVal sBalAnio As String, ByVal sBalMes As String, ByVal nTipo As Integer) As ADODB.Recordset
On Error GoTo Reporte3PatrimonioEfec
        Dim sSql As String
        sSql = "exec stp_sel_Reporte3Patrimonio '" & sBalAnio & "', '" & sBalMes & "', '" & nTipo & "'"
        Set Reporte3PatrimonioEfec = dbConec.CargaRecordSet(sSql)
        Exit Function
Reporte3PatrimonioEfec:
        Set Reporte3PatrimonioEfec = Nothing
        Err.Raise Err.Number, "Financiamiento Recibido", Err.Description
End Function
Public Function ReporteRiesgoCambiario(ByVal sBalAnio As String, ByVal sBalMes As String) As ADODB.Recordset
On Error GoTo ReporteRiesgoCambiario
        Dim sSql As String
        sSql = "exec stp_sel_ReporteRiesgoCambiario '" & sBalAnio & "', '" & sBalMes & "'"
        Set ReporteRiesgoCambiario = dbConec.CargaRecordSet(sSql)
        Exit Function
ReporteRiesgoCambiario:
        Set ReporteRiesgoCambiario = Nothing
        Err.Raise Err.Number, "Financiamiento Recibido", Err.Description
End Function
'****************************
'ALPA 20091230****************************
Public Function ReporteReparticionGastos(ByVal sFecha As String) As ADODB.Recordset
On Error GoTo ReporteReparticionGastos
        Dim sSql As String
        sSql = "exec stp_sel_DistribucionGastos '" & sFecha & "'"
        Set ReporteReparticionGastos = dbConec.CargaRecordSet(sSql)
        Exit Function
ReporteReparticionGastos:
        Set ReporteReparticionGastos = Nothing
        Err.Raise Err.Number, "Financiamiento Recibido", Err.Description
End Function
'EJVG20130124 ***
Public Function ReporteReparticionGastosNew(ByVal sFecha As String) As ADODB.Recordset
On Error GoTo ErrReporteReparticionGastosNew
    Dim sSql As String
    sSql = "exec stp_sel_DistribucionGastosNew '" & sFecha & "'"
    Set ReporteReparticionGastosNew = dbConec.CargaRecordSet(sSql)
    Exit Function
ErrReporteReparticionGastosNew:
    Set ReporteReparticionGastosNew = Nothing
    Err.Raise Err.Number, "Repartición de Gastos", Err.Description
End Function
'END EJVG ********
Public Function ListarCtaContAge() As ADODB.Recordset
On Error GoTo ListarCtaContAge
        Dim sSql As String
        sSql = "exec stp_sel_CtaContAge "
        Set ListarCtaContAge = dbConec.CargaRecordSet(sSql)
        Exit Function
ListarCtaContAge:
        Set ListarCtaContAge = Nothing
        Err.Raise Err.Number, "Listar CtaContAge", Err.Description
End Function
Public Sub InsertaCtaContAge(psCtaCod As String, psAgeCod As String)
   On Error GoTo InsertaCtaContAge
       psSql = "exec stp_ins_CtaContAge '" & psCtaCod & "', '" & psAgeCod & "'"
       dbConec.Ejecutar psSql
   Exit Sub
InsertaCtaContAge:
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaCtaContHisto Method")
End Sub
Public Sub EliminarCtaContAge(psCtaCod As String, psAgeCod As String)
   On Error GoTo EliminarCtaContAge
       psSql = "exec stp_del_CtaContAge '" & psCtaCod & "'"
       dbConec.Ejecutar psSql
   Exit Sub
EliminarCtaContAge:
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaCtaContHisto Method")
End Sub
'*****************************************

'***PEAC 20110722
Public Function CargaCtaContOpeRecipro(Optional LockType As LockTypeEnum = adLockReadOnly, Optional plAlias As Boolean = False, Optional pcTipoCta As String = "0", Optional pcTipoAhoPF As String = "0") As Recordset
On Error GoTo CargaCtaContErr
   
'   If psTabla = "CtaCont" Then
'      psTabla = sCentralCom & psTabla
'   End If
'      If Not plAlias Then
'         psSql = "SELECT cCtaContCod, cCtaContDesc FROM " & psTabla & IIf(psFiltroCta = "", "", " WHERE " & psFiltroCta)
'      Else
'         psSql = "SELECT cCtaContCod Código, cCtaContDesc Descripción FROM " & psTabla & IIf(psFiltroCta = "", "", " WHERE " & psFiltroCta)
'      End If
            
      psSql = "exec stp_sel_CtaContOpeRecipro '" & pcTipoCta & "','" & pcTipoAhoPF & "'"
      
      Set CargaCtaContOpeRecipro = dbConec.CargaRecordSet(psSql, LockType)
   Exit Function
CargaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaContOpeRecipro Method")
End Function
'ALPA 20111222************************************
Public Sub InsertaRelacionCtaContHist(psCtaContCod As String, psCtaContCodHisto As String)
   On Error GoTo InsertaRelacionCtaContHist
       psSql = "exec stp_ins_RelacionCtaContHist '" & psCtaContCod & "', '" & psCtaContCodHisto & "'"
       dbConec.Ejecutar psSql
   Exit Sub
InsertaRelacionCtaContHist:
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaInsertaRelacionCtaContHist Method")
End Sub
'*************************************************
'ALPA 20111223************************************
Public Function ListarBalanceEstadHist(psCtaContCod As String, psPeriodo As String) As ADODB.Recordset
On Error GoTo ErrorListarBalanceEstadHist
        Dim sSql As String
        sSql = "exec stp_sel_VerificarBalanceEstadHist '" & psCtaContCod & " ','" & psPeriodo & "'"
        Set ListarBalanceEstadHist = dbConec.CargaRecordSet(sSql)
        Exit Function
ErrorListarBalanceEstadHist:
        Set ListarBalanceEstadHist = Nothing
        Err.Raise Err.Number, "Listar BalanceEstadHist", Err.Description
End Function
Public Function ListarRelacionCtaContHist(psCtaContCod As String) As String
On Error GoTo ErrorListarRelacionCtaContHist
        Dim sSql As String
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        sSql = "exec stp_sel_RelacionCtaContHist '" & psCtaContCod & "'"
        Set rs = dbConec.CargaRecordSet(sSql)
        If Not (rs.EOF Or rs.BOF) Then
            ListarRelacionCtaContHist = IIf(IsNull(rs!cCtaContCodHisto), "", rs!cCtaContCodHisto)
        Else
            ListarRelacionCtaContHist = ""
        End If
        Exit Function
ErrorListarRelacionCtaContHist:
        ListarRelacionCtaContHist = ""
        Err.Raise Err.Number, "Listar BalanceEstadHist", Err.Description
End Function
Public Sub EliminaCtaContHistorico(psCtaCod As String, Optional psTabla As String = "RelacionCtaContHist")
   On Error GoTo EliminaCtaContHistoricoErr
      If psTabla = "RelacionCtaContHist" Then
         psTabla = sCentralCom & psTabla
      End If
         psSql = "DELETE " & psTabla & " WHERE cCtaContCod LIKE '" & psCtaCod & "'"
         dbConec.Ejecutar psSql
   Exit Sub
EliminaCtaContHistoricoErr:
   Call RaiseError(MyUnhandledError, "RelacionCtaContHist:EliminaCtaContHistorico Method")
End Sub
'*************************************************
Public Function CargarExposicionesAjustadas2A1(ByVal dFecha As Date) As ADODB.Recordset
On Error GoTo ErrorCargarExposicionesAjustadas2A1
        Dim sSql As String
        sSql = "exec stp_sel_ObtenerExposicionesAjustadas2A1 '" & Format(dFecha, "YYYY/MM/DD") & "'"
        Set CargarExposicionesAjustadas2A1 = dbConec.CargaRecordSet(sSql)
        Exit Function
ErrorCargarExposicionesAjustadas2A1:
        Set CargarExposicionesAjustadas2A1 = Nothing
        Err.Raise Err.Number, "Financiamiento Recibido", Err.Description
End Function

Public Function Reporte3Patrimonio2D(ByVal psFecha As String) As ADODB.Recordset
On Error GoTo Reporte3Patrimonio2DErr
        Dim sSql As String
        sSql = "exec stp_sel_BalanceGenxCodigo '" & psFecha & "'"
        Set Reporte3Patrimonio2D = dbConec.CargaRecordSet(sSql)
        Exit Function
Reporte3Patrimonio2DErr:
        Set Reporte3Patrimonio2D = Nothing
        Err.Raise Err.Number, "Financiamiento Recibido", Err.Description
End Function
'MIOL 20130211 RFC138-2012 ****************************************
Public Function ListarCtaContPosCamb(ByVal pnDig As Integer) As ADODB.Recordset
On Error GoTo ListarCtaContPosCamb
        Dim sSql As String
        sSql = "exec stp_sel_CtaContPosCambiaria '" & pnDig & "'"
        Set ListarCtaContPosCamb = dbConec.CargaRecordSet(sSql)
        Exit Function
ListarCtaContPosCamb:
        Set ListarCtaContPosCamb = Nothing
        Err.Raise Err.Number, "Listar CtaContPosCamb", Err.Description
End Function
'END MIOL *********************************************************
