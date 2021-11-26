VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCtaCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim dbConec As COMConecta.DCOMConecta
Dim oConec  As COMConecta.DCOMClasIni
Dim oError As New COMConecta.COMErrorHandling
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sConexion As String
Dim sCentralCom As String

Public Function CargaCtaCont(Optional psFiltroCta As String = "", Optional psTabla As String = "CtaCont", Optional LockType As LockTypeEnum = adLockReadOnly, Optional plAlias As Boolean = False) As ADODB.Recordset
   On Error GoTo CargaCtaContErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
      If Not plAlias Then
         psSql = "SELECT cCtaContCod, cCtaContDesc FROM " & psTabla & IIf(psFiltroCta = "", "", " WHERE " & psFiltroCta)
      Else
         psSql = "SELECT cCtaContCod Código, cCtaContDesc Descripción FROM " & psTabla & IIf(psFiltroCta = "", "", " WHERE " & psFiltroCta)
      End If
      Set CargaCtaCont = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaCtaContErr:
'   Err.Raise vbObjectError + Err.Number, "DCtaCont:CargaCtaCont Method", Err.Description
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:CargaCtaCont Method")
End Function

Public Sub InsertaCtaCont(psCtaCod As String, psCtaDesc As String, psMovNro As String, Optional psTabla As String = "CtaCont", Optional psMoneda As String = "")
   On Error GoTo InsertaCtaContErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
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
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaCtaContErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:InsertaCtaCont Method")
End Sub

Public Sub ActualizaCtaCont(psCtaCod As String, psCtaDesc As String, psMovNro As String, Optional psTabla As String = "CtaCont")
   On Error GoTo ActualizaCtaContErr
      Set dbConec = New COMConecta.DCOMConecta
      dbConec.AbreConexion sConexion
         psSql = "UPDATE " & psTabla & " SET cCtaContDesc = '" & psCtaDesc & "', cUltimaActualizacion = '" & psMovNro & "' WHERE cCtaContCod LIKE '" & psCtaCod & "'"
         dbConec.Ejecutar psSql
      dbConec.CierraConexion
      Set dbConec = Nothing
   Exit Sub
ActualizaCtaContErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:ActualizaCtaCont Method")
End Sub

Public Sub EliminaCtaCont(psCtaCod As String, Optional psTabla As String = "CtaCont")
   On Error GoTo EliminaCtaContErr
      Set dbConec = New COMConecta.DCOMConecta
      dbConec.AbreConexion sConexion
         psSql = "DELETE " & psTabla & " WHERE cCtaContCod LIKE '" & psCtaCod & "'"
         dbConec.Ejecutar psSql
      dbConec.CierraConexion
      Set dbConec = Nothing
   Exit Sub
EliminaCtaContErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:EliminaCtaCont Method")
End Sub

Public Function CargaCtaObj(Optional psCtaCod As String = "", Optional psFiltro As String = "", Optional lObjDesc As Boolean = False) As ADODB.Recordset
   On Error GoTo CargaCtaObjErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
      If lObjDesc Then
         psSql = "SELECT co.cCtaContCod, co.nCtaObjOrden, co.cObjetoCod, o.cObjetoDesc, co.nCtaObjNiv, co.cCtaObjFiltro, co.cCtaObjImpre, o.nObjetoNiv FROM " & sCentralCom & "CtaObj co LEFT JOIN " & sCentralCom & "Objeto o ON o.cObjetoCod =  co.cObjetoCod " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "'", IIf(psFiltro <> "", " WHERE " & psFiltro, ""))
      Else
         psSql = "SELECT cCtaContCod, nCtaObjOrden, cObjetoCod, nCtaObjNiv, cCtaObjFiltro, cCtaObjImpre FROM " & sCentralCom & "CtaObj " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "'", IIf(psFiltro <> "", " WHERE " & psFiltro, ""))
      End If
      Set CargaCtaObj = dbConec.CargaRecordSet(psSql)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaCtaObjErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:CargaCtaObj Method")
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

Public Sub InsertaCtaObj(psCtaCod As String, psCtaObjOrden As String, psObjetoCod As String, pnCtaObjNiv As Integer, psCtaObjFiltro As String, psCtaObjImpre As String, psMovNro As String, Optional pnObjetoNiv As Integer = 0)
Dim lTrans As Boolean
   On Error GoTo InsertaCtaObjErr
   Set dbConec = New COMConecta.DCOMConecta
   
   If dbConec.AbreConexion(sConexion) Then
      dbConec.BeginTrans
      lTrans = True
      psSql = "INSERT " & sCentralCom & "CtaObj (cCtaContCod, cCtaObjOrden, cObjetoCod, nCtaObjNiv, cCtaObjFiltro, cCtaObjImpre, cUltimaActualizacion) " _
            & "VALUES ('" & psCtaCod & "', '" & psCtaObjOrden & "','" & psObjetoCod & "'," & pnCtaObjNiv + pnObjetoNiv & ",'" & psCtaObjFiltro & "','" & psCtaObjImpre & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      If psCtaObjImpre <> "" Then
         Set prs = dbConec.CargaRecordSet(sCentralCom & "spGetTreeObj '" & psObjetoCod & "'," & pnCtaObjNiv & ",'" & psCtaObjFiltro & "'", adLockReadOnly)
         Do While Not prs.EOF
            If prs!nObjetoNiv = pnCtaObjNiv + pnObjetoNiv Then
               InsertaCtaObjFiltro psCtaCod, psCtaObjOrden, prs!cObjetoCod, ImpreFiltro(prs!cObjetoCod, psCtaObjImpre), psMovNro
               dbConec.Ejecutar psSql
            End If
            prs.MoveNext
         Loop
         If prs.State = adStateOpen Then prs.Close: Set prs = Nothing
      End If
      dbConec.CommitTrans
      lTrans = False
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaCtaObjErr:
   If lTrans Then
      dbConec.RollbackTrans
   End If
   dbConec.CierraConexion
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:InsertaCtaObj Method")
End Sub

Public Sub EliminaCtaObj(psCtaCod As String, psObjCod As String)
   On Error GoTo EliminaCtaObjErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
      psSql = "DELETE " & sCentralCom & "CtaObjFiltro WHERE cCtaContCod = '" & psCtaCod & "' and cObjetoCod LIKE '" & psObjCod & "%'"
      dbConec.AdicionaCmdBatch psSql
      psSql = "DELETE " & sCentralCom & "CtaObj WHERE cCtaContCod = '" & psCtaCod & "' and cObjetoCod = '" & psObjCod & "'"
      dbConec.AdicionaCmdBatch psSql
      dbConec.EjecutarBatch
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaCtaObjErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:GrabaCtaObj Method")
End Sub

Public Function CargaCtaObjFiltro(Optional psCtaCod As String = "", Optional psObjetoCod As String = "", Optional psFiltro As String = "", Optional lObjDesc As Boolean = False) As ADODB.Recordset
   On Error GoTo CargaCtaObjFiltroErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
      If lObjDesc Then
         psSql = "SELECT co.cCtaContCod, co.cCtaObjOrden, co.cObjetoCod, o.cObjetoDesc, co.cCtaObjSubCta FROM " & sCentralCom & "CtaObjFiltro co LEFT JOIN " & sCentralCom & "Objeto o ON o.cObjetoCod =  co.cObjetoCod " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "'", IIf(psFiltro <> "", " WHERE " & psFiltro, "")) & IIf(psObjetoCod = "", "", " and co.cObjetoCod LIKE '" & psObjetoCod & "%'")
      Else
         psSql = "SELECT cCtaContCod, cCtaObjOrden, cObjetoCod, cCtaObjSubCta FROM " & sCentralCom & "CtaObjFiltro " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "'", IIf(psFiltro <> "", " WHERE " & psFiltro, "")) & IIf(psObjetoCod = "", "", " and cObjetoCod LIKE '" & psObjetoCod & "%'")
      End If
      Set CargaCtaObjFiltro = dbConec.CargaRecordSet(psSql)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaCtaObjFiltroErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:CargaCtaObjFiltro Method")
End Function

Public Sub InsertaCtaObjFiltro(psCtaCod As String, psCtaObjOrden As String, psObjetoCod As String, psCtaObjSubCta As String, psMovNro As String)
   On Error GoTo InsertaCtaObjFiltroErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
      psSql = "INSERT " & sCentralCom & "CtaObjFiltro (cCtaContCod, cCtaObjOrden, cObjetoCod, cCtaObjSubCta, cUltimaActualizacion) " _
            & "VALUES ('" & psCtaCod & "', '" & psCtaObjOrden & "','" & psObjetoCod & "','" & psCtaObjSubCta & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaCtaObjFiltroErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:InsertaCtaObjFiltro Method")
End Sub

Public Sub EliminaCtaObjFiltro(psCtaCod As String, Optional psCtaObjOrden As String = "", Optional psObjCod As String = "")
   On Error GoTo EliminaCtaObjFiltroErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
      psSql = "DELETE " & sCentralCom & "CtaObjFiltro WHERE cCtaContCod = '" & psCtaCod & "' " & IIf(psCtaObjOrden = "", "", " and cCtaObjOrden = '" & psCtaObjOrden & "' ") & IIf(psObjCod = "", "", " and cObjetoCod = '" & psObjCod & "'")
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaCtaObjFiltroErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:EliminaCtaObjFiltro Method")
End Sub

Public Sub CargaConvertidor()
   On Error GoTo CargaConvertidorErr
   Exit Sub
CargaConvertidorErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:CargaConvertidor Method")
End Sub

Public Sub InsertaConvertidor()
   On Error GoTo InsertaConvertidorErr
   Exit Sub
InsertaConvertidorErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:InsertaConvertidor Method")
End Sub

Public Sub ActualizaConvertidor()
   On Error GoTo ActualizaConvertidorErr
   Exit Sub
ActualizaConvertidorErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:ActualizaConvertidor Method")
End Sub

Public Sub EliminaConvertidor()
   On Error GoTo EliminaConvertidorErr
   Exit Sub
EliminaConvertidorErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:EliminaConvertidor Method")
End Sub

Public Function ExisteCuenta(psCtaContCod As String, Optional plEsInstancia As Boolean = False) As Boolean
   On Error GoTo ExisteCuentaErr
   Set dbConec = New COMConecta.DCOMConecta
   ExisteCuenta = False
   If dbConec.AbreConexion(sConexion) Then
      psSql = "SELECT cCtaContCod FROM " & sCentralCom & "CtaCont WHERE cCtaContCod LIKE '" & psCtaContCod & "%'"
      Set prs = dbConec.CargaRecordSet(psSql)
      If prs.EOF Then
         'MsgBox "Cuenta Contable no existe...!", vbInformation, "!Aviso!"
      Else
         If prs!cCtaContCod <> psCtaContCod Then
         '   MsgBox "Cuenta Contable no existe...!", vbInformation, "!Aviso!"
         End If
      End If
      If plEsInstancia Then
         If prs.RecordCount > 1 Then
         '   MsgBox "Cuenta Contable no es ultima Instancia", vbInformation, "!AViso!"
            prs.Close: Set prs = Nothing
            Exit Function
         End If
      End If
      prs.Close: Set prs = Nothing
   End If
   ExisteCuenta = True
   Exit Function
ExisteCuentaErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:ExisteCuenta Method")
End Function

Public Function GetCtaObjFiltro(psCtaCod As String, psObjCod As String, Optional lAsignaCuenta As Boolean = True) As String
   On Error GoTo GetCtaObjFiltroErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
      psSql = "SELECT cCtaObjSubCta FROM " & sCentralCom & "CtaObj WHERE cCtaContCod = '" & psCtaCod & "' and cObjetoCod = '" & psObjCod & "'"
      Set prs = dbConec.CargaRecordSet(psSql)
      GetCtaObjFiltro = IIf(lAsignaCuenta, psCtaCod, "") & prs!cCtaObjSubCta
      prs.Close: Set prs = Nothing
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
GetCtaObjFiltroErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:GetCtaObjFiltro Method")
End Function

Public Function CtaInstancia(psCtaCod As String, Optional psTabla As String = "CtaCont") As Boolean
   On Error GoTo CtaInstancia
   CtaInstancia = True
   psSql = "SELECT cCtaContCod FROM " & psTabla & " WHERE cCtaContCod LIKE '" & psCtaCod & "%'"
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
      Set prs = dbConec.CargaRecordSet(psSql)
      If prs.RecordCount > 1 Then
         CtaInstancia = False
      End If
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CtaInstancia:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:CuentaInstancia Method")
End Function

Private Sub Class_Initialize()
   Set oConec = New COMConecta.DCOMClasIni
   sConexion = oConec.CadenaConexion
   sCentralCom = oConec.BaseComunes
   Set oConec = Nothing
End Sub


Public Function CargaCtaContClase(Optional psCtaCod As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
On Error GoTo CargaCtaContClaseErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
      psSql = "SELECT cCtaContCod, cCtaCaracter, nCtaMoneda FROM " & sCentralCom & "CtaContClase " & IIf(psCtaCod = "", "", " WHERE '" & psCtaCod & "' LIKE cCtaContCod + '%' ")
      Set CargaCtaContClase = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaCtaContClaseErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DCtaCont:CargaCtaContClase Method")
End Function


'DE ACA EN ADELANTE SON CLASES PARA EL MANTENIMIENTO DE CTAS CONTABLES

Public Function GetOpeCuenta(psCodOpe As String, psDH As String, psTipPers As String, psConcepto As String) As ADODB.Recordset
    Dim sqlOC As String
    Dim rsOC As ADODB.Recordset
    Set rsOC = New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta

    On Error GoTo ErrorGetOpeCuenta
    sqlOC = "Select cCtaContCod From OpeCtaNeg Where cOpeCod = '" & psCodOpe & "' and cOpeCtaDH = '" & psDH & "' and nPersoneria = " & psTipPers & " And nConcepto = " & psConcepto & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rsOC = oCon.CargaRecordSet(sqlOC)
    oCon.CierraConexion
    Set oCon = Nothing
    Set GetOpeCuenta = rsOC
    
    Exit Function
ErrorGetOpeCuenta:
    Err.Raise Err.Number, "GetOpeCuenta", Err.Description
End Function

Public Function EliminaCtaContMant(ByVal pcOpeCod As String, _
                                    ByVal pcCtaContCod As String, _
                                    ByVal pcOpeCtaDH As String, _
                                    ByVal pnPersoneria As String, _
                                    ByVal pnConcepto As String)
Dim sqlE As String

On Error GoTo ErrorEliminaCtaContMant

Set dbConec = New COMConecta.DCOMConecta

    sqlE = "Delete OpeCtaNeg " _
             & " Where cOpeCod =  '" & pcOpeCod & "'" _
             & " and cCtaContCod =  '" & pcCtaContCod & "'" _
             & " and cOpeCtaDH =  '" & pcOpeCtaDH & "'" _
             & " and nPersoneria =  '" & pnPersoneria & "'" _
             & " And nConcepto =  '" & pnConcepto & "'"
        
     dbConec.AbreConexion
     dbConec.Ejecutar sqlE
     dbConec.CierraConexion
Set dbConec = Nothing
Exit Function

ErrorEliminaCtaContMant:
    Err.Raise Err.Number, "Error EliminaCtaContMant", Err.Description
End Function

Public Function InsertaCtaContMant(ByVal pcOpeCod As String, _
                                    ByVal pnConcepto As String, _
                                    ByVal pnPersoneria As String, _
                                    ByVal pcCtaContCod As String, _
                                    ByVal pcOpeCtaDH As String, _
                                    ByVal pdFecSis As Date, _
                                    ByVal psCodAge As String, _
                                    ByVal psCodUser As String)
Dim Sql As String
Dim oMov As COMDMov.DCOMMov

On Error GoTo ErrorInsertaCtaContMant

Set dbConec = New COMConecta.DCOMConecta
Set oMov = New COMDMov.DCOMMov

    Sql = "Insert OpeCtaNeg (cOpeCod,nConcepto,nPersoneria,cCtaContCod,cOpeCtaDH,cUltimaActualizacion) " _
            & "Values('" & pcOpeCod & "','" & pnConcepto & "','" & pnPersoneria & "','" & pcCtaContCod & "','" & pcOpeCtaDH & "','" & oMov.GeneraMovNro(pdFecSis, psCodAge, psCodUser) & "')"

Set oMov = Nothing

     dbConec.AbreConexion
     dbConec.Ejecutar Sql
     dbConec.CierraConexion
Set dbConec = Nothing
Exit Function

ErrorInsertaCtaContMant:
    Err.Raise Err.Number, "Error InsertaCtaContMant", Err.Description
End Function

Public Function ModificaCtaContMant(ByVal pcOpeCod As String, _
                                    ByVal pnConcepto As String, _
                                    ByVal pnPersoneria As String, _
                                    ByVal pcCtaContCod As String, _
                                    ByVal pcOpeCtaDH As String, _
                                    ByVal pdFecSis As Date, _
                                    ByVal psCodAge As String, _
                                    ByVal psCodUser As String, _
                                    ByVal pcOpeCodNew As String, _
                                    ByVal pnConceptoNew As String, _
                                    ByVal pnPersoneriaNew As String, _
                                    ByVal pcCtaContCodNew As String, _
                                    ByVal pcOpeCtaDHNew As String)
Dim Sql As String
Dim oMov As COMDMov.DCOMMov

On Error GoTo ErrorModificaCtaContMant

Set dbConec = New COMConecta.DCOMConecta
Set oMov = New COMDMov.DCOMMov

Sql = " UPDATE OpeCtaNeg " _
            & " set cOpeCod = '" & pcOpeCod & "'," _
            & " cCtaContCod = '" & pcCtaContCod & "'," _
            & " cOpeCtaDH = '" & pcOpeCtaDH & "'," _
            & " nPersoneria = '" & pnPersoneria & "'," _
            & " nConcepto = '" & pnConcepto & "'," _
            & " cUltimaActualizacion = '" & oMov.GeneraMovNro(pdFecSis, psCodAge, psCodUser) & "'" _
            & " Where cOpeCod =  '" & pcOpeCodNew & "'" _
            & " and cCtaContCod =  '" & pcCtaContCodNew & "'" _
            & " and cOpeCtaDH =  '" & pcOpeCtaDHNew & "'" _
            & " and nPersoneria =  '" & pnPersoneriaNew & "'" _
            & " and nConcepto =  '" & pnConceptoNew & "'"

Set oMov = Nothing

dbConec.AbreConexion
dbConec.Ejecutar Sql
dbConec.CierraConexion
     
Set dbConec = Nothing
Exit Function

ErrorModificaCtaContMant:
    Err.Raise Err.Number, "Error ModificaCtaContMant", Err.Description
End Function

Public Sub CargarControlesCtaContMant(ByRef prs As ADODB.Recordset, _
                                    ByVal pnFiltro As Integer)
    Dim Rs As ADODB.Recordset
    Dim Sql As String
    Set Rs = New ADODB.Recordset
    Dim oCons As COMDConstSistema.NCOMConstSistema
    
    On Error GoTo ErrorCargarControlesCtaContMant
    
    Set dbConec = New COMConecta.DCOMConecta
    
    dbConec.AbreConexion
    
    Sql = "Select convert(varchar(8), nPrdConceptoCod) + ' - ' + cDescripcion, nPrdConceptoCod from ProductoConcepto Order By nPrdConceptoCod"
    Set prs = dbConec.CargaRecordSet(Sql)
    
    dbConec.CierraConexion
    Set dbConec = Nothing
    
    Set oCons = New COMDConstSistema.NCOMConstSistema
    pnFiltro = oCons.LeeConstSistema(103)
    Set oCons = Nothing
    Exit Sub
ErrorCargarControlesCtaContMant:
    Err.Raise Err.Number, "Error CargarControles", Err.Description
End Sub

Public Function CargaOperacion(psCodPro As String) As ADODB.Recordset
    
    Dim sqlOpe As String
    
    On Error GoTo ErrorCargaOperacion
    
    sqlOpe = "Select cOpeCod, cOpeDesc From OpeTpo Where cOpeCod Like '" & psCodPro & "' and RTrim(cOpeCod) like '______' Order by cOpeCod"
    
    Set dbConec = New COMConecta.DCOMConecta
    dbConec.AbreConexion
    Set CargaOperacion = dbConec.CargaRecordSet(sqlOpe)
    dbConec.CierraConexion
    Set dbConec = Nothing
    
    Exit Function
    
ErrorCargaOperacion:
    Err.Raise Err.Number, "Carga Operacion", Err.Description
End Function

Public Function GetAsiento(psCodOpe As String) As ADODB.Recordset
    
    Dim sqlAs As String
    
    On Error GoTo ErrorGetAsiento
    
    sqlAs = "Select cOpeCod, cCtaContCod, cOpeCtaDH, nConcepto, nPersoneria  from OpeCtaNeg where cOpeCod = '" & psCodOpe & "' Order By nPersoneria, nConcepto, cOpeCtaDH  "
    
    Set dbConec = New COMConecta.DCOMConecta
    dbConec.AbreConexion
    Set GetAsiento = dbConec.CargaRecordSet(sqlAs)
    dbConec.CierraConexion
    Set dbConec = Nothing
    Exit Function
    
ErrorGetAsiento:
    Err.Raise Err.Number, "Get Asiento", Err.Description
End Function

Public Function ContaFideicomiso(pdFechaI As Date, ntipoCam As Double) As ADODB.Recordset
    
    Dim sqlAs As String
    
    On Error GoTo ErrorContaFideicomiso
    
    sqlAs = "exec stp_sel_ReporteContaFideicomiso '" & Format(pdFechaI, "YYYYMMDD") & "','" & ntipoCam & "'"
    
    Set dbConec = New COMConecta.DCOMConecta
    dbConec.AbreConexion
    Set ContaFideicomiso = dbConec.CargaRecordSet(sqlAs)
    dbConec.CierraConexion
    Set dbConec = Nothing
    
    Exit Function
    
ErrorContaFideicomiso:
    Err.Raise Err.Number, "Get Asiento", Err.Description
End Function

Public Function ObtenerCtaContBalanceMensual(psCtaContCod As String, pdFecDiaProceso As Date, psMoneda As String, pnTipoCambio As Currency) As Currency
On Error GoTo ObtenerCtaContBalanceMensualErr
   Dim oRs As ADODB.Recordset
   Dim oConec As COMConecta.DCOMConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New COMConecta.DCOMConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ObtenerDatosBalanceMensual '" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'," & pnTipoCambio
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = oRs!nSaldoFinImporte
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerCtaContBalanceMensual = nSaldo
Exit Function
ObtenerCtaContBalanceMensualErr:
  Err.Raise Err.Number, "Get Asiento", Err.Description
End Function
'WIOR 20140110 ****************************************************************
Public Function Reporte3PatrimonioEfec(ByVal sBalAnio As String, ByVal sBalMes As String, ByVal nTipo As Integer) As ADODB.Recordset
On Error GoTo Reporte3PatrimonioEfec
        Dim sSql As String
        sSql = "exec stp_sel_Reporte3Patrimonio '" & sBalAnio & "', '" & sBalMes & "', '" & nTipo & "'"
        
        Set dbConec = New COMConecta.DCOMConecta
        dbConec.AbreConexion
        Set Reporte3PatrimonioEfec = dbConec.CargaRecordSet(sSql)
        dbConec.CierraConexion
        Set dbConec = Nothing
        Exit Function
Reporte3PatrimonioEfec:
        Set Reporte3PatrimonioEfec = Nothing
        Err.Raise Err.Number, "Financiamiento Recibido", Err.Description
End Function
'WIOR FIN *********************************************************************
