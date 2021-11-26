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

Dim DBConec As DConecta
Dim oConec  As ClsIni.ClasIni
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sConexion As String
Dim sCentralCom As String

Public Function CargaCtaCont(Optional psFiltroCta As String = "", Optional psTabla As String = "CtaCont", Optional LockType As LockTypeEnum = adLockReadOnly, Optional plAlias As Boolean = False) As Recordset
   On Error GoTo CargaCtaContErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      If Not plAlias Then
         psSql = "SELECT cCtaContCod, cCtaContDesc FROM " & psTabla & IIf(psFiltroCta = "", "", " WHERE " & psFiltroCta)
      Else
         psSql = "SELECT cCtaContCod Código, cCtaContDesc Descripción FROM " & psTabla & IIf(psFiltroCta = "", "", " WHERE " & psFiltroCta)
      End If
      Set CargaCtaCont = DBConec.CargaRecordSet(psSql, LockType)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaCtaContErr:
'   Err.Raise vbObjectError + Err.Number, "DCtaCont:CargaCtaCont Method", Err.Description
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaCont Method")
End Function

Public Sub InsertaCtaCont(psCtaCod As String, psCtaDesc As String, psMovNro As String, Optional psTabla As String = "CtaCont", Optional psMoneda As String = "")
   On Error GoTo InsertaCtaContErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      If psMoneda = "" Or psTabla = "CtaContBase" Then
         psSql = "INSERT " & psTabla & " (cCtaContCod, cCtaContDesc, cUltimaActualizacion) " _
               & "VALUES ('" & psCtaCod & "', '" & psCtaDesc & "','" & psMovNro & "')"
         DBConec.Ejecutar psSql
      Else
         If psTabla = sCentralCom & "CtaCont" Then
            psSql = sCentralCom & "sp_InsertaCtaCont '" & psCtaCod & "','" & psCtaDesc & "', '" & psMovNro & "', '" & psMoneda & "'"
            DBConec.Ejecutar psSql
         End If
      End If
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaCtaCont Method")
End Sub

Public Sub ActualizaCtaCont(psCtaCod As String, psCtaDesc As String, psMovNro As String, Optional psTabla As String = "CtaCont")
   On Error GoTo ActualizaCtaContErr
      Set DBConec = New DConecta
      DBConec.AbreConexion sConexion
         psSql = "UPDATE " & psTabla & " SET cCtaContDesc = '" & psCtaDesc & "', cUltimaActualizacion = '" & psMovNro & "' WHERE cCtaContCod LIKE '" & psCtaCod & "'"
         DBConec.Ejecutar psSql
      DBConec.CierraConexion
      Set DBConec = Nothing
   Exit Sub
ActualizaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:ActualizaCtaCont Method")
End Sub

Public Sub EliminaCtaCont(psCtaCod As String, Optional psTabla As String = "CtaCont")
   On Error GoTo EliminaCtaContErr
      Set DBConec = New DConecta
      DBConec.AbreConexion sConexion
         psSql = "DELETE " & psTabla & " WHERE cCtaContCod LIKE '" & psCtaCod & "'"
         DBConec.Ejecutar psSql
      DBConec.CierraConexion
      Set DBConec = Nothing
   Exit Sub
EliminaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:EliminaCtaCont Method")
End Sub

Public Function CargaCtaObj(Optional psCtaCod As String = "", Optional psFiltro As String = "", Optional lObjDesc As Boolean = False) As ADODB.Recordset
   On Error GoTo CargaCtaObjErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      If lObjDesc Then
         psSql = "SELECT co.cCtaContCod, co.nCtaObjOrden, co.cObjetoCod, o.cObjetoDesc, co.nCtaObjNiv, co.cCtaObjFiltro, co.cCtaObjImpre, o.nObjetoNiv FROM " & sCentralCom & "CtaObj co LEFT JOIN " & sCentralCom & "Objeto o ON o.cObjetoCod =  co.cObjetoCod " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "'", IIf(psFiltro <> "", " WHERE " & psFiltro, ""))
      Else
         psSql = "SELECT cCtaContCod, nCtaObjOrden, cObjetoCod, nCtaObjNiv, cCtaObjFiltro, cCtaObjImpre FROM " & sCentralCom & "CtaObj " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "'", IIf(psFiltro <> "", " WHERE " & psFiltro, ""))
      End If
      Set CargaCtaObj = DBConec.CargaRecordSet(psSql)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
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

Public Sub InsertaCtaObj(psCtaCod As String, psCtaObjOrden As String, psObjetoCod As String, pnCtaObjNiv As Integer, psCtaObjFiltro As String, psCtaObjImpre As String, psMovNro As String, Optional pnObjetoNiv As Integer = 0)
Dim lTrans As Boolean
   On Error GoTo InsertaCtaObjErr
   Set DBConec = New DConecta
   
   If DBConec.AbreConexion(sConexion) Then
      DBConec.BeginTrans
      lTrans = True
      psSql = "INSERT " & sCentralCom & "CtaObj (cCtaContCod, cCtaObjOrden, cObjetoCod, nCtaObjNiv, cCtaObjFiltro, cCtaObjImpre, cUltimaActualizacion) " _
            & "VALUES ('" & psCtaCod & "', '" & psCtaObjOrden & "','" & psObjetoCod & "'," & pnCtaObjNiv + pnObjetoNiv & ",'" & psCtaObjFiltro & "','" & psCtaObjImpre & "','" & psMovNro & "')"
      DBConec.Ejecutar psSql
      If psCtaObjImpre <> "" Then
         Set prs = DBConec.CargaRecordSet(sCentralCom & "spGetTreeObj '" & psObjetoCod & "'," & pnCtaObjNiv & ",'" & psCtaObjFiltro & "'", adLockReadOnly)
         Do While Not prs.EOF
            If prs!nObjetoNiv = pnCtaObjNiv + pnObjetoNiv Then
               InsertaCtaObjFiltro psCtaCod, psCtaObjOrden, prs!cObjetoCod, ImpreFiltro(prs!cObjetoCod, psCtaObjImpre), psMovNro
               DBConec.Ejecutar psSql
            End If
            prs.MoveNext
         Loop
         If prs.State = adStateOpen Then prs.Close: Set prs = Nothing
      End If
      DBConec.CommitTrans
      lTrans = False
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaCtaObjErr:
   If lTrans Then
      DBConec.RollbackTrans
   End If
   DBConec.CierraConexion
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaCtaObj Method")
End Sub

Public Sub EliminaCtaObj(psCtaCod As String, psObjCod As String)
   On Error GoTo EliminaCtaObjErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "DELETE " & sCentralCom & "CtaObjFiltro WHERE cCtaContCod = '" & psCtaCod & "' and cObjetoCod LIKE '" & psObjCod & "%'"
      DBConec.AdicionaCmdBatch psSql
      psSql = "DELETE " & sCentralCom & "CtaObj WHERE cCtaContCod = '" & psCtaCod & "' and cObjetoCod = '" & psObjCod & "'"
      DBConec.AdicionaCmdBatch psSql
      DBConec.EjecutarBatch
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
EliminaCtaObjErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:GrabaCtaObj Method")
End Sub

Public Function CargaCtaObjFiltro(Optional psCtaCod As String = "", Optional psObjetoCod As String = "", Optional psFiltro As String = "", Optional lObjDesc As Boolean = False) As ADODB.Recordset
   On Error GoTo CargaCtaObjFiltroErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      If lObjDesc Then
         psSql = "SELECT co.cCtaContCod, co.cCtaObjOrden, co.cObjetoCod, o.cObjetoDesc, co.cCtaObjSubCta FROM " & sCentralCom & "CtaObjFiltro co LEFT JOIN " & sCentralCom & "Objeto o ON o.cObjetoCod =  co.cObjetoCod " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "'", IIf(psFiltro <> "", " WHERE " & psFiltro, "")) & IIf(psObjetoCod = "", "", " and co.cObjetoCod LIKE '" & psObjetoCod & "%'")
      Else
         psSql = "SELECT cCtaContCod, cCtaObjOrden, cObjetoCod, cCtaObjSubCta FROM " & sCentralCom & "CtaObjFiltro " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "'", IIf(psFiltro <> "", " WHERE " & psFiltro, "")) & IIf(psObjetoCod = "", "", " and cObjetoCod LIKE '" & psObjetoCod & "%'")
      End If
      Set CargaCtaObjFiltro = DBConec.CargaRecordSet(psSql)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaCtaObjFiltroErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaObjFiltro Method")
End Function

Public Sub InsertaCtaObjFiltro(psCtaCod As String, psCtaObjOrden As String, psObjetoCod As String, psCtaObjSubCta As String, psMovNro As String)
   On Error GoTo InsertaCtaObjFiltroErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "INSERT " & sCentralCom & "CtaObjFiltro (cCtaContCod, cCtaObjOrden, cObjetoCod, cCtaObjSubCta, cUltimaActualizacion) " _
            & "VALUES ('" & psCtaCod & "', '" & psCtaObjOrden & "','" & psObjetoCod & "','" & psCtaObjSubCta & "','" & psMovNro & "')"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaCtaObjFiltroErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaCtaObjFiltro Method")
End Sub

Public Sub EliminaCtaObjFiltro(psCtaCod As String, Optional psCtaObjOrden As String = "", Optional psObjCod As String = "")
   On Error GoTo EliminaCtaObjFiltroErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "DELETE " & sCentralCom & "CtaObjFiltro WHERE cCtaContCod = '" & psCtaCod & "' " & IIf(psCtaObjOrden = "", "", " and cCtaObjOrden = '" & psCtaObjOrden & "' ") & IIf(psObjCod = "", "", " and cObjetoCod = '" & psObjCod & "'")
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
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
   Set DBConec = New DConecta
   ExisteCuenta = False
   If DBConec.AbreConexion(sConexion) Then
      psSql = "SELECT cCtaContCod FROM " & sCentralCom & "CtaCont WHERE cCtaContCod LIKE '" & psCtaContCod & "%'"
      Set prs = DBConec.CargaRecordSet(psSql)
      If prs.EOF Then
         MsgBox "Cuenta Contable no existe...!", vbInformation, "!Aviso!"
      Else
         If prs!cCtaContCod <> psCtaContCod Then
            MsgBox "Cuenta Contable no existe...!", vbInformation, "!Aviso!"
         End If
      End If
      If plEsInstancia Then
         If prs.RecordCount > 1 Then
            MsgBox "Cuenta Contable no es ultima Instancia", vbInformation, "!AViso!"
            prs.Close: Set prs = Nothing
            Exit Function
         End If
      End If
      prs.Close: Set prs = Nothing
   End If
   ExisteCuenta = True
   Exit Function
ExisteCuentaErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:ExisteCuenta Method")
End Function

Public Function GetCtaObjFiltro(psCtaCod As String, psObjCod As String, Optional lAsignaCuenta As Boolean = True) As String
   On Error GoTo GetCtaObjFiltroErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "SELECT cCtaObjSubCta FROM " & sCentralCom & "CtaObj WHERE cCtaContCod = '" & psCtaCod & "' and cObjetoCod = '" & psObjCod & "'"
      Set prs = DBConec.CargaRecordSet(psSql)
      GetCtaObjFiltro = IIf(lAsignaCuenta, psCtaCod, "") & prs!cCtaObjSubCta
      prs.Close: Set prs = Nothing
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
GetCtaObjFiltroErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:GetCtaObjFiltro Method")
End Function

Public Function CtaInstancia(psCtaCod As String, Optional psTabla As String = "CtaCont") As Boolean
   On Error GoTo CtaInstancia
   CtaInstancia = True
   psSql = "SELECT cCtaContCod FROM " & psTabla & " WHERE cCtaContCod LIKE '" & psCtaCod & "%'"
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      Set prs = DBConec.CargaRecordSet(psSql)
      If prs.RecordCount > 1 Then
         CtaInstancia = False
      End If
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CtaInstancia:
   Call RaiseError(MyUnhandledError, "DCtaCont:CuentaInstancia Method")
End Function

Private Sub Class_Initialize()
   Set oConec = New ClsIni.ClasIni
   sConexion = oConec.CadenaConexion
   sCentralCom = oConec.BaseComunes
   Set oConec = Nothing
End Sub


Public Function CargaCtaContClase(Optional psCtaCod As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
On Error GoTo CargaCtaContClaseErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "SELECT cCtaContCod, cCtaCaracter, nCtaMoneda FROM " & sCentralCom & "CtaContClase " & IIf(psCtaCod = "", "", " WHERE '" & psCtaCod & "' LIKE cCtaContCod + '%' ")
      Set CargaCtaContClase = DBConec.CargaRecordSet(psSql, LockType)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaCtaContClaseErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaContClase Method")
End Function


