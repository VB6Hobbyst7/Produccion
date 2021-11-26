VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DObjeto"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A8038CC01E4"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Dim DBConec As DConecta
Dim psSql As String
Dim pRs   As ADODB.Recordset

Public sConexion As String
Public sCentralCom As String
Public Function CargaObjeto(Optional psObjCod As String = "", Optional plLike As Boolean = False, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
   On Error GoTo CargaObjetoErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "SELECT cObjetoCod, cObjetoDesc, nObjetoNiv FROM " & sCentralCom & "Objeto" & IIf(psObjCod = "", "", " WHERE cObjetoCod " & IIf(plLike, " LIKE '", "= '") & psObjCod & IIf(plLike, "%'", "'"))
      Set CargaObjeto = DBConec.CargaRecordSet(psSql, LockType)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaObjetoErr:
   Call RaiseError(MyUnhandledError, "DObjeto:CargaObjeto Method")
End Function

Public Sub InsertaObjeto(psObjetoCod As String, psObjetoDesc As String, pnObjetoNiv As Integer, psMovNro As String)
   On Error GoTo InsertaObjetoErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "INSERT " & sCentralCom & "Objeto (cObjetoCod, cObjetoDesc, nObjetoNiv, cUltimaActualizacion) " _
            & "VALUES ('" & psObjetoCod & "','" & psObjetoDesc & "'," & pnObjetoNiv & ",'" & psMovNro & "')"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaObjetoErr:
   Call RaiseError(MyUnhandledError, "DObjeto:InsertaObjeto Method")
End Sub

Public Sub ActualizaObjeto(psObjetoCod As String, psObjetoDesc As String, pnObjetoNiv As Integer, psMovNro As String)
   On Error GoTo ActualizaObjetoErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "UPDATE  " & sCentralCom & "Objeto SET cObjetoDesc = '" & psObjetoDesc & "', nObjetoNiv = " & pnObjetoNiv & ", cUltimaActualizacion = '" & psMovNro & "' WHERE cObjetoCod = '" & psObjetoCod & "'"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
ActualizaObjetoErr:
   Call RaiseError(MyUnhandledError, "DObjeto:ActualizaObjeto Method")
End Sub

Public Sub EliminaObjeto(psObjetoCod As String, Optional plLike As Boolean = False, Optional plIn As Boolean = False)
   On Error GoTo EliminaObjetoErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "DELETE " & sCentralCom & "Objeto WHERE cObjetoCod " & IIf(plIn, " IN (" & psObjetoCod & ")", IIf(plLike, "LIKE", "=") & " '" & psObjetoCod & "' ")
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
EliminaObjetoErr:
   Call RaiseError(MyUnhandledError, "DObjeto:EliminaObjeto Method")
End Sub

Private Sub Class_Initialize()
Dim oConec As New COMConecta.DCOMClasIni
   sConexion = oConec.CadenaConexion
   sCentralCom = oConec.BaseComunes
   Set oConec = Nothing
End Sub

Public Function ObjInstancia(psObjCod As String) As Boolean
   On Error GoTo ObjInstancia
   ObjInstancia = True
   psSql = "SELECT cObjetoCod FROM " & sCentralCom & "Objeto WHERE cObjetoCod LIKE '" & psObjCod & "%'"
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      Set pRs = DBConec.CargaRecordSet(psSql)
      If pRs.RecordCount > 1 Then
         ObjInstancia = False
      End If
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
ObjInstancia:
   Call RaiseError(MyUnhandledError, "DObjeto:ObjInstancia Method")
End Function

