VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMObjeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim dbConec As COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling
Dim psSql As String
Dim prs   As ADODB.Recordset

Public sConexion As String
Public sCentralCom As String
Public Function CargaObjeto(Optional psObjCod As String = "", Optional plLike As Boolean = False, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
   On Error GoTo CargaObjetoErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
      psSql = "SELECT cObjetoCod, cObjetoDesc, nObjetoNiv FROM " & sCentralCom & "Objeto" & IIf(psObjCod = "", "", " WHERE cObjetoCod " & IIf(plLike, " LIKE '", "= '") & psObjCod & IIf(plLike, "%'", "'"))
      Set CargaObjeto = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaObjetoErr:
  oError.RaiseError oError.MyUnhandledError, "DConstantes:GetConstante Method"
End Function

Public Sub InsertaObjeto(psObjetoCod As String, psObjetoDesc As String, pnObjetoNiv As Integer, psMovNro As String)
   On Error GoTo InsertaObjetoErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
      psSql = "INSERT " & sCentralCom & "Objeto (cObjetoCod, cObjetoDesc, nObjetoNiv, cUltimaActualizacion) " _
            & "VALUES ('" & psObjetoCod & "','" & psObjetoDesc & "'," & pnObjetoNiv & ",'" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaObjetoErr:
  oError.RaiseError oError.MyUnhandledError, "DObjeto:InsertaObjeto Method"
End Sub

Public Sub ActualizaObjeto(psObjetoCod As String, psObjetoDesc As String, pnObjetoNiv As Integer, psMovNro As String)
   On Error GoTo ActualizaObjetoErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
      psSql = "UPDATE  " & sCentralCom & "Objeto SET cObjetoDesc = '" & psObjetoDesc & "', nObjetoNiv = " & pnObjetoNiv & ", cUltimaActualizacion = '" & psMovNro & "' WHERE cObjetoCod = '" & psObjetoCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizaObjetoErr:
    oError.RaiseError oError.MyUnhandledError, "DObjeto:InsertaObjeto Method"
End Sub

Public Sub EliminaObjeto(psObjetoCod As String, Optional plLike As Boolean = False, Optional plIn As Boolean = False)
   On Error GoTo EliminaObjetoErr
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
      psSql = "DELETE " & sCentralCom & "Objeto WHERE cObjetoCod " & IIf(plIn, " IN (" & psObjetoCod & ")", IIf(plLike, "LIKE", "=") & " '" & psObjetoCod & "' ")
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaObjetoErr:
   oError.RaiseError oError.MyUnhandledError, "DObjeto:InsertaObjeto Method"
End Sub

Private Sub Class_Initialize()
   Dim oconec As New COMConecta.DCOMClasIni
   sConexion = oconec.CadenaConexion
   sCentralCom = oconec.BaseComunes
   Set oconec = Nothing
End Sub

Public Function ObjInstancia(psObjCod As String) As Boolean
   On Error GoTo ObjInstancia
   ObjInstancia = True
   psSql = "SELECT cObjetoCod FROM " & sCentralCom & "Objeto WHERE cObjetoCod LIKE '" & psObjCod & "%'"
   Set dbConec = New COMConecta.DCOMConecta
   If dbConec.AbreConexion(sConexion) Then
      Set prs = dbConec.CargaRecordSet(psSql)
      If prs.RecordCount > 1 Then
         ObjInstancia = False
      End If
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
ObjInstancia:
   oError.RaiseError oError.MyUnhandledError, "DObjeto:InsertaObjeto Method"
End Function



