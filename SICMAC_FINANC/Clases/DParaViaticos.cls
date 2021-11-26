VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DParaViaticos"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A941CDA0203"
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Public sCentralCom As String

Public Function CargaParametros(Optional psCategCod As String = "", Optional psDestinoCod As String = "", Optional psTranspCod As String = "", Optional psObjetoCod As String = "", Optional LockType As LockTypeEnum = adLockReadOnly, Optional lDescrip As Boolean = False) As Recordset
Dim sCond As String
   On Error GoTo CargaParametrosErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      sCond = ""
      If psCategCod <> "" Then
         sCond = "WHERE v.cCategCod = '" & psCategCod & "' "
      End If
      If psDestinoCod <> "" Then
         sCond = sCond & IIf(sCond = "", "WHERE", "AND") & " v.cDestinoCod = '" & psDestinoCod & "' "
      End If
      If psTranspCod <> "" Then
         sCond = sCond & IIf(sCond = "", "WHERE", "AND") & " v.cTranspcod = '" & psTranspCod & "' "
      End If
      If psObjetoCod <> "" Then
         sCond = sCond & IIf(sCond = "", "WHERE", "AND") & " v.cObjetoCod = '" & psObjetoCod & "' "
      End If
      If lDescrip Then
         psSql = "SELECT v.cCategCod, v.cDestinoCod, v.cTranspCod, " _
               & "v.cObjetoCod, cObjetoDesc = (SELECT cObjetoDesc FROM " & sCentralCom & "Objeto o WHERE o.cObjetoCod = v.cObjetoCod), " _
               & "v.cViaticoAfectoA, cViaticoAfectoADesc = (SELECT cConsDescripcion FROM Constante c WHERE nConsCod = '" & gViaticosAfectoA & "' and c.nConsValor = v.cViaticoAfectoA), " _
               & "v.cViaticoAfectoTope, nViaticoImporte FROM Viaticos v " & sCond
      Else
         psSql = "SELECT cCategCod, cDestinoCod, cTranspCod, v.cObjetoCod, v.cViaticoAfectoA, cViaticoAfectoTope, nViaticoImporte FROM Viaticos v " & sCond
      End If
      Set CargaParametros = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaParametrosErr:
   Call RaiseError(MyUnhandledError, "DParaViaticos:CargaParametros Method")
End Function

Public Sub InsertaParametros(psCategCod As String, psDestinoCod As String, psTranspCod As String, psObjetoCod As String, psViaticoAfectoA As String, psViaticoAfectoTope As String, pnViaticoImporte As Currency, psMovNro As String)
   On Error GoTo InsertaParametrosErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT Viaticos (cCategCod, cDestinoCod, cTranspCod, cObjetoCod, cViaticoAfectoA, cViaticoAfectoTope, nViaticoImporte, cUltimaActualizacion) " _
            & "VALUES ('" & psCategCod & "','" & psDestinoCod & "','" & psTranspCod & "','" & psObjetoCod & "','" & psViaticoAfectoA & "','" & psViaticoAfectoTope & "'," & pnViaticoImporte & ",'" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaParametrosErr:
   Call RaiseError(MyUnhandledError, "DParaViaticos:InsertaParametros Method")
End Sub

Public Sub ActualizaParametros(psCategCod As String, psDestinoCod As String, psTranspCod As String, psObjetoCod As String, psViaticoAfectoA As String, psViaticoAfectoTope As String, pnViaticoImporte As Currency, psMovNro As String)
   On Error GoTo ActualizaParametrosErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "UPDATE Viaticos SET cViaticoAfectoA = '" & psViaticoAfectoA & "', cViaticoAfectoTope = '" & psViaticoAfectoTope & "',  nViaticoImporte = " & pnViaticoImporte & ", cUltimaActualizacion = '" & psMovNro & "' " _
            & "WHERE cCategCod = '" & psCategCod & "' and cDestinoCod = '" & psDestinoCod & "' and cTranspCod = '" & psTranspCod & "' and cObjetoCod = '" & psObjetoCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizaParametrosErr:
   Call RaiseError(MyUnhandledError, "DParaViaticos:ActualizaParametros Method")
End Sub

Public Sub EliminaParametros(psCategCod As String, psDestinoCod As String, psTranspCod As String, psObjetoCod As String, Optional psMovNro As String = "")
   On Error GoTo EliminaParametrosErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE Viaticos " _
            & "WHERE cCategCod = '" & psCategCod & "' and cDestinoCod = '" & psDestinoCod & "' and cTranspCod = '" & psTranspCod & "' and cObjetoCod = '" & psObjetoCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaParametrosErr:
   Call RaiseError(MyUnhandledError, "DParaViaticos:EliminaParametros Method")
End Sub

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oConec As New ClasIni
   sCentralCom = oConec.BaseComunes
   Set oConec = Nothing
End Sub
