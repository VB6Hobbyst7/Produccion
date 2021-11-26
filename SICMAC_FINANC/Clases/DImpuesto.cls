VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DImpuesto"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A80C6000203"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sCentralCom As String

Public Function CargaCtaImpuesto() As Recordset
   On Error GoTo CargaImpuestoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "SELECT Impuesto.cCtaContCod, CtaCont.cCtaContDesc, 1 as nObjetoNiv " & _
              "FROM  " & sCentralCom & "CtaCont CtaCont,  " & sCentralCom & "Impuesto Impuesto " & _
              "WHERE CtaCont.cCtaContCod = Impuesto.cCtaContCod  "
      Set CargaCtaImpuesto = dbConec.CargaRecordSet(psSql, adLockReadOnly)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaImpuestoErr:
   Call RaiseError(MyUnhandledError, "DImpuesto:CargaImpuesto Method")
End Function

Public Function CargaImpuestoFechaValor(Optional psCtaContCod As String = "", Optional pdFecha As Date = 0) As Currency
   On Error GoTo CargaImpuestoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      
      psSql = "SELECT i.cCtaContCod, nImpTasa FROM ImpuestoFecha i WHERE i.cCtaContCod = '" & psCtaContCod & "'" _
          & IIf(pdFecha = 0, "", " and dFechaIniVig = (SELECT Max(dFechaIniVig) FROM ImpuestoFecha f WHERE f.cCtaContCod = i.cCtaContCod and f.dFechaIniVig <= '" & Format(pdFecha, gsFormatoFecha) & "' ) ")
      Set prs = dbConec.CargaRecordSet(psSql)
      If Not prs.EOF Then
         CargaImpuestoFechaValor = prs!nImpTasa
      End If
      RSClose prs
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaImpuestoErr:
   Call RaiseError(MyUnhandledError, "DImpuesto:CargaImpuestoFecha Method")
End Function


Public Function CargaImpuesto(Optional psCtaContCod As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaImpuestoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "SELECT i.cCtaContCod, c.cCtaContDesc, cImpAbrev, nImpTasa, cImpDestino, nCalculo FROM " & sCentralCom & "Impuesto i JOIN " & sCentralCom & "CtaCont c ON c.cCtaContCod = i.cCtaContCod " & IIf(psCtaContCod = "", "", " WHERE i.cCtaContCod ='" & psCtaContCod & "'")
      Set CargaImpuesto = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaImpuestoErr:
   Call RaiseError(MyUnhandledError, "DImpuesto:CargaImpuesto Method")
End Function

Public Sub InsertaImpuesto(psCtaCod As String, psImpAbrev As String, pnImpTasa As Currency, psImpDestino As String, pnImpCalculo As Currency, psMovNro As String)
   On Error GoTo InsertaImpuestoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "Impuesto (cCtaContCod, cImpAbrev, nImpTasa, cImpDestino, nCalculo, cUltimaActualizacion) " _
            & "VALUES ('" & psCtaCod & "','" & psImpAbrev & "'," & pnImpTasa & ",'" & psImpDestino & "'," & pnImpCalculo & ",'" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaImpuestoErr:
   Call RaiseError(MyUnhandledError, "DImpuesto:InsertaImpuesto Method")
End Sub

Public Sub ActualizaImpuesto(psCtaCod As String, psImpAbrev As String, pnImpTasa As Currency, psImpDestino As String, pnImpCalculo As Currency, psMovNro As String)
   On Error GoTo ActualizaImpuestoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "UPDATE " & sCentralCom & "Impuesto SET cImpAbrev = '" & psImpAbrev & "', nImpTasa = " & pnImpTasa & ", cImpDestino = '" & psImpDestino & "', nCalculo = " & pnImpCalculo & ", cUltimaActualizacion = '" & psMovNro & "' WHERE cCtaContCod = '" & psCtaCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizaImpuestoErr:
   Call RaiseError(MyUnhandledError, "DImpuesto:ActualizaImpuesto Method")
End Sub

Public Sub EliminaImpuesto(psCtaCod As String)
   On Error GoTo EliminaImpuestoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "Impuesto WHERE cCtaContCod = '" & psCtaCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaImpuestoErr:
   Call RaiseError(MyUnhandledError, "DImpuesto:EliminaImpuesto Method")
End Sub


Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As New ClasIni
   sCentralCom = oIni.BaseComunes
Set oIni = Nothing
End Sub

