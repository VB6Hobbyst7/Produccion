VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAjusteCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim DBConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sConexion As String
Dim sCentralCom As String

Public Function CargaAjusteInflacion(Optional psCtaCod As String = "", Optional psAjusteCod As String = "", Optional psFiltro As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
Dim sCond As String
   On Error GoTo CargaAjusteInflacionErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      If psCtaCod <> "" Then
         sCond = "WHERE cCtaContCod ='" & psCtaCod & "' "
      End If
      If psAjusteCod <> "" Then
         sCond = sCond & IIf(sCond = "", "WHERE", "AND") & " cAjusteCod = '" & psAjusteCod & "' "
      End If
      If psFiltro <> "" Then
         sCond = sCond & IIf(sCond = "", "WHERE", "AND") & " " & psFiltro
      End If
      psSql = "SELECT cCtaContCod, cAjusteCod, cAjusteDescrip, Convert(varchar(10),dAjusteFecha,103) dAjusteFecha, nAjusteValor1, nAjusteValor2, nAjusteValor3 From AjusteInflacion " & sCond & " Order By cCtaContCod, cAjusteCod "
      Set CargaAjusteInflacion = DBConec.CargaRecordSet(psSql, LockType)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaAjusteInflacionErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:CargaAjusteInflacion Method")
End Function

Public Sub InsertaAjuste(psCtaCod As String, psAjusteCod As String, psDescrip As String, psFecha As String, pnValor1 As Currency, pnValor2 As Currency, pnValor3 As Currency)
   On Error GoTo InsertaAjusteErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "INSERT AjusteInflacion (cCtaContCod, cAjusteCod, cAjusteDescrip, dAjusteFecha, nAjusteValor1, nAjusteValor2, nAjusteValor3) " _
            & "VALUES ('" & psCtaCod & "','" & psAjusteCod & "','" & psDescrip & "', '" & psFecha & "'," & pnValor1 & "," & pnValor2 & "," & pnValor3 & ")"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaAjusteErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:InsertaAjuste Method")
End Sub

Public Sub ActualizaAjuste(psCtaCod As String, psAjusteCod As String, psDescrip As String, psFecha As String, pnValor1 As Currency, pnValor2 As Currency, pnValor3 As Currency)
   On Error GoTo ActualizaAjusteErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "UPDATE AjusteInflacion SET cAjusteDescrip = '" & psDescrip & "', dAjusteFecha = '" & psFecha & "', nAjusteValor1  = " & pnValor1 & ", nAjusteValor2  = " & pnValor2 & ", nAjusteValor3  = " & pnValor3 & " WHERE cCtaContCod = '" & psCtaCod & "' and cAjusteCod = '" & psAjusteCod & "'"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
ActualizaAjusteErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:ActualizaAjuste Method")
End Sub

Public Sub EliminaAjuste(psCtaCod As String, psAjusteCod As String)
   On Error GoTo EliminaAjusteErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "DELETE AjusteInflacion WHERE cCtaContCod = '" & psCtaCod & "' and cAjusteCod = '" & psAjusteCod & "'"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
EliminaAjusteErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:EliminaAjuste Method")
End Sub

Public Function CargaIPM(Optional psFecha As String = "", Optional psFiltro As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
Dim sCond As String
   On Error GoTo CargaImpuestoErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      If psFecha <> "" Then
         sCond = "WHERE dFecha ='" & psFecha & "'"
      End If
      If psFiltro <> "" Then
         sCond = sCond & IIf(psFecha = "", "WHERE", "AND") & " " & psFiltro
      End If
      psSql = "SELECT Convert(varchar(10),dFecha,103) dFecha, Convert(decimal(16,6),nValor) nValor From IPM " & sCond '& " Order By dFecha "
      Set CargaIPM = DBConec.CargaRecordSet(psSql, LockType)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaImpuestoErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:CargaImpuesto Method")
End Function

Public Sub InsertaIPM(psFecha As String, pnValor As Double)
   On Error GoTo InsertaIPMErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "INSERT IPM (dFecha, nValor) VALUES ('" & psFecha & "'," & pnValor & ")"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaIPMErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:InsertaIPM Method")
End Sub

Public Sub ActualizaIPM(psFecha As String, pnValor As Double)
   On Error GoTo ActualizaIPMErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "UPDATE IPM SET nValor  = " & pnValor & " WHERE dFecha = '" & psFecha & "'"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
ActualizaIPMErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:ActualizaIPM Method")
End Sub

Public Sub EliminaIPM(psFecha As String)
   On Error GoTo EliminaIPMErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "DELETE IPM WHERE dFecha = '" & psFecha & "'"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
EliminaIPMErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:EliminaIPM Method")
End Sub
Public Function CargaAjusteHistorico() As Recordset
Dim sCond As String
   On Error GoTo CargaAjusteHistoricoErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then

      psSql = "SELECT da.*, c.cCtaContDesc, cls.cCtaCaracter as cTipoDH " _
           & "FROM AjusteInflacion da JOIN " & sCentralCom & "CtaCont c ON c.cCtaContCod = da.cCtaContCod " _
           & "     JOIN " & sCentralCom & "CtaContClase CLS on da.cCtaContCod LIKE cls.cCtaContCod + '%' "
      Set CargaAjusteHistorico = DBConec.CargaRecordSet(psSql)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaAjusteHistoricoErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:CargaAjusteHistorico Method")
End Function

Public Function VerAjusteCierreAnual(psFechaCie As String, psCond As String) As Recordset
On Error GoTo VerAjusteCierreAnualErr
   Set DBConec = New DConecta
   DBConec.AbreConexion sConexion
   psSql = "SELECT c.cCtaContCod, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte, cls.cCtaCaracter cTipo " _
        & "FROM " & sCentralCom & "CtaCont c LEFT JOIN " _
        & "     (SELECT cs.cCtaContCod, cs.nCtaSaldoImporte FROM CtaSaldo cs WHERE dCtaSaldoFecha = " _
        & "      ( SELECT MAX(a.dCtaSaldoFecha) FROM CtaSaldo a " _
        & "        WHERE  a.cCtaContCod = cs.cCtaContCod and dCtasaldoFecha <= '" & psFechaCie & "') " _
        & "     ) cs on cs.cCtaContCod = c.cCtaContCod, " & sCentralCom & "CtaContClase cls " _
        & "WHERE " & psCond _
        & "     and c.cCtaContCod LIKE cls.cCtaContCod + '%'  "
   Set VerAjusteCierreAnual = DBConec.CargaRecordSet(psSql)
   Set DBConec = Nothing
Exit Function
VerAjusteCierreAnualErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:VerAjusteCierreAnual Method")
End Function

Private Sub Class_Initialize()
Dim oIni As New ClasIni
   sConexion = oIni.CadenaConexion
   sCentralCom = oIni.BaseComunes
Set oIni = Nothing
End Sub



