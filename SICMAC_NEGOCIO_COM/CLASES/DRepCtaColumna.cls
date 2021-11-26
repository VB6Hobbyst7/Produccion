VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRepCtaColumna"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Clase de Datos para Reporte de Cuentas por Columnas"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Dim sConexion As String
Dim sCentralCom As String
Dim DBConec   As DConecta

Dim psSql As String

Public Function CargaRepColumna(Optional psOpeCod As String = "", Optional pnNroCol As Integer = -1, Optional plLike As Boolean = False, Optional LockType As LockTypeEnum = adLockReadOnly, Optional plOpeDesc As Boolean = False) As Recordset
Dim sCond As String
   On Error GoTo CargaRepCtaErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      sCond = IIf(psOpeCod = "", "", " WHERE cOpeCod " & IIf(plLike, " LIKE '", "= '") & psOpeCod & IIf(plLike, "%'", "'"))
      If pnNroCol >= 0 Then
         sCond = sCond & IIf(sCond = "", " WHERE ", " AND ")
         sCond = sCond & " nNroCol = " & pnNroCol
      End If
      If plOpeDesc Then
         psSql = "SELECT rc.cOpeCod, rc.nNroCol, rc.cDescCol, rc.cTotal, ISNULL(o.cOpeDesc,'') as cOpeDesc FROM RepColumna rc LEFT JOIN " & sCentralCom & "OpeTpo o ON o.cOpeCod = rc.cOpeCod " & Replace(sCond, "cOpecod", "rc.cOpecod")
      Else
         psSql = "SELECT cOpeCod, nNroCol, cDescCol, cTotal FROM RepColumna" & sCond
      End If
      Set CargaRepColumna = DBConec.CargaRecordSet(psSql, LockType)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaRepCtaErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:CargaRepCta Method")
End Function

Public Function CargaRepColumnaCta(psCodOpe As String, pnNroCol As Integer, Optional LockType As LockTypeEnum = adLockReadOnly, Optional plCtaDesc As Boolean = False) As Recordset
   On Error GoTo CargaRepColumnaCtaErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      If plCtaDesc Then
         psSql = "SELECT rc.cOpeCod, rc.nNroCol, rc.cCtaContCod, c.cCtaContDesc FROM RepColumnaCta rc JOIN " & sCentralCom & "CtaCont c ON c.cCtaContCod = rc.cCtaContCod " _
               & "WHERE rc.cOpeCod = '" & psCodOpe & "' and rc.nNroCol = " & pnNroCol
      Else
         psSql = "SELECT cOpeCod, nNroCol, cCtaContCod FROM RepColumnaCta WHERE cOpeCod = '" & psCodOpe & "' and nNroCol = " & pnNroCol
      End If
      Set CargaRepColumnaCta = DBConec.CargaRecordSet(psSql, LockType)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaRepColumnaCtaErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:CargaRepColumnaCta Method")
End Function

Public Function MaxNroColumna(psCodOpe As String) As Integer
Dim pRs As ADODB.Recordset
   On Error GoTo MaxNroColumnaErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "SELECT MAX(nNroCol) nNroCol FROM RepColumna WHERE cOpeCod = '" & psCodOpe & "' "
      Set pRs = DBConec.CargaRecordSet(psSql)
      If Not pRs.EOF Then
         MaxNroColumna = pRs!nNroCol
      Else
         MaxNroColumna = 0
      End If
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
MaxNroColumnaErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:MaxNroColumna Method")
End Function

Public Function CargaRepImpresion() As Recordset
   On Error GoTo CargaRepImpresionErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "SELECT RC.COPECOD, ISNULL(OT.COPEDESC,'') cOpeDesc, RC.nNroCol, RC.CDESCCOL, RCTA.CCTACONTCOD, CC.CCTACONTDESC " _
            & "FROM RepColumna RC LEFT JOIN RepColumnaCta RCTA ON RCTA.COPECOD=RC.COPECOD AND RCTA.nNroCol=RC.nNroCol " _
            & "     LEFT JOIN " & sCentralCom & "CtaCont CC ON RCTA.CCTACONTCOD=CC.CCTACONTCOD " _
            & "     LEFT JOIN " & sCentralCom & "OpeTpo OT ON OT.COPECOD=RC.COPECOD  ORDER BY RC.COPECOD "
      Set CargaRepImpresion = DBConec.CargaRecordSet(psSql, adLockReadOnly)
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Function
CargaRepImpresionErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:CargaRepImpresion Method")
End Function


Public Sub InsertaRepColumna(psOpeCod As String, pnNroCol As Integer, psDescCol As String, psTotal As String, psMovNro As String)
   On Error GoTo InsertaRepColumnaErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "INSERT RepColumna (cOpeCod, nNroCol, cDescCol, cTotal, cUltimaActualizacion) " _
            & "VALUES ('" & psOpeCod & "', " & pnNroCol & ", '" & psDescCol & "','" & psTotal & "','" & psMovNro & "')"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaRepColumnaErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:InsertaRepColumna Method")
End Sub

Public Sub InsertaRepColumnaCta(psOpeCod As String, pnNroCol As Integer, psCtaContCod As String, psMovNro As String)
   On Error GoTo InsertaRepColumnaCtaErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "INSERT RepColumnaCta (cOpeCod, nNroCol, cCtaContCod, cUltimaActualizacion) VALUES ('" & psOpeCod & "', " & pnNroCol & ", '" & psCtaContCod & "','" & psMovNro & "')"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
InsertaRepColumnaCtaErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:InsertaRepColumnaCta Method")
End Sub

Public Sub EliminaRepColumna(psOpeCod As String, pnNroCol As Integer)
Dim lsError As String
   On Error GoTo EliminaRepColumnaErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      DBConec.BeginTrans
      psSql = "DELETE RepColumnaCta WHERE cOpeCod = '" & psOpeCod & "' and nNroCol = " & pnNroCol & ""
      DBConec.Ejecutar psSql
      
      psSql = "DELETE RepColumna WHERE cOpeCod = '" & psOpeCod & "' and nNroCol = " & pnNroCol & ""
      DBConec.Ejecutar psSql
      
      DBConec.CommitTrans
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
EliminaRepColumnaErr:
   lsError = Err.Description
   DBConec.RollbackTrans
   Err.Raise Err.Number + MyUnhandledError, "DRepCtaColumna:EliminaRepColumna Method", lsError
End Sub

Public Sub EliminaRepColumnaCta(psOpeCod As String, pnNroCol As Integer, psctacod As String)
   On Error GoTo EliminaRepColumnaCtaErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "DELETE RepColumnaCta  WHERE cOpeCod = '" & psOpeCod & "' and nNroCol = " & pnNroCol & " and cCtaContCod = '" & psctacod & "'"
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
EliminaRepColumnaCtaErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:EliminaRepColumnaCta Method")
End Sub

Public Sub ActualizaRepColumna(psOpeCod As String, pnNroCol As Integer, psDescCol As String, pnTotal As Currency, psMovNro As String)
   On Error GoTo ActualizaRepColumnaErr
   Set DBConec = New DConecta
   If DBConec.AbreConexion(sConexion) Then
      psSql = "UPDATE RepColumna Set cDescCol = '" & psDescCol & "', cTotal = " & pnTotal & ", cUltimaActualizacion = '" & psMovNro & "'  WHERE cOpeCod = '" & psOpeCod & "' and nNroCol = " & pnNroCol
      DBConec.Ejecutar psSql
      DBConec.CierraConexion
   End If
   Set DBConec = Nothing
   Exit Sub
ActualizaRepColumnaErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:ActualizaRepColumna Method")
End Sub

Private Sub Class_Initialize()
   Dim oIni As New COMConecta.DCOMClasIni
   sConexion = oIni.CadenaConexion
   sCentralCom = oIni.BaseComunes
   Set oIni = Nothing
End Sub
