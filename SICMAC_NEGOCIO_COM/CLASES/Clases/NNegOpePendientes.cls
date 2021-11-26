VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NNegOpePendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oCon As DConecta
Dim lsSql As String
Dim prs   As ADODB.Recordset

Public Function CargaOpeVentanillaPendiente(psOpeCod As String, psDH As String, psCodAge As String, pnImporte As Currency, psMoneda As String, Optional psCodPers As String = "") As Recordset
Dim lsOpeFiltro As String
If InStr(psOpeCod, ",") = 0 Then
    lsOpeFiltro = " td.cCodOpe = '" & psOpeCod & "' "
Else
    lsOpeFiltro = " td.cCodOpe IN ('" & psOpeCod & "') "
End If

On Error GoTo CargaOpeVentanillaPendienteErr

   lsSql = "SELECT td.cCodOpe, td.dFecTran, td.nNumTran, oo.cCodPers, p.cNomPers, td.nMonTran, oo.cGlosa, oc.cCtaContCod " _
         & "FROM   TransAho td " _
         & "       JOIN OtrasOpe oo ON oo.nNumTran = td.nNumTran " _
         & "       JOIN DBPersona..Persona p ON p.cCodPers = oo.cCodPers " _
         & "       JOIN TransRef ref ON Convert(money,ref.cNroTranRef) = td.nNumTran and ref.cCodAgeRef = '" & psCodAge & "' " _
         & "       JOIN (SELECT DISTINCT cCodOpe, cCodCnt cCtaContCod FROM DBComunes..OpeCuentaN WHERE cDebeHab = '" & psDH & "') oc ON oc.cCodOpe = td.cCodOpe " _
         & "WHERE  " & lsOpeFiltro & " and td.cFlag is NULL and ref.cNroTran IS NULL and " _
         & "       SubString(td.cCodCta,6,1) = '" & psMoneda & "' " _
         & IIf(psCodPers = "", "", " and oo.cCodPers = '" & psCodPers & "' ")
   Set oCon = New DConecta
   If oCon.AbreConexionRemota(Right(psCodAge, 2), , False, "01") Then
      Set CargaOpeVentanillaPendiente = oCon.CargaRecordSet(lsSql)
   End If
   oCon.CierraConexion
Exit Function
CargaOpeVentanillaPendienteErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ActualizaNegocioReferencia(psNroTran As String, psAgeCod As String, psMovNro As String)
On Error GoTo ActualizaNegocioReferenciaErr
lsSql = "UPDATE TransRef SET cNroTran = '" & psMovNro & "', cCodAge = '" & psAgeCod & "' WHERE cNroTranRef = '" & psNroTran & "' and cCodAgeRef = '" & psAgeCod & "'"
   Set oCon = New DConecta
   If oCon.AbreConexionRemota(Right(psAgeCod, 2), , False, "01") Then
      oCon.Ejecutar lsSql
   End If
   oCon.CierraConexion
Exit Function
ActualizaNegocioReferenciaErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

End Sub
