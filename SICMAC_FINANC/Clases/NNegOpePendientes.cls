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
Dim pnBitCentral As Boolean

Public Function CargaOpeVentanillaPendCtaContCentral(psCtaCod As String, psCodAge As String, psMoneda As String, Optional psCodPers As String = "") As Recordset
Dim lsOpeFiltro As String
On Error GoTo CargaOpeVentanillaPendCtaContErr
   lsSql = "Select M.cOpeCod, Convert(DateTime,LEFT(M.cMovNro,4)+'/'+SubString(M.cMovNro,5,2)+'/'+SubString(M.cMovNro,7,2)) dFecTran, M.nMovNro, T.cPersCodIF, RTRIM(P.cPersNombre) + ' - ' + T.cCtaBco cNomPers, MC.nMonto nMonTran, M.cMovDesc cGlosa " _
         & "From MovCap MC JOIN Mov M JOIN MovTransferBco T ON M.nMovNro = T.nMovNro ON MC.nMovNro = M.nMovNro JOIN Persona P ON P.cPersCod = T.cPersCodIF " _
         & "     JOIN (select Distinct cOpeCod FROM OpeCtaNeg oc WHERE '" & psCtaCod & "' LIKE REPLACE(oc.cCtaContCod,'M','" & psMoneda & "') + '%') oc ON oc.cOpeCod = M.cOpeCod  " _
         & "WHere M.nMovNro NOT IN (Select MR.nMovNroRef From Mov M1 JOIN MovRef MR ON M1.nMovNro = MR.nMovNro " _
         & " And  M1.nMovFlag = " & gMovFlagVigente & " And MR.nMovNroRef = M.nMovNro) And M.nMovFlag = " & gMovFlagVigente & " " _
         & " and T.nMoneda = " & psMoneda & " and SubString(MC.cCtaCod,4,2) = '" & Right(psCodAge, 2) & "' "
   Set oCon = New DConecta
   oCon.AbreConexion
      Set CargaOpeVentanillaPendCtaContCentral = oCon.CargaRecordSet(lsSql)
   oCon.CierraConexion
Exit Function
CargaOpeVentanillaPendCtaContErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CargaOpeVentanillaPendCtaCont(psCtaCod As String, psDH As String, psCodAge As String, pnImporte As Currency, psMoneda As String, Optional psCodPers As String = "") As Recordset
Dim lsOpeFiltro As String
If InStr(psCtaCod, ",") = 0 Then
    lsOpeFiltro = " cCodCnt LIKE '" & Left(psCtaCod, 2) & "_" & Mid(psCtaCod, 4, 22) & "%' "
Else
    lsOpeFiltro = " cCodCnt IN ('" & psCtaCod & "') "
End If
On Error GoTo CargaOpeVentanillaPendCtaContErr

   lsSql = "SELECT td.cCodOpe, td.dFecTran, td.nNumTran, ISNULL(oo.cCodPers,'') cCodPers, ISNULL(p.cNomPers,'') cNomPers, abs(td.nMonTran) nMonTran, ISNULL(oo.cGlosa,'') cGlosa, oc.cCtaContCod " _
         & "FROM   TransAho td " _
         & "       LEFT JOIN OtrasOpe oo ON oo.nNumTran = td.nNumTran " _
         & "       LEFT JOIN DBPersona..Persona p ON p.cCodPers = oo.cCodPers " _
         & "       JOIN TransRef ref ON Convert(money,ref.cNroTranRef) = td.nNumTran and ref.cCodAgeRef = '" & psCodAge & "' " _
         & "       JOIN (SELECT DISTINCT cCodOpe, cCodCnt cCtaContCod FROM DBComunes..OpeCuentaN WHERE " & IIf(psDH = "", "", "cDebeHab = '" & psDH & "' and ") & lsOpeFiltro & " ) oc ON oc.cCodOpe = td.cCodOpe " _
         & "WHERE  td.cFlag is NULL and ref.cNroTran IS NULL and " _
         & "       SubString(td.cCodCta,6,1) = '" & psMoneda & "' " _
         & IIf(psCodPers = "", "", " and oo.cCodPers = '" & psCodPers & "' ") _
         & "ORDER BY td.dFecTran "
   Set oCon = New DConecta
   If oCon.AbreConexion Then 'Remota(Right(psCodAge, 2), , False, "01")
      Set CargaOpeVentanillaPendCtaCont = oCon.CargaRecordSet(lsSql)
   End If
   oCon.CierraConexion
Exit Function
CargaOpeVentanillaPendCtaContErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CargaOpeRegulaVentanillaPendCentral(pnMovNro As Long, psCodAge As String, psMoneda As String) As Recordset
Dim lsOpeFiltro As String
On Error GoTo CargaOpeRegulaVentanillaPendCentralErr
   lsSql = "Select M.cOpeCod, Convert(DateTime,LEFT(M.cMovNro,4)+'/'+SubString(M.cMovNro,5,2)+'/'+SubString(M.cMovNro,7,2)) dFecTran, M.nMovNro, T.cPersCodIF, RTRIM(P.cPersNombre) + ' - ' + T.cCtaBco cNomPers, MC.nMonto nMonTran, M.cMovDesc cGlosa " _
         & "From MovCap MC JOIN Mov M JOIN MovTransferBco T ON M.nMovNro = T.nMovNro ON MC.nMovNro = M.nMovNro JOIN Persona P ON P.cPersCod = T.cPersCodIF " _
         & "     JOIN MovRef mr ON mr.nMovNro = m.nMovNro " _
         & "Where Mr.nMovNroRef = " & pnMovNro _
         & "  And M.nMovFlag = " & gMovFlagVigente & " " _
         & "  And T.nMoneda = " & psMoneda & " and SubString(MC.cCtaCod,4,2) = '" & Right(psCodAge, 2) & "' "
   Set oCon = New DConecta
   oCon.AbreConexion
      Set CargaOpeRegulaVentanillaPendCentral = oCon.CargaRecordSet(lsSql)
   oCon.CierraConexion
Exit Function
CargaOpeRegulaVentanillaPendCentralErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CargaOpeRegulaVentanillaPend(psMovNro As String, psCodAge As String, psCtaCod As String) As Recordset
Dim lsOpeFiltro As String
If InStr(psCtaCod, ",") = 0 Then
    lsOpeFiltro = " cCodCnt LIKE '" & Left(psCtaCod, 2) & "_" & Mid(psCtaCod, 4, 22) & "%' "
Else
    lsOpeFiltro = " cCodCnt IN ('" & psCtaCod & "') "
End If
On Error GoTo CargaOpeRegulaVentanillaPendErr

   lsSql = "SELECT td.cCodOpe, td.dFecTran, td.nNumTran, ISNULL(oo.cCodPers,'') cCodPers, ISNULL(p.cNomPers,'') cNomPers, abs(td.nMonTran) nMonTran, ISNULL(oo.cGlosa,'') cGlosa, oc.cCtaContCod " _
         & "FROM   TransAho td " _
         & "       LEFT JOIN OtrasOpe oo ON oo.nNumTran = td.nNumTran " _
         & "       LEFT JOIN DBPersona..Persona p ON p.cCodPers = oo.cCodPers " _
         & "       JOIN TransRef ref ON Convert(money,ref.cNroTranRef) = td.nNumTran and ref.cCodAgeRef = '" & psCodAge & "' " _
         & "       JOIN (SELECT DISTINCT cCodOpe, cCodCnt cCtaContCod FROM DBComunes..OpeCuentaN WHERE " & lsOpeFiltro & " ) oc ON oc.cCodOpe = td.cCodOpe " _
         & "WHERE  td.cFlag is NULL and ref.cNroTran = '" & psMovNro & "' " _
         & "ORDER BY td.dFecTran "
   Set oCon = New DConecta
   If oCon.AbreConexion Then 'Remota(Right(psCodAge, 2), , False, "01")
      Set CargaOpeRegulaVentanillaPend = oCon.CargaRecordSet(lsSql)
   End If
   oCon.CierraConexion
Exit Function
CargaOpeRegulaVentanillaPendErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CargaIngresoVentanillaPendiente(psOpeCod As String, psCodAge As String, psMoneda As String, Optional psCodPers As String = "", Optional pbBanco As Boolean = False) As Recordset
Dim sql As String
Dim sCadena As String
Dim lsOpeFiltro As String
Dim oCon As New DConecta
If pnBitCentral Then
    If InStr(psOpeCod, ",") = 0 Then
        lsOpeFiltro = " and m.cOpeCod = '" & psOpeCod & "' "
    Else
        lsOpeFiltro = " and m.cOpeCod IN ('" & psOpeCod & "') "
    End If
Else
    If InStr(psOpeCod, ",") = 0 Then
        lsOpeFiltro = " and td.cCodOpe = '" & psOpeCod & "' "
    Else
        lsOpeFiltro = " and td.cCodOpe IN ('" & psOpeCod & "') "
    End If
End If
    If pnBitCentral Then
       If pbBanco = True Then
          sql = "Select M.cOpeCod, Convert(DateTime,LEFT(M.cMovNro,4)+'/'+SubString(M.cMovNro,5,2)+'/'+SubString(M.cMovNro,7,2)) dFecTran, M.nMovNro, T.cPersCodIF , RTRIM(P.cPersNombre) + ' - ' + T.cCtaBco cNomPers, MC.nMonto nMonTran, M.cMovDesc cGlosa " _
              & "From MovCap MC JOIN Mov M JOIN MovTransferBco T ON M.nMovNro = T.nMovNro ON MC.nMovNro = M.nMovNro JOIN Persona P ON P.cPersCod = T.cPersCodIF " _
              & "WHere M.nMovNro NOT IN (Select MR.nMovNroRef From Mov M1 JOIN MovRef MR ON M1.nMovNro = MR.nMovNro " _
 _
              & " and  M1.nMovFlag = " & gMovFlagVigente & " And MR.nMovNroRef = M.nMovNro) And M.nMovFlag = " & gMovFlagVigente & " " _
              & lsOpeFiltro & " and T.nMoneda = " & psMoneda & " and SubString(M.cMovNro,18,2) = '" & Right(psCodAge, 2) & "' "
        Else
          sql = "Select M.cOpeCod, Convert(DateTime,LEFT(M.cMovNro,4)+'/'+SubString(M.cMovNro,5,2)+'/'+SubString(M.cMovNro,7,2)) dFecTran, M.nMovNro , mg.cPersCod cPersCodIF, RTRIM(P.cPersNombre) cNomPers, T.nMovImporte nMonTran, M.cMovDesc cGlosa " _
              & "From Mov M " _
              & "    JOIN MovOpeVarias T ON M.nMovNro = T.nMovNro " _
              & "    JOIN MovGasto mg ON mg.nMovNro = M.nMovNro JOIN Persona P ON P.cPersCod = mg.cPersCod " _
              & "WHere M.nMovNro NOT IN (Select MR.nMovNroRef From Mov M1 JOIN MovRef MR ON M1.nMovNro = MR.nMovNro " _
 _
              & " and  M1.nMovFlag = " & gMovFlagVigente & ") And M.nMovFlag = " & gMovFlagVigente & " " _
              & lsOpeFiltro & " and T.nMoneda = " & psMoneda & " and SubString(M.cMovNro,18,2) = '" & Right(psCodAge, 2) & "' " _
              & IIf(psCodPers = "", "", " and mg.cPersCod = '" & psCodPers & "' ")
        End If
        oCon.AbreConexion
    Else
        sql = "SELECT td.cCodOpe, td.dFecTran, td.nNumTran nMovNro, oo.cCodPers cPersCodIF, p.cNomPers, td.nMonTran, ISNULL(oo.cGlosa,'') cGlosa, '' cCtaContCod " _
            & "FROM   TransAho td " _
            & "       JOIN OtrasOpe oo ON oo.nNumTran = td.nNumTran " _
            & "       JOIN DBPersona..Persona p ON p.cCodPers = oo.cCodPers " _
            & "       JOIN TransRef ref ON Convert(money,ref.cNroTranRef) = td.nNumTran and ref.cCodAgeRef = '" & psCodAge & "' " _
            & "WHERE  td.cFlag is NULL and ref.cNroTran IS NULL and SubString(td.cCodCta,6,1) = '" & psMoneda & "' " _
            & lsOpeFiltro _
            & IIf(psCodPers = "", "", " and oo.cCodPers = '" & psCodPers & "' ")
       oCon.AbreConexion 'Remota Right(psCodAge, 2), True, False, "01"
    End If
    Set CargaIngresoVentanillaPendiente = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function


Public Function ActualizaNegocioReferencia(psNroTran As String, psAgeCod As String, psMovNro As String, Optional psCodOpe As String = "")
On Error GoTo ActualizaNegocioReferenciaErr
lsSql = "UPDATE TransRef SET cNroTran = '" & psMovNro & "', cCodAge = '112" & Mid(psMovNro, 18, 2) & "' WHERE cNroTranRef = " & psNroTran & " and cCodAgeRef = '" & psAgeCod & "' and cNroTran is NULL " & IIf(psCodOpe = "", "", " and cCodOpe = '" & psCodOpe & "'")
   Set oCon = New DConecta
   If oCon.AbreConexion Then 'Remota(Right(psAgeCod, 2), , False, "01")
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
    
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    pnBitCentral = IIf(oConst.LeeConstSistema(gConstSistBitCentral) = "1", True, False)
End Sub
