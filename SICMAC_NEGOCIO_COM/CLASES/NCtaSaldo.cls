VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCtasaldo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A82E291008C"
Option Base 0
Option Explicit
Dim DBConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sConexion As String
Dim sCentralCom As String
Public Function PermiteMntSdoInicial() As Boolean
   On Error GoTo GetFechaSdoInicialErr
   psSql = "SELECT TOP 2 dCtaSaldoFecha FROM CtaSaldo GROUP BY dCtaSaldoFecha "
   Set prs = DBConec.CargaRecordSet(psSql)
   If Not prs.EOF Then
      If prs.RecordCount > 1 Then
         PermiteMntSdoInicial = False
      Else
         PermiteMntSdoInicial = True
      End If
   Else
      PermiteMntSdoInicial = True
   End If
   prs.Close: Set prs = Nothing
   Exit Function
GetFechaSdoInicialErr:
   Call RaiseError(MyUnhandledError, "NCtaSaldo:GetFechaSdoInicial Method")
End Function

Public Function GetFechaSdoInicial() As String
   On Error GoTo GetFechaSdoInicialErr
   psSql = "SELECT MIN(dCtaSaldoFecha) dCtaSaldoFecha FROM CtaSaldo"
   Set prs = DBConec.CargaRecordSet(psSql)
   If Not prs.EOF Then
      If IsNull(prs!dCtaSaldoFecha) Then
         GetFechaSdoInicial = "  /  /    "
      Else
         GetFechaSdoInicial = Format(prs!dCtaSaldoFecha, "dd/mm/yyyy")
      End If
   Else
      GetFechaSdoInicial = "  /  /    "
   End If
   prs.Close: Set prs = Nothing
   Exit Function
GetFechaSdoInicialErr:
   Call RaiseError(MyUnhandledError, "NCtaSaldo:GetFechaSdoInicial Method")
End Function

Public Function ExisteCuentaSaldo(ByVal psCtaCod As String, ByVal psFecha As String) As Boolean
   On Error GoTo ExisteCuentaErr
   Dim clsSdo As New DCtaSaldo
   ExisteCuentaSaldo = False
   Set prs = clsSdo.CargaCtaSaldo(psCtaCod, psFecha)
     If Not (prs.BOF And prs.EOF) Then
         ExisteCuentaSaldo = True
     End If
   prs.Close: Set prs = Nothing
   Set clsSdo = Nothing
   Exit Function
ExisteCuentaErr:
   Call RaiseError(MyUnhandledError, "NCtaSaldo:ExisteCuenta Method")
End Function

Public Sub CierreContableMensual(psFecIni As String, psFecFin As String, psCodUser As String, psFechaAct As String)
Dim sFec As String
Dim sAge As String
Dim sTpo As String
Dim sClave  As String
Dim N       As Currency
Dim sMovNro As String

Dim oMov As DMov
On Error GoTo CierreContableMensualErr
   Dim oGen As New NConstSistemas
   sMovNro = oMov.GeneraMovNro(CDate(Format(psFecFin, "dd/mm/yyyy")), , psCodUser)
   oGen.ActualizaConstSistemas gConstSistCierreMensualCont, sMovNro, psFecFin, , False
   Set oGen = Nothing
   psSql = "Select M.cMovNro, M.nMovNro, ISNULL(O.cOpeGruCod,'') as cTpo " _
        & "From Mov M Join MovCta MC On M.cMovNro = MC.cMovNro " _
        & "           Join " & sCentralCom & "CtaCont C On MC.cCtaContCod = C.cCtaContCod " _
        & "           Join " & sCentralCom & "OpeTpo  O On M.cOpeCod = O.cOpeCod " _
        & "Where M.nMovEstado = " & gMovEstContabMovContable & " and M.nMovFlag <> " & gMovFlagEliminado & " and convert(datetime,LEFT(M.cMovNro,8),101) between " _
        & "      '" & psFecIni & "' And '" & psFecFin & "' AND mc.cCtaContCod <> '' " _
        & "GROUP BY M.cMovNro, ISNULL(O.cOpeGruCod,'') " _
        & "Order by LEFT(M.cMovNro,8), substring(M.cMovNro,18,2), ISNULL(O.cOpeGruCod,'') "
   Set prs = DBConec.CargaRecordSet(psSql)
   Set oMov = New DMov
   sClave = ""
   N = 0
   Do While Not prs.EOF
      DoEvents
      sFec = Left(prs!cMovnro, 8)
      sAge = Mid(prs!cMovnro, 18, 2)
      sTpo = prs!cTpo
      If sClave <> sFec & sAge & sTpo Then
         sClave = sFec & sAge & sTpo
         N = N + 1
      End If
      oMov.ActualizaMovCont prs!nMovNro, , N
      prs.MoveNext
   Loop
   Set oMov = Nothing
Exit Sub
CierreContableMensualErr:
   Call RaiseError(MyUnhandledError, "NCtaSaldo:CierreContableMensual Method")
End Sub

Public Function GetCtaSaldo(psCtaContCod As String, psFecha As String) As Currency
   On Error GoTo GetCtaSaldoErr
   psSql = "SELECT cCtaContCod, dCtaSaldoFecha, nCtaSaldoImporte FROM CtaSaldo cs WHERE cCtaContCod = '" & psCtaContCod & "' AND dCtaSaldoFecha = (SELECT MAX(dCtaSaldoFecha) FROM CtaSaldo cs1 WHERE cs1.cCtaContCod = cs.cCtaContCod and dCtaSaldoFecha <= '" & psFecha & "')"
   Set prs = DBConec.CargaRecordSet(psSql)
   If Not prs.EOF Then
      GetCtaSaldo = prs!nCtaSaldoImporte
   Else
      GetCtaSaldo = 0
   End If
   Exit Function
GetCtaSaldoErr:
   Call RaiseError(MyUnhandledError, "DCtaSaldo:GetCtaSaldo Method")
End Function

Private Sub Class_Initialize()
Dim oIni As New COMConecta.DCOMClasIni
   sConexion = oIni.CadenaConexion
   sCentralCom = oIni.BaseComunes
Set oIni = Nothing

Set DBConec = New DConecta
DBConec.AbreConexion sConexion

End Sub

Private Sub Class_Terminate()
DBConec.CierraConexion
Set DBConec = Nothing
End Sub
