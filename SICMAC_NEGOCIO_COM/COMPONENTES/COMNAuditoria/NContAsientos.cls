VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NContAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsBaseComunes As String
Dim vsBasePesonas As String
Dim lsSql        As String

Private Sub Class_Initialize()
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Set oImp = New COMDConstSistema.DCOMImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni
vsBaseComunes = oIni.BaseComunes
vsBasePesonas = oIni.BasePersonas
Set oIni = Nothing
End Sub

Public Function GetMayorCuenta(psCtaCod As String, psFechaDel As String, psFechaAl As String, Optional pnImporte As Currency = 0, Optional psFiltro As String = "=", Optional pbME As Boolean = False, Optional pbVerDoc As Boolean = True, Optional pbVerPersona As Boolean = False) As ADODB.Recordset
Dim lsCtaCod As String


If Mid(psCtaCod, 3, 1) = "0" Then
    lsCtaCod = Left(psCtaCod, 2) & "_" & Mid(psCtaCod, 4)
Else
    lsCtaCod = psCtaCod
End If

lsSql = "SELECT Left(a.cCtaContCod," & Len(psCtaCod) & ") cCtaContCod, c.cCtaContDesc, M.cMovNro, " & IIf(pbVerDoc, " ISNULL(Docum.cDocAbrev,'') as cDocAbrev, ISNULL(Doc.cDocNro,'') as cDocNro, ISNULL(Convert(varchar(10),Doc.dDocFecha,103),'') as dDocFecha, ", "") & "  " & IIf(pbVerPersona, " dbo.GetPersonaMov(M.nMovNro)  ", " M.cMovDesc ") & " cMovDesc, " _
     & "       ISNULL(SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END),0) as nDebe, " _
     & "       ISNULL(SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END),0) as nHaber " _
    & IIf(pbME, ",ISNULL(SUM(CASE WHEN me.nMovMEImporte > 0 THEN me.nMovMEImporte END),0) as nDebeME, " _
     & "       ISNULL(SUM(CASE WHEN me.nMovMEImporte < 0 THEN me.nMovMEImporte * -1 END),0) as nHaberME ", "") _
     & " ,m.nmovnro FROM   Mov M JOIN MovCta a ON a.nMovNro = M.nMovNro " & IIf(pbME, " LEFT JOIN MovME me ON me.nMovNro = a.nMovNro and me.nMovItem = a.nMovItem ", "") _
     & "             JOIN " & vsBaseComunes & "CtaCont c ON c.cCtaContCod = Left(a.cCtaContCod," & Len(psCtaCod) & ") " _
     & IIf(pbVerDoc, "LEFT JOIN MovDoc Doc ON Doc.nMovNro = M.nMovNro LEFT JOIN " & vsBaseComunes & "Documento Docum ON Docum.nDocTpo = Doc.nDocTpo ", "") _
     & "WHERE  M.nMovEstado = '" & gMovEstContabMovContable & "' and not M.nMovFlag in ('" & gMovFlagEliminado & "','" & gMovFlagModificado & "','" & gMovFlagExtornado & "') and not M.cOpecod Like '70185%' and a.cCtaContCod Like '" & lsCtaCod & "%' and " _
     & "       substring(M.cmovnro,1,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' " _
     & "GROUP BY Left(a.cCtaContCod," & Len(psCtaCod) & "),a.cCtaContcod, c.cCtaContDesc, M.cMovNro, M.nMovNro, " & IIf(pbVerDoc, " Docum.cDocAbrev, Doc.cDocNro, Doc.dDocFecha, ", "") & " M.cMovDesc " _
     & IIf(pnImporte > 0, "HAVING SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END) " & psFiltro & pnImporte & " " _
     & "or SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END) " & psFiltro & pnImporte, "") _
     & "ORDER BY  M.cMovNro, a.cCtaContcod"


Dim oConec As New COMConecta.DCOMConecta
oConec.AbreConexion
Set GetMayorCuenta = oConec.CargaRecordSet(lsSql)
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetMovPersonaARendir(ByVal psMovNro As Long) As ADODB.Recordset
Dim lsCtaCod As String


lsSql = " Select isnull(p.cPersNombre,'')cPersNombre,isnull(mr.nmovnroref,0)nmovnroref" & _
        " from Mov m " & _
        " inner join MovGasto mg ON mg.nmovnro =m.nmovnro " & _
        " inner join MovRef mr ON mr.nmovnro =m.nmovnro " & _
        " inner join PErsona p ON p.cPersCod = mg.cPersCod " & _
        " where m.nmovnro = " & psMovNro & ""

Dim oConec As New COMConecta.DCOMConecta
oConec.AbreConexion
Set GetMovPersonaARendir = oConec.CargaRecordSet(lsSql)
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetMovPersonaARendirRef(ByVal psMovNro As Long) As ADODB.Recordset
Dim lsCtaCod As String


lsSql = " Select p.cPersNombre" & _
        " from Mov m " & _
        " inner join MovARendir mar ON mar.nmovnro=m.nmovnro " & _
        " inner join Persona p ON p.cPersCod=mar.cPersCod " & _
        " where m.nmovnro=" & psMovNro & ""

Dim oConec As New COMConecta.DCOMConecta
oConec.AbreConexion
Set GetMovPersonaARendirRef = oConec.CargaRecordSet(lsSql)
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetCtaEquivalente(psCtaCod As String, psDH As String, psOpeCod As String) As String
Dim prs As New ADODB.Recordset
lsSql = "SELECT cCtaContCod" & IIf(psDH = "D", "H", "D") & " as cCtaCod FROM OpeCtaCta " _
    & "WHERE cOpeCod = '" & psOpeCod & "' and " _
    & "'" & psCtaCod & "' LIKE cCtaContCod" & IIf(psDH = "D", "D", "H") & " + '%'"
Dim oConec As New COMConecta.DCOMConecta
oConec.AbreConexion
Set prs = oConec.CargaRecordSet(lsSql)
If Not prs.EOF Then
   GetCtaEquivalente = prs!cCtaCod
End If
prs.Close: Set prs = Nothing
End Function

Public Sub LlenaArrayAsiento(aAsiento() As String, sCtaCod As String, pnImporte As Currency)
Dim K As Integer, nPos As Integer
nPos = -1
For K = 1 To UBound(aAsiento, 2)
   If aAsiento(0, K) = sCtaCod Then
      nPos = K
   End If
Next
If nPos = -1 Then
   nPos = UBound(aAsiento, 2) + 1
   ReDim Preserve aAsiento(1, nPos)
End If
aAsiento(0, nPos) = sCtaCod
aAsiento(1, nPos) = Val(aAsiento(1, nPos)) + pnImporte
End Sub

Public Function GetAsientoAjusteTipoCambio(psFecha As String, Optional pbCtaBalance As Boolean = True, Optional psCtaOrdenLIKE As String = "") As ADODB.Recordset
Dim sCond1 As String
Dim sCond2 As String
If pbCtaBalance Then
    sCond1 = " not cos.cCtaContCod LIKE '8%' and "
Else
    sCond1 = " cos.cCtaContCod LIKE '" & psCtaOrdenLIKE & "%' and "
End If
lsSql = "SELECT cos.cCtaContCod, cos.nCtaSaldoImporte,  cls.cCtaCaracter " _
      & "FROM   CtaSaldo cos JOIN CtaContClase cls ON cos.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' " _
      & "WHERE  " & sCond1 & " nCtaSaldoImporte <> 0 and cos.cCtaContCod LIKE '__2%' and cls.nCtaMoneda = 0 and Cos.dCtaSaldoFecha = " _
      & "       ( SELECT MAX(a.dCtaSaldoFecha) FROM CtaSaldo a " _
      & "         WHERE  a.cCtaContCod = cos.cCtaContCod and dCtasaldoFecha <= '" & psFecha & "') ORDER BY cos.cCtaContCod "

Dim oConec As New COMConecta.DCOMConecta
oConec.AbreConexion
Set GetAsientoAjusteTipoCambio = oConec.CargaRecordSet(lsSql)
oConec.CierraConexion
Set oConec = Nothing
End Function


