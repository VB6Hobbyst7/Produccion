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
Dim vsConexion As String
Dim vsBaseComunes As String
Dim vsBasePesonas As String
Dim lsSQL        As String

Private Sub Class_Initialize()
Dim oIni As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni
vsConexion = oIni.CadenaConexion
vsBaseComunes = oIni.BaseComunes
vsBasePesonas = oIni.BasePersonas
Set oIni = Nothing
End Sub


Public Sub GrabaAsientoContable(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal pnMonto As Currency, _
                                ByVal rsCtaCont As ADODB.Recordset, _
                                ByVal rsObjetos As ADODB.Recordset, _
                                Optional ByVal psMovNroActual As String = "", _
                                Optional ByVal psMovNroAnt As String = "", _
                                Optional ByVal pnMovNroAnt As Long = 0)
Dim lnMovNro   As Long
Dim lnItem     As Long
Dim lnOrdenObj As Long
Dim lnImporte  As Currency
Dim lsSubCta   As String
Dim lnCol      As Integer
Dim i As Integer

Dim lsAgeCod As String
Dim lsAreaCod As String
Dim oMov       As DMov
Dim lbTransac As Boolean
On Error GoTo GrabaAsientoContableErr
Set oMov = New DMov

lbTransac = True
oMov.BeginTrans
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovCont lnMovNro, pnMonto, 0, "0"

lnItem = 0
If Not rsCtaCont Is Nothing Then
   Do While Not rsCtaCont.EOF
      If rsCtaCont!Código <> "" Then
         For lnCol = 4 To rsCtaCont.Fields.Count
            lnImporte = Val(rsCtaCont(lnCol - 1))
            If lnImporte > 0 Then
               lnItem = lnItem + 1
               If lnCol < 6 Then
                  oMov.InsertaMovCta lnMovNro, lnItem, rsCtaCont!Código, IIf(lnCol = 4, lnImporte, lnImporte * -1)
               Else
                  oMov.InsertaMovMe lnMovNro, lnItem, IIf(lnCol = 6, lnImporte, lnImporte * -1)
               End If
               
               If Not rsObjetos Is Nothing Then
                  rsObjetos.MoveFirst
                  Do While Not rsObjetos.EOF
                     If rsObjetos!ItemCtaCont = rsCtaCont!ItemCtaCont Then
                        lnOrdenObj = rsObjetos!Orden
                        Select Case rsObjetos!cObjetoCod
                           Case ObjCMACAgencias
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!cObjetoCod
                                oMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código, ""
                           Case ObjCMACArea
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!cObjetoCod
                                oMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, lnOrdenObj, "", rsObjetos!Código
                           Case ObjCMACAgenciaArea
                                lsAgeCod = Mid(rsObjetos!Código, 4, 2)
                                lsAreaCod = Mid(rsObjetos!Código, 1, 3)
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!cObjetoCod
                                oMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, lnOrdenObj, lsAgeCod, lsAreaCod
                           Case ObjPersona
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!cObjetoCod
                                oMov.InsertaMovObjPers lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código
                           Case ObjDescomEfectivo
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!cObjetoCod
                                oMov.InsertaMovObjEfectivo lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código
                           Case ObjEntidadesFinancieras
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!cObjetoCod
                                'falta insertar movObjIF
                           Case Else
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código
                        End Select
                     End If
                     rsObjetos.MoveNext
                  Loop
               End If
            End If
         Next
      End If
      rsCtaCont.MoveNext
   Loop
End If
If psMovNroAnt <> "" And pnMovNroAnt > 0 Then  'Es una modificación
   oMov.EliminaMov pnMovNroAnt    'Eliminamos el Mov Modificado
   oMov.ActualizaMovimiento lnMovNro, pnMovNroAnt, False
   oMov.InsertaMovModifica psMovNroActual, psMovNroAnt, psMovNro
End If
oMov.CommitTrans
lbTransac = False
Set oMov = Nothing
Exit Sub
GrabaAsientoContableErr:
   If lbTransac Then
      oMov.RollbackTrans
   End If
   lbTransac = False
   Set oMov = Nothing
   Call RaiseError(MyUnhandledError, "NContAsientos:GrabaAsientoContable Method")
End Sub

Public Function GetAsientoConsulta(psBusCond As String, psOpeCond As String, psAgeCond As String, psDocCond As String, psCtaCond As String, psMovCond As String, psFecCond As String) As ADODB.Recordset
On Error GoTo GetAsientoConsultaErr
Dim oConec As New DConecta
lsSQL = "SELECT m.cMovNro, m.nMovNro, a.nMovItem, a.cCtaContCod, b.cCtaContDesc, " _
     & "       SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END) as nDebe, " _
     & "       SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END) as nHaber, " _
     & "       SUM(CASE WHEN me.nMovMeImporte > 0 THEN me.nMovMeImporte END) as nDebeME, " _
     & "       SUM(CASE WHEN me.nMovMeImporte < 0 THEN me.nMovMeImporte * -1 END) nHaberME " _
     & " FROM  MovCta a      JOIN Mov M ON M.nMovNro = a.nMovNro " _
     & "                LEFT JOIN MovMe Me ON Me.nMovNro = a.nMovNro and Me.nMovItem = a.nMovItem " _
     & "                LEFT JOIN " & vsBaseComunes & "CtaCont b ON a.cCtaContCod = b.cCtaContCod " _
     & "                     JOIN " & vsBaseComunes & "OpeTpo Op ON Op.cOpeCod = M.cOpeCod " _
     & IIf(psDocCond = "", "", " JOIN MovDoc MD ON MD.nMovNro = M.nMovNro ") _
     & "WHERE " & psFecCond & " " & psBusCond & " " & psCtaCond & " " & psMovCond & " " & psAgeCond & " " & psOpeCond & " " & psDocCond & " " _
     & "GROUP BY cMovNro, m.nMovNro, a.nMovItem, a.cCtaContCod, b.cCtaContDesc " _
     & "ORDER BY cMovNro, a.nMovItem"
     
oConec.AbreConexion vsConexion
Set GetAsientoConsulta = oConec.CargaRecordSet(lsSQL)
oConec.CierraConexion
Set oConec = Nothing

Exit Function
GetAsientoConsultaErr:
   Call RaiseError(MyUnhandledError, "NContAsientos:GetAsientoConsulta Method")
End Function

Public Function GetAsientoAjusteTipoCambio(psFecha As String) As ADODB.Recordset
Dim sCond1 As String
Dim sCond2 As String

lsSQL = "SELECT cos.cCtaContCod, cos.nCtaSaldoImporte,  cls.cCtaCaracter " _
      & "FROM   CtaSaldo cos JOIN " & vsBaseComunes & "CtaContClase cls ON cos.cCtaContCod LIKE cls.cCtaContCod + '%' " _
      & "WHERE  nCtaSaldoImporte <> 0 and cos.cCtaContCod LIKE '__2%' and cls.nCtaMoneda = 0 and Cos.dCtaSaldoFecha = " _
      & "       ( SELECT MAX(a.dCtaSaldoFecha) FROM CtaSaldo a " _
      & "         WHERE  a.cCtaContCod = cos.cCtaContCod and dCtasaldoFecha <= '" & psFecha & "') "


Dim oConec As New DConecta
oConec.AbreConexion
Set GetAsientoAjusteTipoCambio = oConec.CargaRecordSet(lsSQL)
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetMayorCuenta(psCtaCod As String, psFechaDel As String, psFechaAl As String, Optional pnImporte As Currency = 0, Optional psFiltro As String = "=") As ADODB.Recordset

lsSQL = "SELECT a.cCtaContCod, c.cCtaContDesc, M.cMovNro, ISNULL(Docum.cDocAbrev,'') as cDocAbrev, ISNULL(Doc.cDocNro,'') as cDocNro, ISNULL(Convert(varchar(10),Doc.dDocFecha,103),'') as dDocFecha, M.cMovDesc, " _
     & "       ISNULL(SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END),0) as nDebe, " _
     & "       ISNULL(SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END),0) as nHaber " _
     & "FROM   Mov M JOIN MovCta a ON a.nMovNro = M.nMovNro " _
     & "             JOIN " & vsBaseComunes & "CtaCont c ON c.cCtaContCod = a.cCtaContCod " _
     & "        LEFT JOIN MovDoc Doc ON Doc.nMovNro = M.nMovNro " _
     & "        LEFT JOIN " & vsBaseComunes & "Documento Docum ON Docum.nDocTpo = Doc.nDocTpo " _
     & "WHERE  M.nMovEstado = '" & gMovEstContabMovContable & "' and M.nMovFlag <> '" & gMovFlagEliminado & "' and a.cCtaContCod Like '" & psCtaCod & "%' and " _
     & "       substring(M.cmovnro,1,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' " _
     & "GROUP BY a.cCtaContCod, c.cCtaContDesc, M.cMovNro, Docum.cDocAbrev, Doc.cDocNro, Doc.dDocFecha, M.cMovDesc " _
     & IIf(pnImporte > 0, "HAVING SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END) " & psFiltro & pnImporte & " " _
     & "or SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END) " & psFiltro & pnImporte, "")
Dim oConec As New DConecta
oConec.AbreConexion
Set GetMayorCuenta = oConec.CargaRecordSet(lsSQL)
oConec.CierraConexion
Set oConec = Nothing
End Function


Public Function GetCtaEquivalente(psCtaCod As String, psDH As String, psOpeCod As String) As String
Dim prs As New ADODB.Recordset
lsSQL = "SELECT cCtaContCod" & IIf(psDH = "D", "H", "D") & " as cCtaCod FROM OpeCtaCta " _
    & "WHERE cOpeCod = '" & psOpeCod & "' and " _
    & "'" & psCtaCod & "' LIKE cCtaContCod" & IIf(psDH = "D", "D", "H") & " + '%'"
Dim oConec As New DConecta
oConec.AbreConexion
Set prs = oConec.CargaRecordSet(lsSQL)
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

