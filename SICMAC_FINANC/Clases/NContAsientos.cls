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
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni
vsBaseComunes = oIni.BaseComunes
vsBasePesonas = oIni.BasePersonas
Set oIni = Nothing
End Sub

Public Sub GrabaAsientoContable(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal pnMonto As Currency, _
                                ByVal rsDoc As ADODB.Recordset, _
                                ByVal rsCtaCont As ADODB.Recordset, _
                                ByVal rsObjetos As ADODB.Recordset, _
                                Optional ByVal psMovNroActual As String = "", _
                                Optional ByVal psMovNroAnt As String = "", _
                                Optional ByVal pnMovNroAnt As Long = 0, Optional ByVal psPersCod As String = "", Optional ByVal psDestino As String = "")
Dim lnMovNro   As Long
Dim lnItem     As Long
Dim lnOrdenObj As Long
Dim lnImporte  As Currency
Dim lsSubCta   As String
Dim lnCol      As Integer
Dim I As Integer

Dim lsAgeCod As String
Dim lsAreaCod As String
Dim oMov       As DMov
Dim lbTransac As Boolean
On Error GoTo GrabaAsientoContableErr
Set oMov = New DMov

Dim oFun As New NConstSistemas
Dim lsCtaConvDol As String
Dim lsCtaConvSol As String
Dim lnTipoDoc As Integer

lsCtaConvDol = oFun.LeeConstSistema(gConstSistCtaConversionMEDol)
lsCtaConvSol = oFun.LeeConstSistema(gConstSistCtaConversiónMESoles)

Set oFun = Nothing

lbTransac = True
oMov.BeginTrans
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovCont lnMovNro, pnMonto, 0, "0"
If psPersCod <> "" Then
    oMov.InsertaMovGasto lnMovNro, psPersCod, psDestino
End If
If Not rsDoc Is Nothing Then
    lnTipoDoc = rsDoc!Tipo
    Do While Not rsDoc.EOF
        If rsDoc!Tipo > 0 And rsDoc!Numero <> "" And ValidaFecha(rsDoc!fecha) = "" Then
            oMov.InsertaMovDoc lnMovNro, rsDoc!Tipo, rsDoc!Numero, Format(rsDoc!fecha, gsFormatoFecha)
        End If
        rsDoc.MoveNext
    Loop
End If

lnItem = 0
If Not rsCtaCont Is Nothing Then
   Do While Not rsCtaCont.EOF
      If rsCtaCont!Código <> "" Then
         For lnCol = 4 To 5
            lnImporte = Val(rsCtaCont(lnCol - 1))
            If lnImporte > 0 Then
              'Nota de Credito
              If lnTipoDoc <> 7 Then
               If lnCol = 4 And Mid(rsCtaCont!Código, 3, 1) = "2" And rsCtaCont!Código <> lsCtaConvSol And rsCtaCont!Código <> lsCtaConvDol Then
                  If nVal(rsCtaCont(lnCol + 1)) = 0 And nVal(rsCtaCont(lnCol - 1)) > 0.03 And Not Left(rsCtaCont!Código, 4) = "5128" And Not Left(rsCtaCont!Código, 4) = "4128" And Not Left(rsCtaCont!Código, 4) = "2124" Then
                     Err.Raise 50001, "NContAsiento: GrabaAsientoContable", "Cuenta Contable de Dolares debe Ingresarse con su monto equivalente en Dólares"
                  End If
               End If
               ' COMENTAR SOLO PARA CONTABILIDAD
               'If lnCol = 5 And Mid(rsCtaCont!Código, 3, 1) = "2" And rsCtaCont!Código <> lsCtaConvSol And rsCtaCont!Código <> lsCtaConvDol Then
               '   If nVal(rsCtaCont(lnCol + 1)) = 0 And nVal(rsCtaCont(lnCol - 1)) > 0.03 And Not Left(rsCtaCont!Código, 4) = "5128" And Not Left(rsCtaCont!Código, 4) = "4128" Then
               '      Err.Raise 50001, "NContAsiento: GrabaAsientoContable", "Cuenta Contable de Dolares debe Ingresarse con su monto equivalente en Dólares"
               '   End If
               'End If
              End If
               
               lnItem = lnItem + 1
               oMov.InsertaMovCta lnMovNro, lnItem, rsCtaCont!Código, IIf(lnCol = 4, lnImporte, lnImporte * -1)
               lnImporte = Val(rsCtaCont(lnCol + 1))
               If lnImporte > 0 Then
                  oMov.InsertaMovMe lnMovNro, lnItem, IIf(lnCol = 4, lnImporte, lnImporte * -1)
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
                           Case ObjDescomEfectivo
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!cObjetoCod
                                oMov.InsertaMovObjEfectivo lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código
                           Case ObjEntidadesFinancieras
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!cObjetoCod
                                oMov.InsertaMovObjIF lnMovNro, lnItem, lnOrdenObj, Mid(rsObjetos!Código, 4, 13), Left(rsObjetos!Código, 2), Mid(rsObjetos!Código, 18, 7)
                           Case ObjBienesServicios
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, ObjBienesServicios
                                oMov.InsertaMovBS lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código
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
   oMov.EliminaMov psMovNroAnt       'Eliminamos el Mov Modificado
   oMov.ActualizaMovimiento psMovNro, pnMovNroAnt, , False
   oMov.InsertaMovModifica psMovNroActual, psMovNroAnt, psMovNro
End If
oMov.ActualizaSaldoMovimiento psMovNro, "+"
oMov.CommitTrans
lbTransac = False
Set oMov = Nothing
Exit Sub
GrabaAsientoContableErr:
   If lbTransac Then
      oMov.RollbackTrans
   End If
   lbTransac = False
   Call RaiseError(MyUnhandledError, "NContAsientos:GrabaAsientoContable Method")
   Set oMov = Nothing
End Sub

Public Sub GrabaAsientoContableMoneda(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal pnMonto As Currency, _
                                ByVal rsDoc As ADODB.Recordset, _
                                ByVal rsCtaCont As ADODB.Recordset, _
                                ByVal rsObjetos As ADODB.Recordset, _
                                Optional ByVal psMovNroActual As String = "", _
                                Optional ByVal psMovNroAnt As String = "", _
                                Optional ByVal pnMovNroAnt As Long = 0, Optional ByVal psPersCod As String = "", _
                                Optional lbRegistraPendiente As Boolean = False, _
                                Optional lbRegulaPendiente As Boolean = False, Optional psAgeCodRef As String = "", _
                                Optional pbBitCentral As Boolean = False, Optional rsPend As ADODB.Recordset)
Dim lnMovNro   As Long
Dim lnItem     As Long
Dim lnOrdenObj As Long
Dim lnImporte  As Currency
Dim lsSubCta   As String
Dim lnCol      As Integer
Dim I As Integer

Dim lsAgeCod  As String
Dim lsAreaCod As String
Dim oMov      As DMov
Dim lbTransac As Boolean
Dim sSQL      As String
Dim lsMsgErr  As String
Dim lsTpoDoc As String


On Error GoTo GrabaAsientoContableMonedaErr
Set oMov = New DMov

lbTransac = True
oMov.BeginTrans
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovCont lnMovNro, pnMonto, 0, "0"
If psPersCod <> "" Then
    oMov.InsertaMovGasto lnMovNro, psPersCod, ""
End If
If Not rsDoc Is Nothing Then
    Do While Not rsDoc.EOF
        lsTpoDoc = rsDoc!Tipo
        If rsDoc!Tipo > 0 And rsDoc!Numero <> "" And ValidaFecha(rsDoc!fecha) = "" Then
            oMov.InsertaMovDoc lnMovNro, rsDoc!Tipo, rsDoc!Numero, Format(rsDoc!fecha, gsFormatoFecha)
        End If
        rsDoc.MoveNext
    Loop
End If

lnItem = 0
If Not rsCtaCont Is Nothing Then
   Do While Not rsCtaCont.EOF
      If rsCtaCont!Código <> "" Then
         For lnCol = 4 To 5
            If Mid(psOpeCod, 3, 1) = "1" Then
                lnImporte = nVal(rsCtaCont(lnCol - 1))
            Else
                lnImporte = nVal(rsCtaCont(lnCol + 1))
            End If
            If lnImporte > 0 Then
               lnItem = lnItem + 1
               oMov.InsertaMovCta lnMovNro, lnItem, rsCtaCont!Código, IIf(lnCol = 4, lnImporte, lnImporte * -1)
               If lbRegulaPendiente And lnItem = 1 Then
                  If Not IsNull(rsPend) Then
                    rsPend.MoveFirst
                    Do While Not rsPend.EOF
                       If rsPend!Ok = 1 Then
                            oMov.InsertaMovRef lnMovNro, rsPend!nMovNro, IIf(pbBitCentral, "", psAgeCodRef)
                            oMov.ActualizaMovPendientesRend rsPend!nMovNro, rsCtaCont!Código, rsPend!Rendicion
                            
                            If psAgeCodRef <> "" Then
                               'Solo para grabar Negocio Distribuido
                               If Not pbBitCentral Then
                                   Dim oCon As New DConecta
                                   oCon.AbreConexion
                                   If oCon.AbreConexion Then 'Remota(Right(psAgeCodRef, 2))
                                       sSQL = "UPDATE TransRef SET cNroTran = '" & psMovNro & "', cCodAge = '112" & Mid(psMovNro, 18, 2) & "' WHERE cNroTranRef = '" & rsPend!nMovNro & "' and cCodAgeRef = '112" & Right(psAgeCodRef, 2) & "' and cCodOpe = '" & rsPend!cCodOpe & "' and cNroTran is NULL "
                                       oCon.Ejecutar sSQL
                                   End If
                                   oCon.CierraConexion
                               End If
                            End If
                        End If
                        rsPend.MoveNext
                    Loop
                  End If
               End If
               
               'Registro de Pendiente
               If lbRegistraPendiente Then
                    If oMov.CuentaEsPendiente(rsCtaCont!Código, , IIf(lnCol = 4, "D", "A")) Then
                       oMov.InsertaMovPendientesRend lnMovNro, rsCtaCont!Código, lnImporte
                    End If
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
                           Case ObjDescomEfectivo
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!cObjetoCod
                                oMov.InsertaMovObjEfectivo lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código
                           Case ObjEntidadesFinancieras
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!cObjetoCod
                                oMov.InsertaMovObjIF lnMovNro, lnItem, lnOrdenObj, Mid(rsObjetos!Código, 4, 13), Left(rsObjetos!Código, 2), Mid(rsObjetos!Código, 18, 7)
                           Case ObjBienesServicios
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, ObjBienesServicios
                                oMov.InsertaMovBS lnMovNro, lnItem, lnOrdenObj, rsObjetos!cObjetoCod
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

'Cambio Liquidaciones
If (psOpeCod = "421009" Or psOpeCod = "422009") And lsTpoDoc = "48" Then
    oMov.InsertaMovRef lnMovNro, lnMovNro, ""
End If


If psMovNroAnt <> "" And pnMovNroAnt > 0 Then  'Es una modificación
   'Antes de eliminar Recalcula Saldo
   oMov.ActualizaMovPendientesRendCambio pnMovNroAnt
   oMov.EliminaMov psMovNroAnt      'Eliminamos el Mov Modificado
   oMov.ActualizaMovimiento psMovNro, pnMovNroAnt, , False
   oMov.InsertaMovModifica psMovNroActual, psMovNroAnt, psMovNro
   
   'Actualiza saldo pendiente con nuevos datos
   oMov.ActualizaMovPendientesRendCambio lnMovNro, "-"
End If
If Mid(psOpeCod, 3, 1) = "2" Then
    oMov.GeneraMovME lnMovNro, psMovNro
End If
oMov.ActualizaSaldoMovimiento psMovNro, "+"
oMov.CommitTrans
lbTransac = False
Set oMov = Nothing
Exit Sub
GrabaAsientoContableMonedaErr:
    lsMsgErr = Err.Description
   If lbTransac Then
      oMov.RollbackTrans
   End If
   lbTransac = False
   Set oMov = Nothing
   Call RaiseError(MyUnhandledError, "NContAsientos:GrabaAsientoContableMoneda Method")
End Sub


'Public Function GetAsientoConsulta(psBusCond As String, psOpeCond As String, psAgeCond As String, psDocCond As String, psCtaCond As String, psMovCond As String, psFecCond As String) As ADODB.Recordset
'On Error GoTo GetAsientoConsultaErr
'Dim oConec As New DConecta
'lsSql = "SELECT m.cMovNro, m.nMovNro, a.nMovItem, a.cCtaContCod, b.cCtaContDesc, " _
'     & "       SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END) as nDebe, " _
'     & "       SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END) as nHaber, " _
'     & "       SUM(CASE WHEN me.nMovMeImporte > 0 THEN me.nMovMeImporte END) as nDebeME, " _
'     & "       SUM(CASE WHEN me.nMovMeImporte < 0 THEN me.nMovMeImporte * -1 END) nHaberME, M.nMovFlag, M.nMovEstado " _
'     & " FROM  MovCta a      JOIN Mov M ON M.nMovNro = a.nMovNro " _
'     & "                LEFT JOIN MovMe Me ON Me.nMovNro = a.nMovNro and Me.nMovItem = a.nMovItem " _
'     & "                LEFT JOIN " & vsBaseComunes & "CtaCont b ON a.cCtaContCod = b.cCtaContCod " _
'     & "                     JOIN " & vsBaseComunes & "OpeTpo Op ON Op.cOpeCod = M.cOpeCod " _
'     & IIf(psDocCond = "", "", " JOIN MovDoc MD ON MD.nMovNro = M.nMovNro ") _
'     & "WHERE " & psFecCond & " " & psBusCond & " " & psCtaCond & " " & psMovCond & " " & psAgeCond & " " & psOpeCond & " " & psDocCond & " " _
'     & "GROUP BY cMovNro, m.nMovNro, a.nMovItem, a.cCtaContCod, b.cCtaContDesc, M.nMovFlag, M.nMovEstado " _
'     & "ORDER BY cMovNro, a.nMovItem"
'
'oConec.AbreConexion
'Set GetAsientoConsulta = oConec.CargaRecordSet(lsSql)
'oConec.CierraConexion
'Set oConec = Nothing
'
'Exit Function
'GetAsientoConsultaErr:
'   Call RaiseError(MyUnhandledError, "NContAsientos:GetAsientoConsulta Method")
'End Function

'Public Function GetMovimientoConsulta(psBusCond As String, psOpeCond As String, psAgeCond As String, psDocCond As String, psCtaCond As String, psMovCond As String, psFecCond As String) As ADODB.Recordset
'On Error GoTo GetMovimientoConsultaErr
'Dim oConec As New DConecta
'lsSql = "SELECT m.cMovNro, m.nMovEstado, m.cOpeCod, ISNULL(m1.nMovMonto,0), m.cMovDesc, m.nMovFlag, m.nMovNro " _
'     & " FROM  Mov m JOIN OpeTpo ot ON ot.cOpeCod = m.cOpecod LEFT JOIN MovCont M1 ON M.nMovNro = m1.nMovNro " _
'     & IIf(psCtaCond = "", "", " JOIN MovCta Mc ON Mc.nMovNro = M.nMovNro ") _
'     & IIf(psDocCond = "", "", " JOIN MovDoc MD ON MD.nMovNro = M.nMovNro ") _
'     & "WHERE " & psFecCond & " " & psBusCond & " " & psCtaCond & " " & psMovCond & " " & psAgeCond & " " & psOpeCond & " " & psDocCond & " " _
'     & "ORDER BY m.cMovNro"
'
'oConec.AbreConexion
'Set GetMovimientoConsulta = oConec.CargaRecordSet(lsSql)
'oConec.CierraConexion
'Set oConec = Nothing
'
'Exit Function
'GetMovimientoConsultaErr:
'   Call RaiseError(MyUnhandledError, "NContAsientos:GetMovimientoConsulta Method")
'End Function

'Public Function GetAsientoAjusteTipoCambio(psFecha As String, Optional pbCtaBalance As Boolean = True, Optional psCtaOrdenLIKE As String = "") As ADODB.Recordset
'Dim sCond1 As String
'Dim sCond2 As String
'If pbCtaBalance Then
'    sCond1 = " not cos.cCtaContCod LIKE '8%' and "
'Else
'    sCond1 = " cos.cCtaContCod LIKE '" & psCtaOrdenLIKE & "%' and "
'End If
'lsSql = "SELECT cos.cCtaContCod, cos.nCtaSaldoImporte,  cls.cCtaCaracter " _
'      & "FROM   CtaSaldo cos JOIN CtaContClase cls ON cos.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' " _
'      & "WHERE  " & sCond1 & " nCtaSaldoImporte <> 0 and cos.cCtaContCod LIKE '__2%' and cls.nCtaMoneda = 0 and Cos.dCtaSaldoFecha = " _
'      & "       ( SELECT MAX(a.dCtaSaldoFecha) FROM CtaSaldo a " _
'      & "         WHERE  a.cCtaContCod = cos.cCtaContCod and dCtasaldoFecha <= '" & psFecha & "') ORDER BY cos.cCtaContCod "
'
'Dim oConec As New DConecta
'oConec.AbreConexion
'Set GetAsientoAjusteTipoCambio = oConec.CargaRecordSet(lsSql)
'oConec.CierraConexion
'Set oConec = Nothing
'End Function


'Public Function GetMayorCuenta(psCtaCod As String, psFechaDel As String, psFechaAl As String, Optional pnImporte As Currency = 0, Optional psFiltro As String = "=", Optional pbME As Boolean = False, Optional pbVerDoc As Boolean = True, Optional pbVerPersona As Boolean = False) As ADODB.Recordset
'Dim lsCtaCod As String
'
'
'If Mid(psCtaCod, 3, 1) = "0" Then
'    lsCtaCod = Left(psCtaCod, 2) & "_" & Mid(psCtaCod, 4)
'Else
'    lsCtaCod = psCtaCod
'End If
'
'lsSql = "SELECT Left(a.cCtaContCod," & Len(psCtaCod) & ") cCtaContCod, c.cCtaContDesc, M.cMovNro, " & IIf(pbVerDoc, " ISNULL(Docum.cDocAbrev,'') as cDocAbrev, ISNULL(Doc.cDocNro,'') as cDocNro, ISNULL(Convert(varchar(10),Doc.dDocFecha,103),'') as dDocFecha, ", "") & "  " & IIf(pbVerPersona, " dbo.GetPersonaMov(M.nMovNro)  ", " M.cMovDesc ") & " cMovDesc, " _
'     & "       ISNULL(SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END),0) as nDebe, " _
'     & "       ISNULL(SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END),0) as nHaber " _
'    & IIf(pbME, ",ISNULL(SUM(CASE WHEN me.nMovMEImporte > 0 THEN me.nMovMEImporte END),0) as nDebeME, " _
'     & "       ISNULL(SUM(CASE WHEN me.nMovMEImporte < 0 THEN me.nMovMEImporte * -1 END),0) as nHaberME ", "") _
'     & "FROM   Mov M JOIN MovCta a ON a.nMovNro = M.nMovNro " & IIf(pbME, " LEFT JOIN MovME me ON me.nMovNro = a.nMovNro and me.nMovItem = a.nMovItem ", "") _
'     & "             JOIN " & vsBaseComunes & "CtaCont c ON c.cCtaContCod = Left(a.cCtaContCod," & Len(psCtaCod) & ") " _
'     & IIf(pbVerDoc, "LEFT JOIN MovDoc Doc ON Doc.nMovNro = M.nMovNro LEFT JOIN " & vsBaseComunes & "Documento Docum ON Docum.nDocTpo = Doc.nDocTpo ", "") _
'     & "WHERE  M.nMovEstado = '" & gMovEstContabMovContable & "' and not M.nMovFlag in ('" & gMovFlagEliminado & "','" & gMovFlagModificado & "') and a.cCtaContCod Like '" & lsCtaCod & "%' and " _
'     & "       substring(M.cmovnro,1,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' " _
'     & "GROUP BY Left(a.cCtaContCod," & Len(psCtaCod) & "), c.cCtaContDesc, M.cMovNro, M.nMovNro, " & IIf(pbVerDoc, " Docum.cDocAbrev, Doc.cDocNro, Doc.dDocFecha, ", "") & " M.cMovDesc " _
'     & IIf(pnImporte > 0, "HAVING SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END) " & psFiltro & pnImporte & " " _
'     & "or SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END) " & psFiltro & pnImporte, "") _
'     & "ORDER BY  M.cMovNro, a.cCtaContcod"
'
'Dim oConec As New DConecta
'oConec.AbreConexion
'Set GetMayorCuenta = oConec.CargaRecordSet(lsSql)
'oConec.CierraConexion
'Set oConec = Nothing
'End Function

'Public Function GetMayorCuenta(psCtaCod As String, psFechaDel As String, psFechaAl As String, Optional pnImporte As Currency = 0, Optional psFiltro As String = "=", Optional pbME As Boolean = False, Optional pbVerDoc As Boolean = True, Optional pbVerPersona As Boolean = False) As ADODB.Recordset
'Dim lsCtaCod As String


'If Mid(psCtaCod, 3, 1) = "0" Then
'    lsCtaCod = Left(psCtaCod, 2) & "_" & Mid(psCtaCod, 4)
'Else
'    lsCtaCod = psCtaCod
'End If

'lsSql = "SELECT Left(a.cCtaContCod," & Len(psCtaCod) & ") cCtaContCod, c.cCtaContDesc, M.cMovNro, " & IIf(pbVerDoc, " ISNULL(Docum.cDocAbrev,'') as cDocAbrev, ISNULL(Doc.cDocNro,'') as cDocNro, ISNULL(Convert(varchar(10),Doc.dDocFecha,103),'') as dDocFecha, ", "") & "  " & IIf(pbVerPersona, " dbo.GetPersonaMov(M.nMovNro)  ", " M.cMovDesc ") & " cMovDesc, " _
'     & "       ISNULL(SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END),0) as nDebe, " _
'     & "       ISNULL(SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END),0) as nHaber " _
'    & IIf(pbME, ",ISNULL(SUM(CASE WHEN me.nMovMEImporte > 0 THEN me.nMovMEImporte END),0) as nDebeME, " _
'     & "       ISNULL(SUM(CASE WHEN me.nMovMEImporte < 0 THEN me.nMovMEImporte * -1 END),0) as nHaberME ", "") _
'     & " ,isnull(p.cPersNombre,'')cPersNombreGasto,isnull(p1.cPersNombre,'')cPersNombreRendir " _
'     & "FROM   Mov M JOIN MovCta a ON a.nMovNro = M.nMovNro " & IIf(pbME, " LEFT JOIN MovME me ON me.nMovNro = a.nMovNro and me.nMovItem = a.nMovItem ", "") _
'     & "             JOIN " & vsBaseComunes & "CtaCont c ON c.cCtaContCod = Left(a.cCtaContCod," & Len(psCtaCod) & ") " _
'     & IIf(pbVerDoc, "LEFT JOIN MovDoc Doc ON Doc.nMovNro = M.nMovNro LEFT JOIN " & vsBaseComunes & "Documento Docum ON Docum.nDocTpo = Doc.nDocTpo ", "") _
'     & " LEFT Join MovGasto mg ON mg.nmovnro=m.nmovnro" _
'     & " LEFT Join Persona p ON p.cPersCod=mg.cPErsCod" _
'     & " LEFT Join MovRef mr ON mr.nmovnro=m.nmovnro" _
'     & " LEFT Join (Select nmovnro,cPErsCod FROM MovARendir  )X  ON x.nmovnro= mr.nmovnroref " _
'     & " LEFT join Persona p1 ON p1.cPersCod = x.cPersCod " _
'     & " WHERE  M.nMovEstado = '" & gMovEstContabMovContable & "' and not M.nMovFlag in ('" & gMovFlagEliminado & "','" & gMovFlagModificado & "') and a.cCtaContCod Like '" & lsCtaCod & "%' and " _
'     & "       substring(M.cmovnro,1,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' " _
'     & "GROUP BY Left(a.cCtaContCod," & Len(psCtaCod) & "), c.cCtaContDesc, M.cMovNro, M.nMovNro, " & IIf(pbVerDoc, " Docum.cDocAbrev, Doc.cDocNro, Doc.dDocFecha, ", "") & " M.cMovDesc " _
'     & IIf(pnImporte > 0, "HAVING SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END) " & psFiltro & pnImporte & " " _
'     & "or SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END) " & psFiltro & pnImporte, "") _
'     & " ,p.cPersNombre,p1.cPersNombre ORDER BY  M.cMovNro, a.cCtaContcod"

'Dim oConec As New DConecta
'oConec.AbreConexion
'Set GetMayorCuenta = oConec.CargaRecordSet(lsSql)
'oConec.CierraConexion
'Set oConec = Nothing
'End Function
'EJVG20130304 Se agregó psTabla
'Public Function GetMayorCuenta(psCtaCod As String, psFechaDel As String, psFechaAl As String, Optional pnImporte As Currency = 0, Optional psFiltro As String = "=", Optional pbME As Boolean = False, Optional pbVerDoc As Boolean = True, Optional pbVerPersona As Boolean = False) As ADODB.Recordset
Public Function GetMayorCuenta(psCtaCod As String, psFechaDel As String, psFechaAl As String, Optional pnImporte As Currency = 0, Optional psFiltro As String = "=", Optional pbME As Boolean = False, Optional pbVerDoc As Boolean = True, Optional pbVerPersona As Boolean = False, Optional psTabla As String = "CtaCont") As ADODB.Recordset
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
     & "             JOIN " & vsBaseComunes & psTabla & " c ON c.cCtaContCod = Left(a.cCtaContCod," & Len(psCtaCod) & ") " _
     & IIf(pbVerDoc, "LEFT JOIN MovDoc Doc ON Doc.nMovNro = M.nMovNro LEFT JOIN " & vsBaseComunes & "Documento Docum ON Docum.nDocTpo = Doc.nDocTpo ", "") _
     & "WHERE  M.nMovEstado = '" & gMovEstContabMovContable & "' and not M.nMovFlag in ('" & gMovFlagEliminado & "','" & gMovFlagModificado & "','" & gMovFlagExtornado & "') and not M.cOpecod Like '70185%' and a.cCtaContCod Like '" & lsCtaCod & "%' and " _
     & "       substring(M.cmovnro,1,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' " _
     & "GROUP BY Left(a.cCtaContCod," & Len(psCtaCod) & "),a.cCtaContcod, c.cCtaContDesc, M.cMovNro, M.nMovNro, " & IIf(pbVerDoc, " Docum.cDocAbrev, Doc.cDocNro, Doc.dDocFecha, ", "") & " M.cMovDesc " _
     & IIf(pnImporte > 0, "HAVING SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END) " & psFiltro & pnImporte & " " _
     & "or SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END) " & psFiltro & pnImporte, "") _
     & "ORDER BY  M.cMovNro, a.cCtaContcod"


Dim oConec As New DConecta
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

Dim oConec As New DConecta
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

Dim oConec As New DConecta
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
Dim oConec As New DConecta
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

Public Function GetAsientoConsulta(psBusCond As String, psOpeCond As String, psAgeCond As String, psDocCond As String, psCtaCond As String, psMovCond As String, psFecCond As String, psMesCerrado As String) As ADODB.Recordset
On Error GoTo GetAsientoConsultaErr
Dim oConec As New DConecta

'*** PEAC 20120503 - se aumento el campo copecod
lsSql = "SELECT m.cMovNro, m.nMovNro, a.nMovItem, a.cCtaContCod, b.cCtaContDesc, " _
     & "       SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END) as nDebe, " _
     & "       SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END) as nHaber, " _
     & "       SUM(CASE WHEN me.nMovMeImporte > 0 THEN me.nMovMeImporte END) as nDebeME, " _
     & "       SUM(CASE WHEN me.nMovMeImporte < 0 THEN me.nMovMeImporte * -1 END) nHaberME, M.nMovFlag, M.nMovEstado " & IIf(psMesCerrado = "", "", ",case when left(cmovnro,8)<='" & psMesCerrado & "' then 'SI' else 'NO' end MCerrado ") & " " _
     & "       , max(m.cOpeCod) cOpeCod " _
     & " FROM  MovCta a      JOIN Mov M ON M.nMovNro = a.nMovNro " _
     & "                LEFT JOIN MovMe Me ON Me.nMovNro = a.nMovNro and Me.nMovItem = a.nMovItem " _
     & "                LEFT JOIN " & vsBaseComunes & "CtaCont b ON a.cCtaContCod = b.cCtaContCod " _
     & "                     JOIN " & vsBaseComunes & "OpeTpo Op ON Op.cOpeCod = M.cOpeCod " _
     & IIf(psDocCond = "", "", " JOIN MovDoc MD ON MD.nMovNro = M.nMovNro ") _
     & "WHERE " & psFecCond & " " & psBusCond & " " & psCtaCond & " " & psMovCond & " " & psAgeCond & " " & psOpeCond & " " & psDocCond & " " _
     & "GROUP BY cMovNro, m.nMovNro, a.nMovItem, a.cCtaContCod, b.cCtaContDesc, M.nMovFlag, M.nMovEstado " _
     & "ORDER BY cMovNro, a.nMovItem"
     
oConec.AbreConexion
Set GetAsientoConsulta = oConec.CargaRecordSet(lsSql)
oConec.CierraConexion
Set oConec = Nothing

Exit Function
GetAsientoConsultaErr:
   Call RaiseError(MyUnhandledError, "NContAsientos:GetAsientoConsulta Method")
End Function


Public Function GetMovimientoConsulta(psBusCond As String, psOpeCond As String, psAgeCond As String, psDocCond As String, psCtaCond As String, psMovCond As String, psFecCond As String, Optional psMesCerrado As String = "") As ADODB.Recordset
On Error GoTo GetMovimientoConsultaErr
Dim oConec As New DConecta

lsSql = "SELECT m.cMovNro, m.nMovEstado, m.cOpeCod, ISNULL(m1.nMovMonto,0), m.cMovDesc, m.nMovFlag, m.nMovNro " & IIf(psMesCerrado = "", "", ",case when left(cmovnro,8)<='" & psMesCerrado & "' then 'SI' else 'NO' end MCerrado ") & "  " _
     & " FROM  Mov m JOIN OpeTpo ot ON ot.cOpeCod = m.cOpecod LEFT JOIN MovCont M1 ON M.nMovNro = m1.nMovNro " _
     & IIf(psCtaCond = "", "", " JOIN MovCta Mc ON Mc.nMovNro = M.nMovNro ") _
     & IIf(psDocCond = "", "", " JOIN MovDoc MD ON MD.nMovNro = M.nMovNro ") _
     & "WHERE " & psFecCond & " " & psBusCond & " " & psCtaCond & " " & psMovCond & " " & psAgeCond & " " & psOpeCond & " " & psDocCond & " " _
     & "ORDER BY m.cMovNro"
     
     
oConec.AbreConexion
Set GetMovimientoConsulta = oConec.CargaRecordSet(lsSql)
oConec.CierraConexion
Set oConec = Nothing

Exit Function
GetMovimientoConsultaErr:
   Call RaiseError(MyUnhandledError, "NContAsientos:GetMovimientoConsulta Method")
End Function

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

Dim oConec As New DConecta
oConec.AbreConexion
Set GetAsientoAjusteTipoCambio = oConec.CargaRecordSet(lsSql)
oConec.CierraConexion
Set oConec = Nothing
End Function
'YIHU RS015-2015 20150406****************************
Public Function GetAsientoContablexCuenta(codcta As String, fecha As Date) As ADODB.Recordset
Dim oConect As New DConecta
    
lsSql = "stp_sel_ObtieneAsientoContablexCuenta '" & codcta & "','" & Format(fecha, "yyyyMMdd") & "'"
oConect.AbreConexion
Set GetAsientoContablexCuenta = oConect.CargaRecordSet(lsSql)
oConect.CierraConexion
Set oConect = Nothing
End Function
'End YIHU********************************************



