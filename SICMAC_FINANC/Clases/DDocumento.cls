VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DDocumento"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A80C4CD03B9"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sCentralCom As String
Public Function CargaDocumento(Optional pnDocTpo As Long = -1, Optional psDocAbrev As String = "", Optional LockType As LockTypeEnum = adLockReadOnly, Optional pnNivel As Integer = 0) As Recordset
Dim sCond As String
   On Error GoTo CargaDocumentoErr
   If pnDocTpo <> -1 Then
      sCond = " WHERE nDocTpo = " & pnDocTpo & " "
   End If
   If psDocAbrev <> "" Then
      sCond = sCond & IIf(sCond = "", "WHERE", "AND") & " cDocAbrev = '" & psDocAbrev & "'"
   End If
   psSql = "SELECT nDocTpo, cDocDesc, " & IIf(pnNivel > 0, pnNivel & " Nivel,", "") & " cDocAbrev FROM " & sCentralCom & "Documento " & sCond
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      Set CargaDocumento = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaDocumentoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:CargaDocumento Method")
End Function

Public Sub InsertaDocumento(pnDocTpo As Long, psDocDesc As String, psDocAbrev As String, psMovNro As String)
   On Error GoTo InsertaDocumentoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "Documento (nDocTpo, cDocDesc, cDocAbrev, cUltimaActualizacion) " _
         & "VALUES (" & pnDocTpo & ", '" & psDocDesc & "','" & psDocAbrev & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaDocumentoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:InsertaDocumento Method")
End Sub

Public Sub ActualizaDocumento(pnDocTpo As Long, psDocDesc As String, psDocAbrev As String, psMovNro As String)
   On Error GoTo ActualizaDocumentoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "UPDATE Documento Set cDocDesc = '" & psDocDesc & "', cDocAbrev = '" & psDocAbrev & "', cUltimaActualizacion = '" & psMovNro & "'  WHERE nDocTpo = " & pnDocTpo & " "
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizaDocumentoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:ActualizaDocumento Method")
End Sub

Public Sub AsignaDocImpuesto(pnDocTpo As Long, psCtaContCod As String, psDocImpDH As String, psDocImpOpc As String, psMovNro As String)
   On Error GoTo AsignaDocImpuestoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "DocImpuesto (nDocTpo, cCtaContCod, cDocImpDH, cDocImpOpc, cUltimaActualizacion) " _
         & "VALUES (" & pnDocTpo & ", '" & psCtaContCod & "','" & psDocImpDH & "','" & psDocImpOpc & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
AsignaDocImpuestoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:AsignaDocImpuesto Method")
End Sub

Public Sub DesasignaDocImpuesto(pnDocTpo As Long, psCtaContCod As String)
   On Error GoTo DesasignaDocImpuestoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "DocImpuesto WHERE nDocTpo = " & pnDocTpo & " and cCtaContCod = '" & psCtaContCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
DesasignaDocImpuestoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:DesasignaDocImpuesto Method")
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
Public Function GetValidaDocProv(ByVal psPersCod As String, ByVal pnDocTpo As Long, ByVal psDocNro As String) As Boolean
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
If Not psPersCod = "" Then
    sql = " SELECT  m.cMovNro " _
         & " FROM   MovGasto a JOIN MovDoc b ON b.nMovNro = a.nMovNro JOIN Mov m ON m.nMovNro = a.nMovNro " _
         & " WHERE  M.nMovFlag not in ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "', '" & gMovFlagModificado & "')  and a.cPersCod = '" & psPersCod & "' and b.nDocTpo = " & pnDocTpo & " and b.cDocNro = '" & psDocNro & "'"
Else
    sql = "SELECT nMovNro FROM MovDoc WHERE nDocTpo = " & pnDocTpo & " and cDocNro = '" & psDocNro & "' "
End If
GetValidaDocProv = False
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetValidaDocProv = True
End If
rs.Close
Set rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function VerificaDoc(ByVal pnDocTpo As TpoDoc, ByVal psNroDoc As String, Optional ByVal psPersCod As String = "") As Boolean
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim sql As String
Set rs = New ADODB.Recordset
Set oConect = New DConecta
VerificaDoc = False
If oConect.AbreConexion = False Then Exit Function
    If pnDocTpo = TpoDocOrdenPago Then
        sSql = "SELECT a.nMovNro FROM  movdoc a " _
                & "  Join MOV M On a.nMovNro = M.nMovNro " _
                & "   WHERE a.nDocTpo = " & pnDocTpo & " and a.cDocNro = '" & psNroDoc & "' " _
                & " And Left(M.cOpeCod,2) Like '4[0123]%'"
         'sSQL = "SELECT a.nMovNro FROM  movdoc a "
         '& "   WHERE a.nDocTpo = " & pnDocTpo & " and a.cDocNro = '" & psNroDoc & "' "
    Else
        sSql = "SELECT a.nMovNro FROM  movdoc a JOIN MOVOBJIF b on a.nMovNro = b.nMovNro  " _
              & "WHERE a.nDocTpo = " & pnDocTpo & " and a.cDocNro = '" & psNroDoc & "' " _
              & "  and b.cPersCod ='" & psPersCod & "' "
    End If
    Set rs = oConect.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
       VerificaDoc = True
    End If
    rs.Close: Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function VerificaCheque(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF) As Boolean
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim sql As String
Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
    
    sSql = " Select  cNroDoc " _
        & " From    Docrec " _
        & " Where   nTpodoc = " & pnTpoDoc & " and cNroDoc = '" & psNroDoc & "' " _
        & "         and cPersCod = '" & psPersCod & "' and cIFtpo = '" & Format(psIFTpo, "00") & "'"
        
    VerificaCheque = False
    Set rs = oConect.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
       VerificaCheque = True
    End If
    rs.Close: Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
End Function

'************************
'** NUEVAS FUNCIONES
'************************
Public Sub EliminaDocumento(pnDocTpo As Long)
Dim lTrans As Boolean
Dim lsError As String
   On Error GoTo EliminaDocumentoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      lTrans = True
      dbConec.BeginTrans
      psSql = "DELETE DocImpuesto WHERE nDocTpo = " & pnDocTpo & " "
      dbConec.Ejecutar psSql
      psSql = "DELETE Documento WHERE nDocTpo = " & pnDocTpo & " "
      dbConec.Ejecutar psSql
      dbConec.CommitTrans
      lTrans = False
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaDocumentoErr:
   lsError = Err.Description
   If lTrans Then
      dbConec.RollbackTrans
   End If
   Err.Raise Err.Number + MyUnhandledError, "DDocumento:EliminaDocumento", lsError
End Sub

Public Function CargaDocImpuesto(pnDocTpo As Long, Optional psCtaCod As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaDocImpuestoErr
   '***Modificado por ELRO el 20130607, según TI-ERS064-2013****
   psSql = "exec stp_sel_ERS0642013_CargaDocImpuesto " & pnDocTpo & ", '" & psCtaCod & "'"
   'psSql = "SELECT d.cCtaContCod, c.cCtaContDesc, " & _
   '    "       CASE WHEN cDocImpDH ='D' THEN 'DEBE' " & _
   '    "            WHEN cDocImpDH ='H' THEN 'HABER' " & _
   '    "       END as Clase , " & _
   '    "       CASE WHEN cDocImpOpc ='1' THEN 'Obligatorio' " & _
   '    "            WHEN cDocImpOpc ='2' THEN 'Opcional' " & _
   '    "       END as Caracter , " _
   '    & "       d.cDocImpDH, d.cDocImpOpc, I.cImpAbrev,  I.nImpTasa, I.cImpDestino, nCalculo " _
   '    & " FROM  " & sCentralCom & "CtaCont c " _
   '    & "      JOIN " & sCentralCom & "DocImpuesto d ON d.cCtaContCod = c.cCtaContCod " _
   '    & "      JOIN " & sCentralCom & "Impuesto I ON I.cCtaContCod = D.cCtaContCod " _
   '    & IIf(pnDocTpo <> -1, "WHERE d.nDocTpo= " & pnDocTpo, "") & IIf(psCtaCod <> "", IIf(pnDocTpo <> -1, " WHERE ", " AND ") & "c.cCtaContCod = '" & psCtaCod & "'", "") & " Order By nCalculo"
   '***Fin Modificado por ELRO el 20130607, según TI-ERS064-2013
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      Set CargaDocImpuesto = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaDocImpuestoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:CargaDocImpuesto Method")
End Function

'*** PEAC 20100921
Public Function CargaDatosPorcenPoli(ByVal psCodAge As String, ByVal pnTipoSeguro As Integer) As Recordset
   On Error GoTo CargaDatosPorcenPoliErr

    psSql = " exec stp_sel_ObtieneDatosPoliSeguroPatri '" & psCodAge & "'," & pnTipoSeguro

   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      Set CargaDatosPorcenPoli = dbConec.CargaRecordSet(psSql, 1)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaDatosPorcenPoliErr:
   Call RaiseError(MyUnhandledError, "DDocumento:CargaDatosPorcenPoli Method")
End Function

'***Agregado por ELRO el 20111109, segun Acta 323-2011/TI-D
Public Function registrarPagoAdelantadoCab(ByVal psDesPagAde As String, _
                                           ByVal psCtaContCod As String, _
                                           ByVal pnRubro As Integer, _
                                           ByVal pnNumMes As Integer, _
                                           ByVal pdFecIni As String, _
                                           ByVal pdFecVen As String, _
                                           ByVal pnImporte As Currency, _
                                           ByVal pcAgeCod As String, _
                                           ByVal psMov As String) As Boolean

On Error GoTo registrarPagoAdelantadoCabErr
    Dim sSql As String
    Dim oDConecta As New DConecta
    
    registrarPagoAdelantadoCab = False

    sSql = " exec stp_ins_RegiatrarPagoAdelantadoCab '" & psDesPagAde & "', '" & psCtaContCod & "', " & pnRubro & ", '" & pnNumMes & "', '" & pdFecIni & "', '" & pdFecVen & "', " & pnImporte & ", '" & pcAgeCod & "', '" & psMov & "'"

    If oDConecta.AbreConexion() Then
        Call oDConecta.Ejecutar(sSql)
        registrarPagoAdelantadoCab = True
        oDConecta.CierraConexion
    End If
    
    
    Set oDConecta = Nothing
    Exit Function

registrarPagoAdelantadoCabErr:
Call RaiseError(MyUnhandledError, "DDocumento:registrarPagoAdelantadoCab Method")
End Function

Public Function recuperarPagoAdelantadoCab(ByVal pIdPagAdeCab As Integer) As ADODB.Recordset

On Error GoTo recuperarPagoAdelantadoCabErr
    Dim sSql As String
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta

    sSql = " exec stp_sel_RecuperarPagoAdelantadoCab " & pIdPagAdeCab & " "

    If oDConecta.AbreConexion() Then
        Set recuperarPagoAdelantadoCab = oDConecta.CargaRecordSet(sSql)
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
recuperarPagoAdelantadoCabErr:
Call RaiseError(MyUnhandledError, "DDocumento:recuperarPagoAdelantadoCab Method")
End Function

Public Function actualizarPagoAdelantadoCab(ByVal pIdPagAdeCab As Integer, _
                                            ByVal psDesPagAde As String, _
                                            ByVal psCtaContCod As String, _
                                            ByVal pnRubro As Integer, _
                                            ByVal pnNumMes As Integer, _
                                            ByVal pdFecIni As String, _
                                            ByVal pdFecVen As String, _
                                            ByVal pnImporte As Currency, _
                                            ByVal pcAgeCod As String, _
                                            ByVal psMov As String) As Boolean

On Error GoTo actualizarPagoAdelantadoCabErr
    Dim sSql As String
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta
    
    actualizarPagoAdelantadoCab = False

    sSql = " exec stp_upd_ActualizarPagoAdelantadoCab " & pIdPagAdeCab & ", '" & psDesPagAde & "', '" & psCtaContCod & "', " & pnRubro & "," & pnNumMes & ", '" & pdFecIni & "', '" & pdFecVen & "', " & pnImporte & ", '" & pcAgeCod & "', '" & psMov & "'"

    If oDConecta.AbreConexion() Then
        oDConecta.Ejecutar (sSql)
        actualizarPagoAdelantadoCab = True
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
actualizarPagoAdelantadoCabErr:
Call RaiseError(MyUnhandledError, "DDocumento:actualizarPagoAdelantadoCab Method")
End Function

Public Function listarRubrosPagosAdelantados() As ADODB.Recordset

On Error GoTo listarRubrosPagosAdelantadosErr
    Dim sSql As String
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta

    sSql = "exec stp_sel_ListarRubrosPagosAdelantados"

    If oDConecta.AbreConexion() Then
        Set listarRubrosPagosAdelantados = oDConecta.CargaRecordSet(sSql)
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
listarRubrosPagosAdelantadosErr:
Call RaiseError(MyUnhandledError, "DDocumento:listarRubrosPagosAdelantados Method")
End Function

Public Function buscarPagosAdelantados(ByVal pnRubro As Integer, ByVal psMoneda As String, Optional ByVal bTodo As Boolean = False) As ADODB.Recordset

On Error GoTo buscarPagosAdelantadosErr
    Dim sSql As String
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta
    
    If bTodo Then
        sSql = "exec stp_sel_BuscarPagosAdelantados " & pnRubro & ", '" & psMoneda & "', " & 1 & ""
    Else
        sSql = "exec stp_sel_BuscarPagosAdelantados " & pnRubro & ", '" & psMoneda & "', " & 0 & ""
    End If
    
    
    If oDConecta.AbreConexion() Then
        Set buscarPagosAdelantados = oDConecta.CargaRecordSet(sSql)
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
buscarPagosAdelantadosErr:
Call RaiseError(MyUnhandledError, "DDocumento:buscarPagosAdelantados Method")

End Function

Public Function listarAgenciasPagosAdelantados() As ADODB.Recordset

On Error GoTo listarAgenciasPagosAdelantadosErr
    Dim sSql As String
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta

    sSql = "exec stp_sel_ListarAgenciasPagosAdelantados"

    If oDConecta.AbreConexion() Then
        Set listarAgenciasPagosAdelantados = oDConecta.CargaRecordSet(sSql)
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
listarAgenciasPagosAdelantadosErr:
Call RaiseError(MyUnhandledError, "DDocumento:listarAgenciasPagosAdelantados Method")

End Function

Public Function NumeroAmortizacionPagoAdelantado(ByVal pIdPagAdeCab As Integer, _
                                                 Optional ByVal pnMes As Integer = 0, _
                                                 Optional ByVal pnAnio As Integer = 0, _
                                                 Optional ByVal bTodo As Boolean = False) As Integer

On Error GoTo NumeroAmortizacionPagoAdelantadoErr
    Dim sSql As String
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta
    Dim rsNumeroAmortizacion As ADODB.Recordset
    Set rsNumeroAmortizacion = New ADODB.Recordset
    
    NumeroAmortizacionPagoAdelantado = 0
    
    If bTodo Then
        sSql = "exec stp_sel_NumeroAmortizacionPagoAdelantado " & pIdPagAdeCab & ", " & pnMes & ", " & pnAnio & ", " & 1 & " "
    Else
        sSql = "exec stp_sel_NumeroAmortizacionPagoAdelantado " & pIdPagAdeCab & ", " & pnMes & ", " & pnAnio & ", " & 0 & " "
    End If

    If oDConecta.AbreConexion() Then
        Set rsNumeroAmortizacion = oDConecta.CargaRecordSet(sSql)
        If Not rsNumeroAmortizacion.BOF And Not rsNumeroAmortizacion.EOF Then
            NumeroAmortizacionPagoAdelantado = rsNumeroAmortizacion!nRegistro
        End If
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
NumeroAmortizacionPagoAdelantadoErr:
Call RaiseError(MyUnhandledError, "DDocumento:NumeroAmortizacionPagoAdelantado Method")

End Function


Public Function generarAsientoAmortizacionPagoAdelantado(ByVal psMovNro As String, _
                                                         ByVal psOpeCod As String, _
                                                         ByVal psMovDesc As String, _
                                                         ByVal pcCtaConDebe As String, _
                                                         ByVal pnMes As Integer, _
                                                         ByVal pnAnio As Integer, _
                                                         ByVal pnMoneda As Integer, _
                                                         ByVal rsListaGenerarAsiento As ADODB.Recordset, _
                                                         ByRef psError As String) As Long
            

Dim oDMov As DMov
Dim lnMovItem  As Long
Dim lnMovNro   As Long
Dim lbTrans    As Boolean
Dim lnCuotasAmortizadas, lnCuotaAmortizada, lnConfirmarInsertaMovCtaDebe, lnConfirmarInsertaMovCtaHaber  As Integer
Dim lnAmortizacion, lnTotalAmortizacion As Currency
Dim lbConfirmarAmortizacion As Boolean
Dim lsMsgErr, lcCtaConDebe  As String
Dim larrCodAge() As String
Dim I As Integer
Dim nAmortizacionXAgencia As Currency
Dim nAmortizacionUltimaAgencia As Currency
Dim nAgencias As Integer



On Error GoTo generarAsientoAmortizacionPagoAdelantadoErr

Set oDMov = New DMov

generarAsientoAmortizacionPagoAdelantado = 0

oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable
lnMovNro = oDMov.GetnMovNro(psMovNro)

'Guardarmos las cuentas del debe y haber
lnMovItem = 0
lnAmortizacion = 0
lnTotalAmortizacion = 0
lnConfirmarInsertaMovCtaDebe = 1
lnConfirmarInsertaMovCtaHaber = 1
lbConfirmarAmortizacion = False

If Not rsListaGenerarAsiento Is Nothing Then
    If rsListaGenerarAsiento.State = adStateOpen Then
        If Not rsListaGenerarAsiento.EOF And Not rsListaGenerarAsiento.BOF Then
            Do While Not rsListaGenerarAsiento.EOF
                lnCuotasAmortizadas = NumeroAmortizacionPagoAdelantado(rsListaGenerarAsiento!IdPagAdeCab, _
                                                                      , _
                                                                       , _
                                                                       True)
                lnCuotaAmortizada = NumeroAmortizacionPagoAdelantado(rsListaGenerarAsiento!IdPagAdeCab, _
                                                                       pnMes, _
                                                                       pnAnio, _
                                                                       False)
                If lnCuotaAmortizada = 0 Then
                    If rsListaGenerarAsiento!nNumMes = lnCuotasAmortizadas + 1 Then
                        lnAmortizacion = rsListaGenerarAsiento!nImporte - (Round((rsListaGenerarAsiento!nImporte / rsListaGenerarAsiento!nNumMes), 2) * (rsListaGenerarAsiento!nNumMes - 1))
                    Else
                        lnAmortizacion = Round(rsListaGenerarAsiento!nImporte / rsListaGenerarAsiento!nNumMes, 2)
                    End If
                    'Guardamos la cuenta en el debe
                    lnMovItem = lnMovItem + 1
                    lcCtaConDebe = ""
                    larrCodAge = Split(rsListaGenerarAsiento!cAgeCod, ",")
                    nAgencias = CInt(UBound(larrCodAge))
                    
                    If nAgencias >= 1 Then
                        nAmortizacionXAgencia = Round(lnAmortizacion / (nAgencias + 1), 2)
                        nAmortizacionUltimaAgencia = lnAmortizacion - (nAmortizacionXAgencia * nAgencias)
                        For I = 0 To nAgencias
                            If larrCodAge(I) <> "" Then
                                lcCtaConDebe = pcCtaConDebe & larrCodAge(I)
                            Else
                                lcCtaConDebe = pcCtaConDebe
                            End If
                            
                            If nAgencias > I Then
                                lnConfirmarInsertaMovCtaDebe = oDMov.InsertaMovCta(lnMovNro, lnMovItem, lcCtaConDebe, nAmortizacionXAgencia)
                                If lnConfirmarInsertaMovCtaDebe > 0 Then
                                    psError = lcCtaConDebe
                                    generarAsientoAmortizacionPagoAdelantado = -1
                                    oDMov.RollbackTrans
                                    Exit Function
                                End If
                            Else
                                lnConfirmarInsertaMovCtaDebe = oDMov.InsertaMovCta(lnMovNro, lnMovItem, lcCtaConDebe, nAmortizacionUltimaAgencia)
                                If lnConfirmarInsertaMovCtaDebe > 0 Then
                                    psError = lcCtaConDebe
                                    generarAsientoAmortizacionPagoAdelantado = -1
                                    oDMov.RollbackTrans
                                    Exit Function
                                End If
                            End If
                            lnMovItem = lnMovItem + 1
                        Next I
                    Else
                        If Trim(CStr(rsListaGenerarAsiento!cAgeCod)) <> "" Then
                          lcCtaConDebe = pcCtaConDebe & CStr(rsListaGenerarAsiento!cAgeCod)
                        Else
                            lcCtaConDebe = pcCtaConDebe
                        End If
                        lnConfirmarInsertaMovCtaDebe = oDMov.InsertaMovCta(lnMovNro, lnMovItem, lcCtaConDebe, lnAmortizacion)
                        If lnConfirmarInsertaMovCtaDebe > 0 Then
                            psError = lcCtaConDebe
                            generarAsientoAmortizacionPagoAdelantado = -1
                            oDMov.RollbackTrans
                            Exit Function
                        End If
                    End If
                    'Guardamos la cuenta en el Haber
                    If lnConfirmarInsertaMovCtaDebe = 0 Then
                        lnMovItem = lnMovItem + 1
                        lnConfirmarInsertaMovCtaHaber = oDMov.InsertaMovCta(lnMovNro, lnMovItem, rsListaGenerarAsiento!cCtaContCod, lnAmortizacion * -1)
                        
                        'Registra la amortizacion en la tabla PagoAdelantadoDet
                        If lnConfirmarInsertaMovCtaHaber = 0 Then
                            lbConfirmarAmortizacion = registrarAmortizacionPagoAdelantadoCab(rsListaGenerarAsiento!IdPagAdeCab, _
                                                                                             lnAmortizacion, _
                                                                                             pnMes, _
                                                                                             pnAnio, _
                                                                                             psMovNro)
                            
                            'Sumando las amortizaciones
                            lnTotalAmortizacion = lnTotalAmortizacion + lnAmortizacion
                            If lbConfirmarAmortizacion = False Then
                                psError = "No se registro los datos de la amortización del Pago Adelantado"
                                generarAsientoAmortizacionPagoAdelantado = -3
                                oDMov.RollbackTrans
                                Exit Function
                            End If
                            
                        Else
                            psError = rsListaGenerarAsiento!cCtaContCod
                            generarAsientoAmortizacionPagoAdelantado = -2
                            oDMov.RollbackTrans
                            Exit Function
                        End If
                    Else
                        psError = lcCtaConDebe
                        generarAsientoAmortizacionPagoAdelantado = -1
                        oDMov.RollbackTrans
                        Exit Function
                    End If
                End If
                
            rsListaGenerarAsiento.MoveNext
            Loop
        End If
    End If
End If

If lnCuotaAmortizada = 0 Then
    oDMov.InsertaMovCont lnMovNro, lnTotalAmortizacion, "0", "0"
                                       
    If pnMoneda = gMonedaExtranjera Then
        oDMov.GeneraMovME lnMovNro, psMovNro
    End If
    
    oDMov.ActualizaSaldoMovimiento psMovNro, "+"
End If

oDMov.CommitTrans
lbTrans = False
If lnConfirmarInsertaMovCtaDebe = 0 And _
   lnConfirmarInsertaMovCtaHaber = 0 And _
   lbConfirmarAmortizacion Then
    generarAsientoAmortizacionPagoAdelantado = lnMovNro
End If
Set oDMov = Nothing
lnMovItem = 0
lnMovNro = 0
lnCuotasAmortizadas = 0
lnCuotaAmortizada = 0
lnConfirmarInsertaMovCtaDebe = 0
lnConfirmarInsertaMovCtaHaber = 0
lnAmortizacion = 0#
lnTotalAmortizacion = 0#
lbConfirmarAmortizacion = False
lcCtaConDebe = ""
I = 0
nAmortizacionXAgencia = 0#
nAmortizacionUltimaAgencia = 0#
nAgencias = 0
Exit Function
generarAsientoAmortizacionPagoAdelantadoErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "generarAsientoAmortizacionPagoAdelantado", lsMsgErr & " Error al grabar el asiento de la amortización del Pago Adelanatdo " & psMovDesc
End Function

Public Function registrarAmortizacionPagoAdelantadoCab(ByVal pIdPagAdeCab As Integer, _
                                                       ByVal pnAmortizacion As Currency, _
                                                       ByVal pnMes As Integer, _
                                                       ByVal pnAnio As Integer, _
                                                       ByVal psMov As String) As Boolean

On Error GoTo registrarAmortizacionPagoAdelantadoCabErr
    Dim sSql As String
    Dim oDConecta As New DConecta
    
    registrarAmortizacionPagoAdelantadoCab = False

    sSql = " exec stp_ins_RegistrarAmortizacionPagoAdelantado " & pIdPagAdeCab & ", " & pnAmortizacion & ", " & pnMes & ", " & pnAnio & ", '" & psMov & "'"

    If oDConecta.AbreConexion() Then
        Call oDConecta.Ejecutar(sSql)
        registrarAmortizacionPagoAdelantadoCab = True
        oDConecta.CierraConexion
    End If
    
    
    Set oDConecta = Nothing
    Exit Function

registrarAmortizacionPagoAdelantadoCabErr:
Call RaiseError(MyUnhandledError, "DDocumento:registrarAmortizacionPagoAdelantadoCab Method")
End Function

Public Function listarAmortizacionesPagoAdelantado(Optional ByVal pIdPagAdeCab As Integer = 0, _
                                                   Optional ByVal pcMov As String = "") As ADODB.Recordset
On Error GoTo listarAmortizacionesPagoAdelantadoErr
    Dim sSql As String
    Dim oDConecta As New DConecta
    Set oDConecta = New DConecta

    sSql = "exec stp_sel_ListarAmortizacionesPagoAdelantado  " & pIdPagAdeCab & ", '" & pcMov & "'"

    If oDConecta.AbreConexion() Then
        Set listarAmortizacionesPagoAdelantado = oDConecta.CargaRecordSet(sSql)
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
listarAmortizacionesPagoAdelantadoErr:
Call RaiseError(MyUnhandledError, "DDocumento:listarAmortizacionesPagoAdelantado Method")
End Function


Public Function eliminarPagoAdelantadoCab(ByVal pIdPagAdeCab As Integer) As Boolean
On Error GoTo eliminarPagoAdelantadoCabErr
    Dim sSql As String
    Dim oDConecta As New DConecta
    
    eliminarPagoAdelantadoCab = False

    sSql = " exec stp_del_EliminarPagoAdelantadoCab " & pIdPagAdeCab

    If oDConecta.AbreConexion() Then
        Call oDConecta.Ejecutar(sSql)
        eliminarPagoAdelantadoCab = True
        oDConecta.CierraConexion
    End If
    
    
    Set oDConecta = Nothing
    Exit Function

eliminarPagoAdelantadoCabErr:
Call RaiseError(MyUnhandledError, "DDocumento:eliminarPagoAdelantadoCab Method")

End Function

 Public Function imprimirAmortizacionesPagosAdelantados(ByVal pcMov As String, _
                                                        ByVal psNomCmac As String, _
                                                        ByVal psFecSis As String, _
                                                        ByVal psMoneda As String) As String
    Dim sCadImp As String
    Dim rsLista As ADODB.Recordset
    Set rsLista = New Recordset
 
    Set rsLista = listarAmortizacionesPagoAdelantado(, pcMov)
    
    If Not rsLista.BOF And Not rsLista.EOF Then
    
    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & psNomCmac & Space(70) & psFecSis & Chr(10)
    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & Space(50) & "AMORTIZACIONES DE PAGOS ADELANTADOS" & Chr(10)
    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & "MOVIMIENTO: " & pcMov & Chr(10)
    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & "MONEDA: " & psMoneda & Chr(10)
    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & String(141, "-") & Chr(10)
    sCadImp = sCadImp & ImpreFormat("PAGO ADELANTADO", 64) & ImpreFormat("IMPORTE", 10) & ImpreFormat("AMO. ACU.", 10) & ImpreFormat("MES AMO.", 13) & ImpreFormat("AMO. MES", 10) & ImpreFormat("TOT. AMO.", 15) & ImpreFormat("SALDO", 9) & Chr(10)
    sCadImp = sCadImp & String(141, "-") & Chr(10)
    Do While Not rsLista.EOF
        sCadImp = sCadImp & ImpreFormat(rsLista!cDesPagAde, 60) & ImpreFormat(rsLista!nImporte, 10) & ImpreFormat(rsLista!nAmoAcu, 11) & ImpreFormat(rsLista!cMes, 11) & ImpreFormat(rsLista!nAmortizacion, 10) & ImpreFormat(rsLista!nTotAmt, 10) & ImpreFormat(rsLista!nSaldo, 10) & Chr(10)
        rsLista.MoveNext
    Loop
    
    imprimirAmortizacionesPagosAdelantados = sCadImp
    
    Else
        sCadImp = ""
        Exit Function
    End If
 
 End Function

Public Function recuperarRubroPagosAdelantados(Optional ByVal pbTodo As Boolean = True, _
                                               Optional ByVal pnRubro As Integer = 0) As ADODB.Recordset


On Error GoTo recuperarRubroPagosAdelantadosErr
    Dim sSql As String
    Dim oDConecta As New DConecta
    Set oDConecta = New DConecta

    If pbTodo Then
        sSql = "exec stp_sel_RecuperarRubroPagosAdelantados  " & 1 & ", " & pnRubro & ""
    Else
        sSql = "exec stp_sel_RecuperarRubroPagosAdelantados  " & 0 & ", " & pnRubro & ""
    End If
    

    If oDConecta.AbreConexion() Then
        Set recuperarRubroPagosAdelantados = oDConecta.CargaRecordSet(sSql)
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
recuperarRubroPagosAdelantadosErr:
Call RaiseError(MyUnhandledError, "DDocumento:recuperarRubroPagosAdelantados Method")

End Function

Public Function registrarRubroPagoAdelantado(ByVal pnConsValor As Integer, _
                                             ByVal pcForCtaCon As String, _
                                             ByVal pcMov As String) As Boolean

On Error GoTo registrarRubroPagoAdelantadoErr
    Dim sSql As String
    Dim oDConecta As New DConecta
    
    registrarRubroPagoAdelantado = False

    sSql = " exec stp_ins_RegistrarRubroPagoAdelantadoCab " & pnConsValor & ", '" & pcForCtaCon & "', '" & pcMov & "'"

    If oDConecta.AbreConexion() Then
        Call oDConecta.Ejecutar(sSql)
        registrarRubroPagoAdelantado = True
        oDConecta.CierraConexion
    End If
    
    
    Set oDConecta = Nothing
    Exit Function

registrarRubroPagoAdelantadoErr:
Call RaiseError(MyUnhandledError, "DDocumento:registrarRubroPagoAdelantado Method")
End Function

Public Function modificarRubroPagoAdelantado(ByVal pIdRubroPagoAdelantadoCab As Integer, _
                                             ByVal pnConsValor As Integer, _
                                             ByVal pcForCtaCon As String, _
                                             ByVal pcMov As String) As Boolean

On Error GoTo modificarRubroPagoAdelantadoErr
    Dim sSql As String
    Dim oDConecta As New DConecta
    
    modificarRubroPagoAdelantado = False

    sSql = " exec stp_upd_ModificarRubroPagoAdelantadoCab " & pIdRubroPagoAdelantadoCab & " ," & pnConsValor & ", '" & pcForCtaCon & "', '" & pcMov & "'"

    If oDConecta.AbreConexion() Then
        Call oDConecta.Ejecutar(sSql)
        modificarRubroPagoAdelantado = True
        oDConecta.CierraConexion
    End If
        
    Set oDConecta = Nothing
    Exit Function

modificarRubroPagoAdelantadoErr:
Call RaiseError(MyUnhandledError, "DDocumento:modificarRubroPagoAdelantado Method")
End Function

Public Function eliminarRubroPagoAdelantado(ByVal pIdRubroPagoAdelantadoCab As Integer) As Boolean

On Error GoTo eliminarRubroPagoAdelantadoErr
    Dim sSql As String
    Dim oDConecta As New DConecta
    
    eliminarRubroPagoAdelantado = False

    sSql = " exec stp_del_EliminarRubroPagoAdelantadoCab " & pIdRubroPagoAdelantadoCab & ""

    If oDConecta.AbreConexion() Then
        Call oDConecta.Ejecutar(sSql)
        eliminarRubroPagoAdelantado = True
        oDConecta.CierraConexion
    End If
    
    
    Set oDConecta = Nothing
    Exit Function

eliminarRubroPagoAdelantadoErr:
Call RaiseError(MyUnhandledError, "DDocumento:eliminarRubroPagoAdelantado Method")
End Function

Public Function recuperarCodigoRubro(Optional ByVal pcConsDescripcion As String = "") As Integer

On Error GoTo recuperarCodigoRubroErr
    Dim sSql As String
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta
    Dim rsCodigoRubro As ADODB.Recordset
    Set rsCodigoRubro = New ADODB.Recordset
    
    recuperarCodigoRubro = 0
    
    sSql = "exec stp_sel_RecuperarCodigoConstante '" & pcConsDescripcion & "'"
    If oDConecta.AbreConexion() Then
        Set rsCodigoRubro = oDConecta.CargaRecordSet(sSql)
        If Not rsCodigoRubro.BOF And Not rsCodigoRubro.EOF Then
            recuperarCodigoRubro = rsCodigoRubro!Codigo
        End If
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
recuperarCodigoRubroErr:
Call RaiseError(MyUnhandledError, "DDocumento:recuperarCodigoRubro Method")

End Function

Public Function registrarConstanteRubro(ByVal pnConsValor As Integer, _
                                             ByVal pcConsDescripcion As String) As Boolean

On Error GoTo registrarConstanteRubroErr
    Dim sSql As String
    Dim oDConecta As New DConecta
    
    registrarConstanteRubro = False

    sSql = " exec stp_ins_RegistrarConstanteRubro " & pnConsValor & ", '" & pcConsDescripcion & "'"

    If oDConecta.AbreConexion() Then
        Call oDConecta.Ejecutar(sSql)
        registrarConstanteRubro = True
        oDConecta.CierraConexion
    End If
    
    
    Set oDConecta = Nothing
    Exit Function

registrarConstanteRubroErr:
Call RaiseError(MyUnhandledError, "DDocumento:registrarConstanteRubro Method")
End Function

Public Function modificarConstanteRubro(ByVal pnConsValor As Integer, _
                                        ByVal pcConsDescripcion As String) As Boolean

On Error GoTo modificarConstanteRubroErr
    Dim sSql As String
    Dim oDConecta As New DConecta
    
    modificarConstanteRubro = False

    sSql = " exec stp_upd_ModificarConstanteRubro " & pnConsValor & ", '" & pcConsDescripcion & "'"

    If oDConecta.AbreConexion() Then
        Call oDConecta.Ejecutar(sSql)
        modificarConstanteRubro = True
        oDConecta.CierraConexion
    End If
    
    
    Set oDConecta = Nothing
    Exit Function

modificarConstanteRubroErr:
Call RaiseError(MyUnhandledError, "DDocumento:modificarConstanteRubro Method")
End Function

Public Function eliminarConstanteRubro(ByVal pnConsValor As Integer) As Boolean

On Error GoTo eliminarConstanteRubroErr
    Dim sSql As String
    Dim oDConecta As New DConecta
    
    eliminarConstanteRubro = False

    sSql = " exec stp_del_EliminarConstanteRubro " & pnConsValor & ""

    If oDConecta.AbreConexion() Then
        Call oDConecta.Ejecutar(sSql)
        eliminarConstanteRubro = True
        oDConecta.CierraConexion
    End If
    
    
    Set oDConecta = Nothing
    Exit Function

eliminarConstanteRubroErr:
Call RaiseError(MyUnhandledError, "DDocumento:eliminarConstanteRubro Method")
End Function
'***Fin Agregado por ELRO**********************************

'*** PEAC 20121001
Public Function BuscaNumCorrelaDocDDJJ(ByVal pnSerie As String, ByVal pnTipoDoc As String) As String ' ADODB.Recordset

On Error GoTo BuscaNumCorrelaDocDDJJErr
    Dim sSql As String
    Dim oDConecta As New DConecta
    Set oDConecta = New DConecta
    Dim rsBuscaCorre As ADODB.Recordset
    Set rsBuscaCorre = New ADODB.Recordset
    
    sSql = "exec stp_sel_ObtieneCorrelaDDJJ '" & pnSerie & "','" & pnTipoDoc & "'"

    If oDConecta.AbreConexion() Then
        Set rsBuscaCorre = oDConecta.CargaRecordSet(sSql)
        BuscaNumCorrelaDocDDJJ = rsBuscaCorre!cNum
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
BuscaNumCorrelaDocDDJJErr:
Call RaiseError(MyUnhandledError, "DDocumento:BuscaNumCorrelaDocDDJJ Method")

End Function
'EJVG20121130 ***
Public Function VerificaChequeCMAC(ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIFCod As String, ByVal psNroCheque As String) As Boolean
    Dim oConect As New DConecta
    Dim rs As New ADODB.Recordset
    Dim sSql As String
    VerificaChequeCMAC = False
    oConect.AbreConexion
    sSql = "Exec stp_sel_RecuperaChequeMovDoc '" & psIFTpo & "','" & psPersCod & "','" & psCtaIFCod & "','" & psNroCheque & "'"
    Set rs = oConect.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
       VerificaChequeCMAC = True
    End If
    Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
End Function
'END EJVG *******
