VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NConstSistemas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsFormatoFecha As String
Public Sub Inicio(ByVal psFormatoFecha As String)
vsFormatoFecha = psFormatoFecha
End Sub
Public Function ActualizaConstSistemas(ByVal pnConsSistCod As ConstSistemas, ByVal psMovUltAct As String, _
                              Optional ByVal psConsSistValor As String = "", _
                              Optional ByVal psConsSistDesc As String = "") As Integer
Dim oConSist As New DConstSistemas
    ActualizaConstSistemas = oConSist.ActualizaConsSistemas(pnConsSistCod, psMovUltAct, psConsSistValor, psConsSistDesc)
Set oConSist = Nothing
End Function
Public Function LeeConstSistema(ByVal psConstSistCod As ConstSistemas) As String
Dim rsVar As Recordset
Dim sSQL As String
Dim oCon  As DConecta
Set oCon = New DConecta

If oCon.AbreConexion = False Then Exit Function
sSQL = "Select nConsSisDesc, nConsSisValor From ConstSistema where nConsSisCod =" & psConstSistCod & ""
Set rsVar = New Recordset
Set rsVar = oCon.CargaRecordSet(sSQL)
LeeConstSistema = ""
If Not rsVar.EOF And Not rsVar.BOF Then
    LeeConstSistema = rsVar("nConsSisValor")
End If
rsVar.Close
Set rsVar = Nothing
Set oCon = Nothing
End Function
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
End Sub

Public Function GetRutaAcceso(psAgenciaCod As String) As String
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sql = "Select cRuta From RutaActualizacion Where cAgeCod = '" & psAgenciaCod & "'"
    
    oCon.AbreConexion
    GetRutaAcceso = oCon.CargaRecordSet(sql).Fields(0)
    oCon.CierraConexion
     
    Set oCon = Nothing
End Function
Public Function LeeConstSistemaTodos() As Recordset
Dim rs As Recordset
Dim sSQL As String
Dim oCon  As DConecta
Set oCon = New DConecta
Set rs = New Recordset
If oCon.AbreConexion = False Then Exit Function
sSQL = "Select nConsSisCod,nConsSisDesc, nConsSisValor From ConstSistema order by nConsSisCod"

Set rs = oCon.CargaRecordSet(sSQL)
Set LeeConstSistemaTodos = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Sub InsertaValor(psConsSisCod As Integer, psDescri As String, psValor As String, psMovNro)
On Error GoTo InsertaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "INSERT Constsistema (nConsSisCod, nConsSisDesc, nConsSisValor, cUltimaActuzalizacion) " _
            & "VALUES (" & psConsSisCod & ",'" & psDescri & "','" & psValor & "','" & psMovNro & "')"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
InsertaValorErr:
   Call RaiseError(MyUnhandledError, "NConstSistema:InsertaValor Method")
End Sub
Public Sub ActualizaValor(psConsSisCod As Integer, psDescri As String, psValor As String, psMovNro)
   On Error GoTo ActualizaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "UPDATE ConstSistema SET nConsSisDesc = '" & psDescri & "', nConsSisValor = '" & psValor & "', cUltimaActuzalizacion = '" & psMovNro & "' WHERE nConsSisCod = " & psConsSisCod & ""
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
ActualizaValorErr:
   Call RaiseError(MyUnhandledError, "DImpuesto:ActualizaValor Method")
End Sub
Public Function LeeConstSistemaCod() As Integer
Dim rsVar As Recordset
Dim sSQL As String
Dim oCon  As DConecta
Set oCon = New DConecta

If oCon.AbreConexion = False Then Exit Function
sSQL = "Select Max(nConsSisCod) as Codigo from  Constsistema"
Set rsVar = New Recordset
Set rsVar = oCon.CargaRecordSet(sSQL)
LeeConstSistemaCod = 0
If Not rsVar.EOF And Not rsVar.BOF Then
    LeeConstSistemaCod = rsVar("Codigo")
End If
rsVar.Close
Set rsVar = Nothing
Set oCon = Nothing

End Function
Public Sub EliminaValor(psConsSisCod As String)
   On Error GoTo EliminaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "DELETE Constsistema WHERE nConsSiscod = '" & psConsSisCod & "'"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
EliminaValorErr:
   Call RaiseError(MyUnhandledError, "DImpuesto:EliminaValor Method")
End Sub

'*** PEAC 20100706
Public Function LeeParametrosEncaje() As Recordset
Dim rs As Recordset
Dim sSQL As String
Dim oCon  As DConecta
Set oCon = New DConecta
Set rs = New Recordset
If oCon.AbreConexion = False Then Exit Function

sSQL = " exec stp_sel_ObtieneParametrosEncaje "

Set rs = oCon.CargaRecordSet(sSQL)
Set LeeParametrosEncaje = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

'*** PEAC 20100706
Public Sub EliminaValorParamEncaje(pnCodigo As Integer, psFecha As String)
   On Error GoTo EliminaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = " exec stp_del_ValorParamEncaje " & pnCodigo & ",'" & psFecha & "'"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
EliminaValorErr:
   Call RaiseError(MyUnhandledError, "DImpuesto:EliminaValor Method")
End Sub

'*** PEAC 20100706
Public Sub InsertaValorEncaje(pnCodigo As Integer, psFecha As String, psDescri As String, pnValor As Double)
On Error GoTo InsertaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
   
      psSql = " exec stp_ins_ValorParametroEncaje " & pnCodigo & ",'" & psFecha & "','" & psDescri & "'," & pnValor
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
InsertaValorErr:
   Call RaiseError(MyUnhandledError, "NConstSistema:InsertaValorEncaje")
End Sub

'*** PEAC 20100706
Public Sub ActualizaValorParametroEncaje(pnCodigo As Integer, psFecha As String, pnValor As Double)
   On Error GoTo ActualizaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = " exec stp_upd_ValorParametroEncaje " & pnCodigo & ",'" & psFecha & "'," & pnValor
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
ActualizaValorErr:
   Call RaiseError(MyUnhandledError, "DImpuesto:ActualizaValorParametroEncaje")
End Sub

'*** PEAC 20100706
Public Function BuscaParametrosEncaje(pnCodigo As Integer, psFecha As String) As Recordset
Dim rs As Recordset
Dim sSQL As String
Dim oCon  As DConecta
Set oCon = New DConecta
Set rs = New Recordset
If oCon.AbreConexion = False Then Exit Function

sSQL = " exec stp_sel_BuscaParametrosEncaje " & pnCodigo & ",'" & psFecha & "'"

Set rs = oCon.CargaRecordSet(sSQL)
Set BuscaParametrosEncaje = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

'*** PEAC 20110525
Public Function LeeEntidadesOpeReciprocas() As Recordset
Dim rs As Recordset
Dim sSQL As String
Dim oCon  As DConecta
Set oCon = New DConecta
Set rs = New Recordset
If oCon.AbreConexion = False Then Exit Function

sSQL = " exec stp_sel_EntidadesOpeReciprocas "

Set rs = oCon.CargaRecordSet(sSQL)
Set LeeEntidadesOpeReciprocas = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

'*** PEAC 20110525
Public Sub InsertaEntiOpeRecipro(psNumRuc As String, psCodEnti As String, pnSector As Integer, pnEstado As Integer, psAbrevia As String, psMovNro As String) ', psDisAho As String, psDisPF As String, psDepAh As String, psDepPF As String, psIngAh As String, psIngPF As String, psGasAh As String, psGasPF As String, psMovNro As String)
On Error GoTo InsertaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      
      'psSql = "exec stp_ins_EntiOpeRecipro '" & psNumRuc & "','" & psCodEnti & "'," & pnSector & "," & pnEstado ''& ",'" & psDisAho & "','" & psDisPF & "','" & psDepAh & "','" & psDepPF & "','" & psIngAh & "','" & psIngPF & "','" & psGasAh & "','" & psGasPF & "','" & psMovNro & "'"
        psSql = "exec stp_ins_EntiOpeRecipro '" & psNumRuc & "','" & psCodEnti & "'," & pnSector & "," & pnEstado & ",'" & psAbrevia & "','" & psMovNro & "'"
            
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
InsertaValorErr:
   Call RaiseError(MyUnhandledError, "NConstSistema:InsertaEntiOpeRecipro Method")
End Sub

'*** PEAC 20110525
Public Sub ActualizaEntiOpeRecipro(psNumRuc As String, psCodEnti As String, pnSector As Integer, pnEstado As Integer, psAbrevia As String, psMovNro As String) ', psDisAho As String, psDisPF As String, psDepAh As String, psDepPF As String, psIngAh As String, psIngPF As String, psGasAh As String, psGasPF As String, psMovNro As String)
   On Error GoTo ActualizaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      'psSql = "exec stp_upd_EntiOpeRecipro '" & psNumRuc & "','" & psCodEnti & "'," & pnSector & "," & pnEstado & ",'" & psDisAho & "','" & psDisPF & "','" & psDepAh & "','" & psDepPF & "','" & psIngAh & "','" & psIngPF & "','" & psGasAh & "','" & psGasPF & "','" & psMovNro & "'"
      psSql = "exec stp_upd_EntiOpeRecipro '" & psNumRuc & "','" & psCodEnti & "'," & pnSector & "," & pnEstado & ",'" & psAbrevia & "','" & psMovNro & "'"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
ActualizaValorErr:
   Call RaiseError(MyUnhandledError, "DImpuesto:ActualizaEntiOpeRecipro Method")
End Sub

'*** PEAC 20110525
Public Sub EliminaEntiOpeRecipro(psCodEntidad As String, pnSector As Integer)
   On Error GoTo EliminaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "exec stp_del_EntiOpeRecipro '" & psCodEntidad & "'," & pnSector
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
EliminaValorErr:
   Call RaiseError(MyUnhandledError, "DImpuesto:EliminaEntiOpeRecipro Method")
End Sub

'*** PEAC 20110610
Public Function LeeRutasViaticos() As Recordset
Dim rs As Recordset
Dim sSQL As String
Dim oCon  As DConecta
Set oCon = New DConecta
Set rs = New Recordset
If oCon.AbreConexion = False Then Exit Function

sSQL = " exec stp_sel_RutasViaticos "

Set rs = oCon.CargaRecordSet(sSQL)
Set LeeRutasViaticos = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

'*** PEAC 20110610
Public Sub InsertaRutasViaticos(psRuta As String)
On Error GoTo InsertaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      
      psSql = "exec stp_ins_RutasViaticos '" & psRuta & "'"
      
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
InsertaValorErr:
   Call RaiseError(MyUnhandledError, "NConstSistema:InsertaRutasViaticos Method")
End Sub

'*** PEAC 20110610
Public Sub ActualizaRutasViaticos(pnCod As Integer, psRuta As String)
   On Error GoTo ActualizaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "exec stp_upd_RutasViaticos " & pnCod & ",'" & psRuta & "'"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
ActualizaValorErr:
   Call RaiseError(MyUnhandledError, "DImpuesto:ActualizaRutasViaticos Method")
End Sub

'*** PEAC 20110630
Public Function LeeCuentasContEntiFinan(psTipCta As String, psPersCod As String) As Recordset
Dim rs As Recordset
Dim sSQL As String
Dim oCon  As DConecta
Set oCon = New DConecta
Set rs = New Recordset
If oCon.AbreConexion = False Then Exit Function

sSQL = "exec stp_sel_ObtieneCtasEntiFinan '" & psPersCod & "','" & psTipCta & "'"

Set rs = oCon.CargaRecordSet(sSQL)
Set LeeCuentasContEntiFinan = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

'*** PEAC 20110630
Public Sub InsertaCtasContEntiFinanOpeRecipro(psTipAho As String, psCtaCont As String, psTipoCta As String, psPersCod As String)

On Error GoTo InsertaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      
      psSql = "exec stp_ins_CtaContEntiFinanOpeRecipro '" & psTipAho & "','" & psCtaCont & "','" & psTipoCta & "','" & psPersCod & "'"
      
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
InsertaValorErr:
   Call RaiseError(MyUnhandledError, "NConstSistema:InsertaCtasContEntiFinanOpeRecipro Method")
End Sub

'*** PEAC 20110630
Public Sub EliminaCtasContEntiFinanOpeRecipro(psTipAho As String, psCtaCont As String, psTipoCta As String, psPersCod As String)
   On Error GoTo EliminaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   If dbconec.AbreConexion() Then
      psSql = "exec stp_del_CtaContEntiFinanOpeRecipro '" & psTipAho & "','" & psCtaCont & "','" & psTipoCta & "','" & psPersCod & "'"
      dbconec.Ejecutar psSql
      dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   Exit Sub
EliminaValorErr:
   Call RaiseError(MyUnhandledError, "DImpuesto:EliminaCtasContEntiFinanOpeRecipro Method")
End Sub

'*** PEAC 20110630
Public Function BuscaCtasContEntiFinanOpeRecipro(psTipAho As String, psCtaCont As String, psTipoCta As String, psPersCod As String) As Boolean

On Error GoTo InsertaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   If dbconec.AbreConexion() Then
      
        psSql = "exec stp_sel_CtaContEntiFinanOpeRecipro '" & psTipAho & "','" & psCtaCont & "','" & psTipoCta & "','" & psPersCod & "'"
      
        Set rs = dbconec.CargaRecordSet(psSql)

        dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   
    If (rs.EOF And rs.BOF) Then
        BuscaCtasContEntiFinanOpeRecipro = False
    Else
        BuscaCtasContEntiFinanOpeRecipro = True
    End If
   
   Exit Function
InsertaValorErr:
   Call RaiseError(MyUnhandledError, "NConstSistema:InsertaCtasContEntiFinanOpeRecipro Method")
   
End Function

'*** PEAC 20110722
Public Function BuscaEntiOpeRecipro(psPersCod As String, pnSector As Integer) As Boolean

On Error GoTo InsertaValorErr
   Dim dbconec As DConecta
   Dim psSql As String
   Set dbconec = New DConecta
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   If dbconec.AbreConexion() Then
      
        psSql = "exec stp_sel_BuscaEntiOpeRecipro '" & psPersCod & "'," & pnSector
      
        Set rs = dbconec.CargaRecordSet(psSql)

        dbconec.CierraConexion
   End If
   Set dbconec = Nothing
   
    If (rs.EOF And rs.BOF) Then
        BuscaEntiOpeRecipro = False
    Else
        BuscaEntiOpeRecipro = True
    End If
   
   Exit Function
InsertaValorErr:
   Call RaiseError(MyUnhandledError, "NConstSistema:BuscaEntiOpeRecipro Method")
   
End Function

'*** PASI 20140401 TI-ERS050-2014
Public Sub EliminaRutasViaticos(ByVal pnCodRuta As Integer)
    On Error GoTo ErrorEliminaRutasViaticos
    Dim dbconec As DConecta
    Dim psSql As String
    Set dbconec = New DConecta
    If dbconec.AbreConexion() Then
        psSql = "stp_del_RutasViaticos " & pnCodRuta
        dbconec.Ejecutar psSql
        dbconec.CierraConexion
    End If
    Set dbconec = Nothing
    Exit Sub
ErrorEliminaRutasViaticos:
    Call RaiseError(MyUnhandledError, "DImpuesto:EliminaRutasViaticos Method")
End Sub
'*** END PASI
