VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAgencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function GetAgenciaPorcentajeGastos() As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim oConec As DConecta
Set oConec = New DConecta
Dim sql As String
On Error GoTo ErrorGetAgenciaPorcentajeGastos
    sql = "exec stp_sel_AgenciaPorcentajeGastos"
    oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sql)
        Set GetAgenciaPorcentajeGastos = rs
        'RSClose rs
    oConec.CierraConexion
    Exit Function
ErrorGetAgenciaPorcentajeGastos:
    Err.Raise Err.Number, "AgenciaPorcentajeGastos", Err.Description
End Function

Public Sub ActualizarAgenciaPorcentajeGastos(ByVal psCodAge As String, pnPorGastos As Currency, pnPorCarteratAho As Currency, pnPorIngFinan As Currency)
Dim rs As ADODB.Recordset
Dim oConec As DConecta
Set oConec = New DConecta
Dim sql As String
On Error GoTo ErrorActualizarAgenciaPorcentajeGastos
    
    '*** PEAC 20100708
    'sql = "exec stp_upd_AgenciaPorcentajeGastos '" & psCodAge & "' , " & pnPorGastos
    sql = "exec stp_upd_AgenciaPorcentajeGastos '" & psCodAge & "' , " & pnPorGastos & "," & pnPorCarteratAho & "," & pnPorIngFinan
    
    oConec.AbreConexion
    oConec.CargaRecordSet (sql)
    oConec.CierraConexion
    Exit Sub
ErrorActualizarAgenciaPorcentajeGastos:
    Err.Raise Err.Number, "ActualizarAgenciaPorcentajeGastos", Err.Description
End Sub

Public Function GetAgencias(ByVal cCodAge As String) As String
        Dim oCon As DConecta
        Dim sSql As String
        Dim rs As ADODB.Recordset
        Set oCon = New DConecta
        If oCon.AbreConexion = False Then Exit Function

        sSql = "Select cAgeDescripcion From Agencias where cAgeCod='" & cCodAge & "'"
        Set rs = New ADODB.Recordset
        Set rs = oCon.CargaRecordSet(sSql)
        If Not (rs.EOF Or rs.BOF) Then
            GetAgencias = rs!CAGEDESCRIPCION & Space(20)
        Else
            GetAgencias = Space(20)
        End If
        
        
        oCon.CierraConexion
        Set oCon = Nothing
End Function
Public Function GetAgenciaPorcentajeGastosxValor() As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim oConec As DConecta
Set oConec = New DConecta
Dim sql As String
On Error GoTo ErrorGetAgenciaPorcentajeGastosxValor
    sql = "exec stp_sel_AgenciaPorcentajeGastosxvalor"
    oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sql)
        Set GetAgenciaPorcentajeGastosxValor = rs
        'RSClose rs
    oConec.CierraConexion
    Exit Function
ErrorGetAgenciaPorcentajeGastosxValor:
    Err.Raise Err.Number, "AgenciaPorcentajeGastosxValor", Err.Description
End Function

'*** PEAC 20100907
Public Function GetAgenciaPorcentajeSeguPatri() As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim oConec As DConecta
Set oConec = New DConecta
Dim sql As String
On Error GoTo ErrorGetAgenciaPorcentajeSeguPatri
    sql = "exec stp_sel_ObtienePorcentajeSeguPatri"
    oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sql)
        Set GetAgenciaPorcentajeSeguPatri = rs
    oConec.CierraConexion
    Exit Function
ErrorGetAgenciaPorcentajeSeguPatri:
    Err.Raise Err.Number, "GetAgenciaPorcentajeSeguPatri", Err.Description
End Function

'*** PEAC 20100907
Public Sub ActualizarAgenciaPorcentajeSeguPatri(ByVal psCodAge As String, ByVal pnTipoSegu As Integer, pnPorGastos As Currency)
Dim rs As ADODB.Recordset
Dim oConec As DConecta
Set oConec = New DConecta
Dim sql As String
On Error GoTo ErrorActualizarAgenciaPorcentajeSeguPatri
    
    sql = "exec stp_upd_PorcentajeSeguPatri '" & psCodAge & "' , " & pnTipoSegu & "," & pnPorGastos
    
    oConec.AbreConexion
    oConec.CargaRecordSet (sql)
    oConec.CierraConexion
    Exit Sub
ErrorActualizarAgenciaPorcentajeSeguPatri:
    Err.Raise Err.Number, "ActualizarAgenciaPorcentajeSeguPatri", Err.Description
End Sub

'*** PEAC 20100921
Public Function ObtieneAgencias() As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim oConec As DConecta
Set oConec = New DConecta
Dim sql As String
On Error GoTo ErrorObtieneAgencias

    sql = " EXEC stp_sel_ObtieneAgencias "
    
    oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sql)
        Set ObtieneAgencias = rs
    oConec.CierraConexion
    
    Exit Function
ErrorObtieneAgencias:
    Err.Raise Err.Number, "ObtieneAgencias", Err.Description
End Function

'*** PEAC 20100921
Public Function CargaPorcenPoliSeguro(Optional ByVal pnTipoPoli As Integer = 0) As Recordset
   
Dim dbConec As DConecta
Set dbConec = New DConecta
Dim psSql As String
   
   
   On Error GoTo CargaPorcenPoliSeguroErr
   
    If pnTipoPoli = 0 Then
        psSql = " exec stp_sel_ObtienePorcentajeSeguPatri "
    Else
        psSql = " exec stp_sel_ObtienePorcentajeSeguPatri " & pnTipoPoli
    End If
       
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      Set CargaPorcenPoliSeguro = dbConec.CargaRecordSet(psSql, 1)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaPorcenPoliSeguroErr:
   Call RaiseError(MyUnhandledError, "DDocumento:CargaPorcenPoliSeguro Method")
End Function

'*** PEAC 20100922
Public Sub ActualizarAgenciaPorcenPoliSeguPatri(ByVal psCodAge As String, pnTipoSeguro As Currency, pnPorcentaje As Currency)
Dim rs As ADODB.Recordset
Dim oConec As DConecta
Set oConec = New DConecta
Dim sql As String
On Error GoTo ErrorActualizarAgenciaPorcenPoliSeguPatri
    
    sql = "exec stp_upd_PorcentajeSeguPatri '" & psCodAge & "' , " & pnTipoSeguro & "," & pnPorcentaje
    
    oConec.AbreConexion
    oConec.CargaRecordSet (sql)
    oConec.CierraConexion
    Exit Sub
ErrorActualizarAgenciaPorcenPoliSeguPatri:
    Err.Raise Err.Number, "ActualizarAgenciaPorcenPoliSeguPatri", Err.Description
End Sub

'*** PEAC 20100924
Public Function DistribuyePoliSeguPatri(ByVal psFecha As String, ByVal pnTipoPoli As Integer, ByVal pnTCFijo As Double, ByVal pnTCVenta As Double) As Recordset
   
Dim dbConec As DConecta
Set dbConec = New DConecta
Dim psSql As String
   
   On Error GoTo CargaPorcenPoliSeguroErr
   
    psSql = " exec stp_sel_CalculaDistriPoliSeguPatri '" & psFecha & "'," & pnTipoPoli & "," & pnTCFijo & "," & pnTCVenta
       
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      Set DistribuyePoliSeguPatri = dbConec.CargaRecordSet(psSql, 1)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaPorcenPoliSeguroErr:
   Call RaiseError(MyUnhandledError, "DDocumento:CargaPorcenPoliSeguro Method")
End Function

'*** PEAC 20110303
Public Function CargaViaticoVisitaAgencias() As Recordset

Dim dbConec As DConecta
Set dbConec = New DConecta
Dim psSql As String
      
    On Error GoTo CargaViaticoVisitaAgenciasErr

    psSql = " exec stp_sel_CargaViaticoVisitaAgencias "
       
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
       Set CargaViaticoVisitaAgencias = dbConec.CargaRecordSet(psSql, 1)
       dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
CargaViaticoVisitaAgenciasErr:
   Call RaiseError(MyUnhandledError, "DDocumento:CargaViaticoVisitaAgencias Method")
End Function

'*** PEAC 20110530
Public Function ObtieneResumenCredPigno(ByVal psFecha As String, ByVal psAgencias As String) As Recordset

Dim dbConec As DConecta
Set dbConec = New DConecta
Dim psSql As String
   
   On Error GoTo ObtieneResumenCredPignoErr
   
    psSql = " exec stp_sel_ObtieneResumenCredPignoRecup '" & psFecha & "','" & psAgencias & "'"
       
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      Set ObtieneResumenCredPigno = dbConec.CargaRecordSet(psSql, 1)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
ObtieneResumenCredPignoErr:
   Call RaiseError(MyUnhandledError, "DDocumento:ObtieneResumenCredPigno Method")
End Function

'*** PEAC 20110705
Public Function ObtieneResumenGastosJudiciales(psAges As String, psTipCred As String, psFecDel As String, psFecAl As String, psMoneda As String) As Recordset

Dim dbConec As DConecta
Set dbConec = New DConecta
Dim psSql As String
   
   On Error GoTo ObtieneResumenGastosJudicialesErr

    psSql = " exec stp_sel_ObtieneResumenGastosJudPorAges '" & psAges & "','" & psTipCred & "','" & psFecDel & "','" & psFecAl & "','" & psMoneda & "'"

   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      Set ObtieneResumenGastosJudiciales = dbConec.CargaRecordSet(psSql, 1)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
ObtieneResumenGastosJudicialesErr:
   Call RaiseError(MyUnhandledError, "DDocumento:ObtieneResumenGastosJudiciales Method")
End Function

'***Agregado por ELRO el 20111026, según Acta 277-2011/TI-D
Public Function listarAniosCobertura() As ADODB.Recordset
    Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sSql As String
    
    On Error GoTo listarAniosCoberturaErr
    sSql = "exec stp_sel_ListarAniosCobertura"
    
    If oConec.AbreConexion() Then
      Set listarAniosCobertura = oConec.CargaRecordSet(sSql, 1)
      oConec.CierraConexion
    End If
    
    Set oConec = Nothing
    Exit Function
    
listarAniosCoberturaErr:
       Call RaiseError(MyUnhandledError, "listarAniosCobertura Method")
End Function

Public Function recuperarCoberturaAgencia(ByVal pnAnio As Integer, _
                                          ByVal pnMoneda As Integer) As ADODB.Recordset
   Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sSql As String
    
    On Error GoTo recuperarCoberturaAgenciaErr
    sSql = "exec stp_sel_RecuperarCoberturasAgencias '" & pnAnio & "', '" & pnMoneda & "'"
    
    If oConec.AbreConexion() Then
      Set recuperarCoberturaAgencia = oConec.CargaRecordSet(sSql, 1)
      oConec.CierraConexion
    End If
    
    Set oConec = Nothing
    Exit Function
    
recuperarCoberturaAgenciaErr:
       Call RaiseError(MyUnhandledError, "recuperarCoberturaAgencia Method")
End Function

Public Sub generarCoberturaAgencia(ByVal psCodAge As String, _
                                   ByVal pnMes As Integer, _
                                   ByVal pnAnio As Integer, _
                                   ByVal pnMoneda As Integer, _
                                   ByVal pnCobertura As Currency, _
                                   ByVal psMov As String)
    Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sSql As String
    
    On Error GoTo generarCoberturaAgenciaErr
    sSql = "exec stp_ins_GenerarCoberturaAgencia '" & psCodAge & "', '" & pnMes & "', '" & pnAnio & "', '" & pnMoneda & "', '" & pnCobertura & "', '" & psMov & "'"
    
    If oConec.AbreConexion() Then
        oConec.CargaRecordSet (sSql)
        oConec.CierraConexion
    End If
    
    Set oConec = Nothing
    Exit Sub
    
generarCoberturaAgenciaErr:
       Call RaiseError(MyUnhandledError, "generarCoberturaAgencia Method")

End Sub

Public Sub modificarCoberturaAgencia(ByVal psCodAge As String, _
                                   ByVal pnMes As Integer, _
                                   ByVal pnAnio As Integer, _
                                   ByVal pnMoneda As Integer, _
                                   ByVal pnCobertura As Currency, _
                                   ByVal psMov As String)
    Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sSql As String
    
    On Error GoTo modificarCoberturaAgenciaErr
    sSql = "exec stp_upd_ModificarCoberturaAgencia '" & psCodAge & "', '" & pnMes & "', '" & pnAnio & "', '" & pnMoneda & "', '" & pnCobertura & "', '" & psMov & "'"
    
    If oConec.AbreConexion() Then
        oConec.CargaRecordSet (sSql)
        oConec.CierraConexion
    End If
    
    Set oConec = Nothing
    Exit Sub
    
modificarCoberturaAgenciaErr:
       Call RaiseError(MyUnhandledError, "modificarCoberturaAgencia Method")

End Sub
'***Fin Agregado por ELRO

'*** PEAC 20120621
Public Function ObtieneComiBcoNac(ByVal psFecDel As String) As ADODB.Recordset
   Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sSql As String
    
    On Error GoTo ObtieneComiBcoNacErr
    
    sSql = "exec stp_sel_ComisionBcoNac '" & psFecDel & "'"

    If oConec.AbreConexion() Then
      Set ObtieneComiBcoNac = oConec.CargaRecordSet(sSql, 1)
      oConec.CierraConexion
    End If
    
    Set oConec = Nothing
    Exit Function

ObtieneComiBcoNacErr:
       Call RaiseError(MyUnhandledError, "ObtieneComiBcoNac Method")
End Function

'*** PEAC 20120628
Public Function CargaAsientoContComisionBN(ByVal psFecha As String) As ADODB.Recordset
    Dim oConec As DConecta
    Set oConec = New DConecta
    Dim psSql As String
    
    On Error GoTo CargaAsientoContComisionBNErr

      psSql = "exec stp_sel_CargaAsientoContComisionBN '" & psFecha & "'"

    If oConec.AbreConexion() Then
      Set CargaAsientoContComisionBN = oConect.CargaRecordSet(psSql, 1)
       oConec.CierraConexion
    End If
      Set oConec = Nothing
      
   Exit Function
CargaAsientoContComisionBNErr:
   Call RaiseError(MyUnhandledError, "DMov:CargaAsientoContComisionBN Method")
End Function
'EJVG20121113 ***
Public Function RecuperaAgenciasRptInfoEstadColocBCRP() As ADODB.Recordset
    Dim oConec As New DConecta
    Dim lsSql As String
    On Error GoTo ErrRecuperaAgenciasRptInfoEstadColocBCRP
    lsSql = "exec stp_sel_AgenciasRptInfoEstadColocBCRP"
    oConec.AbreConexion
    Set RecuperaAgenciasRptInfoEstadColocBCRP = oConec.CargaRecordSet(lsSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrRecuperaAgenciasRptInfoEstadColocBCRP:
    Err.Raise Err.Number, "RecuperaAgenciasRptInfoEstadColocBCRP", Err.Description
End Function
'END EJVG *******

'*** PEAC 20130613
Public Function CargaCOAparaPVS(psPeriodo As String, psCtaIgv As String, psOpeCod As String, Optional psDocTpo As String = "") As ADODB.Recordset
   Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sSql As String
    
    On Error GoTo CargaCOAparaPVSErr
    
    sSql = "exec stp_sel_ReportesTransferenciaComprobantesParaCOA '" & psPeriodo & "','" & psCtaIgv & "','" & gPersIdRUC & "','" & gMovEstContabMovContable & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "','" & gMovFlagModificado & "', '" & psOpeCod & "', '" & psDocTpo & "'"

    If oConec.AbreConexion() Then
      Set CargaCOAparaPVS = oConec.CargaRecordSet(sSql, 1)
      oConec.CierraConexion
    End If
    
    Set oConec = Nothing
    Exit Function

CargaCOAparaPVSErr:
       Call RaiseError(MyUnhandledError, "CargaCOAparaPVS Method")
End Function
