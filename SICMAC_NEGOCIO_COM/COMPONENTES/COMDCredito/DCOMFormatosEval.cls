VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMFormatosEval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************************************************************
'Descripcion: Modulo de Clase para los formatos de Evaluacion
'Referencia: ERS004-2016 (Formatos de Evaluacion)
'Creado por: LUCV20160524 | PEAC |
'Ultima Modificacion: 20160524
'*****************************************************************************
'*** PEAC 20160523

Option Explicit
Public oDCOMConecta As COMConecta.DCOMConecta

Public Function RecuperaConfigFornatosEval() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_sel_formatosEvaluacion "
    oCon.AbreConexion
    Set RecuperaConfigFornatosEval = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20160523
Public Function ActualizaConfFormatosEval(ByVal pnCodForm As Integer, ByVal pnMontoMin As Double, ByVal onMontoMax As Double) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_upd_ConfigFormatosEval " & pnCodForm & " , " & pnMontoMin & " , " & onMontoMax
    oCon.AbreConexion
    Set ActualizaConfFormatosEval = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20160525
Public Function AgregaCredFormEvalConfigProd(ByVal pnCodForm As Integer, ByVal pcProd As String, ByVal pcSubProd As String, ByVal pnAplica As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_ins_CredFormEvalConfigProd " & pnCodForm & " , '" & pcProd & "' , '" & pcSubProd & "', " & pnAplica
    oCon.AbreConexion
    Set AgregaCredFormEvalConfigProd = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20160523
Public Function ObtenerProducto(ByVal pnConsCod As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_productos " & pnConsCod
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerProducto = rs
End Function

'*** PEAC 20160523
Public Function ObtenerSubProducto(ByVal pnConsCod As Integer, ByVal pnConsValor As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_SubProductos " & pnConsCod & ", " & pnConsValor
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerSubProducto = rs
End Function

'*** PEAC 20160525
Public Function ObtenerFormatosSeleccionados(ByVal pcTpoProd As String, ByVal pcTpoSubProd As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_FormatosSeleccionados '" & pcTpoProd & "', '" & pcTpoSubProd & "' "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFormatosSeleccionados = rs
End Function

'*** PEAC 20160525
Public Function ObtenerFormatosChkParaElegir(ByVal pcTpoProd As String, ByVal pcTpoSubProd As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_CredFormEvalFormatosElegidos '" & pcTpoProd & "', '" & pcTpoSubProd & "' "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFormatosChkParaElegir = rs
End Function

'*** PEAC 20160525
Public Function LimpiaFormatosProdEval(ByVal pcTpoProd As String, ByVal pcTpoSubProd As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_del_FormatosSeleccionados '" & pcTpoProd & "', '" & pcTpoSubProd & "' "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set LimpiaFormatosProdEval = rs
End Function

'*** PEAC 20160530
Public Function LimpiaProdEvalRatios(ByVal pcTpoProd As String, ByVal pcTpoSubProd As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_del_CredFormEvalRatios '" & pcTpoProd & "', '" & pcTpoSubProd & "' "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set LimpiaProdEvalRatios = rs
End Function


'*** PEAC 20160529
Public Function AgregaCredFormEvalRatios(ByVal pcProd As String, ByVal pcSubProd As String, ByVal pnIfis As Integer, ByVal pnIndispen As Integer, _
        ByVal pnMesesDel As Integer, ByVal pnMesesAl As Integer, ByVal pnMontoDel As Double, ByVal pnMontoAl As Double, ByVal pnAceptableDel As Double, ByVal pnAceptableAl As Double, _
        ByVal pnCriticoDel As Double, ByVal pnCriticoAl As Double, ByVal pnTipoCli As Integer, ByVal pnRatio As Integer, ByVal pnLimite As Integer, ByVal pnLimitePorcentaje As Integer) As ADODB.Recordset
        
'pnIfis pnIndispen pnMesesDel pnMesesAl pnMontoDel pnMontoAl pnAceptableDel pnAceptableAl pnTipoCli pnRatio

    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_ins_CredFormEvalRatios '" & pcProd & "' , '" & pcSubProd & "', " & pnIfis & ", " & pnMesesDel & ", " & pnMesesAl & ", " & pnTipoCli & "," & pnMontoDel & ", " & _
                            pnMontoAl & ", " & pnRatio & ", " & pnAceptableDel & ", " & pnAceptableAl & ", " & pnCriticoDel & ", " & pnCriticoAl & ", " & pnIndispen & ", " & pnLimite & "," & pnLimitePorcentaje
    'CTI3 ERS0032020 - Se agrregó los valores de límite y pnLimitePorcentaje por producto consumo
'@cTpoProd char(3), @cTpoSubProd char(3), @nIfis int, @nMesesDel int, @nMesesAl int, @pnTipoCli int, @nMontoDel money,@nMontoAl money,
'@nRatio int, @nAceptableDel money, @nAceptableAl money,@nCriticoDel money, @nCriticoAl money, @nIndispen int
    
    oCon.AbreConexion
    Set AgregaCredFormEvalRatios = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20160529
Public Function ObtenerFormatosEvalRatios(ByVal pcTpoProd As String, ByVal pcTpoSubProd As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_CredFormEvalRatios '" & pcTpoProd & "', '" & pcTpoSubProd & "' "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFormatosEvalRatios = rs
End Function
'*** PEAC 20160608
Public Function ObtenerFormatosEvalActivos(ByVal pcCtaCod As String, ByVal pnCodForm As Integer, Optional ByVal pcFecReg As String = "19000101") As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_CredFormEvalActivo '" & pcCtaCod & "', " & pnCodForm & ", '" & pcFecReg & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFormatosEvalActivos = rs
End Function
'*** PEAC 20160608
Public Function ObtenerFormatosEvalActivosPlantilla(ByVal pnCodForm As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_CredFormEvalActivoPlantilla " & pnCodForm
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFormatosEvalActivosPlantilla = rs
End Function

'*** PEAC 20160608
Public Function ObtenerFormatosEvalPasivos(ByVal pcCtaCod As String, ByVal pnCodForm As Integer, Optional ByVal pcFecReg As String = "19000101") As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_CredFormEvalPasivos '" & pcCtaCod & "', " & pnCodForm & ", '" & pcFecReg & "' "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFormatosEvalPasivos = rs
End Function
'*** PEAC 20160608
Public Function ObtenerFormatosEvalPasivosPlantilla(ByVal pnCodForm As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_CredFormEvalPasivosPlantilla " & pnCodForm
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFormatosEvalPasivosPlantilla = rs
End Function

'*** PEAC 20160614
Public Function ObtenerFormatosEvalEstadoGanPerd(ByVal pcCtaCod As String, ByVal pnCodForm As Integer, Optional ByVal pcFecReg As String = "19000101") As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_CredFormEvalEstadoGanPerd '" & pcCtaCod & "', " & pnCodForm & ", '" & pcFecReg & "' "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFormatosEvalEstadoGanPerd = rs
End Function

'*** PEAC 20160614
Public Function ObtenerFormatosEvalEstadoGanPerdPlantilla(ByVal pnCodForm As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_CredFormEvalEstadoGananciaPerdidaPlantilla " & pnCodForm
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFormatosEvalEstadoGanPerdPlantilla = rs
End Function

'*** PEAC 20160629
Public Function AgregaCredFormEvalEstFinActivosPasivosFormato6(ByVal pcCtaCod As String, ByVal pnCodForm As Integer, ByVal pcFecReg As String, ByVal pnAuditado As Integer, _
                                            ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnPP As Double, ByVal pnPE As Double, ByVal pnTotal As Double) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_ins_CredFormEvalEstFinActivosPasivosFormato6 '" & pcCtaCod & "', " & pnCodForm & ", '" & pcFecReg & "', " & pnAuditado & _
                                            ", " & pnConsCod & ", " & pnConsValor & ", " & pnPP & ", " & pnPE & ", " & pnTotal
    oCon.AbreConexion
    Set AgregaCredFormEvalEstFinActivosPasivosFormato6 = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20160629
Public Function AgregaCredFormEvalEstFinActivosPasivosFormato6Det(ByVal pcCtaCod As String, ByVal pnCodForm As Integer, ByVal pcFecReg As String, ByVal pnAuditado As Integer, _
                                            ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnPP As Double, ByVal pnPE As Double, ByVal pnTotal As Double) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_ins_CredFormEvalEstFinActivosPasivosFormato6Det '" & pcCtaCod & "', " & pnCodForm & ", '" & pcFecReg & "', " & pnAuditado & _
                                            ", " & pnConsCod & ", " & pnConsValor & ", " & pnPP & ", " & pnPE & ", " & pnTotal
    oCon.AbreConexion
    Set AgregaCredFormEvalEstFinActivosPasivosFormato6Det = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20160629
Public Function AgregaCredFormEvalEstFinActivosPasivosFormato6DetDetalle( _
ByVal pcCtaCod As String, _
ByVal pnCodForm As Integer, _
ByVal pcFecReg As String, _
ByVal pnAuditado As Integer, _
ByVal pnConsCod As Integer, _
ByVal pnConsValor As Integer, _
ByVal pnTipoPatrimonio As Integer, _
ByVal pcFechaDet As String, _
ByVal pcDescripcionDet As String, _
ByVal pnPP As Double, _
ByVal pnPE As Double, _
ByVal pcCodIfi As String) As ADODB.Recordset

    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_ins_CredFormEvalEstFinActivosPasivosFormato6DetDetalle '" & pcCtaCod & "', " & pnCodForm & ", '" & pcFecReg & "', " & pnAuditado & _
                                            ", " & pnConsCod & ", " & pnConsValor & ", " & pnTipoPatrimonio & ",'" & pcFechaDet & "','" & pcDescripcionDet & "', " & pnPP & ", " & pnPE & ", '" & pcCodIfi & "' "
    oCon.AbreConexion
    Set AgregaCredFormEvalEstFinActivosPasivosFormato6DetDetalle = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20160629
Public Function AgregaCredFormEvalEstFinEstGanPerFormato6(ByVal pcCtaCod As String, ByVal pnCodForm As Integer, ByVal pcFecReg As String, ByVal pnAuditado As Integer, _
                                            ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnMonto As Double) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_ins_CredFormEvalEstFinEstGanPerFormato6 '" & pcCtaCod & "', " & pnCodForm & ", '" & pcFecReg & "', " & pnAuditado & _
                                            ", " & pnConsCod & ", " & pnConsValor & ", " & pnMonto
    oCon.AbreConexion
    Set AgregaCredFormEvalEstFinEstGanPerFormato6 = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20160629
Public Function ValidaFechaCredFormEval(ByVal pcCtaCod As String, ByVal pnCodForm As Integer, ByVal pcFechaEval As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ValidaFechaCredFormEval '" & pcCtaCod & "'," & pnCodForm & ",'" & pcFechaEval & "' "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ValidaFechaCredFormEval = rs
End Function

'*** PEAC 20160629
Public Function ObtenerFormatosEvalActivosDet(ByVal pcCtaCod As String, ByVal pnCodForm As Integer, Optional ByVal pcFecReg As String = "19000101") As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_CredFormEvalActivoDetForm6 '" & pcCtaCod & "', " & pnCodForm & ", '" & pcFecReg & "' "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFormatosEvalActivosDet = rs
End Function

'*** PEAC 20160629
Public Function ObtenerFormatosEvalPasivosDet(ByVal pcCtaCod As String, ByVal pnCodForm As Integer, Optional ByVal pcFecReg As String = "19000101") As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_CredFormEvalPasivoDetForm6 '" & pcCtaCod & "', " & pnCodForm & ", '" & pcFecReg & "' "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFormatosEvalPasivosDet = rs
End Function

'*** PEAC 20160703
Public Function RecuperaDatosInfCome(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer

    On Error GoTo ErrorRecuperaInfCome

    sSql = "exec stp_sel_ObtieneDatosInformeComercial '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosInfCome = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function
ErrorRecuperaInfCome:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*** PEAC 20160703
Public Function RecuperaCaliSbs(ByVal psDocNat As String, ByVal psDocJur As String, Optional ByVal psTpoDoc As Integer = 1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCaliSbs
  
    sSql = "exec stp_sel_ObtieneCaliSbs '" & psDocNat & "','" & psDocJur & "', " & psTpoDoc & " "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCaliSbs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCaliSbs:
                  Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'*** PEAC 20160703
Public Function RecuperaDatosFinan(ByVal psCtaCod As String, Optional ByVal pnNumForm As Integer = 0) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaDatosFinan
      
    'sSql = "exec stp_sel_ObtieneDatosFinancieros '" & psCtaCod & "'"
    sSql = "exec stp_sel_ObtieneRatioIndicador '" & psCtaCod & "'"
       
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosFinan = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosFinan:
    Err.Raise Err.Number, "Error En Proceso.", Err.Description
End Function

'*** PEAC 20160703
Public Function RecuperaCapacidadPago(ByVal pcCtaCod As String, ByVal pnCodForm As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCapacidadPago
  
    sSql = "exec stp_sel_ObtieneRatioIndicador '" & pcCtaCod & "'," & pnCodForm
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCapacidadPago = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCapacidadPago:
                  Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'*** PEAC 20160706
Public Function ObtenerFormatosEvalCoeficienteFinan(ByVal pcCtaCod As String, ByVal pnCodForm As Integer, Optional ByVal pcFecReg As String = "19000101") As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_CredFormEvalCoeficienteFinan '" & pcCtaCod & "', " & pnCodForm & ", '" & pcFecReg & "' "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFormatosEvalCoeficienteFinan = rs
End Function

'*** PEAC 20160629
Public Function AgregaCredFormEvalPropuCredFormato6(ByVal pcCtaCod As String, ByVal pnCodForm As Integer, ByVal pcFecVis As String, _
                ByVal pcEntornoFam As String, ByVal pcGiroUbi As String, ByVal pcExpeCred As String, ByVal pcConsistInfor As String, _
                ByVal pcColaterales As String, ByVal pcDestino As String) As ADODB.Recordset

    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_ins_CredFormEvalPropuestaFormato1_5 '" & pcCtaCod & "', " & pnCodForm & ", '" & pcFecVis & "', '" & pcEntornoFam & _
                                            "', '" & pcGiroUbi & "', '" & pcExpeCred & "', '" & pcConsistInfor & "','" & pcColaterales & "','" & pcDestino & "'"
    oCon.AbreConexion
    Set AgregaCredFormEvalPropuCredFormato6 = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20160629
Public Function AgregaCredFormEvalReferidosFormato6(ByVal pcCtaCod As String, ByVal pnCodForm As Integer, ByVal pnCodRef As Integer, _
                ByVal pcNombre As String, ByVal pcDNINom As String, ByVal pcTelf As String, ByVal pcReferido As String, ByVal pcDNIRef As String) As ADODB.Recordset

    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_ins_CredFormEvalReferidos '" & pcCtaCod & "', " & pnCodForm & ", " & pnCodRef & ", '" & pcNombre & _
                                            "', '" & pcDNINom & "', '" & pcTelf & "', '" & pcReferido & "','" & pcDNIRef & "' "
    oCon.AbreConexion
    Set AgregaCredFormEvalReferidosFormato6 = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20160629
Public Function AgregaCredFormEvalCoeficienteFinanFormato6(ByVal pcCtaCod As String, ByVal pnCodForm As Integer, ByVal pcFecReg As String, _
                ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnMonto As Double) As ADODB.Recordset

    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_ins_CredFormEvalCoeficienteFinan '" & pcCtaCod & "', " & pnCodForm & ", '" & pcFecReg & "', " & pnConsCod & _
                                            ", " & pnConsValor & ", " & pnMonto
    oCon.AbreConexion
    Set AgregaCredFormEvalCoeficienteFinanFormato6 = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'*** PEAC 20160730
Public Function RecuperaDatosCredEvalFlujoCajaForm6(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_CredFormEvalFlujoCajaForm6 '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalFlujoCajaForm6 = rs
    Set rs = Nothing
End Function

'*** PEAC 20160730
Public Function RecuperaDatosCredEvalGastosFamForm6(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_CredFormEvalGastosFamForm6 '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalGastosFamForm6 = rs
    Set rs = Nothing
End Function

'*** PEAC 20160730
Public Function RecuperaDatosCredEvalOtrosIngresosForm6(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_CredFormEvalOtrosIngresosForm6 '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalOtrosIngresosForm6 = rs
    Set rs = Nothing
End Function

'*** PEAC 20160810
Public Function RecuperaDatosCredEvalFechasEEFFEvalForm6(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_FechasEEFFEvalForm6 '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalFechasEEFFEvalForm6 = rs
    Set rs = Nothing
End Function

'*** PEAC 20160810
Public Function CargarIfisDetalleForm6(Optional pnConsCod As Integer = -1, Optional pnConsValor As Integer = -1) As Recordset 'LUCV20161115, Modificó->Según ERS068-2016 (Agregó parametros)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrCargarIfisDetalleForm6
    sSql = "EXEC stp_sel_CargarIfisDetalleForm6 " & pnConsCod & ", " & pnConsValor & " " 'LUCV20161115, Modificó->Según ERS068-2016 (Agregó parametros)
    oConecta.AbreConexion
    Set CargarIfisDetalleForm6 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrCargarIfisDetalleForm6:
    Err.Raise Err.Number, "Error En Proceso <<CargarIfisDetalleForm6>>", Err.Description
End Function

'*** PEAC 20160812
Public Function ObtieneUltFecEEFFForm6(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrObtieneUltFecEEFFForm6
    sSql = "EXEC stp_sel_UltFechaEvalForm6 '" & pcCtaCod & "' "
    oConecta.AbreConexion
    Set ObtieneUltFecEEFFForm6 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrObtieneUltFecEEFFForm6:
    Err.Raise Err.Number, "Error En Proceso <<ObtieneUltFecEEFFForm6>>", Err.Description
End Function
'*** PEAC 20160820
Public Function ObtieneDatosComunes(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrObtieneDatosComunes
    sSql = "EXEC stp_sel_ObtieneDatosComunesFormEval '" & pcCtaCod & "' "
    oConecta.AbreConexion
    Set ObtieneDatosComunes = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrObtieneDatosComunes:
    Err.Raise Err.Number, "Error En Proceso <<ObtieneDatosComunes>>", Err.Description
End Function


'_____.....:::::********** LUCV20160526 **********:::::....._____
'LUCV20160524
Public Function RecuperaCredFormEvalProductoEstadoExposicion(ByVal psCtaCod As String) As ADODB.Recordset ' Estado y Exposcion con este crédito
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCredFormEvalProductoEstadoExposicion
    sSql = "stp_sel_RecuperaCredFormEvalProductoEstadoExposicion '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCredFormEvalProductoEstadoExposicion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCredFormEvalProductoEstadoExposicion:
    Err.Raise Err.Number, "Recuperar Datos de exposición con este crédito y estado del credito", Err.Description
End Function


'LUCV20160524
Public Function RecuperaSolicitudDatoBasicosEval(ByVal psCtaCod As String) As ADODB.Recordset 'CabeceraFormato(1,2,3,4,5)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaSolicitudDatoBasicosEval
    
    sSql = "stp_sel_DatosBasicosSolicitudEval '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSolicitudDatoBasicosEval = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaSolicitudDatoBasicosEval:
    Err.Raise Err.Number, "Recuperar Solicitud Datos Basicos", Err.Description
End Function

Public Function RecuperaDeudaLineaCreditoNU(ByVal psCtaCod As String) As ADODB.Recordset 'En fila de FeGastosFamiliares(1,2,3,4,5)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDeudaLineaCreditoNU
    
    sSql = "stp_sel_ObtenerDeudaLineaCreditoNU '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDeudaLineaCreditoNU = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDeudaLineaCreditoNU:
    Err.Raise Err.Number, "Recuperar Solicitud Datos Basicos", Err.Description
End Function


'LUCV20160620
Public Function RecuperaSolicitudDatoBasicosSolicitud(ByVal psCtaCod As String) As ADODB.Recordset 'PreCarda-> ObtenerFormato
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaSolicitudDatoBasicosSolicitud
    
    sSql = "stp_sel_DatosBasicosSolicitud '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSolicitudDatoBasicosSolicitud = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaSolicitudDatoBasicosSolicitud:
    Err.Raise Err.Number, "Recuperar Solicitud Datos Basicos", Err.Description
End Function

'CTI3 ERS0032020*******************************************************************
'LUCV20160524
'Public Function AsignarFormato(ByVal psPrd As String, ByVal psSPrd As String, ByVal pnMonto As Double) ', Optional ByRef pnMin As Double, Optional ByRef pnMax As Double) As Integer
Public Function AsignarFormato(ByVal psPrd As String, ByVal psSPrd As String, ByVal pnMonto As Double, ByVal nTpoCliente As Integer, ByVal nTpoIng As Long)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

        'sSql = "stp_sel_AsignarCredFormEval '" & psPrd & "','" & psSPrd & "'," & pnMonto
        sSql = "stp_sel_AsignarCredFormEval_new '" & psPrd & "','" & psSPrd & "'," & pnMonto & ", " & nTpoCliente & ", " & nTpoIng
        Set oCon = New COMConecta.DCOMConecta
        oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(sSql)
        oCon.CierraConexion
        Set oCon = Nothing
        
        If rs.RecordCount > 0 Then
            If Not (rs.EOF And rs.BOF) Then
                AsignarFormato = CInt(rs!ncodForm)
                'pnMin = CDbl(rs!nMontoMin)
                'pnMax = CDbl(rs!nMontoMax)
            Else
                AsignarFormato = 0
                'pnMin = 0
                'pnMax = 0
            End If
        Else
            AsignarFormato = 0
            'pnMin = 0
            'pnMax = 0
        End If
        
                          
                  
                  
          
    
    Set rs = Nothing
End Function

'LUCV20160525: Carga de Datos
Public Function RecuperaColocacCredEval(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperaCredFormEval '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaColocacCredEval = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosCredEvalGastosNeg(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosCredFormEvalGastosNego '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalGastosNeg = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosCredEvalGastosFam(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosCredFormEvalGastosFam '" & psCtaCod & "'"
        Set oCon = New COMConecta.DCOMConecta
        oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(sSql)
        oCon.CierraConexion
        Set oCon = Nothing
        
        Set RecuperaDatosCredEvalGastosFam = rs
        Set rs = Nothing
End Function
Public Function RecuperaDatosCredEvalOtrosIngr(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosCredFormEvalOtrosIng '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalOtrosIngr = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosCredEvalCuotaIfis(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosCredFormEvalCuotaIfis '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalCuotaIfis = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosCredEvalPropuesta(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosCredFormEvalPropuesta '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalPropuesta = rs
    Set rs = Nothing
End Function
'Public Function RecuperaDatosCapacidadPagoNeta(ByVal psCtaCod As String) As ADODB.Recordset
'    Dim sSql As String
'    Dim oCon As COMConecta.DCOMConecta
'    Dim rs As ADODB.Recordset
'
'    sSql = "exec stp_sel_RecuperaDatosCapacidadPagoNeta'" & psCtaCod & "'"
'    Set oCon = New COMConecta.DCOMConecta
'    oCon.AbreConexion
'    Set rs = oCon.CargaRecordSet(sSql)
'    oCon.CierraConexion
'    Set oCon = Nothing
'    Set RecuperaDatosCapacidadPagoNeta = rs
'    Set rs = Nothing
'End Function

Public Function RecuperaDatosCredEvalVentaCosto(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosVentaCosto '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set RecuperaDatosCredEvalVentaCosto = rs
    Set rs = Nothing
End Function

Public Function RecuperaDatosReferidos(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosReferidos'" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set RecuperaDatosReferidos = rs
    Set rs = Nothing
End Function

Public Function RecuperaCuotaPropuesta(ByVal psCtaCod As String) As ADODB.Recordset 'CabeceraFormato(1,2,3,4,5)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCuotaPropuesta
    
    sSql = "stp_sel_RecuperaDatosCuotaPropuesta '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCuotaPropuesta = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCuotaPropuesta:
    Err.Raise Err.Number, "Recuperar Solicitud Datos Basicos", Err.Description
End Function


'LUCV20160525: Carga Datos Grillas(FlexEdit) ************************************************
'LUCV, Lista feGastosNegocio
Public Function RecuperaDatosFlexGastosNegocio(ByVal psCtaCod As String, ByVal pnFormato As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_FlexCredFormEvalGastoNego '" & psCtaCod & "', " & pnFormato & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexGastosNegocio = rs
    Set rs = Nothing
End Function
'LUCV, Lista feOtrosIngresos
Public Function RecuperaDatosFlexOtrosIngresos(ByVal pnFormato As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_FlexCredFormEvalOtrosIngresos '" & pnFormato & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexOtrosIngresos = rs
    Set rs = Nothing
End Function

'LUCV, Lista feGastosFamiliares
Public Function RecuperaDatosFlexGastosFamiliares(ByVal pcCtaCod As String, ByVal pnFormato As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_FlexCredFormEvalGastosFamiliares '" & pcCtaCod & "', " & pnFormato & " "
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexGastosFamiliares = rs
    Set rs = Nothing
End Function

'LUCV, Lista feBalanceGeneral
Public Function RecuperaDatosFlexBalanceGeneral(ByVal psCtaCod As String, ByVal pnFormato As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_FlexCredFormEvalBalanceGeneral '" & psCtaCod & "', " & pnFormato & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexBalanceGeneral = rs
    Set rs = Nothing
End Function

'LUCV, Lista feActivos
Public Function RecuperaDatosFlexActivos(ByVal pnFormato As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_FlexCredFormEvalActivos '" & pnFormato & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexActivos = rs
    Set rs = Nothing
End Function

Public Function RecuperaDatosFlexPasivos(ByVal psCtaCod As String, ByVal pnFormato As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_FlexCredFormEvalPasivos '" & psCtaCod & "'," & pnFormato & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexPasivos = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosFlexFlujoCaja(ByVal pnFormato As Integer, ByVal pcCtaCod As String, Optional ByVal pnTipoFlujo As Integer = 0) As ADODB.Recordset
'LUCV20171015, Agregó según ERS0512017: Optional ByVal pnTipoFlujo As Integer = 0
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_FlexCredFormEvalFlujoCaja   " & pnFormato & ", '" & pcCtaCod & "', " & pnTipoFlujo & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexFlujoCaja = rs
    Set rs = Nothing
End Function

Public Function RecuperaDatosFlexFlexDeclaracionPdt(ByVal pnFormato As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_FlexCredFormEvalDeclaracionPDT '" & pnFormato & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexFlexDeclaracionPdt = rs
    Set rs = Nothing
End Function


Public Function RecuperaDatosRatios(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosRatiosIndicadores '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosRatios = rs
    Set rs = Nothing
End Function


Public Function RecuperaDatosRatiosAceptableCritico(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RatiosFinancierosAceptableCritico '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosRatiosAceptableCritico = rs
    Set rs = Nothing
End Function


Public Function RecuperaDatosActivos(ByVal psCtaCod As String, ByVal pnCodActivo As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosActivos '" & psCtaCod & "', " & pnCodActivo & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosActivos = rs
    Set rs = Nothing
End Function

Public Function RecuperaDatosPasivos(ByVal psCtaCod As String, ByVal pnCodPasivo As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosPasivos '" & psCtaCod & "', " & pnCodPasivo & " "
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosPasivos = rs
    Set rs = Nothing
End Function


Public Function RecuperaDatosActivoPasivo(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosActivoPasivo1_6 '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosActivoPasivo = rs
    Set rs = Nothing
End Function

Public Function RecuperaDatosIfiCuota(ByVal psCtaCod As String, ByVal pnCodForm As Integer, ByVal pnConsCod As Integer, Optional ByVal pnConsValor As Integer = 0, Optional ByVal pnTipoFlujoCaja As Integer = 0) As ADODB.Recordset
'JOEP20171015 Segun ERS051-2017 Flujo de Caja Optional ByVal pnConsValor As Integer = 0
'LUCV20171015, Agregó según ERS05612017: Optional ByVal pnTipoFlujoCaja As Integer = 0

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosCuotasIfis '" & psCtaCod & "', " & pnCodForm & ", " & pnConsCod & ", " & pnConsValor & ", " & pnTipoFlujoCaja & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosIfiCuota = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosCredEvalFlujoCaja(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosCredFormEvalFlujoCaja '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalFlujoCaja = rs
    Set rs = Nothing
End Function


Public Function RecuperaDatosCredEvalPDT(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosCredFormEvalDeclaracionPdt '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalPDT = rs
    Set rs = Nothing
End Function

Public Function RecuperaDatosCredEvalPDTDet(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperaDatosCredFormEvalDeclaracionPdtDet '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosCredEvalPDTDet = rs
    Set rs = Nothing
End Function
'LUCV20160525: Fin(FlexEdit) *************************************************************************

'***** LUCV20160524 - Para verificar si el formato fue verficado
Public Function RecuperaColocacCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaColocacCred
    sSql = "Select * From ColocacCred Where cCtaCod = '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocacCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'LUCV20160622 ->**********
Public Sub GrabarCredFormEvalCabeceraFormato1_5(ByVal psCtaCod As String, _
                                        ByVal pnInsertActualiza As Integer, _
                                        ByVal psGiroNeg As String, _
                                        ByVal pnExpEmpAnio As Integer, _
                                        ByVal pnExpEmpMes As Integer, _
                                        ByVal pnTiempoLocalAnio As Integer, _
                                        ByVal pnTiempoLocalMes As Integer, _
                                        ByVal pnUltEndeuda As Double, _
                                        ByVal psFecUltEndeuda As String, _
                                        ByVal pnCondLocal As Integer, _
                                        ByVal psCondLocalOtros As String, _
                                        ByVal pnExposicion As Double, _
                                        ByVal psFecEvaluacion As String, _
                                        ByVal pnFormato As Integer, _
                                        ByVal psComentario As String, _
                                        Optional ByVal pnInversionFlujo As Double = 0, _
                                        Optional ByVal pnFinanciamientoFlujo As Double = 0)
                                        'LUCV20171015 Agregó: pnInversionFlujo, pnFinanciamientoFlujo
                                        'ByVal pnIngreso As Double, _
                                        'ByVal pnEgreso As Double, _
                                        'ByVal pnMargen As Double, _

    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    If pnInsertActualiza = 1 Then
        lsSQL = "exec stp_ins_GrabarCredFormEvalFormato1_6 '"
    Else
        lsSQL = "exec stp_upd_GrabarCredFormEvalFormato1_6 '"
    End If

    lsSQL = lsSQL & psCtaCod & "', " & pnFormato & ", '" & psGiroNeg & "'," & _
                    pnExpEmpAnio & "," & pnExpEmpMes & "," & pnTiempoLocalAnio & "," & _
                    pnTiempoLocalMes & "," & pnCondLocal & ",'" & psCondLocalOtros & "'," & _
                    pnExposicion & ",'" & psFecEvaluacion & "'," & pnUltEndeuda & ",'" & _
                    psFecUltEndeuda & "', '" & psComentario & "', " & pnInversionFlujo & ", " & pnFinanciamientoFlujo & ""
    
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub GrabarCredFormEvalPropuestaFormato1_5(ByVal psCtaCod As String, _
                                        ByVal pnInsertActualiza As Integer, _
                                        ByVal psFechaVisita As String, _
                                        ByVal psEntorFami As String, _
                                        ByVal psGiroUbi As String, _
                                        ByVal psExperiencia As String, _
                                        ByVal psFormalidad As String, _
                                        ByVal psColaterales As String, _
                                        ByVal psDestino As String, _
                                        ByVal pnFormato As Integer)
                                                                                            
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    If pnInsertActualiza = 1 Then
        lsSQL = "exec stp_ins_CredFormEvalPropuestaFormato1_5 '"
    Else
        lsSQL = "exec stp_upd_CredFormEvalPropuestaFormato1_5 '"
    End If

    lsSQL = lsSQL & psCtaCod & "', " & pnFormato & ", '" & psFechaVisita & "','" & _
                    psEntorFami & "','" & psGiroUbi & "','" & psExperiencia & "','" & _
                    psFormalidad & "','" & psColaterales & "','" & psDestino & "'"

    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'Fin LUCV20160622 <-**********

Public Sub GrabarCredFormEvalRatioIndicador1_5(ByVal psCtaCod As String, _
                                                    ByVal pnInsertActualiza As Integer, _
                                                    ByVal pnCodCapaPagoNeta As Integer, _
                                                    ByVal pnCapPagoNeto As Double, _
                                                    ByVal pnCodEndeuPat As Integer, _
                                                    ByVal pnEndeuPat As Double, _
                                                    ByVal pnCodIngrNeto As Integer, _
                                                    ByVal pnIngreNetoNego As Double, _
                                                    ByVal pnCodExecMens As Integer, _
                                                    ByVal pnExceMensual As Double, _
                                                    ByVal pnFormato As Integer)
                                                                                                        
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    If pnInsertActualiza = 1 Then
        lsSQL = "exec stp_ins_CredFormEvalRatioIndicador1_6 '"
    Else
        lsSQL = "exec stp_upd_CredFormEvalRatioIndicador1_6 '"
    End If

    lsSQL = lsSQL & psCtaCod & "', " & pnFormato & ", " & pnCodCapaPagoNeta & "," & pnCodEndeuPat & "," & pnCodIngrNeto & "," & _
                    pnCodExecMens & "," & pnCapPagoNeto & ", " & pnEndeuPat & ", " & pnIngreNetoNego & ", " & pnExceMensual & ""

    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

'LUCV20160622 ->**********
Public Sub InsertaCredFormEvalGastoNego(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnCodGasto As Double, ByVal pnMonto As Currency)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_CredFormEvalGastoNego '" & psCtaCod & "', " & pnFormato & ", " & pnCodGasto & ", " & pnMonto & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Sub InsertaCredFormEvalGastoFami(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnCodGasto As Double, ByVal pnMonto As Currency)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_CredFormEvalGastoFami '" & psCtaCod & "', " & pnFormato & ", " & pnCodGasto & ", " & pnMonto & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Sub InsertaCredFormEvalOtrosIngr(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnCodGasto As Double, ByVal pnMonto As Currency)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_CredFormEvalOtrosIngr '" & psCtaCod & "', " & pnFormato & ", " & _
                                                    pnCodGasto & ", " & pnMonto & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
 
 Public Sub RecuperaValorConstante(ByVal pnConsCod As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_CredFormEvalActiPasi " & pnConsCod & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub InsertaCredFormEvalActiPasi(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal psFechaReg As String, _
                                      ByVal nConsCod As Integer, ByVal nConsvalor As Integer, ByVal pnMonto As Currency, _
                                      Optional ByVal pnPP As Currency = 0, Optional ByVal pnPE As Currency = 0)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_CredFormEvalActiPasi '" & psCtaCod & "', " & pnFormato & ", '" & psFechaReg & "'," & _
                                                    nConsCod & ", " & nConsvalor & "," & pnMonto & ", " & pnPP & ", " & pnPE & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub InsertaCredFormEvalActiPasiDet(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal psFechaReg As String, _
                                      ByVal nConsCod As Integer, ByVal nConsvalor As Integer, ByVal pnMonto As Currency, _
                                      Optional ByVal pnPP As Currency = 0, Optional ByVal pnPE As Currency = 0)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_CredFormEvalActiPasiDet '" & psCtaCod & "', " & pnFormato & ", '" & psFechaReg & "'," & _
                                                   nConsCod & ", " & nConsvalor & "," & pnMonto & ", " & pnPP & ", " & pnPE & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
    
    
End Sub

Public Sub InsertaCredFormEvalCuotaIFI(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnConsCod As Integer, _
                                       ByVal pnConsValor As Integer, ByVal pnNroCuota As Integer, ByVal psDescripcion As String, ByVal pnMonto As Currency, Optional ByVal pnTipoFlujoCaja As Integer = 0)
    'LUCV20171015, Agregó: Optional ByVal pnTipoFlujoCaja As Integer = -1
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_CredFormEvalCuotaIfi '" & psCtaCod & "', " & pnFormato & ", " & pnConsCod & " , " & pnConsValor & " ," & _
                                                pnNroCuota & ",'" & psDescripcion & "', " & pnMonto & ", " & pnTipoFlujoCaja & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub InsertaCredFormEvalActiPasiCtaCobrar(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnTipoPatrimonio As Integer, ByVal pnNroCtaCobrar As Integer, ByVal pnConsCod As Integer, _
                                              ByVal pnConsValor As Integer, ByVal psFechaInventario As String, ByVal psDescripcion As String, ByVal pnMonto As Currency)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
       
    lsSQL = "exec stp_ins_CredFormEvalActiPasiCtaCobrar '" & psCtaCod & "', " & pnFormato & ", " & pnConsCod & " , " & pnConsValor & ", " & pnTipoPatrimonio & _
                                                        "," & pnNroCtaCobrar & ", '" & psFechaInventario & "','" & psDescripcion & "', " & pnMonto & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub InsertaCredFormEvalActiPasiInventario(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnTipoPatrimonio As Integer, ByVal pnNroInventario As Integer, ByVal pnConsCod As Integer, _
                                              ByVal pnConsValor As Integer, ByVal psFechaInventario As String, ByVal psDescripcion As String, _
                                              ByVal pnCantidad As Double, ByVal pnUnidadMed As String, ByVal pnCostoUnit As Double, ByVal pnMonto As Currency)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
                                                            
    lsSQL = "exec stp_ins_CredFormEvalActiPasiInventario '" & psCtaCod & "', " & pnFormato & ", " & pnConsCod & " , " & pnConsValor & _
                                                          " , " & pnTipoPatrimonio & ", " & pnNroInventario & ", '" & psDescripcion & "'," & _
                                                            pnCantidad & ", '" & pnUnidadMed & "' ," & pnCostoUnit & ", " & pnMonto & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub


Public Sub InsertaCredFormEvalActiPasiActivoFijo(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnNroActiFijo As Integer, ByVal pnTipoPatrimonio As Integer, ByVal pnConsCod As Integer, _
                                                ByVal pnConsValor As Integer, ByVal psDescripcion As String, ByVal pnCantidad As Double, ByVal pnPrecio As Double, ByVal pnTotal As Currency)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
           
    lsSQL = "exec stp_ins_CredFormEvalActiPasiActFijo '" & psCtaCod & "', " & pnFormato & "," & pnConsCod & " , " & pnConsValor & ", " & _
                                    pnTipoPatrimonio & "," & pnNroActiFijo & ", '" & psDescripcion & "'," & pnCantidad & ", " & pnPrecio & ", " & pnTotal & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub InsertaCredFormEvalActiPasiSobregiros(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnTipoPatrimonio As Integer, _
                                                ByVal pnNroActiFijo As Integer, ByVal psEntidad As String, ByVal psCredito As String, ByVal pnTEA As Double, ByVal pnCuotas As Integer, _
                                                ByVal pnMontoCuota As Double, pnCuotaPend As Double)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
           
    lsSQL = "exec stp_ins_CredFormEvalActiPasiSobregiros '" & psCtaCod & "', " & pnFormato & "," & pnConsCod & " , " & pnConsValor & ", " & _
                                    pnTipoPatrimonio & "," & pnNroActiFijo & ", '" & psEntidad & "','" & psCredito & "', " & pnTEA & ", " & pnCuotas & ", " & pnMontoCuota & ", " & pnCuotaPend & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub InsertaCredFormEvalReferidos(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnNroRef As Integer, _
                                        ByVal psNombres As String, ByVal nDNI As String, _
                                        ByVal nTelefono As String, ByVal psReferido As String, ByVal nNroDNI As String)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_CredFormEvalReferidos '" & psCtaCod & "', " & pnFormato & _
                                                 ", " & pnNroRef & ", '" & psNombres & "', '" & nDNI & _
                                                 "', '" & nTelefono & "', '" & psReferido & "', '" & nNroDNI & "'"
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub InsertaCredFormFlujoCajaMensual(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnConsCod As Integer, _
                                       ByVal pnConsValor As Integer, ByVal psConcepto As String, ByVal pnMonto As Currency, Optional ByVal pnTipoFlujoCaja As Integer = 0)
    'LUCV20171015, Agregó según ERS0512017: pnTipoFlujoCaja
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    
    lsSQL = "exec stp_ins_CredFormEvalFlujoCajaMensual '" & psCtaCod & "', " & pnFormato & ", " & pnConsCod & " , " & _
                                                        pnConsValor & " ,'" & psConcepto & "', " & pnMonto & ", " & pnTipoFlujoCaja & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub InsertaCredFormEvalPDT(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnConsCod As Integer, _
                                       ByVal pnConsValor As Integer, ByVal pnMes1 As Integer, ByVal pnMes2 As Integer, ByVal pnMes3 As Integer, _
                                       ByVal pnAnio1 As Integer, ByVal pnAnio2 As Integer, ByVal pnAnio3 As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_CredFormEvalDeclaracionPdt '" & psCtaCod & "', " & pnFormato & ", " & pnConsCod & " , " & _
                                                        pnConsValor & " ," & pnMes1 & ", " & pnMes2 & ", " & pnMes3 & ", " & pnAnio1 & ", " & pnAnio2 & ", " & pnAnio3 & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub InsertaCredFormEvalPDTDet(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnConsCod As Integer, _
                                       ByVal pnConsValor As Integer, ByVal pnMontoMes1 As Long, ByVal pnMontoMes2 As Long, ByVal pnMontoMes3 As Long, _
                                       ByVal pnPromedio As Long, ByVal pnVentasDecl As Double)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_CredFormEvalDeclaracionPdtDet '" & psCtaCod & "', " & pnFormato & ", " & pnConsCod & " , " & pnConsValor & " ," & _
                                                        pnMontoMes1 & ", " & pnMontoMes2 & ", " & pnMontoMes3 & ", " & pnPromedio & ", " & pnVentasDecl & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'LUCV20160622, Fin<-**********


'LUCV20160622 ->**********
Public Sub EliminaCredFormEvalGrillaGastoNegocio(ByVal psCtaCod As String, ByVal pnFormato As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_del_CredFormEvalGastoNego '" & psCtaCod & "', " & pnFormato & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Sub EliminaCredFormEvalGrillaGastoFami(ByVal psCtaCod As String, ByVal pnFormato As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_del_CredFormEvalGastoFami '" & psCtaCod & "', " & pnFormato & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Sub EliminaCredFormEvalGrillaOtrosIngre(ByVal psCtaCod As String, ByVal pnFormato As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_del_CredFormEvalOtrosIngr '" & psCtaCod & "', " & pnFormato & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub EliminaCredFormEvalGrillaActiPasi(ByVal psCtaCod As String, ByVal pnFormato As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_del_CredFormEvalActiPasi'" & psCtaCod & "', " & pnFormato & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub EliminaCredFormEvalGrillaActiPasiDet(ByVal psCtaCod As String, ByVal pnFormato As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_del_CredFormEvalActiPasiDet '" & psCtaCod & "', " & pnFormato & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub EliminaCredFormEvalGrillaOtrasIfis(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnConsCod As Integer, Optional ByVal pnConsValor As Integer = 0, Optional ByVal pnTipoFlujoCaja As Integer = 0)
'JOEP20171015, Agregó según ERS0512017: Optional ByVal pnConsValor As Integer = -1
'LUCV20171015, Agregó según ERS0512017: Optional ByVal pnTipoFlujoCaja As Integer = 0
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_del_CredFormEvalOtrasIfis '" & psCtaCod & "', " & pnFormato & ", " & pnConsCod & "," & pnConsValor & "," & pnTipoFlujoCaja & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub EliminaCredFormEvalGrillaReferidos(ByVal psCtaCod As String, ByVal pnFormato As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_del_CredFormEvalReferidos '" & psCtaCod & "', " & pnFormato & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub EliminaCredFormEvalFlujoCaja(ByVal psCtaCod As String, ByVal pnFormato As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_del_CredFormEvalFlujoCajaMensual '" & psCtaCod & "', " & pnFormato & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub EliminaCredFormEvalVentasCostos(ByVal psCtaCod As String, ByVal pnFormato As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
     lsSQL = "exec stp_del_CredFormEvalVentasCosto '" & psCtaCod & "', " & pnFormato & " "
     
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub EliminaCredFormEvalPDT(ByVal psCtaCod As String, ByVal pnFormato As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
     lsSQL = "exec stp_del_EliminaCredFormEvalPDT '" & psCtaCod & "', " & pnFormato & " "
     
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub EliminaCredFormEvalPDTDet(ByVal psCtaCod As String, ByVal pnFormato As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
     lsSQL = "exec stp_del_EliminaCredFormEvalPDTDet '" & psCtaCod & "', " & pnFormato & " "
     
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'LUCV20160622 <-**********

'LUCV20160623 ->**********
Public Function RecuperarDatosInformeVisitaFormato1_6(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec spt_sel_RecuperarDatosInformeVisitaFormato1_6 '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosInformeVisitaFormato1_6 = rs
    Set rs = Nothing
End Function
'LUCV20160623 <-**********

''JOEP20160530: Cargar Datos
''----------------------------------------------------------------------------------------------
'Public Sub ActualziarfrmCredFormEvalFormatoNuevaLineaAutomatica(ByVal cNumLineaAut As String, ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal cPersCodLinea As String, _
'                                                            ByVal cDescripcion As String, ByVal dFecha As String, ByVal nMontoMax As Double, _
'                                                            ByVal nCuotaMax As Double, ByVal nPlazoMax As Integer, ByVal nMoneda As Integer, _
'                                                            ByVal nTasaInt As Double, ByVal cPersCodAna As String, ByVal nMontoRiesgoUnico As Double, ByVal cObservacion As String, ByVal dEstado As Date, ByVal nEstado As Integer)
'
'Dim lsSQL As String
'    lsSQL = "exec stp_upd_ActualizarNuevaLineaAutomatica '"
'    lsSQL = lsSQL & cNumLineaAut & "','" & cCtaCod & "'," & nCodForm & ",'" & cPersCodLinea & "','" & cDescripcion & "','" & _
'                    dFecha & "'," & nMontoMax & "," & nCuotaMax & "," & nPlazoMax & "," & _
'                    nMoneda & "," & nTasaInt & ",'" & cPersCodAna & "'," & nMontoRiesgoUnico & ",'" & cObservacion & "','" & dEstado & "'," & nEstado & ""
'
'    coConex.Ejecutar lsSQL
'
'End Sub
'
''JOEP20160530
'Public Sub ActualziarfrmCredFormEvalFormatoNuevaLineaAutomatica_PersonaFirma(ByVal cNumLineaAut As String, cCtaCod As String, nCodForm As Integer, cPersCod As String)
'Dim lsSQL As String
'        lsSQL = "exec stp_upd_ActualizarNuevaLineaAutomaticaPersonaFirma '"
'        lsSQL = lsSQL & cNumLineaAut & "','" & cCtaCod & "'," & nCodForm & ",'" & cPersCod & "'"
'
'            coConex.Ejecutar lsSQL
'End Sub
'
''JOEP20160530
'Public Sub ActualziarfrmCredFormEvalFormatoNuevaLineaAutomatica_Garantias(ByVal cNumLineaAut As String, cCtaCod As String, nCodForm As Integer, cNumGarant As String)
'Dim lsSQL As String
'        lsSQL = "exec stp_upd_ActualizarNuevaLineaAutomaticaGarantias '"
'        lsSQL = lsSQL & cNumLineaAut & "','" & cCtaCod & "'," & nCodForm & ",'" & cNumGarant & "'"
'
'            coConex.Ejecutar lsSQL
'End Sub
''para Guardar Datos de la Cabecera(Linea Automatica)
'
''JOEP20160530
'Public Sub GrabarfrmCredFormEvalFormatoNuevaLineaAutomatica(ByVal cNumLineaAut As String, ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal cPersCodLinea As String, _
'                                                            ByVal cDescripcion As String, ByVal dFecha As String, ByVal nMontoMax As Double, _
'                                                            ByVal nCuotaMax As Double, ByVal nPlazoMax As Integer, ByVal nMoneda As Integer, _
'                                                            ByVal nTasaInt As Double, ByVal cPersCodAna As String, ByVal nMontoRU, ByVal cObservacion As String, ByVal dEstado As Date, ByVal nEstado As Integer)
'
'Dim rs As New ADODB.Recordset
'Dim lsSQL As String
'    lsSQL = "exec stp_ins_CredFormEvalNuevaLineaCredAutomatica '"
'    lsSQL = lsSQL & cNumLineaAut & "','" & cCtaCod & "','" & nCodForm & "','" & cPersCodLinea & "','" & cDescripcion & "','" & _
'                    dFecha & "','" & nMontoMax & "','" & nCuotaMax & "','" & nPlazoMax & "','" & _
'                    nMoneda & "','" & nTasaInt & "','" & cPersCodAna & "'," & nMontoRU & ",'" & cObservacion & "','" & dEstado & "'," & nEstado & ""
'
'    coConex.Ejecutar lsSQL
'
'End Sub
'
''FIN para Guardar Datos de la Cabecera(Linea Automatica)
''JOEP20160530
'Public Sub GrabarfrmCredFormEvalFormatoNuevaLineaAutomaticaEStado(ByVal cNumLineaAut As String, ByVal nMontoMax As Double, ByVal dEstado As Date, _
'                                                                  ByVal nEstado As Integer, ByVal cPersCodLinea As String)
'
''Dim rs As New ADODB.Recordset
'Dim lsSQL As String
'    lsSQL = "exec stp_ins_CredFormEvalLineaCredAutoEstado '"
'    lsSQL = lsSQL & cNumLineaAut & "'," & nMontoMax & ",'" & dEstado & "'," & nEstado & ",'" & cPersCodLinea & "'"
'
'    coConex.Ejecutar lsSQL
'End Sub
'
''JOEP20160530
'Public Sub GrabarfrmCredFormEvalFormatoNuevaLineaAutomatica_PersonaFirma(cNumLineaAut As String, cCtaCod As String, nCodForm As Integer, cPersCod As String)
'
'    Dim lsSQL As String
'        lsSQL = "exec stp_ins_CredFormEvalLineaCredAutoPersonaFirmar '"
'        lsSQL = lsSQL & cNumLineaAut & "','" & cCtaCod & "'," & nCodForm & ",'" & cPersCod & "'"
'
'            coConex.Ejecutar lsSQL
'
'End Sub
'
''JOEP20160530
'Public Sub GrabarfrmCredFormEvalFormatoNuevaLineaAutomatica_Garantias(cNumLineaAut As String, cCtaCod As String, nCodForm As Integer, cNumGarant As String)
'
'    Dim lsSQL As String
'        lsSQL = "exec stp_ins_CredFormEvalLineaCredAutoGarantias '"
'        lsSQL = lsSQL & cNumLineaAut & "','" & cCtaCod & "'," & nCodForm & ",'" & cNumGarant & "'"
'
'            coConex.Ejecutar lsSQL
'
'End Sub
'
''JOEP20160530
'Public Sub ActualizarfrmCredFormEvalFormatoLineaCredAutomatica(ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal dFechaExpeCaja As String, _
'                                                           ByVal cLinea As String)
'Dim lsSQL As String
'        lsSQL = "exec stp_upd_ActualizarFormatoLineaAutomatica'"
'        lsSQL = lsSQL & cCtaCod & "'," & nCodForm & ",'" & dFechaExpeCaja & "','" & cLinea & "'"
'
'        coConex.Ejecutar lsSQL
'End Sub
'
''JOEP20160530
'Public Sub GrabarfrmCredFormEvalFormatoLineaCredAutomatica(ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal cActividad As String, ByVal nExposiCred As Double, _
'                                                           ByVal nUltEndeSBS As Double, ByVal dUltEndeuSBS As String, ByVal dFechaExpeCaja As String, _
'                                                           ByVal cLinea As String)
'    Dim lsSQL As String
'        lsSQL = "exec stp_ins_CredFormEvalFormatoLineaCredAutomatica'"
'        lsSQL = lsSQL & cCtaCod & "'," & nCodForm & ",'" & cActividad & "'," & nExposiCred & "," & nUltEndeSBS & ",'" & dUltEndeuSBS & "','" & dFechaExpeCaja & "','" & cLinea & "'"
'
'        coConex.Ejecutar lsSQL
'End Sub
'
''JOEP20160530
'Public Sub GrabarfrmCredFormEvalFormatoParalelo(ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal cActividad As String, _
'                                                ByVal dFechaExpeCaja As String, ByVal nUltEndeSBS As Double, ByVal nNCreditos As Integer, _
'                                                ByVal dUltEndeuSBS As String, ByVal cCampaña As String, ByVal nExposiCred As Double)
'    Dim oConecta As COMConecta.DCOMConecta
'    Dim lsSQL As String
'
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'
'
'        lsSQL = "exec stp_ins_CredFormEvalFormatoParaleloInfCliente'"
'        lsSQL = lsSQL & cCtaCod & "'," & nCodForm & ",'" & cActividad & "','" & dFechaExpeCaja & "'," & nUltEndeSBS & _
'                        "," & nNCreditos & ",'" & dUltEndeuSBS & "','" & cCampaña & "'," & nExposiCred & ""
'
'        oConecta.Ejecutar lsSQL
'End Sub

'Eliminar JOEP 20160726
Public Sub EliminaCredFormEvalPropuesta(ByVal psCtaCod As String, ByVal pnFormato As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_del_EliminarPropuesta '" & psCtaCod & "', " & pnFormato & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub EliminaCredFormEvalBoletaPago(ByVal psCtaCod As String, ByVal pnFormato As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_del_EliminarBoletaPagoSinConvenio '" & psCtaCod & "', " & pnFormato & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub EliminaCredFormEvalReciboHonorarios(ByVal psCtaCod As String, ByVal pnFormato As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_del_EliminarReciboHonorariosSinConvenio '" & psCtaCod & "', " & pnFormato & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Sub EliminaCredFormEvalNegocio(ByVal psCtaCod As String, ByVal pnFormato As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_sel_ERS004_2016_EliminarNegocio '" & psCtaCod & "', " & pnFormato & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Eliminar

'JOEP 20160811
Public Function ObtenerCapPagoConConvenio(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ObtenerCapPagoConConvenio '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set ObtenerCapPagoConConvenio = rs
    
    Set rs = Nothing
End Function
'FIN JOEP 20160811

'JOEP20160530
Public Sub GrabarfrmCredFormEvalFormatoParalelo(ByVal cCtaCod As String, ByVal ncodForm As Integer, ByVal cActividad As String, _
                                                ByVal dFechaExpeCaja As String, ByVal nUltEndeSBS As Double, ByVal nNCreditos As Integer, _
                                                ByVal dUltEndeuSBS As String, ByVal cCampaña As String, ByVal nExposiCred As Double)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
                                                                        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    
        lsSQL = "exec stp_ins_CredFormEvalFormatoParaleloInfCliente'"
        lsSQL = lsSQL & cCtaCod & "'," & ncodForm & ",'" & cActividad & "','" & dFechaExpeCaja & "'," & nUltEndeSBS & _
                        "," & nNCreditos & ",'" & dUltEndeuSBS & "','" & cCampaña & "'," & nExposiCred & ""
        
        oConecta.Ejecutar lsSQL
End Sub


'JOEP20160530
Public Sub dGrabarfrmCredFormEvalFormatoParaleloEvaluacion(ByVal cCtaCod As String, ByVal ncodForm As Integer, ByVal nMontoApro As Double, ByVal nSaldoActual As Double, ByVal nImporte As Double, ByVal nCapPago As Double, ByVal nIngresoNeto As Double, _
                                                            ByVal nIncreIngresoDatos As Double, _
                                                            ByVal nIngreEstMontos As Double, ByVal nIncreIngresoEstiMontos As Double, ByVal nMargenbruto As Double, ByVal nOtrsoGastos As Double, ByVal nConsuFamili As Double, ByVal nCuotaCredVig As Double, ByVal nOtrosIng As Double, _
                                                            ByVal nMargenBrutoCaja As Double, ByVal nIngreResumen As Double, ByVal nIncreIngreResumen As Double, ByVal nMontoParalelo As Double, ByVal nMontoPropuesto As Double)
   Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
                                                                        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
  
        lsSQL = "exec stp_ins_CredFormEvalFormatoParaleloEvaluacion'"
        lsSQL = lsSQL & cCtaCod & "'," & ncodForm & "," & nMontoApro & "," & nSaldoActual & "," & nImporte & "," & nCapPago & "," & nIngresoNeto & "," _
                        & nIncreIngresoDatos & "," & nIngreEstMontos & "," & nIncreIngresoEstiMontos & "," & nMargenbruto & "," & nOtrsoGastos & "," & nConsuFamili & "," & nCuotaCredVig & "," & nOtrosIng & "," & nMargenBrutoCaja & "," & nIngreResumen & "," & nIncreIngreResumen & "," & nMontoParalelo & "," & nMontoPropuesto & ""

            oConecta.Ejecutar lsSQL
End Sub

'JOEP20160530
Public Sub GrabarfrmCredFormEvalFormatoParaleloPropuestaCredito(ByVal cCtaCod As String, ByVal nCodForma As Integer, ByVal dFecVisita As String, ByVal cDestino As String, ByVal cEntornoFami As String, ByVal cGiroubicacion As String, ByVal cExpeCrediticia As String, ByVal cFormalidaNegocio As String, ByVal cGarantias As String)
      
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
                                                                        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
        lsSQL = "exec stp_ins_CredFormEvalFormatoParaleloPropuesta'"
        lsSQL = lsSQL & cCtaCod & "'," & nCodForma & ",'" & dFecVisita & "','" & cEntornoFami & "','" & cExpeCrediticia & "','" & cGarantias & "','" & cDestino & "','" & cGiroubicacion & "','" & cFormalidaNegocio & "'"
                                                                                
    oConecta.Ejecutar lsSQL
End Sub

'JOEP20160530
Public Sub ActualizarfrmCredFormFormatoParalelo(ByVal psCtaCod As String, ByVal ncodForm As Integer, ByVal dFechaExpeCaja As String)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
                                                                        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
        lsSQL = "exec stp_upd_ActualizarFormatoParaleloInfromacionNegocio'"
        lsSQL = lsSQL & psCtaCod & "'," & ncodForm & ",'" & Format(dFechaExpeCaja, "yyyyMMdd") & "'"
        
    oConecta.Ejecutar lsSQL
    
End Sub

'JOEP20160530
Public Sub ActualizarfrmCredFormFormatoParaleloEvaluacion(ByVal cCtaCod As String, ByVal ncodForm As Integer, ByVal nMontoApro As Double, ByVal nSaldoActual As Double, ByVal nVentas As Double, _
ByVal nCapPago As Double, ByVal nIngresoNeto As Double, ByVal nIncreIngresoDatos As Double, ByVal nIngreEstMontos As Double, ByVal nIncreIngresoEstiMontos As Double, _
ByVal nMargenbruto As Double, ByVal nOtrsoGastos As Double, ByVal nConsuFamili As Double, ByVal nCuotaCredVig As Double, ByVal nOtrosIng As Double, ByVal nMargenBrutoCaja As Double, _
ByVal nIngreResumen As Double, ByVal nIncreIngreResumen As Double, ByVal nMontoParalelo As Double, ByVal nMontoPropuesto As Double)
   
   Dim oConecta As COMConecta.DCOMConecta
   Dim lsSQL As String
                                                                        
   Set oConecta = New COMConecta.DCOMConecta
   oConecta.AbreConexion
    
        lsSQL = "exec stp_upd_ActualizarFormatoParaleloEvaluacion'"
        lsSQL = lsSQL & cCtaCod & "'," & ncodForm & "," & nMontoApro & "," & nSaldoActual & "," & nVentas & "," & nCapPago & "," & nIngresoNeto & "," & nIncreIngresoDatos & "," & nIngreEstMontos & "," _
        & nIncreIngresoEstiMontos & "," & nMargenbruto & "," & nOtrsoGastos & "," & nConsuFamili & "," & nCuotaCredVig & "," & nOtrosIng & "," & nMargenBrutoCaja & "," & nIngreResumen & "," _
        & nIncreIngreResumen & "," & nMontoParalelo & "," & nMontoPropuesto & ""
        
   oConecta.Ejecutar lsSQL
End Sub

'JOEP20160530
Public Sub ActualizarfrmCredFormFormatoParaleloPropuestaCredito(ByVal cCtaCod As String, ByVal nCodForma As Integer, ByVal dFecVisita As Date, ByVal cDestino As String, ByVal cEntornoFami As String, ByVal cGiroubicacion As String, ByVal cExpeCrediticia As String, ByVal cFormalidaNegocio As String, ByVal cGarantias As String)
    
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
                                                                        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
        lsSQL = "exec stp_upd_ActualizarFormatoParaleloPropuestaCredito'"
        lsSQL = lsSQL & cCtaCod & "'," & nCodForma & ",'" & Format(dFecVisita, "yyyyMMdd") & "','" & cEntornoFami & "','" & cExpeCrediticia & "','" & cGarantias & "','" & cDestino & "','" & cGiroubicacion & "','" & cFormalidaNegocio & "'"
        
        oConecta.Ejecutar lsSQL
End Sub
'------------------------------------------------------------------------------------------------

'Formato con Convenio Recuperar Datos Total CAbecera
'JOEP20160530
Public Function RecuperarDatosTotalConsumoConConvenio(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalConConvenio '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoConConvenio = rs
    
    Set rs = Nothing
End Function
'FIN Formato con Convenio Recuperar Datos Total CAbecera

'Formato con Convenio Recuperar Datos Total Evaluacion Meses
'JOEP20160530
Public Function RecuperarDatosTotalConsumoConConvenioEvalMeses(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalConConvenioEvalMeses '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoConConvenioEvalMeses = rs
    
    Set rs = Nothing
End Function
'FIN Formato con Convenio Recuperar Datos Total Evaluacion Meses

'Formato con Convenio Recuperar Datos Total Promedios
'JOEP20160530
Public Function RecuperarDatosTotalConsumoConConvenioPromedios(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalConConvenioPromedios '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoConConvenioPromedios = rs
    
    Set rs = Nothing
End Function
'FIN Formato con Convenio Recuperar Datos Total Promedios

'Formato con Convenio Recuperar Datos Total Referidos
'JOEP20160530
Public Function RecuperarDatosTotalConsumoConConvenioReferidos(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalConConvenioReferidos '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoConConvenioReferidos = rs
    
    Set rs = Nothing
End Function
'FIN Formato con Convenio Recuperar Datos Total Referidos

'Formato con Convenio Recuperar Datos Total Remu Bruta Total 1
'JOEP20160530
Public Function RecuperarDatosTotalConsumoConConvenioRemuBrutaTota_1(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalConConvenioRemBrutaTotal_1 '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoConConvenioRemuBrutaTota_1 = rs
    
    Set rs = Nothing
End Function
'FIN Formato con Convenio Recuperar Datos Total Remu Bruta Total 1

'Formato con Convenio Recuperar Datos Total Remu Bruta Total 2
'JOEP20160530
Public Function RecuperarDatosTotalConsumoConConvenioRemuBrutaTota_2(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalConConvenioRemBrutaTotal_2 '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoConConvenioRemuBrutaTota_2 = rs
    
    Set rs = Nothing
End Function
'FIN Formato con Convenio Recuperar Datos Total Remu Bruta Total 2

'Formato con Convenio Recuperar Datos Total Descuento 1
'JOEP20160530
Public Function RecuperarDatosTotalConsumoConConvenioDescuento_1(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalConConvenioDescuento_1 '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoConConvenioDescuento_1 = rs
    
    Set rs = Nothing
End Function
'FIN Formato con Convenio Recuperar Datos Total Descuento 1

'Formato con Convenio Recuperar Datos Total Descuento 2
'JOEP20160530
Public Function RecuperarDatosTotalConsumoConConvenioDescuento_2(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalConConvenioDescuento_2 '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoConConvenioDescuento_2 = rs
    
    Set rs = Nothing
End Function
'FIN Formato con Convenio Recuperar Datos Total Descuento 2

'Formato Con Convenio
'JOEP20160530
Public Function RecuperarDatosFormatoConConvenio(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarFormatoConConvenio'" & pcCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosFormatoConConvenio = rs
   
    Set rs = Nothing
End Function
'FIN Formato Con Convenio

'Mostrar Informe de Visita Con Convenio
'JOEP20160530
Public Function MostrarDatosInformeVisitaFormatoConConvenio(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec spt_sel_MostraInfromeVisitaConsumoConConvenio'" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarDatosInformeVisitaFormatoConConvenio = rs
    
    Set rs = Nothing
End Function
'FIN Mostrar Informe de Visita Con Convenio

'Guardar Formato Con Convenio
'JOEP20160530
Public Sub GuardarFormatoConConvenio_InfCliente(ByVal cCtaCod As String, ByVal ncodForm As Integer, ByVal cActividad As String, ByVal nAntgAnios As Integer, _
                                                ByVal nAntgMes As Integer, ByVal nUltEndeSBS As Double, ByVal dUltEndeuSBS As Date, ByVal nTipoAportacion As Integer, ByVal nTipoInstitucion As Integer, ByVal nSectorSalud As Integer, ByVal nTipoPlanilla As Integer, ByVal cInstiConv As String, ByVal nMontoSol As Double, _
                                                ByVal nNumCuotas As Integer, ByVal nExposiCred As Double, ByVal dFecFteIng As Date, ByVal cComentario As String)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
                                                                        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    lsSQL = "exec  stp_ins_ConsumoConCovenio_InfCliente '"
    lsSQL = lsSQL & cCtaCod & "'," & ncodForm & ",'" & cActividad & "'," & nAntgAnios & "," & nAntgMes & "," & _
                    nUltEndeSBS & ",'" & Format(dUltEndeuSBS, "yyyyMMdd") & "'," & nTipoAportacion & "," & nTipoInstitucion & "," & nSectorSalud & "," & nTipoPlanilla & ",'" & cInstiConv & "'," & nMontoSol & "," & nNumCuotas & "," & _
                    nExposiCred & ",'" & Format(dFecFteIng, "yyyyMMdd") & "','" & cComentario & "'"
    
    oConecta.Ejecutar lsSQL

End Sub
'FIN Guardar Formato Con Convenio

'Actualizar Formato Con Convenio
'JOEP20160530
Public Sub ActualizarFormatoConConvenio_InfCliente(ByVal cCtaCod As String, ByVal ncodForm As Integer, ByVal cActividad As String, ByVal nAntgAnios As Integer, _
                                                ByVal nAntgMes As Integer, ByVal nUltEndeSBS As Double, ByVal dUltEndeuSBS As Date, ByVal nTipoAportacion As Integer, ByVal nTipoInstitucion As Integer, ByVal nSectorSalud As Integer, ByVal nTipoPlanilla As Integer, ByVal cInstiConv As String, ByVal nMontoSol As Double, _
                                                ByVal nNumCuotas As Integer, ByVal nExposiCred As Double, ByVal dFecFteIng As Date, ByVal cComentario As String)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
                                                                        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_upd_ActualizarFormatoConConvenioCabecera '"
    lsSQL = lsSQL & cCtaCod & "'," & ncodForm & ",'" & cActividad & "'," & nAntgAnios & "," & nAntgMes & "," & _
                    nUltEndeSBS & ",'" & Format(dUltEndeuSBS, "yyyyMMdd") & "'," & nTipoAportacion & "," & nTipoInstitucion & "," & nSectorSalud & "," & nTipoPlanilla & ",'" & cInstiConv & "'," & nMontoSol & "," & nNumCuotas & "," & _
                    nExposiCred & ",'" & Format(dFecFteIng, "yyyyMMdd") & "','" & cComentario & "'"
    
    oConecta.Ejecutar lsSQL

End Sub
'FIN Actualizar Formato Con Convenio

'Guardar Datos Formato Con Convenio EvalMes1
'JOEP20160530
Public Sub GuardarFormatoConConvenio_EvalMes1(ByVal cCtaCod As String, ByVal dFechaMes1 As Integer, ByVal nAnoMes1 As Integer, ByVal nRemBruTotalMes1 As Double, nDescMes1, ByVal nIngrNetoMes1 As Double, ByVal EvalMes1 As Integer)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_ins_ConsumoConCovenio_EvalMes1_2 '"
    lsSQL = lsSQL & cCtaCod & "'," & dFechaMes1 & "," & nAnoMes1 & "," & nRemBruTotalMes1 & "," & nDescMes1 & "," & _
                    nIngrNetoMes1 & "," & EvalMes1 & ""
    
    oConecta.Ejecutar lsSQL

End Sub
'FIN Guardar Datos Formato Con Convenio EvalMes1

'Actualizar Datos Formato Con Convenio EvalMes1
'JOEP20160530
Public Sub ActualizarFormatoConConvenio_EvalMes1(ByVal cCtaCod As String, ByVal dFechaMes1 As Integer, ByVal nAnoMes1 As Integer, ByVal nRemBruTotalMes1 As Double, nDescMes1, ByVal nIngrNetoMes1 As Double, ByVal EvalMes1 As Integer)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_upd_ActualizarFormatoConConvenioEvalMes1_2 '"
    lsSQL = lsSQL & cCtaCod & "'," & dFechaMes1 & "," & nAnoMes1 & "," & nRemBruTotalMes1 & "," & nDescMes1 & "," & _
                    nIngrNetoMes1 & "," & EvalMes1 & ""
    
    oConecta.Ejecutar lsSQL

End Sub
'FIN Actualizar Datos Formato Con Convenio EvalMes1

'Guardar Datos Formato Con Convenio EvalMes2
'JOEP20160530
Public Sub GuardarFormatoConConvenio_EvalMes2(ByVal cCtaCod As String, ByVal dFechaMes2 As Integer, ByVal nAnoMes2 As Integer, ByVal nRemBruTotalMes2 As Double, nDescMes2, ByVal nIngrNetoMes2 As Double, ByVal EvalMes2 As Integer)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_ins_ConsumoConCovenio_EvalMes1_2 '"
    lsSQL = lsSQL & cCtaCod & "'," & dFechaMes2 & "," & nAnoMes2 & "," & nRemBruTotalMes2 & "," & nDescMes2 & "," & _
                    nIngrNetoMes2 & "," & EvalMes2 & ""
    
    oConecta.Ejecutar lsSQL

End Sub
'FIN Guardar Datos Formato Con Convenio EvalMes2

'Actualizar Datos Formato Con Convenio EvalMes2
'JOEP20160530
Public Sub ActualizarFormatoConConvenio_EvalMes2(ByVal cCtaCod As String, ByVal dFechaMes2 As Integer, ByVal nAnoMes2 As Integer, ByVal nRemBruTotalMes2 As Double, nDescMes2, ByVal nIngrNetoMes2 As Double, ByVal EvalMes2 As Integer)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_upd_ActualizarFormatoConConvenioEvalMes1_2 '"
    lsSQL = lsSQL & cCtaCod & "'," & dFechaMes2 & "," & nAnoMes2 & "," & nRemBruTotalMes2 & "," & nDescMes2 & "," & _
                    nIngrNetoMes2 & "," & EvalMes2 & ""
    
    oConecta.Ejecutar lsSQL

End Sub
'FIN Actualizar Datos Formato Con Convenio EvalMes2

'Guardar Datos Formato Con Convenio Promedio
'JOEP20160530
Public Sub GuardarFormatoConConvenio_Promedio(ByVal cCtaCod As String, ByVal nRemBruTotalPromedio As Double, ByVal nDescuentoPromedio As String, ByVal nIngNetoPromedio As Double, ByVal nMonMaxIngDescuento As Double, ByVal nCapPagoPorc As Integer, ByVal nCapPagoTotal As Double)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_ins_ConsumoConCovenio_Promedio '"
    lsSQL = lsSQL & cCtaCod & "'," & nRemBruTotalPromedio & "," & nDescuentoPromedio & "," & nIngNetoPromedio & "," & nMonMaxIngDescuento & "," & nCapPagoPorc & "," & nCapPagoTotal & ""
    
    oConecta.Ejecutar lsSQL

End Sub
'FIN Guardar Datos Formato Con Convenio Promedio

'Actualizar Datos Formato Con Convenio Promedio
'JOEP20160530
Public Sub ActualizarFormatoConConvenio_Promedio(ByVal cCtaCod As String, ByVal nRemBruTotalPromedio As Double, ByVal nDescuentoPromedio As String, ByVal nIngNetoPromedio As Double, ByVal nMonMaxIngDescuento As Double, ByVal nCapPagoPorc As Integer, ByVal nCapPagoTotal As Double)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_upd_ActualizarFormatoConCovenioPromedio '"
    lsSQL = lsSQL & cCtaCod & "'," & nRemBruTotalPromedio & "," & nDescuentoPromedio & "," & nIngNetoPromedio & "," & nMonMaxIngDescuento & "," & nCapPagoPorc & "," & nCapPagoTotal & ""
    
    oConecta.Ejecutar lsSQL

End Sub
'FIN Actualizar Datos Formato Con Convenio Promedio

'Guardar Datos Formato Con Convenio
'JOEP20160530
Public Sub GuardarFormatoConConvenio_PropuestaCredito(ByVal cCtaCod As String, ByVal ncodForm As Integer, ByVal dFecVisita As Date, ByVal cEntornoFami As String, ByVal cGiroUbica As String, ByVal cExpeCrediticia As String, ByVal cFormaNegocio As String, ByVal cColateGarantia As String, ByVal cDestino As String)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_ins_ConsumoConCovenio_PropuestaCredito '"
    lsSQL = lsSQL & cCtaCod & "'," & ncodForm & ",'" & Format(dFecVisita, "yyyyMMdd") & "','" & cEntornoFami & "','" & cGiroUbica & "','" & cExpeCrediticia & "','" & cFormaNegocio & "','" & cColateGarantia & "','" & cDestino & "'"
    
    oConecta.Ejecutar lsSQL
End Sub
'FIN Guardar Datos Formato Con Convenio

'Actualizar Datos Formato Con Convenio
'JOEP20160530
Public Sub ActualizarFormatoConConvenio_PropuestaCredito(ByVal cCtaCod As String, ByVal ncodForm As Integer, ByVal dFecVisita As Date, ByVal cEntornoFami As String, ByVal cGiroUbica As String, ByVal cExpeCrediticia As String, ByVal cFormaNegocio As String, ByVal cColateGarantia As String, ByVal cDestino As String)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_upd_ActualizarFormatoConConvenioPropuestaCredito '"
    lsSQL = lsSQL & cCtaCod & "'," & ncodForm & ",'" & Format(dFecVisita, "yyyyMMdd") & "','" & cEntornoFami & "','" & cGiroUbica & "','" & cExpeCrediticia & "','" & cFormaNegocio & "','" & cColateGarantia & "','" & cDestino & "'"
    
    oConecta.Ejecutar lsSQL
End Sub
'FIN Actualizar Datos Formato Con Convenio

'Guardar Datos Formato Con Convenio Tabla Remuneracion Bruta total
'JOEP20160530
Public Sub GuardarFormatoConConvenio_RemBrutaTotal(ByVal cCtaCod As String, ByVal EvalMes1 As Integer, ByVal nCodRemBruTot As Integer, ByVal cConcepto As String, ByVal nMonto As Double)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_ins_ConsumoConCovenio_RemuneracionBrutaTotal '"
    lsSQL = lsSQL & cCtaCod & "'," & EvalMes1 & "," & nCodRemBruTot & ",'" & cConcepto & "'," & nMonto & ""
    
    oConecta.Ejecutar lsSQL
End Sub
'FIN Guardar Datos Formato Con Convenio Tabla Remuneracion Bruta total

'Actualizar Datos Formato Con Convenio Tabla Remuneracion Bruta total
'JOEP20160530
Public Sub ActualizarFormatoConConvenio_RemBrutaTotal(ByVal cCtaCod As String, ByVal EvalMes1 As Integer, ByVal nCodRemBruTot As Integer, ByVal cConcepto As String, ByVal nMonto As Double)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_upd_ActualizarFormatoConCovenioRemuneracionBrutaTotal '"
    lsSQL = lsSQL & cCtaCod & "'," & EvalMes1 & "," & nCodRemBruTot & ",'" & cConcepto & "'," & nMonto & ""
    
    oConecta.Ejecutar lsSQL
End Sub
'FIN Actualizar Datos Formato Con Convenio Tabla Remuneracion Bruta total

'JOEP 20160812 Eliminar Remuneracion Bruta Total
Public Sub EliminarRemBrutaTotal(ByVal cCtaCod As String, ByVal EvalMes1 As Integer)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_del_EliminarRemuBrutaTotal '"
    lsSQL = lsSQL & cCtaCod & "'," & EvalMes1 & ""
    
    oConecta.Ejecutar lsSQL
End Sub
'FIN JOEP 20160812 Eliminar Remuneracion Bruta Total

'Guardar Datos Formato Con Convenio Tabla Descuento
'JOEP20160530
Public Sub GuardarFormatoConConvenio_Descuento(ByVal cCtaCod As String, ByVal EvalMes1 As Integer, ByVal nCodDesc As Integer, ByVal cConcepto As String, ByVal nMonto As Double)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_ins_ConsumoConCovenio_Descuento '"
    lsSQL = lsSQL & cCtaCod & "'," & EvalMes1 & "," & nCodDesc & ",'" & cConcepto & "'," & nMonto & ""
    
    oConecta.Ejecutar lsSQL
End Sub
'FIN Guardar Datos Formato Con Convenio Tabla Descuento

'Actualizar Datos Formato Con Convenio Tabla Descuento
'JOEP20160530
Public Sub ActualizarFormatoConConvenio_Descuento(ByVal cCtaCod As String, ByVal EvalMes1 As Integer, ByVal nCodDesc As Integer, ByVal cConcepto As String, ByVal nMonto As Double)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_upd_ActualizarFormatoConCovenioDescuento '"
    lsSQL = lsSQL & cCtaCod & "'," & EvalMes1 & "," & nCodDesc & ",'" & cConcepto & "'," & nMonto & ""
    
    oConecta.Ejecutar lsSQL
End Sub
'FIN Actualizar Datos Formato Con Convenio Tabla Descuento

'JOEP 20160812 Eliminar Descuento
Public Sub EliminarDescuento(ByVal cCtaCod As String, ByVal EvalMes As Integer)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_del_EliminarDescuento '"
    lsSQL = lsSQL & cCtaCod & "'," & EvalMes & ""
    
    oConecta.Ejecutar lsSQL
End Sub
'Fin JOEP 20160812 Eliminar Descuento

'Mostrar Ingresos

'FIN Mostrar Ingresos
Public Function MostrarIngresos(ByVal psCtaCod As String, ByVal pnCodForm As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS004_2016_MostrarIngresos '" & psCtaCod & "'," & pnCodForm & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarIngresos = rs
    
    Set rs = Nothing
End Function

'Mostrar Egresos
Public Function MostrarEgresos(ByVal psCtaCod As String, ByVal pnCodForm As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS004_2016_MostrarEgresos'" & psCtaCod & "'," & pnCodForm & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarEgresos = rs
    
    Set rs = Nothing
End Function

'FIN Mostrar Egresos

'Mostrar Cuotas Ifis
Public Function MostrarCuotasIfis(ByVal psCtaCod As String, ByVal ncodForm As Integer, ByVal nConsCod As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_MostrarCuotaIfis '" & psCtaCod & "'," & ncodForm & "," & nConsCod & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarCuotasIfis = rs
    
    Set rs = Nothing
End Function

'Fin Mostrar Cuotas Ifis

'Mostrar Informe de Visita Sin Convenio
'JOEP20160530
Public Function MostrarDatosInformeVisitaFormatoSinConvenio(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec spt_sel_MostraInfromeVisitaConsumoSinConvenio'" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarDatosInformeVisitaFormatoSinConvenio = rs
    
    Set rs = Nothing
End Function
'FIN Mostrar Informe de Visita Sin Convenio

Public Function MostrarFormatoSinConvenioInfVisCabecera(ByVal psCtaCod As String, ByVal ncodForm As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec spt_sel_ERS004_2016_MostraConsumoSinConvenioInfVisCabecera'" & psCtaCod & "'," & ncodForm & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarFormatoSinConvenioInfVisCabecera = rs
    
    Set rs = Nothing
End Function


'Convenio Sin Convenio
'JOEP20160530
Public Function RecuperarDatosConsumoSinConvenio(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ConsumoSinConvenio'" & pcCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosConsumoSinConvenio = rs
   
    Set rs = Nothing
End Function
'FIN Convenio Sin Convenio

'Guardar Datos Formato sin Convenio
'JOEP20160530
Public Sub GuardarFormatoSinConvenio_InfCliente(ByVal cCtaCod As String, ByVal ncodForm As Integer, ByVal cActividad As String, ByVal nAntgAnios As Integer, _
                                                ByVal nAntgMes As Integer, ByVal nUltEndeSBS As Double, ByVal nTipoAportacion As Integer, ByVal dUltEndeuSBS As Date, ByVal nMontoSol As Double, _
                                                ByVal nNumCuotas As Integer, ByVal nExposiCred As Double, ByVal nNunDepend As Integer, ByVal cPersCodEmpleado As String, _
                                                ByVal dFecFteIng As Date, ByVal cComentario As String)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_ins_ConsumoSinCovenio_InfCliente '"
    lsSQL = lsSQL & cCtaCod & "'," & ncodForm & ",'" & cActividad & "'," & nAntgAnios & "," & nAntgMes & "," & _
                    nUltEndeSBS & "," & nTipoAportacion & ",'" & Format(dUltEndeuSBS, "yyyyMMdd") & "'," & nMontoSol & "," & nNumCuotas & "," & _
                    nExposiCred & "," & nNunDepend & ",'" & cPersCodEmpleado & "','" & Format(dFecFteIng, "yyyyMMdd") & "','" & cComentario & "'"
    
    oConecta.Ejecutar lsSQL

End Sub
'FIN Guardar Datos Formato sin Convenio

'Actualizar Datos Formato sin Convenio
'JOEP20160530
Public Sub ActualizarFormatoSinConvenio_InfCliente(ByVal cCtaCod As String, ByVal ncodForm As Integer, ByVal cActividad As String, ByVal nAntgAnios As Integer, _
                                                ByVal nAntgMes As Integer, ByVal nUltEndeSBS As Double, ByVal nTipoAportacion As Integer, ByVal dUltEndeuSBS As Date, ByVal nMontoSol As Double, _
                                                ByVal nNumCuotas As Integer, ByVal nExposiCred As Double, ByVal nNunDepend As Integer, ByVal cPersCodEmpleado As String, _
                                                ByVal dFecFteIng As Date, ByVal cComentario As String)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_upd_ActualizarFormatoSinConvenio_InfCliente '"
    lsSQL = lsSQL & cCtaCod & "'," & ncodForm & ",'" & cActividad & "'," & nAntgAnios & "," & nAntgMes & "," & _
                    nUltEndeSBS & "," & nTipoAportacion & ",'" & Format(dUltEndeuSBS, "yyyyMMdd") & "'," & nMontoSol & "," & nNumCuotas & "," & _
                    nExposiCred & "," & nNunDepend & ",'" & cPersCodEmpleado & "','" & Format(dFecFteIng, "yyyyMMdd") & "','" & cComentario & "'"
    
    oConecta.Ejecutar lsSQL

End Sub
'FIN Actualizar Datos Formato sin Convenio

'Guardar Datos Formato sin Convenio
'JOEP20160530
Public Sub GuardarFormatoSinConvenio_PropuestaCredito(ByVal cCtaCod As String, ByVal ncodForm As Integer, ByVal dFecVisita As Date, ByVal cEntornoFami As String, ByVal cGiroUbica As String, ByVal cExpeCrediticia As String, ByVal cFormaNegocio As String, ByVal cColateGarantia As String, ByVal cDestino As String)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_ins_ConsumoSinCovenio_PropuestaCredito '"
    lsSQL = lsSQL & cCtaCod & "'," & ncodForm & ",'" & Format(dFecVisita, "yyyyMMdd") & "','" & cEntornoFami & "','" & cGiroUbica & "','" & cExpeCrediticia & "','" & cFormaNegocio & "','" & cColateGarantia & "','" & cDestino & "'"
    
    oConecta.Ejecutar lsSQL
End Sub
'FIN Guardar Datos Formato sin Convenio.

'Actualizar Datos Formato sin Convenio
'JOEP20160530
Public Sub ActualizarFormatoSinConvenio_PropuestaCredito(ByVal cCtaCod As String, ByVal ncodForm As Integer, ByVal dFecVisita As Date, ByVal cEntornoFami As String, ByVal cGiroUbica As String, ByVal cExpeCrediticia As String, ByVal cFormaNegocio As String, ByVal cColateGarantia As String, ByVal cDestino As String)
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    lsSQL = "exec  stp_upd_ActualizarFormatoSinCovenioPropuestaCredito '"
    lsSQL = lsSQL & cCtaCod & "'," & ncodForm & ",'" & Format(dFecVisita, "yyyyMMdd") & "','" & cEntornoFami & "','" & cGiroUbica & "','" & cExpeCrediticia & "','" & cFormaNegocio & "','" & cColateGarantia & "','" & cDestino & "'"
    
    oConecta.Ejecutar lsSQL
End Sub
'FIN Actulizar Datos Formato sin Convenio


'Actualizar Datos Formato sin Convenio
'JOEP20160530
Public Sub ActualizarCredFormEvalSinConvenioReferidos(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnNroRef As Integer, _
                                        ByVal psNombres As String, ByVal nDNI As String, _
                                        ByVal nTelefono As String, ByVal psReferido As String, ByVal nNroDNI As String)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_upd_ActualizarFormatoSinCovenioReferidos '" & psCtaCod & "', " & pnFormato & _
                                                 ", " & pnNroRef & ", '" & psNombres & "', '" & nDNI & _
                                                 "', '" & nTelefono & "', '" & psReferido & "', '" & nNroDNI & "'"
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Actualizar Datos Formato sin Convenio

'Actualizar Datos Formato sin Convenio
'JOEP20160530
Public Sub ActualizarCredFormEvalConConvenioReferidos(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnNroRef As Integer, _
                                        ByVal psNombres As String, ByVal nDNI As String, _
                                        ByVal nTelefono As String, ByVal psReferido As String, ByVal nNroDNI As String)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec  stp_upd_ActualizarFormatoConCovenioReferidos'" & psCtaCod & "', " & pnFormato & _
                                                 ", " & pnNroRef & ", '" & psNombres & "', '" & nDNI & _
                                                 "', '" & nTelefono & "', '" & psReferido & "', '" & nNroDNI & "'"
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Actualizar Datos Formato sin Convenio

'Actualizar Datos Formato sin Convenio
'JOEP20160530
Public Sub OpcionalActualizarCredFormEvalSinConvenioReferidos(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnNroRef As Integer, _
                                        ByVal psNombres As String, ByVal nDNI As String, _
                                        ByVal nTelefono As String, ByVal psReferido As String, ByVal nNroDNI As String)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_upd_ActualizarFormatoSinCovenioReferidos '" & psCtaCod & "', " & pnFormato & _
                                                 ", " & pnNroRef & ", '" & psNombres & "', '" & nDNI & _
                                                 "', '" & nTelefono & "', '" & psReferido & "', '" & nNroDNI & "'"
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Actualizar Datos Formato sin Convenio


'Actualizar Datos Formato sin Convenio Grilla Ingresos
'JOEP20160530
Public Sub ActualizarCredFormEvalSinConvenioOtrosIngr(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnCodGasto As Double, ByVal pnMonto As Double)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_upd_ActualizarFormatoSinCovenioOtrosIngr '" & psCtaCod & "', " & pnFormato & ", " & _
                                                    pnCodGasto & ", " & pnMonto & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Actualizar Datos Formato sin Convenio Grillo Ingresos

'Guardar Datos Formato sin Convenio Grilla Boleta de Pago
'JOEP20160530
Public Sub GuardarFormatoConConvenio_BoletaPago(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal pnMontoNeto As Double, ByVal pnMontoBruto As Double) 'Agrego pnMontoBruto ACTA Nº 112-2018 JOEP20180614
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_ConsumoSinCovenio_BoletaPago '" & psCtaCod & "', " & pnFormato & ", " & _
                                                    pnAnio & ", " & pnMes & "," & pnMontoBruto & "," & pnMontoNeto & "" 'Agrego pnMontoBruto ACTA Nº 112-2018 JOEP20180614
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Guardar Datos Formato sin Convenio Grilla Boleta de Pago

'Guardar Datos Formato sin Convenio Grilla Boleta de Pago
'JOEP20160530
Public Sub ActualizarFormatoSinConvenio_BoletaPago(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal pnMonto As Double)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_upd_ActualizarFormatoSinCovenioBoletaPago '" & psCtaCod & "', " & pnFormato & ", " & _
                                                    pnAnio & ", " & pnMes & "," & pnMonto & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Guardar Datos Formato sin Convenio Grilla Boleta de Pago

'Actualizar Datos Formato sin Convenio Grilla Boleta de Pago
'JOEP20160530
Public Sub ActualizarFormatoConConvenio_BoletaPago(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal pnMonto As Double)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_upd_ActualizarFormatoSinCovenioBoletaPago '" & psCtaCod & "', " & pnFormato & ", " & _
                                                    pnAnio & ", " & pnMes & "," & pnMonto & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Actualizar Datos Formato sin Convenio Grilla Boleta de Pago

'Guardar Datos Formato sin Convenio Grilla Negocio
'JOEP20160530
Public Sub GuardarFormatoConConvenio_GastosNegocio(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnCodGasto As Integer, ByVal pnMonto As Double)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_ConsumoSinCovenio_GastosNegocio '" & psCtaCod & "', " & pnFormato & ", " & _
                                                    pnCodGasto & ", " & pnMonto & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Guardar Datos Formato sin Convenio Grilla Negocio

'Actualizar Datos Formato sin Convenio Grilla Negocio
'JOEP20160530
Public Sub ActualizarFormatoSinConvenio_GastosNegocio(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnCodGasto As Integer, ByVal pnMonto As Double)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_upd_ActualizarFormatoSinCovenioGastoNegocio '" & psCtaCod & "', " & pnFormato & ", " & _
                                                    pnCodGasto & ", " & pnMonto & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Actualizar Datos Formato sin Convenio Grilla Negocio

'Actualizar Datos Formato sin Convenio Grilla Cuotas IFIS
'JOEP20160530
Public Sub ActualizarCredFormEvalSinConvenioCuotaIFI(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnConsCod As Integer, _
                                       ByVal pnConsValor As Integer, ByVal pnNroCuota As Integer, ByVal psDescripcion As String, ByVal pnMonto As Double, ByVal nEstado As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_upd_ActualizarFormatoSinCovenioCuotaIfi '" & psCtaCod & "', " & pnFormato & ", " & pnConsCod & " , " & pnConsValor & " ," & _
                                                pnNroCuota & ",'" & psDescripcion & "', " & pnMonto & "," & nEstado & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Actulizar Datos Formato sin Convenio Grilla Cuotas IFIS

'Actualizar Datos Formato sin Convenio Grilla Cuotas IFIS
'JOEP20160530
Public Sub ActualizarCredFormEvalCuotaIFI(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnConsCod As Integer, _
                                       ByVal pnConsValor As Integer, ByVal pnNroCuota As Integer, ByVal psDescripcion As String, ByVal pnMonto As Double)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_upd_ActualizarFormatoSinCovenioCuotaIfi '" & psCtaCod & "', " & pnFormato & ", " & pnConsCod & " , " & pnConsValor & " ," & _
                                                pnNroCuota & ",'" & psDescripcion & "', " & pnMonto & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Actualizar Datos Formato sin Convenio Grilla Cuotas IFIS

'Guardar Datos Formato sin Convenio Grilla Negocio Ventas Costos
'JOEP20160530
Public Sub GuardarFormatoSinConvenio_NegocioVentasCostos(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnIngNegocio As Double, ByVal pnEgrVenta As Double, ByVal pnMargBruto As Double, Optional ByVal pnIngNetoNegocio As Double = -1)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_ConsumoSinCovenio_NegocioVentasCosto '" & psCtaCod & "', " & pnFormato & ", " & _
                                                    pnIngNegocio & ", " & pnEgrVenta & " ," & pnMargBruto & "," & pnIngNetoNegocio & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Guardar Datos Formato sin Convenio Grilla Negocio Ventas Costos

'Actualizar Datos Formato sin Convenio Grilla Negocio Ventas Costos
'JOEP20160530
Public Sub ActualizarFormatoSinConvenio_NegocioVentasCostos(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnIngNegocio As Double, ByVal pnEgrVenta As Double, ByVal pnMargBruto As Double, Optional ByVal pnIngNetoNegocio As Double = -1)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_upd_ActualizarFormatoSinCovenioNegocioGastosVenta '" & psCtaCod & "', " & pnFormato & ", " & _
                                                    pnIngNegocio & ", " & pnEgrVenta & " ," & pnMargBruto & "," & pnIngNetoNegocio & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Actualizar Datos Formato sin Convenio Grilla Negocio Ventas Costos

'Guardar Datos Formato sin Convenio Grilla Recibo por Honorarios
'JOEP20160530
Public Sub GuardarFormatoConConvenio_ReciboHonorarios(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal pnMonto As Double)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_ConsumoSinCovenio_ReciboHonorarios '" & psCtaCod & "'," & pnFormato & "," & _
                                                    pnAnio & "," & pnMes & "," & pnMonto & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Guardar Datos Formato sin Convenio Grilla Recibo por Honorarios

'Actualizar Datos Formato sin Convenio Grilla Recibo por Honorarios
'JOEP20160530
Public Sub ActualizarFormatoSinConvenio_ReciboHonorarios(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal pnMonto As Double)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_upd_ActualizarFormatoSinCovenioReciboHonorarios '" & psCtaCod & "'," & pnFormato & "," & _
                                                    pnAnio & "," & pnMes & "," & pnMonto & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Actulizar Datos Formato sin Convenio Grilla Recibo por Honorarios

'Actualizar Datos Formato sin Convenio Grilla Recibo por Honorarios
'JOEP20160530
Public Sub ActualizarFormatoConConvenio_ReciboHonorarios(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal pnMonto As Double)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_upd_ActualizarFormatoSinCovenioReciboHonorarios '" & psCtaCod & "'," & pnFormato & "," & _
                                                    pnAnio & "," & pnMes & "," & pnMonto & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Actulizar Datos Formato sin Convenio Grilla Recibo por Honorarios

'Actualizar Datos Formato sin Convenio Grillo Egresos
'JOEP20160530
Public Sub ActualizarCredFormEvalSinConvenioGastoFami(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnCodGasto As Double, ByVal pnMonto As Double)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_upd_ActualizarFormatoSinCovenioGastoFami '" & psCtaCod & "', " & pnFormato & ", " & pnCodGasto & ", " & pnMonto & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN Actualizar Datos Formato sin Convenio Grillo Egresos

'Mostrar Datos Formato sin Convenio ComboBox
'JOEP20160530
Public Function MostrarComboFecha() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_MostraComboFecha"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarComboFecha = rs
    Set rs = Nothing

End Function
'FIN Guardar Datos Formato sin Convenio Grillo Egresos

'Recupera Datos Total de sin Convenio
'JOEP20160530
Public Function RecuperarDatosTotalConsumoSinConvenio(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalConsumoSinConvenio '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoSinConvenio = rs
    
    Set rs = Nothing
End Function
'FIN Recupera Datos Total de sin Convenio

'Recupera Datos Total de sin Convenio Propuesta de Credito
'JOEP20160530
Public Function RecuperarConsumoSinConvenioPropuestaCredito(ByVal psCtaCod As String, ByVal ncodForm As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS004_2016_RecuperarConsumoSinConvenioPropuestaCredito '" & psCtaCod & "'," & ncodForm & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarConsumoSinConvenioPropuestaCredito = rs
    
    Set rs = Nothing
End Function
'FIN Recupera Datos Total de sin Convenio Propuesta de Credito

'Recupera Datos Total de sin Convenio Referidos
'JOEP20160530
Public Function RecuperarDatosTotalConsumoSinConvenioReferidos(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalReferidos '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoSinConvenioReferidos = rs
    
    Set rs = Nothing
End Function
'FIN Recupera Datos Total de sin Convenio Referidos

'Recupera Datos Total de sin Convenio Ingresos
'JOEP20160530
Public Function RecuperarDatosTotalConsumoSinConvenioIngresos(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalIngrsos '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoSinConvenioIngresos = rs
    
    Set rs = Nothing
End Function
'FIN Recupera Datos Total de sin Convenio Ingresos

'Recupera Datos Total de sin Convenio Egresos
'JOEP20160530
Public Function RecuperarDatosTotalConsumoSinConvenioEgresos(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalEgresos '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoSinConvenioEgresos = rs
    
    Set rs = Nothing
End Function
'FIN Recupera Datos Total de sin Convenio Egresos

'Recupera Datos Total de sin Convenio Boleta de Pago
'JOEP20160530
Public Function RecuperarDatosTotalConsumoSinConvenioBoletaPago(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalBoletaPago '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoSinConvenioBoletaPago = rs
    
    Set rs = Nothing
End Function
'FIN Recupera Datos Total de sin Convenio Boleta de Pago

'Recupera Datos Total de sin Convenio Recibo Honorarios
'JOEP20160530
Public Function RecuperarDatosTotalConsumoSinConvenioReciboHonorarios(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalReciboHonorarios '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoSinConvenioReciboHonorarios = rs
    
    Set rs = Nothing
End Function
'FIN Recupera Datos Total de sin Convenio Recibo Honorarios

'Recupera Datos Total de sin Convenio Cuotas IFIS
'JOEP20160530
Public Function RecuperarDatosTotalConsumoSinConvenioCuotasIfis(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalCuotasIFIS '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoSinConvenioCuotasIfis = rs
    
    Set rs = Nothing
End Function
'FIN Recupera Datos Total de sin Convenio Cuotas IFIS

'Recupera Datos Total de sin Convenio Gastos Negocio
'JOEP20160530
Public Function RecuperarDatosTotalConsumoSinConvenioGastoNegocio(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalGastoNegocio '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoSinConvenioGastoNegocio = rs
    
    Set rs = Nothing
End Function
'FIN Recupera Datos Total de sin Convenio Gastos Negocio

'Recupera Datos Total de sin Convenio Gastos Negocio Ventas Costo
'JOEP20160530
Public Function RecuperarDatosTotalConsumoSinConvenioGastoNegocioVentasCosto(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_RecuperarDatosTotalVentasCosto '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalConsumoSinConvenioGastoNegocioVentasCosto = rs
    
    Set rs = Nothing
End Function
'FIN Recupera Datos Total de sin Convenio Gastos Negocio Ventas Costo

''JOEP20160530
'Public Function RecuperarDatosLineAut_Form1_6(ByVal pcPersCod As String)
'    Dim sSql As String
'    Dim oConecta As COMConecta.DCOMConecta
'
'    On Error GoTo ErrorRecuperarDatosLineAut_Form1_6
'
'    sSql = "stp_sel_RecuperarDatosLineaAut_Formatos1_6 '" & pcPersCod & "'"
'
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set RecuperarDatosLineAut_Form1_6 = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorRecuperarDatosLineAut_Form1_6:
'    Err.Raise Err.Number, "Recuperar Solicitud Datos Basicos", Err.Description
'End Function

''JOEP20160530
'Public Function RecuperarEsatdoNuevaLineaAutomatica(ByVal pcNumLineaAut As String)
'    Dim oConec As COMConecta.DCOMConecta
'    Dim sSql As String
'    Dim rs As ADODB.Recordset
'
'    sSql = "Select nEstado From CredFormEvalLineaCredAutomatica Where cNumLineaAut='" & pcNumLineaAut & "'"
'
'    Set oConec = New COMConecta.DCOMConecta
'    oConec.AbreConexion
'    Set rs = oConec.CargaRecordSet(sSql)
'    oConec.CierraConexion
'    Set oConec = Nothing
'
'    If Not rs.EOF And Not rs.BOF Then
'        RecuperarEsatdoNuevaLineaAutomatica = rs!nEstado
'    End If
'    Set rs = Nothing
'End Function
'
''JOEP20160530
'Public Function RecuperarClienteHojaAprobacionLineaAutomatica(ByVal nNumLineaAut As String) As ADODB.Recordset
'    Dim sSql As String
'    Dim oCon As COMConecta.DCOMConecta
'    Dim rs As ADODB.Recordset
'
'    sSql = "exec stp_sel_MostrarHojaAprobacionLineaAutomaticaCliente'" & nNumLineaAut & "'"
'
'    Set oCon = New COMConecta.DCOMConecta
'    oCon.AbreConexion
'    Set rs = oCon.CargaRecordSet(sSql)
'    oCon.CierraConexion
'    Set oCon = Nothing
'
'    Set RecuperarClienteHojaAprobacionLineaAutomatica = rs
'
'    Set rs = Nothing
'End Function

'JOEP20160530
'Public Function RecuperarDatosHojaAprobacionLineaAutomatica(ByVal nNumLineaAut As String) As ADODB.Recordset
'    Dim sSql As String
'    Dim oCon As COMConecta.DCOMConecta
'    Dim rs As ADODB.Recordset
'
'    sSql = "exec stp_sel_MostrarHojaAprobacionLineaAutomaticaDatos'" & nNumLineaAut & "'"
'
'    Set oCon = New COMConecta.DCOMConecta
'    oCon.AbreConexion
'    Set rs = oCon.CargaRecordSet(sSql)
'    oCon.CierraConexion
'    Set oCon = Nothing
'
'    Set RecuperarDatosHojaAprobacionLineaAutomatica = rs
'
'    Set rs = Nothing
'End Function
'
''JOEP20160530
'Public Function RecuperarAnalistaHojaAprobacionLineaAutomatica(ByVal nNumLineaAut As String) As ADODB.Recordset
'    Dim sSql As String
'    Dim oCon As COMConecta.DCOMConecta
'    Dim rs As ADODB.Recordset
'
'    sSql = "exec stp_sel_MostrarHojaAprobacionAnalista'" & nNumLineaAut & "'"
'
'    Set oCon = New COMConecta.DCOMConecta
'    oCon.AbreConexion
'    Set rs = oCon.CargaRecordSet(sSql)
'    oCon.CierraConexion
'    Set oCon = Nothing
'
'    Set RecuperarAnalistaHojaAprobacionLineaAutomatica = rs
'
'    Set rs = Nothing
'End Function

'JOEP20160530
'Public Function RecuperarPersonaFirmaHojaAprobacionLineaAutomatica(ByVal nNumLineaAut As String) As ADODB.Recordset
'    Dim sSql As String
'    Dim oCon As COMConecta.DCOMConecta
'    Dim rs As ADODB.Recordset
'
'    sSql = "exec stp_sel_MostrarHojaAprobacionPersonaFirma'" & nNumLineaAut & "'"
'
'    Set oCon = New COMConecta.DCOMConecta
'    oCon.AbreConexion
'    Set rs = oCon.CargaRecordSet(sSql)
'    oCon.CierraConexion
'    Set oCon = Nothing
'
'    Set RecuperarPersonaFirmaHojaAprobacionLineaAutomatica = rs
'
'    Set rs = Nothing
'End Function
'
''JOEP20160530
'Public Function RecuperarGarantiasHojaAprobacionLineaAutomatica(ByVal nNumLineaAut As String) As ADODB.Recordset
'    Dim sSql As String
'    Dim oCon As COMConecta.DCOMConecta
'    Dim rs As ADODB.Recordset
'
'    sSql = "exec stp_sel_MostrarHojaAprobacionGarantias'" & nNumLineaAut & "'"
'
'    Set oCon = New COMConecta.DCOMConecta
'    oCon.AbreConexion
'    Set rs = oCon.CargaRecordSet(sSql)
'    oCon.CierraConexion
'    Set oCon = Nothing
'
'
'    Set RecuperarGarantiasHojaAprobacionLineaAutomatica = rs
'
'    Set rs = Nothing
'End Function

'JOEP20160530******************************************************
Public Function RecuperarDatosInformeVisitaFormatoParalelo(ByVal psCtaCod As String) As ADODB.Recordset

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec spt_sel_MostraInfromeVisitaFormatoParalelo'" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosInformeVisitaFormatoParalelo = rs
    
    Set rs = Nothing
End Function

'Public Function RecuperarDatosCredEvalFormatoParaleloEvaluacion(ByVal psCtaCod As String) As ADODB.Recordset
'    Dim sSql As String
'    Dim oCon As COMConecta.DCOMConecta
'    Dim rs As ADODB.Recordset
'
'    sSql = "exec stp_sel_DatosBasicosFormatoParaleloEvaluacion '" & psCtaCod & "'"
'
'    Set oCon = New COMConecta.DCOMConecta
'    oCon.AbreConexion
'    Set rs = oCon.CargaRecordSet(sSql)
'    oCon.CierraConexion
'    Set oCon = Nothing
'
'    Set RecuperarDatosCredEvalFormatoParaleloEvaluacion = rs
'
'    Set rs = Nothing
'End Function

Public Function RecuperarDatosTotalFormatoParalelo(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_MantenimietoFormatoParalelo '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosTotalFormatoParalelo = rs
    
    Set rs = Nothing
End Function

'***********************************************************************************************************
'JOEP20160530
Public Function RecuperarDatosCredEvalFPEvaluacion(ByVal psCtaCod As String) As ADODB.Recordset

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_DatosBasicosFPEvaluacion '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosCredEvalFPEvaluacion = rs
   
    Set rs = Nothing
    
End Function

Public Function ValidarParalelo(ByVal cPersCod As String) As ADODB.Recordset

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS004_2016_ValidarParalelo '" & cPersCod & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set ValidarParalelo = rs

    Set rs = Nothing

End Function

Public Function RecuperarDatosCredEvalFormatoParalelo(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_DatosBasicosFormatoParalelo '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosCredEvalFormatoParalelo = rs

    Set rs = Nothing
End Function

Public Function RecuperarDatosCredEvalFormatoLineaCredAutomatica(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_DatosBasicosFormatoLineaCredAutomatico '" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarDatosCredEvalFormatoLineaCredAutomatica = rs
   
    Set rs = Nothing
End Function

Public Function RecuperarComboLinea() As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_LineCredAutomatica"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarComboLinea = rs
    Set rs = Nothing
End Function

Public Function AsignarAnalista(ByVal psPersCod) As ADODB.Recordset

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_AsignarAnalista'" & psPersCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set AsignarAnalista = rs
    Set rs = Nothing

End Function

Public Function MostrarRiesgoUnico(ByVal nPersCod As String) As ADODB.Recordset

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_MostrarRiesgoUnico'" & nPersCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarRiesgoUnico = rs
    Set rs = Nothing

End Function

Public Function EvaluarLineaAutomatica(ByVal pcCtaCod) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_DatosBasicosFormatoLineaCredAutomatico '" & pcCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set EvaluarLineaAutomatica = rs
   
    Set rs = Nothing
End Function

Public Function MostrarEstadoActual(ByVal pcPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_LineaAutomaticaEstadoActual '" & pcPersCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarEstadoActual = rs
    Set rs = Nothing
End Function

Public Function MostrarCodigoLinea() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_CredFormEvalNuevaLineaAutomaticaCodigo"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set MostrarCodigoLinea = rs
    Set rs = Nothing
End Function

Public Function RecuperarMontoRULineaAutomatica(ByVal nNumLineaAut As String)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec spt_sel_RecuperaDatosBasicosLineaAutomatica '" & nNumLineaAut & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperarMontoRULineaAutomatica = rs
   
    Set rs = Nothing
End Function

Public Function AsignarFormatoLineaAutomatica(ByVal pnMonto As Double) ' Para Linea Credito
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim AsignarFormato As Integer

    sSql = "stp_sel_AsignarCredFormEvalLineaAutomatica  " & pnMonto & " "
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    If rs.RecordCount > 0 Then
        If Not (rs.EOF And rs.BOF) Then
            AsignarFormatoLineaAutomatica = CInt(rs!ncodForm)
'            pnMin = CDbl(rs!nMontoMin)
'            pnMax = CDbl(rs!nMontoMax)
        Else
            AsignarFormatoLineaAutomatica = 0
'            pnMin = 0
'            pnMax = 0
        End If
    Else
        AsignarFormatoLineaAutomatica = 0
'        pnMin = 0
'        pnMax = 0
    End If
   
    Set rs = Nothing
End Function

Public Function RecuperaDatosFlexRemBrutaTotalMes1(ByVal psCtaCod As String, ByVal pnFormato As Integer, _
        Optional ByVal pnTpoInsConv As Integer, Optional ByVal bSecSalud As Boolean) As ADODB.Recordset
    
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperarDatosRemBrutaTotalMes1 '" & psCtaCod & "'," & pnTpoInsConv & "," & IIf(bSecSalud = True, 1, 0)
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexRemBrutaTotalMes1 = rs
    Set rs = Nothing
End Function

Public Function RecuperaDatosFlexRemBrutaTotalMes2(ByVal psCtaCod As String, ByVal pnFormato As Integer, _
                Optional ByVal pnTpoInsConv As Integer, Optional ByVal bSecSalud As Boolean) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_RecuperarDatosRemBrutaTotalMes2 '" & psCtaCod & "'," & pnTpoInsConv & "," & IIf(bSecSalud = True, 1, 0)
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexRemBrutaTotalMes2 = rs
    Set rs = Nothing
End Function
Public Function CredFormEvalObtieneDescuentoConv(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_CredFormEvalObtieneDescuentoConv '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set CredFormEvalObtieneDescuentoConv = rs
    Set rs = Nothing
End Function
'JOEP FIN*******************************************
'EJVG20160712 ***
Public Function RecuperaCreditoxEvaluacion(ByVal psCtaCod As String) As Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta

    On Error GoTo ErrExisteFormatoEvaluacion

    sSql = "EXEC stp_sel_RecuperaCredxEvaluacion '" & psCtaCod & "'"

    oConecta.AbreConexion
    Set RecuperaCreditoxEvaluacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrExisteFormatoEvaluacion:
    Err.Raise Err.Number, "Error En Proceso <<RecuperaCreditoxEvaluacion>>", Err.Description
End Function
Public Function RecuperaFormatoEvaluacion(ByVal psCtaCod As String) As Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta

    On Error GoTo ErrExisteFormatoEvaluacion

    sSql = "EXEC stp_sel_CredFormEval '" & psCtaCod & "'"

    oConecta.AbreConexion
    Set RecuperaFormatoEvaluacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrExisteFormatoEvaluacion:
    Err.Raise Err.Number, "Error En Proceso <<ExisteFormatoEvaluacion>>", Err.Description
End Function
Public Sub EliminaFormatoEvaluacion(ByVal psCtaCod As String, Optional ByVal pcFechaReg As String = "19000101")
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta

    On Error GoTo ErrEliminaFormatoEvaluacion

    sSql = "EXEC spt_del_CredFormEval '" & psCtaCod & "','" & pcFechaReg & "' "

    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrEliminaFormatoEvaluacion:
    Err.Raise Err.Number, "Error En Proceso <<EliminaFormatoEvaluacion>>", Err.Description
End Sub
Public Sub RecalculaIndicadoresyRatiosEvaluacion(ByVal pcCtaCod As String, Optional ByVal pnMontoCuota As Double = 0#, Optional ByVal pcFecReg As String = "19000101")
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    
    On Error GoTo ErrRecalculaIndicadoresyRatiosEvaluacion
    
    sSql = "EXEC stp_upd_CredformEvalratiosIndicadores '" & pcCtaCod & "'," & pnMontoCuota & ",'" & pcFecReg & "'"
    
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrRecalculaIndicadoresyRatiosEvaluacion:
    Err.Raise Err.Number, "Error En Proceso <<RecalculaIndicadoresyRatiosEvaluacion>>", Err.Description
End Sub
'END EJVG *******
'RECO20160718**************************************************
Public Function CargarOtrasIfis(Optional pnConsCod As Integer = -1, Optional pnConsValor As Integer = -1) As Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrCargarOtrasIfis
    sSql = "EXEC stp_sel_CargarOtrasIfis " & pnConsCod & ", " & pnConsValor & "" 'LUCV20161115, Modificó->Según ERS068-2016
    oConecta.AbreConexion
    Set CargarOtrasIfis = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrCargarOtrasIfis:
    Err.Raise Err.Number, "Error En Proceso <<CargarOtrasIfis>>", Err.Description
End Function
Public Function ObtieneDetalleInventario(ByVal psCtaCod As String, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnTipoPat As Integer)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrObtieneDetalleInventario
    sSql = "EXEC stp_sel_FormatoObtieneDetalleInventario '" & psCtaCod & "'," & pnConsCod & "," & pnConsValor & "," & pnTipoPat
    oConecta.AbreConexion
    Set ObtieneDetalleInventario = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrObtieneDetalleInventario:
    Err.Raise Err.Number, "Error En Proceso <<ObtieneDetalleInventario>>", Err.Description
End Function
Public Function ObtieneDetalleActiFijo(ByVal psCtaCod As String, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnTipoPat As Integer)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrObtieneDetalleActiFijo
    sSql = "EXEC stp_sel_FormatoObtieneDetalleActiFijo '" & psCtaCod & "'," & pnConsCod & "," & pnConsValor & "," & pnTipoPat
    oConecta.AbreConexion
    Set ObtieneDetalleActiFijo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrObtieneDetalleActiFijo:
    Err.Raise Err.Number, "Error En Proceso <<ObtieneDetalleActiFijo>>", Err.Description
End Function
Public Function RecuperaDatosCtaCobrar(ByVal psCtaCod As String, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnTipoPat As Integer, _
                                      Optional ByVal pnTpoConsulta As Integer = 1)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrRecuperaDatosCtaCobrar
    
    If pnTpoConsulta = 1 Then
        sSql = "EXEC stp_sel_CredFormEvalRecuperaDatosCtaCobrarIfis '" & psCtaCod & "'," & pnConsCod & "," & pnConsValor & "," & pnTipoPat
    Else
        sSql = "EXEC stp_sel_CredFormEvalRecuperaDatosCtaCobrar '" & psCtaCod & "'," & pnConsCod & "," & pnConsValor & "," & pnTipoPat
    End If
    oConecta.AbreConexion
    Set RecuperaDatosCtaCobrar = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrRecuperaDatosCtaCobrar:
    Err.Raise Err.Number, "Error En Proceso <<RecuperaDatosCtaCobrar>>", Err.Description
End Function
Public Sub ActualizaDatosLineaCredNoutilizada(ByVal psCtaCod As String, ByVal pnMontoDeuda As Currency, ByVal pnMontoCuota As Currency)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
    
    On Error GoTo ErrActualizaDatosLineaCredNoutilizada
    
    sSql = "EXEC stp_upd_ActualizaDatosLineaCredNoutilizada '" & psCtaCod & "'," & pnMontoDeuda & "," & pnMontoCuota

    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrActualizaDatosLineaCredNoutilizada:
    Err.Raise Err.Number, "Error En Proceso <<ActualizaDatosLineaCredNoutilizada>>", Err.Description
End Sub
'RECO FIN *****************************************************
'*** FRHU 20160823
Public Function CreditoTieneFormatoEvaluacion(ByVal psCtaCod As String) As Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    
    On Error GoTo ErrCreditoTieneFormatoEvaluacion

    sSql = "EXEC stp_sel_CredTieneFormatoEvaluacion '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set CreditoTieneFormatoEvaluacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ErrCreditoTieneFormatoEvaluacion:
    Err.Raise Err.Number, "Error En Proceso <<CreditoTieneFormatoEvaluacion>>", Err.Description
End Function
'*** FIN FRHU 20160823

'**** PTI120170523, Agregó Según ERS014-2017
Public Function RecuperaCreditosEvaluadosFormatos(ByVal psPersCod As String, ByVal pnTipo As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_ERS0142017_RecuperaCreditosEvaluadosFormatos '" & psPersCod & "', " & pnTipo & ""

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosEvaluadosFormatos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'**** Fin PTI120170523

'LUCV20170302, Según ANEXO 001-2017
Public Sub dBeginTrans()
    'If pbConcurrencia Then
    '    coConex.ConexionActiva.Execute "SET TRANSACTION ISOLATION LEVEL SERIALIZABLE"
    'End If
    oDCOMConecta.BeginTrans
End Sub
Public Sub dRollbackTrans()
    oDCOMConecta.RollbackTrans
    oDCOMConecta.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub
Public Sub dCommitTrans()
    oDCOMConecta.CommitTrans
    'coConex.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub

Public Function RecuperaAlertaTempranaConfigCapaPago() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ANEXO0012017_AlertaTempranaConfigCapaPago "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set RecuperaAlertaTempranaConfigCapaPago = rs
End Function

Public Function RecuperaAlertaTempranaConfigCapaPagoDet(Optional psProdCabCod As String = -1, _
                                                    Optional psTpoProdCod As String = -1, _
                                                    Optional psCredCabCod As String = -1, _
                                                    Optional psTpoCredCod As String = -1) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ANEXO0012017_AlertaTempranaConfigCapaPagoDet '" & psProdCabCod & "', '" & psTpoProdCod & "', '" & psCredCabCod & "', '" & psTpoCredCod & "' "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set RecuperaAlertaTempranaConfigCapaPagoDet = rs
End Function

Public Function RecuperaAlertaTempranaConfigRatioOtro() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ANEXO0012017_AlertaTempranaConfigRatioOtro "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set RecuperaAlertaTempranaConfigRatioOtro = rs
End Function

Public Function ObtenerCodAlertaTempranaConfig() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ANEXO0012017_AlertaTempranaConfigCodAlerta"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerCodAlertaTempranaConfig = rs
End Function
Public Sub GrabarAlertasTempranas(ByVal pcCtaCod As String, Optional ByVal pcOpeTpo As String = "000000", Optional ByVal pcFecReg As String = "19000101")
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorGrabarAlertasTempranas
    
    sSql = "EXEC stp_ins_ANEXO0012017_AlertasTempranas '" & pcCtaCod & "','" & pcOpeTpo & "','" & pcFecReg & "' "

    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorGrabarAlertasTempranas:
    Err.Raise Err.Number, "Error En Proceso <<GrabarAlertasTempranas>>", Err.Description
End Sub

Public Function GrabarAlertaTempranaConfig(ByVal pnCodAlerta As Integer, _
                                           ByVal psMovNro As String, _
                                           ByVal pnValor As Double, _
                                           ByVal pnUnidad As Integer, _
                                           ByVal pnEstado As Integer, _
                                           Optional ByVal pnTipoCliente As Integer = -1, _
                                           Optional ByVal pnProducto As Integer = -1, _
                                           Optional ByVal pnSubProducto As Integer = -1, _
                                           Optional ByVal pnCredito As Integer = -1, _
                                           Optional ByVal pnTipoCredito As Integer = -1, _
                                           Optional ByVal psNombreRatio As String = "", _
                                           Optional ByVal psFormula As String = "") As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_ins_ANEXO0012017_AlertaTempranaConfig " & pnCodAlerta & " , '" & psMovNro & "' , " & pnValor & ", " & pnUnidad & ", " & pnEstado & ", " & _
           pnTipoCliente & ", '" & pnProducto & "', '" & pnSubProducto & "', '" & pnCredito & "' ,'" & pnTipoCredito & "', '" & psNombreRatio & "', '" & psFormula & "'"
    
    oCon.AbreConexion
    Set GrabarAlertaTempranaConfig = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function EliminarAlertaTempranaConfig(Optional ByVal pnCodAlerta As Integer = -1, _
                                             Optional ByVal pnProducto As Integer = -1, _
                                             Optional ByVal pnSubProducto As Integer = -1, _
                                             Optional ByVal pnCredito As Integer = -1, _
                                             Optional ByVal pnTipoCredito As Integer = -1) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "exec stp_del_ANEXO0012017_AlertaTempranaConfig " & pnCodAlerta & ", " & pnProducto & ", " & pnSubProducto & ", " & pnCredito & ", " & pnTipoCredito & " "
    
    oCon.AbreConexion
    Set EliminarAlertaTempranaConfig = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function ObtenerEstadoLimite() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ANEXO0012017_EstadoLimite "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerEstadoLimite = rs
End Function

Public Function ObtenerValoresUnidadMedida() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ANEXO0012017_ObtenerValoresUnidadMedida"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerValoresUnidadMedida = rs
End Function

Public Function ObtieneListadoAlertasTempranas(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_ANEXO0012017_ObtieneListadoAlertasTempranas '" & pcCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneListadoAlertasTempranas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function ObtieneListadoSeccionAlertasTempranas(ByVal pcPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_ANEXO0012017_ObtieneListadoSeccionAlertasTempranas '" & pcPersCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneListadoSeccionAlertasTempranas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function VerificarExisteAlertaTemprana(ByVal psCtaCod As String) As Boolean
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New ADODB.Recordset
    
    Set oConecta = New COMConecta.DCOMConecta
    sSql = "exec stp_sel_ANEXO0012017_verificarExisteAutorizaciones '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not (oRs.BOF Or oRs.EOF) Then
        VerificarExisteAlertaTemprana = True
    Else
        VerificarExisteAlertaTemprana = False
    End If
End Function
'Fin LUCV20170302

'**********->LUCV20170915-Agregó, según: ERS051-2017 **********->
Public Function RecuperaDatosFlexActivosEEFF() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_FlexActivosEEFF "
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexActivosEEFF = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosFlexPasivosEEFF() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_FlexPasivosEEFF "
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexPasivosEEFF = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosFlexEstadoResultadoEEFF() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS0512017_FlexEstadoResultadoEEFF"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexEstadoResultadoEEFF = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosFlexFlujoEfectivoEEFF() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS0512017_FlexFLujoEfectivoEEFF"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexFlujoEfectivoEEFF = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosFlexCambiosPatrimonioEEFF() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_ERS0512017_FlexCambiosPatrimonioEEFF"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosFlexCambiosPatrimonioEEFF = rs
    Set rs = Nothing
End Function
Public Function CargarDetalleItemEEFF(Optional pnConsCod As Integer = -1, Optional pnConsValor As Integer = -1) As Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrCargarDetalleItemEEFF
    sSql = "EXEC stp_sel_ERS0512017_CargarItemEEFF " & pnConsCod & ", " & pnConsValor & ""
    oConecta.AbreConexion
    Set CargarDetalleItemEEFF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrCargarDetalleItemEEFF:
    Err.Raise Err.Number, "Error En Proceso <<CargarDetalleItemEEFF>>", Err.Description
End Function
Public Sub InsertaPersEstadoFinancieroDetNew(ByVal pnCodEF As Integer, _
                                             ByVal pnEstFinanOrdenItem As Integer, _
                                             ByVal pnEstFinanTipo As Integer, _
                                             ByVal pnEstFinanSubTipoCod As Integer, _
                                             ByVal pnEstFinanSubTipoValor As Integer, _
                                             ByVal pnEstFinanSubTipoValorGrupo As Integer, _
                                             ByVal pnMonto As Currency, _
                                             Optional ByVal pnPorcentaje As Currency = 0, _
                                             Optional ByVal pnEstFinanCapitalSocial As Currency = 0, _
                                             Optional ByVal pnEstFinanCapitalSocialAdicional As Currency = 0, _
                                             Optional ByVal pnEstFinanReservaLegal As Currency = 0, _
                                             Optional ByVal pnEstFinanResultadoAcumulado As Currency = 0, _
                                             Optional ByVal pnEstFinanUtilidadEjercicio As Currency = 0)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_ERS0512017_InsertaPersEstadoFinancieroDetNew " & pnCodEF & ", " & pnEstFinanOrdenItem & "," & pnEstFinanTipo & "," & pnEstFinanSubTipoCod & ", " & _
                                                                           pnEstFinanSubTipoValor & ", " & pnEstFinanSubTipoValorGrupo & ", " & pnMonto & ", " & pnPorcentaje & ", " & _
                                                                           pnEstFinanCapitalSocial & ", " & pnEstFinanCapitalSocialAdicional & "," & pnEstFinanReservaLegal & ", " & _
                                                                           pnEstFinanResultadoAcumulado & ", " & pnEstFinanUtilidadEjercicio & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub InsertaPersEstadoFinancieroDetNewDet(ByVal pnCodEF As Integer, _
                                             ByVal pnEstFinanSubTipoCod As Integer, _
                                             ByVal pnEstFinanSubTipoValor As Integer, _
                                             ByVal pnEstFinanSubTipoValorGrupo As Integer, _
                                             ByVal pcEstFinanDescripcion As String, _
                                             ByVal pnMonto As Currency)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    
    lsSQL = "exec stp_ins_ERS0512017_InsertaPersEstadoFinancieroDetNewDet " & pnCodEF & "," & pnEstFinanSubTipoCod & ", " & pnEstFinanSubTipoValor & ", " & pnEstFinanSubTipoValorGrupo & ", '" & pcEstFinanDescripcion & "',  " & pnMonto & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Sub EliminaPersEstadoFinancieroDetNew(ByVal pnCodEF As String, ByVal pnConsValor As Integer)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_del_ERS0512017_EliminaPersEstadoFinancieroDetNew " & pnCodEF & ", " & pnConsValor & " "
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Function RecuperaDatosActivosEEFF(ByVal pnCodEF As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_RecuperaDatosActivosEEFF " & pnCodEF & ""
        Set oCon = New COMConecta.DCOMConecta
        oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(sSql)
        oCon.CierraConexion
        Set oCon = Nothing
        
        Set RecuperaDatosActivosEEFF = rs
        Set rs = Nothing
End Function
Public Function RecuperaDatosPasivosEEFF(ByVal pnCodEF As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_RecuperaDatosPasivosEEFF " & pnCodEF & ""
        Set oCon = New COMConecta.DCOMConecta
        oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(sSql)
        oCon.CierraConexion
        Set oCon = Nothing
        
        Set RecuperaDatosPasivosEEFF = rs
        Set rs = Nothing
End Function
Public Function RecuperaDatosEstadoResultadoEEFF(ByVal pnCodEF As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_RecuperaDatosEstadoResultadoEEFF " & pnCodEF & ""
        Set oCon = New COMConecta.DCOMConecta
        oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(sSql)
        oCon.CierraConexion
        Set oCon = Nothing
        
        Set RecuperaDatosEstadoResultadoEEFF = rs
        Set rs = Nothing
End Function
Public Function RecuperaDatosFlujoEfectivoEEFF(ByVal pnCodEF As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_RecuperaDatosFlujoEfectivoEEFF " & pnCodEF & ""
        Set oCon = New COMConecta.DCOMConecta
        oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(sSql)
        oCon.CierraConexion
        Set oCon = Nothing
        
        Set RecuperaDatosFlujoEfectivoEEFF = rs
        Set rs = Nothing
End Function
Public Function RecuperaDatosCambiosPatrimonioEEFF(ByVal pnCodEF As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_RecuperaDatosCambiosPatrimonioEEFF " & pnCodEF & ""
        Set oCon = New COMConecta.DCOMConecta
        oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(sSql)
        oCon.CierraConexion
        Set oCon = Nothing
        
        Set RecuperaDatosCambiosPatrimonioEEFF = rs
        Set rs = Nothing
End Function
Public Function RecuperaDatosPersEstadosFinancieros(ByVal pnCodEF As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_RecuperaDatosPersEstadosFinancieros " & pnCodEF & ""
        Set oCon = New COMConecta.DCOMConecta
        oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(sSql)
        oCon.CierraConexion
        Set oCon = Nothing
        
        Set RecuperaDatosPersEstadosFinancieros = rs
        Set rs = Nothing
End Function
Public Function RecuperaDatosDetalleSubItem(ByVal pnCodEF As Integer, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, Optional ByVal pnConsValorGrupo As Integer = 1)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaDatosDetalleSubItem
    
    sSql = "EXEC stp_sel_ERS0512017_RecuperaDatosDetalleSubItem  " & pnCodEF & "," & pnConsCod & "," & pnConsValor & ", " & pnConsValorGrupo & ""

    oConecta.AbreConexion
    Set RecuperaDatosDetalleSubItem = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaDatosDetalleSubItem:
    Err.Raise Err.Number, "Error En Proceso <<RecuperaDatosCtaCobrar>>", Err.Description
End Function
Public Sub ActualizaPersEstadoFinanciero(ByVal pnCodEF As Integer, _
                                         ByVal pnAuditado As Integer, _
                                         ByVal psFechaPatIni As String, _
                                         ByVal psFechaPatFin As String)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_upd_ERS0512017_ActualizaPersEstadoFinanciero " & pnCodEF & ", " & pnAuditado & ", '" & psFechaPatIni & "','" & psFechaPatFin & "'"
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'->***** LUCV20171015 ->*****
Public Function RecuperaDatosComparativoActivo(ByVal psPersCod As String, ByVal pcFechaIni As String, ByVal pcFechaFin As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_RecuperaDatosComparativoActivo '" & psPersCod & "', '" & pcFechaIni & "', '" & pcFechaFin & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosComparativoActivo = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosComparativoPasivoPat(ByVal psPersCod As String, ByVal pcFechaIni As String, ByVal pcFechaFin As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_RecuperaDatosComparativoPasivoPat '" & psPersCod & "', '" & pcFechaIni & "', '" & pcFechaFin & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosComparativoPasivoPat = rs
    Set rs = Nothing
End Function
Public Function RecuperaDatosComparativoEERR(ByVal psPersCod As String, ByVal pcFechaIni As String, ByVal pcFechaFin As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_RecuperaDatosComparativoEERR '" & psPersCod & "', '" & pcFechaIni & "', '" & pcFechaFin & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaDatosComparativoEERR = rs
    Set rs = Nothing
End Function

Public Function RecuperaFlujoCajaRptObtieneDatosCabecera(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosCabecera '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaFlujoCajaRptObtieneDatosCabecera = rs
    Set rs = Nothing
End Function

Public Function RecuperaFlujoCajaRptObtieneDatosCuotas(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosCuotas '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaFlujoCajaRptObtieneDatosCuotas = rs
    Set rs = Nothing
End Function

Public Function RecuperaFlujoCajaRptObtieneDatosConceptos(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosConceptos '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaFlujoCajaRptObtieneDatosConceptos = rs
    Set rs = Nothing
End Function
Public Function RecuperaFlujoCajaRptObtieneDatosParametros(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosParametros '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set RecuperaFlujoCajaRptObtieneDatosParametros = rs
    Set rs = Nothing
End Function
Public Function VerificaActualizacionEEFF(ByVal psPersCod As String, ByVal psFechaInicial As String, ByVal psFechaFinal As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "EXEC stp_sel_ERS0512017_ValidaActualizacionEEFF '" & psPersCod & "','" & psFechaInicial & "','" & psFechaFinal & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set VerificaActualizacionEEFF = rs
    Set rs = Nothing
End Function
'<-***** Fin LUCV20171015<-*****
'**********<-LUCV20170915 Fin <-**********

'**********->INICIO EAAS20170915, Según ERS051-2017->*******
Public Function RecuperarNumeroEdicionesEF(ByVal psPersCod As String, ByVal pnCodEstFinan As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New ADODB.Recordset
    
    Set oConecta = New COMConecta.DCOMConecta
    sSql = "exec RecuperarNumeroEdicionesEF '" & psPersCod & "','" & pnCodEstFinan & "'"
    oConecta.AbreConexion
    Set RecuperarNumeroEdicionesEF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function VerificarExisteAutorizacionEF(ByVal psPersCod As String, ByVal pnCodEstFinan As Integer, ByVal psUser As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New ADODB.Recordset
    Set oConecta = New COMConecta.DCOMConecta
    sSql = "exec VerificarExisteAutorizacionEF '" & psPersCod & "','" & pnCodEstFinan & "','" & psUser & "'"
    oConecta.AbreConexion
    Set VerificarExisteAutorizacionEF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    
End Function
Public Function GrabarEstadoFinancieroEdicion(ByVal psCodUser As String, ByVal pnCodEstFinan As Integer, ByVal pgdFecSis As Date, pnTipoAccion As Integer, pcEstadoEF As String, plcMovNroEF As String, pcMaquina As String)
    Dim sFechaEdicion As String
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorGrabarEstadoFinancieroEdicion
    sFechaEdicion = Format(pgdFecSis, "yyyymmdd hh:mm:ss")
    sSql = "EXEC stp_ins_ERS0512017_EstadoFinancieroEdicion '" & psCodUser & "','" & pnCodEstFinan & "','" & sFechaEdicion & "','" & pnTipoAccion & "','" & pcEstadoEF & "','" & plcMovNroEF & "','" & pcMaquina & "' "
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorGrabarEstadoFinancieroEdicion:
    Err.Raise Err.Number, "Error En Proceso <<GrabarEstadoFinancieroEdicion>>", Err.Description
End Function

Public Function RecuperaEEFF(ByVal pnCodEF As Integer, ByVal pnOp As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaEEFF
    sSql = "stp_sel_ERS0512017_RecuperaEEFF '" & pnCodEF & "','" & pnOp & "'"
    oConecta.AbreConexion
    Set RecuperaEEFF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaEEFF:
    Err.Raise Err.Number, "Error En Proceso <<RecuperaEEFF>>", Err.Description
End Function
Public Function RecuperaUsuarioSolicitanteEdicionEEFF(psUser) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorRecuperaUsuarioSolicitanteEdicionEEFF
    sSql = "EXEC stp_sel_ERS0512017_RecuperaUsuarioSolicitanteEdicionEEFF '" & psUser & "'"
    oConecta.AbreConexion
    Set RecuperaUsuarioSolicitanteEdicionEEFF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaUsuarioSolicitanteEdicionEEFF:
    Err.Raise Err.Number, "Error En Proceso <<RecuperaUsuarioSolicitanteEdicionEEFF>>", Err.Description
End Function

Public Function GrabarSolicitudEdicionEF(ByVal psMotivo As String, ByVal pnCodEF As Integer, ByVal pgsCodUser As String, plcMovNroEF As String)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorGrabarSolicitudEdicionEF
    sSql = "EXEC stp_ins_ERS0512017_GrabarSolicitudEdicionEF '" & psMotivo & "','" & pnCodEF & "','" & pgsCodUser & "','" & plcMovNroEF & "'"
    oConecta.AbreConexion
    Set GrabarSolicitudEdicionEF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorGrabarSolicitudEdicionEF:
    Err.Raise Err.Number, "Error En Proceso <<GrabarSolicitudEdicionEF>>", Err.Description
End Function
Public Function GrabarAprobacionSolicitudEdicionEF(ByVal psMotivo As String, ByVal pnCodEF As Integer, ByVal pgsCodUser As String, ByVal pgsCodUserAutoriza As String, plcMovNroEF As String)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorGrabarAprobacionSolicitudEdicionEF
    sSql = "EXEC stp_upd_ERS0512017_GrabarAprobacionSolicitudEdicionEF '" & pnCodEF & "','" & pgsCodUser & "','" & psMotivo & "','" & pgsCodUserAutoriza & "','" & plcMovNroEF & "'"
    oConecta.AbreConexion
    Set GrabarAprobacionSolicitudEdicionEF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorGrabarAprobacionSolicitudEdicionEF:
    Err.Raise Err.Number, "Error En Proceso <<GrabarAprobacionSolicitudEdicionEF>>", Err.Description
End Function
Public Function ObtieneTipoPermiso(pbTipoPermiso As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorObtieneTipoPermiso
    sSql = "EXEC stp_sel_ERS0512017_VerificaPermisoCargo '" & pbTipoPermiso & "'"
    oConecta.AbreConexion
    Set ObtieneTipoPermiso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtieneTipoPermiso:
    Err.Raise Err.Number, "Error En Proceso <<ObtieneTipoPermiso>>", Err.Description
End Function
Public Function GrabarAnulacionSolicitudEdicionEF(ByVal psMotivo As String, ByVal pnCodEF As Integer, ByVal pgsCodUser As String, ByVal pgsCodUserAutoriza As String, plcMovNroEF As String)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorGrabarAnulacionSolicitudEdicionEF
    sSql = "EXEC stp_upd_ERS0512017_GrabarAnulacionSolicitudEdicionEF '" & pnCodEF & "','" & pgsCodUser & "','" & psMotivo & "','" & pgsCodUserAutoriza & "','" & plcMovNroEF & "'"
    oConecta.AbreConexion
    Set GrabarAnulacionSolicitudEdicionEF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorGrabarAnulacionSolicitudEdicionEF:
    Err.Raise Err.Number, "Error En Proceso <<GrabarAnulacionSolicitudEdicionEF>>", Err.Description
End Function
Public Function VerificaMigracionEstadoFinanciero(ByVal pnCodEF As Integer)
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta
    On Error GoTo ErrorVerificaMigracionEstadoFinanciero
    sSql = "EXEC stp_sel_ERS0512017_VerificaMigracionEstadoFinanciero '" & pnCodEF & "'"
    oConecta.AbreConexion
    Set VerificaMigracionEstadoFinanciero = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorVerificaMigracionEstadoFinanciero:
    Err.Raise Err.Number, "Error En Proceso <<VerificaMigracionEstadoFinanciero>>", Err.Description
End Function
Public Function VerificaPermisoEdicionEF(ByVal pnConsSisCod As Integer, ByVal pgsCodCargo As String) As Boolean
    Dim oGen As COMDConstSistema.DCOMGeneral
    Dim i As Integer
    Dim ObtieneTipoPermisoCredEval As Integer
    Dim sCargosFormatosEval As String
    Dim sCodCargo() As String
    Set oGen = New COMDConstSistema.DCOMGeneral
    
    sCargosFormatosEval = oGen.LeeConstSistema(pnConsSisCod)
    sCodCargo = Split(sCargosFormatosEval, ",")
    ObtieneTipoPermisoCredEval = 0
    
    For i = 0 To UBound(sCodCargo)
        If pgsCodCargo = sCodCargo(i) Then
            ObtieneTipoPermisoCredEval = 1
            VerificaPermisoEdicionEF = True
            Exit Function
        End If
    Next i
    VerificaPermisoEdicionEF = False
End Function
'<-********** FIN EAAS 20170926 <-**********
'->********** 'Inicio JOEP20171015 Segun ERS051-2017 Flujo de Caja **********->
Public Sub InsertaCredFormEvalParamFlujoCaja(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnIncVC As Double, _
                                                ByVal pnIncCM As Double, ByVal pnIncPP As Double, _
                                                ByVal pnIncGV As Double, ByVal pnIncC As Double)
    Dim oCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    lsSQL = "exec stp_ins_ERS0512017_insertaParametroFlujoCaja '" & psCtaCod & "', " & pnFormato & _
                                                 ", " & pnIncVC & "," & pnIncCM & "," & pnIncPP & _
                                                 ", " & pnIncGV & "," & pnIncC & ""
    oCon.ConexionActiva.Execute lsSQL
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Function RecuperaDatosParamFlujoCaja(ByVal psCtaCod As String, ByVal pnCodForm As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    sSql = "exec stp_sel_ERS0512017_RecuperaDatosParamFlujoCaja '" & psCtaCod & "', " & pnCodForm & " "
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing

    Set RecuperaDatosParamFlujoCaja = rs
    Set rs = Nothing
End Function

'<-********** Fin 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja **********
'JOEP20190919 ERS042 CP-2018
Public Function RecuperaProdMultFormato(ByVal psProdCod As String) As Recordset
    Dim sSql As String
    Dim oConecta As New COMConecta.DCOMConecta

    On Error GoTo ErrExisteFormatoEvaluacion

    sSql = "EXEC stp_sel_VerificaMultFormatos '" & psProdCod & "'"

    oConecta.AbreConexion
    Set RecuperaProdMultFormato = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrExisteFormatoEvaluacion:
    Err.Raise Err.Number, "Error En Proceso <<RecuperaProdMultFormato>>", Err.Description
End Function
'JOEP20190919 ERS042 CP-2018

'CTI320200110 ERS003-2020. Agreg?
Public Function GrabarRSEAdmisionCuota(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_ins_RSE_Admision_Cuota '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set GrabarRSEAdmisionCuota = rs
    Set rs = Nothing
End Function
Public Function GrabarRSERepAdmision(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_ins_RSE_Admision '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set GrabarRSERepAdmision = rs
    Set rs = Nothing
End Function

Public Sub ActualizarRSECuotaEstimada(ByVal psCtaCod As String)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    sSql = "exec stp_RSE_upd_CuotasEstimada '" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'Fin CTI320200110


