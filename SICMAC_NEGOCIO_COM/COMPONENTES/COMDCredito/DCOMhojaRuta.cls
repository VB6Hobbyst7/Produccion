VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMhojaRuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oCon As New DCOMConecta
Private Type Param
    valor As Variant
    comilla As Boolean
    nuloNegativo As Boolean
    nuloVacio As Boolean
End Type
Public Function getValCombo(ByVal t As String)
    If t = "" Then
        getValCombo = -1
        Exit Function
    End If
    getValCombo = CInt(Trim(Right(t, 5)))
End Function
Private Function ArmaProcedimiento(ByVal nomProcedure As String, ByRef params() As Param) As String ' creado por VAPI
    Dim sSql As String
    sSql = "EXEC " & nomProcedure & " "
    Dim i As Integer
    For i = 0 To UBound(params)
                
        If params(i).nuloNegativo Then
            If params(i).valor < 0 Then
                params(i).valor = "NULL"
            End If
        End If
        
        If params(i).comilla Then
            If params(i).nuloVacio Then
                If params(i).valor = "" Then
                    sSql = sSql & "NULL"
                Else
                    sSql = sSql & "'" & params(i).valor & "'"
                End If
            Else
                sSql = sSql & "'" & params(i).valor & "'"
            End If
        Else
            If UCase(params(i).valor) = "VERDADERO" Then
                params(i).valor = "1"
            ElseIf UCase(params(i).valor) = "FALSO" Then
                params(i).valor = "0"
            End If
            
            sSql = sSql & params(i).valor
        End If
        If Not (i = UBound(params)) Then
          sSql = sSql & ","
        End If
    Next i
    ArmaProcedimiento = sSql
End Function
'AGREGADO POR VAPI PRESOLICITUD DE CREDITOS 20170411 ERS 001-2017
Public Function ObtenerPreSolicitudes(ByVal psCuserAnalista As String, ByVal bAmpliado As Integer) As ADODB.Recordset
On Error GoTo Error
If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "EXEC stp_sel_hojaRuta_obtenerPreSolicitud '" & psCuserAnalista & "'," & bAmpliado
    Set ObtenerPreSolicitudes = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta ObtenerPreSolicitudes", Err.Description
End Function

Public Function ObtenerPreSolicitudesXid(ByVal nPresolicitudId As Integer) As ADODB.Recordset
On Error GoTo Error
If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "EXEC stp_sel_hojaRuta_obtenerPreSolicitudxID " & nPresolicitudId
    Set ObtenerPreSolicitudesXid = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta ObtenerPreSolicitudesXid", Err.Description
End Function

Public Function ActualizarcCtaPreSol(ByVal nPresolicitudId As Integer, ByVal cCtaCodSolicitud As String) As ADODB.Recordset
On Error GoTo Error
If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "EXEC stp_sel_upd_registarcCtaPreSol " & nPresolicitudId & ",'" & cCtaCodSolicitud & "'"
    Set ActualizarcCtaPreSol = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta ActualizarcCtaPreSol", Err.Description
End Function


'FIN AGREGADO POR VAPI
Public Function GeneraCarteraDiariaPromocion(ByVal psCperiodo As String, ByVal psCuserAnalista As String, ByVal bGeneraDiaSiguiente As Integer) As ADODB.Recordset
On Error GoTo Error
If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "EXEC stp_ERS0232015_generarCarteraDiariaPromocion '" & psCperiodo & "','" & psCuserAnalista & "'," & bGeneraDiaSiguiente
    Set GeneraCarteraDiariaPromocion = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta GeneraCarteraDiariaPromocion", Err.Description
End Function

Public Function GeneraCarteraDiariaMora(ByVal psCperiodo As String, ByVal psCuserAnalista As String, ByVal bGeneraDiaSiguiente As Integer) As ADODB.Recordset
On Error GoTo Error
If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "EXEC stp_ERS0232015_generarCarteraDiariaMora '" & psCperiodo & "','" & psCuserAnalista & "'," & bGeneraDiaSiguiente
    Set GeneraCarteraDiariaMora = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta GeneraCarteraDiariaMora", Err.Description
End Function


Public Function ObtenerHojaRutaDia(ByVal psCuserAnalista As String, ByVal bPromocion As String) As ADODB.Recordset
On Error GoTo Error
If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "EXEC stp_ERS0232015_sel_HojaRutaAsignacionDiariaAnalista '" & psCuserAnalista & "'," & bPromocion
    Set ObtenerHojaRutaDia = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta ObtenerHojaRutaDia", Err.Description
End Function

Public Sub RegistraVisitaDiariaPromocion(ByVal cCuserAnalista As String, _
                                         ByVal cPersCodCliente As String, _
                                         ByVal dFechaRegistro As String, _
                                         ByVal bGestionado As Boolean, _
                                         ByVal bEntrevistaTercero As Boolean, _
                                         ByVal cPersTercero As String, _
                                         ByVal nPersTerceroRelac As Integer, _
                                         ByVal cGlosaNoGestion As String, _
                                         ByVal nResultadoVisitaPromo As Integer, _
                                         ByVal dFechaProxVisita As String, _
                                         ByVal cGlosaResultadoPromo As String, _
                                         ByVal nMotivoNoGestionPromo As Integer, _
                                         ByVal nLineaRutaId As Integer)

On Error GoTo Error
If oCon.AbreConexion = False Then Exit Sub
    Dim sSql As String
    Dim p(12) As Param
    p(0).valor = cCuserAnalista: p(0).comilla = True
    p(1).valor = cPersCodCliente: p(1).comilla = True
    p(2).valor = dFechaRegistro: p(2).comilla = True
    p(3).valor = bGestionado: p(3).comilla = False
    p(4).valor = bEntrevistaTercero: p(4).comilla = False
    p(5).valor = cPersTercero: p(5).comilla = True: p(5).nuloVacio = True
    p(6).valor = nPersTerceroRelac: p(6).comilla = False: p(6).nuloNegativo = True
    p(7).valor = cGlosaNoGestion: p(7).comilla = True: p(7).nuloVacio = True
    p(8).valor = nResultadoVisitaPromo: p(8).comilla = False: p(8).nuloNegativo = True
    p(9).valor = dFechaProxVisita: p(9).comilla = True: p(9).nuloVacio = True
    p(10).valor = cGlosaResultadoPromo: p(10).comilla = True: p(10).nuloVacio = True
    p(11).valor = nMotivoNoGestionPromo: p(11).comilla = False: p(11).nuloNegativo = True
    p(12).valor = nLineaRutaId: p(12).comilla = False
    sSql = ArmaProcedimiento("stp_ERS0232015_ins_HojaRutaResultadoVisitaAnalistaPromo", p)
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Exit Sub
Error:
    Err.Raise Err.Number, "Hoja Ruta RegistraVisitaDiariaPromocion", Err.Description
End Sub


Public Sub RegistraVisitaDiariaPromocionNoPlaneada(ByVal cCuserAnalista As String, _
                                         ByVal cPersCodCliente As String, _
                                         ByVal dFechaRegistro As String, _
                                         ByVal bGestionado As Boolean, _
                                         ByVal bEntrevistaTercero As Boolean, _
                                         ByVal cPersTercero As String, _
                                         ByVal nPersTerceroRelac As Integer, _
                                         ByVal cGlosaNoGestion As String, _
                                         ByVal nResultadoVisitaPromo As Integer, _
                                         ByVal dFechaProxVisita As String, _
                                         ByVal cGlosaResultadoPromo As String, _
                                         ByVal nMotivoNoGestionPromo As Integer, _
                                         ByVal nLineaRutaId As Integer)

On Error GoTo Error
If oCon.AbreConexion = False Then Exit Sub
    Dim sSql As String
    Dim p(12) As Param
    p(0).valor = cCuserAnalista: p(0).comilla = True
    p(1).valor = cPersCodCliente: p(1).comilla = True
    p(2).valor = dFechaRegistro: p(2).comilla = True
    p(3).valor = bGestionado: p(3).comilla = False
    p(4).valor = bEntrevistaTercero: p(4).comilla = False
    p(5).valor = cPersTercero: p(5).comilla = True: p(5).nuloVacio = True
    p(6).valor = nPersTerceroRelac: p(6).comilla = False: p(6).nuloNegativo = True
    p(7).valor = cGlosaNoGestion: p(7).comilla = True: p(7).nuloVacio = True
    p(8).valor = nResultadoVisitaPromo: p(8).comilla = False: p(8).nuloNegativo = True
    p(9).valor = dFechaProxVisita: p(9).comilla = True: p(9).nuloVacio = True
    p(10).valor = cGlosaResultadoPromo: p(10).comilla = True: p(10).nuloVacio = True
    p(11).valor = nMotivoNoGestionPromo: p(11).comilla = False: p(11).nuloNegativo = True
    p(12).valor = nLineaRutaId: p(12).comilla = False
    sSql = ArmaProcedimiento("stp_ERS0232015_ins_HojaRutaResultadoVisitaAnalistaNoPlaPromo", p)
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Exit Sub
Error:
    Err.Raise Err.Number, "Hoja Ruta RegistraVisitaDiariaPromocionNoPlaneada", Err.Description
End Sub

Public Sub RegistraVisitaDiariaMora(ByVal cCuserAnalista As String, _
                                         ByVal cPersCodCliente As String, _
                                         ByVal dFechaRegistro As String, _
                                         ByVal bGestionado As Boolean, _
                                         ByVal bEntrevistaTercero As Boolean, _
                                         ByVal cPersTercero As String, _
                                         ByVal nPersTerceroRelac As Integer, _
                                         ByVal cGlosaNoGestion As String, _
                                         ByVal nResultadoVisitaMora As Integer, _
                                         ByVal nMotivoIncumplimiento As Integer, _
                                         ByVal bCompromiso As Boolean, _
                                         ByVal dFechaCompromiso As String, _
                                         ByVal nMonedaCompromiso As Integer, _
                                         ByVal nMontoCompromiso As Double, _
                                         ByVal nMotivoNoGestionMora As Integer, _
                                         ByVal nLineaRutaId As Integer)

On Error GoTo Error
If oCon.AbreConexion = False Then Exit Sub
    Dim sSql As String
    Dim p(15) As Param
    
    p(0).valor = cCuserAnalista: p(0).comilla = True
    p(1).valor = cPersCodCliente: p(1).comilla = True
    p(2).valor = dFechaRegistro: p(2).comilla = True
    p(3).valor = bGestionado: p(3).comilla = False
    p(4).valor = bEntrevistaTercero: p(4).comilla = False
    p(5).valor = cPersTercero: p(5).comilla = True: p(5).nuloVacio = True
    p(6).valor = nPersTerceroRelac: p(6).comilla = False: p(6).nuloNegativo = True
    p(7).valor = cGlosaNoGestion: p(7).comilla = True: p(7).nuloVacio = True
    p(8).valor = nResultadoVisitaMora: p(8).comilla = False: p(8).nuloNegativo = True
    p(9).valor = nMotivoIncumplimiento: p(9).comilla = False: p(9).nuloNegativo = True
    p(10).valor = bCompromiso: p(10).comilla = False
    p(11).valor = dFechaCompromiso: p(11).comilla = True: p(11).nuloVacio = True
    p(12).valor = nMonedaCompromiso: p(12).comilla = False: p(12).nuloNegativo = True
    p(13).valor = nMontoCompromiso: p(13).comilla = False: p(13).nuloNegativo = True
    p(14).valor = nMotivoNoGestionMora: p(14).comilla = False: p(14).nuloNegativo = True
    p(15).valor = nLineaRutaId: p(15).comilla = False: p(15).nuloNegativo = True
    sSql = ArmaProcedimiento("stp_ERS0232015_ins_HojaRutaResultadoVisitaAnalistaMora", p)
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Exit Sub
Error:
    Err.Raise Err.Number, "Hoja Ruta RegistraVisitaDiariaMora", Err.Description
End Sub

Public Function RegistraVisitaDiariaMoraNoPlaneada(ByVal cCuserAnalista As String, _
                                         ByVal cPersCodCliente As String, _
                                         ByVal dFechaRegistro As String, _
                                         ByVal bGestionado As Boolean, _
                                         ByVal bEntrevistaTercero As Boolean, _
                                         ByVal cPersTercero As String, _
                                         ByVal nPersTerceroRelac As Integer, _
                                         ByVal cGlosaNoGestion As String, _
                                         ByVal nResultadoVisitaMora As Integer, _
                                         ByVal nMotivoIncumplimiento As Integer, _
                                         ByVal bCompromiso As Boolean, _
                                         ByVal dFechaCompromiso As String, _
                                         ByVal nMonedaCompromiso As Integer, _
                                         ByVal nMontoCompromiso As Double, _
                                         ByVal nMotivoNoGestionMora As Integer, _
                                         ByVal nLineaRutaId As Integer) As Integer
 
On Error GoTo Error
If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    Dim p(15) As Param
    Dim rs As ADODB.Recordset
    p(0).valor = cCuserAnalista: p(0).comilla = True
    p(1).valor = cPersCodCliente: p(1).comilla = True
    p(2).valor = dFechaRegistro: p(2).comilla = True
    p(3).valor = bGestionado: p(3).comilla = False
    p(4).valor = bEntrevistaTercero: p(4).comilla = False
    p(5).valor = cPersTercero: p(5).comilla = True: p(5).nuloVacio = True
    p(6).valor = nPersTerceroRelac: p(6).comilla = False: p(6).nuloNegativo = True
    p(7).valor = cGlosaNoGestion: p(7).comilla = True: p(7).nuloVacio = True
    p(8).valor = nResultadoVisitaMora: p(8).comilla = False: p(8).nuloNegativo = True
    p(9).valor = nMotivoIncumplimiento: p(9).comilla = False: p(9).nuloNegativo = True
    p(10).valor = bCompromiso: p(10).comilla = False
    p(11).valor = dFechaCompromiso: p(11).comilla = True: p(11).nuloVacio = True
    p(12).valor = nMonedaCompromiso: p(12).comilla = False: p(12).nuloNegativo = True
    p(13).valor = nMontoCompromiso: p(13).comilla = False: p(13).nuloNegativo = True
    p(14).valor = nMotivoNoGestionMora: p(14).comilla = False: p(14).nuloNegativo = True
    p(15).valor = nLineaRutaId: p(15).comilla = False: p(15).nuloNegativo = True
    sSql = ArmaProcedimiento("stp_ERS0232015_ins_HojaRutaResultadoVisitaAnalistaNoPlaMora", p)
    Set rs = oCon.CargaRecordSet(sSql)
    RegistraVisitaDiariaMoraNoPlaneada = rs!nRes
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta RegistraVisitaDiariaMora", Err.Description
End Function

Public Function obtenerNumeroVisitasPendientes(ByVal cUserAnalista As String, ByVal bSoloDia) As Integer
On Error GoTo Error
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    Dim p(1) As Param
    Dim rs As ADODB.Recordset
    p(0).valor = cUserAnalista: p(0).comilla = True
    p(1).valor = bSoloDia: p(1).comilla = False
    sSql = ArmaProcedimiento("stp_ERS0232015_sel_numeroVisitasPendientes", p)
    Set rs = oCon.CargaRecordSet(sSql)
    obtenerNumeroVisitasPendientes = rs!nPendientes
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta obtenerNumeroVisitasPendientes", Err.Description
End Function
Public Function obtenerNumeroVisitasRegistradasHoy(ByVal cUserAnalista As String) As Integer
On Error GoTo Error
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    Dim p(0) As Param
    Dim rs As ADODB.Recordset
    p(0).valor = cUserAnalista: p(0).comilla = True
    sSql = ArmaProcedimiento("stp_ERS0232015_sel_numeroVisitasRegistradasHoy", p)
    Set rs = oCon.CargaRecordSet(sSql)
    obtenerNumeroVisitasRegistradasHoy = rs!nTotal
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta obtenerNumeroVisitasRegistradasHoy", Err.Description
End Function

Public Function esHoraLimite() As Boolean
On Error GoTo Error
    If oCon.AbreConexion = False Then Exit Function
    Dim rs As ADODB.Recordset
    Set rs = oCon.CargaRecordSet("EXEC stp_ERS0232015_sel_esHoraLimite")
    esHoraLimite = IIf(rs!nRes = 1, True, False)
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta esHoraLimite", Err.Description
End Function

Public Function haConfiguradoAgencia(ByVal cPeriodo As String, ByVal cCodAge As String) As Boolean
On Error GoTo Error
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    Dim p(1) As Param
    Dim rs As ADODB.Recordset
    p(0).valor = cPeriodo: p(0).comilla = True
    p(1).valor = cCodAge: p(1).comilla = True
    sSql = ArmaProcedimiento("stp_ERS0232015_sel_haConfiguradoAgencia", p)
    Set rs = oCon.CargaRecordSet(sSql)
    haConfiguradoAgencia = (rs!nCant > 0)
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta haConfiguradoAgencia", Err.Description
End Function

Public Sub solicitarVisto(ByVal cUserAnalista As String, ByVal cMovNroRegistro As String)
On Error GoTo Error
    If oCon.AbreConexion = False Then Exit Sub
    Dim sSql As String
    Dim p(1) As Param
    Dim rs As ADODB.Recordset
    
    With p(0)
        .valor = cUserAnalista
        .comilla = True
    End With
    
    With p(1)
        .valor = cMovNroRegistro
        .comilla = True
    End With
    
    sSql = ArmaProcedimiento("stp_ERS0232015_ins_SolicitarVisto", p)
    Set rs = oCon.Ejecutar(sSql)
    oCon.CierraConexion
    Exit Sub
Error:
    Err.Raise Err.Number, "Hoja Ruta solicitarVisto", Err.Description
End Sub
Public Function tieneVistoPendiente(ByVal cUserAnalista As String) As Boolean
On Error GoTo Error
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    Dim p(0) As Param
    Dim rs As ADODB.Recordset
    
    With p(0)
        .valor = cUserAnalista
        .comilla = True
    End With
    

    sSql = ArmaProcedimiento("stp_ERS0232015_sel_tieneVistoPendiente", p)
    Set rs = oCon.CargaRecordSet(sSql)
    tieneVistoPendiente = (rs!nPendientes > 0)
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta tieneVistoPendiente", Err.Description
End Function
Public Function obtenerVistosPendientes(ByVal cAgeCod As String) As ADODB.Recordset
On Error GoTo Error
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    Dim p(0) As Param
    Dim rs As ADODB.Recordset
    
    With p(0)
        .valor = cAgeCod
        .comilla = True
    End With
    
    sSql = ArmaProcedimiento("stp_ERS0232015_sel_ListarVistosPendientes", p)
    Set obtenerVistosPendientes = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta obtenerVistosPendientes", Err.Description
End Function

Public Sub darVisto(ByVal nIdVisto As Integer, ByVal cGlosa As String, ByVal cUserVisto As String)

On Error GoTo Error
If oCon.AbreConexion = False Then Exit Sub
    Dim sSql As String
    Dim p(2) As Param
    
    With p(0)
        .valor = nIdVisto
        .comilla = False
    End With
    
    With p(1)
        .valor = cGlosa
        .comilla = True
    End With
    
    With p(2)
        .valor = cUserVisto
        .comilla = True
    End With

    sSql = ArmaProcedimiento("stp_ERS0232015_upd_DarVisto", p)
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Exit Sub
Error:
    Err.Raise Err.Number, "Hoja Ruta darVisto", Err.Description
End Sub

Public Sub recibirVisto(ByVal cUserAnalista As String)

On Error GoTo Error
If oCon.AbreConexion = False Then Exit Sub
    Dim sSql As String
    Dim p(0) As Param

    With p(0)
        .valor = cUserAnalista
        .comilla = True
    End With

    sSql = ArmaProcedimiento("stp_ERS0232015_upd_RecibirVisto", p)
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Exit Sub
Error:
    Err.Raise Err.Number, "Hoja Ruta recibirVisto", Err.Description
End Sub

Public Function ObtenerConfNumVisitasAnalista(ByVal cUserAnalista As String, ByVal cPeriodo As String) As Boolean
On Error GoTo Error
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    Dim p(1) As Param
    Dim rs As ADODB.Recordset
    
    p(0).valor = cUserAnalista: p(0).comilla = True
    p(1).valor = cPeriodo: p(1).comilla = True
    sSql = ArmaProcedimiento("stp_sel_ObtenerConfNumVisitasAnalista", p)
    Set rs = oCon.CargaRecordSet(sSql)
    
    If Not (rs.EOF And rs.BOF) Then
        ObtenerConfNumVisitasAnalista = True
    Else
        ObtenerConfNumVisitasAnalista = False
    End If
    
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta ObtenerConfNumVisitasAnalista", Err.Description
End Function

Public Function ObtenerVarConfig(ByVal nCod As Integer) As String
On Error GoTo Error
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    Dim p(0) As Param
    Dim rs As ADODB.Recordset
    p(0).valor = nCod
    
    sSql = ArmaProcedimiento("stp_ERS0232015_sel_HojaRutaConfigCierre", p)
    Set rs = oCon.CargaRecordSet(sSql)
    
    If Not (rs.EOF And rs.BOF) Then
        ObtenerVarConfig = rs!cValor
    Else
        ObtenerVarConfig = ""
    End If
    
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta obtenerVistosPendientes", Err.Description
End Function

Public Function ObtieneAgenciaHojaRutaNew(ByVal psCodAge As String) As Boolean
On Error GoTo Error

    Dim sAgencias As String
    Dim sCod() As String
    Dim i As Integer

    sAgencias = ObtenerVarConfig(6)
    
    sCod = Split(sAgencias, ",")
    ObtieneAgenciaHojaRutaNew = False
    
    For i = 0 To UBound(sCod)
        If psCodAge = sCod(i) Then
            ObtieneAgenciaHojaRutaNew = True
            Exit Function
        End If
    Next i
    
    Exit Function
Error:
    MsgBox Err.Description, vbCritical, "Error"
End Function

'vapi: 20160523 para el bloqueo del sicmact por incumplimiento
Public Function puedeGenerar(ByVal cUser As String, ByVal bDiaSiguiente As Integer) As Integer
On Error GoTo Error
If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "EXEC stp_sel_HojaRutaPuedeGenerar '" & cUser & "'," & bDiaSiguiente
    puedeGenerar = oCon.CargaRecordSet(sSql)!nRes
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta puedeGenerar", Err.Description
End Function
Public Sub SolicitarVistoHojaRuta(ByVal cUserAnalista As String, ByVal cCodAge As String, ByVal nMotivo As Integer)
On Error GoTo Error
If oCon.AbreConexion = False Then Exit Sub
    Dim sSql As String
    sSql = "EXEC stp_ERS0232015_ins_SolicitarVisto '" & cUserAnalista & "','" & cCodAge & "'," & nMotivo
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Exit Sub
Error:
    Err.Raise Err.Number, "Hoja Ruta SolicitarVistoHojaRuta", Err.Description
End Sub

Public Function participaHojaRuta(ByVal cUser As String) As Boolean
On Error GoTo Error
If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "EXEC stp_sel_ParticipaHojaRuta '" & cUser & "'"
    participaHojaRuta = oCon.CargaRecordSet(sSql)!nCant > 0
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta participaHojaRuta", Err.Description
End Function

'MARG ERS027-2017***
Public Function puedeRechazarPresolicitud(ByVal cUser As String) As Boolean
On Error GoTo Error
If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "EXEC stp_sel_ERS027_2017_puedeRechazarPresolicitud '" & cUser & "'"
    puedeRechazarPresolicitud = oCon.CargaRecordSet(sSql)!res
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Hoja Ruta puedeRechazarPresolicitud", Err.Description
End Function
'END MARG************
'MARG ERS027-2017***
Public Function InsertarPresolicitudRechazo(pnPresolicitudId As Integer, pnConsValorMotivo As Integer, pcComentario As String, pcUserRegistro As String) As Integer
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim oRs As New ADODB.Recordset

    On Error GoTo ErrorInsertarInsertarPresolicitudRechazo

    sSql = " exec stp_ins_ERS027_2017_InsertarPresolicitudRechazo " & pnPresolicitudId & "," & pnConsValorMotivo & ",'" & pcComentario & "','" & pcUserRegistro & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set oRs = oCon.Ejecutar(sSql)
    If Not oRs.BOF And Not oRs.EOF Then
        If oRs.RecordCount > 0 Then
            InsertarPresolicitudRechazo = oRs!nId
        Else
            InsertarPresolicitudRechazo = 0
        End If
    Else
        InsertarPresolicitudRechazo = 0
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrorInsertarInsertarPresolicitudRechazo:
    Err.Raise Err.Number, "InsertarPresolicitudRechazo ", Err.Description
    InsertarPresolicitudRechazo = 0
End Function
'AGREGADO POR PTI1 ERS027-2017
Public Function existePresolicitud(sCodigo As String, sUsuario As String, bAmpliacion As Integer) As ADODB.Recordset
On Error GoTo Error
If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = " exec stp_sel_hojaRuta_obtenerPreSolicitudxUsuario '" & sCodigo & "','" & sUsuario & "'," & bAmpliacion
    Set existePresolicitud = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Exit Function
Error:
    Err.Raise Err.Number, "Obtener Presolicitud de usuario", Err.Description
End Function
'FIN AGREGADO PTI1

'END MARG ***********

