VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMServicioRecaudo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub EditarConvenioRecaudo( _
        ByVal sCodConvenio As String, ByVal nTipoCobro As Integer, _
        ByVal nTipoCalculo As Integer, ByVal nComision As Double, _
        ByVal nMinimo As Double, ByVal nMaximo As Double, _
        ByVal nDistribucion As Integer, ByVal nDistEmpresa As Double, _
        ByVal nDistCliente As Double, ByVal sCtaCod As String, _
        ByVal sNombreConvenio As String, ByVal sDescripcion As String, _
        ByVal cMovNro As String, ByVal nEstado As Integer, _
        ByVal nPagoMin As Double, ByVal nPagoMax As Double, _
        ByVal cCorreo As String, ByVal cContacto As String, _
        ByVal nFrecuencia As Integer, ByVal dProximaFecha As Date, _
        ByVal psConvenioCompartido As String, _
        ByVal nTipoMora As Integer, ByVal nMora As Double, _
        ByVal sServicio As String, ByVal bUnicode As Integer, _
        ByVal nResumenWeb As Integer, ByVal bTieneWS As Boolean) 'CTI1 ERS027-2019

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorEditarConvenioRecaudo
    
        sSql = "exec stp_upd_ConvenioRecaudo '"
        sSql = sSql & sCodConvenio & "',"
        sSql = sSql & nTipoCobro & ","
        sSql = sSql & nTipoCalculo & ","
        sSql = sSql & nComision & ","
        sSql = sSql & nMinimo & ","
        sSql = sSql & nMaximo & ","
        sSql = sSql & nDistribucion & ","
        sSql = sSql & nDistEmpresa & ","
        sSql = sSql & nDistCliente & ",'"
        sSql = sSql & sCtaCod & "','"
        sSql = sSql & sNombreConvenio & "','"
        sSql = sSql & sDescripcion & "',"
        sSql = sSql & nEstado & ", '"
        sSql = sSql & cMovNro & "', "
        sSql = sSql & nPagoMin & ", "
        sSql = sSql & nPagoMax & ", '"
        sSql = sSql & cCorreo & "', '"
        sSql = sSql & cContacto & "', "
        sSql = sSql & nFrecuencia & ", '"
        sSql = sSql & Format(dProximaFecha, "yyyymmdd") & "', '"
        sSql = sSql & psConvenioCompartido & "', "
        sSql = sSql & nTipoMora & ", "
        sSql = sSql & nMora & ", '" & sServicio & "',"
        sSql = sSql & bUnicode & "," & nResumenWeb & ","
        sSql = sSql & IIf(bTieneWS, 1, 0) 'CTI1 ERS027-2019
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
    
ErrorEditarConvenioRecaudo:
    Err.Raise Err.Number, "EditarConvenioRecaudo", "Error al editar convenio, contactarse con el area de TI"

End Sub

Public Sub AgregaConvenioRecaudo( _
        ByVal sPerscod As String, ByVal nTipoCobro As Integer, _
        ByVal nTipoCalculo As Integer, ByVal nComision As Double, _
        ByVal nMinimo As Double, ByVal nMaximo As Double, _
        ByVal nDistribucion As Integer, ByVal nDistEmpresa As Double, _
        ByVal nDistCliente As Double, ByVal nTipoConvenio As Integer, _
        ByVal sCtaCod As String, ByVal sNombreConvenio As String, _
        ByVal sDescripcion As String, ByVal sNroMov As String, _
        ByVal nEstado As Integer, ByVal nPagoMin As Double, _
        ByVal nPagoMax As Double, ByVal cCorreo As String, _
        ByVal cContacto As String, ByVal nFrecuencia As Integer, _
        ByVal dProximaFecha As Date, ByVal psConvenioCompartido As String, _
        ByVal nTipoMora As Integer, ByVal nMora As Double, ByVal sServicio As String, _
        ByVal bUnicode As Integer, ByVal nResumenWeb As Integer, _
        ByVal bTieneWS As Boolean) 'CTI1 ERS027-2019

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorAgregaConvenioRecaudo
    
        sSql = "exec stp_ins_ConvenioRecaudo '"
        sSql = sSql & sPerscod & "',"
        sSql = sSql & nTipoCobro & ","
        sSql = sSql & nTipoCalculo & ","
        sSql = sSql & nComision & ","
        sSql = sSql & nMinimo & ","
        sSql = sSql & nMaximo & ","
        sSql = sSql & nDistribucion & ","
        sSql = sSql & nDistEmpresa & ","
        sSql = sSql & nDistCliente & ","
        sSql = sSql & nTipoConvenio & ",'"
        sSql = sSql & sCtaCod & "','"
        sSql = sSql & sNombreConvenio & "','"
        sSql = sSql & sDescripcion & "','"
        sSql = sSql & sNroMov & "',"
        sSql = sSql & nEstado & ", "
        sSql = sSql & nPagoMin & ", "
        sSql = sSql & nPagoMax & ", '"
        sSql = sSql & cCorreo & "', '"
        sSql = sSql & cContacto & "', "
        sSql = sSql & nFrecuencia & ", '"
        sSql = sSql & Format(dProximaFecha, "yyyymmdd") & "', '"
        sSql = sSql & psConvenioCompartido & "', "
        sSql = sSql & nTipoMora & ", "
        sSql = sSql & nMora & ", '"
        sSql = sSql & sServicio & "',"
        sSql = sSql & bUnicode & "," & nResumenWeb & ","
        sSql = sSql & IIf(bTieneWS, 1, 0) 'CTI1 ERS027-2019
        'nMora
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.BeginTrans
        oConecta.CargaRecordSet (sSql)
        oConecta.CommitTrans
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
    
ErrorAgregaConvenioRecaudo:
    oConecta.RollbackTrans
    oConecta.CierraConexion
    Set oConecta = Nothing
    Err.Raise Err.Number, "Agregar Convenio", "Error al registrar nuevo convenio, contactarse con el area de TI"

End Sub

Public Function GetTipoProductoAhorro(ByVal sCuenta As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    On Error GoTo ErrorGetTipoProductoAhorro
    
        sSql = "exec stp_sel_TipoProductoAhorro '" & sCuenta & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
                
        oConecta.AbreConexion
        Set GetTipoProductoAhorro = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
                
        Exit Function
    
ErrorGetTipoProductoAhorro:
    Err.Raise Err.Number, "Tipo de Producto", "Error al verificar el tipo de producto en la BD, contactarse con el area de TI"

End Function
Public Function verificaPagoTrama(ByVal pnConvenio As String, ByVal pnCliente As String, ByVal pnTrama As String) As Boolean
    
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    On Error GoTo ErrorGetVerificaPagoTrama
    
        sSql = "exec stp_sel_VerificaPagoTrama '" & pnConvenio & "','" & pnCliente & "','" & pnTrama & "'"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set rs = oConecta.CargaRecordSet(sSql)
        If Not rs.EOF And Not rs.BOF Then
            If rs!resultado = 1 Then
                verificaPagoTrama = False
            Else
                verificaPagoTrama = True
            End If
        End If
        oConecta.CierraConexion
        Set oConecta = Nothing
                
        Exit Function
    
ErrorGetVerificaPagoTrama:
    Err.Raise Err.Number, "Pago Trama", "Error al verificar los pagos de cuotas tramas en la BD, contactarse con el area de TI"

End Function

Public Sub BajaConvenioRecaudo(ByVal sCodConvenio As String)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorBajaConvenioRecaudo
    
        sSql = "exec stp_upd_bajaConvenioRecaudo '"
        sSql = sSql & sCodConvenio & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
    
ErrorBajaConvenioRecaudo:

    Err.Raise Err.Number, "BajaConvenioRecaudo", "Error al dar de baja convenio, contactarse con el area de TI"

End Sub

Public Function AgregarCobroRecaudoDetalle( _
                    ByVal sCodConvenio As String, _
                    ByVal sId As String, _
                    ByVal sCodCliente As String, _
                    ByVal nTipoDoi As Integer, _
                    ByVal sDoi As String, _
                    ByVal sNomCliente As String, _
                    ByVal sServicio As String, _
                    ByVal sConcepto As String, _
                    ByVal nImporte As Double _
                    ) As Integer
    
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rsError As ADODB.Recordset
    
    On Error GoTo ErrorAgregarCobroRecaudoDetalle
    
        sSql = "exec stp_ins_CobroRecaudoDetalle '"
        sSql = sSql & sCodConvenio & "', '"
        sSql = sSql & sId & "', '"
        sSql = sSql & sCodCliente & "', "
        sSql = sSql & nTipoDoi & ", '"
        sSql = sSql & sDoi & "', '"
        sSql = sSql & sNomCliente & "', '"
        sSql = sSql & sServicio & "', '"
        sSql = sSql & sConcepto & "', "
        sSql = sSql & nImporte & ""
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set rsError = oConecta.CargaRecordSet(sSql)
        AgregarCobroRecaudoDetalle = CInt(rsError!Error)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Set rsError = Nothing
        Exit Function
    
ErrorAgregarCobroRecaudoDetalle:
    Err.Raise Err.Number, "AgregarCobroRecaudoDetalle", "Error al registrar detalle de cobro recaudo, contactarse con el area de TI"

End Function

Public Function CargaTemporal( _
    ByVal sCodConvenio As String, _
    ByVal sRutaTrama As String, _
    ByVal sRutaLog As String, _
    ByVal sRutaFormato As String) As ADODB.Recordset

    Dim sSql As String
    
    Dim oConecta As COMConecta.DCOMConecta
    Dim rsRespuestaCarga As ADODB.Recordset
    
    On Error GoTo Error
        
        sSql = sSql & "STP_INS_CARGARTRAMA '"
        sSql = sSql & sRutaTrama & "', '"
        sSql = sSql & sRutaLog & "', '"
        sSql = sSql & sRutaFormato & "', '"
        sSql = sSql & sCodConvenio & "' "
        
        
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        
        Set rsRespuestaCarga = oConecta.CargaRecordSet(sSql)
        
        oConecta.CierraConexion
        Set oConecta = Nothing
        Set CargaTemporal = rsRespuestaCarga
        Set rsRespuestaCarga = Nothing
        Exit Function
    
Error:

    Err.Raise Err.Number, "Carga de Trama a Temporal", "Error mientras se intentaba cargar Trama"

End Function

Public Function ValidarTrama( _
    ByVal sCodConvenio As String, _
    ByVal sTipValidacion As String, _
    ByVal nCantRegistros As Double, _
    ByVal sfecPrescripcion As String, _
    ByVal sMovNro As String, _
    ByVal sRutaTrama As String, _
    ByVal sRutaLog As String, _
    ByVal sRutaFormato As String) As ADODB.Recordset

    Dim sSql As String
    
    Dim sResultado As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rsErrores As ADODB.Recordset
    
    On Error GoTo ErrorAgregarCobroRecaudo

        'Aplicando Validaciones
        If sTipValidacion = "C" Then
            sSql = "exec stp_sel_ValidacionCompleta '" & sCodConvenio & "', " & nCantRegistros
        ElseIf sTipValidacion = "I" Then
            sSql = "exec stp_sel_ValidacionIncompleta '" & sCodConvenio & "', " & nCantRegistros
        ElseIf sTipValidacion = "P" Then
            sSql = "exec stp_sel_ValidacionImportes '" & sCodConvenio & "', " & nCantRegistros
        End If
   
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        
        'Aplicando Validaciones
        Set rsErrores = oConecta.CargaRecordSet(sSql)
            
        oConecta.CierraConexion
        Set oConecta = Nothing
        Set ValidarTrama = rsErrores
        Set rsErrores = Nothing
        Exit Function
    
ErrorAgregarCobroRecaudo:

    Err.Raise Err.Number, "Validacion de Trama", "Error durante el proceso de vaidacion, contactarse con el Area de TI"

End Function

Public Function CargarTrama( _
    ByVal sCodConvenio As String, _
    ByVal sTipValidacion As String, _
    ByVal nCantRegistros As Double, _
    ByVal sfecPrescripcion As String, _
    ByVal sMovNro As String, _
    ByVal sRutaTrama As String, _
    ByVal sRutaLog As String) As Boolean

    Dim sSql As String
    Dim sSql2 As String
    
    Dim oConecta As COMConecta.DCOMConecta
    Dim rsErrores As ADODB.Recordset
    
    On Error GoTo ErrorAgregarCobroRecaudo
    
    sSql = "exec stp_ins_reemplazarTramaConvenioCabecera '"
    sSql = sSql & sCodConvenio & "','"
    sSql = sSql & sTipValidacion & "',"
    sSql = sSql & nCantRegistros & ",'"
    sSql = sSql & sfecPrescripcion & "', '"
    sSql = sSql & sMovNro & "'"
            
    sSql2 = "exec stp_ins_CargarCobroDetalle '"
    sSql2 = sSql2 & sCodConvenio & "' "
               
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.BeginTrans
        
    'Actualizar cabecera de convenio
    oConecta.CargaRecordSet (sSql)
                
    'Carga Trama a Tabla Final
    oConecta.CargaRecordSet (sSql2)
                
    oConecta.CommitTrans
            
    oConecta.CierraConexion
    Set oConecta = Nothing
    Set rsErrores = Nothing
    CargarTrama = True
    Exit Function
    
ErrorAgregarCobroRecaudo:
    oConecta.RollbackTrans
    CargarTrama = False

End Function


Public Function GetInteres(ByVal nCapital As Double, ByVal nTasa As Double, _
            ByVal nPlazo As Long, Optional nTipoInteres As TipoCalculoInteres = TpoCalcIntSimple) As Double

Dim nIntEfe As Double
    If nTipoInteres = TpoCalcIntSimple Then
        GetInteres = Round((nTasa / 36000) * nPlazo * nCapital, 2)
    ElseIf nTipoInteres = TpoCalcIntCompuesto Then
        GetInteres = Round(((nTasa / 36000 + 1) ^ nPlazo - 1) * nCapital, 2)
    ElseIf nTipoInteres = TpoCalcIntAdelantado Then
        nIntEfe = ((nTasa / 36000 + 1) ^ nPlazo - 1)
        nIntEfe = nIntEfe / (1 + nIntEfe)
        GetInteres = Round(nIntEfe * nCapital, 2)
    End If
    
End Function

Public Function getMontoConceptoServicio(Optional ByVal strCodigoConcepto As String, _
                                          Optional ByVal strCodigoConvenio As String) As _
                                          ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorMontoConcepto
        sSql = "exec stp_sel_consultaMontoDeudaXConvenio "
        sSql = sSql & " '" & strCodigoConcepto & "', "
        sSql = sSql & " '" & strCodigoConvenio & "' "
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set getMontoConceptoServicio = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorMontoConcepto:
    Set getMontoConceptoServicio = Nothing
    Err.Raise Err.Number, "Buscar Monto ", Err.Description

End Function

Public Function GetBuscaConvenioXCodigo(Optional ByVal strCodigoConvevio As String) As ADODB.Recordset
    Dim sSql As String
    
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo errorGetBuscaConvenioXCodigo
        sSql = "exec stp_sel_consultaConvenioXCodigo"
        sSql = sSql & " '" & strCodigoConvevio & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set GetBuscaConvenioXCodigo = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
        
errorGetBuscaConvenioXCodigo:
    Set GetBuscaConvenioXCodigo = Nothing
    Err.Raise Err.Number, "Buscar Convenio por codigo", Err.Description
    
End Function

Public Function GetBUscarConvenioXCodigo(ByVal strCodigoConvenio As String)
     Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorGetBUscarConvenioXCodigo
    
        sSql = "exec stp_sel_consultaConvenioRecaudoXCodigo "
        sSql = sSql & " '" & strCodigoConvenio & "'"
             
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set GetBUscarConvenioXCodigo = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
    
ErrorGetBUscarConvenioXCodigo:
    Set GetBUscarConvenioXCodigo = Nothing
    Err.Raise Err.Number, "Buscar Convenio ", Err.Description
    
End Function

Public Function GetBuscarConvenio(Optional ByVal strNombreEmprea As String = "", _
                                  Optional ByVal strDescripcionConvenio As String = "", _
                                  Optional ByVal strNombreConvenio As String = "%") _
                                  As ADODB.Recordset
    
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorGetBuscarConvenio
    
        sSql = "exec stp_sel_buscarConvenioRecaudo "
        sSql = sSql & " '" & strNombreEmprea & "','" & strDescripcionConvenio & "', '"
        sSql = sSql & strNombreConvenio & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set GetBuscarConvenio = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
    
ErrorGetBuscarConvenio:
    Set GetBuscarConvenio = Nothing
    Err.Raise Err.Number, "Buscar Convenio ", Err.Description
    
End Function

Public Function GetBuscarUsuarioRecaudo(Optional ByVal strNombreCliente As String = "", _
                                        Optional ByVal strDOI As String = "", _
                                        Optional ByVal strCodigoCliente As String = "", _
                                        Optional ByVal strCodigoConvenio As String = "", _
                                        Optional ByVal BuscarPor As Integer = 3) _
                                        As ADODB.Recordset
    'Buscar Por:
    ' 1 = Codigo
    ' 2 = DOI
    ' 3 = Nombre
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo Error
        If BuscarPor = 1 Then
            sSql = "exec stp_sel_consultaUsuarioConvenioByCodigo "
            sSql = sSql & " '" & strCodigoConvenio & "', '" & strCodigoCliente & "'"
        ElseIf BuscarPor = 2 Then
            sSql = "exec stp_sel_consultaUsuarioConvenioByDOI "
            sSql = sSql & " '" & strCodigoConvenio & "', '" & strDOI & "'"
        ElseIf BuscarPor = 3 Then
            sSql = "exec stp_sel_consultaUsuarioConvenio "
            sSql = sSql & " '" & strNombreCliente & "', '" & strDOI & "','" & _
            strCodigoCliente & "','" & strCodigoConvenio & "'"
        End If
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set GetBuscarUsuarioRecaudo = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
Error:
    Set GetBuscarUsuarioRecaudo = Nothing
    Err.Raise Err.Number, "Buscar Usuario ", Err.Description
End Function

Public Function listaMovExtornoServicioRecaudo(ByVal fecha As String, Optional ByVal nroMovi As String = "", _
                                               Optional nroConvenio As String = "", _
                                               Optional sNombreConvenio As String = "") As Recordset
    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    If nroMovi = "" Then nroMovi = "%%"
    If nroConvenio = "" Then nroConvenio = "%%"
    If sNombreConvenio = "" Then sNombreConvenio = "%%"
    
    On Error GoTo ErrorListaMovExtorno
    sql = "exec stp_sel_listaMovExtornarServicioRecaudo '" & fecha & "','" & gDepositoRecaudo & "','" & _
          gIFTRecaudo & "','" & gCargoEmpresaRecaudo & "','"
    
    sql = sql & nroMovi & "','"
    sql = sql & nroConvenio & "', '"
    sql = sql & sNombreConvenio & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    
    Set listaMovExtornoServicioRecaudo = oConecta.CargaRecordSet(sql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorListaMovExtorno:
    Set listaMovExtornoServicioRecaudo = Nothing
    Err.Raise Err.Number, "Lista de Movimiento para Extornar", Err.Description
    
End Function

Public Function getListaConceptosCobrarXConvenioPV(Optional ByVal strCodigoConvenio As String) _
                                                                As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorListaConceptos
        sSql = "exec stp_sel_listaConceptosCobrarRecaudoPV "
        sSql = sSql & " '" & strCodigoConvenio & "' "
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set getListaConceptosCobrarXConvenioPV = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorListaConceptos:
    Set getListaConceptosCobrarXConvenioPV = Nothing
    Err.Raise Err.Number, "Buscar Usuario ", Err.Description
    
End Function

Public Function getListaPagosRecaudo(ByVal ctipo As String, ByVal cCodConvenio As String) As Recordset
    
    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorListaPagos
        sql = "exec stp_sel_listaPagosConvenio "
        sql = sql & "'" & ctipo & "', "
        sql = sql & "'" & cCodConvenio & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set getListaPagosRecaudo = oConecta.CargaRecordSet(sql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
        
ErrorListaPagos:
    Set getListaPagosRecaudo = Nothing
    Err.Raise Err.Number, "Buscar Monto ", Err.Description
        
End Function

Public Function getListaPagosRecaudoSV(ByVal sCodConvenio As String, _
                                       ByVal dFechaDesde As String, _
                                       ByVal dFechaHasta As String) As Recordset
    
    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorListaPagos
        sql = "exec stp_sel_listaPagosConvenioSV "
        sql = sql & "'" & sCodConvenio & "', "
        sql = sql & "'" & Format$(dFechaDesde, "yyyy/mm/dd") & "', "
        sql = sql & "'" & Format$(dFechaHasta, "yyyy/mm/dd") & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set getListaPagosRecaudoSV = oConecta.CargaRecordSet(sql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
        
ErrorListaPagos:
    Set getListaPagosRecaudoSV = Nothing
    Err.Raise Err.Number, "Buscar Monto ", Err.Description
        
End Function

Public Function actualizaXEnvioReporteCobros(ByVal nNumeroCobro As Integer, _
                                             ByVal fechaReporte As Date, ByVal oConecta As COMConecta.DCOMConecta) As Boolean
                                            
    Dim sql As String
    Dim bRespuesta As Boolean
    
    On Error GoTo ErrorActualizaReporteEnvio
        
        sql = "exec stp_upt_envioCobroRecaudoDetalle "
        sql = sql & nNumeroCobro & ","
        sql = sql & "'" & Format$(fechaReporte, "yyyymmdd") & "'"
                
        oConecta.CargaRecordSet sql
        
        bRespuesta = True
        
        actualizaXEnvioReporteCobros = bRespuesta
        Exit Function
        
ErrorActualizaReporteEnvio:

    actualizaXEnvioReporteCobros = False
    Err.Raise Err.Number, "Actualizar Generacion de Reporte ", Err.Description
    
End Function

Public Function registraOperacion(ByVal cCodConvenio As String)
            
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo registraOperacion
    
        sSql = "exec stp_upd_IngresarOperacionConvenio '"
        sSql = sSql & cCodConvenio & "'"
               
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
    
registraOperacion:
    Err.Raise Err.Number, "Agregar Convenio", "Error al registrar nuevo convenio, contactarse con el area de TI"
    
End Function

Public Function getCantidadRegistrosConvenio(ByVal sCodConvenio As String) As Boolean

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rsRegistros As ADODB.Recordset
    
    On Error GoTo Error

    Set oConecta = New COMConecta.DCOMConecta
    sSql = "exec stp_sel_CantidadRegistrosConvenio '"
    sSql = sSql & sCodConvenio & "' "
    oConecta.AbreConexion
    Set rsRegistros = oConecta.CargaRecordSet(sSql)
    If Not rsRegistros.EOF And Not rsRegistros.BOF Then
        If rsRegistros!nValor >= 1 Then
            getCantidadRegistrosConvenio = True
        Else
            getCantidadRegistrosConvenio = False
        End If
    Else
        getCantidadRegistrosConvenio = False
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
Error:
    getCantidadRegistrosConvenio = False
    Err.Raise Err.Number, "Verificando el Numero de Registros de un Convenio", Err.Description
End Function

Public Function getImporteMinMax(ByVal sCodConvenio As String, ByVal sId As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rsImporteMinMax As ADODB.Recordset
    
    On Error GoTo Error

    Set oConecta = New COMConecta.DCOMConecta
    
    sSql = "stp_sel_ObtenerImporteConvenioMinMax '"
    sSql = sSql & sCodConvenio & "', '" & sId & "' "

    oConecta.AbreConexion
    Set rsImporteMinMax = oConecta.CargaRecordSet(sSql)
    
    If Not rsImporteMinMax Is Nothing Then
        If Not rsImporteMinMax.EOF And Not rsImporteMinMax.BOF Then
            Set getImporteMinMax = rsImporteMinMax
        Else
            Set getImporteMinMax = Nothing
        End If
    Else
        Set getImporteMinMax = Nothing
    End If
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
        
Error:

    Set getImporteMinMax = Nothing
    Err.Raise Err.Number, "Verificando el Numero de Registros de un Convenio", Err.Description

End Function

'RIRO20150220    *************************************************************
Public Function CalculaComisionRecaudo(ByVal psCodConvenio As String, _
                                       ByVal psCodCliente As String, _
                                       ByVal pvConceptos As Variant, _
                                       ByRef pnComiCli As Double, _
                                       ByRef pnComiEmp As Double) As Boolean

On Error GoTo Error
    
    Dim bComision As Boolean
    Dim rsConvenio As New ADODB.Recordset
    
    'variables de Tipo de Cobro: Operacion / Concepto
    Dim nTipoCobro As Integer
    
    'variables de Calculo de comision
    Dim nCalculoComision As Integer
    Dim nComision As Double
    Dim nComMin As Double, nComMax As Double
    
    'variables distribucion de la comision
    Dim nDistComision As Integer
    Dim nDistEmp As Double, nDistCliente As Double
    
    'Otras Varibles ***************º
    Dim nFilas As Integer
    Dim nColumnas As Integer
    Dim i As Integer
    Dim nTemporal As Double
    'fin de otras variables *******
    
    Set rsConvenio = GetBUscarConvenioXCodigo(psCodConvenio)
    pnComiCli = 0
    pnComiEmp = 0
    
    nTipoCobro = rsConvenio!nTipoCobro
    nCalculoComision = rsConvenio!nTipoCalculo
    nDistComision = rsConvenio!nDistribucion
    
    nComision = rsConvenio!nComision
    nComMin = rsConvenio!nMinimo
    nComMax = rsConvenio!nMaximo
    nDistCliente = rsConvenio!nDistCliente
    nDistEmp = rsConvenio!nDistEmpresa
    
    nFilas = UBound(pvConceptos, 2) ' numero de filas
    nColumnas = UBound(pvConceptos, 1) ' numero de columnas
    
    'Cobro por Concepto ******************************************************
    If nTipoCobro = 0 Then
    
        'Comision Fija
        If nCalculoComision = 0 Then
            'Comision Fija == > Distribución fija
            If nDistComision = 0 Then
                pnComiCli = nFilas * nDistCliente
                pnComiEmp = nFilas * nDistEmp
            'Comision Fija == > Distribución porcentual
            Else
                pnComiCli = nFilas * (nDistCliente / 100#) * nComision
                pnComiEmp = nFilas * (nDistEmp / 100#) * nComision
            End If
        'Comision Procentual
        Else
            'Comision Procentual == > Distribución Porcentual
            If nDistComision = 1 Then
                For i = 1 To nFilas
                    nTemporal = (nComision / 100#) * (pvConceptos(2, i) + pvConceptos(3, i))
                    If nTemporal < nComMin Then
                        nTemporal = nComMin
                    ElseIf nTemporal > nComMax Then
                        nTemporal = nComMax
                    End If
                    pnComiEmp = pnComiEmp + nTemporal * (nDistEmp / 100#)
                    pnComiCli = pnComiCli + nTemporal * (nDistCliente / 100#)
                Next
            End If
        End If
    'Cobro por Operacion *****************************************************
    ElseIf nTipoCobro = 1 Then
        'Comision Fija
        If nCalculoComision = 0 Then
        
            'Comision Fija == > Distribución fija
            If nDistComision = 0 Then
                pnComiCli = nDistCliente
                pnComiEmp = nDistEmp
            
            'Comision Fija == > Distribución porcentual
            Else
                pnComiCli = (nDistCliente / 100#) * nComision
                pnComiEmp = (nDistEmp / 100#) * nComision
            
            End If
        'Comision Porcentual
        Else
            'Comision Porcentual == > Distribución porcentual
            If nDistComision = 1 Then
                For i = 1 To nFilas
                    nTemporal = nTemporal + pvConceptos(2, i) + pvConceptos(3, i)
                Next
                nTemporal = nTemporal * (nComision / 100#)
                If nTemporal < nComMin Then
                    nTemporal = nComMin
                ElseIf nTemporal > nComMax Then
                    nTemporal = nComMax
                End If
            
                pnComiCli = (nDistCliente / 100#) * nTemporal
                pnComiEmp = (nDistEmp / 100#) * nTemporal
            End If
        End If
    End If
    CalculaComisionRecaudo = True
    Exit Function
Error:
    CalculaComisionRecaudo = False
End Function
'END RIRO ************************************************************************


