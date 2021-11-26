VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMSrvRecaudoWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents Req As WinHttp.WinHttpRequest
Attribute Req.VB_VarHelpID = -1

Public Function VerificarConvenioRecaudoWebService(ByVal strCodigoConvenio As String) As Boolean
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rsRegistros As ADODB.Recordset
    On Error GoTo Error
            sSql = "exec stp_sel_VerificarConvenioRecaudoWebService '" & strCodigoConvenio & "'"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set rsRegistros = oConecta.CargaRecordSet(sSql)
        VerificarConvenioRecaudoWebService = rsRegistros!bTieneWS
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
Error:
     VerificarConvenioRecaudoWebService = False
    Err.Raise Err.Number, "VerificarConvenioRecaudoWebService ", Err.Description
End Function

Public Function ObtenerCodConvenioRecaudoWebService(ByVal strCodigoConvenio As String) As Integer
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rsRegistros As ADODB.Recordset
    On Error GoTo Error
            sSql = "exec stp_sel_ObtenerCodConvenioRecaudoWebService '" & strCodigoConvenio & "'"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set rsRegistros = oConecta.CargaRecordSet(sSql)
         ObtenerCodConvenioRecaudoWebService = rsRegistros!nCodConvenioRecaudoWS
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
Error:
    ObtenerCodConvenioRecaudoWebService = 0
    Err.Raise Err.Number, "ObtenerCodConvenioRecaudoWebService ", Err.Description
End Function

Public Function ConsultarSuministroELUC(ByVal codSuministro As String, ByVal urlSimaynas As String) As ADODB.Recordset
    Dim rFlag As ADODB.Recordset
    Dim postData As String
    On Error GoTo ErrorConsultarSuministroELUC
    
    Set Req = New WinHttp.WinHttpRequest
    postData = "codigoSuministro=" & codSuministro
    
    With Req
        .Open "POST", urlSimaynas & "/ServicioRecaudoELUC/consultar_Suministro?" & postData, Async:=False 'Descomentar al pasar a Produccion
        '.Open "POST", "http://localhost:65229/ServicioRecaudoELUC/consultar_Suministro?" & postData, Async:=False 'Pruebas
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=""UTF-8"""
        .Send
    End With
    
    Set rFlag = New ADODB.Recordset
    rFlag.Fields.Append "cFlag", adBSTR, 10
    
    If Req.Status = 200 Then
        If Req.ResponseText = "NoResult" Then
            rFlag.Open
            Set ConsultarSuministroELUC = rFlag
            Exit Function
        End If
        Set ConsultarSuministroELUC = ConsultarSuministroELUCResponse(Req.ResponseText)
    Else
        MsgBox "No se pudo realizar la consulta satisfactoriamente", vbExclamation, "Mensaje"
        rFlag.Open
        Set ConsultarSuministroELUC = rFlag
        Exit Function
    End If
    Exit Function
ErrorConsultarSuministroELUC:
    Set ConsultarSuministroELUC = Nothing
    Err.Raise Err.Number, "ConsultarSuministroELUC ", Err.Description
End Function

Private Function ConsultarSuministroELUCResponse(ByVal response As String) As ADODB.Recordset
' procesar la respuesta recibida del servicio Web
Dim R As ADODB.Recordset
Dim p As Object
Dim sJsonData As String
Dim i As Integer

sJsonData = Replace(response, """", "")

Dim codigoSuministro As Long
Dim codigoComprobante As String
Dim RazonSocial As String
Dim CodigoPeriodoComercial As Long
Dim MontoTotal As Currency
Dim FechaVencimiento As Date

Set p = JSON.parse(sJsonData)
If JSON.GetParserErrors <> "" Then
    MsgBox JSON.GetParserErrors, vbOKCancel, "Error"
    Exit Function
End If

Set R = New ADODB.Recordset
' Define campos
With R.Fields
    .Append "cNomCliente", adBSTR, 500
    .Append "nTipoDoi", adSmallInt
    .Append "cDoi", adBSTR, 8
    .Append "cConcepto", adBSTR, 20
    .Append "cCodCliente", adBSTR, 20
    .Append "nMora", adCurrency
    .Append "cCodigoComprobante", adBSTR, 50
    .Append "nMontoTotal", adCurrency
    .Append "cFechaVencimiento", adBSTR, 10
End With

' Abre el recordset desconectado
R.Open

For i = 1 To p.Count
    codigoSuministro = CLng(Trim(p.Item(i).Item("CodigoSuministro")))
    codigoComprobante = Trim(p.Item(i).Item("CodigoComprobante"))
    RazonSocial = Trim(p.Item(i).Item("RazonSocial"))
    CodigoPeriodoComercial = CLng(Trim(p.Item(i).Item("CodigoPeriodoComercial")))
    MontoTotal = CCur(Trim(p.Item(i).Item("MontoTotal")))
    FechaVencimiento = CDate(Trim(p.Item(i).Item("FechaVencimiento")))
    
    ' Agregar data al recordset
    R.AddNew _
        Array("cNomCliente", "nTipoDoi", "cDoi", "cConcepto", "cCodCliente", "nMora", _
              "cCodigoComprobante", "nMontoTotal", "cFechaVencimiento"), _
        Array(RazonSocial, 1, "00000000", CodigoPeriodoComercial, codigoSuministro, 0, _
              codigoComprobante, MontoTotal, FechaVencimiento)
Next i

Set ConsultarSuministroELUCResponse = R

End Function

'CTI7 INI Se debe listar solo 1 de los pagos parciales
Public Function ConsultarSuministroELUCFirstRecord(ByVal records As ADODB.Recordset) As ADODB.Recordset
Dim R As ADODB.Recordset
Set R = New ADODB.Recordset

With R.Fields
    .Append "cNomCliente", adBSTR, 500
    .Append "nTipoDoi", adSmallInt
    .Append "cDoi", adBSTR, 8
    .Append "cConcepto", adBSTR, 20
    .Append "cCodCliente", adBSTR, 20
    .Append "nMora", adCurrency
    .Append "cCodigoComprobante", adBSTR, 50
    .Append "nMontoTotal", adCurrency
    .Append "cFechaVencimiento", adBSTR, 10
End With
R.Open

If records.RecordCount > 0 Then
    Do While Not records.EOF
        If R.RecordCount = 0 Then
            R.AddNew _
                Array("cNomCliente", "nTipoDoi", "cDoi", "cConcepto", "cCodCliente", "nMora", _
                      "cCodigoComprobante", "nMontoTotal", "cFechaVencimiento"), _
                Array(records!cNomCliente, records!nTipoDoi, records!cDoi, records!cConcepto, records!cCodCliente, records!nMora, _
                      records!cCodigoComprobante, records!nMontoTotal, records!cFechaVencimiento)
        End If
        records.MoveNext
    Loop
    If records.EOF Then
        records.MoveFirst
    End If
End If
Set ConsultarSuministroELUCFirstRecord = R
End Function
'CTI7 INI Se debe listar solo 1 de los pagos parciales

'CTI7 CAMBIO PARAMETROS
'Public Function GenerarComprobanteELUC(ByVal CodigoCobranza As String, ByVal codigoComprobante As String, _
'                                       ByVal codigoCaja As String, ByVal codigoContratista As String, _
'                                       ByVal codigoContratistaContrato As String, ByVal montoCobrado As String, _
'                                       ByVal montoRecibido As String, ByVal fechaCobranza As String, _
'                                       ByVal urlSimaynas As String, ByVal sCodConvenio As String, _
'                                       ByVal sCodCliente As String) As String
Public Function GenerarComprobanteELUC(ByVal codigoSuministro As String, ByVal codigoComprobante As String, _
                                       ByVal montoRecibido As String, ByVal cMovNro As String, ByVal urlSimaynas As String) As String
    Dim postData As String
    On Error GoTo ErrorGenerarComprobanteELUC
    
    Set Req = New WinHttp.WinHttpRequest
    postData = "codigoSuministro=" & codigoSuministro & "&codigoComprobante=" & codigoComprobante & _
               "&montoRecibido=" & montoRecibido & "&pcMovNro=" & cMovNro ' & "&fechaCobranza=" & fechaCobranza'
               '"&codigoCaja=" & codigoCaja & "&codigoContratista=" & codigoContratista & _'
               '"&codigoContratistaContrato=" & codigoContratistaContrato & "&montoCobrado=" & montoCobrado & _'
    With Req
        .Open "POST", urlSimaynas & "/ServicioRecaudoELUC/generar_Comprobante?" & postData, Async:=False 'Descomentar al pasar a Produccion
        '.Open "POST", "http://localhost:65229/ServicioRecaudoELUC/generar_Comprobante?" & postData, Async:=False 'Pruebas
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=""UTF-8"""
        .Send
    End With
    
    If Req.Status = 200 Then
        If Req.ResponseText = "NoResult" Then
            GenerarComprobanteELUC = ""
            Exit Function
        End If
        'CTI7
        'GenerarComprobanteELUC = GenerarComprobanteELUCResponse(Req.ResponseText, sCodConvenio, _
        '                                                        codigoComprobante, sCodCliente)
        GenerarComprobanteELUC = GenerarComprobanteELUCResponse(Req.ResponseText)
    Else
        MsgBox "No se pudo realizar satisfactoriamente la operación", vbExclamation, "Mensaje"
        GenerarComprobanteELUC = ""
        Exit Function
    End If
    Exit Function
ErrorGenerarComprobanteELUC:
    GenerarComprobanteELUC = ""
    Err.Raise Err.Number, "GenerarComprobanteELUC ", Err.Description
End Function

'CTI7
'Private Function GenerarComprobanteELUCResponse(ByVal response As String, ByVal sCodConvenio As String, _
'                                                ByVal sCodigoComprobante As String, ByVal sCodCliente As String) As String
Private Function GenerarComprobanteELUCResponse(ByVal response As String) As String
' procesar la respuesta recibida del servicio Web
Dim R As ADODB.Recordset
Dim p As Object
Dim sJsonData As String
Dim i As Integer

sJsonData = Replace(response, """", "")

Dim sCodigoCobranza As String
Dim sMensaje As String

Set p = JSON.parse(sJsonData)
If JSON.GetParserErrors <> "" Then
    MsgBox JSON.GetParserErrors, vbOKCancel, "Error"
    GenerarComprobanteELUCResponse = False
    Exit Function
End If

For i = 1 To p.Count
    sCodigoCobranza = Trim(p.Item(i).Item("CodigoCobranza"))
    sMensaje = Trim(p.Item(i).Item("Mensaje"))
Next i

'CTI7
'If sMensaje = "CORRECTO" And sCodigoCobranza <> "0" Then
If sMensaje = "EXITOSO" And sCodigoCobranza <> "0" Then
    GenerarComprobanteELUCResponse = sCodigoCobranza
Else
    GenerarComprobanteELUCResponse = ""
End If

End Function

Public Function ObtenerValoresRecaudoWebService(ByVal nCodConvenioRecaudoWS As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo Error
            sSql = "exec stp_sel_ObtenerValoresServicioWeb " & nCodConvenioRecaudoWS
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtenerValoresRecaudoWebService = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
Error:
     Set ObtenerValoresRecaudoWebService = Nothing
    Err.Raise Err.Number, "ObtenerValoresRecaudoWebService ", Err.Description
End Function

'CTI7
'Public Function AnularComprobanteELUC(ByVal CodigoCobranza As String, ByVal justificacion As String, _
'                                       ByVal loginUsuario As String, ByVal urlSimaynas As String) As Boolean
Public Function AnularComprobanteELUC(ByVal codigoSuministro As String, ByVal nMovNro As Double, ByVal urlSimaynas As String) As Boolean
    Dim postData As String
    On Error GoTo ErrorAnularComprobanteELUC
    
    Set Req = New WinHttp.WinHttpRequest
    postData = "codigoSuministro=" & codigoSuministro & "&pnMovNro=" & nMovNro '&
               '"&loginUsuario=" & loginUsuario
    With Req
        .Open "POST", urlSimaynas & "/ServicioRecaudoELUC/anular_Comprobante?" & postData, Async:=False 'Descomentar al pasar a Produccion
        '.Open "POST", "http://localhost:65229/ServicioRecaudoELUC/anular_Comprobante?" & postData, Async:=False 'Pruebas
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=""UTF-8"""
        .Send
    End With
    
    If Req.Status = 200 Then
        If Req.ResponseText = "NoResult" Then
            AnularComprobanteELUC = False
            Exit Function
        End If
        
        If Req.ResponseText = "CORRECTO" Then
            AnularComprobanteELUC = True
        Else
            AnularComprobanteELUC = False
        End If
    Else
        MsgBox "No se pudo realizar satisfactoriamente la operación", vbExclamation, "Mensaje"
        AnularComprobanteELUC = False
        Exit Function
    End If
    Exit Function
ErrorAnularComprobanteELUC:
    AnularComprobanteELUC = False
    Err.Raise Err.Number, "AnularComprobanteELUC ", Err.Description
End Function

Public Function ObtenerCodigoCobranzaWS(ByVal sCodConvenio As String, ByVal sCodCliente As String) As String
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rsRegistros As ADODB.Recordset
    On Error GoTo Error
            sSql = "exec stp_sel_ObtenerCodigoCobranzaWS '" & sCodConvenio & "','" & sCodCliente & "'"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set rsRegistros = oConecta.CargaRecordSet(sSql)
        ObtenerCodigoCobranzaWS = rsRegistros!sCodigoCobranza
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
Error:
    ObtenerCodigoCobranzaWS = ""
    Err.Raise Err.Number, "ObtenerCodigoCobranzaWS ", Err.Description
End Function

'CTI1 ERS027-2019 ELOR
Public Function ConsultarSuministroELOR(ByVal codSuministro As String, ByVal urlSimaynas As String) As ADODB.Recordset
    Dim rFlag As ADODB.Recordset
    Dim postData As String
    On Error GoTo ErrorConsultarSuministroELOR
    
    Set Req = New WinHttp.WinHttpRequest
    postData = "codigoSuministro=" & codSuministro
    
    With Req
        .Open "POST", urlSimaynas & "/ServicioRecaudoELOR/consultar_Suministro?" & postData, Async:=False 'Descomentar al pasar a Produccion
        '.Open "POST", "http://localhost:65229/ServicioRecaudoELOR/consultar_Suministro?" & postData, Async:=False 'Pruebas
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=""UTF-8"""
        .Send
    End With
    
    Set rFlag = New ADODB.Recordset
    rFlag.Fields.Append "cFlag", adBSTR, 10
    
    If Req.Status = 200 Then
        If Req.ResponseText = "NoResult" Then
            rFlag.Open
            Set ConsultarSuministroELOR = rFlag
            Exit Function
        End If
        Set ConsultarSuministroELOR = ConsultarSuministroELORResponse(Req.ResponseText)
    Else
        MsgBox "No se pudo realizar la consulta satisfactoriamente", vbExclamation, "Mensaje"
        rFlag.Open
        Set ConsultarSuministroELOR = rFlag
        Exit Function
    End If
    Exit Function
ErrorConsultarSuministroELOR:
    Set ConsultarSuministroELOR = Nothing
    Err.Raise Err.Number, "ConsultarSuministroELOR", Err.Description
End Function

Private Function ConsultarSuministroELORResponse(ByVal response As String) As ADODB.Recordset
' procesar la respuesta recibida del servicio Web
Dim R As ADODB.Recordset
Dim p As Object
Dim sJsonData As String
Dim i As Integer

sJsonData = Replace(response, """", "")

Dim codigoSuministro As String
Dim Nombres As String
Dim CodigoPeriodoComercial As String
Dim MontoTotal As Currency
Dim FechaVencimiento As String
Dim Mensaje As String
Dim Pagara As String

Set p = JSON.parse(sJsonData)
If JSON.GetParserErrors <> "" Then
    MsgBox JSON.GetParserErrors, vbOKCancel, "Error"
    Exit Function
End If

Set R = New ADODB.Recordset
' Define campos
With R.Fields
    .Append "cNomCliente", adBSTR, 500
    .Append "nTipoDoi", adSmallInt
    .Append "cDoi", adBSTR, 8
    .Append "cConcepto", adBSTR, 20
    .Append "cCodCliente", adBSTR, 20
    .Append "nMora", adCurrency
    .Append "cCodigoComprobante", adBSTR, 50
    .Append "nMontoTotal", adCurrency
    .Append "cFechaVencimiento", adBSTR, 10
End With

' Abre el recordset desconectado
R.Open

For i = 1 To p.Count
    codigoSuministro = Trim(p.Item(i).Item("SUMINISTRO"))
    Nombres = Trim(p.Item(i).Item("NOMBRES"))
    CodigoPeriodoComercial = Trim(Trim(p.Item(i).Item("Periodo"))) 'MESULT")))
    MontoTotal = CCur(Trim(p.Item(i).Item("IMPORTEPARCIAL"))) 'TOTAL")))
    FechaVencimiento = Format$(Trim(p.Item(i).Item("FECHAVENCIMIENTO")), "dd/mm/yyyy")
    Mensaje = Trim(p.Item(i).Item("Mensaje"))
    Pagara = Trim(p.Item(i).Item("PAGARA"))
    
    If Mensaje = "OK" And Pagara = "SI" Then
        ' Agregar data al recordset
        R.AddNew _
            Array("cNomCliente", "nTipoDoi", "cDoi", "cConcepto", "cCodCliente", "nMora", _
                  "cCodigoComprobante", "nMontoTotal", "cFechaVencimiento"), _
            Array(Nombres, 1, "00000000", CodigoPeriodoComercial, codigoSuministro, 0, _
                  codigoSuministro, MontoTotal, FechaVencimiento)
    End If
Next i

Set ConsultarSuministroELORResponse = R

End Function

'CTI7 INI Se debe listar solo 1 de los pagos parciales
Public Function ConsultarSuministroELORFirstRecord(ByVal records As ADODB.Recordset) As ADODB.Recordset
Dim R, RTmp As ADODB.Recordset
Set R = New ADODB.Recordset
Set RTmp = New ADODB.Recordset
Set RTmp = records

With R.Fields
    .Append "cNomCliente", adBSTR, 500
    .Append "nTipoDoi", adSmallInt
    .Append "cDoi", adBSTR, 8
    .Append "cConcepto", adBSTR, 20
    .Append "cCodCliente", adBSTR, 20
    .Append "nMora", adCurrency
    .Append "cCodigoComprobante", adBSTR, 50
    .Append "nMontoTotal", adCurrency
    .Append "cFechaVencimiento", adBSTR, 10
End With
R.Open

If records.RecordCount > 0 Then
    Do While Not records.EOF
        If R.RecordCount = 0 Then
            R.AddNew _
                Array("cNomCliente", "nTipoDoi", "cDoi", "cConcepto", "cCodCliente", "nMora", _
                      "cCodigoComprobante", "nMontoTotal", "cFechaVencimiento"), _
                Array(records!cNomCliente, records!nTipoDoi, records!cDoi, records!cConcepto, records!cCodCliente, records!nMora, _
                      records!cCodigoComprobante, records!nMontoTotal, records!cFechaVencimiento)
        End If
        records.MoveNext
    Loop
    If records.EOF Then
        records.MoveFirst
    End If
End If
Set ConsultarSuministroELORFirstRecord = R
End Function
'CTI7 INI Se debe listar solo 1 de los pagos parciales

Public Function RealizarPagoELOR(ByVal contrato As String, ByVal monto As String, ByVal cMovNro As String, ByVal urlSimaynas As String) As String
    Dim postData As String
    On Error GoTo ErrorRealizarPagoELOR
    
    Set Req = New WinHttp.WinHttpRequest
    postData = "pcContrato=" & contrato & "&pcMonto=" & monto & "&pcMovNro=" & cMovNro
    With Req
        .Open "POST", urlSimaynas & "/ServicioRecaudoELOR/realizar_Pago?" & postData, Async:=False 'Descomentar al pasar a Produccion
        '.Open "POST", "http://localhost:65229/ServicioRecaudoELOR/realizar_Pago?" & postData, Async:=False 'Pruebas
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=""UTF-8"""
        .Send
    End With
    
    If Req.Status = 200 Then
        If Req.ResponseText = "NoResult" Then
            RealizarPagoELOR = ""
            Exit Function
        End If
        If Req.ResponseText = "SI" Then
            RealizarPagoELOR = "SI"
        Else
            RealizarPagoELOR = ""
        End If
    Else
        MsgBox "No se pudo realizar satisfactoriamente la operación", vbExclamation, "Mensaje"
        RealizarPagoELOR = ""
        Exit Function
    End If
    Exit Function
ErrorRealizarPagoELOR:
    RealizarPagoELOR = ""
    Err.Raise Err.Number, "RealizarPagoELOR ", Err.Description
End Function
'CTI1 ERS027-2019 ELOR
