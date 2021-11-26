VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'modificado por requerimiento de funcional
'fecha: 20060328
'autor: Pedro Mucha
'motivo: se agregó un campo de comentario en la tabla ColocParametro
'y se modifico la consulta de seleccion.
Public Function RecuperaDatos(Optional ByVal psProducto As String = "", Optional ByVal pnOrden As Integer = 1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
' pnOrden ==> 0 = (cParamDesc) // 1 = (nParamVar)
On Error GoTo ErrorRecuperaDatos

'    sSql = "Select nParamVar,cParamDesc,nParamValor,ISNULL(cParamCom,'') AS cParamCom From ColocParametro " & _
           "Where nParamVar like '" & psProducto & "%' " & _
           "Order by " & IIf(pnOrden = 1, "cParamDesc", "nParamVar")
           
    sSql = "stp_sel_ObtieneParamCredMant '" & psProducto & "' ," & pnOrden '*** PEAC 20160920
           
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaDatos:
    Err.Raise Err.Number, "Recupera Datos", Err.Description

End Function

Public Function RecuperaValorParametro(ByVal pnCodigo As Long) As Double
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorRecuperaValorParametro
    sSql = "Select nParamValor from ColocParametro Where nParamVar = " & pnCodigo
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If R.RecordCount > 0 Then
        RecuperaValorParametro = R!nParamValor
    Else
        RecuperaValorParametro = 0
    End If
    Exit Function

ErrorRecuperaValorParametro:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
    

End Function

Public Function ExisteParametro(ByVal psVariable As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    sSql = "Select * from ColocParametro where nParamVar = '" & psVariable & "'"
    Set oConecta = New COMConecta.DCOMConecta
    Call oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    Call oConecta.CierraConexion
    Set oConecta = Nothing
    
    If R.RecordCount > 0 Then
        ExisteParametro = True
    Else
        ExisteParametro = False
    End If
    
End Function

'se modifico este método
'fecha: 20060328
'autor: Pedro Mucha
'motivo: Dado que se agregó un nuevo campo en la tabla colocParametro
'es necesario agregar este campo en la consulta de inserción
Public Sub NuevoParametro(ByVal psVariable As String, ByVal psDesc As String, _
ByVal pnValor As Double, ByVal psComentario As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorNuevoParametro
        
    sSql = "INSERT INTO ColocParametro(nParamVar,cParamDesc,nParamValor,cParamCom)"
    sSql = sSql & " VALUES('" & psVariable & "','" & psDesc & "'," & _
    Format(pnValor, "#0.00") & ",'" & psComentario & "')"
    
    Set oConecta = New COMConecta.DCOMConecta
    Call oConecta.AbreConexion
    Call oConecta.Ejecutar(sSql)
    Call oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorNuevoParametro:
    Err.Raise Err.Number, "Nuevo Parametro", Err.Description
End Sub

'Se modifico este método
'fecha: 20060328
'autor: Pedro Mucha
'motivo: dado que se agregó un nuevo campo es necesario incluirlo en la
'consulta de modificacion
'Public Sub ModificarParametro(ByVal psVariable As String, _
'ByVal psDesc As String, ByVal pnValor As Double, ByVal psComentario As String) 'FRHU20131129
Public Sub ModificarParametro(ByVal psVariable As String, _
Optional ByVal psDesc As String = "", Optional ByVal pnValor As Double = 0, Optional ByVal psComentario As String = "")

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorModificarParametro
    'FRHU20131129
'    sSql = "UPDATE ColocParametro SET cParamDesc='" & Replace(psDesc, "'", "''") & _
'    "', nParamValor = " & Format(pnValor, "#0.00") & ", cParamCom = '" & psComentario & "'"
'    sSql = sSql & " Where nParamVar = " & psVariable
    
        sSql = "UPDATE ColocParametro SET "
    If psDesc <> "" Then
        sSql = sSql & " cParamDesc = '" & Replace(psDesc, "'", "''") & "',"
    End If
    If psComentario <> "" Then
        sSql = sSql & " cParamCom = '" & psComentario & "', "
    End If
        sSql = sSql & " nParamValor = " & Format(pnValor, "#0.00")
        sSql = sSql & " Where nParamVar = " & psVariable
    
    Set oConecta = New COMConecta.DCOMConecta
    Call oConecta.AbreConexion
    Call oConecta.Ejecutar(sSql)
    Call oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificarParametro:
    Err.Raise Err.Number, "Modifcar Parametro", Err.Description
End Sub

Public Sub EliminarParametro(ByVal psVariable As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorEliminarParametro
        
        sSql = "DELETE ColocParametro "
        sSql = sSql & " Where nParamVar = " & psVariable
        
    Set oConecta = New COMConecta.DCOMConecta
    Call oConecta.AbreConexion
    Call oConecta.Ejecutar(sSql)
    Call oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorEliminarParametro:
    Err.Raise Err.Number, "Eliminar Parametro", Err.Description
End Sub

'MAVM 20121113 ***
'Public Function RecuperaParametrosCalendario(ByRef pnTramoNoConsPorcen As Double, ByRef pnPlazoMiViv As Integer)
Public Function RecuperaParametrosCalendario(ByRef pnTramoNoConsPorcen As Double, ByRef pnPlazoMiViv As Integer, Optional ByRef pnPlazoMiVivMax As Integer)
'***
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorRecuperaValorParametro
    'MAVM 20121113 ***
    'sSQL = "SELECT nParamValor FROM ColocParametro WHERE nParamvar =3056"
    sSql = "SELECT nParamValor FROM ColocParametro WHERE nParamvar = " & gColocMiVivPlazoMin
    '***
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    
    pnPlazoMiViv = IIf(R.EOF, 0, R!nParamValor)
    
    'MAVM 20120606 ***
    sSql = "SELECT nParamValor FROM ColocParametro WHERE nParamvar = " & gColocMiVivPlazoMax
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    pnPlazoMiVivMax = IIf(R.EOF, 0, R!nParamValor)
    '***
    
    'MAVM 20121113 ***
    'sSQL = "SELECT nParamValor FROM ColocParametro WHERE nParamvar =" & gColocMiVivTramo
    sSql = "SELECT nParamValor FROM ColocParametro WHERE nParamvar = " & gColocMiVivBonoBuenPagador
    '***
    
    Set R = oConecta.CargaRecordSet(sSql)
    
    pnTramoNoConsPorcen = IIf(R.EOF, 0, R!nParamValor)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaValorParametro:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
    

End Function

'MAVM 20121113 ***
Public Function RecuperaParametrosCalendarioMiViv(ByRef pnValVentaMinMiViv As Double, ByRef pnValVentaMaxMiViv As Double, ByRef pnPorcentCIMiViv As Double, ByRef pnBonoBuenPagador As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorRecuperaParametrosCalendarioMiViv
    
    sSql = "SELECT nParamValor FROM ColocParametro WHERE nParamvar = " & gColocMiVivValVentMin
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    pnValVentaMinMiViv = IIf(R.EOF, 0, R!nParamValor)
    
    sSql = "SELECT nParamValor FROM ColocParametro WHERE nParamvar = " & gColocMiVivValVentMax
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    pnValVentaMaxMiViv = IIf(R.EOF, 0, R!nParamValor)
    
    sSql = "SELECT nParamValor FROM ColocParametro WHERE nParamvar = " & gColocMiVivPorcentCI
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    pnPorcentCIMiViv = IIf(R.EOF, 0, R!nParamValor) 'gColocMiVivBonoBuenPagador
    
    sSql = "SELECT nParamValor FROM ColocParametro WHERE nParamvar = " & gColocMiVivBonoBuenPagador
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    pnBonoBuenPagador = IIf(R.EOF, 0, R!nParamValor) '
       
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaParametrosCalendarioMiViv:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'***

'MAVM 20130209 ***
Public Function RecuperaDatosPeriodo(ByVal psValor As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
On Error GoTo ErrorRecuperaDatosPeriodo
    sSql = "Select ROW_NUMBER()OVER(ORDER BY cConsDescripcion) Correlativo, C.nConsValor, cConsDescripcion, ISNULL(nMeses, 0) nMeses From Constante C Left Join PeriodoMaximoDestino P on C.nConsValor = P.nConsValor And P.iEstado = 1 " & _
           "Where nConsCod = " & gColocDestino & " And C.nConsValor <> " & gColocDestino & " And bEstado = 1 And C.nConsValor Like '" & psValor & "'" & _
           "Order by cConsDescripcion"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosPeriodo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaDatosPeriodo:
    Err.Raise Err.Number, "Recupera Datos Periodo", Err.Description
End Function

Public Sub NuevoPeriodo(ByVal psValor As Integer, ByVal pnMeses As Integer, ByVal pnMovNro As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorNuevoPeriodo
        
    sSql = "Insert Into PeriodoMaximoDestino(nConsValor, nMeses, cMovNro, iEstado)"
    sSql = sSql & " VALUES(" & psValor & "," & pnMeses & ", '" & pnMovNro & "', " & 1 & ")"
    
    Set oConecta = New COMConecta.DCOMConecta
    Call oConecta.AbreConexion
    Call oConecta.Ejecutar(sSql)
    Call oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorNuevoPeriodo:
    Err.Raise Err.Number, "Nuevo Parametro", Err.Description
End Sub

Public Sub ModificarDatosPeriodo(ByVal psValor As Integer)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorModificarDatosPeriodo
    
    sSql = "Update PeriodoMaximoDestino Set iEstado = 0 Where nConsValor = " & psValor
    
    Set oConecta = New COMConecta.DCOMConecta
    Call oConecta.AbreConexion
    Call oConecta.Ejecutar(sSql)
    Call oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorModificarDatosPeriodo:
    Err.Raise Err.Number, "Modificar Datos Periodo", Err.Description
End Sub

Public Sub InsertarAutorizExcedePlazo(ByVal psCtaCod As String, ByVal psMovNro As String, ByVal pnPeriodo As Integer, ByVal pnTEA As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorInsertarAutorExcedePlazo
        
    sSql = "Insert Into ColocacExcedePlazo(cCtaCod, cMovNro, nPeriodo, nTEA, nEstado)"
    sSql = sSql & " VALUES('" & psCtaCod & "','" & psMovNro & "', " & pnPeriodo & ", " & pnTEA & ", " & 0 & ")"
    
    Set oConecta = New COMConecta.DCOMConecta
    Call oConecta.AbreConexion
    Call oConecta.Ejecutar(sSql)
    Call oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorInsertarAutorExcedePlazo:
    Err.Raise Err.Number, "Inserta Autorizacion Excede Plazo", Err.Description
End Sub

Public Function ListarCreditoExcedePlazo() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorListarCreditoExcedePlazo

    sSql = "exec stp_sel_ListarCreditoExcedePlazo"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ListarCreditoExcedePlazo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function
ErrorListarCreditoExcedePlazo:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'***
'WIOR 20130719 ************************************************************************************************
Public Function InsertaParametrosAgro(ByVal pnTipo As Long, ByVal psSubTipo As String, Optional ByVal pnEstado As Integer = 1, Optional pnMin As Long = 0) As Long
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

On Error GoTo ErrorInsertaParametrosAgro

    sSql = "exec stp_ins_CredParamAgro " & pnTipo & ",'" & psSubTipo & "'," & pnEstado & "," & pnMin
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    
    If Not (rs.EOF And rs.BOF) Then
        InsertaParametrosAgro = CDbl(rs!nSubTipo)
    Else
        InsertaParametrosAgro = 0
    End If
    
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function
ErrorInsertaParametrosAgro:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Sub ActualizaParametrosAgro(ByVal pnTipo As Long, ByVal pnSubTipo As Long, ByVal psSubTipo As String, Optional ByVal pnEstado As Integer = 1, Optional pnMin As Long = 0)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorActualizaParametrosAgro
       
    sSql = "exec stp_upd_CredParamAgro " & pnTipo & "," & pnSubTipo & ",'" & psSubTipo & "'," & pnEstado & "," & pnMin
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Sub
ErrorActualizaParametrosAgro:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Sub
Public Function ObtenerParametrosAgro(Optional ByVal pnTipo As Long = 0, Optional ByVal psSubTipo As String = "%", Optional ByVal psEstado As String = "1", Optional ByVal pnSubProd As Integer = 0) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorObtenerParametrosAgro

    If pnSubProd = 0 Then
        sSql = "exec stp_sel_CredParamAgro " & pnTipo & ",'" & psSubTipo & "','" & psEstado & "'"
    Else
    
        sSql = "exec stp_sel_CredParamAgroXSubProd " & pnSubProd
    End If
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ObtenerParametrosAgro = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function
ErrorObtenerParametrosAgro:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ExisteCredAgrico(ByVal pnSubTipo As Long) As Double
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

On Error GoTo ErrorExisteCredAgrico

    sSql = "exec stp_sel_ExisteCredAgrico " & pnSubTipo
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        ExisteCredAgrico = CDbl(rs!CantCred)
    Else
        ExisteCredAgrico = 0
    End If
    
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function
ErrorExisteCredAgrico:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'WIOR FIN ****************************************************************************************************
'-------------------------------------- FRHU 20121202
Public Function RecuperaValidaClientePreferencial(ByVal pcPersCod As String, ByVal pnMinCalSbsNorm As Integer, _
ByVal pnEndeMaxUniFami As Integer, ByVal pnEdadMaxClie As Integer, Optional ByVal npPerContMini As Integer = 0) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
On Error GoTo ErrorRecuperaDatos
        
    sSql = "exec stp_sel_ConfiguracionClientePreferencial " & "'" & pcPersCod & "'" & " , " & npPerContMini & "," & pnMinCalSbsNorm & ", " & pnEndeMaxUniFami & ", " & pnEdadMaxClie

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaValidaClientePreferencial = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaDatos:
    Err.Raise Err.Number, "Recupera Datos", Err.Description

End Function
'-------------------------------------- FIN FRHU 20121202
'**************** FRHU 20140512 ERS068-2014
Public Function RecuperaParametroSeguroTarjetas() As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
On Error GoTo ErrorRecuperaDatos
        
    sSql = "exec stp_sel_SegTarjetaParametros"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaParametroSeguroTarjetas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaDatos:
    Err.Raise Err.Number, "Recupera Datos", Err.Description

End Function
Public Sub ActualizarSegTarjetaParametros(ByVal pnParamValor As Double, ByVal pnParamVar As Integer)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorActualizarSegTarjetaParametros
       
    sSql = "exec stp_upd_SegTarjetaParametros " & pnParamValor & "," & pnParamVar
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Sub
ErrorActualizarSegTarjetaParametros:
    Err.Raise Err.Number, "En Proceso", Err.Description
End Sub
'**** FIN FRHU 20140512
'APRI20171121 ERS028-2017
Public Function RecuperaValorParametroDesg(ByVal pnCodigo As Integer) As Double
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorRecuperaValorParametroDesg
    sSql = "EXEC STP_SEL_DESGPARAMETRO " & pnCodigo
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        RecuperaValorParametroDesg = R!nValor
    Else
        RecuperaValorParametroDesg = 0
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaValorParametroDesg:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
    

End Function
Public Function RecuperaValorCriterioDesg(ByVal pnCodigo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorRecuperaValorCriterioDesg
    sSql = "EXEC STP_SEL_CRITERIOASEGURABILIDADID " & pnCodigo
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaValorCriterioDesg = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaValorCriterioDesg:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
    

End Function
Public Function CumpleCriterioEdad(ByVal pcPersCod As String, ByVal dFecha As Date, ByVal pnCriterio As Integer) As Boolean
    Dim rs As New ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    On Error GoTo ErrCumpleCriterioEdad
    lsSQL = "SELECT bCumpleCriterioEdad = dbo.fn_cumplecriterioedad( '" & pcPersCod & "','" & Format(dFecha, "YYYY-MM-DD") & "'," & pnCriterio & ")"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        CumpleCriterioEdad = rs!bCumpleCriterioEdad
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrCumpleCriterioEdad:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'END APRI

