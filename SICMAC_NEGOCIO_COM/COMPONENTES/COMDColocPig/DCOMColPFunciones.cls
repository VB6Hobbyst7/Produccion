VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMColPFunciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim coConex As New COMConecta.DCOMConecta
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String
Dim oError As New COMConecta.COMErrorHandling


Public Function dObtieneCuentasPersona(ByVal psCodPers As String, ByVal psEstados As String, ByVal psMoneda As String) As ADODB.Recordset
Dim sql As String
Dim rs As New ADODB.Recordset
On Error GoTo dError
    'Sql = " Select Pro.cCtaCod, P.cPersNombre from ProductoPersona PP"
    'Sql = Sql & " Inner Join Persona P on P.cPersCod  = PP.cPersCod"
    'Sql = Sql & " Inner Join Producto Pro on Pro.cCtaCod = PP.cCtaCod"
    'Sql = Sql & " Where PP.nPrdPersRelac = 10 And Pro.nPrdEstado in (" & psEstados & ")"
    'Sql = Sql & " and P.cPersCod = '" & psCodPers & "' and substring(Pro.cCtaCod,9,1)='" & psMoneda & "'"
    sql = " select Pro.cCtaCod, nSaldo from ProductoPersona PP"
    sql = sql & " Inner Join Producto Pro on Pro.cCtaCod = PP.cCtaCod"
    sql = sql & " where cperscod='" & psCodPers & "'"
    sql = sql & " and nPrdPersRelac=10 "
    sql = sql & " and nPrdEstado in (" & psEstados & ") "
    sql = sql & " and substring(Pro.cCtaCod,9,1)='" & psMoneda & "' Order by Pro.cCtaCod"
    
    Set rs = coConex.CargaRecordSet(sql)
    Set dObtieneCuentasPersona = rs
    Set rs = Nothing
    Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneContratosPersona>>", Err.Description

End Function


Public Function dObtieneContratosPersona(ByVal psCodPers As String, ByVal psEstados As String, _
                                        Optional ByVal pbSoloCreditos As Boolean = False) As Recordset

Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New ADODB.Recordset

On Error GoTo dError


lsSQL = "SELECT  P.cCtaCod, P.nPRdEstado FROM ProductoPersona PP " _
    & "INNER JOIN Producto P ON PP.cCtaCod = P.cCtaCod " _
    & IIf(pbSoloCreditos, "INNER JOIN ColocacCred C ON C.cCtaCod=P.cCtaCod ", "") _
    & "WHERE PP.cPersCod = '" & psCodPers & "' " _
    & " AND P.nPrdEstado in (" & psEstados & ")"

Set dObtieneContratosPersona = coConex.CargaRecordSet(lsSQL)

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneContratosPersona>>", Err.Description
End Function

'Devuelve Descripcion de Zona de un cliente
Public Function dObtieneNombreZona(ByVal psCodigoZona As String) As String
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Dim lsNombreZona As String

On Error GoTo dError

Set lrs = New ADODB.Recordset

    lsSQL = "SELECT cDesZon FROM " & csCentralCom & "Zonas WHERE cCodZon = '" & psCodigoZona & "'"
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    If lrs.BOF And lrs.EOF Then
        lsNombreZona = ""
    Else
        lsNombreZona = Trim(lrs!cDesZon)
    End If
    lrs.Close
Set lrs = Nothing
dObtieneNombreZona = lsNombreZona

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Nombre Zona en <<dObtieneNombreZona>>", Err.Description

End Function

'Devuelve Ciudad de Zona de un cliente
Public Function dObtieneCiudadZona(ByVal psCodigoZona As String) As String
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Dim lsNombreCiudad As String

On Error GoTo dError

Set lrs = New ADODB.Recordset


    lsSQL = "SELECT cDesZon FROM " & csCentralCom & "Zonas WHERE cCodZon IN ('" & "1" & Mid(psCodigoZona, 2, 2) & "000000000" & "','" & "2" & Mid(psCodigoZona, 2, 4) & "0000000" & "')"
    Set lrs = coConex.CargaRecordSet(lsSQL)
    If lrs.BOF And lrs.EOF Then
        lsNombreCiudad = ""
    Else
        Do While Not lrs.EOF
            lsNombreCiudad = lsNombreCiudad & " " & Trim(lrs!cDesZon)
            lrs.MoveNext
        Loop
    End If
    lrs.Close
Set lrs = Nothing
dObtieneCiudadZona = lsNombreCiudad

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Ciudad Zona  <<dObtieneCiudadZona>>", Err.Description

End Function


'Devuelve Descripcion de Zona de un cliente
Public Function dObtieneNombreZonaPersona(ByVal psCodPers As String) As String
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Dim lsNombreZona As String

On Error GoTo dError

Set lrs = New ADODB.Recordset

    lsSQL = "SELECT cUbiGeoDescripcion FROM " & csCentralCom & "UbicacionGeografica WHERE cUbiGeoCod " & _
            "in (Select cPersDireccUbiGeo From Persona where cPersCod = '" & psCodPers & "' ) "
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    If lrs.BOF And lrs.EOF Then
        lsNombreZona = ""
    Else
        lsNombreZona = Trim(lrs!cUbiGeoDescripcion)
    End If
    lrs.Close
Set lrs = Nothing
dObtieneNombreZonaPersona = lsNombreZona

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Nombre Zona en <<dObtieneNombreZona>>", Err.Description

End Function


'Devuelve Ciudad de Zona de un cliente
Public Function dObtieneCiudadZonaPersona(ByVal psCodPers As String) As String
'Dim lrs As ADODB.Recordset
'Dim lsSQL As String
'Dim lsNombreCiudad As String
'
'On Error GoTo dError
'
'Set lrs = New ADODB.Recordset
'
'    lsSQL = "SELECT cDesZon FROM " & csCentralCom & "Zonas " & _
'            "WHERE cCodZon IN ('" & "1" & Mid(psCodigoZona, 2, 2) & "000000000" & "','" & "2" & Mid(psCodigoZona, 2, 4) & "0000000" & "')"
'    Set lrs = coConex.CargaRecordSet(lsSQL)
'    If lrs.BOF And lrs.EOF Then
'        lsNombreCiudad = ""
'    Else
'        Do While Not lrs.EOF
'            lsNombreCiudad = lsNombreCiudad & " " & Trim(lrs!cDesZon)
'            lrs.MoveNext
'        Loop
'    End If
'    lrs.Close
'Set lrs = Nothing
'dObtieneCiudadZonaPersona = lsNombreCiudad
'
'Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Ciudad Zona  <<dObtieneCiudadZona>>", Err.Description

End Function

Public Function dObtieneMovNroAnterior(ByVal psCtaCod As String, ByVal psOpeCod As String) As Long
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Dim lnMovNro As String

On Error GoTo dError

Set lrs = New ADODB.Recordset


    lsSQL = "SELECT nMovNro FROM MovCol WHERE cCtaCod = '" & psCtaCod & "' AND cOpecod ='" & psOpeCod & "' "
    Set lrs = coConex.CargaRecordSet(lsSQL)
    If lrs.BOF And lrs.EOF Then
        lnMovNro = 0
    Else
        Do While Not lrs.EOF
            lnMovNro = lrs!nMovNro
            lrs.MoveNext
        Loop
    End If
    lrs.Close
Set lrs = Nothing
dObtieneMovNroAnterior = lnMovNro

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Ciudad Zona  <<dObtieneCiudadZona>>", Err.Description

End Function

Public Function dObtieneRecordSet(ByVal psSql As String) As ADODB.Recordset
Dim lrs As ADODB.Recordset

On Error GoTo dError

'Set lrs = New ADODB.Recordset
Set dObtieneRecordSet = coConex.CargaRecordSet(psSql)

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Recordset <<dObtieneRecordSet>>", Err.Description

End Function


Public Function dObtieneAgencias(Optional pbOrdenCodigo As Boolean = True) As Recordset
    Dim lsSQL As String
    Dim lsOrdena As String
    
    On Error GoTo ErrorAge
    
    lsOrdena = " ORDER BY " & IIf(pbOrdenCodigo = True, " cAgeCod ", " cAgeDescripcion ")
    lsSQL = "SELECT cAgeCod, cAgeDescripcion FROM " & csCentralCom & "Agencias" & lsOrdena
    
    Set dObtieneAgencias = coConex.CargaRecordSet(lsSQL)
    
Exit Function
ErrorAge:
    Call oError.RaiseError(oError.MyUnhandledError, "DCOMCONSTANTEs:GetConstante Method")
End Function

'***************************************************
'* Verifica si una fecha se encuentra en la tabla feriados
'***************************************************
Function dVerSiFeriado(ByVal psFecVer As String, ByRef psmensaje As String) As Boolean
Dim lsSQL As String
Dim lrs As ADODB.Recordset
Dim oFun As New COMFunciones.FCOMFechas

lsSQL = oFun.ValidaFecha(psFecVer)
If lsSQL = "" Then
    ' Determina si una fecha es feriado
    psFecVer = Format(psFecVer, "mm/dd/yyyy")
    Set lrs = New ADODB.Recordset
    
    lsSQL = "Select dFeriado From Feriado Where dFeriado = '" & psFecVer & "' "
    
    Set lrs = coConex.CargaRecordSet(lsSQL)
    If lrs.BOF And lrs.EOF Then
        dVerSiFeriado = False
    Else
        dVerSiFeriado = True
    End If
    lrs.Close
    Set lrs = Nothing
Else
    psmensaje = lsSQL
    dVerSiFeriado = False
End If

End Function

Private Sub Class_Initialize()
    Dim loIni As New COMConecta.DCOMClasIni

    csConexion = loIni.CadenaConexion
    csNegocio = loIni.BaseNegocio
    csCentralPer = loIni.BasePersonas
    csCentralCom = loIni.BaseComunes
    csCentralImg = loIni.BaseImagenes
    csAdminist = loIni.BaseAdministracion
    If coConex.AbreConexion(csConexion) = False Then
        Call oError.RaiseError(oError.MyUnhandledError, "DColPContrato:Initialize. Error en Conexion a Base de datos")
    End If
End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

'*** PEAC 20190508
Public Function dObtieneFechaVencimientoFeriado(ByVal psFecVenc As String, ByVal psAge As String, ByVal psFecSis As String) As Recordset
    Dim lsSQL As String
    Dim lsOrdena As String
    
    On Error GoTo ErrorAge
    
    lsSQL = "exec stp_sel_FechaVencimientoFeriado '" & psFecVenc & "','" & psAge & "','" & psFecSis & "' "
    
    Set dObtieneFechaVencimientoFeriado = coConex.CargaRecordSet(lsSQL)
    
Exit Function
ErrorAge:
    Call oError.RaiseError(oError.MyUnhandledError, "DCOMColPFunciones:dObtieneFechaVencimientoFeriado Method")
End Function
'APRI20190515
Public Function dObtieneInfoClientePig(ByVal psPersCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim lsOrdena As String
    
    On Error GoTo ErrorAge
    
    lsSQL = "exec stp_sel_ClienteInfoPigno '" & psPersCod & "'"
    
    Set dObtieneInfoClientePig = coConex.CargaRecordSet(lsSQL)
    
Exit Function
ErrorAge:
    Call oError.RaiseError(oError.MyUnhandledError, "DCOMColPFunciones:dObtieneInfoClientePig Method")
End Function
'END APRI
