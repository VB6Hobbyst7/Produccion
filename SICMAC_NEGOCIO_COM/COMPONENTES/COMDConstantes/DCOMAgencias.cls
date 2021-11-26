VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMAgencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function NombreAgencia(ByVal psAgeCod As String) As String
Dim sSQL As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
    sSQL = "select cAgeDescripcion from Agencias Where cAgeCod = '" & psAgeCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If R.RecordCount > 0 Then
        NombreAgencia = Trim(UCase(R!cAgeDescripcion))
    Else
        NombreAgencia = ""
    End If
    R.Close
End Function

'peac 20071224
Public Function ObtieneAgencias() As ADODB.Recordset
Dim sSQL As String
Dim oConecta As COMConecta.DCOMConecta
    
    'Modify By GITU 14-06-2013
    'sSQL = "select cAgeDescripcion cConsDescripcion, cAgecod nConsValor from agencias where nEstado=1"
    
    sSQL = "stp_sel_ObtieneAgencias"
    'End GITU
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneAgencias = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'madm 20100723
Public Function ObtieneAgenciasIqt() As ADODB.Recordset
Dim sSQL As String
Dim oConecta As COMConecta.DCOMConecta

    sSQL = "exec stp_sel_AgenciaIqt"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneAgenciasIqt = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function


Public Function RecuperaAgencias(Optional lsCodAge As String) As ADODB.Recordset
Dim sSQL As String
Dim sSqlFiltro As String
Dim oConecta As COMConecta.DCOMConecta

    '*** PEAC 20080308
    'If lsCodAge <> "" Then sSqlFiltro = " where cAgeCod = '" & lsCodAge & "'"
    
    'ARCV 30-01-2007
    'sSQL = "Select * From Agencias " & sSqlFiltro
    '***PEAC 20080220 -- se modifico para que  no figure la frase (CIUDAD) en cUbiGeoDescripcion
    'sSQL = "SELECT A.*,REPLACE(UG.cUbiGeoDescripcion,'(CIUDAD)','')cUbiGeoDescripcion," & _
           " Dep=(SELECT cUbiGeoDescripcion FROM UbicacionGeografica WHERE LEFT(cUbiGeoCod,1)='1' AND SUBSTRING(cUbiGeoCod,2,2)= SUBSTRING(UG.cUbiGeoCod,2,2) ), " & _
           " Prov=(SELECT cUbiGeoDescripcion FROM UbicacionGeografica WHERE LEFT(cUbiGeoCod,1)='2' AND SUBSTRING(cUbiGeoCod,2,4)= SUBSTRING(UG.cUbiGeoCod,2,4) ), " & _
           " Dist=(SELECT cUbiGeoDescripcion FROM UbicacionGeografica WHERE LEFT(cUbiGeoCod,1)='3' AND SUBSTRING(cUbiGeoCod,2,6)= SUBSTRING(UG.cUbiGeoCod,2,6) ) " & _
           " FROM Agencias A INNER JOIN UbicacionGeografica UG ON A.cUbiGeoCod = UG.cUbiGeoCod " & sSqlFiltro
    '-------
    
    '***PEAC 20080308
    If lsCodAge <> "" Then
        sSQL = "exec stp_sel_RecuperaAgencias '" & lsCodAge & "'"
    Else
        sSQL = "exec stp_sel_RecuperaAgencias "
    End If
    
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaAgencias = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function VerficaZonaAgencia(ByVal lsCodAge As String, ByVal lsCodAgeCta As String) As Boolean
Dim sSQL As String
Dim lsUbigeoAge As String
Dim lsUbigeoAgeCta As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    
    'sSQL = "select substring(cUbiGeoCod,1,6) Age from Agencias  where cAgeCod='" & lsCodAge & "'"
    sSQL = "select substring(cUbiGeoCod,1,5) Age from Agencias  where cAgeCod='" & lsCodAge & "'" 'JUEZ 20140521 por incidente INC1405210013
    Set rs = New ADODB.Recordset
    Set rs = oConecta.CargaRecordSet(sSQL)
      If Not (rs.EOF And rs.BOF) Then
         lsUbigeoAge = rs!Age
      End If
    rs.Close
    
    
    'sSQL = "select substring(cUbiGeoCod,1,6) AgeCta from Agencias  where cAgeCod='" & lsCodAgeCta & "'"
    sSQL = "select substring(cUbiGeoCod,1,5) AgeCta from Agencias  where cAgeCod='" & lsCodAgeCta & "'" 'JUEZ 20140521 por incidente INC1405210013
    Set rs = New ADODB.Recordset
    Set rs = oConecta.CargaRecordSet(sSQL)
      If Not (rs.EOF And rs.BOF) Then
        lsUbigeoAgeCta = rs!AgeCta
      End If
    rs.Close
    
    If Trim(lsUbigeoAge) <> Trim(lsUbigeoAgeCta) Then
        VerficaZonaAgencia = True
    Else
        VerficaZonaAgencia = False
    End If
    
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'**DAOR 20070201
'**Recuperar agencias del banco de la nación. para desembolsos por corresponsalia
Public Function RecuperaAgenciasBancoNacion() As ADODB.Recordset
Dim sSQL As String
Dim oConecta As COMConecta.DCOMConecta

    sSQL = "select cCodAge,vAgencia FROM AgenciasBancoNacion order by vAgencia"
    '-------
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaAgenciasBancoNacion = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
'RECO20151222 ERS040-2015******************************************
Public Function ObtieneDatosAgencia(ByVal psAgeCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As COMConecta.DCOMConecta

    sSQL = "stp_sel_ObtieneDatosAgencia '" & psAgeCod & "'"
    '-------
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneDatosAgencia = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'RECO FIN**********************************************************
'JUEZ 20160407 *********************************************************
Public Function ObtieneCodigoAgencia(ByVal psAgeDesc As String) As String
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    sSQL = "stp_sel_ObtieneCodigoAgencia '" & Trim(psAgeDesc) & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If R.RecordCount > 0 Then
        ObtieneCodigoAgencia = R!cAgeCod
    Else
        ObtieneCodigoAgencia = ""
    End If
End Function
'END JUEZ **************************************************************
