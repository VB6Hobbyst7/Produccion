VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMLineaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************
'***     Rutina:           DLineaCredito
'***     Descripcion:      Clase que permite el Mantenimeinto de las Lineas de Credito
'***     Creado por:        NSSE
'***     Maquina:           07SIST_08
'***     Fecha-Tiempo:         14/06/2001 04:22:46 PM
'***     Ultima Modificacion: Lo Ultimo que se Modifico
'*****************************************************************************************

Option Explicit
Private gConsPersona As String
Private gConsComunes As String
Private gConsImagenes As String
Private oConn As COMConecta.DCOMConecta

'*****************************************************************************************
'******* Recupera las Lineas de Credito de un Producto
'*****************************************************************************************
Public Function RecuperaLineasProducto(ByVal psProducto As String, ByVal psMoneda As String) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
    
    On Error GoTo ErrorRecuperaLineasProducto
    sSql = "Select CL.cLineaCred, CL.cDescripcion, CL.nPlazoMax, CL.nPlazoMin, CL.nMontoMax, CL.nMontoMin, "
    sSql = sSql & " CLT.nTasaIni, CLT.nTasaFin, CLT2.nTasaIni as nTasaGraciaIni, CLT2.nTasaFin as nTasaGraciaFin,CLT3.nTasaIni as nTasaMoraIni, CLT3.nTasaFin as nTasaMoraFin "
    sSql = sSql & " From ColocLineaCredito CL Inner Join ColocLineaCreditoTasa CLT ON CL.cLineaCred = CLT.cLineaCred AND CLT.nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntCompNormal
    sSql = sSql & "                           Left Join ColocLineaCreditoTasa CLT2 ON CL.cLineaCred = CLT2.cLineaCred AND CLT2.nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntGracia
    sSql = sSql & "                           Left Join ColocLineaCreditoTasa CLT3 ON CL.cLineaCred = CLT3.cLineaCred AND CLT3.nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntMoratNormal
    sSql = sSql & " Where substring(CL.cLineaCred,7,3) = '" & psProducto & "' AND substring(CL.cLineaCred,5,1) = '" & psMoneda & "' Order By cDescripcion"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaLineasProducto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaLineasProducto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaLineasProductoArbol(ByVal psProducto As String, ByVal psMoneda As String, _
                                Optional ByVal pbPreferencial As Boolean = False, _
                                Optional ByVal psAgeCod As String = "", _
                                Optional ByVal pnPlazo As Integer = 0, _
                                Optional ByVal pnMonto As Double = 0, _
                                Optional ByVal pnCuotas As Integer = 0, _
                                Optional ByVal pnTasa As Currency = -1, _
                                Optional ByVal pnDiasGracia As Integer = 0, _
                                Optional ByVal pdFecSis As Date = "1900/01/01", _
                                Optional ByVal cPersCodMivienda As String = "", _
                                Optional ByVal pnIdCampana As Integer = 0) As ADODB.Recordset 'arlo20200617

'Se Agrego la Agencia para el Cusco
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim sparte1 As String
Dim sParte2 As String
Dim sParte3 As String

    If (pnPlazo * pnCuotas) <= 365 Then
        sparte1 = "CP1-0000" + psMoneda + "1" + psProducto + "00"
    'sParte2 = "LP2-0000" + psMoneda + "2" + psProducto + "00"
    Else
        sparte1 = "LP2-0000" + psMoneda + "2" + psProducto + "00"
    End If

    On Error GoTo ErrorRecuperaLineasProductoArbol


'arlo 20200618 begin
'    sSql = "Select top 1 CASE substring(CL.cLineaCred,6,1) WHEN '1' THEN 'CP1-' + CL.cLineaCred WHEN '2' THEN 'LP2-' + CL.cLineaCred END AS cLineaCred, T.cLineaCreditoDes +'-'+ CL.cDescripcion + ' /'+ CONVERT(VARCHAR(20),NMONTOMIN) + ' - '+ CONVERT(VARCHAR(20),NMONTOMAX) + ' / ' + CONVERT(VARCHAR(20),nRanTasDesde) + ' - '+ CONVERT(VARCHAR(20),nRanTasHasta) + space(100) + CL.cPersCod  AS cDescripcion, '2' as Nivel  "
'    sSql = sSql & " From ColocLineaCredito CL JOIN  LineaCredito T ON T.CLINEACREDITOCOD=CL.CLINEACRED and rtrim(isnull(CL.cCtaIFCod,''))<>'' "
'    sSql = sSql & " JOIN LineaCreditoPrioridad LCP ON T.CLINEACREDITOCOD like left(LCP.CLINEACREDITOCOD,6)+'___'+right(LCP.CLINEACREDITOCOD,4) "
'    If psAgeCod <> "" Then
'        sSql = sSql & " INNER JOIN ColocLineaCreditoAgencia CLCA ON CLCA.cLineaCred=CL.cLineaCred "
'    End If
'
'    sSql = sSql & " Where substring(CL.cLineaCred,7,3) = '" & psProducto & "' AND substring(CL.cLineaCred,5,1) = '" & psMoneda & "' "
'    If psAgeCod <> "" Then
'        sSql = sSql & "  and LCP.cAgeCod=CLCA.cAgeCod "
'    End If
'    sSql = sSql & " AND CL.bEstado=1 "
'
'    If Not (psProducto = "853" Or psProducto = "854") Then
'        If pnMonto > 0 Then
'            sSql = sSql & " AND " & pnMonto & " BETWEEN ISNULL(NMONTOMIN,0) AND ISNULL(NMONTOMAX,0)"
'        End If
'        If pnPlazo > 0 Then
'            sSql = sSql & " AND " & ((pnPlazo * pnCuotas) + pnDiasGracia) & " <= DateDiff(d,'" & Format(pdFecSis, "YYYY/MM/DD") & "',T.dFechaMax) "
'        End If
'        If pnTasa <> -1 And psProducto <> "755" Then
'            sSql = sSql & "AND " & pnTasa & " between nRanTasDesde and nRanTasHasta "
'        End If
'        If pbPreferencial = True Then
'            sSql = sSql & " AND CL.bPreferencial=1"
'        End If
'        If psAgeCod <> "" Then
'            sSql = sSql & " AND CLCA.cAgeCod='" & psAgeCod & "'"
'        End If
'    Else
'        sSql = sSql & " AND T.cPersCodMivienda='" & cPersCodMivienda & "'"
'        If psAgeCod <> "" Then
'            sSql = sSql & " AND CLCA.cAgeCod='" & psAgeCod & "'"
'        End If
'    End If
'
'    sSql = sSql & " ORDER BY LCP.nPrioridad"
'arlo 20200618 begin

    'arlo20200618
    sSql = "exec  stp_sel_DevolverLineaCredito '" & psAgeCod & "','" & psProducto & "'," & pnMonto & "," & pnPlazo & "," & pnTasa & ",'" & psMoneda & "'," & pnCuotas & "," & pnDiasGracia & ",'" & cPersCodMivienda & "'," & pnIdCampana & ""

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaLineasProductoArbol = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaLineasProductoArbol:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'**********************************************************
'***** Saldo de Linea de Credito por Colocar
'**********************************************************
Public Function SaldoColocacionLineaCredito(ByVal psLineaCred As String) As Double
Dim sSql As String
Dim R As ADODB.Recordset
Dim nMontoTotal As Double
Dim nSaldoLineas As Double
    On Error GoTo ErrorSaldoColocacionLineaCredito
    Set oConn = New COMConecta.DCOMConecta
    sSql = "Select nMontoTotal,nMontoColocado from ColocLineaCreditoSaldo Where cLineaCred = '" & psLineaCred & "'"
    oConn.AbreConexion
    Set R = oConn.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        nMontoTotal = CDbl(Format(R!nMontoTotal, "#0.00"))
        'R.Close
        'Set R = Nothing
        'sSql = "Select sum(nSaldo) as nTotalSaldos from ColocLineaCredito Where cLineaCred like '" & psLineaCred & "______'"
        'Set R = oConn.CargaRecordSet(sSql)
        'nSaldoLineas = IIf(IsNull(R!nTotalSaldos), 0, R!nTotalSaldos)
        nSaldoLineas = CDbl(Format(R!nMontoColocado, "#0.00"))
        SaldoColocacionLineaCredito = CDbl(Format(nMontoTotal - nSaldoLineas, "#0.00"))
    Else
        SaldoColocacionLineaCredito = -1
    End If
    R.Close
    Set R = Nothing
    oConn.CierraConexion
    Set oConn = Nothing
    Exit Function
    
ErrorSaldoColocacionLineaCredito:
    Err.Raise Err.Number, "Saldo de Colocacion de Linea de Credito", Err.Description
End Function
Public Sub CreaSaldoLineaCredito(ByVal psLineaCred As String, Optional pbSinTransac As Boolean = True)
Dim sSql As String
    
    On Error GoTo ErrorCreaSaldoLineaCredito
    If pbSinTransac Then
        Set oConn = New COMConecta.DCOMConecta
        oConn.AbreConexion
    End If
    sSql = "INSERT INTO ColocLineaCreditoSaldo(cLineaCred,nMontoTotal,nSaldoCap,nMontoColocado)"
    sSql = sSql & "VALUES('" & psLineaCred & "',0.00,0.00,0.00)"
    oConn.ConexionActiva.Execute sSql
    If pbSinTransac Then
        oConn.CierraConexion
        Set oConn = Nothing
    End If
    Exit Sub
    
ErrorCreaSaldoLineaCredito:
    Err.Raise Err.Number, "LineaCreditoSaldo", Err.Description
End Sub

Public Sub ActualizarLinea(ByVal psLineaCredCod As String, ByVal psDescription As String, _
                        ByVal pbEstado As Integer, ByVal pnPlazoMax As Integer, ByVal pnPlazoMin As Integer, _
                        ByVal pnMontoMax As Double, ByVal pnMontoMin As Double, ByVal psPersCod As String, _
                        Optional ByVal pbPreferencial As Boolean = False, Optional ByVal pMatAgencias As Variant)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer

    On Error GoTo ErrorActualizarLinea
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "UPDATE ColocLineaCredito SET cDescripcion = '" & psDescription & "',"
    sSql = sSql & " bEstado = " & Trim(Str(pbEstado)) & ","
    sSql = sSql & " nPlazoMax = " & Format(pnPlazoMax, "#0") & ","
    sSql = sSql & " nPlazoMin = " & Format(pnPlazoMin, "#0") & ","
    sSql = sSql & " nMontoMax = " & Format(pnMontoMax, "#0.00") & ","
    sSql = sSql & " nMontoMin = " & Format(pnMontoMin, "#0.00") & ","
    sSql = sSql & " cPersCod = '" & psPersCod & "',"
    sSql = sSql & " bPreferencial=" & IIf(pbPreferencial = True, 1, 0)
    sSql = sSql & " WHERE cLineaCred = '" & psLineaCredCod & "'"
    oConecta.ConexionActiva.Execute sSql
    
   'CUSCO
   sSql = "DELETE FROM ColocLineaCreditoAgencia WHERE cLineaCred='" & psLineaCredCod & "'"
   oConecta.ConexionActiva.Execute sSql
   
    If IsArray(pMatAgencias) Then
        For i = 0 To UBound(pMatAgencias) - 1
            sSql = "INSERT INTO ColocLineaCreditoAgencia(cLineaCred,cAgeCod)"
            sSql = sSql & " VALUES('" & psLineaCredCod & "','" & pMatAgencias(i) & "')"
            oConecta.ConexionActiva.Execute sSql
        Next i
    End If
    '************************************
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Sub
ErrorActualizarLinea:
    Err.Raise Err.Number, "Actualiza Linea", Err.Description
    
End Sub
Public Sub EliminaTasasLinea(ByVal psLineaCredCod As String, ByVal psTasaTipo As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorEliminaLineaTasas
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "DELETE ColocLineaCreditoTasa Where cLineaCred = '" & psLineaCredCod & "' and nColocLinCredTasaTpo = '" & psTasaTipo & "'"
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorEliminaLineaTasas:
    Err.Raise Err.Number, "Eliminar Tasas de Linea", Err.Description
    
End Sub
Public Sub ActualizarLineaTasas(ByVal psLineaCredCod As String, ByVal psTasaTipo As String, ByVal pnTasaIni As Double, ByVal pnTasaFin As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorActualizarLineaTasas
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "UPDATE ColocLineaCreditoTasa Set "
    sSql = sSql & " nTasaIni = " & Format(pnTasaIni, "#0.0000") & ","
    sSql = sSql & " nTasaFin = " & Format(pnTasaFin, "#0.0000")
    sSql = sSql & " Where cLineaCred = '" & psLineaCredCod & "' And nColocLinCredTasaTpo = " & Trim(psTasaTipo)
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorActualizarLineaTasas:
    Err.Raise Err.Number, "Actualiza Tasas de Linea", Err.Description
    
End Sub

Public Function BuscaLineas(ByVal psBuscar As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorBuscaLineas
    sSql = "Select cLinecred from " & gConsComunes & "ColocLineCredito Where cLineaCred like '" & psBuscar & "%' Order by cLineaCred"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set BuscaLineas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorBuscaLineas:
    Err.Raise Err.Number, "Buscar Linea", Err.Description
End Function

Public Function RecuperaLineasTasas(ByVal psLineaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperaLineasTasas
    sSql = "Select C.nColocLinCredTasaTpo, C.nTasaIni, C.nTasaFin, CT.cConsDescripcion "
    sSql = sSql & " from " & gConsComunes & "ColocLineaCreditoTasa C inner join " & gConsComunes & "Constante CT ON C.nColocLinCredTasaTpo = CT.nConsValor "
    sSql = sSql & " where CT.nConsCod = " & gColocLineaCredTasas & " And  C.cLineaCred = '" & psLineaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaLineasTasas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaLineasTasas:
    Err.Raise Err.Number, "Tasas de Linea", Err.Description
    
End Function

Public Function RecuperaFondos() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "Select L.cLineaCred, P.cPersNombre, L.cAbrev "
    sSql = sSql & " from ColocLineaCredito L Inner Join Persona P ON P.cPersCod = L.cPersCod"
    sSql = sSql & " Where Len(cLineaCred) = 2 Order by cPersNombre  "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaFondos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
End Function

Public Function RecuperaInstitucionesFinancieras() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "select I.*, P.cPersNombre from "
    sSql = sSql & " InstitucionFinanc I Inner Join Persona P ON I.cPersCod = P.cPersCod"
    sSql = sSql & " where I.cIFTpo = '05'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaInstitucionesFinancieras = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
End Function

Public Function RecuperaSubFondos(ByVal psFondo As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "Select DISTINCT SUBSTRING(cLineaCred,1,4) as cSubFondo, cDescripcion, cAbrev "
    sSql = sSql & " From ColocLineaCredito "
    sSql = sSql & " Where Len(cLineaCred) = 5 and cLineaCred like '" & psFondo & "%'"
    sSql = sSql & " Group by SUBSTRING(cLineaCred,1,4), cDescripcion, cAbrev"
    sSql = sSql & " Order by cDescripcion"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSubFondos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

End Function

'**Modificado por DAOR 20070409, se aumentó el parametro pbConsiderarEstado
'**que indicará si se considera o no el estado de la linea
Public Function RecuperaLineadeCredito(ByVal psLineaCred As String, Optional ByVal pbConsiderarEstado As Boolean = True) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
On Error GoTo ErrorRecuperaLineasCredito
    sSql = "Select CL.cPersCod, LS.nMoneda, CL.cLineaCred, CL.cDescripcion, CL.nPlazoMax, CL.nPlazoMin, CL.nMontoMax, CL.nMontoMin, P.cPersNombre + space(50) + P.cPersCod as PersCod, convert(int,CL.bEstado), "
    sSql = sSql & " CLT.nTasaIni, CLT.nTasaFin, CLT2.nTasaIni as nTasaGraciaIni, CLT2.nTasaFin as nTasaGraciaFin,CLT3.nTasaIni as nTasaMoraIni, CLT3.nTasaFin as nTasaMoraFin "
    sSql = sSql & " From " & gConsComunes & "ColocLineacredito CL Inner Join Persona P ON CL.cPersCod = P.cPersCod"
    sSql = sSql & "         Left join ColocLineaCreditoSaldo LS ON CL.cLineaCred = LS.cLineaCred "
    sSql = sSql & "                           Left Join ColocLineaCreditoTasa CLT ON CL.cLineaCred = CLT.cLineaCred AND CLT.nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntCompNormal
    sSql = sSql & "                           Left Join ColocLineaCreditoTasa CLT2 ON CL.cLineaCred = CLT2.cLineaCred AND CLT2.nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntGracia
    sSql = sSql & "                           Left Join ColocLineaCreditoTasa CLT3 ON CL.cLineaCred = CLT3.cLineaCred AND CLT3.nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntMoratNormal
    sSql = sSql & " Where CL.cLineaCred = '" & psLineaCred & "'"
    'Comentado por DAOR 20070409
    'sSQL = sSQL & " AND CL.bEstado=1"
    
    '**DAOR 20070409***************************
    If pbConsiderarEstado Then
        sSql = sSql & " AND CL.bEstado=1"
    End If
    '******************************************
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaLineadeCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaLineasCredito:
    Err.Raise Err.Number, "Linea de Credito", Err.Description
End Function

Public Function RecuperaLineasCredito(ByVal pnLongitud As Integer, Optional ByVal psCriterio As String = "", Optional ByVal pcVer As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
On Error GoTo ErrorRecuperaLineasCredito
    
    If UCase$(pcVer) = "ECSA" Then
        sSql = " Select cFondo = F.cLineaCred, Fondo = F.cDescripcion,"
        sSql = sSql & " cSubFondo = SF.cLineaCred, SubFondo = SF.cDescripcion,"
        sSql = sSql & " cLineaCred = LC.cLineaCred, cDescripcion = LC.cLineaCred + '-' + LC.cDescripcion,"
        sSql = sSql & " Moneda = "
        sSql = sSql & " Case"
        sSql = sSql & "    When SubString(LC.cLineaCred, 5, 1) = 1 Then 'Soles'"
        sSql = sSql & "    When SubString(LC.cLineaCred, 5, 1) = 2 Then 'Dolares'"
        sSql = sSql & "    When SubString(LC.cLineaCred, 5, 1) = 3 Then 'VAC'"
        sSql = sSql & "    Else 'No Definido'"
        sSql = sSql & " End, "
        sSql = sSql & " LC.bEstado, LC.nPlazoMax, LC.nPlazoMin, LC.nMontoMax,"
        sSql = sSql & " LC.nMontoMin, LC.cPersCod, LC.cIFTpo, LC.cAbrev, LC.bPreferencial,"
        sSql = sSql & " LCS.nMontoTotal , LCS.nSaldoCap, LCS.nMontoColocado"
        sSql = sSql & " From ColocLineaCredito LC"
        sSql = sSql & " Inner Join ColocLineaCredito SF On SubString(LC.cLineaCred, 1, 4) = SubString(SF.cLineaCred, 1, 4)"
        sSql = sSql & " Inner Join ColocLineaCredito F On SubString(LC.cLineaCred, 1, 2) = SubString(F.cLineaCred, 1, 2)"
        sSql = sSql & " Left Join ColocLineaCreditoSaldo LCS On LC.cLineaCred = LCS.cLineaCred"
        sSql = sSql & " Where Len(LC.cLineaCred) = " & pnLongitud & " And Len(sF.cLineaCred) = 4 And Len(F.cLineaCred) = 2 And LC.bEstado = 1"
        sSql = sSql & " ORder By 1, 3, 5"
        
    Else
         sSql = "Select CL.cLineaCred, CL.cDescripcion, CL.nPlazoMax, CL.nPlazoMin, CL.nMontoMax, CL.nMontoMin, P.cPersNombre + space(50) + P.cPersCod as PersCod, convert(int,CL.bEstado) as nEstado"
         sSql = sSql & " From " & gConsComunes & "ColocLineacredito CL Inner Join Persona P ON CL.cPersCod = P.cPersCod "
        
         Select Case pnLongitud
             Case 1
                 sSql = sSql & " WHERE LEN(RTrim(CL.cLineaCred))<=2"
             Case 2
                 sSql = sSql & " WHERE LEN(RTrim(CL.cLineaCred))<=5"
             Case 3
                 sSql = sSql & " WHERE LEN(RTrim(CL.cLineaCred))<=6"
             Case 4
                 sSql = sSql & " WHERE LEN(RTrim(CL.cLineaCred))<=9"
             Case 5
                 sSql = sSql & " WHERE LEN(RTrim(CL.cLineaCred))=13" 'Ahora son 13
         End Select
         If psCriterio <> "" Then
             sSql = sSql & " AND cLineaCred like '" & psCriterio & "%' "
         End If
         sSql = sSql & " AND RTRIM(CL.cLineaCred)<>'' ORDER BY cLineaCred"
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaLineasCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaLineasCredito:
    Err.Raise Err.Number, "Linea de Credito", Err.Description
End Function

Public Function RecuperaLineasCreditoAgencia(ByVal psCodAge As String, ByVal psCodProducto As String, ByVal psMoneda As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
On Error GoTo ErrorRecuperaLineasCredito
    
    
    sSql = sSql & " Select C.cLineaCred,cDescripcion,nPlazoMax,nPlazoMin,nMontoMax,nMontoMin, "
    sSql = sSql & " nTasaCom =(select nTasaIni from dbo.ColocLineaCreditoTasa CT where C.cLineaCred=CT.cLineaCred and nColocLinCredTasaTpo=1 ) ,"
    sSql = sSql & " nTasaMora =(select nTasaIni from dbo.ColocLineaCreditoTasa CT where C.cLineaCred=CT.cLineaCred and nColocLinCredTasaTpo=3 ) ,"
    sSql = sSql & " nTasaGracia =(select nTasaIni from dbo.ColocLineaCreditoTasa CT where C.cLineaCred=CT.cLineaCred and nColocLinCredTasaTpo=5 )"
    sSql = sSql & " From dbo.ColocLineaCredito C"
    sSql = sSql & " Join  ColocLineaCreditoAgencia CA ON C.cLineaCred=CA.cLineaCred"
    sSql = sSql & " WHERE cAgeCod ='" & psCodAge & "' and substring(C.cLineaCred,7,3)='" & psCodProducto & "' and substring(C.cLineaCred,5,1)= '" & psMoneda & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaLineasCreditoAgencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaLineasCredito:
    Err.Raise Err.Number, "Linea de Credito", Err.Description
End Function


Private Sub Class_Initialize()
Dim CIni As COMConecta.DCOMClasIni
    Set CIni = New COMConecta.DCOMClasIni
    gConsPersona = CIni.BasePersonas
    gConsComunes = CIni.BaseComunes
    gConsImagenes = CIni.BaseImagenes
    Set CIni = Nothing
End Sub

Public Sub IniciaGrabado()
    Set oConn = New COMConecta.DCOMConecta
    On Error GoTo ErrorIniciaGrabado
    oConn.AbreConexion
    oConn.ConexionActiva.BeginTrans
    Exit Sub
    
ErrorIniciaGrabado:
    Err.Raise Err.Number, "Inicia Grabado de Linea", Err.Description

End Sub

Public Function Correlativo(ByVal psLinea As String) As String
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
    
    'Se Aumentaron los 2 digitos del Paquete
    'sSql = "select MAX(SUBSTRING(cLineaCred,10,2)) as nCorrel from ColocLineaCredito Where cLineaCred like '" & psLinea & "%'"
    sSql = "select MAX(SUBSTRING(cLineaCred,12,2)) as nCorrel from ColocLineaCredito Where cLineaCred like '" & psLinea & "%'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If R.RecordCount > 0 And R!nCorrel <> "" Then
        Correlativo = Right("00" & Trim(Str(IIf(IsNull(R!nCorrel), 0, R!nCorrel) + 1)), 2)
    Else
        Correlativo = "01"
    End If
    R.Close
    Set R = Nothing
End Function

Public Sub NuevaLineaCredito(ByVal psLineaCredCod As String, ByVal psDescription As String, _
    ByVal pbEstado As Integer, ByVal pnPlazoMax As Integer, ByVal pnPlazoMin As Integer, _
    ByVal pnMontoMax As Double, ByVal pnMontoMin As Double, ByVal psPersCod As String, _
    Optional ByVal psAbrev As String = "", Optional ByVal pbPreferencial As Boolean = False, _
    Optional ByVal pMatAgencias As Variant)

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
    
    On Error GoTo ErrorNuevaLineaCredito
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "INSERT INTO ColocLineaCredito(cLineaCred, cDescripcion, bEstado, nPlazoMax, nPlazoMin, nMontoMax, nMontoMin,cPersCod,cAbrev,bPreferencial) "
    sSql = sSql & " VALUES('" & psLineaCredCod & "','" & psDescription & "'," & Trim(Str(pbEstado)) & "," & Format(pnPlazoMax, "#0") & "," & Format(pnPlazoMin, "#0") & "," & Format(pnMontoMax, "#0.00") & "," & Format(pnMontoMin, "#0.00") & ",'" & psPersCod & "','" & Trim(psAbrev) & "'," & IIf(pbPreferencial = True, 1, 0) & ")"
    oConecta.ConexionActiva.Execute sSql
    
    'Para el Saldo de Linea de Credito
    sSql = "INSERT INTO ColocLineaCreditoSaldo(cLineaCred, nMontoTotal, nSaldoCap, nMontoColocado, nMoneda, nMontoReservado)"
    sSql = sSql & " VALUES('" & psLineaCredCod & "',0.00,0.00,0.00,0,0.00)"
    oConecta.ConexionActiva.Execute sSql
    
    '**** Manejo de Lineas de Credito por Agencia ****
    If IsArray(pMatAgencias) Then
        For i = 0 To UBound(pMatAgencias) - 1
            sSql = "INSERT INTO ColocLineaCreditoAgencia(cLineaCred,cAgeCod)"
            sSql = sSql & " VALUES('" & psLineaCredCod & "','" & pMatAgencias(i) & "')"
            oConecta.ConexionActiva.Execute sSql
        Next i
    End If
    '**************************************************
    
    '03-05-2006
    'Insercion en la Tabla CredSaldosAdeudoLinea cuando exista la Linea
'    Dim rs As ADODB.Recordset
'    Dim lsCodPaq As String
'    sSQL = "SELECT cCodPaq  FROM CredSaldosAdeudo WHERE LEFT(cCodPaq,5)=LEFT(" & psLineaCredCod & ",5)"
'    Set rs = oConecta.CargaRecordSet(sSQL)
'    If Not rs.EOF Then
'        lsCodPaq = rs!cCodPaq
'        sSQL = "SELECT cCodPaq FROM CredSaldosAdeudoLinea WHERE cCodPaq='" & lsCodPaq & "' AND cLineaCred='" & psLineaCredCod & "'"
'        Set rs = oConecta.CargaRecordSet(sSQL)
'        'Solo si no existe
'        If rs.EOF Then
'            sSQL = "INSERT INTO CredSaldosAdeudoLinea(cCodPaq,cLineaCred)" & _
'                "VALUES('" & lsCodPaq & "','" & psLineaCredCod & "')"
'            oConecta.ConexionActiva.Execute sSQL
'        End If
'    End If
    '**************************************************
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Sub
    
ErrorNuevaLineaCredito:
    Err.Raise Err.Number, "Nueva LineaCredito", Err.Description
    
End Sub

Public Sub NuevaLineaCreditoTasas(ByVal psLineaCredCod As String, ByVal psTasaTipo As String, ByVal pnTasaIni As Double, ByVal pnTasaFin As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorNuevaLineaCreditoTasas
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "INSERT INTO ColocLineaCreditoTasa(cLineaCred, nColocLinCredTasaTpo, nTasaIni,nTasaFin) "
    sSql = sSql & " VALUES('" & psLineaCredCod & "'," & psTasaTipo & "," & Format(pnTasaIni, "#0.0000") & "," & Format(pnTasaFin, "#0.0000") & ")"
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Sub
    
ErrorNuevaLineaCreditoTasas:
    Err.Raise Err.Number, "Nueva LienaCredito Tasas", Err.Description
End Sub

Public Sub FinalizaGrabado()
On Error GoTo ErrorFinalizaGrabado
    oConn.ConexionActiva.CommitTrans
    oConn.CierraConexion
    Set oConn = Nothing
    Exit Sub
    
ErrorFinalizaGrabado:
    oConn.ConexionActiva.RollbackTrans
    Err.Raise Err.Number, "FinalizaGrabado", Err.Description
End Sub

Public Sub EliminaLineaCredito(ByVal psLineaCred As String)
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorEliminaLineaCredito
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute "DELETE ColocLineaCredito Where cLineaCred = '" & psLineaCred & "'"
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorEliminaLineaCredito:
    Err.Raise Err.Number, "Elimina Linea de Credito", Err.Description
End Sub

Public Function RecuperaInstitucion(ByVal psLineaCred As String) As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
Dim sSql As String
    
    On Error GoTo ErrorRecuperaInstitucion
    RecuperaInstitucion = ""
    sSql = "Select cPersCod from ColocLineacredito where cLineaCred = '" & Mid(psLineaCred, 1, 2) & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        RecuperaInstitucion = R!cPersCod
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorRecuperaInstitucion:
    Err.Raise Err.Number, "", Err.Description
End Function

Public Function ObtenerLinea(ByVal psCtaCod As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Set oConec = New COMConecta.DCOMConecta
    sSql = "Select cLineaCred from colocaciones where cctacod='" & psCtaCod & "'"
    
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    ObtenerLinea = rs!cLineaCred
    Set rs = Nothing
End Function


Public Function ObtenerDatosLinea(ByVal pcLineaCred As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    Set oConec = New COMConecta.DCOMConecta
    sSql = "Select nPlazoMax,nPlazoMin,nMontoMax,nMontoMin From "
    sSql = sSql & " ColocLineaCredito Where cLineaCred='" & pcLineaCred & "'"
    
    oConec.AbreConexion
    Set ObtenerDatosLinea = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerTasaLinea(ByVal pcLineaCred As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select nColocLinCredTasaTpo,nTasaIni From ColocLineaCreditoTasa Where cLineaCred='" & pcLineaCred & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerTasaLinea = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerPreferencialLinea(ByVal pcLineaCred As String) As Boolean
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select bPreferencial=ISNULL(bPreferencial,0) From ColocLineaCredito Where cLineaCred='" & pcLineaCred & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
       ObtenerPreferencialLinea = rs!bPreferencial
    End If
    Set rs = Nothing
End Function

'*******************************************
'       Funcion que Devuelve los Valores para el llenado Inicial del Formulario
'********************************************
Public Function Cargar_Datos_Objetos_LineaCredito(ByRef prsFondos As ADODB.Recordset, _
                                                ByRef prsProductos As ADODB.Recordset, _
                                                ByRef prsAgencias As ADODB.Recordset)
    Dim oConec As COMConecta.DCOMConecta
    Dim oCred As DCOMCredito
    Dim sSql As String
    Dim rs As ADODB.Recordset
    'CUSCO
    Dim loCargaAg As COMDColocPig.DCOMColPFunciones
    
    On Error GoTo ErrorCargar_Datos_Objetos_LineaCredito
    
    Set prsFondos = RecuperaFondos
    
    Set oCred = New DCOMCredito
    Set prsProductos = oCred.RecuperaProductosDeCredito
    
    
    Set loCargaAg = New COMDColocPig.DCOMColPFunciones
    Set prsAgencias = loCargaAg.dObtieneAgencias(True)
    Set loCargaAg = Nothing
    
    
    'Set oConec = New COMConecta.DCOMConecta
    'oConec.AbreConexion
    'Set Rs = oConec.CargaRecordSet(sSQL)
    'oConec.CierraConexion
    'Set oConec = Nothing
    
    'If Not Rs.EOF And Not Rs.BOF Then
    '   Cargar_Datos_Objetos_LineaCredito = Rs!bPreferencial
    'End If
    'Set Rs = Nothing
    Exit Function
    
ErrorCargar_Datos_Objetos_LineaCredito:
    Err.Raise Err.Number, "Carga Datos Objetos", Err.Description

End Function

Public Function CargarDetalleLineaCredito(ByVal psLineaCred As String, _
                                        ByRef prsTasas As ADODB.Recordset, _
                                        ByRef prsAgencias As ADODB.Recordset)
Dim sSql As String
Dim oConec As COMConecta.DCOMConecta

On Error GoTo ErrorCargarDetalleLineaCredito

Set prsTasas = RecuperaLineasTasas(psLineaCred)

sSql = "SELECT cLineaCred,LCA.cAgeCod,A.cAgeDescripcion FROM ColocLineaCreditoAgencia LCA INNER JOIN Agencias A ON LCA.cAgeCod=A.cAgeCod" & _
       " WHERE cLineaCred='" & psLineaCred & "'"

Set oConec = New COMConecta.DCOMConecta
oConec.AbreConexion
Set prsAgencias = oConec.CargaRecordSet(sSql)
oConec.CierraConexion
Set oConec = Nothing

Exit Function

ErrorCargarDetalleLineaCredito:
    Err.Raise Err.Number, "Carga Detalle Linea Credito", Err.Description
End Function

'*************************************
'Se Aumento para el manejo de paquetes
Public Function RecuperaPaquetes(ByVal psFiltro As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    'sSQL = " SELECT DISTINCT SUBSTRING(cLineaCred,10,2) as Paquete From ColocLineaCredito WHERE cLineaCred like '" & psFiltro & "%'"
    sSql = " SELECT DISTINCT SUBSTRING(cLineaCred,10,2) as Paquete From ColocLineaCredito WHERE cLineaCred like '" & psFiltro & "%' AND LEN(cLineaCred)=13"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPaquetes = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

End Function

Public Function CorrelativoPaquete(ByVal psFiltro As String) As String
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
    
    sSql = "select MAX(SUBSTRING(cLineaCred,10,2)) as nCorrel from ColocLineaCredito Where cLineaCred like '" & psFiltro & "%'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If R.RecordCount > 0 Then
        CorrelativoPaquete = Right("00" & Trim(Str(IIf(IsNull(R!nCorrel), 0, R!nCorrel) + 1)), 2)
    Else
        CorrelativoPaquete = "01"
    End If
    R.Close
    Set R = Nothing
End Function
'*************************************

'*******************************************
'       Funcion que Devuelve los Creditos por Linea
'********************************************
Public Function Creditos_X_LineaCredito(ByVal psLineaCod As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo Error_Creditos_X_LineaCredito
    sSql = "Select C.cCtaCod, Per.cPersNombre, P.nPrdEstado, K.cConsDescripcion, P.nSaldo " _
    & "From Colocaciones C " _
    & "Inner Join Producto P On C.cCtaCod = P.cCtaCod " _
    & "Inner Join ProductoPersona PP On P.cCtaCod = PP.cCtaCod " _
    & "Inner Join Persona Per On PP.cPersCod = Per.cPersCod " _
    & "Inner Join ColocacCred CC On C.cCtaCod = CC.cCTaCod " _
    & "Inner Join Constante K On P.nPrdEstado = K.nConsValor And K.nConsCod = 3001 " _
    & "Where cLineaCred = '" & psLineaCod & "' And PP.nPrdPersRelac = 20 " _
    & "And P.nPrdEstado In (2020, 2021, 2022, 2030, 2031, 2032) "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set Creditos_X_LineaCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
Error_Creditos_X_LineaCredito:
    Err.Raise Err.Number, "Créditops por Linea de Crédito", Err.Description

End Function

'**DAOR 20070404
'**Función que devuelve las lineas de créditos de las cuentas de origen
'**de un crédito Refinanciado, En formato de Arbol que será utilizado en el proceso de refinanciación
Public Function RecuperaLineasCredOrigenRefinanciadoArbol(psCtaCod As String) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaLineasCredOrigenRefinanciadoArbol

'    sSql = "Select  case substring(CL.cLineaCred,6,1) WHEN '1' THEN 'CP1-' + CL.cLineaCred WHEN '2' THEN 'LP2-' + CL.cLineaCred END AS cLineaCred, " & _
'            " CL.cDescripcion + ' /'+ CONVERT(VARCHAR(20),CL.nMontoMin) + ' - '+ CONVERT(VARCHAR(20),CL.nMontoMax) + " & _
'            " ' / ' + convert(varchar(20),PT.nTasaInteres) + ' - '+ convert(varchar(20),nTasaInteres) + ' - ' + CR.cCtaCodRef as cDescripcion, " & _
'            " '1' as Nivel " & _
'            " From ColocacRefinanc CR inner join Colocaciones C on CR.cCtaCodRef=C.cCtaCod " & _
'            " inner join ColocLineaCredito CL on C.cLineaCred=CL.cLineaCred " & _
'            " inner join ProductoTasaInteres PT on C.cCtaCod = PT.cCtaCod and PT.nPrdTasaInteres= " & COMDConstantes.gColocLineaCredTasasIntCompNormal & _
'            " Where   CR.cCtaCod='" & psCtaCod & "' and CR.nEstado=(select max(nEstado) from ColocacRefinanc where cCtaCod='" & psCtaCod & "') " & _
'            " Order By cLineaCred,Nivel "

    sSql = "Select  case substring(CL.cLineaCred,6,1) WHEN '1' THEN 'CP1-' + CL.cLineaCred WHEN '2' THEN 'LP2-' + CL.cLineaCred END AS cLineaCred, " & _
            " CL.cDescripcion + ' /'+ CONVERT(VARCHAR(20),CL.nMontoMin) + ' - '+ CONVERT(VARCHAR(20),CL.nMontoMax) + " & _
            " ' / ' + convert(varchar(20),PT.nTasaInteres) + ' - '+ convert(varchar(20),nTasaInteres) + ' - ' + CR.cCtaCodRef as cDescripcion, " & _
            " '1' as Nivel " & _
            " From ColocacRefinanc CR inner join Colocaciones C on CR.cCtaCodRef=C.cCtaCod " & _
            " inner join ColocLineaCredito CL on C.cLineaCred=CL.cLineaCred " & _
            " inner join ProductoTasaInteres PT on C.cCtaCod = PT.cCtaCod and PT.nPrdTasaInteres= " & COMDConstantes.gColocLineaCredTasasIntCompNormal & _
            " inner join ( " & _
            "   select C.cLineaCred,min(C.dVigencia) as dVigencia " & _
            "   from ColocacRefinanc CR inner join Colocaciones C on CR.cCtaCodRef=C.cCtaCod " & _
            "   where CR.cCtaCod='" & psCtaCod & "'" & _
            "   group by C.cLineaCred " & _
            "    ) as T on C.cLineaCred=T.cLineaCred and C.dVigencia=T.dVigencia " & _
            " Where   CR.cCtaCod='" & psCtaCod & "' and CR.nEstado=(select max(nEstado) from ColocacRefinanc where cCtaCod='" & psCtaCod & "') " & _
            " Order By cLineaCred,Nivel "
            
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaLineasCredOrigenRefinanciadoArbol = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaLineasCredOrigenRefinanciadoArbol:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'**DAOR 20070407
'**Función que devuelve la linea de crédito de la cuenta de origen
'**en el proceso de refinanciación
Public Function RecuperaLineadeCredOrigenRefinanciado(ByVal psCtaCod As String, ByVal psLineaCred As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
On Error GoTo ErrorRecuperaLineadeCredOrigenRefinanciado
    sSql = "Select  CL.cPersCod,CL.cLineaCred,CL.cDescripcion,CL.nPlazoMax,CL.nPlazoMin, " & _
        " CL.nMontoMax,CL.nMontoMin,(P.cPersNombre + space(50) + P.cPersCod) as PersCod, " & _
        " Convert(int,CL.bEstado) as nEstado, PT1.nTasaInteres as nTasaIni, PT1.nTasaInteres as nTasaFin, " & _
        " PT2.nTasaInteres as nTasaMoraIni, PT2.nTasaInteres as nTasaMoraFin, " & _
        " PT3.nTasaInteres as nTasaGraciaIni,PT3.nTasaInteres as nTasaGraciaFin " & _
        "From Colocaciones C inner join ColocLineacredito CL on C.cLineaCred=CL.cLineaCred " & _
        " Inner Join Persona P ON CL.cPersCod = P.cPersCod " & _
        " Left Join ProductoTasaInteres PT1 ON C.cCtaCod = PT1.cCtaCod and PT1.nPrdTasaInteres = " & COMDConstantes.gColocLineaCredTasasIntCompNormal & _
        " Left Join ProductoTasaInteres PT2 ON C.cCtaCod = PT2.cCtaCod and PT2.nPrdTasaInteres = " & COMDConstantes.gColocLineaCredTasasIntMoratNormal & _
        " Left Join ProductoTasaInteres PT3 ON C.cCtaCod = PT3.cCtaCod and PT3.nPrdTasaInteres = " & COMDConstantes.gColocLineaCredTasasIntGracia & _
        "Where C.cCtaCod = '" & psCtaCod & "' and CL.cLineaCred='" & psLineaCred & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaLineadeCredOrigenRefinanciado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaLineadeCredOrigenRefinanciado:
    Err.Raise Err.Number, "Linea de Credito", Err.Description
End Function


Public Function ObtenerTasaLineaCredOrigenRefinanciado(ByVal psCtaCod As String, ByVal pcLineaCred As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select nColocLinCredTasaTpo,nTasaIni From ColocLineaCreditoTasa Where cLineaCred='" & pcLineaCred & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerTasaLineaCredOrigenRefinanciado = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'ALPA 20150114******************************************************************************************

Public Function ObtieneRangosMontosxTarifarios(Optional psParProd As String, Optional pdFecha As Date) As ADODB.Recordset  'psHabRang
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    sSql = "Exec stp_sel_Catalogo_GetRangosxProducto '" & psParProd & "','" & Format(pdFecha, "yyyymmdd") & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set ObtieneRangosMontosxTarifarios = rs
    Set oConec = Nothing
End Function '***NAGL 20190115

Public Function RecuperaLineadeCreditoProductoCrediticio(ByVal psTpoProdCod As String, ByVal pnCampaniaId As Integer, ByVal psPersCod As String, ByVal psLineaCred As String, ByVal psLineaCredDescripcion As String, ByVal pnMoneda As Integer, ByVal nMonto As Currency, ByVal pbPreferencial As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
On Error GoTo ErrorRecuperaLineadeCreditoProductoCrediticio

    sSql = "exec stp_sel_ObtenerTasaLineasCredito '" & psTpoProdCod & "'," & pnCampaniaId & ",'" & psPersCod & "','" & psLineaCred & "','" & psLineaCredDescripcion & "'," & pnMoneda & "," & nMonto & "," & pbPreferencial
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaLineadeCreditoProductoCrediticio = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaLineadeCreditoProductoCrediticio:
    Err.Raise Err.Number, "Linea de Credito-Producto Crediticio", Err.Description
End Function
Public Sub InsertaProductoCrediticio(ByVal psTpoProdCod As String, ByVal pnCampanaId As Integer, ByVal pnSoles As Integer, ByVal pnDolares As Integer, ByVal pnPersoneriaN As Integer, ByVal pnPersoneriaJ As Integer, ByVal pnPlazo As Integer, ByVal psPlazoMin As Integer, ByVal psPlazoMax As Integer, ByVal pbCalificacion As Integer, ByVal pbCalA As Integer, ByVal pbCalB As Integer, ByVal pbCalC As Integer, ByVal pbCalD As Integer, ByVal pbCalE As Integer, ByVal pnGenero As Integer, ByVal pnGeneroM As Integer, ByVal pnGeneroF As Integer, ByVal pnCalificacion As Integer, ByVal pnCalificacionSBSPor As Integer, ByVal pnEdad As Integer, ByVal pnEdadMin As Integer, ByVal pnEdadMax As Integer)
    On Error GoTo InsertaProductoCrediticioErr
    Dim SQL As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    SQL = "exec stp_ins_ProductoCrediticio '" & psTpoProdCod & "'," & pnCampanaId & "," & pnSoles & "," & pnDolares & "," & pnPersoneriaN & "," & pnPersoneriaJ & "," & pnPlazo & "," & psPlazoMin & "," & psPlazoMax & "," & pbCalificacion & "," & pbCalA & "," & pbCalB & "," & pbCalC & "," & pbCalD & "," & pbCalE & "," & pnGenero & "," & pnGeneroM & "," & pnGeneroF & "," & pnCalificacion & "," & pnCalificacionSBSPor & "," & pnEdad & "," & pnEdadMin & "," & pnEdadMax & ",1"
    oCon.Ejecutar SQL
    oCon.CierraConexion
    Exit Sub
InsertaProductoCrediticioErr:
    Err.Raise Err.Number, "DLineaCreditoV2:InsertaProductoCrediticio Method", Err.Description
End Sub

Public Sub InsertaProductoCrediticioNew(ByVal psTpoProdCod As String, ByVal pnCampanaId As Integer, ByVal pnSoles As Integer, ByVal pnDolares As Integer)
    On Error GoTo InsertaProductoCrediticioErr
    Dim SQL As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    SQL = "exec stp_ins_ProductoCrediticioNew '" & psTpoProdCod & "'," & pnCampanaId & "," & pnSoles & "," & pnDolares & ",1"
    oCon.Ejecutar SQL
    oCon.CierraConexion
    Exit Sub
InsertaProductoCrediticioErr:
    Err.Raise Err.Number, "DLineaCreditoV2:InsertaProductoCrediticio Method", Err.Description
End Sub 'NAGL Según ERS046-2018 20181219

Public Sub InsertaProductoCrediticioTasas(ByVal psTpoProdCod As String, ByVal pnCampanaId As Integer, ByVal pnTipoTasa As Integer, ByVal pnMoneda As Integer, ByVal pnOrden As Integer, ByVal pnMontoDesde As Currency, ByVal pnMontoHasta As Currency, ByVal pnTipoMinino As Currency, ByVal pnTipoMaximo As Currency, ByVal pbPreferencial As Integer)
    On Error GoTo InsertaProductoCrediticioTasasErr
    Dim SQL As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    SQL = "exec stp_ins_ProductoCrediticioTasas '" & psTpoProdCod & "'," & pnCampanaId & ",' " & pnTipoTasa & "'," & pnMoneda & "," & pnOrden & "," & pnMontoDesde & "," & pnMontoHasta & "," & pnTipoMinino & "," & pnTipoMaximo & "," & pbPreferencial
    oCon.Ejecutar SQL
    oCon.CierraConexion
    
    Exit Sub
InsertaProductoCrediticioTasasErr:
    Err.Raise Err.Number, "DLineaCreditoV2:InsertaProductoCrediticioTasas Method", Err.Description
End Sub
Public Sub EliminarProductoCreditocioTasas(ByVal psTpoProdCod As String, ByVal pnCampanaId As Integer, ByVal pnMoneda As Integer)
    On Error GoTo InsertaProductoCrediticioTasasErr
    Dim SQL As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    SQL = "exec stp_del_ProductoCreditocioTasas '" & psTpoProdCod & "'," & pnCampanaId & "," & pnMoneda & ""
    oCon.Ejecutar SQL
    oCon.CierraConexion
    Exit Sub
InsertaProductoCrediticioTasasErr:
    Err.Raise Err.Number, "DLineaCreditoV2:EliminarProductoCreditocioTasas Method", Err.Description
End Sub
Public Sub InsertaProductoCreditocioAgencia(ByVal psTpoProdCod As String, ByVal pnCampanaId As Integer, ByVal psAgeCod As String, ByVal pnLogico As Integer)
    On Error GoTo InsertaProductoCreditocioAgenciaErr
    Dim SQL As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    SQL = "exec stp_ins_ProductoCreditocioAgencia '" & psTpoProdCod & "'," & pnCampanaId & ",'" & psAgeCod & "'," & pnLogico
    oCon.Ejecutar SQL
    oCon.CierraConexion
    Exit Sub
InsertaProductoCreditocioAgenciaErr:
    Err.Raise Err.Number, "DLineaCreditoV2:InsertaProductoCreditocioAgencia Method", Err.Description
End Sub
Public Sub InsertaProductoCreditocioDestino(ByVal psTpoProdCod As String, ByVal pnCampanaId As Integer, ByVal psDestinoCod As String, ByVal pnLogico As Integer)
    On Error GoTo InsertaProductoCreditocioDestinoErr
    Dim SQL As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    SQL = "exec stp_ins_ProductoCreditocioDestino '" & psTpoProdCod & "'," & pnCampanaId & ",'" & psDestinoCod & "'," & pnLogico
    oCon.Ejecutar SQL
    oCon.CierraConexion
    'InsertaLineaCreditoAgencia = 0
    Exit Sub
InsertaProductoCreditocioDestinoErr:
    Err.Raise Err.Number, "DLineaCreditoV2:InsertaProductoCreditocioDestino Method", Err.Description
End Sub
Public Function ObtenerProductoCrediticio(ByVal psTpoProdCod As String, ByVal pnIdCampana, Optional psTipo As String = "") As ADODB.Recordset 'NAGL Agregó 20181215 psTipo
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ProductoCrediticio '" & psTpoProdCod & "','" & pnIdCampana & "','" & psTipo & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    Set ObtenerProductoCrediticio = rs
    oCon.CierraConexion
End Function
Public Function ObtenerProductoCrediticioTasas(ByVal psTpoProdCod As String, ByVal pnIdCampana As Integer, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ProductoCrediticioTasas '" & psTpoProdCod & "'," & pnIdCampana & ",'" & pnMoneda & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerProductoCrediticioTasas = rs
End Function
Public Function ObtenerProductoCreditocioDestino(ByVal psTpoProdCod As String, ByVal pnCampaniaId As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ProductoCreditocioDestino '" & psTpoProdCod & "'," & pnCampaniaId & ""
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerProductoCreditocioDestino = rs
End Function
Public Function ObtenerConstanteLineaCredito(ByVal pnConsCod As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ConstanteLineaCredito '" & pnConsCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerConstanteLineaCredito = rs
End Function
Public Function ObtenerProductoCreditocioAgencia(ByVal psTpoProdCod As String, ByVal pnIdCampana As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ProductoCreditocioAgencia '" & psTpoProdCod & "','" & pnIdCampana & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerProductoCreditocioAgencia = rs
End Function
Public Sub DeshabilitarProductoCrediticio(ByVal psTpoProdCod As String, ByVal nEstado As Integer, ByVal pnIdCampana As Integer)
    On Error GoTo DeshabilitarProductoCrediticioErr
    Dim SQL As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    SQL = "exec stp_upd_ProductoCrediticio_deshabilitar '" & psTpoProdCod & "'," & pnIdCampana & "," & nEstado & ""
    oCon.Ejecutar SQL
    oCon.CierraConexion
    Exit Sub
DeshabilitarProductoCrediticioErr:
    Err.Raise Err.Number, "DLineaCreditoV2:DeshabilitarProductoCrediticio Method", Err.Description
End Sub
Public Function ObtenerCampanaConfguracion() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ObtieneCampanasConfiguracion "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    Set ObtenerCampanaConfguracion = rs
    oCon.CierraConexion
End Function
'*******************************************************************************************************
'***FRHU 20170915 ERS049-2017
Public Sub InsertarColocHistorialTasaMaxima(ByVal psCtaCod As String, ByVal pnTasaIntIni As Double, ByVal pnTasaIntFin As Double, ByVal pnTasaInt As Double, ByVal sMovNro As String, ByVal pnOperacion As Integer, ByVal psCodCargo As String)
    On Error GoTo InsertarColocHistorialTasaMaximaErr
    Dim SQL As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    SQL = "exec stp_ins_del_ColocHistorialTasaMaxima '" & psCtaCod & "'," & pnTasaIntIni & "," & pnTasaIntFin & "," & pnTasaInt & ",'" & sMovNro & "'," & pnOperacion & ",'" & psCodCargo & "'"
    oCon.Ejecutar SQL
    oCon.CierraConexion
    Exit Sub
InsertarColocHistorialTasaMaximaErr:
    Err.Raise Err.Number, "DCOMLineaCredito:InsertarColocHistorialTasaMaxima Method", Err.Description
End Sub
'FIN FRHU 20170915 ERS049-2017

