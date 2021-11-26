VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLineaCredito"
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
Private oConn As DConecta

'*****************************************************************************************
'******* Recupera las Lineas de Credito de un Producto
'*****************************************************************************************
Public Function RecuperaLineasProducto(ByVal psProducto As String, ByVal psMoneda As String) As ADODB.Recordset
Dim oConecta As DConecta
Dim sSql As String
    
    On Error GoTo ErrorRecuperaLineasProducto
    sSql = "Select CL.cLineaCred, CL.cDescripcion, CL.nPlazoMax, CL.nPlazoMin, CL.nMontoMax, CL.nMontoMin, "
    sSql = sSql & " CLT.nTasaIni, CLT.nTasaFin, CLT2.nTasaIni as nTasaGraciaIni, CLT2.nTasaFin as nTasaGraciaFin,CLT3.nTasaIni as nTasaMoraIni, CLT3.nTasaFin as nTasaMoraFin "
    sSql = sSql & " From ColocLineaCredito CL Inner Join ColocLineaCreditoTasa CLT ON CL.cLineaCred = CLT.cLineaCred AND CLT.nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntCompNormal
    sSql = sSql & "                           Left Join ColocLineaCreditoTasa CLT2 ON CL.cLineaCred = CLT2.cLineaCred AND CLT2.nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntGracia
    sSql = sSql & "                           Left Join ColocLineaCreditoTasa CLT3 ON CL.cLineaCred = CLT3.cLineaCred AND CLT3.nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntMoratNormal
    sSql = sSql & " Where substring(CL.cLineaCred,7,3) = '" & psProducto & "' AND substring(CL.cLineaCred,5,1) = '" & psMoneda & "' Order By cDescripcion"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaLineasProducto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaLineasProducto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaLineasProductoArbol(ByVal psProducto As String, ByVal psMoneda As String, _
Optional ByVal pbPreferencial As Boolean = False) As ADODB.Recordset
Dim oConecta As DConecta
Dim sSql As String
Dim sparte1 As String
Dim sParte2 As String
Dim sParte3 As String

sparte1 = "CP1-0000" + psMoneda + "1" + psProducto + "00"
sParte2 = "LP2-0000" + psMoneda + "2" + psProducto + "00"

    On Error GoTo ErrorRecuperaLineasProductoArbol

    sSql = "Select '" & sparte1 & "'  as cLineaCred, 'CORTO PLAZO' as cDescripcion, '1' as Nivel "
    sSql = sSql & "UNION "
    sSql = sSql & "Select '" & sParte2 & "'   as cLineaCred, 'LARGO PLAZO' as cDescripcion, '1' as Nivel "
    sSql = sSql & "UNION "
    sSql = sSql & "Select CASE substring(CL.cLineaCred,6,1) WHEN '1' THEN 'CP1-' + CL.cLineaCred WHEN '2' THEN 'LP2-' + CL.cLineaCred END AS cLineaCred, CL.cDescripcion + ' /'+ CONVERT(VARCHAR(20),NMONTOMIN) + ' - '+ CONVERT(VARCHAR(20),NMONTOMAX) + ' / ' + CONVERT(VARCHAR(20),nTasaIni) + ' - '+ CONVERT(VARCHAR(20),nTasaFin) AS cDescripcion, '2' as Nivel  "
    sSql = sSql & " From ColocLineaCredito CL JOIN  COLOCLINEACREDITOTASA T ON T.CLINEACRED = CL.CLINEACRED AND T.nColocLinCredTasaTpo = 1 "
    sSql = sSql & " Where substring(CL.cLineaCred,7,3) = '" & psProducto & "' AND substring(CL.cLineaCred,5,1) = '" & psMoneda & "' "
    sSql = sSql & " AND CL.bEstado=1 "
    If pbPreferencial = True Then
    sSql = sSql & " AND CL.bPreferencial=1"
    End If
    sSql = sSql & " ORDER BY cLineaCred, Nivel"
    Set oConecta = New DConecta
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
    Set oConn = New DConecta
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
        Set oConn = New DConecta
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
Optional ByVal pbPreferencial As Boolean = False)
Dim sSql As String
Dim oConecta As DConecta
    On Error GoTo ErrorActualizarLinea
    Set oConecta = New DConecta
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
    Exit Sub
    oConecta.CierraConexion
    Set oConecta = Nothing
ErrorActualizarLinea:
    Err.Raise Err.Number, "Actualiza Linea", Err.Description
    
End Sub
Public Sub EliminaTasasLinea(ByVal psLineaCredCod As String, ByVal psTasaTipo As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaLineaTasas
    Set oConecta = New DConecta
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
Dim oConecta As DConecta

    On Error GoTo ErrorActualizarLineaTasas
    Set oConecta = New DConecta
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
Dim oConecta As DConecta

On Error GoTo ErrorBuscaLineas
    sSql = "Select cLinecred from " & gConsComunes & "ColocLineCredito Where cLineaCred like '" & psBuscar & "%' Order by cLineaCred"
    Set oConecta = New DConecta
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
Dim oConecta As DConecta

On Error GoTo ErrorRecuperaLineasTasas
    sSql = "Select C.cLineaCred, C.nColocLinCredTasaTpo, C.nTasaIni, C.nTasaFin, CT.cConsDescripcion "
    sSql = sSql & " from " & gConsComunes & "ColocLineaCreditoTasa C inner join " & gConsComunes & "Constante CT ON C.nColocLinCredTasaTpo = CT.nConsValor "
    sSql = sSql & " where CT.nConsCod = " & gColocLineaCredTasas & " And  C.cLineaCred = '" & psLineaCod & "'"
    
    Set oConecta = New DConecta
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
Dim oConecta As DConecta

    sSql = "Select L.cLineaCred, P.cPersNombre, L.cAbrev "
    sSql = sSql & " from ColocLineaCredito L Inner Join Persona P ON P.cPersCod = L.cPersCod"
    sSql = sSql & " Where Len(cLineaCred) = 2 Order by cPersNombre  "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaFondos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
End Function

Public Function RecuperaInstitucionesFinancieras() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    sSql = "select I.*, P.cPersNombre from "
    sSql = sSql & " InstitucionFinanc I Inner Join Persona P ON I.cPersCod = P.cPersCod"
    sSql = sSql & " where I.cIFTpo = '05'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaInstitucionesFinancieras = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
End Function

Public Function RecuperaSubFondos(ByVal psFondo As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    sSql = "Select DISTINCT SUBSTRING(cLineaCred,1,4) as cSubFondo, cDescripcion, cAbrev "
    sSql = sSql & " From ColocLineaCredito "
    sSql = sSql & " Where Len(cLineaCred) = 5 and cLineaCred like '" & psFondo & "%'"
    sSql = sSql & " Group by SUBSTRING(cLineaCred,1,4), cDescripcion, cAbrev"
    sSql = sSql & " Order by cDescripcion"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaSubFondos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

End Function

Public Function RecuperaLineadeCredito(ByVal psLineaCred As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
On Error GoTo ErrorRecuperaLineasCredito
    sSql = "Select CL.cPersCod, LS.nMoneda, CL.cLineaCred, CL.cDescripcion, CL.nPlazoMax, CL.nPlazoMin, CL.nMontoMax, CL.nMontoMin, P.cPersNombre + space(50) + P.cPersCod as PersCod, convert(int,CL.bEstado), "
    sSql = sSql & " CLT.nTasaIni, CLT.nTasaFin, CLT2.nTasaIni as nTasaGraciaIni, CLT2.nTasaFin as nTasaGraciaFin,CLT3.nTasaIni as nTasaMoraIni, CLT3.nTasaFin as nTasaMoraFin "
    sSql = sSql & " From " & gConsComunes & "ColocLineacredito CL Inner Join Persona P ON CL.cPersCod = P.cPersCod"
    sSql = sSql & "         Left join ColocLineaCreditoSaldo LS ON CL.cLineaCred = LS.cLineaCred "
    sSql = sSql & "                           Left Join ColocLineaCreditoTasa CLT ON CL.cLineaCred = CLT.cLineaCred AND CLT.nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntCompNormal
    sSql = sSql & "                           Left Join ColocLineaCreditoTasa CLT2 ON CL.cLineaCred = CLT2.cLineaCred AND CLT2.nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntGracia
    sSql = sSql & "                           Left Join ColocLineaCreditoTasa CLT3 ON CL.cLineaCred = CLT3.cLineaCred AND CLT3.nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntMoratNormal
    sSql = sSql & " Where CL.cLineaCred = '" & psLineaCred & "'"
    sSql = sSql & " AND CL.bEstado=1"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaLineadeCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaLineasCredito:
    Err.Raise Err.Number, "Linea de Credito", Err.Description
End Function




Public Function RecuperaLineasCredito(ByVal pnLongitud As Integer, Optional ByVal psCriterio As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
On Error GoTo ErrorRecuperaLineasCredito
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
            sSql = sSql & " WHERE LEN(RTrim(CL.cLineaCred))=11"
    End Select
    If psCriterio <> "" Then
        sSql = sSql & " AND cLineaCred like '" & psCriterio & "%' "
    End If
    sSql = sSql & " AND RTRIM(CL.cLineaCred)<>'' ORDER BY cLineaCred"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaLineasCredito = oConecta.CargaRecordSet(sSql)
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
    Set oConn = New DConecta
    On Error GoTo ErrorIniciaGrabado
    oConn.AbreConexion
    oConn.ConexionActiva.BeginTrans
    Exit Sub
    
ErrorIniciaGrabado:
    Err.Raise Err.Number, "Inicia Grabado de Linea", Err.Description

End Sub

Public Function Correlativo(ByVal psLinea As String) As String
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset
    sSql = "select MAX(SUBSTRING(cLineaCred,10,2)) as nCorrel from ColocLineaCredito Where cLineaCred like '" & psLinea & "%'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If R.RecordCount > 0 Then
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
    Optional ByVal psAbrev As String = "", Optional ByVal pbPreferencial As Boolean = False)

Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorNuevaLineaCredito
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "INSERT INTO ColocLineaCredito(cLineaCred, cDescripcion, bEstado, nPlazoMax, nPlazoMin, nMontoMax, nMontoMin,cPersCod,cAbrev,bPreferencial) "
    sSql = sSql & " VALUES('" & psLineaCredCod & "','" & psDescription & "'," & Trim(Str(pbEstado)) & "," & Format(pnPlazoMax, "#0") & "," & Format(pnPlazoMin, "#0") & "," & Format(pnMontoMax, "#0.00") & "," & Format(pnMontoMin, "#0.00") & ",'" & psPersCod & "','" & Trim(psAbrev) & "'," & IIf(pbPreferencial = True, 1, 0) & ")"
    oConecta.ConexionActiva.Execute sSql
    'Para el Saldo de Linea de Credito
    sSql = "INSERT INTO ColocLineaCreditoSaldo(cLineaCred, nMontoTotal, nSaldoCap, nMontoColocado, nMoneda, nMontoReservado)"
    sSql = sSql & " VALUES('" & psLineaCredCod & "',0.00,0.00,0.00,0,0.00)"
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Sub
    
ErrorNuevaLineaCredito:
    Err.Raise Err.Number, "Nueva LienaCredito", Err.Description
    
End Sub

Public Sub NuevaLineaCreditoTasas(ByVal psLineaCredCod As String, ByVal psTasaTipo As String, ByVal pnTasaIni As Double, ByVal pnTasaFin As Double)
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorNuevaLineaCreditoTasas
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminaLineaCredito
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute "DELETE ColocLineaCredito Where cLineaCred = '" & psLineaCred & "'"
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorEliminaLineaCredito:
    Err.Raise Err.Number, "Elimina Linea de Credito", Err.Description
End Sub

Public Function RecuperaInstitucion(ByVal psLineaCred As String) As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset
Dim sSql As String
    
    On Error GoTo ErrorRecuperaInstitucion
    RecuperaInstitucion = ""
    sSql = "Select cPersCod from ColocLineacredito where cLineaCred = '" & Mid(psLineaCred, 1, 2) & "'"
    Set oConecta = New DConecta
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

Public Function ObtenerLinea(ByVal psctacod As String) As String
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    sSql = "Select cLineaCred from colocaciones where cctacod='" & psctacod & "'"
    
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    ObtenerLinea = rs!cLineaCred
    Set rs = Nothing
End Function


Public Function ObtenerDatosLinea(ByVal pcLineaCred As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    Set oConec = New DConecta
    sSql = "Select nPlazoMax,nPlazoMin,nMontoMax,nMontoMin From "
    sSql = sSql & " ColocLineaCredito Where cLineaCred='" & pcLineaCred & "'"
    
    oConec.AbreConexion
    Set ObtenerDatosLinea = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerTasaLinea(ByVal pcLineaCred As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select nColocLinCredTasaTpo,nTasaIni From ColocLineaCreditoTasa Where cLineaCred='" & pcLineaCred & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ObtenerTasaLinea = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerPreferencialLinea(ByVal pcLineaCred As String) As Boolean
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select bPreferencial From ColocLineaCredito Where cLineaCred='" & pcLineaCred & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
       ObtenerPreferencialLinea = rs!bPreferencial
    End If
    Set rs = Nothing
End Function
