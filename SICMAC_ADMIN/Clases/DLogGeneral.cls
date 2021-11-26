VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogGeneral"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsCentralPer As String
Dim vsCentralCom As String
Dim vsCentralImg As String
Dim sSql As String
Dim oconect As DConecta

Public Enum TpoCargaUsu
    UsuUnRegistro = 0
    UsuTodosArea = 1
    UsuTodos = 2
End Enum

Public Enum TpoCargaArea
    AreaTotal = 0
    AreaAutoriza = 1
    AreaComite = 2
    AreaPersona = 3
End Enum

Public Enum TpoCargaAgencia
    AgeUnRegistro = 1
    AgeTotal = 5
End Enum


Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    Dim oIni As ClasIni
    
    Set oIni = New ClasIni
        vsConexion = oIni.CadenaConexion
        vsCentralPer = oIni.BasePersonas
        vsCentralCom = oIni.BaseComunes
        vsCentralImg = oIni.BaseImagenes
    Set oIni = Nothing
    'EJVG20131007 ***
    Set oconect = New DConecta
    If oconect.AbreConexion = False Then
        Call RaiseError(MyUnhandledError, "DLogGeneral:Initialize Method. Error en Conexion a Base de datos")
    End If
    'END EJVG
End Sub
'EJVG20131007 ***
Public Sub dBeginTrans()
    oconect.BeginTrans
End Sub
Public Sub dCommitTrans()
    oconect.CommitTrans
End Sub
Public Sub dRollbackTrans()
    oconect.RollBackTrans
End Sub
Private Sub Class_Terminate()
    oconect.CierraConexion
    Set oconect = Nothing
End Sub
'END EJVG *******
Public Function GetnMovNro(ByVal psMovNro As String) As Long
Dim oConec As DConecta
Set oConec = New DConecta
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
If oConec.AbreConexion() Then
    sql = "Select nMovNro From Mov where cMovNro ='" & psMovNro & "'"
    Set rs = oConec.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        GetnMovNro = rs!nMovNro
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
End If
End Function

Public Function GetsMovNro(ByVal pnMovNro As Integer) As String
Dim oConec As DConecta
Set oConec = New DConecta
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
If oConec.AbreConexion() Then
    sql = "Select cMovNro From Mov where nMovNro = " & pnMovNro & " "
    Set rs = oConec.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        GetsMovNro = rs!cMovNro
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
End If
End Function

'Devuelve un String con un Nro de Mov
Public Function GeneraMov(ByVal pdFecSis As Date, ByVal psCodCmac As String, _
    ByVal psCodAge As String, ByVal psCodUser As String) As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    If oConec.AbreConexion() Then
        GeneraMov = Format(pdFecSis & " " & oConec.GetHoraServer, gsFormatoMovFechaHora) & psCodCmac & Right(psCodAge, 2) & "00" & psCodUser
    Else
        GeneraMov = ""
    End If
    Set oConec = Nothing
End Function

'Devuelve un RecordSet con las Areas
Public Function CargaArea(ByVal psTpoCarga As TpoCargaArea, Optional ByVal psCodigo As String = "") As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If psTpoCarga = AreaTotal Then
            'Todos las Areas
            sSql = "SELECT a.cAreaCod, a.cAreaDescripcion, a.cAreaEstruc " & _
                " FROM  Areas A " & _
                " ORDER BY a.cAreaDescripcion "
        ElseIf psTpoCarga = AreaAutoriza Then
            If Len(Trim(psCodigo)) = 0 Then
                sSql = "SELECT a.cAreaCod, a.cAreaDescripcion " & _
                    " FROM  Areas A " & _
                    " WHERE a.nLogTramite = " & gLogAreaTraEstadoAcepta & " " & _
                    " ORDER BY a.cAreaDescripcion  "
            Else
                sSql = "SELECT a.cAreaCod, a.cAreaDescripcion " & _
                    " FROM  Areas A " & _
                    " WHERE a.nLogTramite = " & gLogAreaTraEstadoAcepta & " AND a.cAreaCod = '" & psCodigo & "'"
            End If
        ElseIf psTpoCarga = AreaComite Then
            sSql = "SELECT a.cAreaCod, a.cAreaDescripcion, rh.cPersCod, p.cPersNombre " & _
                " FROM  Areas A JOIN RRHH RH ON a.cAreaCod = rh.cAreaCod AND rh.dCese IS NULL " & _
                "       JOIN " & vsCentralPer & "Persona P ON rh.cPersCod = p.cPersCod " & _
                " WHERE a.nLogComite = " & gLogAreaTraEstadoActivado & " " & _
                " ORDER BY a.cAreaDescripcion  "
        ElseIf psTpoCarga = AreaPersona Then
            sSql = "SELECT a.cAreaCod, a.cAreaDescripcion " & _
                " FROM  Areas A JOIN RRHH RH ON a.cAreaCod = rh.cAreaCod AND rh.dCese IS NULL " & _
                " WHERE rh.cPersCod = '" & psCodigo & "'"
        Else
            sSql = ""
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaArea = rs
End Function
'Devuelve un RecordSet con el Area siguiente para el trámite
Public Function CargaAreaSuperior(ByVal psAreaStru As String, Optional ByVal pbAreaFinal As Boolean = False) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If pbAreaFinal = False Then
            sSql = " SELECT cAreaCod, cAreaDescripcion " & _
                " FROM Areas " & _
                " WHERE cAreaEstruc = substring('" & psAreaStru & "',1,len('" & psAreaStru & "') - 2) " & _
                " AND nLogTramite >= " & gLogAreaTraEstadoActivado & " " & _
                " "
                ' AND (cLogTramite <> '" & gLogAreaTraEstadoAcepta & "' OR )
        Else
            sSql = " SELECT cAreaCod, cAreaDescripcion " & _
                " FROM Areas " & _
                " WHERE nLogTramite = " & gLogAreaTraEstadoAcepta & " " & _
                " AND cAreaEstruc <> '" & psAreaStru & "'"
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaAreaSuperior = rs
End Function

'Devuelve un RecordSet con la relación usuario-area
Public Function CargaUsuario(ByVal pnTpoCarga As TpoCargaUsu, ByVal psUserArea As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If pnTpoCarga = UsuUnRegistro Then
            'Un usuario
            sSql = "SELECT rh.cAreaCod , a.cAreaDescripcion, p.cPersNombre, a.cAreaEstruc, isnull(a.cLogTramite,'0') cLogTramite " & _
                " FROM  RRHH RH INNER JOIN Areas A ON rh.cAreaCod = a.cAreaCod " & _
                "       JOIN " & vsCentralPer & "PERSONA P ON P.cPersCod = RH.cPersCod  " & _
                " WHERE  rh.cUser = '" & psUserArea & "' "
        ElseIf pnTpoCarga = UsuTodosArea Then
            'Todos los usuarios por area
            sSql = "SELECT rh.cPersCod, rh.cUser, p.cPersNombre, p.cPersEstado " & _
                " FROM  RRHH RH INNER JOIN Areas A ON rh.cAreaCod = a.cAreaCod and rh.cAreaCod = '" & psUserArea & "'" & _
                "       JOIN " & vsCentralPer & "PERSONA P ON P.cPersCod = RH.cPersCod  " & _
                " ORDER BY p.cPersNombre "
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaUsuario = rs
End Function

'Devuelve un RecordSet de acuerdo a los periodos para trabajar
Public Function CargaPeriodo() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT nPerNro " & _
            " FROM LogPeriodo " & _
            " WHERE nPerEstado = " & gLogProvEstadoActivado & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaPeriodo = rs
End Function

'Devuelve un RecordSet de acuerdo a la constante deseada
Public Function CargaParametro(ByVal pnParProd As Integer, ByVal pnParCod As Integer) As Currency
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT IsNull(p.nParValor,0) Campo " & _
            " FROM Parametro P " & _
            " WHERE p.nParProd = " & pnParProd & " AND p.nParCod = " & pnParCod & ""
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    CargaParametro = rs!Campo
End Function

'Devuelve un RecordSet de acuerdo a la constante deseada
Public Function CargaConstante(ByVal pnConsCod As Integer, _
Optional ByVal pbPrimeroValor As Boolean = True, _
Optional ByVal pnOtro As Integer = 0) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = IIf(pbPrimeroValor, "SELECT c.nConsValor, c.cConsDescripcion ", "SELECT c.cConsDescripcion, c.nConsValor ")
        If pnOtro = 0 Then
            sSql = sSql & _
                " FROM Constante C " & _
                " WHERE nConsCod = " & pnConsCod & " AND Not nConsValor = " & pnConsCod & " " & _
                " ORDER BY c.nConsValor "
        Else
            sSql = sSql & _
                " FROM Constante C " & _
                " WHERE nConsCod = " & pnConsCod & " AND Not nConsValor = " & pnConsCod & " " & _
                " AND nConsValor = " & pnOtro & " " & _
                " ORDER BY c.nConsValor "
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaConstante = rs
End Function


'Carga un RecordSet de acuerdo a la constante deseada
Public Function ConsCarga(ByVal pnConsCod As Integer, ByVal pnConsValor As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT c.nConsCod, c.nConsValor, c.cConsDescripcion " & _
            " FROM Constante C " & _
            " WHERE nConsCod = " & pnConsCod & " AND nConsValor = " & pnConsValor & " "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set ConsCarga = rs
End Function
'Agrega una Constante
Public Function ConsGraba(ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, _
ByVal psConsDescripcion As String) As Integer
    Dim sActualiza As String
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        sSql = "INSERT INTO Constante(nConsCod, nConsValor, cConsDescripcion) " & _
            " VALUES (" & pnConsCod & "," & pnConsValor & ",'" & psConsDescripcion & "')"
        oConec.Ejecutar sSql
        
        ConsGraba = 0
        oConec.CierraConexion
    Else
        ConsGraba = 1
    End If
End Function
'Modifica una Constante
Public Function ConsModifi(ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, _
ByVal psConsDescripcion As String) As Integer
    Dim sActualiza As String
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        
        sSql = "UPDATE Constante SET cConsDescripcion = '" & psConsDescripcion & "' " & _
            " WHERE nConsCod = " & pnConsCod & " AND nConsValor = " & pnConsValor & ""
        oConec.Ejecutar sSql
        
        ConsModifi = 0
        oConec.CierraConexion
    Else
        ConsModifi = 1
    End If
End Function
'Devuelve un nuevo código para las Constantes (nuevo cConsValor)
Public Function ConsCodNue(ByVal pnConsCod As Integer) As Integer
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sSql = "SELECT isnull(MAX(c.nConsValor),0) Campo" & _
            " FROM Constante C WHERE c.nConsCod = " & pnConsCod & " AND " & _
            " c.nConsValor <> " & pnConsCod & ""
        
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        ConsCodNue = rs!Campo + 1
    Else
        ConsCodNue = 0
    End If
End Function


'Devuelve un RecordSet con las Areas
Public Function CargaAgencia(ByVal pnTpoCarga As TpoCargaAgencia, _
Optional ByVal psAgencia As String = "") As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        If pnTpoCarga = AgeTotal Then
            'Todos las Areas
            sSql = "SELECT a.cAgeCod, a.cAgeDescripcion, '1' " & _
                " FROM  Agencias A " & _
                " ORDER BY a.cAgeCod "
        ElseIf pnTpoCarga = AgeUnRegistro Then
            sSql = "SELECT a.cAgeCod, a.cAgeDescripcion, a.cAgeDireccion " & _
                " FROM  Agencias A " & _
                " WHERE a.cAgeCod = '" & psAgencia & "' "
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set CargaAgencia = rs
End Function

Public Function GetOpeSerie(psOpeCod As String) As String
    Dim sql As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sql = "Select cSerie From Opeserie Where cOpeCod = '" & psOpeCod & "'"
        Set rs = oConec.CargaRecordSet(sql)
    End If
    
    If rs.EOF And rs.BOF Then
        GetOpeSerie = ""
    Else
        GetOpeSerie = rs.Fields(0)
    End If
    
    rs.Close
    
    Set oConec = Nothing
    Set rs = Nothing
End Function

Public Function GetOpeDocTpo(psOpeCod As String) As String
    Dim sql As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    
    If oConec.AbreConexion() Then
        sql = "Select nDocTpo From OpeDoc Where cOpeCod = '" & psOpeCod & "'"
        Set rs = oConec.CargaRecordSet(sql)
    End If
    
    If rs.EOF And rs.BOF Then
        GetOpeDocTpo = ""
    Else
        GetOpeDocTpo = rs.Fields(0)
    End If
    
    rs.Close
    
    Set oConec = Nothing
    Set rs = Nothing
End Function


Public Function CargaMotRechazos() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "Select cDescRechazo,cCodRechazo  from  LogReqMotRechazo where cestadoRech = 0 "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaMotRechazos = rs
End Function

Public Function GetEstadoRequerimiento(ByVal psReqCod As String) As String
Dim oConec As DConecta
Set oConec = New DConecta
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
If oConec.AbreConexion() Then
    sql = "select nLogReqEstado from logrequerimiento where  nLogReqCod = '" & Trim(psReqCod) & "'"
    Set rs = oConec.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        GetEstadoRequerimiento = rs!nLogReqEstado
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
End If
End Function

Public Function ActualizaEstReq(ByVal pscodreq As String, pscodrechazo As String, psEstado As TpoReqEstado, psaprorech As String) As Integer
On Error GoTo InsertaMovRefErr
    Dim sSql As String, sActualiza As String
    Dim oConec As DConecta
    ActualizaEstReq = 1
    Set oConec = New DConecta
    If oConec.AbreConexion() Then
       Select Case psaprorech
       Case "A"
            sSql = "Update logrequerimiento " _
            & " set nLogReqEstado =  " & psEstado & ", " _
            & " cCodRechazo = NULL  " _
            & " where nLogReqCod  = '" & pscodreq & "' "
       Case "R"
            sSql = "Update logrequerimiento " _
            & " set  cCodRechazo = '" & pscodrechazo & "', " _
            & " nLogReqEstado =  " & psEstado & " " _
            & " where nLogReqCod  = '" & pscodreq & "' "
       End Select
       oConec.Ejecutar sSql
       ActualizaEstReq = 0
       oConec.CierraConexion
    End If
    Exit Function
InsertaMovRefErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaMovRef Method")
End Function

'WIOR 20120811 ******************************************************************
Public Function RegistrarContrato(ByVal pcNContrato As String, ByVal pcPersCodProv As String, ByVal pcAreaUsuaria As String, _
                                ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pnMoneda As Integer, ByVal pnMonto As Double, _
                                ByVal pnCuotas As Integer, ByVal pnTipoPago, ByVal pnTipoContrato As Integer, ByVal pdFecFirma As Date, _
                                ByVal pcGlosa As String, ByVal pcNombreArchivo As String, ByVal pnTipoGarantia As Integer, _
                                ByVal pnMonedaGarantia As Integer, ByVal pnMontoGarantia As Double, ByVal pnEstado As Integer) As Integer
On Error GoTo ErrorRegistrarContrato
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        sSql = "stp_ins_ContratoProveedor '" & pcNContrato & "','" & pcPersCodProv & "','" & pcAreaUsuaria & "','" & _
                Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "'," & pnMoneda & "," & pnMonto & "," & _
                pnCuotas & "," & pnTipoPago & "," & pnTipoContrato & ",'" & Format(pdFecFirma, "yyyymmdd") & "','" & pcGlosa & "','" & _
                pcNombreArchivo & "'," & pnTipoGarantia & "," & pnMonedaGarantia & "," & pnMontoGarantia & "," & pnEstado
        
        
        oConec.Ejecutar sSql
        
        RegistrarContrato = 0
        oConec.CierraConexion
    Else
        RegistrarContrato = 1
    End If
    Exit Function
ErrorRegistrarContrato:
    Err.Raise Err.Number, "Registrar Contrato", Err.Description
End Function
Public Function RegistrarCronogramaContrato(ByVal pcNContrato As String, ByVal pnnpago As Integer, ByVal pdFecPago As Date, ByVal pnMoneda As Integer, ByVal pnMonto As Double, ByVal pnEstado As Integer, ByVal pnNAdenda As Integer) As Integer
On Error GoTo ErrorRegistrarCronogramaContrato
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        'sSql = "stp_ins_ContratoCronograma '" & pcNContrato & "'," & pnNPago & ",'" & Format(pdFecPago, "yyyymmdd") & "'," & pnMoneda & "," & pnMonto & "," & pnEstado & "," & pnNAdenda 'Comentado ERS0772014
        sSql = "stp_ins_ERS0772014_ContratoCronograma '" & pcNContrato & "'," & pnnpago & ",'" & Format(pdFecPago, "yyyymmdd") & "'," & pnMoneda & "," & pnMonto & "," & pnEstado & "," & pnNAdenda
        
        
        oConec.Ejecutar sSql
        
        RegistrarCronogramaContrato = 0
        oConec.CierraConexion
    Else
        RegistrarCronogramaContrato = 1
    End If
    Exit Function
ErrorRegistrarCronogramaContrato:
    Err.Raise Err.Number, "Registrar Contrato", Err.Description
End Function

Public Function ExisteNContrato(ByVal pcNContrato As String, Optional ByVal pncontref As Integer = 0) As Integer
On Error GoTo ErrorRegistrarCronogramaContrato
    Dim oConec As DConecta
    Dim rs As Recordset
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
'        If pncontref = 0 Then
'            sSql = "select cNContrato from ContratoProveedor where cNContrato='" & pcNContrato & "'"
        'Else
            sSql = "stp_sel_ERS0772014_ExisteContratoNContrato '" & pcNContrato & "'"  'PASI20140822 TI-ERS077-2014
        'End If
        
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        If rs.RecordCount > 0 Then
            ExisteNContrato = 0
        Else
            ExisteNContrato = 1
        End If
    Else
        ExisteNContrato = 1
    End If
    Exit Function
ErrorRegistrarCronogramaContrato:
    Err.Raise Err.Number, "Registrar Contrato", Err.Description
End Function

Public Function ListarContratosPorEstado(ByVal pnEstado As String) As ADODB.Recordset
On Error GoTo ErrorListarContratosPorEstado
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        'sSql = "stp_sel_ContratosyAdendasPorEstado '" & pnEstado & "'" 'Comentado PASIERS0772014
        sSql = "stp_sel_ERS0772014_ContratosyAdendasPorEstado '" & pnEstado & "'" 'PASIERS0772014
    
        Set ListarContratosPorEstado = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    Else
        Set ListarContratosPorEstado = Nothing
    End If
    Exit Function
ErrorListarContratosPorEstado:
    Err.Raise Err.Number, "Listar Contrato", Err.Description
End Function
Public Function NombreArchivoNContrato(ByVal pcNContrato As String, Optional ByVal pncontref As Integer = 0) As String 'pnContRef agregado pasiERS0772014
On Error GoTo ErrorNombreArchivoNContrato
    Dim oConec As DConecta
    Dim rs As Recordset
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        'sSql = "select cNombreArchivo from ContratoProveedor where cNContrato='" & pcNContrato & "'" comentado pasi20140823 ti-ers077-2014
        If pncontref = 0 Then
            sSql = "stp_sel_NombreArchivoContrato '" & pcNContrato & "'"
        Else
            sSql = "stp_sel_ERS0772014_NombreArchivoContrato '" & pcNContrato & "'," & pncontref
        End If
        
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        If rs.RecordCount > 0 Then
            NombreArchivoNContrato = Trim(rs!cNombreArchivo)
        Else
            NombreArchivoNContrato = ""
        End If
    Else
        NombreArchivoNContrato = ""
    End If
    Exit Function
ErrorNombreArchivoNContrato:
    Err.Raise Err.Number, "Archivo de Contrato", Err.Description
End Function
Public Function ListarCronogramaDeContratos(ByVal psncontrato As String, Optional ByVal pncontref As Integer = 0) As ADODB.Recordset 'nContRef Agregado PASI20140823 TI-ERS077-2014
On Error GoTo ErrorListarCronogramaDeContratos
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        If pncontref = 0 Then
            sSql = "stp_sel_RS591518_DetalleCronograma '" & psncontrato & "'," & pncontref
        Else
            sSql = "stp_sel_ERS0772014_DetalleCronograma '" & psncontrato & "'," & pncontref
        End If
        Set ListarCronogramaDeContratos = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    Else
        Set ListarCronogramaDeContratos = Nothing
    End If
    Exit Function
ErrorListarCronogramaDeContratos:
    Err.Raise Err.Number, "Registrar Contrato", Err.Description
End Function
Public Function ListarDatosContratos(ByVal psncontrato As String, Optional ByVal pncontref As Integer = 0) As ADODB.Recordset  'pnContRef Agregado PASIERS0772014
On Error GoTo ErrorListarDatosContratos
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        If pncontref = 0 Then
            sSql = "stp_sel_ContratosPorNumero '" & psncontrato & "'"
        Else
            sSql = "stp_sel_ERS0772014_ContratosPorNumero '" & psncontrato & "'," & pncontref 'PASIERS0772014
        End If
        Set ListarDatosContratos = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    Else
        Set ListarDatosContratos = Nothing
    End If
    Exit Function
ErrorListarDatosContratos:
    Err.Raise Err.Number, "Datos del Contrato", Err.Description
End Function

Public Function ObtenerCuotasContratos(ByVal psncontrato As String, Optional ByVal pncontref As Integer = 0) As ADODB.Recordset  'pnContRef Agregado PASI20140823 ti-ers077-2014
On Error GoTo ErrorObtenerCuotasContratos
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        If pncontref = 0 Then
            sSql = "stp_sel_ObtenerCantidadCuotasContrato '" & psncontrato & "'"
        Else
            sSql = "stp_sel_ERS0772014_ObtenerCantidadCuotasContrato '" & psncontrato & "'," & pncontref 'PASIERS0772014
        End If
        Set ObtenerCuotasContratos = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    Else
        Set ObtenerCuotasContratos = Nothing
    End If
    Exit Function
ErrorObtenerCuotasContratos:
    Err.Raise Err.Number, "Obtener Cuotas del Contrato", Err.Description
End Function
Public Function ObtenerUltCuotaContratos(ByVal psncontrato As String, Optional ByVal pncontref As Integer = 0) As Integer
On Error GoTo ErrorListarDatosContratos
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        'sSql = " select isnull(COUNT(nNPago),0) UltNPago From ContratoCronograma where cNContrato='" & psNContrato & "' and not nEstado=4 "
        If pncontref = 0 Then
            sSql = "SELECT ISNULL(MAX(CC.nNPago),0) UltNPago From ContratoCronograma CC INNER JOIN ContratoProveedor CP ON CP.cNContrato = CC.cNContrato AND CP.cNAdenda = CC.cNAdenda WHERE CP.cNContrato='" & psncontrato & "'" 'EJVG20131203
        Else
            sSql = "stp_sel_ERS0772014_UltimaCuotaContrato '" & psncontrato & "'," & pncontref  'PASIERS0772014
        End If
        
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        
        If rs.RecordCount > 0 Then
            ObtenerUltCuotaContratos = CInt(Trim(rs!UltNPago))
        Else
            ObtenerUltCuotaContratos = 0
        End If
    Else
         ObtenerUltCuotaContratos = 0
    End If
    Exit Function
ErrorListarDatosContratos:
    Err.Raise Err.Number, "Datos del Contrato", Err.Description
End Function
Public Function ObtenerUltAdendaContratos(ByVal psncontrato As String, Optional ByVal pncontref As Integer = 0) As Integer 'pnContRef Agregado PASI20140823 ti-ers077-2014
On Error GoTo ErrorListarDatosContratos
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        'sSql = " select isnull(max(cNAdenda),0) UltNAdenda from ContratoAdenda where cNContrato='" & psNContrato & "'"
        If pncontref = 0 Then
            sSql = " select isnull(max(cNAdenda),0) UltNAdenda from ContratoProveedor where cNContrato='" & psncontrato & "'" 'EJVG20131203
        Else
            sSql = "stp_sel_ERS0772014_UltimaAdendaContrato '" & psncontrato & "'," & pncontref 'PASIERS0772014
        End If
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        
        If rs.RecordCount > 0 Then
            ObtenerUltAdendaContratos = CInt(Trim(rs!UltNAdenda))
        Else
            ObtenerUltAdendaContratos = 0
        End If
    Else
         ObtenerUltAdendaContratos = 0
    End If
    Exit Function
ErrorListarDatosContratos:
    Err.Raise Err.Number, "Datos del Contrato", Err.Description
End Function

Public Function ExisteAdendaContratos(ByVal psncontrato As String, Optional ByVal pncontref As Integer = 0) As Integer 'pnContRef Agregado PASI20140823 ti-ers077-2014
On Error GoTo ErrorListarDatosContratos
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        'EJVG20131203 ***
        'sSql = " select isnull(max(cNAdenda),0) UltNAdenda from ContratoAdenda where cNContrato='" & psNContrato & "'" & _
               " and nTipo=1 and not nEstado=4 "
        If pncontref = 0 Then
            sSql = " select isnull(max(cNAdenda),0) UltNAdenda from ContratoProveedor where cNContrato='" & psncontrato & "'"
        Else
            sSql = "stp_sel_ERS0772014_ExisteAdendasContrato '" & psncontrato & "'," & pncontref 'PASIERS0772014
        End If
        'END EJVG *******

        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        
        If rs.RecordCount > 0 Then
            ExisteAdendaContratos = CInt(Trim(rs!UltNAdenda))
        Else
            ExisteAdendaContratos = 0
        End If
    Else
         ExisteAdendaContratos = 0
    End If
    Exit Function
ErrorListarDatosContratos:
    Err.Raise Err.Number, "Datos del Contrato", Err.Description
End Function
Public Function ListarDatosAdenda(ByVal psncontrato As String, ByVal pnNAdendas As Integer, Optional ByVal pncontref As Integer = 0) As ADODB.Recordset 'pnContRef Agregado PASI20140823 ti-ers077-2014
On Error GoTo ErrorListarDatosContratos
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        If pncontref = 0 Then
            sSql = "stp_sel_ContratosAdenasPorNumero '" & psncontrato & "'," & pnNAdendas
        Else
            sSql = "stp_sel_ERS0772014_ContratosAdenasPorNumero '" & psncontrato & "'," & pnNAdendas & "," & pncontref  'PASIERS0772014
        End If
    
        Set ListarDatosAdenda = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    Else
        Set ListarDatosAdenda = Nothing
    End If
    Exit Function
ErrorListarDatosContratos:
    Err.Raise Err.Number, "Datos del Contrato", Err.Description
End Function

Public Function RegistrarAdenda(ByVal pcNContrato As String, ByVal pnNAdenda As Integer, ByVal pnTipo As Integer, _
                                ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pnMoneda As Integer, ByVal pnMonto As Double, _
                                ByVal pcGlosa As String, ByVal pnEstado As Integer, Optional ByVal pnCuotaDesde As Integer = 0, _
                                Optional ByVal pnCuotaHasta As Integer = 0, Optional ByVal psNombreArchivo As String = "") As Integer 'WIOR 20130131  psNombreArchivo
On Error GoTo ErrorRegistrarAdenda
    Dim sActualiza As String
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        sSql = "stp_ins_ContratoAdenda '" & pcNContrato & "'," & pnNAdenda & "," & pnTipo & ",'" & _
                Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "'," & _
                pnMoneda & "," & pnMonto & ",'" & pcGlosa & "'," & pnEstado & "," & pnCuotaDesde & _
                "," & pnCuotaHasta & ",'" & psNombreArchivo & "'"
        
        
        oConec.Ejecutar sSql
        
        RegistrarAdenda = 0
        oConec.CierraConexion
    Else
        RegistrarAdenda = 1
    End If
    Exit Function
ErrorRegistrarAdenda:
    Err.Raise Err.Number, "Registrar Adenda", Err.Description
End Function

Public Function ListarDatosAdendasPorContrato(ByVal psncontrato As String, Optional ByVal pncontref As Integer = 0) As ADODB.Recordset 'pnContRef Agregado PASIERS0772014
On Error GoTo ErrorListarDatosAdendas
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        If pncontref = 0 Then
            sSql = "stp_sel_ListarAdendasPorContratoEstado '" & psncontrato & "'"
        Else
            sSql = "stp_sel_ERS0772014_ListarAdendasPorContratoEstado '" & psncontrato & "'," & pncontref 'PASIERS0772014
        End If
        Set ListarDatosAdendasPorContrato = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    Else
        Set ListarDatosAdendasPorContrato = Nothing
    End If
    Exit Function
ErrorListarDatosAdendas:
    Err.Raise Err.Number, "Datos de la Adenda", Err.Description
End Function

Public Function RegistrarExtornoAdenda(ByVal pcNContrato As String, ByVal pncontref As Integer, Optional ByVal pnNAdenda As Integer = 0, Optional ByVal pnTipo As Integer = 0, Optional ByVal pcGlosa As String = "", Optional pcMovExt As String = "", Optional ByVal pnTipoExt As Integer = 0) As Integer 'pncontref agregado pasi20140823 ti-ers077-2014
On Error GoTo ErrorRegistrarExtornoAdenda
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
        If pnTipoExt = 0 Then
            sSql = "insert into ExtornoAdenda(cNContrato,cNAdenda,nTipo,cGlosa,cMovExt,cnContRef) values('" & pcNContrato & "'," & pnNAdenda & "," & pnTipo & ",'" & pcGlosa & "','" & pcMovExt & "'," & pncontref & " )"
        ElseIf pnTipoExt = 1 Then
            sSql = "insert into ExtornoContrato(cNContrato,cGlosa,cMovExt,cnContref) values('" & pcNContrato & "','" & pcGlosa & "','" & pcMovExt & "'," & pncontref & ")"
        End If
        
        oConec.Ejecutar sSql
        
        RegistrarExtornoAdenda = 0
        oConec.CierraConexion
    Else
        RegistrarExtornoAdenda = 1
    End If
    Exit Function
ErrorRegistrarExtornoAdenda:
    Err.Raise Err.Number, "Registrar Extorno Adenda", Err.Description
End Function
Public Function ExtornarAdenda(ByVal pcNContrato As String, ByVal pnNAdenda As Integer, Optional ByVal pnTipo As Integer = 0) As Integer
On Error GoTo ErrorExtornarAdenda
    Dim oConec As DConecta
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        sSql = "update ContratoAdenda set nEstado=4 where cNContrato='" & pcNContrato & "' and cNAdenda=" & pnNAdenda
            oConec.Ejecutar sSql
            
        If pnTipo = 1 Then
            sSql = "update ContratoCronograma set nEstado=4 where cNContrato='" & pcNContrato & "' and cNAdenda=" & pnNAdenda
            
            oConec.Ejecutar sSql
        End If

        ExtornarAdenda = 0
        oConec.CierraConexion
    Else
        ExtornarAdenda = 1
    End If
    Exit Function
ErrorExtornarAdenda:
    Err.Raise Err.Number, "Extornar Adenda", Err.Description
End Function
Public Function ExtornarContrato(ByVal pcNContrato As String, Optional ByVal pncontref = 0) As Integer
On Error GoTo ErrorExtornarContrato
    Dim oConec As DConecta

    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        'sSql = "update ContratoProveedor set nEstado=5 where cNContrato='" & pcNContrato & "'" 'Comentado PASI20140821
        If pncontref = 0 Then
            sSql = "stp_upd_ExtornaContrato'" & pcNContrato & "'" 'PASI20140821 TI-ERS077-2014
        Else
            sSql = "stp_upd_ERS0772004_ExtornaContrato'" & pcNContrato & "'," & pncontref 'PASI20140821 TI-ERS077-2014
        End If
        
        oConec.Ejecutar sSql
        
        ExtornarContrato = 0
        oConec.CierraConexion
    Else
        ExtornarContrato = 1
    End If
    Exit Function
ErrorExtornarContrato:
    Err.Raise Err.Number, "Extornar Contrato", Err.Description
End Function
Public Function ListaCuotasAdenda(ByVal psncontrato As String, ByVal pnNAdendas As Integer, Optional ByVal pncontref As Integer = 0) As ADODB.Recordset 'pnContRef Agregado PASI20140823 ti-ers077-2014
On Error GoTo ErrorListaCuotasAdenda
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        'sSql = " select cNContrato,cNAdenda,nNPago,dFecPago, " & _
               " Case when nMoneda=1 then 'SOLES' ELSE 'DOLARES'END Moneda " & _
               " ,nMoneda,nMonto,nEstado from ContratoCronograma " & _
               " where cncontrato='" & psNContrato & "' " & _
               " and cnadenda=" & pnNAdendas & " Order By nNPago "
        sSql = "Exec stp_sel_ERS0622013_ListaCuotasAdenda '" & psncontrato & "'," & pnNAdendas & "," & pncontref  'EJVG20131205"
    
        Set ListaCuotasAdenda = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    Else
        Set ListaCuotasAdenda = Nothing
    End If
    Exit Function
ErrorListaCuotasAdenda:
    Err.Raise Err.Number, "Cuotas de la Adenda del Contrato", Err.Description
End Function
Public Function ObtenerUITActual() As Double
On Error GoTo ErrorObtenerUitActual
    Dim oConec As DConecta
    Dim rs As Recordset
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        sSql = " SELECT top 1 nUIT FROM UIT order by nPeriodo desc"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        
        If rs.RecordCount > 0 Then
            ObtenerUITActual = CDbl(Trim(rs!nUIT))
        Else
            ObtenerUITActual = 0
        End If
        
    Else
        ObtenerUITActual = 0
    End If
    Exit Function
ErrorObtenerUitActual:
    Err.Raise Err.Number, "Valor de UIT", Err.Description
End Function
'WIOR FIN ***********************************************************************
'WIOR 20130131 ******************************************************************
Public Function ListarContratosGeneral(ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal psTipo As String, ByVal pnEstado As String) As ADODB.Recordset
On Error GoTo ErrorListarContratosGeneral
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        sSql = "exec stp_sel_ContratosGeneral '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & psTipo & "','" & pnEstado & "'"
    
        Set ListarContratosGeneral = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    Else
        Set ListarContratosGeneral = Nothing
    End If
    Exit Function
ErrorListarContratosGeneral:
    Err.Raise Err.Number, "Listar Contratos General", Err.Description
End Function
Public Function ListarAdendasGeneral(ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal psTipo As String, ByVal pnEstado As String) As ADODB.Recordset
On Error GoTo ErrorListarAdendasGeneral
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        sSql = "exec stp_sel_AdendasContratosGeneral '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & psTipo & "','" & pnEstado & "'"
    
        Set ListarAdendasGeneral = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    Else
        Set ListarAdendasGeneral = Nothing
    End If
    Exit Function
ErrorListarAdendasGeneral:
    Err.Raise Err.Number, "Listar Adendas General", Err.Description
End Function
Public Function NombreArchivoAdenda(ByVal pcNContrato As String, ByVal pcNAdenda As String) As String
On Error GoTo ErrorNombreArchivoAdenda
    Dim oConec As DConecta
    Dim rs As Recordset
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        sSql = "select isnull(cNombreArchivo,'')cNombreArchivo from ContratoAdenda where cNContrato='" & pcNContrato & "' and cNAdenda='" & pcNAdenda & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        If rs.RecordCount > 0 Then
            NombreArchivoAdenda = Trim(rs!cNombreArchivo)
        Else
            NombreArchivoAdenda = ""
        End If
    Else
        NombreArchivoAdenda = ""
    End If
    Exit Function
ErrorNombreArchivoAdenda:
    Err.Raise Err.Number, "Archivo de Adenda", Err.Description
End Function
Public Function OrdenProvisionada(ByVal pnMovNro As String) As Boolean
On Error GoTo ErrorOrdenProvisionada
    Dim oConec As DConecta
    Dim rs As Recordset
    Set oConec = New DConecta
    
    OrdenProvisionada = False
    If oConec.AbreConexion() Then
            
        sSql = "exec stp_sel_OrdenCompraProvisionada  '" & pnMovNro & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        If rs.RecordCount > 0 Then
            If Trim(Trim(rs!nMovNro)) <> "" Then
                OrdenProvisionada = True
            End If
        Else
            OrdenProvisionada = False
        End If
    Else
        OrdenProvisionada = False
    End If
    Exit Function
ErrorOrdenProvisionada:
    Err.Raise Err.Number, "Orden Provisionada", Err.Description
End Function

Public Sub InsActComprobante(ByVal psNCompronbante As String, ByVal pnTipComp As Integer, _
                            ByVal pdFecEmision As Date, ByVal psCodGen As String, _
                            ByVal psCodigo As String, ByVal pnTipRel As Integer, _
                            ByVal psOpeCod As String, ByVal pnTipOpe As Integer)
                            
On Error GoTo ErrorInsActComprobante
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
    
        sSql = "exec stp_ins_InsActCompronateLog '" & psNCompronbante & "'," & pnTipComp & ",'" _
                & Format(pdFecEmision, "yyyymmdd") & "','" & psCodGen & "','" & psCodigo & "'," _
                & pnTipRel & ",'" & psOpeCod & "'," & pnTipOpe
                
        oConec.Ejecutar (sSql)
        oConec.CierraConexion
    End If
    
    Exit Sub
ErrorInsActComprobante:
    Err.Raise Err.Number, "Insertar-Actualizar Comprobante", Err.Description
End Sub
Public Function ObtenerComprobante(ByVal psCodGen As String, ByVal psCodigo As String, _
                                ByVal pnTipRel As Integer, ByVal psOpeCod As String) As ADODB.Recordset
                                
On Error GoTo ErrorObtenerComprobante
Dim oConec As DConecta
Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        sSql = "exec stp_sel_CompronateLog '" & psCodGen & "','" & psCodigo & "'," & pnTipRel & ",'" & psOpeCod & "'"
    
        Set ObtenerComprobante = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    Else
        Set ObtenerComprobante = Nothing
    End If
    Exit Function
ErrorObtenerComprobante:
    Err.Raise Err.Number, "Listar Adendas General", Err.Description
End Function
Public Function ObtenerComprobantePend() As ADODB.Recordset
On Error GoTo ErrorObtenerComprobantePend
    Dim oConec As DConecta
    Dim rs As Recordset
    
    Set oConec = New DConecta
    
    If oConec.AbreConexion() Then
            
        sSql = "exec st_sel_ObtenerComprobantesPendientes"
    
        Set ObtenerComprobantePend = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    Else
        Set ObtenerComprobantePend = Nothing
    End If
    Exit Function
ErrorObtenerComprobantePend:
    Err.Raise Err.Number, "Obtener Comprobantes Pendientes", Err.Description
End Function
'WIOR 20130131 ******************************************************************
'EJVG20131007 ***
Public Function RegistrarContrato_NEW(ByVal pcNContrato As String, ByVal pcPersCodProv As String, ByVal pcAreaUsuaria As String, _
                                        ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pnMoneda As Integer, ByVal pnMonto As Double, _
                                        ByVal pnCuotas As Integer, ByVal pnTipoPago, ByVal pnTipoContrato As Integer, ByVal pdFecFirma As Date, _
                                        ByVal pcGlosa As String, ByVal pcNombreArchivo As String, ByVal pnTipoGarantia As Integer, _
                                        ByVal pnMonedaGarantia As Integer, ByVal pnMontoGarantia As Double, ByVal pnEstado As Integer, _
                                        ByVal pnTpoContratacion As Integer, ByVal psobjeto As String) As Integer
    On Error GoTo ErrorRegistrarContrato
    Dim lsSQL As String
    RegistrarContrato_NEW = 1
    lsSQL = "stp_ins_ContratoProveedor '" & pcNContrato & "','" & pcPersCodProv & "','" & pcAreaUsuaria & "','" & _
            Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "'," & pnMoneda & "," & pnMonto & "," & _
            pnCuotas & "," & pnTipoPago & "," & pnTipoContrato & ",'" & Format(pdFecFirma, "yyyymmdd") & "','" & pcGlosa & "','" & _
            pcNombreArchivo & "'," & pnTipoGarantia & "," & pnMonedaGarantia & "," & pnMontoGarantia & "," & pnEstado & "," _
            & pnTpoContratacion & ",'" & psobjeto & "'"
    oconect.Ejecutar lsSQL
    RegistrarContrato_NEW = 0
    Exit Function
ErrorRegistrarContrato:
    Err.Raise Err.Number, "Error al Registrar Contrato", Err.Description
End Function
Public Function RegistrarCronogramaContrato_NEW(ByVal pcNContrato As String, ByVal pnnpago As Integer, ByVal pdFecPago As Date, ByVal pnMoneda As Integer, ByVal pnMonto As Double, ByVal pnEstado As Integer, ByVal pnNAdenda As Integer, ByVal nContRef As Integer) As Integer 'PASI2040822 Ti-ERS077-2014 se agrego el parametro nContRef
    On Error GoTo ErrorRegistrarCronogramaContrato
    Dim lsSQL As String
    RegistrarCronogramaContrato_NEW = 1
    lsSQL = "stp_ins_ContratoCronograma '" & pcNContrato & "'," & pnnpago & ",'" & Format(pdFecPago, "yyyymmdd") & "'," & pnMoneda & "," & pnMonto & "," & pnEstado & "," & pnNAdenda & "," & nContRef
    oconect.Ejecutar lsSQL
    RegistrarCronogramaContrato_NEW = 0
    Exit Function
ErrorRegistrarCronogramaContrato:
    Err.Raise Err.Number, "Error al Registrar Cronograma de Contrato", Err.Description
End Function
Public Function ListaTpoDocOrigenActaConformidad(ByVal psTpoDocOrigen As String) As ADODB.Recordset
On Error GoTo ErrListaTpoDocOrigenActaConformidad
    Dim oConec As New DConecta
    Dim lsSQL As String
    oConec.AbreConexion
    lsSQL = "Exec stp_sel_ERS0622013_ListaTpoDocOrigenActaConformidad '" & psTpoDocOrigen & "'"
    Set ListaTpoDocOrigenActaConformidad = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrListaTpoDocOrigenActaConformidad:
    Err.Raise Err.Number, "Error al recuperar Tipo Documento Origen Acta Conformidad", Err.Description
End Function
Public Function ListaOrdenCompraxActaConformidad(ByVal psAreaAgeCod As String, ByVal psCtaContCod As String) As ADODB.Recordset
On Error GoTo ErrListaOrdenCompraxActaConformidad
    Dim oConec As New DConecta
    Dim lsSQL As String
    oConec.AbreConexion
    lsSQL = "Exec stp_sel_ERS0622013_ListaOCxActaConformidad '" & psAreaAgeCod & "','" & psCtaContCod & "'"
    Set ListaOrdenCompraxActaConformidad = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrListaOrdenCompraxActaConformidad:
    Err.Raise Err.Number, "Error al recuperar Lista de Orden de Compra para Acta de Conformidad", Err.Description
End Function
Public Function ListaOrdenCompraDetxActaConformidad(ByVal pnMovNro As Long, ByVal psAreaAgeCod As String, ByVal psCtaContCod As String) As ADODB.Recordset
On Error GoTo ErrListaOrdenCompraDetxActaConformidad
    Dim oConec As New DConecta
    Dim lsSQL As String
    oConec.AbreConexion
    lsSQL = "Exec stp_sel_ERS0622013_ListaOCDetxActaConformidad " & pnMovNro & ",'" & psAreaAgeCod & "','" & psCtaContCod & "'"
    Set ListaOrdenCompraDetxActaConformidad = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrListaOrdenCompraDetxActaConformidad:
    Err.Raise Err.Number, "Error al recuperar Detalle de Orden de Compra para Acta de Conformidad", Err.Description
End Function
Public Function ListaOrdenServicioxActaConformidad(ByVal psAreaAgeCod As String, ByVal psCtaContCod As String) As ADODB.Recordset
On Error GoTo ErrListaOrdenServicioxActaConformidad
    Dim oConec As New DConecta
    Dim lsSQL As String
    oConec.AbreConexion
    lsSQL = "Exec stp_sel_ERS0622013_ListaOSxActaConformidad '" & psAreaAgeCod & "','" & psCtaContCod & "'"
    Set ListaOrdenServicioxActaConformidad = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrListaOrdenServicioxActaConformidad:
    Err.Raise Err.Number, "Error al recuperar Lista de Orden de Servicio para Acta de Conformidad", Err.Description
End Function
Public Function ListaOrdenServicioDetxActaConformidad(ByVal pnMovNro As Long, ByVal psAgeCod As String, ByVal psCtaContCod As String) As ADODB.Recordset
On Error GoTo ErrListaOrdenServicioDetxActaConformidad
    Dim oConec As New DConecta
    Dim lsSQL As String
    oConec.AbreConexion
    lsSQL = "Exec stp_sel_ERS0622013_ListaOSDetxActaConformidad " & pnMovNro & ",'" & psAgeCod & "','" & psCtaContCod & "'"
    Set ListaOrdenServicioDetxActaConformidad = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrListaOrdenServicioDetxActaConformidad:
    Err.Raise Err.Number, "Error al recuperar Detalle de Orden de Servicio para Acta de Conformidad", Err.Description
End Function
Public Function OCSxActaConformidad(ByVal pnMovNro As Long) As ADODB.Recordset
On Error GoTo ErrOCSxActaConformidad
    Dim oConec As New DConecta
    Dim lsSQL As String
    oConec.AbreConexion
    lsSQL = "Exec stp_sel_ERS0622013_OCSxActaConformidad " & pnMovNro
    Set OCSxActaConformidad = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrOCSxActaConformidad:
    Err.Raise Err.Number, "Error al recuperar Datos de Orden de Compra/Servicio para Acta de Conformidad", Err.Description
End Function
Public Function ActaConformidad(ByVal pnMovNro As Long) As ADODB.Recordset
On Error GoTo ErrActaConformidad
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0622013_ActaConformidad " & pnMovNro
    Set ActaConformidad = oconect.CargaRecordSet(lsSQL)
    Exit Function
ErrActaConformidad:
    Err.Raise Err.Number, "Error al recuperar Acta de Conformidad", Err.Description
End Function
Public Function ActaConformidadxImpresion(ByVal pnMovNro As Long) As ADODB.Recordset
On Error GoTo ErrActaConformidad
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0622013_ActaConformidadxImpresion " & pnMovNro
    Set ActaConformidadxImpresion = oconect.CargaRecordSet(lsSQL)
    Exit Function
ErrActaConformidad:
    Err.Raise Err.Number, "Error al recuperar Acta de Conformidad por Impresión", Err.Description
End Function
Public Function GetProveedorxActaConformidadLibre(ByVal psPersCod As String, ByVal pnMoneda As Integer) As ADODB.Recordset
On Error GoTo ErrActaConformidad
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0622013_ObtieneDatosProveedor '" & psPersCod & "'," & pnMoneda
    Set GetProveedorxActaConformidadLibre = oconect.CargaRecordSet(lsSQL)
    Exit Function
ErrActaConformidad:
    Err.Raise Err.Number, "Error al recuperar Datos de Proveedor", Err.Description
End Function
Public Function ListaActaConformidadXExtorno(ByVal psUsuario As String) As ADODB.Recordset
On Error GoTo ErrActaConformidad
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0622013_ListaActaConformidadXExtorno '" & psUsuario & "'"
    Set ListaActaConformidadXExtorno = oconect.CargaRecordSet(lsSQL)
    Exit Function
ErrActaConformidad:
    Err.Raise Err.Number, "Error al recuperar lista de acta de conformidad x extornos", Err.Description
End Function
Public Function ListaDocumentosActaConformidad() As ADODB.Recordset
On Error GoTo ErrListaDocumentosActaConformidad
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0622013_ListaDocumentosActaConformidad"
    Set ListaDocumentosActaConformidad = oconect.CargaRecordSet(lsSQL)
    Exit Function
ErrListaDocumentosActaConformidad:
    Err.Raise Err.Number, "Error al recuperar lista de Documentos de Acta de Conformidad", Err.Description
End Function
Public Function ListaActaConformidadDetxComprobante(ByVal psActaConformidadNro As String) As ADODB.Recordset
On Error GoTo ErrListaActaConformidadDetxComprobante
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0622013_ListaActaConformidadDetXComprobante '" & psActaConformidadNro & "'"
    Set ListaActaConformidadDetxComprobante = oconect.CargaRecordSet(lsSQL)
    Exit Function
ErrListaActaConformidadDetxComprobante:
    Err.Raise Err.Number, "Error al recuperar detalle de Acta de Conformidad", Err.Description
End Function
Public Function ListaComprobantesxImpresion() As ADODB.Recordset
On Error GoTo ErrListaComprobantesxImpresion
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0622013_ListaComprobantesxImpresion"
    Set ListaComprobantesxImpresion = oconect.CargaRecordSet(lsSQL)
    Exit Function
ErrListaComprobantesxImpresion:
    Err.Raise Err.Number, "Error al recuperar Lista de Comprobantes", Err.Description
End Function
Public Function ListaHistorialComprobante(ByVal psPersCod As String, ByVal pnDocTpo As Integer, ByVal psDocNro As String) As ADODB.Recordset
On Error GoTo ErrListaHistorialComprobante
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0622013_HistorialComprobante '" & psPersCod & "'," & pnDocTpo & ",'" & psDocNro & "'"
    Set ListaHistorialComprobante = oconect.CargaRecordSet(lsSQL)
    Exit Function
ErrListaHistorialComprobante:
    Err.Raise Err.Number, "Error al recuperar Historial de Comprobante", Err.Description
End Function
Public Sub RegistrarAdenda_NEW(ByVal pcNContrato As String, ByVal pncontref As Integer, ByVal pnNAdenda As Integer, ByVal pnTipo As Integer, _
                                ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pnMoneda As Integer, ByVal pnMonto As Double, _
                                ByVal pcGlosa As String, ByVal pnEstado As Integer, Optional ByVal pnCuotaDesde As Integer = 0, _
                                Optional ByVal pnCuotaHasta As Integer = 0, Optional ByVal psNombreArchivo As String = "") 'pnContRef Agregado PASI20140823 ti-ers077-2014
On Error GoTo ErrRegistrarAdenda_NEW
    Dim lsSQL As String
    lsSQL = "Exec stp_ins_ContratoAdenda '" & pcNContrato & "'," & pnNAdenda & "," & pnTipo & ",'" & _
                Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "'," & _
                pnMoneda & "," & pnMonto & ",'" & pcGlosa & "'," & pnEstado & "," & pnCuotaDesde & _
                "," & pnCuotaHasta & ",'" & psNombreArchivo & "'," & pncontref
    oconect.Ejecutar lsSQL
    Exit Sub
ErrRegistrarAdenda_NEW:
    Err.Raise Err.Number, "Registrar Adenda", Err.Description
End Sub
Public Sub ActualizarContratoProveedor(ByVal psncontrato As String, Optional ByVal pnNAdenda As Integer = -1) 'pnContRef Agregado PASI20140823 ti-ers077-2014
On Error GoTo ErrActualizarContratoProveedor
    Dim lsSQL As String
    Dim lsActualiza As String
    
    If pnNAdenda <> -1 Then
        lsActualiza = lsActualiza & ",cNAdenda = " & pnNAdenda
    End If
    If Len(lsActualiza) > 0 Then
        lsActualiza = Mid(lsActualiza, 2, Len(lsActualiza))
        lsSQL = "UPDATE CP SET " & lsActualiza & " FROM ContratoProveedor CP WHERE CP.cNContrato = '" & psncontrato & "' "
        oconect.Ejecutar lsSQL
    End If
    Exit Sub
ErrActualizarContratoProveedor:
    Err.Raise Err.Number, "Actualizar Contrato Proveedor", Err.Description
End Sub
Public Function MigrarCronogramaxAdenda(ByVal psncontrato As String, ByVal pnNAdendaNew As Integer, Optional ByVal pncontref As Integer = 0) As Integer 'pnContRef Agregado PASIERS0772014
On Error GoTo ErrMigrarCronogramaxAdenda
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0622013_MigrarCronogramaContratoxAdenda '" & psncontrato & "'," & pnNAdendaNew & "," & pncontref
    Set rs = oconect.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        MigrarCronogramaxAdenda = rs!nUltCuota
    End If
    Set rs = Nothing
    Exit Function
ErrMigrarCronogramaxAdenda:
    Err.Raise Err.Number, "Migrar Cronograma x Adenda", Err.Description
End Function
Public Function InsertaContratoAdendaRel(ByVal psncontrato As String, ByVal pnNAdenda As Integer, ByVal pnnpago As Integer, Optional ByVal pncontref As Integer = 0) As Integer 'pnContRef Agregado PASIERS0772014
On Error GoTo ErrInsertaContratoAdendaRel
    Dim lsSQL As String
    lsSQL = "Exec stp_ins_ERS0622013_ContratoAdendaRel '" & psncontrato & "'," & pnNAdenda & "," & pnnpago & "," & pncontref
    oconect.Ejecutar lsSQL
    Exit Function
ErrInsertaContratoAdendaRel:
    Err.Raise Err.Number, "Error al insertar en Relacion Adenda Relación", Err.Description
End Function
Public Function RealizaronPagosdeAdenda(ByVal psncontrato As String, ByVal pnNAdenda As Integer, Optional ByVal pncontref As Integer = 0) As Boolean 'pnContRef Agregado PASI20140823 ti-ers077-2014
    On Error GoTo ErrRealizaronPagosdeAdenda
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0622013_RealizaronPagosdeAdenda '" & psncontrato & "'," & pnNAdenda & "," & pncontref
    Set rs = oconect.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        RealizaronPagosdeAdenda = rs!bRealizaronPago
    End If
    Set rs = Nothing
    Exit Function
ErrRealizaronPagosdeAdenda:
    Err.Raise Err.Number, "Error al consultar Pagos de cuotas afectadas por adenda", Err.Description
End Function
Public Function ExtornarAdenda_NEW(ByVal psncontrato As String, ByVal pnNAdenda As Integer, Optional ByVal pncontref As Integer = 0) As Boolean 'pnContRef Agregado PASI20140823 ti-ers077-2014
    On Error GoTo ErrExtornarAdenda_NEW
    Dim lsSQL As String
    If pncontref = 0 Then
        lsSQL = "Exec stp_del_ERS0622013_ExtornaAdenda '" & psncontrato & "'," & pnNAdenda & "," & pncontref
    Else
        lsSQL = "Exec stp_del_ERS0772014_ExtornaAdenda '" & psncontrato & "'," & pnNAdenda & "," & pncontref
    End If
    oconect.Ejecutar lsSQL
    Exit Function
ErrExtornarAdenda_NEW:
    Err.Raise Err.Number, "Error al extornar adenda", Err.Description
End Function
Public Function RegistrarExtornoAdenda_NEW(ByVal psncontrato As String, ByVal pnNAdenda As Integer, Optional ByVal pnTipo As Integer = 0, Optional ByVal pcGlosa As String = "", Optional pcMovExt As String = "", Optional ByVal pncontref As Integer = 0) As Integer 'pnContRef Agregado PASI20140823 ti-ers077-2014
    On Error GoTo ErrRegistrarExtornoAdenda_NEW
    Dim lsSQL As String
    lsSQL = "insert into ExtornoAdenda(cNContrato,cNAdenda,nTipo,cGlosa,cMovExt,cnContRef) values ('" & psncontrato & "'," & pnNAdenda & "," & pnTipo & ",'" & pcGlosa & "','" & pcMovExt & "'," & pncontref & " )"
    oconect.Ejecutar lsSQL
    Exit Function
ErrRegistrarExtornoAdenda_NEW:
    Err.Raise Err.Number, "Registrar Extorno Adenda", Err.Description
End Function
Public Function ActualizaMontoCronograma(ByVal psncontrato As String, ByVal pnNAdenda As Integer, ByVal pnnpago As Integer, ByVal pnMontoAdenda As Currency, ByVal pncontref As Integer) As Integer 'pnContRef Agregado PASI20140823 ti-ers077-2014) As Integer
    On Error GoTo ErrActualizaMontoCronograma
    Dim lsSQL As String
    lsSQL = "Exec stp_upd_ERS0622013_ActualizaMontoCronograma '" & psncontrato & "'," & pnNAdenda & "," & pnnpago & "," & pnMontoAdenda & "," & pncontref
    oconect.Ejecutar lsSQL
    Exit Function
ErrActualizaMontoCronograma:
    Err.Raise Err.Number, "Error al actualizar el monto de adenda al cronograma", Err.Description
End Function
Public Function RealizoPagoContrato(ByVal psncontrato As String, Optional ByVal pncontref As Integer = 0) As Boolean 'pnContRef agregado PASIERS0772014
    On Error GoTo ErrRealizoPagoContrato
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    If pncontref = 0 Then
        lsSQL = "Exec stp_sel_ERS0622013_RealizaronPago '" & psncontrato & "'"
    Else
        lsSQL = "Exec stp_sel_ERS0772014_RealizaronPago '" & psncontrato & "'," & pncontref
    End If
    Set rs = oconect.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        RealizoPagoContrato = rs!bRealizaronPago
    End If
    Set rs = Nothing
    Exit Function
ErrRealizoPagoContrato:
    Err.Raise Err.Number, "Error al consultar Pagos de cuotas contrato", Err.Description
End Function
'Public Function ListaContratoxActaConformidad(ByVal psAreaAgeCod As String, ByVal pnTpoContratacion As Integer, ByVal pnMoneda As Moneda) As ADODB.Recordset 'Comentado PASI20140926 ERS0772014
'On Error GoTo ErrListaContratoxActaConformidad
'    Dim oConec As New DConecta
'    Dim lsSQL As String
'    oConec.AbreConexion
'    lsSQL = "Exec stp_sel_ERS0622013_ListaContratoxActaConformidad '" & psAreaAgeCod & "'," & pnTpoContratacion & "," & pnMoneda
'    Set ListaContratoxActaConformidad = oConec.CargaRecordSet(lsSQL)
'    oConec.CierraConexion
'    Set oConec = Nothing
'    Exit Function
'ErrListaContratoxActaConformidad:
'    Err.Raise Err.Number, "Error al recuperar Lista de Contratos para Acta de Conformidad", Err.Description
'End Function
Public Function ListaContratoDetxActaConformidad(ByVal psncontrato As String) As ADODB.Recordset
On Error GoTo ErrListaContratoDetxActaConformidad
    Dim oConec As New DConecta
    Dim lsSQL As String
    oConec.AbreConexion
    lsSQL = "Exec stp_sel_ERS0622013_ListaContratoDetxActaConformidad '" & psncontrato & "'"
    Set ListaContratoDetxActaConformidad = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrListaContratoDetxActaConformidad:
    Err.Raise Err.Number, "Error al recuperar detalle de Contrato para Acta de Conformidad", Err.Description
End Function
Public Function ContratoxActaConformidad(ByVal psncontrato As String) As ADODB.Recordset
On Error GoTo ErrContratoxActaConformidad
    Dim oConec As New DConecta
    Dim lsSQL As String
    oConec.AbreConexion
    lsSQL = "Exec stp_sel_ERS0622013_ContratoxActaConformidad '" & psncontrato & "'"
    Set ContratoxActaConformidad = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrContratoxActaConformidad:
    Err.Raise Err.Number, "Error al recuperar Contrato para Acta de Conformidad", Err.Description
End Function
Public Function RealizoPagoContratoxCuota(ByVal psncontrato As String, ByVal pnNAdenda As Integer, ByVal pnnpago As Integer, ByVal pncontref As Integer) As Integer 'pnContRef Agregado PASI20140823 ti-ers077-2014) As Integer
    On Error GoTo ErrRealizoPagoContratoxCuota
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0622013_RealizoPagoContrato '" & psncontrato & "'," & pnNAdenda & "," & pnnpago & "," & pncontref
    oconect.Ejecutar lsSQL
    Exit Function
ErrRealizoPagoContratoxCuota:
    Err.Raise Err.Number, "Registrar Extorno Adenda", Err.Description
End Function
'END EJVG *******

'PASI20140724 ERS0772014 ******************************************
Public Function ComboContratoSuministro(ByVal pnEstado As String, ByVal pnMoneda As Integer) As ADODB.Recordset
    On Error GoTo ERRORComboContratoSuministro
    Dim rs As Recordset
    
    Set oconect = New DConecta
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ComboContratoSuministro '" & pnEstado & "'," & pnMoneda
        Set ComboContratoSuministro = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    Else
        Set ComboContratoSuministro = Nothing
    End If
    Exit Function
ERRORComboContratoSuministro:
    Err.Raise Err.Number, "Contratos de Suministro no se Cargaron", Err.Description
End Function
Public Function RegistraContratoProveedor(ByVal psncontrato As String, _
                                            ByVal cPersCodProv As String, _
                                            ByVal cAreaUsuaria As String, _
                                            ByVal psAreaUsuaria2 As String, _
                                            ByVal pdFecIni As Date, _
                                            ByVal pdFecFin As Date, _
                                            ByVal pnMoneda As Integer, _
                                            ByVal pnMonto As Double, _
                                            ByVal pnCuotas As Integer, _
                                            ByVal pnTipoPago As Integer, _
                                            ByVal pnObjContrato As Integer, _
                                            ByVal pdFecFirma As Date, _
                                            ByVal psGlosa As String, _
                                            ByVal psNombreArchivo As String, _
                                            ByVal pnTipoGarantia As Integer, _
                                            ByVal pnMonedaGarantia As Integer, _
                                            ByVal pnMontoGarantia As Double, _
                                            ByVal pnEstado As Integer, _
                                            ByVal pnTipoContrato As Integer, _
                                            ByVal psObjetoCod As String) As Integer
    On Error GoTo ErrorRegistraContratoProveedor
    Dim rs As ADODB.Recordset
        sSql = "stp_ins_ERS0772014_ContratoProveedorNew '" & psncontrato & "'," _
                                            & "'" & cPersCodProv & "'," _
                                            & "'" & cAreaUsuaria & "'," _
                                             & "'" & psAreaUsuaria2 & "'," _
                                            & "'" & Format(pdFecIni, "yyyymmdd") & "'," _
                                            & "'" & Format(pdFecFin, "yyyymmdd") & "'," _
                                            & pnMoneda & "," _
                                            & pnMonto & "," _
                                            & pnCuotas & "," _
                                            & pnTipoPago & "," _
                                            & pnObjContrato & "," _
                                            & "'" & Format(pdFecFirma, "yyyymmdd") & "'," _
                                            & "'" & psGlosa & "'," _
                                            & "'" & psNombreArchivo & "'," _
                                            & pnTipoGarantia & "," _
                                            & pnMonedaGarantia & "," _
                                            & pnMontoGarantia & "," _
                                            & pnEstado & "," _
                                            & pnTipoContrato & "," _
                                            & "'" & psObjetoCod & "'"
        Set rs = oconect.CargaRecordSet(sSql)
        
        If Not rs.EOF And Not rs.BOF Then
             RegistraContratoProveedor = rs!nId
        End If
    Exit Function
ErrorRegistraContratoProveedor:
    Err.Raise Err.Number, "Error al Registrar Contrato", Err.Description
End Function
Public Function RegistrarItemContrato(ByVal psncontrato As String, _
                                        ByVal pncontref As Integer, _
                                        ByVal pnadenda As Integer, _
                                        ByVal psBSCod As String, _
                                        ByVal pnCant As Integer, _
                                        ByVal pnPrecUnit As Double, _
                                        ByVal psCtaCont As String) 'Exito
    On Error GoTo ErrorRegistrarItemContrato
    sSql = "stp_ins_ContratoItem '" & psncontrato & "'," _
                                    & pncontref & "," _
                                    & pnadenda & "," _
                                    & "'" & psBSCod & "'," _
                                    & pnCant & "," _
                                    & pnPrecUnit & "," _
                                    & "'" & psCtaCont & "'"
    oconect.Ejecutar sSql
    Exit Function
ErrorRegistrarItemContrato:
    Err.Raise Err.Number, "Error al Registrar el Item del Contrato", Err.Description
End Function
Public Function ComparaBSContrato(ByVal psNContRef As Integer, ByVal psObjetos As String) As Boolean
    On Error GoTo ErrorComparaItemsContratoSuministro
        Dim lsSQL As String
        Dim oConec As DConecta
        Dim rs As Recordset
        
        Set oConec = New DConecta
        If oConec.AbreConexion Then
            lsSQL = "stp_sel_ERS0772014_ComparaBSContrato " & psNContRef & ",'" & psObjetos & "'"
            Set rs = oConec.CargaRecordSet(lsSQL)
            If Not rs.EOF And Not rs.BOF Then
                If rs!Cant = 0 Then
                    ComparaBSContrato = False
                Else
                    ComparaBSContrato = True
                End If
            End If
            oConec.CierraConexion
        Else
            ComparaBSContrato = False
        End If
    Exit Function
ErrorComparaItemsContratoSuministro:
    Err.Raise Err.Number, "Error al Comparar Items del Contrato", Err.Description
End Function
Public Function ComparaSaldoBSContrato(ByVal pncontref As Integer, ByVal psobjeto As String, ByVal pnMonto As Double) As Boolean
    On Error GoTo ErrorComparaSaldoItemsContratoSuministro
    Dim lsSQL As String
    Dim oConec As DConecta
    Dim rs As Recordset
    Set oConec = New DConecta
    If oConec.AbreConexion Then
        'lsSQL = "stp_sel_ERS0772014_ComparaSaldoBSContrato " & pnContref & ",'" & psAgeDest & "','" & psobjeto & "'," & pnMonto
        lsSQL = "stp_sel_ERS0772014_ComparaSaldoBSContrato " & pncontref & ",'" & psobjeto & "'," & pnMonto
        Set rs = oConec.CargaRecordSet(lsSQL)
        If Not rs.EOF And Not rs.BOF Then
            If rs!MontoDisp = 0 Then
                ComparaSaldoBSContrato = False
            Else
                ComparaSaldoBSContrato = True
            End If
            oConec.CierraConexion
        Else
            ComparaSaldoBSContrato = False
        End If
    End If
    Exit Function
ErrorComparaSaldoItemsContratoSuministro:
    Err.Raise Err.Number, "Error al Comparar Items del Contrato", Err.Description
End Function
Public Function ComparaSaldoBSEditContrato(ByVal pncontref As Integer, ByVal psobjeto As String, ByVal pnMonto As Double, ByVal psMovNro As String) As Boolean
    On Error GoTo ErrorComparaSaldoItemsContratoSuministro
    Dim lsSQL As String
    Dim oConec As DConecta
    Dim rs As Recordset
    Set oConec = New DConecta
    If oConec.AbreConexion Then
        'lsSQL = "stp_sel_ERS0772014_ComparaSaldoBSEditContrato " & pnContref & ",'" & psAdeDest & "','" & psobjeto & "'," & pnMonto & ",'" & psMovnro & "'"
        lsSQL = "stp_sel_ERS0772014_ComparaSaldoBSEditContrato " & pncontref & ",'" & psobjeto & "'," & pnMonto & ",'" & psMovNro & "'"
        Set rs = oConec.CargaRecordSet(lsSQL)
        If Not rs.EOF And Not rs.BOF Then
            If rs!MontosDisp = 0 Then
                ComparaSaldoBSEditContrato = False
            Else
                ComparaSaldoBSEditContrato = True
            End If
            oConec.CierraConexion
        Else
            ComparaSaldoBSEditContrato = False
        End If
    End If
    Exit Function
ErrorComparaSaldoItemsContratoSuministro:
    Err.Raise Err.Number, "Error al Comparar Items del Contrato", Err.Description
End Function
Public Function ExisteMovContrato(ByVal pnMovNro As Long) As String
    Dim lsSQL As String
    Dim oConec As DConecta
    Dim rs As Recordset
    Set oConec = New DConecta
    oConec.AbreConexion
    lsSQL = "stp_sel_ERS0772014_ExisteMovContrato '" & pnMovNro & "'"
    Set rs = oConec.CargaRecordSet(lsSQL)
    If Not rs.EOF And Not rs.BOF Then
        ExisteMovContrato = rs!cNContrato
    Else
        ExisteMovContrato = ""
    End If
    oConec.CierraConexion
End Function
Public Function ExistenCuotasPagadas(ByVal psncontrato As String, Optional ByVal pncontref As Integer = 0) As Integer
    Dim lsSQL As String
    Dim oConec As DConecta
    Dim rs As Recordset
    Set oConec = New DConecta
    oConec.AbreConexion
    lsSQL = "stp_sel_ERS0772014_ExistenCuotasPagadas '" & psncontrato & "'," & pncontref
    Set rs = oConec.CargaRecordSet(lsSQL)
        ExistenCuotasPagadas = rs!Estado
    oConec.CierraConexion
End Function
Public Function CargaTipoContrato() As ADODB.Recordset
    Set oconect = New DConecta
    If oconect.AbreConexion() Then
        sSql = "stp_sel_ERS0772014_TipoContrato"
        Set CargaTipoContrato = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    End If
End Function
Public Function ObtenerContratoAdendaReajuste(ByVal psncontrato As String, _
                                                ByVal pncontref As Integer) As ADODB.Recordset
    On Error GoTo ERRORObtenerAdendaReajusteContrato
        If oconect.AbreConexion() Then
            sSql = "stp_sel_ERS0772014_ContratoAdendaReajuste '" & psncontrato & "'," _
                                                    & pncontref '& "," _
                                                    & pnAdenda
            Set ObtenerContratoAdendaReajuste = oconect.CargaRecordSet(sSql)
            oconect.CierraConexion
        Else
            Set ObtenerContratoAdendaReajuste = Nothing
        End If
        Exit Function
ERRORObtenerAdendaReajusteContrato:
    Err.Raise Err.Number, "Datos del Contrato", Err.Description
End Function
Public Function RegistraContratoAdendaReajuste(ByVal psncontrato As String, _
                                                ByVal pncontref As Integer, _
                                                ByVal pnadenda As Integer, _
                                                ByVal pnTipo As Integer, _
                                                ByVal pnMoneda As Integer, _
                                                ByVal pnMonto As Currency, _
                                                ByVal pdFecIni As Date, _
                                                ByVal pdFecFin As Date, _
                                                ByVal psNombreArchivo)
                                                
On Error GoTo ErrRegistraContratoAdendaReajuste
    sSql = "stp_ins_ERS0772014_ContratoAdendaReajuste '" & psncontrato & "'," _
                                            & pncontref & "," _
                                            & pnadenda & "," _
                                            & pnTipo & "," _
                                            & pnMoneda & "," _
                                            & pnMonto & "," _
                                            & "'" & Format(pdFecIni, "yyyymmdd") & "'," _
                                            & "'" & Format(pdFecFin, "yyyymmdd") & "'," _
                                            & "'" & psNombreArchivo & "'"
    oconect.Ejecutar sSql
    Exit Function
ErrRegistraContratoAdendaReajuste:
    Err.Raise Err.Number, "Registrar Adenda", Err.Description
End Function
Public Function MigraContratoAdendaReajuste(ByVal psncontrato As String, _
                                            ByVal pncontref As Integer, _
                                            ByVal pnadenda As Integer, _
                                            ByVal pnTipo As Integer)
On Error GoTo ErrMigraContratoAdendaReajuste
    Dim rs As ADODB.Recordset
    sSql = "stp_ins_ERS07772014_MigraContratoAdendaReajuste '" & psncontrato & "'," _
                                                & pncontref & "," _
                                                & pnadenda & "," _
                                                & pnTipo
    oconect.Ejecutar (sSql)
    Exit Function
ErrMigraContratoAdendaReajuste:
    Err.Raise Err.Number, "Migrar Adenda Reajuste", Err.Description
End Function
Public Function ActualizaMontoAdendaReajuste(ByVal psncontrato As String, _
                                            ByVal pncontref As Integer, _
                                            ByVal pnadenda As Integer, _
                                            ByVal pnMontoAdenda As Currency)
    On Error GoTo ErrActualizaMontoAdendaReajuste
        sSql = "stp_upd_ERS0772014_ActualizaMontoAdendaReajuste '" & psncontrato & "'," _
                                                    & pncontref & "," _
                                                    & pnadenda & "," _
                                                    & pnMontoAdenda
    oconect.Ejecutar (sSql)
    Exit Function
ErrActualizaMontoAdendaReajuste:
    Err.Raise Err.Number, "Error al actualizar el monto de adenda", Err.Description
End Function
Public Function ListaContratoABDetxRegistrarComprobante(ByVal psncontrato As String, _
                                            ByVal pncontref As Integer, _
                                            ByVal pnadenda As Integer)
    On Error GoTo ERRORListaContratoABDetxRegistrarComprobante
    If oconect.AbreConexion() Then
        sSql = "stp_sel_ERS0772014_ListaContratoABDetxRegistrarComprobante '" & psncontrato & "'," _
                                                                        & pncontref & "," _
                                                                        & pnadenda
        Set ListaContratoABDetxRegistrarComprobante = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    Else
        Set ListaContratoABDetxRegistrarComprobante = Nothing
    End If
    Exit Function
ERRORListaContratoABDetxRegistrarComprobante:
    Err.Raise Err.Number, "Datos del Item del Contrato", Err.Description
End Function
Public Function ListaContratoObraDetxRegistrarComprobante(ByVal psncontrato As String, _
                                                            ByVal pncontref As Integer, _
                                                            ByVal pnadenda As Integer)
    On Error GoTo ERRORListaContratoObraDetxRegistrarComprobante
    If oconect.AbreConexion() Then
        sSql = "stp_sel_ERS0772014_ListaContratoObraDetxRegistrarComprobante '" & psncontrato & "'," _
                                                                        & pncontref & "," _
                                                                        & pnadenda
        Set ListaContratoObraDetxRegistrarComprobante = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    Else
        ListaContratoObraDetxRegistrarComprobante = Nothing
    End If
    Exit Function
ERRORListaContratoObraDetxRegistrarComprobante:
    Err.Raise Err.Number, "Datos del Item del Contrato", Err.Description
End Function
'Public Function RegistrarContratoSaldo(ByVal psNContrato As String, ByVal pnContRef As Integer, ByVal pnSaldo As Currency) 'CPASI
'    On Error GoTo ERRORRegistrarContratoSaldo
'    sSql = "stp_ins_ContratoSaldo '" & psNContrato & "'," _
'                                    & pnContRef & "," _
'                                    & pnSaldo
'    oconect.Ejecutar sSql
'    Exit Function
'ERRORRegistrarContratoSaldo:
'    Err.Raise Err.Number, "Error al Registrar el Saldo del Contrato", Err.Description
'End Function
Public Function RegistrarContratoItemRel(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnadenda As Integer, ByVal pNItemEst As Integer)
    On Error GoTo ERRORRegistrarContratoItemRel
    Dim rs As ADODB.Recordset
    sSql = "stp_ins_ContratoItemRel '" & psncontrato & "'," _
                                    & pncontref & "," _
                                    & pnadenda & "," _
                                    & pNItemEst
    Set rs = oconect.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
            RegistrarContratoItemRel = rs!nId
    End If
    Exit Function
ERRORRegistrarContratoItemRel:
    Err.Raise Err.Number, "Error al Registrar el Item del Contrato", Err.Description
End Function
Public Function ListaOrdenCompraxRegistroComprobante(ByVal psAreaAgeCod As String, ByVal psCtaCont As String, ByVal pdFecIni As Date, ByVal pdFecFin As Date) As ADODB.Recordset
    On Error GoTo ERRORListaOrdenCompraxRegistroComprobante
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ListaOCxRegistroComprobante '" & psAreaAgeCod & "'," _
                                                            & "'" & psCtaCont & "'," _
                                                            & "'" & Format(pdFecIni, "yyyymmdd") & "'," _
                                                            & "'" & Format(pdFecFin, "yyyymmdd") & "'"
                                    
        Set ListaOrdenCompraxRegistroComprobante = oconect.CargaRecordSet(sSql)
    Else
        Set ListaOrdenCompraxRegistroComprobante = Nothing
    End If
    Exit Function
ERRORListaOrdenCompraxRegistroComprobante:
    Err.Raise Err.Number, "Error al Cargar las Ordenes", Err.Description
End Function
Public Function ListaOrdenServicioxRegistroComprobante(ByVal psAreaAgeCod As String, ByVal psCtaCont As String, ByVal pdFecIni As Date, ByVal pdFecFin As Date) As ADODB.Recordset
    On Error GoTo ERRORListaOrdenServicioxRegistroComprobante
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ListaOSxRegistroComprobante '" & psAreaAgeCod & "'," _
                                                            & "'" & psCtaCont & "'," _
                                                            & "'" & Format(pdFecIni, "yyyymmdd") & "'," _
                                                            & "'" & Format(pdFecFin, "yyyymmdd") & "'"
                                    
        Set ListaOrdenServicioxRegistroComprobante = oconect.CargaRecordSet(sSql)
    Else
        Set ListaOrdenServicioxRegistroComprobante = Nothing
    End If
    Exit Function
ERRORListaOrdenServicioxRegistroComprobante:
    Err.Raise Err.Number, "Error al Cargar las Ordenes", Err.Description
End Function
Public Function ListaOrdenCompraDetxRegistroComprobante(ByVal pnMovNro As Long, ByVal psAreaAgeCod As String, ByVal psCtaCont As String) As ADODB.Recordset
    On Error GoTo ERRORListaOrdenCompraDetxRegistroComprobante
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ListaOCDetxRegistroComprobante " & pnMovNro & "," _
                                                                & "'" & psAreaAgeCod & "'," _
                                                                & "'" & psCtaCont & "'"
        Set ListaOrdenCompraDetxRegistroComprobante = oconect.CargaRecordSet(sSql)
    Else
        Set ListaOrdenCompraDetxRegistroComprobante = Nothing
    End If
    Exit Function
ERRORListaOrdenCompraDetxRegistroComprobante:
    Err.Raise Err.Number, "Error al Cargar los Items", Err.Description
End Function
Public Function ListaOrdenServicioDetxRegistroComprobante(ByVal pnMovNro As Long, ByVal psAreaAgeCod As String, ByVal psCtaCont As String) As ADODB.Recordset
    On Error GoTo ERRORListaOrdenServicioDetxRegistroComprobante
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ListaOSDetxRegistroComprobante " & pnMovNro & "," _
                                                                & "'" & psAreaAgeCod & "'," _
                                                                & "'" & psCtaCont & "'"
        Set ListaOrdenServicioDetxRegistroComprobante = oconect.CargaRecordSet(sSql)
    Else
        Set ListaOrdenServicioDetxRegistroComprobante = Nothing
    End If
    Exit Function
ERRORListaOrdenServicioDetxRegistroComprobante:
    Err.Raise Err.Number, "Error al Cargar los Items", Err.Description
End Function
Public Function ListaContratoxRegistroComprobante(ByVal psAreaAgeCod As String, ByVal pnTipoContrato As Integer, ByVal pnMoneda As Integer) As ADODB.Recordset
On Error GoTo ErrorListaContratoxRegistroComprobante
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ListaContratoxRegistroComprobante '" & psAreaAgeCod & "'," _
                                                                        & pnTipoContrato & "," _
                                                                        & pnMoneda
        Set ListaContratoxRegistroComprobante = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    Else
        Set ListaContratoxRegistroComprobante = Nothing
    End If
    Exit Function
ErrorListaContratoxRegistroComprobante:
    Err.Raise Err.Number, "Listar Contrato", Err.Description
End Function
Public Function ListaContratoSADetxRegistrarComprobante(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnadenda As Integer) As ADODB.Recordset
    On Error GoTo ERRORListaContratoSADetxRegistrarComprobante
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ListaContratoSADetxRegistrarComprobante '" & psncontrato & "'," _
                                                                            & pncontref & "," _
                                                                            & pnadenda
        Set ListaContratoSADetxRegistrarComprobante = oconect.CargaRecordSet(sSql)
    Else
        Set ListaContratoSADetxRegistrarComprobante = Nothing
    End If
    Exit Function
ERRORListaContratoSADetxRegistrarComprobante:
    Err.Raise Err.Number, "Error al Cargar los Items", Err.Description
End Function
Public Function ListaOrdenDetxComprobante(ByVal pnMovNro As Long) As ADODB.Recordset
On Error GoTo ERRORListaOrdenDetxComprobante
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_OrdenDetxComprobante " & pnMovNro
        Set ListaOrdenDetxComprobante = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    Else
        Set ListaOrdenDetxComprobante = Nothing
    End If
    Exit Function
ERRORListaOrdenDetxComprobante:
    Err.Raise Err.Number, "Listar Det. Orden.", Err.Description
End Function
Public Function ListaContratoDetxComprobante(ByVal psncontrato As String, ByVal pncontref As Integer) As ADODB.Recordset
On Error GoTo ERRORListaContratoDetxComprobante
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ContratoDetxComprobante '" & psncontrato & "'," _
                                                            & pncontref
        Set ListaContratoDetxComprobante = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    Else
        Set ListaContratoDetxComprobante = Nothing
    End If
    Exit Function
ERRORListaContratoDetxComprobante:
    Err.Raise Err.Number, "Listar Det. Contrato.", Err.Description
End Function
Public Function ObtenerProveedorxRegistroComprobanteLibre(ByVal psPersCod As String, ByVal pnMoneda As Integer) As ADODB.Recordset
    On Error GoTo ErrObtenerProveedorxRegistroComprobanteLibre
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ObtieneDatosProveedor '" & psPersCod & "'," _
                                                & pnMoneda
        Set ObtenerProveedorxRegistroComprobanteLibre = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    Else
        Set ObtenerProveedorxRegistroComprobanteLibre = Nothing
    End If
    Exit Function
ErrObtenerProveedorxRegistroComprobanteLibre:
    Err.Raise Err.Number, "Error al recuperar Datos de Proveedor", Err.Description
End Function
Public Function RegistraActaPendxComprobante(ByVal pnMovNroRef As Long, ByVal psAreaUsuaria As String)
    On Error GoTo ErrorRegistraActaPendxComprobante
    If oconect.AbreConexion() Then
        sSql = "stp_ins_ERS0772014_AreaxActaConformidadxContrato " & pnMovNroRef & "," _
                                                            & "'" & psAreaUsuaria & "'"
        oconect.Ejecutar sSql
    End If
    Exit Function
ErrorRegistraActaPendxComprobante:
    Err.Raise Err.Number, "Error al registrar Actas Pendientes", Err.Description
End Function
Public Function ObtenerAreaxActaConformidadxContrato(ByVal psncontrato As String, ByVal pncontref As Integer)
    On Error GoTo ERRORObtenerAreaxActaConformidadxContrato
    If oconect.AbreConexion() Then
        sSql = "stp_sel_ERS0772014_AreaxActaConformidadxContrato '" & psncontrato & "'," _
                                                                & pncontref
        Set ObtenerAreaxActaConformidadxContrato = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    Else
        Set ObtenerAreaxActaConformidadxContrato = Nothing
    End If
    Exit Function
ERRORObtenerAreaxActaConformidadxContrato:
    Err.Raise Err.Number, "Error al cargar Datos del Contrato", Err.Description
End Function
Public Function ListaComprobantexActaConformidad(ByVal psAreaAgeCod As String, ByVal pnMoneda As Integer) As ADODB.Recordset
    On Error GoTo ERRORListaComprobantexActaConformidad
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ListaComprobantexActaConformidad '" & psAreaAgeCod & "'," _
                                                                    & pnMoneda
        Set ListaComprobantexActaConformidad = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    Else
        Set ListaComprobantexActaConformidad = Nothing
    End If
    Exit Function
ERRORListaComprobantexActaConformidad:
    Err.Raise Err.Number, "Error al cargar Datos del Comprobante", Err.Description
End Function
Public Function ListaComprobanteLibrexActaConformidad(ByVal psAreaAgeCod As String, ByVal pnMoneda As Integer) As ADODB.Recordset
    On Error GoTo ERRORListaComprobanteLibrexActaConformidad
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ListaComprobanteLibrexActaConformidad '" & psAreaAgeCod & "'," _
                                                                    & pnMoneda
        Set ListaComprobanteLibrexActaConformidad = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    Else
        Set ListaComprobanteLibrexActaConformidad = Nothing
    End If
    Exit Function
ERRORListaComprobanteLibrexActaConformidad:
    Err.Raise Err.Number, "Error al cargar Datos del Comprobante", Err.Description
End Function
Public Function ListaOrdenxActaConformidad(ByVal pnMovNro As Long) As ADODB.Recordset
    On Error GoTo ERRORListaOrdenxActaConformidad
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_OrdenxActaConformidad " & pnMovNro
        Set ListaOrdenxActaConformidad = oconect.CargaRecordSet(sSql)
    Else
        Set ListaOrdenxActaConformidad = Nothing
    End If
    Exit Function
ERRORListaOrdenxActaConformidad:
    Err.Raise Err.Number, "Error al cargar Datos de la Orden", Err.Description
End Function
Public Function ListaContratoxActaConformidad(ByVal pnMovNro As Long, ByVal psAreaAgeCod As String) As ADODB.Recordset
    On Error GoTo ERRORListaContratoxActaConformidad
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ContratoxActaConformidad " & pnMovNro & ",'" & psAreaAgeCod & "'"
        Set ListaContratoxActaConformidad = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    Else
        Set ListaContratoxActaConformidad = Nothing
    End If
    Exit Function
ERRORListaContratoxActaConformidad:
    Err.Raise Err.Number, "Error al cargar Datos del Comprobante", Err.Description
End Function
Public Function ListaOrdenLibrexActaConformidad(ByVal pnMovNro As Long) As ADODB.Recordset
    On Error GoTo ERRORListaOrdenLibrexActaConformidad
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_OrdenLibrexActaconformidad " & pnMovNro
        Set ListaOrdenLibrexActaConformidad = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    Else
        Set ListaOrdenLibrexActaConformidad = Nothing
    End If
    Exit Function
ERRORListaOrdenLibrexActaConformidad:
    Err.Raise Err.Number, "Error al cargar Datos del Comprobante", Err.Description
End Function
Public Function ActaConformidad_ERS0772014_xImpresion(ByVal pnMovNro As Long) As ADODB.Recordset
    On Error GoTo ErrActaConformidad
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ActaConformidadxImpresion " & pnMovNro
        Set ActaConformidad_ERS0772014_xImpresion = oconect.CargaRecordSet(sSql)
    Else
        Set ActaConformidad_ERS0772014_xImpresion = Nothing
    End If
    Exit Function
ErrActaConformidad:
    Err.Raise Err.Number, "Error al recuperar Acta de Conformidad por Impresión", Err.Description
End Function
Public Function ActaConformidadOrdenDet_ERS0772014_xImpresion(ByVal pnMovNro As Long, ByVal pnMoneda As Integer) As ADODB.Recordset
    On Error GoTo ErrActaConformidad
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ActaConformidadOrdenDetxImpresion " & pnMovNro & "," & pnMoneda
        Set ActaConformidadOrdenDet_ERS0772014_xImpresion = oconect.CargaRecordSet(sSql)
    Else
        Set ActaConformidadOrdenDet_ERS0772014_xImpresion = Nothing
    End If
Exit Function
ErrActaConformidad:
    Err.Raise Err.Number, "Error al recuperar Acta de Conformidad por Impresión", Err.Description
End Function
Public Function ActaConformidadContDet_ERS0772014_xImpresion(ByVal pnMovNro As Long, ByVal pnMoneda As Integer) As ADODB.Recordset
    On Error GoTo ErrActaConformidad
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ActaConformidadContDetxImpresion " & pnMovNro & "," & pnMoneda
        Set ActaConformidadContDet_ERS0772014_xImpresion = oconect.CargaRecordSet(sSql)
    Else
        Set ActaConformidadContDet_ERS0772014_xImpresion = Nothing
    End If
Exit Function
ErrActaConformidad:
    Err.Raise Err.Number, "Error al recuperar Acta de Conformidad por Impresión", Err.Description
End Function
Public Function RegistrarContratoServicio(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal psNAdenda As Integer, ByVal pnnpago As Integer, ByVal psAgeDet As String, ByVal psCtaCont As String, ByVal psDescripcion As String, ByVal pnMovItem As Integer, ByVal pnMovImporte As Currency)
     On Error GoTo ERRORRegistrarContratoServicio
        sSql = "stp_ins_ERS0772014_ContratoServicio '" & psncontrato & "'," _
                                                                            & pncontref & "," _
                                                                            & psNAdenda & "," _
                                                                            & pnnpago & "," _
                                                                            & "'" & psAgeDet & "'," _
                                                                            & "'" & psCtaCont & "'," _
                                                                            & "'" & psDescripcion & "'," _
                                                                            & pnMovItem & "," _
                                                                            & pnMovImporte
        oconect.Ejecutar sSql
        Exit Function
ERRORRegistrarContratoServicio:
 Err.Raise Err.Number, "Error al registrar los Items de Servicio", Err.Description
End Function
Public Function RegistrarContratoBienes(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnadenda As Integer, ByVal psAgeDet As String, ByVal psCtaCont As String, ByVal psDescripcion As String, ByVal pnMovItem As Integer, ByVal psBSCod As String, ByVal pnCant As Integer, ByVal pnMovImporte As Currency, ByVal psCBEst As Integer, ByVal psMovNro As String, ByVal psMovNroAnt As String) As Currency
  On Error GoTo ErrorRegistraContratoBienes
  Dim rs As ADODB.Recordset
        sSql = "stp_ins_ERS0772014_ContratoBienes '" & psncontrato & "'," _
                                                                            & pncontref & "," _
                                                                            & pnadenda & "," _
                                                                            & "'" & psAgeDet & "'," _
                                                                            & "'" & psCtaCont & "'," _
                                                                            & "'" & psDescripcion & "'," _
                                                                            & pnMovItem & "," _
                                                                            & "'" & psBSCod & "'," _
                                                                            & pnCant & "," _
                                                                            & pnMovImporte & "," _
                                                                            & psCBEst & "," _
                                                                            & "'" & psMovNro & "'," _
                                                                            & "'" & psMovNroAnt & "'"
       Set rs = oconect.CargaRecordSet(sSql)
        If Not rs.EOF Then
             RegistrarContratoBienes = rs!nImporte
        End If
        Set rs = Nothing
        Exit Function
ErrorRegistraContratoBienes:
 Err.Raise Err.Number, "Error al registrar los Items de Servicio", Err.Description
End Function
Public Function ListaBienesContrato(ByVal psncontrato As String, ByVal pncontref As Integer)
    On Error GoTo ERRORListaBienesContrato
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ObtenerBienesContrato '" & psncontrato & "'," _
                                                                                    & pncontref
        Set ListaBienesContrato = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    Else
        Set ListaBienesContrato = Nothing
    End If
    Exit Function
ERRORListaBienesContrato:
    Err.Raise Err.Number, "Error al cargar Datos del Contrato", Err.Description
End Function
Public Sub RegistraSaldoContrato(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnSaldo As Currency)
    On Error GoTo ERRORActualizaSaldoContrato
    sSql = "stp_ins_ERS0772014_ContratoSaldo '" & psncontrato & "'," _
                                                & pncontref & "," _
                                                & pnSaldo _
                                                
    oconect.Ejecutar sSql
    Exit Sub
ERRORActualizaSaldoContrato:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaSaldoContrato Method")
End Sub
'Public Function ActualizaSaldoRedContrato(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnSaldo As Currency, Optional ByVal pnAdendaRef As Integer = 0) As Currency
'    On Error GoTo ERRORActualizaSaldoContrato
'    Dim rs As ADODB.Recordset
'    sSql = "stp_ins_ERS0772014_ContratoSaldoRed '" & psncontrato & "'," _
'                                                & pncontref & "," _
'                                                & pnSaldo & "," _
'                                                & pnAdendaRef
'    Set rs = oconect.CargaRecordSet(sSql)
'     If Not rs.EOF Then
'             ActualizaSaldoRedContrato = rs!nImporte
'        End If
'        Set rs = Nothing
'    Exit Function
'ERRORActualizaSaldoContrato:
'    Call RaiseError(MyUnhandledError, "DMov:ActualizaSaldoContrato Method")
'End Function
Public Function MigrarBienesContratoxAdenda(ByVal psncontrato As String, ByVal pnNAdendaNew As Integer, ByVal pncontref As Integer) As Integer
    On Error GoTo ErrMigrarBienesContratoxAdenda
    Dim rs As New ADODB.Recordset
    sSql = "Exec stp_sel_ERS0772014_MigrarBienesContratoxAdenda '" & psncontrato & "'," & pnNAdendaNew & "," & pncontref
    Set rs = oconect.CargaRecordSet(sSql)
    If Not rs.EOF Then
        MigrarBienesContratoxAdenda = rs!UltItem
    End If
    Set rs = Nothing
    Exit Function
ErrMigrarBienesContratoxAdenda:
    Err.Raise Err.Number, "Migrar Bienes x Adenda", Err.Description
End Function
Public Function ListaServiciosContrato(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnnPagoIni As Integer, ByVal pnnPagoFin As Integer)
    On Error GoTo ERRORListaServiciosContrato
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ObtenerServiciosContrato '" & psncontrato & "'," _
                                                                                    & pncontref & "," _
                                                                                    & pnnPagoIni & "," _
                                                                                    & pnnPagoFin
        Set ListaServiciosContrato = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    Else
        Set ListaServiciosContrato = Nothing
    End If
    Exit Function
ERRORListaServiciosContrato:
    Err.Raise Err.Number, "Error al cargar Datos del Contrato", Err.Description
End Function
Public Function MigrarServicioxAdenda(ByVal psncontrato As String, ByVal pnNAdendaNew As Integer, ByVal pncontref As Integer) As Integer 'pnContRef Agregado PASI20140823 ti-ers077-2014) As Integer
On Error GoTo ErrMigrarServicioxAdenda
    Dim rs As New ADODB.Recordset
    sSql = "Exec stp_sel_ERS0772014_MigrarServiciosContratoxAdenda '" & psncontrato & "'," & pnNAdendaNew & "," & pncontref
    Set rs = oconect.CargaRecordSet(sSql)
    If Not rs.EOF Then
        MigrarServicioxAdenda = rs!nUltCuota
    End If
    Set rs = Nothing
    Exit Function
ErrMigrarServicioxAdenda:
    Err.Raise Err.Number, "Migrar Servicio x Adenda", Err.Description
End Function
Public Function ActualizaMontoServicio(ByVal psncontrato As String, ByVal pnNAdenda As Integer, ByVal pnnpago As Integer, ByVal pnMontoAdenda As Currency, ByVal pncontref As Integer, ByVal psAgeDest As String, ByVal psCtaCont As String) As Integer
    On Error GoTo ErrActualizaMontoCronograma
    Dim lsSQL As String
    lsSQL = "Exec stp_upd_ERS0772014_ActualizaMontoServicio '" & psncontrato & "'," & pnNAdenda & "," & pnnpago & "," & pnMontoAdenda & "," & pncontref & ",'" & psAgeDest & "','" & psCtaCont & "'"
    oconect.Ejecutar lsSQL
    Exit Function
ErrActualizaMontoCronograma:
    Err.Raise Err.Number, "Error al actualizar el monto de adenda al cronograma", Err.Description
End Function
Public Function ListaContratoServicioxRegistrarComprobante(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnnpago As Integer) As ADODB.Recordset
    On Error GoTo ERRORListaContratoServicioxRegistrarComprobante
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ListaContServicioxRegistroComprobante '" & psncontrato & "'," _
                                                                            & pncontref & "," _
                                                                            & pnnpago
        Set ListaContratoServicioxRegistrarComprobante = oconect.CargaRecordSet(sSql)
    Else
        Set ListaContratoServicioxRegistrarComprobante = Nothing
    End If
    Exit Function
ERRORListaContratoServicioxRegistrarComprobante:
    Err.Raise Err.Number, "Error al Cargar los Items", Err.Description
End Function
Public Function ListaComprobanteExtorno(ByVal psUsuario As String, ByVal psOpeCod As String) As ADODB.Recordset
    On Error GoTo ErrListaComprobanteExtorno
    sSql = "stp_sel_ERS0772014_ListaComprobanteXExtorno '" & psUsuario & "','" & psOpeCod & "'"
    Set ListaComprobanteExtorno = oconect.CargaRecordSet(sSql)
    Exit Function
ErrListaComprobanteExtorno:
    Err.Raise Err.Number, "Error al recuperar lista de comprobantes x extornos", Err.Description
End Function
Public Function RegistrarContratoAdendaServicioRel(ByVal psncontrato As String, _
                                                                            ByVal pncontref As Integer, _
                                                                            ByVal pnNAdenda As Integer, _
                                                                            ByVal pnnpago As Integer, _
                                                                            ByVal psAgeDest As String, _
                                                                            ByVal psCtaCont As String, _
                                                                            ByVal psDescripcion As String, _
                                                                            ByVal pnImporte As Currency)
     On Error GoTo ERRORRegistrarContratoAdendaServicioRel
    sSql = "stp_ins_ERS0772014_ContratoAdendaServiciosRel '" & psncontrato & "'," _
                                                                                    & pncontref & "," _
                                                                                    & pnNAdenda & "," _
                                                                                    & pnnpago & "," _
                                                                                    & "'" & psAgeDest & "'," _
                                                                                    & "'" & psCtaCont & "'," _
                                                                                    & "'" & psDescripcion & "'," _
                                                                                    & pnImporte
    oconect.Ejecutar sSql
    Exit Function
ERRORRegistrarContratoAdendaServicioRel:
 Err.Raise Err.Number, "Error al registrar los Items de Servicio", Err.Description
End Function
Public Function RegistrarContratoAdendaBienRel(ByVal psncontrato As String, _
                                                                            ByVal pncontref As Integer, _
                                                                            ByVal pnNAdenda As Integer, _
                                                                            ByVal psAgeDest As String, _
                                                                            ByVal psCtaCont As String, _
                                                                            ByVal psDescripcion As String, _
                                                                            ByVal psBSCod As String, _
                                                                            ByVal pnCant As Integer, _
                                                                            ByVal pnPrecio As Currency, _
                                                                            ByVal pnImporte As Currency)
     On Error GoTo ERRORRegistrarContratoAdendaBienRel
    sSql = "stp_ins_ERS0772014_ContratoAdendaBienRel '" & psncontrato & "'," _
                                                                                    & pncontref & "," _
                                                                                    & pnNAdenda & "," _
                                                                                    & "'" & psAgeDest & "'," _
                                                                                    & "'" & psCtaCont & "'," _
                                                                                    & "'" & psDescripcion & "'," _
                                                                                    & "'" & psBSCod & "'," _
                                                                                    & pnCant & "," _
                                                                                    & pnPrecio & "," _
                                                                                    & pnImporte
    oconect.Ejecutar sSql
    Exit Function
ERRORRegistrarContratoAdendaBienRel:
 Err.Raise Err.Number, "Error al registrar los Items de Bienes", Err.Description
End Function
Public Function ListaContratoServicioAdendaDet(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnNAdenda As Integer) As ADODB.Recordset
    On Error GoTo ErrorListaContServAdenda
        If oconect.AbreConexion Then
            sSql = "stp_sel_ERS0772014_ContratoServicioAdendaDet '" & psncontrato & "'," & pncontref & "," & pnNAdenda
            Set ListaContratoServicioAdendaDet = oconect.CargaRecordSet(sSql)
        Else
            Set ListaContratoServicioAdendaDet = Nothing
        End If
    Exit Function
ErrorListaContServAdenda:
    Err.Raise Err.Number, "Error al cargar los Servicios", Err.Description
End Function
Public Function ListaContratoAdendaReajusteDet(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnNAdenda As Integer)
    On Error GoTo ErrListaContratoAdendaReajusteDet
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ContratoAdendaReajusteDet '" & psncontrato & "'," & pncontref & "," & pnNAdenda
        Set ListaContratoAdendaReajusteDet = oconect.CargaRecordSet(sSql)
    Else
        Set ListaContratoAdendaReajusteDet = Nothing
    End If
    Exit Function
ErrListaContratoAdendaReajusteDet:
    Err.Raise Err.Number, "Error al cargar los Datos de la Adenda", Err.Description
End Function
Public Function ListaContratoBienAdendaDet(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnNAdenda As Integer) As ADODB.Recordset
    On Error GoTo ErrListaContratoBienAdendaDet
        If oconect.AbreConexion Then
            sSql = "stp_sel_ListaContratoBienAdendaDet '" & psncontrato & "'," & pncontref & "," & pnNAdenda
            Set ListaContratoBienAdendaDet = oconect.CargaRecordSet(sSql)
        Else
            Set ListaContratoBienAdendaDet = Nothing
        End If
        Exit Function
ErrListaContratoBienAdendaDet:
    Err.Raise Err.Number, "Error al cargar los Bienes", Err.Description
End Function
Public Sub ActualizarContratoProveedorNew(ByVal psncontrato As String, ByVal pncontref As Integer, Optional ByVal pnNAdenda As Integer = -1) 'pnContRef Agregado PASI20140823 ti-ers077-2014
On Error GoTo ErrActualizarContratoProveedor
    Dim lsSQL As String
    Dim lsActualiza As String
    
    If pnNAdenda <> -1 Then
        lsActualiza = lsActualiza & ",cNAdenda = " & pnNAdenda
    End If
    If Len(lsActualiza) > 0 Then
        lsActualiza = Mid(lsActualiza, 2, Len(lsActualiza))
        lsSQL = "UPDATE CP SET " & lsActualiza & " FROM ContratoProveedorNew CP WHERE CP.cNContrato = '" & psncontrato & "' and CP.cnContRef = " & pncontref
        oconect.Ejecutar lsSQL
    End If
    Exit Sub
ErrActualizarContratoProveedor:
    Err.Raise Err.Number, "Actualizar Contrato Proveedor", Err.Description
End Sub
Public Function RealizaronPagosdeAdendaBien(ByVal psncontrato As String, ByVal pnNAdenda As Integer, Optional ByVal pncontref As Integer = 0) As Boolean 'PASIERS0772014
    On Error GoTo ErrRealizaronPagosdeAdenda
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0772014_RealizaronPagosdeAdendaBien '" & psncontrato & "'," & pnNAdenda & "," & pncontref
    Set rs = oconect.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        RealizaronPagosdeAdendaBien = rs!bRealizaronPago
    End If
    Set rs = Nothing
    Exit Function
ErrRealizaronPagosdeAdenda:
    Err.Raise Err.Number, "Error al consultar Pagos de Bienes afectadas por adenda", Err.Description
End Function
Public Function ExistenPagosValoriza(ByVal psncontrato As String, ByVal pnNAdenda As Integer, Optional ByVal pncontref As Integer = 0) As Boolean 'PASIERS0772014
    On Error GoTo ErrExistenPagosValoriza
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0772014_ExistenPagosValoriza '" & psncontrato & "'," & pnNAdenda & "," & pncontref
    Set rs = oconect.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        ExistenPagosValoriza = IIf(rs!bSaldo = 1, True, False)
    End If
    Set rs = Nothing
    Exit Function
ErrExistenPagosValoriza:
    Err.Raise Err.Number, "Error al consultar Pagos de Obras afectadas por adenda", Err.Description
End Function
Public Function ActualizaSaldoxAdenda(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnadenda As Integer)
    On Error GoTo ErrActualizaSaldoxAdenda
        sSql = "stp_upd_ERS0772014_ActualizaSaldoxAdenda '" & psncontrato & "'," & pncontref & "," & pnadenda
        oconect.Ejecutar sSql
        Exit Function
ErrActualizaSaldoxAdenda:
    Err.Raise Err.Number, "Registrar Extorno Adenda", Err.Description
End Function
Public Function ListaTpoDocOrigenComprobante() As ADODB.Recordset
    On Error GoTo ErrListaTpoDocOrigenComprobante
        If oconect.AbreConexion Then
            sSql = "stp_sel_ERS0772014_ListaTpoDocOrigenComprobante "
            Set ListaTpoDocOrigenComprobante = oconect.CargaRecordSet(sSql)
        Else
            Set ListaTpoDocOrigenComprobante = Nothing
        End If
        Exit Function
ErrListaTpoDocOrigenComprobante:
    Err.Raise Err.Number, "Error al recuperar Tipo Documento Origen de Comprobante", Err.Description
End Function
Public Function ObtieneTipoPagoContratoServicio(ByVal psncontrato As String, ByVal pncontref As Integer) As Integer
    On Error GoTo ErrObtieneTipoPagoContratoServicio
    Dim rs As New ADODB.Recordset
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ObtieneTipoPagoContratoServicio '" & psncontrato & "'," & pncontref
        Set rs = oconect.CargaRecordSet(sSql)
        If Not rs.EOF Then
             ObtieneTipoPagoContratoServicio = rs!nTipoPago
        Else
             ObtieneTipoPagoContratoServicio = 0
        End If
    End If
    Exit Function
ErrObtieneTipoPagoContratoServicio:
    Err.Raise Err.Number, "Error al recuperar Tipo de Pago del Contrato", Err.Description
End Function
Public Function ExisteSaldoContrato(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnMonto As Currency) As Boolean
    On Error GoTo ErrExisteSaldoContrato
    Dim rs As New ADODB.Recordset
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS0772014_ExisteSaldoContrato '" & psncontrato & "'," & pncontref & "," & pnMonto
        Set rs = oconect.CargaRecordSet(sSql)
        If Not rs.EOF Then
            ExisteSaldoContrato = IIf(rs!Estado = 0, False, True)
        Else
            ExisteSaldoContrato = False
        End If
    End If
    Exit Function
ErrExisteSaldoContrato:
    Err.Raise Err.Number, "Error al recuperar el Saldo del Contrato", Err.Description
End Function
Public Function InsertaItemServicioxPagoVariable(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnadenda As Integer, ByVal pnnpago As Integer)
    On Error GoTo ErrInsertaItemServicioxPagoVariable
        sSql = "stp_ins_ERS0772014_AgregaItemServicioxPagoVariable '" & psncontrato & "'," & pncontref & "," & pnadenda & "," & pnnpago
        oconect.Ejecutar sSql
        Exit Function
ErrInsertaItemServicioxPagoVariable:
    Err.Raise Err.Number, "Error al Insertar los Items del Servicio", Err.Description
End Function
Public Function ObtieneMovxImpresionOrden(ByVal pnMovNro As Long) As String
    On Error GoTo ErrObtieneMovxImpresionOrden
    Dim rs As ADODB.Recordset
    If oconect.AbreConexion Then
        sSql = "stp_sel_ERS1372014_ObtienecMovxImpresionOrden " & pnMovNro
        Set rs = oconect.CargaRecordSet(sSql)
        If Not rs.EOF Then
            ObtieneMovxImpresionOrden = rs!cMovNro
        Else
            ObtieneMovxImpresionOrden = ""
        End If
    End If
    Exit Function
ErrObtieneMovxImpresionOrden:
    Err.Raise Err.Number, "Error al recuperar el número de movimiento-", Err.Description
End Function
Public Function ListaTpoDocOrigenComprobanteLibre() As ADODB.Recordset
    On Error GoTo ErrListaTpoDocOrigenComprobanteLibre
        If oconect.AbreConexion Then
            sSql = "stp_sel_ERS0772014_ListaTpoDocOrigenComprobanteLibre "
            Set ListaTpoDocOrigenComprobanteLibre = oconect.CargaRecordSet(sSql)
        Else
            Set ListaTpoDocOrigenComprobanteLibre = Nothing
        End If
        Exit Function
ErrListaTpoDocOrigenComprobanteLibre:
    Err.Raise Err.Number, "Error al recuperar Tipo Documento Origen de Comprobante Libre", Err.Description
End Function
Public Sub ActualizaSaldoContrato(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnSaldo As Currency, Optional ByVal pnAdendaRef As Integer = 0)
    On Error GoTo ERRORActualizaSaldoContrato
    sSql = "stp_upd_ERS0772014_ActualizaSaldoContrato '" & psncontrato & "'," _
                                                & pncontref & "," _
                                                & pnSaldo & "," _
                                                & pnAdendaRef
    oconect.Ejecutar sSql
    Exit Sub
ERRORActualizaSaldoContrato:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaSaldoContrato Method")
End Sub
'End PASI ****************************
'PASIERS0472015*******************************
Public Function ObtieneOrdenesxSeguimiento(ByVal psCtaCont As String, ByVal pdFecIni As Date, pdFecFin As Date) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_ERS0472015_ObtieneOrdenesxSeguimiento '" & psCtaCont & "','" & Format(pdFecIni, "YYYYMMdd") & "','" & Format(pdFecFin, "YYYYMMdd") & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set ObtieneOrdenesxSeguimiento = rs
End Function
Public Function ObtieneOrdenesxSeguimientoxNroOrden(ByVal psNroOrden As String, ByVal pdFecIni As Date, pdFecFin As Date) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_ERS0472015_ObtieneOrdenesxSeguimientoxNroOrden '" & psNroOrden & "','" & Format(pdFecIni, "YYYYMMdd") & "','" & Format(pdFecFin, "YYYYMMdd") & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set ObtieneOrdenesxSeguimientoxNroOrden = rs
End Function
Public Function ObtieneOrdenDet(ByVal psDocNro As String, ByVal pnDocTpo As Integer) As ADODB.Recordset
     Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_ERS0472015_ObtieneOrdenDet '" & psDocNro & "'," & pnDocTpo
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set ObtieneOrdenDet = rs
End Function
Public Function ObtieneOrdenCab(ByVal psDocNro As String, ByVal pnDocTpo As Integer) As ADODB.Recordset
     Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    If oConec.AbreConexion() Then
        sSql = "exec stp_sel_ERS0472015_ObtieneOrdenCab '" & psDocNro & "'," & pnDocTpo
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set ObtieneOrdenCab = rs
End Function
'end PASI
'PASI20150917 ERS0472015
Public Function obtieneExisteItemComprobanteOrden(ByVal pnItem As Integer, ByVal pnMovNro As Long) As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oconect = New DConecta
    If oconect.AbreConexion() Then
        sSql = "stp_sel_ERS0472015_ObtieneExisteItemComprobantexOrden " & pnItem & "," & pnMovNro
         Set rs = oconect.CargaRecordSet(sSql)
         If Not rs.EOF And Not rs.BOF Then
             obtieneExisteItemComprobanteOrden = rs!existe
        Else
             obtieneExisteItemComprobanteOrden = 0
        End If
        oconect.CierraConexion
    Else
         obtieneExisteItemComprobanteOrden = 0
    End If
End Function
Public Function ObtieneItemDetxOC(ByVal pnMovNro As Long, ByVal pnItem As Integer) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oconect = New DConecta
    If oconect.AbreConexion() Then
        sSql = "stp_sel_ERS0472015_ObtieneItemDetxOC " & pnMovNro & "," & pnItem
        Set rs = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    End If
    Set ObtieneItemDetxOC = rs
End Function
Public Function ObtieneItemDetxOS(ByVal pnMovNro As Long, ByVal pnItem As Integer) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oconect = New DConecta
    If oconect.AbreConexion() Then
        sSql = "stp_sel_ERS0472015_ObtieneItemDetxOS " & pnMovNro & "," & pnItem
        Set rs = oconect.CargaRecordSet(sSql)
        oconect.CierraConexion
    End If
    Set ObtieneItemDetxOS = rs
End Function
'Public Function ObtieneMovComprobantexOrden(ByVal pnMovNro As Long) As Long
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Set oConect = New DConecta
'    If oConect.AbreConexion() Then
'        sSQL = "stp_sel_ERS0472015_ObtieneMovComprobantexOrden " & pnMovNro
'        Set rs = oConect.CargaRecordSet(sSQL)
'    End If
'    ObtieneMovComprobantexOrden = CLng(rs!nMovNro)
'End Function
Public Function ObtieneMaxOrdPago() As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oconect = New DConecta
    If oconect.AbreConexion() Then
        sSql = "stp_sel_ERS0782015_ObtieneMaxOrdPago"
        Set rs = oconect.CargaRecordSet(sSql)
        ObtieneMaxOrdPago = rs!nMaxOrd
    End If
End Function
Public Function ObtieneExisteNOrden(ByVal nRangoIni As Long, ByVal nRangoFin As Long) As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oconect = New DConecta
    If oconect.AbreConexion() Then
        sSql = "stp_sel_ERS0782015_ObtieneExisteNOrden " & nRangoIni & "," & nRangoFin
        Set rs = oconect.CargaRecordSet(sSql)
        ObtieneExisteNOrden = rs!nExiste
    End If
End Function
'END PASI

'TORE RFC190219000001 - 27/07/2019
Public Sub EnviarMail(ByVal psHost As String, ByVal psEmailEnvia As String, ByVal psEmailDestino As String, ByVal psAsunto As String, ByVal psContenido As String, _
                      Optional ByVal psEmailCC As String = "", Optional ByVal psEmailCCO As String = "", Optional ByVal psRutaDocAdjunto As String = "")

Dim oSendMail As clsSendMail
Dim bAuthLogin      As Boolean
Dim bPopLogin       As Boolean
Dim bHtml           As Boolean
Dim MyEncodeType    As ENCODE_METHOD
Dim etPriority      As MAIL_PRIORITY
Dim bReceipt        As Boolean

    Set oSendMail = New clsSendMail
    bHtml = True

    psContenido = FormateaContenido(psContenido)

    With oSendMail
        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)

        .SMTPHost = psHost                          ' Required the fist time, optional thereafter
        .from = psEmailEnvia                        ' Required the fist time, optional thereafter
        '.FromDisplayName = Nombre Envia            ' Optional, saved after first use
        .Recipient = psEmailDestino                 ' Required, separate multiple entries with delimiter character
        '.RecipientDisplayName = Nombre Destino     ' Optional, separate multiple entries with delimiter character
        .CcRecipient = psEmailCC                    ' Optional, separate multiple entries with delimiter character
        '.CcDisplayName = Nombre CC                 ' Optional, separate multiple entries with delimiter character
        .BccRecipient = psEmailCCO                  ' Optional, separate multiple entries with delimiter character
        '.ReplyToAddress = txtFrom.Text             ' Optional, used when different than 'From' address
        .Subject = psAsunto                         ' Optional
        .Message = psContenido                      ' Optional
        .Attachment = Trim(psRutaDocAdjunto)        ' Optional, separate multiple entries with delimiter character

        .AsHTML = bHtml                              ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MyEncodeType                     ' Optional, default = MIME_ENCODE
        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised

        .Send                                       ' Required
    End With

    Set oSendMail = Nothing
    
End Sub
Private Function FormateaContenido(ByVal psCadena As String) As String

    psCadena = "<font style='font-family:Calibri Light,Arial; font-size:14px'>" & psCadena & "<p><p>" & _
               "<b>TECNOLOGIA DE INFORMACIÓN</b></font>"

    FormateaContenido = psCadena
End Function

