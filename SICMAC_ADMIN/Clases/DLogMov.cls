VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsServerAdmin As String
Dim vsServerComunes As String
Dim vsServerImagenes As Variant
Dim vsServerNegocio As String
Dim vsServerPersonas As String
Dim vsFormatoFecha As String

Dim oconect As DConecta

Public Enum TpoGrabaSelPostor
    SelPosConsulta = 1
    SelPosAbsolucion = 2
    SelPosObservacion = 3
End Enum

Public Sub RollBackTrans()
    oconect.RollBackTrans
End Sub

Public Sub CommitTrans()
    oconect.CommitTrans
End Sub

Public Sub BeginTrans()
    oconect.BeginTrans
End Sub

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

Dim oIni As ClasIni
Set oIni = New ClasIni

vsConexion = oIni.CadenaConexion
vsServerAdmin = oIni.BaseAdministracion
vsServerComunes = oIni.BaseComunes
vsServerImagenes = oIni.BaseImagenes
vsServerNegocio = oIni.BaseNegocio
vsServerPersonas = oIni.BasePersonas
 
Set oconect = New DConecta
If oconect.AbreConexion = False Then
    Call RaiseError(MyUnhandledError, "DLogMov:Initialize Method. Error en Conexion a Base de datos")
End If
End Sub

Private Sub Class_Terminate()
oconect.CierraConexion
Set oconect = Nothing
End Sub

Public Function GetnMovNro(ByVal psMovNro As String) As Long
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

sql = "Select nMovNro From Mov where cMovNro ='" & psMovNro & "'"
Set rs = oconect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetnMovNro = rs!nMovNro
End If
rs.Close
Set rs = Nothing
End Function

'Ejecuta procesos Batch
Public Function EjecutaBatch() As Integer
On Error GoTo EjecutaBatchErr
    EjecutaBatch = oconect.EjecutarBatch
    Exit Function
EjecutaBatchErr:
    Call RaiseError(MyUnhandledError, "DLogMov:EjecutaBatch Method")
End Function

'Inserta Mov
Public Function InsertaMov(ByVal psMovNro As String, ByVal psOpeCod As String, _
    ByVal psMovDesc As String, ByVal pnMovEstado As Integer) As Integer
On Error GoTo InsertaMovErr
    Dim sSql As String, sActualiza As String
    InsertaMov = 1
    
    sSql = "INSERT INTO Mov (cMovNro, cOpeCod, " & _
        " cMovDesc, nMovEstado, nMovFlag) " & _
    " VALUES ('" & psMovNro & "','" & psOpeCod & "','" & _
        psMovDesc & "'," & pnMovEstado & "," & gMovFlagVigente & ")"
    
    oconect.Ejecutar sSql
    
    InsertaMov = 0
    Exit Function

InsertaMovErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaMov Method")
End Function
'Inserta MovRef
Public Function InsertaMovRef(ByVal psMovNro As String, ByVal psMovNroRef As String) As Integer
On Error GoTo InsertaMovRefErr
    Dim sSql As String, sActualiza As String
    InsertaMovRef = 1
    
    sSql = "INSERT INTO MovRef (cMovNro, cMovNroRef) " & _
    " VALUES ('" & psMovNro & "','" & psMovNroRef & "')"
    
    oconect.AdicionaCmdBatch sSql
    
    InsertaMovRef = 0
    Exit Function

InsertaMovRefErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaMovRef Method")
End Function

'Inserta LogRequerimiento
Public Function InsertaRequeri(ByVal pnReqNro As Long, ByVal pnPeriodo As Integer, _
    ByVal pnReqTpo As Integer, ByVal psNecesi As String, ByVal psRequeri As String) As Integer
On Error GoTo InsertaRequeriErr
    Dim sSql As String, sActualiza As String
    InsertaRequeri = 1
    
    sSql = "INSERT INTO LogRequerimiento (nLogReqNro, nLogReqPeriodo, nLogReqTpo, " & _
        " cLogReqNecesidad, cLogReqRequerimiento) " & _
    " VALUES (" & pnReqNro & "," & pnPeriodo & "," & pnReqTpo & ",'" & _
        psNecesi & "','" & psRequeri & "')"
    ','" & psActualizacion & "'
    oconect.Ejecutar sSql
    
    InsertaRequeri = 0
    Exit Function

InsertaRequeriErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaRequeri Method")
End Function

'Inserta LogReqTramite
Public Function InsertaReqTramite(ByVal pnReqNro As Long, ByVal pnReqTraNro As Long, _
    ByVal psAreaCod As String, ByVal psAreaCodDes As String, ByVal psReqTraCom As String, _
    ByVal pnReqTraEst As Integer, ByVal pnReqTraMoneda As Integer, _
    ByVal psActualizacion As String) As Integer
On Error GoTo InsertaReqTramiteErr
    Dim sSql As String, sActualiza As String
    InsertaReqTramite = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogReqTramite (nLogReqNro, nLogReqTraNro, cAreaCod, " & _
        " cAreaCodDes, cLogReqTraComentario, nLogReqTraEstado, nLogReqTraMoneda, " & _
        " cUltimaActualizacion)" & _
    " VALUES (" & pnReqNro & "," & pnReqTraNro & ",'" & psAreaCod & "','" & _
        psAreaCodDes & "','" & psReqTraCom & "'," & pnReqTraEst & "," & pnReqTraMoneda & ",'" & _
        sActualiza & "')"
    
    oconect.Ejecutar sSql
    
    InsertaReqTramite = 0
    Exit Function

InsertaReqTramiteErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaReqTramite Method")
End Function

'Inserta LogReqDetalle
Public Function InsertaReqDetalle(ByVal pnReqNro As Long, ByVal pnReqTraNro As Long, _
    ByVal psBSCod As String, ByVal pnReqPrecio As Currency, _
    ByVal psActualizacion As String) As Integer
On Error GoTo InsertaReqDetalleErr
    Dim sSql As String, sActualiza As String
    
    InsertaReqDetalle = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    sSql = "INSERT INTO LogReqDetalle (nLogReqNro, nLogReqTraNro, cBSCod, " & _
        " nLogReqDetPrecio, cUltimaActualizacion)" & _
    " VALUES (" & pnReqNro & "," & pnReqTraNro & ",'" & psBSCod & "'," & _
        pnReqPrecio & ",'" & sActualiza & "') "
    
    oconect.Ejecutar sSql
    
    InsertaReqDetalle = 0
    Exit Function

InsertaReqDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaReqDetalle Method")
End Function

'Actualiza LogReqDetalle
Public Function ActualizaReqDetalle(ByVal pnReqNro As Long, ByVal pnReqTraNro As Long, _
    ByVal psBSCod As String, ByVal psPlaPreNro As String, _
    ByVal psPlaRubro As String, ByVal psPreCtaContCod As String, _
    ByVal psActualizacion As String, ByVal psPeriodo As String) As Integer
On Error GoTo ActualizaReqDetalleErr
    Dim sSql As String, sActualiza As String
    
    ActualizaReqDetalle = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    sSql = " UPDATE LogReqDetalle " _
         & " SET nPresuCod = '" & psPlaPreNro & "', cPresuRubroCod = '" & psPlaRubro & "'," _
         & " cCtaContCod = '" & psPreCtaContCod & "', cUltimaActualizacion = '" & sActualiza & "', nPresuAnio = " & psPeriodo & "" _
         & " WHERE nLogReqNro = " & pnReqNro & " AND nLogReqTraNro = " & pnReqTraNro & " " _
         & "       AND cBSCod = '" & psBSCod & "'"
    oconect.Ejecutar sSql
    
    ActualizaReqDetalle = 0
    Exit Function

ActualizaReqDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaReqDetalle Method")
End Function

'Inserta LogReqDetMes
Public Function InsertaReqDetMes(ByVal pnReqNro As Long, ByVal pnReqTraNro As Long, _
    ByVal psBSCod As String, ByVal pnReqDetMesNro As Integer, ByVal pnReqDetMesCant As Currency) As Integer
On Error GoTo InsertaReqDetMesErr
    Dim sSql As String, sActualiza As String
    InsertaReqDetMes = 1
    
    sSql = "INSERT INTO LogReqDetMes (nLogReqNro, nLogReqTraNro, cBSCod, " & _
        " nLogReqDetMesNro, nLogReqDetMesCantidad)" & _
    " VALUES (" & pnReqNro & "," & pnReqTraNro & ",'" & psBSCod & "'," & _
        pnReqDetMesNro & "," & pnReqDetMesCant & ")"
    
    oconect.Ejecutar sSql
    
    InsertaReqDetMes = 0
    Exit Function

InsertaReqDetMesErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaReqDetMes Method")
End Function

'Actualiza LogReqDetMes (solo en tipo 3 - consolidado)
'Libera los items por una seleccion cancelada o desierta
Public Function ActualizaReqDetMes(ByVal pnSelNro As Long, _
    ByVal psActualizacion As String) As Integer
On Error GoTo ActualizaReqDetMesErr
    Dim sSql As String, sActualiza As String
    
    ActualizaReqDetMes = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    sSql = "UPDATE LogReqDetMes SET  " & _
        "       nLogSelNro = Null " & _
        " FROM LogReqDetMes  " & _
        " WHERE nLogSelNro = " & pnSelNro & " "
    
    oconect.Ejecutar sSql
    
    ActualizaReqDetMes = 0
    Exit Function

ActualizaReqDetMesErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaReqDetMes Method")
End Function


'Actualiza LogReqDetMes (solo en tipo 3 - consolidado)
Public Function ActualizaReqDetMesConsol(ByVal psReqNros As String, ByVal psMeses As String, _
    ByVal psBS As String, ByVal pnSelNro As Long, _
    ByVal psActualizacion As String) As Integer
On Error GoTo ActualizaReqDetMesConsolErr
    Dim sSql As String, sActualiza As String
    
    ActualizaReqDetMesConsol = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    sSql = "UPDATE LogReqDetMes SET  " & _
        "       nLogSelNro = " & pnSelNro & " " & _
        " FROM LogReqDetMes RDM JOIN MOV M ON rdm.nLogReqNro = m.nMovNro " & _
        " WHERE m.cMovNro IN ('" & psReqNros & "') AND rdm.nLogReqDetMesNro IN (" & psMeses & ") " & _
        "       AND rdm.cBSCod IN ('" & psBS & "') "
    
    oconect.Ejecutar sSql
    
    ActualizaReqDetMesConsol = 0
    Exit Function

ActualizaReqDetMesConsolErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaReqDetMesConsol Method")
End Function

'Actualiza LogRequerimiento
Public Function ActualizaRequeri(ByVal pnReqNro As Long, ByVal pnPeriodo As Integer, _
    ByVal pnReqTpo As Integer, ByVal psNecesi As String, ByVal psRequeri As String) As Integer
On Error GoTo ActualizaRequeriErr
    Dim sSql As String, sActualiza As String
    ActualizaRequeri = 1
    
    sSql = "UPDATE LogRequerimiento SET nLogReqPeriodo = " & pnPeriodo & ", " & _
        " nLogReqTpo = " & pnReqTpo & ", cLogReqNecesidad = '" & psNecesi & "', " & _
        " cLogReqRequerimiento = '" & psRequeri & "'" & _
        " WHERE nLogReqNro = " & pnReqNro & " "
    
    oconect.Ejecutar sSql
    
    ActualizaRequeri = 0
    Exit Function

ActualizaRequeriErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaRequeri Method")
End Function
'Actualiza LogRequerimiento (Tipo Cambio)
Public Function ActualizaRequeriTipCambio(ByVal pnReqNro As Long, _
ByVal pnTipCambio As Currency) As Integer
On Error GoTo ActualizaRequeriTipCambioErr
    Dim sSql As String, sActualiza As String
    ActualizaRequeriTipCambio = 1
    
    sSql = "UPDATE LogRequerimiento SET nLogReqTipCambio = " & pnTipCambio & " " & _
        " WHERE nLogReqNro = " & pnReqNro & " "
    
    oconect.Ejecutar sSql
    
    ActualizaRequeriTipCambio = 0
    Exit Function

ActualizaRequeriTipCambioErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaRequeriTipCambio Method")
End Function


'Inserta LogReqTramite
Public Function ActualizaReqTramite(ByVal pnReqNro As Long, ByVal pnReqTraNro As Long, _
    ByVal psAreaCodDes As String, ByVal pnReqTraEst As Integer, ByVal psReqTraCom As String, _
    ByVal pnReqTraMoneda As Integer, ByVal psActualizacion As String) As Integer
On Error GoTo ActualizaReqTramiteErr
    Dim sSql As String, sActualiza As String
    ActualizaReqTramite = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    'psAreaCodDes = "" And psReqTraCom = ""
    If pnReqTraEst = 0 Then
        sSql = "UPDATE LogReqTramite SET nLogReqTraMoneda = " & pnReqTraMoneda & ", " & _
            "       cUltimaActualizacion = '" & sActualiza & "' " & _
            " WHERE nLogReqNro = " & pnReqNro & " AND nLogReqTraNro = " & pnReqTraNro & " "
    Else
        sSql = "UPDATE LogReqTramite SET cAreaCodDes = '" & psAreaCodDes & "', " & _
            "       cLogReqTraComentario = '" & psReqTraCom & "', nLogReqTraEstado = " & pnReqTraEst & ", " & _
            "       nLogReqTraMoneda = " & pnReqTraMoneda & ", cUltimaActualizacion = '" & sActualiza & "' " & _
            " WHERE nLogReqNro = " & pnReqNro & " AND nLogReqTraNro = " & pnReqTraNro & " "
    End If
        
    oconect.Ejecutar sSql
    
    ActualizaReqTramite = 0
    Exit Function

ActualizaReqTramiteErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaReqTramite Method")
End Function

'Elimina LogReqDetalle
Public Function EliminaReqDetalle(ByVal pnReqNro As Long, ByVal pnReqTraNro As Long) As Integer
On Error GoTo EliminaReqDetalleErr
    Dim sSql As String, sActualiza As String
    EliminaReqDetalle = 1
    
    If Val(pnReqTraNro) = 0 Then
        sSql = "DELETE LogReqDetalle " & _
            " WHERE nLogReqNro = " & pnReqNro & " "
    Else
        sSql = "DELETE LogReqDetalle " & _
            " WHERE nLogReqNro = " & pnReqNro & " AND nLogReqTraNro = " & pnReqTraNro & " "
    End If
    
    oconect.Ejecutar sSql
    
    EliminaReqDetalle = 0
    Exit Function

EliminaReqDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:EliminaReqDetalle Method")
End Function

'Elimina LogReqDetalle
Public Function EliminaReqDetMes(ByVal pnReqNro As Long, ByVal pnReqTraNro As Long) As Integer
On Error GoTo EliminaReqDetMesErr
    Dim sSql As String, sActualiza As String
    EliminaReqDetMes = 1
    
    If Len(Trim(pnReqTraNro)) = 0 Then
        sSql = "DELETE LogReqDetMes " & _
            " WHERE nLogReqNro = " & pnReqNro & " "
    Else
        sSql = "DELETE LogReqDetMes " & _
            " WHERE nLogReqNro = " & pnReqNro & " AND nLogReqTraNro = " & pnReqTraNro & " "
    End If
        
    oconect.Ejecutar sSql
    
    EliminaReqDetMes = 0
    Exit Function

EliminaReqDetMesErr:
    Call RaiseError(MyUnhandledError, "DLogMov:EliminaReqDetMes Method")
End Function

'Inserta LogReqCon
Public Function InsertaReqCon(ByVal pnReqNro As Long, ByVal pnReqConNro As Long) As Integer
On Error GoTo InsertaReqConErr
    Dim sSql As String, sActualiza As String
    InsertaReqCon = 1
    
    sSql = "INSERT INTO LogReqCon (nLogReqNro, nLogReqConNro) " & _
        " VALUES (" & pnReqNro & "," & pnReqConNro & ")"
    
    oconect.Ejecutar sSql
    
    InsertaReqCon = 0
    Exit Function

InsertaReqConErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaReqCon Method")
End Function

'Inserta LogSeleccion
Public Function InsertaSeleccion(ByVal pnSelNro As Long, ByVal pdResolu As Date, _
ByVal psResNro As String, ByVal psAreaCod As String, ByVal pnSelTpo As Integer, _
ByVal psActualizacion As String) As Integer
On Error GoTo InsertaSeleccionErr
    Dim sSql As String, sActualiza As String
    InsertaSeleccion = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogSeleccion (nLogSelNro, dLogSelRes, cLogSelResNro, " & _
        " nLogSelEstado, cAreaCod, nLogSelTpo, cUltimaActualizacion) " & _
        " VALUES (" & pnSelNro & ",'" & Format(pdResolu, gsFormatoFecha) & "','" & psResNro & "','" & _
        gLogSelEstadoInicioRes & "','" & psAreaCod & "'," & pnSelTpo & ",'" & sActualiza & "')"
    
    oconect.Ejecutar sSql
    
    InsertaSeleccion = 0
    Exit Function

InsertaSeleccionErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSeleccion Method")
End Function

'Inserta LogReqDetalle
Public Function InsertaSelDetalle(ByVal pnSelNro As Long, _
    ByVal psBSCod As String, ByVal pnSelPrecio As Currency, _
    ByVal pnSelCantidad As Currency, ByVal psActualizacion As String) As Integer
On Error GoTo InsertaSelDetalleErr
    Dim sSql As String, sActualiza As String
    
    InsertaSelDetalle = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    sSql = "INSERT INTO LogSelDetalle (nLogSelNro, cBSCod, " & _
        " nLogSelDetCantidad, nLogSelDetPrecio, cUltimaActualizacion)" & _
    " VALUES (" & pnSelNro & ",'" & psBSCod & "'," & _
        pnSelCantidad & "," & pnSelPrecio & ",'" & sActualiza & "') "
    
    oconect.Ejecutar sSql
    
    InsertaSelDetalle = 0
    Exit Function

InsertaSelDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelDetalle Method")
End Function

'Inserta LogSelComite
Public Function InsertaSelComite(ByVal pnSelNro As Long, ByVal psAreaCod As String, _
ByVal psPersCod As String, ByVal psActualizacion As String) As Integer
On Error GoTo InsertaSelComiteErr
    Dim sSql As String, sActualiza As String
    InsertaSelComite = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogSelComite (nLogSelNro, cAreaCod, cPersCod, cUltimaActualizacion) " & _
        " VALUES (" & pnSelNro & ",'" & psAreaCod & "','" & psPersCod & "','" & sActualiza & "')"
    
    oconect.Ejecutar sSql
    
    InsertaSelComite = 0
    Exit Function

InsertaSelComiteErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelComite Method")
End Function

'Inserta LogSelCronograma
Public Function InsertaSelCronograma(ByVal pnSelNro As Long, ByVal pnCroNro As Integer, _
ByVal pdInicial As Date, ByVal pdFinal As Date, ByVal psActualizacion As String) As Integer
On Error GoTo InsertaSelCronogramaErr
    Dim sSql As String, sActualiza As String
    InsertaSelCronograma = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogSelCronograma(nLogSelNro, nLogSelCroNro, " & _
        " dLogSelCroIni, dLogSelCroFin, cUltimaActualizacion) " & _
        " VALUES (" & pnSelNro & ",'" & pnCroNro & "','" & _
        Format(pdInicial, gsFormatoFecha) & "','" & Format(pdFinal, gsFormatoFecha) & "','" & sActualiza & "')"
    
    oconect.Ejecutar sSql
    
    InsertaSelCronograma = 0
    Exit Function

InsertaSelCronogramaErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelCronograma Method")
End Function

'Actualiza LogSeleccion
Public Function ActualizaSeleccion(ByVal pnSelNro As Long, ByVal pdResolu As Date, _
ByVal psResNro As String, ByVal psAreaCod As String, ByVal pnSelTpo As Integer, _
ByVal psActualizacion As String) As Integer
On Error GoTo ActualizaSeleccionErr
    Dim sSql As String, sActualiza As String
    ActualizaSeleccion = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "UPDATE LogSeleccion SET dLogSelRes = '" & Format(pdResolu, gsFormatoFecha) & "', " & _
        "       cLogSelResNro = '" & psResNro & "', cAreaCod = '" & psAreaCod & "'," & _
        "       nLogSelTpo = " & pnSelTpo & ", cUltimaActualizacion = '" & sActualiza & "' " & _
        " WHERE nLogSelNro = " & pnSelNro & " "
        
    oconect.Ejecutar sSql
    
    ActualizaSeleccion = 0
    Exit Function

ActualizaSeleccionErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaSeleccion Method")
End Function

'Actualiza LogSeleccion - LogSelEstado
Public Function ActualizaSeleccionEstado(ByVal pnSelNro As Long, _
ByVal pnEstado As Integer, ByVal psActualizacion As String) As Integer
On Error GoTo ActualizaSeleccionEstadoErr
    Dim sSql As String, sActualiza As String
    ActualizaSeleccionEstado = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "UPDATE LogSeleccion SET nLogSelEstado = " & pnEstado & "," & _
        "       cUltimaActualizacion = '" & sActualiza & "'" & _
        " WHERE nLogSelNro = " & pnSelNro & " "
        
    oconect.Ejecutar sSql
    
    ActualizaSeleccionEstado = 0
    Exit Function

ActualizaSeleccionEstadoErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaSeleccionEstado Method")
End Function

'Actualiza LogSeleccion (Detalle)
Public Function ActualizaSeleccionDetalle(ByVal pnSelNro As Long, _
ByVal pnMoneda As Integer, ByVal pnTipCambio As Currency, _
ByVal psActualizacion As String) As Integer
On Error GoTo ActualizaSeleccionDetalleErr
    Dim sSql As String, sActualiza As String
    ActualizaSeleccionDetalle = 1

    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    sSql = "UPDATE LogSeleccion SET nLogSelMoneda = " & pnMoneda & ", " & _
        "       nLogSelTipCambio = " & pnTipCambio & ", nLogSelEstado = " & gLogSelEstadoBases & "," & _
        "       cUltimaActualizacion = '" & sActualiza & "'" & _
        " WHERE nLogSelNro = " & pnSelNro & " "
    oconect.Ejecutar sSql

    ActualizaSeleccionDetalle = 0
    Exit Function

ActualizaSeleccionDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaSeleccionDetalle Method")
End Function

'Actualiza LogSeleccion (Referencia)
Public Function ActualizaSeleccionReferencia(ByVal pnSelNro As Long, _
ByVal pnCostoBase As Currency, ByVal pnValReferencia As Currency, _
ByVal pnSisAdjudica As Integer, _
ByVal psActualizacion As String) As Integer
On Error GoTo ActualizaSeleccionReferenciaErr
    Dim sSql As String, sActualiza As String
    ActualizaSeleccionReferencia = 1

    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    sSql = "UPDATE LogSeleccion SET nLogSelCostoBase = " & pnCostoBase & ", " & _
        "       nLogSelValorRefe = " & pnValReferencia & ", nLogSelEstado = " & gLogSelEstadoParametro & "," & _
        "       nLogSelSisAdj = " & pnSisAdjudica & ", cUltimaActualizacion = '" & sActualiza & "'" & _
        " WHERE nLogSelNro = " & pnSelNro & " "
    oconect.Ejecutar sSql

    ActualizaSeleccionReferencia = 0
    Exit Function

ActualizaSeleccionReferenciaErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaSeleccionReferencia Method")
End Function

'Actualiza LogSeleccion (Adju)
Public Function ActualizaSeleccionAdju(ByVal pnSelNro As Long, ByVal pnSelCotNro As Integer, _
ByVal psActualizacion As String) As Integer
On Error GoTo ActualizaSeleccionAdjuErr
    Dim sSql As String, sActualiza As String
    ActualizaSeleccionAdju = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    sSql = "UPDATE LogSeleccion SET nLogSelCotNro = " & pnSelCotNro & ", " & _
        "       cUltimaActualizacion = '" & sActualiza & "', nLogSelEstado = " & gLogSelEstadoProcAdju & " " & _
        " WHERE nLogSelNro = " & pnSelNro & " "
    
    oconect.Ejecutar sSql
    
    ActualizaSeleccionAdju = 0
    Exit Function

ActualizaSeleccionAdjuErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaSeleccionAdju Method")
End Function

'Inserta LogSelParametro
Public Function InsertaSelParametro(ByVal pnSelNro As Long, ByVal pnSelParTpo As Integer, _
ByVal pnSelParNro As Integer, ByVal pnSelParPuntaje As Integer, ByVal pnSelParNroTpo As Integer, _
ByVal psActualizacion As String) As Integer
On Error GoTo InsertaSelParametroErr
    Dim sSql As String, sActualiza As String
    InsertaSelParametro = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogSelParametro (nLogSelNro, nLogSelParTpo, nLogSelParNro, " & _
        "       nLogSelParPuntaje, nLogSelParNroTpo, cUltimaActualizacion) " & _
        " VALUES (" & pnSelNro & "," & pnSelParTpo & "," & pnSelParNro & ", " & _
        "" & pnSelParPuntaje & "," & pnSelParNroTpo & ",'" & sActualiza & "')"
    
    oconect.Ejecutar sSql
    
    InsertaSelParametro = 0
    Exit Function

InsertaSelParametroErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelParametro Method")
End Function
'Inserta LogSelParDetalle
Public Function InsertaSelParDetalle(ByVal pnSelNro As Long, ByVal pnSelParTpo As Integer, _
ByVal pnSelParNro As Integer, ByVal pnSelParDetNro As Integer, ByVal pnSelParDetPuntaje As Integer, _
ByVal pnSelParDetIni As Long, ByVal pnSelParDetFin As Long) As Integer
On Error GoTo InsertaSelParDetalleErr
    Dim sSql As String, sActualiza As String
    InsertaSelParDetalle = 1
    
    sSql = "INSERT INTO LogSelParDetalle (nLogSelNro, nLogSelParTpo, nLogSelParNro, " & _
        "       nLogSelParDetNro, nLogSelParDetPuntaje, " & _
        "       nLogSelParDetIni, nLogSelParDetFin ) " & _
        " VALUES (" & pnSelNro & "," & pnSelParTpo & "," & pnSelParNro & ", " & _
        "" & pnSelParDetNro & "," & pnSelParDetPuntaje & "," & _
        "" & pnSelParDetIni & "," & pnSelParDetFin & ")"
    
    oconect.Ejecutar sSql
    
    InsertaSelParDetalle = 0
    Exit Function

InsertaSelParDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelParDetalle Method")
End Function
'Inserta LogSelCotPar
Public Function InsertaSelCotPar(ByVal pnSelNro As Long, ByVal pnSelCotNro As Integer, _
ByVal pnSelParTpo As Integer, ByVal pnSelParNro As Integer, ByVal pnSelCotParValor As Integer, _
ByVal psActualizacion As String) As Integer
On Error GoTo InsertaSelCotParErr
    Dim sSql As String, sActualiza As String
    InsertaSelCotPar = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogSelCotPar (nLogSelNro, nLogSelCotNro, nLogSelParTpo, " & _
        "       nLogSelParNro, nLogSelCotParValor, cUltimaActualizacion) " & _
        " VALUES (" & pnSelNro & "," & pnSelCotNro & "," & pnSelParTpo & "," & _
        pnSelParNro & ", " & pnSelCotParValor & ",'" & sActualiza & "')"
    
    oconect.Ejecutar sSql
    
    InsertaSelCotPar = 0
    Exit Function

InsertaSelCotParErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelCotPar Method")
End Function

'Inserta LogSelPublica
Public Function InsertaSelPublica(ByVal pnSelNro As Long, ByVal psPersCod As String, _
ByVal pdPubIni As Date, ByVal pdPubFin As Date, ByVal psActualizacion As String) As Integer
On Error GoTo InsertaSelPublicaErr
    Dim sSql As String, sActualiza As String
    InsertaSelPublica = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

    sSql = "INSERT INTO LogSelPublica (nLogSelNro, cPersCod, dLogSelPubIni, " & _
        "       dLogSelPubFin, cUltimaActualizacion) " & _
        " VALUES (" & pnSelNro & ",'" & psPersCod & "','" & Format(pdPubIni, gsFormatoFecha) & "', " & _
        "'" & Format(pdPubFin, gsFormatoFecha) & "','" & sActualiza & "')"
    
    oconect.Ejecutar sSql
    
    InsertaSelPublica = 0
    Exit Function

InsertaSelPublicaErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelPublica Method")
End Function

'Inserta LogSelPostor
Public Function InsertaSelPostor(ByVal pnSelNro As Long, ByVal psPersCod As String, _
ByVal pdPostor As String, ByVal psActualizacion As String) As Integer
On Error GoTo InsertaSelPostorErr
    Dim sSql As String, sActualiza As String
    InsertaSelPostor = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

    sSql = "INSERT INTO LogSelPostor (nLogSelNro, cPersCod, dLogSelPosEntrega, " & _
        "       cUltimaActualizacion) " & _
        " VALUES (" & pnSelNro & ",'" & psPersCod & "','" & Format(pdPostor, gsFormatoFecha) & "','" & _
        sActualiza & "')"
    
    oconect.Ejecutar sSql
    
    InsertaSelPostor = 0
    Exit Function

InsertaSelPostorErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelPostor Method")
End Function

'Elimina LogSelPostor
Public Function EliminaSelPostor(ByVal pnSelNro As Long) As Integer
On Error GoTo EliminaSelPostorErr
    Dim sSql As String, sActualiza As String
    EliminaSelPostor = 1
    
    sSql = "DELETE LogSelPostor " & _
        " WHERE nLogSelNro = " & pnSelNro & " "
    
    oconect.Ejecutar sSql
    
    EliminaSelPostor = 0
    Exit Function

EliminaSelPostorErr:
    Call RaiseError(MyUnhandledError, "DLogMov:EliminaSelPostor Method")
End Function


'Actualiza LogSelPostor
Public Function ActualizaSelPostor(ByVal pnTpoGraba As TpoGrabaSelPostor, _
ByVal pnSelNro As Long, ByVal psPersCod As String, ByVal psTexto As String, _
ByVal psActualizacion As String) As Integer
On Error GoTo ActualizaSelPostorErr
    Dim sSql As String, sActualiza As String
    ActualizaSelPostor = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    If pnTpoGraba = SelPosConsulta Then
        sSql = "UPDATE LogSelPostor SET cLogSelPosConsulta = '" & psTexto & "', " & _
            "       cUltimaActualizacion = '" & sActualiza & "'" & _
            " WHERE nLogSelNro = " & pnSelNro & " AND cPersCod = '" & psPersCod & "'"
    ElseIf pnTpoGraba = SelPosAbsolucion Then
        sSql = "UPDATE LogSelPostor SET cLogSelPosAbsolucion = '" & psTexto & "', " & _
            "       cUltimaActualizacion = '" & sActualiza & "'" & _
            " WHERE nLogSelNro = " & pnSelNro & " AND cPersCod = '" & psPersCod & "'"
    ElseIf pnTpoGraba = SelPosObservacion Then
        sSql = "UPDATE LogSelPostor SET cLogSelPosObservacion = '" & psTexto & "', " & _
            "       cUltimaActualizacion = '" & sActualiza & "'" & _
            " WHERE nLogSelNro = " & pnSelNro & " AND cPersCod = '" & psPersCod & "'"
    End If
    oconect.Ejecutar sSql
    
    ActualizaSelPostor = 0
    Exit Function

ActualizaSelPostorErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaSelPostor Method")
End Function

'Inserta LogSelCotiza
Public Function InsertaSelCotiza(ByVal pnSelNro As Long, ByVal pnSelCotNro As Integer, _
ByVal psPersCod As String, ByVal psActualizacion As String) As Integer
On Error GoTo InsertaSelCotizaErr
    Dim sSql As String, sActualiza As String
    InsertaSelCotiza = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

    sSql = "INSERT INTO LogSelCotiza (nLogSelNro, nLogSelCotNro, " & _
        "       cPersCod, cUltimaActualizacion) " & _
        " VALUES (" & pnSelNro & "," & pnSelCotNro & ",'" & _
                psPersCod & "','" & sActualiza & "')"
    
    oconect.Ejecutar sSql
    
    InsertaSelCotiza = 0
    Exit Function

InsertaSelCotizaErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelCotiza Method")
End Function

'Inserta LogSelCotDetalle
Public Function InsertaSelCotDetalle(ByVal pnSelNro As Long, ByVal pnSelCotNro As Integer, _
ByVal psBSCod As String, ByVal pnCantidad As Currency, ByVal pnPrecio As Currency, _
ByVal psActualizacion As String) As Integer
On Error GoTo InsertaSelCotDetalleErr
    Dim sSql As String, sActualiza As String
    InsertaSelCotDetalle = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

    sSql = "INSERT INTO LogSelCotDetalle (nLogSelNro, nLogSelCotNro, cBSCod, " & _
        "       nLogSelCotDetCantidad, nLogSelCotDetPrecio, cUltimaActualizacion) " & _
        " VALUES (" & pnSelNro & "," & pnSelCotNro & ",'" & psBSCod & "'," & _
                pnCantidad & "," & pnPrecio & ",'" & sActualiza & "')"
    
    oconect.Ejecutar sSql
    
    InsertaSelCotDetalle = 0
    Exit Function

InsertaSelCotDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelCotDetalle Method")
End Function

'Actualiza LogSelCotizacion
Public Function ActualizaSelCotizacion(ByVal pnSelNro As Long, ByVal pnSelCotNro As Integer, _
ByVal pnGarTpo As Integer, ByVal psGarNro As String, ByVal psGarComenta As String, _
ByVal psActualizacion As String) As Integer
On Error GoTo ActualizaSelCotizacionErr
    Dim sSql As String, sActualiza As String
    ActualizaSelCotizacion = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

    sSql = "UPDATE LogSelCotiza SET nLogSelCotGarTpo = " & pnGarTpo & ",  " & _
        "       cLogSelCotGarNro = '" & psGarNro & "', cLogSelCotGarComentario = '" & psGarComenta & "', " & _
        "       cUltimaActualizacion = '" & sActualiza & "' " & _
        " WHERE nLogSelNro = " & pnSelNro & " AND nLogSelCotNro = " & pnSelCotNro & " "
    
    oconect.Ejecutar sSql
    
    ActualizaSelCotizacion = 0
    Exit Function

ActualizaSelCotizacionErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaSelCotizacion Method")
End Function

'Actualiza LogSelCotDetalle
Public Function ActualizaSelCotDetalle(ByVal pnSelNro As Long, ByVal pnSelCotNro As Integer, _
ByVal psBSCod As String, ByVal pnCantidad As Currency, ByVal pnPrecio As Currency, _
ByVal psActualizacion As String) As Integer
On Error GoTo ActualizaSelCotDetalleErr
    Dim sSql As String, sActualiza As String
    ActualizaSelCotDetalle = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

    sSql = "UPDATE LogSelCotDetalle SET nLogSelCotDetCantidad = " & pnCantidad & ",  " & _
        "       nLogSelCotDetPrecio = " & pnPrecio & ", cUltimaActualizacion = '" & sActualiza & "' " & _
        " WHERE nLogSelNro = " & pnSelNro & " AND nLogSelCotNro = " & pnSelCotNro & "" & _
        "       AND cBSCod = '" & psBSCod & "'"
    
    oconect.Ejecutar sSql
    
    ActualizaSelCotDetalle = 0
    Exit Function

ActualizaSelCotDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaSelCotDetalle Method")
End Function


'Actualiza LogSelParCot
Public Function ActualizaSelCotPar(ByVal pnSelNro As Long, ByVal pnSelCotNro As Integer, _
ByVal pnSelParTpo As Integer, ByVal pnSelParNro As Integer, ByVal pnValor As Integer, _
ByVal psActualizacion As String) As Integer
On Error GoTo ActualizaSelCotParErr
    Dim sSql As String, sActualiza As String
    ActualizaSelCotPar = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

    sSql = "UPDATE LogSelCotPar SET nLogSelCotParValor = " & pnValor & ",  " & _
        "       cUltimaActualizacion = '" & sActualiza & "' " & _
        " WHERE nLogSelNro = " & pnSelNro & " AND nLogSelCotNro = " & pnSelCotNro & "" & _
        "       AND nLogSelParTpo = " & pnSelParTpo & " AND nLogSelParNro = " & pnSelParNro & " "
    
    oconect.Ejecutar sSql
    
    ActualizaSelCotPar = 0
    Exit Function

ActualizaSelCotParErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaSelCotPar Method")
End Function

'Inserta LogContratacion
Public Function InsertaContratacion(ByVal pnConNro As Long, ByVal pnSelNro As Long, _
ByVal pdContratacion As Date, ByVal psAgencia As String, ByVal psEntrega As String, _
ByVal psActualizacion As String) As Integer
On Error GoTo InsertaContratacionErr
    Dim sSql As String, sActualiza As String
    InsertaContratacion = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogContratacion (nLogConNro, nLogSelNro, dLogConReg, " & _
        " nLogConEstado, cAgeCod, cLogConEntrega, cUltimaActualizacion) " & _
        " VALUES (" & pnConNro & "," & pnSelNro & ",'" & Format(pdContratacion, gsFormatoFecha) & "'," & _
        gLogConEstadoInicio & ",'" & psAgencia & "','" & psEntrega & "','" & sActualiza & "')"
    
    oconect.Ejecutar sSql
    
    InsertaContratacion = 0
    Exit Function

InsertaContratacionErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaContratacion Method")
End Function

'Inserta LogConDetalle
Public Function InsertaConDetalle(ByVal pnConNro As Long, ByVal psBSCod As String, _
ByVal pnCantidad As Currency, ByVal pnPrecio As Currency, _
ByVal psActualizacion As String) As Integer
On Error GoTo InsertaConDetalleErr
    Dim sSql As String, sActualiza As String
    InsertaConDetalle = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogConDetalle (nLogConNro, cBsCod, nLogConDetCantidad, " & _
        " nLogConDetPrecio, cUltimaActualizacion) " & _
        " VALUES (" & pnConNro & ",'" & psBSCod & "'," & pnCantidad & "," & _
        pnPrecio & ",'" & sActualiza & "')"
    
    oconect.Ejecutar sSql
    
    InsertaConDetalle = 0
    Exit Function

InsertaConDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaConDetalle Method")
End Function


'Inserta LogSerCon
Public Function InsertaSerCon(ByVal pnSerConNro As Long, ByVal pnSerConTpo As Integer, _
ByVal psNumero As String, ByVal psDescripcion As String, ByVal pnMonto As Currency, _
ByVal pdInicio As Date, ByVal pdFinal As Date, _
ByVal psActualizacion As String) As Integer
On Error GoTo InsertaSerConErr
    Dim sSql As String, sActualiza As String
    InsertaSerCon = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

    sSql = "INSERT INTO LogSerCon (nLogSerConNro, nLogSerConTpo, " & _
        "       cPersCod, cLogSerConDescripcion, nLogSerConMonto, " & _
        "       dLogSerConInicio, dLogSerConFinal, cUltimaActualizacion) " & _
        " VALUES (" & pnSerConNro & "," & pnSerConTpo & ",'" & _
                psNumero & "','" & psDescripcion & "'," & pnMonto & ",'" & _
                Format(pdInicio, gsFormatoFecha) & "','" & Format(pdFinal, gsFormatoFecha) & "','" & _
                sActualiza & "')"
    oconect.Ejecutar sSql
    
    InsertaSerCon = 0
    Exit Function

InsertaSerConErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSerCon Method")
End Function

'Elimina un Servicios-Contratos
Public Function EliminaSerCon(ByVal pnMovNro As String) As Integer
On Error GoTo EliminaSerConErr
    Dim sSql As String
    
    EliminaSerCon = 1
    
    sSql = "UPDATE Mov SET nMovFlag = " & gMovFlagEliminado & " " & _
        " WHERE nMovNro = " & pnMovNro & " "
    
    oconect.Ejecutar sSql
    
    EliminaSerCon = 0
    Exit Function

EliminaSerConErr:
    Call RaiseError(MyUnhandledError, "DLogMov:EliminaSerCon Method")
End Function

'Inserta LogSerConDetalle
Public Function InsertaSerConDet(ByVal pnSerConNro As Long, ByVal pdFecha As Date, _
ByVal pnMonto As Currency) As Integer
On Error GoTo InsertaSerConDetErr
    Dim sSql As String, sActualiza As String
    InsertaSerConDet = 1
    
    sSql = "INSERT INTO LogSerConDetalle (nLogSerConNro, dLogSerConDetFecha, " & _
        "       nLogSerConDetImporte) " & _
        " VALUES (" & pnSerConNro & ",'" & Format(pdFecha, gsFormatoFecha) & "'," & _
                pnMonto & ")"
    oconect.Ejecutar sSql
    
    InsertaSerConDet = 0
    Exit Function

InsertaSerConDetErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSerConDet Method")
End Function

'Inserta LogSerConGarantia
Public Function InsertaSerConGar(ByVal pnSerConNro As Long, ByVal pnGarTpo As Integer, _
ByVal psGarNro As String, ByVal psGarDescri As String) As Integer
On Error GoTo InsertaSerConGarErr
    Dim sSql As String, sActualiza As String
    InsertaSerConGar = 1
    
    sSql = "INSERT INTO LogSerConGarantia (nLogSerConNro, nLogSerConGarTpo, " & _
        "       cLogSerConGarNro, cLogSerConGarDescripcion) " & _
        " VALUES (" & pnSerConNro & "," & pnGarTpo & ",'" & _
                psGarNro & "','" & psGarDescri & "')"
    oconect.Ejecutar sSql
    
    InsertaSerConGar = 0
    Exit Function

InsertaSerConGarErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSerConGar Method")
End Function

'Inserta AlmacenBS
Public Function InsertaAlmacenBS(ByVal pnAlmacen As Integer, ByVal psBSCod As String, _
ByVal pnStock As Currency) As Integer
On Error GoTo InsertaAlmacenBSErr
    Dim sSql As String, sActualiza As String
    InsertaAlmacenBS = 1
    
    sSql = "INSERT INTO AlmacenBS(nAlmCod, cBSCod, nAlmBSStock)" & _
        " VALUES (" & pnAlmacen & ",'" & psBSCod & "'," & pnStock & ")"
    oconect.Ejecutar sSql
    
    InsertaAlmacenBS = 0
    Exit Function

InsertaAlmacenBSErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaAlmacenBS Method")
End Function

'Actualiza AlmacenBS
Public Function ActualizaAlmacenBS(ByVal pnAlmacen As Integer, ByVal psBSCod As String, _
ByVal pnCantid As Currency, Optional ByVal pbSuma As Boolean = True) As Integer
On Error GoTo ActualizaAlmacenBSErr
    Dim sSql As String, sActualiza As String
    Dim sTmp As String
    ActualizaAlmacenBS = 1
    
    'sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sTmp = IIf(pbSuma = True, " + ", " - ")
    sSql = "UPDATE AlmacenBS SET nAlmBSStock = nAlmBSStock " & sTmp & pnCantid & "  " & _
        " WHERE nAlmCod = " & pnAlmacen & " AND cBSCod = '" & psBSCod & "'"
    
    oconect.Ejecutar sSql
    
    ActualizaAlmacenBS = 0
    Exit Function

ActualizaAlmacenBSErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaAlmacenBS Method")
End Function

Public Function ActualizaReqControlConsol(pnCodigo As Integer, pnPeriodo As Integer, pntipreq As Integer, pnEstado As Integer, psActualiza As String) As Integer
On Error GoTo EliminaSerConErr
    Dim sSql As String
    
    ActualizaReqControlConsol = 1
    
    sSql = "update LogReqControlConsol" _
           & " set nLogConsolEstado = 3, " _
           & " cUltimaActualizacion = '" & psActualiza & "' " _
           & " where nLogControlCod = " & pnCodigo & " " _
           & " and  nLogConsolEstado =" & pnEstado & " " _
           & "  and  nLogReqPeriodo = " & pnPeriodo & " " _
           & "  and  nLogReqTpo = " & pntipreq & " "
      oconect.Ejecutar sSql
      ActualizaReqControlConsol = 0
      Exit Function
EliminaSerConErr:
    Call RaiseError(MyUnhandledError, "DLogMov:EliminaSerCon Method")
End Function

Public Function ProcesaReqConsol(pnPeriodo As Integer, pnTipoReq As Integer, psActualiza As String) As Integer
    On Error GoTo GeneraMovNroErr
    Dim sql As String
    sql = "spLogProcesaLogReqConsol_01  " & pnPeriodo & " ," & pnTipoReq & ",'" & psActualiza & "'"
    oconect.Ejecutar (sql)
    ProcesaReqConsol = 0
    Exit Function
GeneraMovNroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function

Public Function EliminaReqConsol(pnPeriodo As Integer, pnTipoReq As Integer, psActualiza As String, pnCodConsol As Integer) As Integer
    On Error GoTo GeneraMovNroErr
    Dim sql As String
         sql = " Update logrequerimiento " _
              & " Set nLogReqConsolidado = 0  " _
              & " where  nLogReqCod  in(select  Distinct nLogReqCod from logreqconsol_01 where  NLogControlCod = " & pnCodConsol & ") and  " _
              & " nLogReqPeriodo = " & pnPeriodo & "  and  nLogReqTpo = " & pnTipoReq & " "
          oconect.Ejecutar (sql)
            
          sql = "Delete from logreqconsol_01 where nLogReqPeriodo = " & pnPeriodo & "  and  nLogReqTpo = " & pnTipoReq & " and NLogControlCod = " & pnCodConsol & " "
          oconect.Ejecutar (sql)
          sql = " Update LogReqControlConsol " _
          & " set  nLogConsolEstado = 2  ," _
          & " cUltimaActualizacion  ='" & psActualiza & "' " _
          & "  where nLogReqPeriodo = " & pnPeriodo & "  and  nLogReqTpo = " & pnTipoReq & "  and  nLogConsolEstado  = 1 and NLogControlCod = " & pnCodConsol & "  "
          oconect.Ejecutar (sql)
          EliminaReqConsol = 0
    Exit Function
GeneraMovNroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function

Public Function ActualizaReqListaPrecios(ByVal psBSCod As String, ByVal pnPeriodo As Integer, ByVal psReqNro As String, ByVal pnPrecioRefer As Currency, ByVal psActualizacion As String) As Integer
On Error GoTo ActualizaReqListaPreciosErr
    Dim sSql As String, sActualiza As String
    ActualizaReqListaPrecios = 1
    sActualiza = Left(psActualizacion, 8) & Format(oconect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    sSql = "Update  LogReqBsPrecios " _
    & " set nPrecReferencial = " & pnPrecioRefer & ", " _
    & " cUltimaActualizacion = '" & sActualiza & "', " _
    & " nLogReqCod = '" & psReqNro & "' " _
    & " where  rtrim(cBSCod) ='" & psBSCod & "' and  nperiodo  = " & pnPeriodo & " "
    oconect.Ejecutar sSql
    ActualizaReqListaPrecios = 0
    Exit Function
ActualizaReqListaPreciosErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaReqDetalle Method")
End Function
'EJVG20140210 ***
Public Sub CierreSaldosAlmacen(ByVal pnAlmTpo As Integer, ByVal pnAlmCod As Integer, ByVal pdFecha As Date)
    On Error GoTo ErrCierreSaldosAlmacen
    Dim lsSQL As String
    lsSQL = "Exec stp_ins_CierreSaldosAlmacen  " & pnAlmTpo & " ," & pnAlmCod & ",'" & Format(pdFecha, "yyyymmdd") & "'"
    oconect.Ejecutar (lsSQL)
    Exit Sub
ErrCierreSaldosAlmacen:
    Call RaiseError(MyUnhandledError, "DLogMov:CierreSaldosAlmacen Method")
End Sub
'END EJVG *******
