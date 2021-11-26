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

Dim oConect As DConecta

Private Sub Class_Initialize()
Dim oini As ClasIni
Set oini = New ClasIni

vsConexion = oini.CadenaConexion
vsServerAdmin = oini.BaseAdministracion
vsServerComunes = oini.BaseComunes
vsServerImagenes = oini.BaseImagenes
vsServerNegocio = oini.BaseNegocio
vsServerPersonas = oini.BasePersonas
 
Set oConect = New DConecta
If oConect.AbreConexion = False Then
    Call RaiseError(MyUnhandledError, "DLogMov:Initialize Method. Error en Conexion a Base de datos")
End If
End Sub

Private Sub Class_Terminate()
oConect.CierraConexion
Set oConect = Nothing
End Sub

'Ejecuta procesos Batch
Public Function EjecutaBatch() As Integer
On Error GoTo EjecutaBatchErr
    EjecutaBatch = oConect.EjecutarBatch
    Exit Function
EjecutaBatchErr:
    Call RaiseError(MyUnhandledError, "DLogMov:EjecutaBatch Method")
End Function

'Inserta Mov
Public Function InsertaMov(ByVal psMovNro As String, ByVal psOpeCod As String, _
    ByVal psMovDesc As String, ByVal psMovEstado As String, _
    Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaMovErr
    Dim sSql As String, sActualiza As String
    InsertaMov = 1
    
    sSql = "INSERT INTO Mov (cMovNro, cOpeCod, " & _
        " cMovDesc, cMovEstado) " & _
    " VALUES ('" & psMovNro & "','" & psOpeCod & "','" & _
        psMovDesc & "','" & psMovEstado & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaMov = 0
    Exit Function

InsertaMovErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaMov Method")
End Function
'Inserta MovRef
Public Function InsertaMovRef(ByVal psMovNro As String, ByVal psMovNroRef As String, _
    Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaMovRefErr
    Dim sSql As String, sActualiza As String
    InsertaMovRef = 1
    
    sSql = "INSERT INTO MovRef (cMovNro, cMovNroRef) " & _
    " VALUES ('" & psMovNro & "','" & psMovNroRef & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaMovRef = 0
    Exit Function

InsertaMovRefErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaMovRef Method")
End Function

'Inserta LogRequerimiento
Public Function InsertaRequeri(ByVal psReqNro As String, ByVal psPeriodo As String, _
    ByVal psReqTpo As String, ByVal psNecesi As String, ByVal psRequeri As String, _
    Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaRequeriErr
    Dim sSql As String, sActualiza As String
    InsertaRequeri = 1
    
    sSql = "INSERT INTO LogRequerimiento (cLogReqNro, cLogReqPeriodo, cLogReqTpo, " & _
        " cLogReqNecesidad, cLogReqRequerimiento) " & _
    " VALUES ('" & psReqNro & "','" & psPeriodo & "','" & psReqTpo & "','" & _
        psNecesi & "','" & psRequeri & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaRequeri = 0
    Exit Function

InsertaRequeriErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaRequeri Method")
End Function

'Inserta LogReqTramite
Public Function InsertaReqTramite(ByVal psReqNro As String, ByVal psReqTraNro As String, _
    ByVal psAreaCod As String, ByVal psAreaCodDes As String, ByVal psReqTraCom As String, ByVal psReqTraEst As String, _
    ByVal pnReqTraFlu As Integer, ByVal psActualizacion As String, _
    Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaReqTramiteErr
    Dim sSql As String, sActualiza As String
    InsertaReqTramite = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogReqTramite (cLogReqNro, cLogReqTraNro, cAreaCod, " & _
        " cAreaCodDes, cLogReqTraComentario, cLogReqTraEstado, nLogReqTraFlujo, " & _
        " cUltimaActualizacion)" & _
    " VALUES ('" & psReqNro & "','" & psReqTraNro & "','" & psAreaCod & "','" & _
        psAreaCodDes & "','" & psReqTraCom & "','" & psReqTraEst & "','" & pnReqTraFlu & "','" & _
        sActualiza & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaReqTramite = 0
    Exit Function

InsertaReqTramiteErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaReqTramite Method")
End Function

'Inserta LogReqDetalle
Public Function InsertaReqDetalle(ByVal psReqNro As String, ByVal psReqTraNro As String, _
    ByVal psBSCod As String, ByVal psReqRefMoneda As String, ByVal pnReqRefPrecio As Currency, _
    ByVal pnReqPrecio As Currency, ByVal psCtaContCod As String, ByVal psActualizacion As String, _
    Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaReqDetalleErr
    Dim sSql As String, sActualiza As String
    
    InsertaReqDetalle = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    If Len(Trim(psCtaContCod)) = 0 Then
        sSql = "INSERT INTO LogReqDetalle (cLogReqNro, cLogReqTraNro, cBSCod, " & _
            " cLogReqDetRefMoneda, nLogReqDetRefPrecio, nLogReqDetPrecio, " & _
            " cUltimaActualizacion)" & _
        " VALUES ('" & psReqNro & "','" & psReqTraNro & "','" & psBSCod & "','" & _
            psReqRefMoneda & "'," & pnReqRefPrecio & "," & pnReqPrecio & ",'" & _
            sActualiza & "') "
    Else
        sSql = "INSERT INTO LogReqDetalle (cLogReqNro, cLogReqTraNro, cBSCod, " & _
            " cLogReqDetRefMoneda, nLogReqDetRefPrecio, nLogReqDetPrecio, " & _
            " cCtaContCod, cUltimaActualizacion)" & _
        " VALUES ('" & psReqNro & "','" & psReqTraNro & "','" & psBSCod & "','" & _
            psReqRefMoneda & "'," & pnReqRefPrecio & "," & pnReqPrecio & ",'" & _
            psCtaContCod & "','" & sActualiza & "') "
    End If
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaReqDetalle = 0
    Exit Function

InsertaReqDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaReqDetalle Method")
End Function

'Inserta LogReqDetMes
Public Function InsertaReqDetMes(ByVal psReqNro As String, ByVal psReqTraNro As String, _
    ByVal psBSCod As String, ByVal pnReqDetMesNro As Integer, ByVal pnReqDetMesCant As Currency, _
    Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaReqDetMesErr
    Dim sSql As String, sActualiza As String
    InsertaReqDetMes = 1
    
    sSql = "INSERT INTO LogReqDetMes (cLogReqNro, cLogReqTraNro, cBSCod, " & _
        " nLogReqDetMesNro, nLogReqDetMesCantidad)" & _
    " VALUES ('" & psReqNro & "','" & psReqTraNro & "','" & psBSCod & "'," & _
        pnReqDetMesNro & "," & pnReqDetMesCant & ")"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaReqDetMes = 0
    Exit Function

InsertaReqDetMesErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaReqDetMes Method")
End Function


'Actualiza LogRequerimiento
Public Function ActualizaRequeri(ByVal psReqNro As String, ByVal psPeriodo As String, _
    ByVal psReqTpo As String, ByVal psNecesi As String, ByVal psRequeri As String, _
    Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo ActualizaRequeriErr
    Dim sSql As String, sActualiza As String
    ActualizaRequeri = 1
    
    sSql = "UPDATE LogRequerimiento SET cLogReqPeriodo = '" & psPeriodo & "', " & _
        " cLogReqTpo = '" & psReqTpo & "', cLogReqNecesidad = '" & psNecesi & "', " & _
        " cLogReqRequerimiento = '" & psRequeri & "'" & _
        " WHERE cLogReqNro = '" & psReqNro & "'"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    ActualizaRequeri = 0
    Exit Function

ActualizaRequeriErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaRequeri Method")
End Function

'Inserta LogReqTramite
Public Function ActualizaReqTramite(ByVal psReqNro As String, ByVal psReqTraNro As String, _
    ByVal psAreaCodDes As String, ByVal psReqTraEst As String, ByVal psReqTraCom As String, _
    ByVal pnReqTraFlu As Integer, ByVal psActualizacion As String, _
    Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo ActualizaReqTramiteErr
    Dim sSql As String, sActualiza As String
    ActualizaReqTramite = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "UPDATE LogReqTramite SET cAreaCodDes = '" & psAreaCodDes & "', " & _
        " cLogReqTraComentario = '" & psReqTraCom & "', cLogReqTraEstado = '" & psReqTraEst & "', " & _
        " nLogReqTraFlujo = " & pnReqTraFlu & ", cUltimaActualizacion = '" & sActualiza & "' " & _
        " WHERE cLogReqNro = '" & psReqNro & "' AND cLogReqTraNro = '" & psReqTraNro & "' "
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    ActualizaReqTramite = 0
    Exit Function

ActualizaReqTramiteErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaReqTramite Method")
End Function

'Elimina LogReqDetalle
Public Function EliminaReqDetalle(ByVal psReqNro As String, ByVal psReqTraNro As String, _
    Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo EliminaReqDetalleErr
    Dim sSql As String, sActualiza As String
    EliminaReqDetalle = 1
    
    If Len(Trim(psReqTraNro)) = 0 Then
        sSql = "DELETE LogReqDetalle " & _
            " WHERE cLogReqNro = '" & psReqNro & "' "
    Else
        sSql = "DELETE LogReqDetalle " & _
            " WHERE cLogReqNro = '" & psReqNro & "' AND cLogReqTraNro = '" & psReqTraNro & "'"
    End If
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    EliminaReqDetalle = 0
    Exit Function

EliminaReqDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:EliminaReqDetalle Method")
End Function

'Elimina LogReqDetalle
Public Function EliminaReqDetMes(ByVal psReqNro As String, ByVal psReqTraNro As String, _
    Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo EliminaReqDetMesErr
    Dim sSql As String, sActualiza As String
    EliminaReqDetMes = 1
    
    If Len(Trim(psReqTraNro)) = 0 Then
        sSql = "DELETE LogReqDetMes " & _
            " WHERE cLogReqNro = '" & psReqNro & "' "
    Else
        sSql = "DELETE LogReqDetMes " & _
            " WHERE cLogReqNro = '" & psReqNro & "' AND cLogReqTraNro = '" & psReqTraNro & "'"
    End If
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    EliminaReqDetMes = 0
    Exit Function

EliminaReqDetMesErr:
    Call RaiseError(MyUnhandledError, "DLogMov:EliminaReqDetMes Method")
End Function

'********************************************************************************************
'************************ PLAN DE OBTENCION *************************************************
'********************************************************************************************
'Inserta LogReqObt
Public Function InsertaReqObt(ByVal psReqNro As String, ByVal psObtNro As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaReqObtErr
    Dim sSql As String, sActualiza As String
    InsertaReqObt = 1
    
    sSql = "INSERT INTO LogReqObt (cLogReqNro, cLogObtNro) " & _
        " VALUES ('" & psReqNro & "','" & psObtNro & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaReqObt = 0
    Exit Function

InsertaReqObtErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaReqObt Method")
End Function

'Inserta LogObtencion
Public Function InsertaObtencion(ByVal psObtTpo As String, ByVal psObtNro As String, _
ByVal psPeriodo As String, ByVal psActualizacion As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaObtencionErr
    Dim sSql As String, sActualiza As String
    InsertaObtencion = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogObtencion (cLogObtNro, cLogObtTpo, cLogObtPeriodo, " & _
        " cLogObtEstado, cUltimaActualizacion) " & _
        " VALUES ('" & psObtNro & "','" & psObtTpo & "','" & psPeriodo & "','" & _
        gLogObtEstadoInicio & "','" & sActualiza & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaObtencion = 0
    Exit Function

InsertaObtencionErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaObtencion Method")
End Function

'Inserta LogObtDetalle
Public Function InsertaObtDetalle(ByVal psObtNro As String, ByVal psBSCod As String, _
ByVal pnCantidad As Currency, ByVal pnPrecio As Currency, ByVal psCtaCont As String, _
ByVal psActualizacion As String, Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaObtDetalleErr
    Dim sSql As String, sActualiza As String
    InsertaObtDetalle = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    If Trim(psCtaCont) = "" Then
        sSql = "INSERT INTO LogObtDetalle (cLogObtNro, cBSCod, " & _
            " nLogObtCantidad, nLogObtPrecio, cUltimaActualizacion) " & _
            " VALUES ('" & psObtNro & "','" & psBSCod & "'," & _
            pnCantidad & "," & pnPrecio & ",'" & sActualiza & "')"
    Else
        sSql = "INSERT INTO LogObtDetalle (cLogObtNro, cBSCod, " & _
            " nLogObtCantidad, nLogObtPrecio, cCtaContCod, cUltimaActualizacion) " & _
            " VALUES ('" & psObtNro & "','" & psBSCod & "'," & _
            pnCantidad & "," & pnPrecio & ",'" & psCtaCont & "','" & sActualiza & "')"
    End If
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaObtDetalle = 0
    Exit Function

InsertaObtDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaObtDetalle Method")
End Function

'Actualiza LogObtencion
Public Function ActualizaObtencion(ByVal psObtNro As String, ByVal psEstado As LogObtEstado, _
ByVal psActualizacion As String, Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo ActualizaObtencionErr
    Dim sSql As String, sActualiza As String
    ActualizaObtencion = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "UPDATE LogObtencion SET cLogObtEstado = '" & psEstado & "'," & _
        "       cUltimaActualizacion = '" & sActualiza & "'" & _
        " WHERE cLogObtNro = '" & psObtNro & "' "
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    ActualizaObtencion = 0
    Exit Function

ActualizaObtencionErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaObtencion Method")
End Function

'Actualiza LogObtDetalle
Public Function ActualizaObtDetalle(ByVal psObtNro As String, ByVal psBSCod As String, _
ByVal psCtaCont As String, ByVal psActualizacion As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo ActualizaObtDetalleErr
    Dim sSql As String, sActualiza As String
    ActualizaObtDetalle = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "UPDATE LogObtDetalle SET cCtaContCod = '" & psCtaCont & "'," & _
        "       cUltimaActualizacion = '" & sActualiza & "'" & _
        " WHERE cLogObtNro = '" & psObtNro & "' " & _
        "       AND cBSCod = '" & psBSCod & "' "
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    ActualizaObtDetalle = 0
    Exit Function

ActualizaObtDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaObtDetalle Method")
End Function


'Inserta LogObtAdq
Public Function InsertaObtAdq(ByVal psObtNro As String, ByVal psAdqNro As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaObtAdqErr
    Dim sSql As String, sActualiza As String
    InsertaObtAdq = 1
    
    sSql = "INSERT INTO LogObtAdq (cLogObtNro, cLogAdqNro) " & _
        " VALUES ('" & psObtNro & "','" & psAdqNro & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaObtAdq = 0
    Exit Function

InsertaObtAdqErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaObtAdq Method")
End Function

'Inserta LogAdquisicion
Public Function InsertaAdquisicion(ByVal psAdqTpo As String, ByVal psAdqNro As String, _
ByVal psPeriodo As String, ByVal psActualizacion As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaAdquisicionErr
    Dim sSql As String, sActualiza As String
    InsertaAdquisicion = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogAdquisicion (cLogAdqNro, cLogAdqTpo, cLogAdqPeriodo, " & _
        " cLogAdqEstado, cUltimaActualizacion) " & _
        " VALUES ('" & psAdqNro & "','" & psAdqTpo & "','" & psPeriodo & "','" & _
        gLogAdqEstadoInicio & "','" & sActualiza & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaAdquisicion = 0
    Exit Function

InsertaAdquisicionErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaAdquisicion Method")
End Function

'Inserta LogAdqDetalle
Public Function InsertaAdqDetalle(ByVal psAdqNro As String, ByVal psBSCod As String, _
ByVal pnCantidad As Currency, ByVal pnPrecio As Currency, ByVal psActualizacion As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaAdqDetalleErr
    Dim sSql As String, sActualiza As String
    InsertaAdqDetalle = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogAdqDetalle (cLogAdqNro, cBSCod, " & _
        " nLogAdqDetCantidad, nLogAdqDetPrecio, cUltimaActualizacion) " & _
        " VALUES ('" & psAdqNro & "','" & psBSCod & "'," & _
        pnCantidad & "," & pnPrecio & ",'" & sActualiza & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaAdqDetalle = 0
    Exit Function

InsertaAdqDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaAdqDetalle Method")
End Function

'Actualiza LogAdquisicion
Public Function ActualizaAdquisicion(ByVal psAdqNro As String, ByVal psEstado As LogAdqEstado, _
ByVal psActualizacion As String, Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo ActualizaAdquisicionErr
    Dim sSql As String, sActualiza As String
    ActualizaAdquisicion = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "UPDATE LogAdquisicion SET cLogAdqEstado = '" & psEstado & "'," & _
        "       cUltimaActualizacion = '" & sActualiza & "'" & _
        " WHERE cLogAdqNro = '" & psAdqNro & "' "
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    ActualizaAdquisicion = 0
    Exit Function

ActualizaAdquisicionErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaAdquisicion Method")
End Function

'Inserta LogSeleccion
Public Function InsertaSeleccion(ByVal psSelNro As String, ByVal pdResolu As Date, _
ByVal psResNro As String, ByVal psAreaCod As String, ByVal psPersCod As String, _
ByVal psActualizacion As String, Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaSeleccionErr
    Dim sSql As String, sActualiza As String
    InsertaSeleccion = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogSeleccion (cLogSelNro, dLogSelRes, cLogSelResNro, " & _
        " cLogSelEstado, cAreaCod, cPersCod, cUltimaActualizacion) " & _
        " VALUES ('" & psSelNro & "','" & Format(pdResolu, "mm/dd/yyyy") & "','" & psResNro & "','" & _
        gLogSelEstadoInicioRes & "','" & psAreaCod & "','" & psPersCod & "','" & sActualiza & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaSeleccion = 0
    Exit Function

InsertaSeleccionErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSeleccion Method")
End Function

'Inserta LogSelComite
Public Function InsertaSelComite(ByVal psSelNro As String, ByVal psAreaCod As String, _
ByVal psPersCod As String, ByVal psActualizacion As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaSelComiteErr
    Dim sSql As String, sActualiza As String
    InsertaSelComite = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogSelComite (cLogSelNro, cAreaCod, cPersCod, cUltimaActualizacion) " & _
        " VALUES ('" & psSelNro & "','" & psAreaCod & "','" & psPersCod & "','" & sActualiza & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaSelComite = 0
    Exit Function

InsertaSelComiteErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelComite Method")
End Function

'Actualiza LogSeleccion
Public Function ActualizaSeleccion(ByVal psSelNro As String, ByVal pdResolu As Date, _
ByVal psResNro As String, ByVal psAreaCod As String, ByVal psPersCod As String, _
ByVal psActualizacion As String, Optional psEstado As LogSelEstado, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo ActualizaSeleccionErr
    Dim sSql As String, sActualiza As String
    ActualizaSeleccion = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    If psResNro <> "" And psAreaCod <> "" And psPersCod <> "" Then
        sSql = "UPDATE LogSeleccion SET dLogSelRes = '" & Format(pdResolu, "mm/dd/yyyy") & "', " & _
            "       cLogSelResNro = '" & psResNro & "', cAreaCod = '" & psAreaCod & "'," & _
            "       cPersCod = '" & psPersCod & "', cUltimaActualizacion = '" & sActualiza & "' " & _
            " WHERE cLogSelNro = '" & psSelNro & "' "
    Else
        sSql = "UPDATE LogSeleccion SET cLogSelEstado = '" & psEstado & "'," & _
            "       cUltimaActualizacion = '" & sActualiza & "'" & _
            " WHERE cLogSelNro = '" & psSelNro & "' "
    End If
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    ActualizaSeleccion = 0
    Exit Function

ActualizaSeleccionErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaSeleccion Method")
End Function

'Actualiza LogSeleccion (Base)
Public Function ActualizaSeleccionBase(ByVal psSelNro As String, ByVal pnCostoBase As Currency, _
ByVal psAdqNro As String, ByVal psSisAdj As String, ByVal psActualizacion As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo ActualizaSeleccionBaseErr
    Dim sSql As String, sActualiza As String
    ActualizaSeleccionBase = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    sSql = "UPDATE LogSeleccion SET cLogSelCotNro = '" & psAdqNro & "', " & _
        "       cLogSelSisAdj = '" & psSisAdj & "', nLogSelCostoBase = " & pnCostoBase & "," & _
        "       cUltimaActualizacion = '" & sActualiza & "', cLogSelEstado = '" & gLogSelEstadoBases & "' " & _
        " WHERE cLogSelNro = '" & psSelNro & "' "
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    ActualizaSeleccionBase = 0
    Exit Function

ActualizaSeleccionBaseErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaSeleccionBase Method")
End Function

'Actualiza LogSeleccion (Adju)
Public Function ActualizaSeleccionAdju(ByVal psSelNro As String, ByVal psSelCotNro As String, _
ByVal psActualizacion As String, Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo ActualizaSeleccionAdjuErr
    Dim sSql As String, sActualiza As String
    ActualizaSeleccionAdju = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    sSql = "UPDATE LogSeleccion SET cLogSelCotNro = '" & psSelCotNro & "', " & _
        "       cUltimaActualizacion = '" & sActualiza & "', cLogSelEstado = '" & gLogSelEstadoProcAdju & "' " & _
        " WHERE cLogSelNro = '" & psSelNro & "' "
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    ActualizaSeleccionAdju = 0
    Exit Function

ActualizaSeleccionAdjuErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaSeleccionAdju Method")
End Function

'Inserta LogSelParametro
Public Function InsertaSelParametro(ByVal psSelNro As String, ByVal psSelParTpo As String, _
ByVal pnSelParNro As Integer, ByVal pnSelParPuntaje As Integer, ByVal psActualizacion As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaSelParametroErr
    Dim sSql As String, sActualiza As String
    InsertaSelParametro = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "INSERT INTO LogSelParametro (cLogSelNro, cLogSelParTpo, nLogSelParNro, " & _
        "       nLogSelParPuntaje, cUltimaActualizacion) " & _
        " VALUES ('" & psSelNro & "','" & psSelParTpo & "'," & pnSelParNro & ", " & _
        "" & pnSelParPuntaje & ",'" & sActualiza & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaSelParametro = 0
    Exit Function

InsertaSelParametroErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelParametro Method")
End Function

'Inserta LogSelPublica
Public Function InsertaSelPublica(ByVal psSelNro As String, ByVal psPersCod As String, _
ByVal pdPubIni As Date, ByVal pdPubFin As Date, ByVal psActualizacion As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaSelPublicaErr
    Dim sSql As String, sActualiza As String
    InsertaSelPublica = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

    sSql = "INSERT INTO LogSelPublica (cLogSelNro, cPersCod, dLogSelPubIni, " & _
        "       dLogSelPubFin, cUltimaActualizacion) " & _
        " VALUES ('" & psSelNro & "','" & psPersCod & "','" & Format(pdPubIni, "mm/dd/yyyy") & "', " & _
        "'" & Format(pdPubFin, "mm/dd/yyyy") & "','" & sActualiza & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaSelPublica = 0
    Exit Function

InsertaSelPublicaErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelPublica Method")
End Function

'Inserta LogSelPostor
Public Function InsertaSelPostor(ByVal psSelNro As String, ByVal psPersCod As String, _
ByVal psActualizacion As String, Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaSelPostorErr
    Dim sSql As String, sActualiza As String
    InsertaSelPostor = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

    sSql = "INSERT INTO LogSelPostor (cLogSelNro, cPersCod, cUltimaActualizacion) " & _
        " VALUES ('" & psSelNro & "','" & psPersCod & "','" & sActualiza & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaSelPostor = 0
    Exit Function

InsertaSelPostorErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelPostor Method")
End Function

'Elimina LogSelPostor
Public Function EliminaSelPostor(ByVal psSelNro As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo EliminaSelPostorErr
    Dim sSql As String, sActualiza As String
    EliminaSelPostor = 1
    
    sSql = "DELETE LogSelPostor " & _
        " WHERE cLogSelNro = '" & psSelNro & "' "
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    EliminaSelPostor = 0
    Exit Function

EliminaSelPostorErr:
    Call RaiseError(MyUnhandledError, "DLogMov:EliminaSelPostor Method")
End Function


'Actualiza LogSelPostor
Public Function ActualizaSelPostor(ByVal psSelNro As String, ByVal psPersCod As String, _
ByVal psObserva As String, ByVal psActualizacion As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo ActualizaSelPostorErr
    Dim sSql As String, sActualiza As String
    ActualizaSelPostor = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)
    
    sSql = "UPDATE LogSelPostor SET cLogSelPosObserva = '" & psObserva & "', " & _
        "       cUltimaActualizacion = '" & sActualiza & "'" & _
        " WHERE cLogSelNro = '" & psSelNro & "' AND cPersCod = '" & psPersCod & "'"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    ActualizaSelPostor = 0
    Exit Function

ActualizaSelPostorErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaSelPostor Method")
End Function

'Inserta LogSelCotiza
Public Function InsertaSelCotiza(ByVal psSelNro As String, ByVal psSelCotNro As String, _
ByVal psPersCod As String, ByVal psActualizacion As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaSelCotizaErr
    Dim sSql As String, sActualiza As String
    InsertaSelCotiza = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

    sSql = "INSERT INTO LogSelCotiza (cLogSelNro, cLogSelCotNro, " & _
        "       cPersCod, cUltimaActualizacion) " & _
        " VALUES ('" & psSelNro & "','" & psSelCotNro & "','" & _
                psPersCod & "','" & sActualiza & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaSelCotiza = 0
    Exit Function

InsertaSelCotizaErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelCotiza Method")
End Function

'Inserta LogSelCotDetalle
Public Function InsertaSelCotDetalle(ByVal psSelCotNro As String, ByVal psBSCod As String, _
ByVal pnCantidad As Currency, ByVal pnPrecio As Currency, ByVal psActualizacion As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo InsertaSelCotDetalleErr
    Dim sSql As String, sActualiza As String
    InsertaSelCotDetalle = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

    sSql = "INSERT INTO LogSelCotDetalle (cLogSelCotNro, cBSCod, " & _
        "       nLogSelCotDetCantidad, nLogSelCotDetPrecio, cUltimaActualizacion) " & _
        " VALUES ('" & psSelCotNro & "','" & psBSCod & "'," & _
                pnCantidad & "," & pnPrecio & ",'" & sActualiza & "')"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    InsertaSelCotDetalle = 0
    Exit Function

InsertaSelCotDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:InsertaSelCotDetalle Method")
End Function


'Actualiza LogSelCotDetalle
Public Function ActualizaSelCotDetalle(ByVal psSelCotNro As String, ByVal psBSCod As String, _
ByVal pnCantidad As Currency, ByVal pnPrecio As Currency, ByVal psActualizacion As String, _
Optional lbEjecBatch As Boolean = False) As Integer
On Error GoTo ActualizaSelCotDetalleErr
    Dim sSql As String, sActualiza As String
    ActualizaSelCotDetalle = 1
    
    sActualiza = Left(psActualizacion, 8) & Format(oConect.GetHoraServer, "hhmmss") & Mid(psActualizacion, 15)

    sSql = "UPDATE LogSelCotDetalle SET nLogSelCotDetCantidad = " & pnCantidad & ",  " & _
        "       nLogSelCotDetPrecio = " & pnPrecio & ", cUltimaActualizacion = '" & sActualiza & "' " & _
        " WHERE cLogSelCotNro = '" & psSelCotNro & "' AND cBSCod = '" & psBSCod & "'"
    
    If lbEjecBatch Then
        oConect.AdicionaCmdBatch sSql
    Else
        oConect.Ejecutar sSql
    End If
    ActualizaSelCotDetalle = 0
    Exit Function

ActualizaSelCotDetalleErr:
    Call RaiseError(MyUnhandledError, "DLogMov:ActualizaSelCotDetalle Method")
End Function


